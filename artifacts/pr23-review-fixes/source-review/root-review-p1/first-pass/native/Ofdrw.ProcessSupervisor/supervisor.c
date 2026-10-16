#define _GNU_SOURCE
#include <errno.h>
#include <fcntl.h>
#include <limits.h>
#include <poll.h>
#include <signal.h>
#include <stdint.h>
#include <stdlib.h>
#include <sys/resource.h>
#include <sys/socket.h>
#include <sys/types.h>
#include <sys/wait.h>
#include <time.h>
#include <unistd.h>
#ifdef __linux__
#include <sys/syscall.h>
#endif
#ifdef __APPLE__
#include <libproc.h>
#endif
extern char **environ;
/* Versioned, local private protocol: four network-order uint32 words:
 * magic, kind (started/exit/start-error/cleaned/helper-error), target pid, value.
 * Host sends one K byte or closes the socket to request cleanup. */
enum { MAGIC=0x4f465301, STARTED=1, EXITED=2, START_ERROR=3, CLEANED=4, FAILED=5 };
enum { START_LIMIT_MS=6000, CLEAN_LIMIT_MS=5000, WRITE_LIMIT_MS=1000 };
static volatile sig_atomic_t interrupted;
static void on_signal(int signal_number) { (void)signal_number; interrupted=1; }
static int64_t milliseconds(void) {
    struct timespec value;
    if(clock_gettime(CLOCK_MONOTONIC,&value)<0) return -1;
    return (int64_t)value.tv_sec*1000+value.tv_nsec/1000000;
}
static int nonblocking(int fd) {
    int flags=fcntl(fd,F_GETFL);
    return flags<0?-1:fcntl(fd,F_SETFL,flags|O_NONBLOCK);
}
static int frame(int kind,pid_t pid,int value) {
    uint32_t words[4]={MAGIC,(uint32_t)kind,(uint32_t)pid,(uint32_t)value};
    unsigned char bytes[16];
    for(int i=0;i<4;i++) for(int j=0;j<4;j++) bytes[i*4+j]=(unsigned char)(words[i]>>(24-j*8));
    size_t offset=0; int64_t now=milliseconds();if(now<0)return -1;
    int64_t deadline=now+WRITE_LIMIT_MS;
    while(offset<sizeof(bytes)) {
        now=milliseconds();if(now<0)return -1;
        if(now>=deadline){errno=ETIMEDOUT;return -1;}
        ssize_t count=write(3,bytes+offset,sizeof(bytes)-offset);
        if(count>0) { offset+=(size_t)count;continue; }
        if(count<0&&errno==EINTR)continue;
        if(count==0||(errno!=EAGAIN&&errno!=EWOULDBLOCK))return -1;
        now=milliseconds();if(now<0)return -1;
        int64_t left=deadline-now;if(left<=0){errno=ETIMEDOUT;return -1;}
        struct pollfd descriptor={3,POLLOUT,0};
        if(poll(&descriptor,1,(int)left)<0&&errno!=EINTR)return -1;
    }
    return 0;
}
static int observe(pid_t pid,siginfo_t *status) {
    int result;do{result=waitid(P_PID,(id_t)pid,status,WEXITED|WNOHANG|WNOWAIT);}while(result<0&&errno==EINTR);
    return result;
}
static int exit_code(const siginfo_t *status) {
    return status->si_code==CLD_EXITED?status->si_status:128+status->si_status;
}
#ifdef __APPLE__
static int only_zombies(pid_t group) {
    pid_t members[4096];errno=0;int bytes=proc_listpgrppids(group,members,sizeof(members));
    if(bytes==0&&errno!=0)return 0;
    if(bytes<0||(size_t)bytes>=sizeof(members))return 0;
    for(int i=0;i<bytes/(int)sizeof(pid_t);i++) {
        if(!members[i])continue;
        struct proc_bsdshortinfo info;
        int length=proc_pidinfo(members[i],PROC_PIDT_SHORTBSDINFO,0,&info,sizeof(info));
        if(length==0&&errno==ESRCH)continue;
        if(length!=(int)sizeof(info)||info.pbsi_pgid!=(uint32_t)group||info.pbsi_status!=5)return 0;
    }
    return 1;
}
#endif
/* Called once, while target is still our waitable child. No caller ever signals
 * the numeric group after this function reaps or abandons that reservation. */
static int cleanup(pid_t pid,int group_owned,int started,int *reported) {
    siginfo_t status={0};if(observe(pid,&status)<0)return errno;
    if(kill(group_owned?-pid:pid,SIGKILL)<0) {
        int error=errno;
#ifdef __APPLE__
        if(error!=ESRCH&&!(group_owned&&error==EPERM&&status.si_pid==pid&&only_zombies(pid)))return error;
#else
        if(error!=ESRCH)return error;
#endif
    }
    int64_t initial=milliseconds();int clock_error=initial<0?errno:0;
    int64_t deadline=initial+CLEAN_LIMIT_MS;
    for(int attempt=0;attempt<500;attempt++) {
        status.si_pid=0;if(observe(pid,&status)<0)return errno;
        if(status.si_pid==pid) {
            if(started&&!*reported){frame(EXITED,pid,exit_code(&status));*reported=1;}
            int result;do{result=waitpid(pid,0,WNOHANG);}while(result<0&&errno==EINTR);
            return result==pid?clock_error:(result<0?errno:EIO);
        }
        if(!clock_error){int64_t now=milliseconds();if(now<0)clock_error=errno;else if(now>=deadline)return ETIMEDOUT;}
        struct timespec pause={0,10000000};nanosleep(&pause,0);
    }
    return clock_error?clock_error:ETIMEDOUT;
}
static int close_unrelated(void) {
    int64_t now=milliseconds();if(now<0)return -1;int64_t deadline=now+START_LIMIT_MS;
#ifdef __linux__
#ifdef SYS_close_range
    if(syscall(SYS_close_range,4,UINT_MAX,0)==0)return 0;
    /* A seccomp EPERM may forbid close_range while ordinary close works. */
#endif
#endif
    struct rlimit limit;if(getrlimit(RLIMIT_NOFILE,&limit)<0)return -1;
    if(limit.rlim_cur==RLIM_INFINITY||limit.rlim_cur>INT_MAX){errno=EOVERFLOW;return -1;}
    for(int fd=4;fd<(int)limit.rlim_cur;fd++){
        if((fd&255)==0){if(interrupted){errno=EINTR;return -1;}now=milliseconds();if(now<0)return -1;if(now>=deadline){errno=ETIMEDOUT;return -1;}}
        close(fd);
    }
    return 0;
}
static int make_pipe(int pair[2]) {
    if(pipe(pair)<0)return -1;
    if(fcntl(pair[0],F_SETFD,FD_CLOEXEC)<0||fcntl(pair[1],F_SETFD,FD_CLOEXEC)<0){int error=errno;close(pair[0]);close(pair[1]);errno=error;return -1;}
    return 0;
}
/* Target child has not been granted exec yet. Its only writers are fixed-size
 * native readiness/error records, so neither this write nor grant read allocates. */
static void child_record(int descriptor,int stage,int error) {
    int values[2]={stage,error};const char *data=(const char*)values;size_t left=sizeof(values);
    while(left){ssize_t count=write(descriptor,data,left);if(count<0&&errno==EINTR)continue;if(count<=0)break;data+=count;left-=(size_t)count;}
}
int main(int argc,char **argv) {
    if(argc<3)return 64;
    int socket_type;socklen_t socket_length=sizeof(socket_type);
    if(getsockopt(3,SOL_SOCKET,SO_TYPE,&socket_type,&socket_length)<0||socket_type!=SOCK_STREAM)return 64;
    /* posix_spawn starts a fresh single-threaded executable, never a fork of CLR.
     * Save ignored dispositions/mask before changing supervisor-only policy. */
    const int signals[]={SIGCHLD,SIGPIPE,SIGHUP,SIGINT,SIGTERM};
    struct sigaction original[5],action={0};sigset_t original_mask,empty;
    for(int i=0;i<5;i++)if(sigaction(signals[i],0,&original[i])<0)return 70;
    if(sigprocmask(SIG_SETMASK,0,&original_mask)<0)return 70;
    action.sa_handler=SIG_DFL;sigemptyset(&action.sa_mask);
    if(sigaction(SIGCHLD,&action,0)<0)return 70;
    action.sa_handler=SIG_IGN;if(sigaction(SIGPIPE,&action,0)<0)return 70;
    action.sa_handler=on_signal;
    for(int i=2;i<5;i++)if(sigaction(signals[i],&action,0)<0)return 70;
    sigemptyset(&empty);if(sigprocmask(SIG_SETMASK,&empty,0)<0)return 70;
    if(fcntl(3,F_SETFD,FD_CLOEXEC)<0||nonblocking(3)<0||close_unrelated()<0){frame(FAILED,0,errno);return 70;}
    int errors[2],grant[2];
    if(make_pipe(errors)<0){frame(FAILED,0,errno);return 70;}
    if(make_pipe(grant)<0){int error=errno;close(errors[0]);close(errors[1]);frame(FAILED,0,error);return 70;}
    pid_t pid=fork();
    if(pid<0){frame(FAILED,0,errno);return 70;}
    if(pid==0) {
        /* The only post-fork path: POSIX async-signal-safe operations, no CLR,
         * malloc, stdio, dynamic loader, application hooks, or callbacks. */
        close(errors[0]);close(grant[1]);close(3);
        for(int i=0;i<5;i++)if(sigaction(signals[i],&original[i],0)<0){child_record(errors[1],2,errno);_exit(127);}
        if(sigprocmask(SIG_SETMASK,&original_mask,0)<0||setpgid(0,0)<0){child_record(errors[1],2,errno);_exit(127);}
        child_record(errors[1],1,0);
        char permission;ssize_t count;do{count=read(grant[0],&permission,1);}while(count<0&&errno==EINTR);
        close(grant[0]);if(count!=1||permission!='G')_exit(125);
        execve(argv[1],argv+2,environ);child_record(errors[1],2,errno);_exit(127);
    }
    close(errors[1]);close(grant[0]);close(1);close(2);
    int group_owned=0,started=0,reported=0,failure=0,exec_error=0;
    size_t received=0;int record[2]={0,0};
    int64_t initial=milliseconds();
    if(initial<0){failure=errno;goto finish;}
    int64_t start_deadline=initial+START_LIMIT_MS;
    if(nonblocking(errors[0])<0||nonblocking(grant[1])<0){failure=errno;goto finish;}
    for(;;) {
        if(interrupted){failure=EINTR;break;}
        int64_t now=milliseconds();if(now<0){failure=errno;break;}
        if(!started&&now>=start_deadline){failure=ETIMEDOUT;break;}
        struct pollfd descriptors[2]={{3,POLLIN,0},{errors[0],POLLIN,0}};
        int ready=poll(descriptors,started?1:2,10);
        if(ready<0){if(errno==EINTR)continue;failure=errno;break;}
        /* Control wins if readiness and cancellation arrive together. Until G is
         * sent the child cannot execute target code: exact-child kill is safe. */
        if(descriptors[0].revents) {
            char command;ssize_t count=read(3,&command,1);
            if(count<0&&(errno==EINTR||errno==EAGAIN||errno==EWOULDBLOCK))continue;
            if(count<0){failure=errno;break;}
            if(count==1&&command!='K')failure=EPROTO;
            break;
        }
        if(!started&&descriptors[1].revents) {
            ssize_t count=read(errors[0],((char*)record)+received,sizeof(record)-received);
            if(count<0){if(errno==EINTR||errno==EAGAIN||errno==EWOULDBLOCK)continue;failure=errno;break;}
            if(count==0) {
                if(received||!group_owned){failure=EPROTO;break;}
                close(errors[0]);errors[0]=-1;started=1;
                if(frame(STARTED,pid,0)<0){failure=errno;break;}
            } else {
                received+=(size_t)count;
                if(received==sizeof(record)) {
                    received=0;
                    if(record[0]==2){exec_error=record[1];if(exec_error<=0)failure=EPROTO;break;}
                    if(record[0]!=1||record[1]!=0||group_owned){failure=EPROTO;break;}
                    /* Child is blocked on grant and cannot run target code.
                     * Verify group before permitting exec or any group signal. */
                    if(getpgid(pid)!=pid){failure=EPROTO;break;}
                    group_owned=1;char permission='G';
                    ssize_t written;do{written=write(grant[1],&permission,1);}while(written<0&&errno==EINTR);
                    if(written!=1){failure=written<0?errno:EIO;break;}
                    close(grant[1]);grant[1]=-1;
                }
            }
        }
        if(started&&!reported) {
            siginfo_t status={0};if(observe(pid,&status)<0){failure=errno;break;}
            if(status.si_pid==pid){if(frame(EXITED,pid,exit_code(&status))<0){failure=errno;break;}reported=1;}
        }
    }
finish:
    if(grant[1]>=0)close(grant[1]);
    if(errors[0]>=0)close(errors[0]);
    int clean_error=cleanup(pid,group_owned,started,&reported);
    if(exec_error)frame(START_ERROR,pid,exec_error);
    else if(failure||clean_error)frame(FAILED,pid,failure?failure:clean_error);
    else frame(CLEANED,pid,0);
    close(3);return failure||clean_error?70:0;
}
