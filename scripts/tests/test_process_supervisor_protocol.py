"""Exercise the shipped native owner with a private socket, never global PID scans."""
import errno
try:
    import fcntl
except ImportError:  # Windows has no Unix descriptor API
    fcntl = None
import os
import pathlib
import platform
import signal
import socket
import struct
import subprocess
import sys
import tempfile
import time
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[2]
MAGIC = 0x4F465301


@unittest.skipIf(os.name == "nt", "Windows uses its existing native Job backend")
class ProcessSupervisorProtocolTests(unittest.TestCase):
    def setUp(self):
        architecture = {"arm64": "arm64", "aarch64": "arm64", "x86_64": "x64", "AMD64": "x64"}[platform.machine()]
        rid = ("osx-" if sys.platform == "darwin" else "linux-") + architecture
        self.helper = ROOT / "native/Ofdrw.ProcessSupervisor/assets" / rid / "ofdrw-supervisor"
        self.processes = []
        self.sockets = []

    def tearDown(self):
        # Closing the private endpoint asks the native owner to kill/reap only
        # its retained child group. No PID-name enumeration is involved.
        for endpoint in self.sockets:
            endpoint.close()
        for process in self.processes:
            try:
                process.communicate(timeout=8)
            except subprocess.TimeoutExpired:
                # Exact Popen-owned child only; timeout remains a test failure.
                process.kill()
                process.communicate(timeout=3)
                self.fail("Native owner exceeded its cleanup deadline")

    def launch(self, executable="/bin/sh", arguments=("-c", "exit 0"), ignored=False, argv0="target-argv0", foreign_descriptor=None):
        peer, child = socket.socketpair()
        peer.settimeout(8)
        self.sockets.append(peer)
        child_descriptor = child.fileno()

        def prepare():
            os.dup2(child_descriptor, 3)
            if ignored:
                signal.signal(signal.SIGCHLD, signal.SIG_IGN)

        process = subprocess.Popen([str(self.helper), executable, argv0, *arguments],
                                   pass_fds=tuple(set((3, child_descriptor) + (() if foreign_descriptor is None else (foreign_descriptor,)))), preexec_fn=prepare,
                                   stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        child.close()
        self.processes.append(process)
        return process, peer

    def frame(self, peer):
        data = bytearray()
        while len(data) < 16:
            part = peer.recv(16 - len(data))
            self.assertTrue(part, "Supervisor closed a partial frame")
            data.extend(part)
        magic, kind, pid, value = struct.unpack("!4I", data)
        self.assertEqual(MAGIC, magic)
        self.assertGreater(pid, 0)
        return kind, pid, value

    def test_actual_zero_and_nonzero_status_under_inherited_ignore(self):
        for code in (0, 7):
            with self.subTest(code=code):
                process, peer = self.launch(arguments=("-c", f"exit {code}"), ignored=True)
                started = self.frame(peer)
                exited = self.frame(peer)
                self.assertEqual((1, started[1], 0), started)
                self.assertEqual((2, started[1], code), exited)
                peer.sendall(b"K")
                self.assertEqual((4, started[1], 0), self.frame(peer))
                self.assertEqual(0, process.wait(timeout=3))

    def test_native_exec_error_without_shell_fallback(self):
        with tempfile.TemporaryDirectory() as folder:
            invalid = pathlib.Path(folder) / "invalid"
            invalid.write_text("this is not an executable\n")
            invalid.chmod(0o700)
            for target, expected in ((str(invalid), errno.ENOEXEC), (str(invalid) + "-missing", errno.ENOENT)):
                with self.subTest(target=target):
                    process, peer = self.launch(target, ())
                    kind, _, actual = self.frame(peer)
                    self.assertEqual((3, expected), (kind, actual))
                    self.assertEqual(4, self.frame(peer)[0])
                    self.assertEqual(b"", process.communicate(timeout=3)[0])
                    self.assertEqual(0, process.returncode)

    def test_exact_original_argv0_and_two_output_pipes(self):
        process, peer = self.launch(arguments=("-c", 'printf "%s" "$0"; printf "error" >&2'), argv0="original 中文")
        self.assertEqual(1, self.frame(peer)[0])
        self.assertEqual(2, self.frame(peer)[0])
        peer.sendall(b"K")
        self.assertEqual(4, self.frame(peer)[0])
        output, error = process.communicate(timeout=3)
        self.assertEqual("original 中文", output.decode())
        self.assertEqual(b"error", error)

    def test_parent_first_exit_keeps_group_reserved(self):
        process, peer = self.launch(arguments=("-c", "sleep 30 & exit 7"))
        started = self.frame(peer)
        self.assertEqual((2, started[1], 7), self.frame(peer))
        peer.sendall(b"K")
        self.assertEqual((4, started[1], 0), self.frame(peer))
        process.communicate(timeout=3)  # inherited capture writers have closed

    def test_protocol_eof_cleans_target_without_ack(self):
        process, peer = self.launch(arguments=("-c", "sleep 30 & wait"))
        self.assertEqual(1, self.frame(peer)[0])
        peer.close()
        process.communicate(timeout=3)
        self.assertEqual(0, process.returncode)

    def test_corrupt_command_cleans_and_reports_failure(self):
        process, peer = self.launch(arguments=("-c", "sleep 30 & wait"))
        started = self.frame(peer)
        peer.sendall(b"?")
        self.assertEqual((2, started[1], 137), self.frame(peer))
        self.assertEqual((5, started[1], errno.EPROTO), self.frame(peer))
        self.assertEqual(70, process.wait(timeout=3))

    def test_supervisor_termination_runs_cleanup_and_preserves_unrelated_process(self):
        unrelated = subprocess.Popen(["/bin/sleep", "30"])
        try:
            process, peer = self.launch(arguments=("-c", "sleep 30 & wait"))
            self.assertEqual(1, self.frame(peer)[0])
            process.send_signal(signal.SIGTERM)
            self.assertEqual(2, self.frame(peer)[0])
            self.assertEqual((5, errno.EINTR), self.frame(peer)[::2])
            process.communicate(timeout=3)
            self.assertIsNone(unrelated.poll())
        finally:
            unrelated.terminate()
            unrelated.wait(timeout=3)

    def test_unrelated_inheritable_descriptor_is_closed_before_target_exec(self):
        with tempfile.TemporaryFile() as file:
            foreign = fcntl.fcntl(file.fileno(), fcntl.F_DUPFD, 64)
            os.set_inheritable(foreign, True)
            try:
                script = f"import os,sys\ntry: os.fstat({foreign})\nexcept OSError: sys.exit(0)\nsys.exit(9)"
                process, peer = self.launch(sys.executable, ("-c", script), foreign_descriptor=foreign)
                started = self.frame(peer)
                self.assertEqual((2, started[1], 0), self.frame(peer))
                peer.sendall(b"K")
                self.assertEqual((4, started[1], 0), self.frame(peer))
                self.assertEqual(0, process.wait(timeout=3))
            finally:
                os.close(foreign)

    def test_immediate_eof_during_startup_is_bounded(self):
        started = time.monotonic()
        process, peer = self.launch(arguments=("-c", "sleep 30 & wait"))
        peer.close()
        process.communicate(timeout=3)
        self.assertLess(time.monotonic() - started, 3)
