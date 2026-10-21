import java.nio.file.*;
import java.security.*;
import java.security.cert.Certificate;
import org.ofdrw.crypto.OFDEncryptor;
import org.ofdrw.crypto.OFDDecryptor;
import org.ofdrw.crypto.enryptor.UserPasswordEncryptor;
import org.ofdrw.crypto.enryptor.UserCertEncryptor;
import org.ofdrw.crypto.decryptor.UserPasswordDecryptor;
import org.ofdrw.crypto.decryptor.UserCertDecryptor;
import org.ofdrw.crypto.integrity.OFDIntegrity;
import org.ofdrw.crypto.integrity.OFDIntegrityVerifier;
import org.ofdrw.crypto.integrity.GMProtectSigner;
import org.ofdrw.crypto.integrity.GMProtectVerifier;
import org.ofdrw.gm.cert.PEMLoader;
import org.bouncycastle.crypto.params.ECPrivateKeyParameters;
import org.bouncycastle.jcajce.provider.asymmetric.util.ECUtil;

public class R14Java {
  public static void main(String[] a) throws Exception {
    Path source=Paths.get(a[1]), output=Paths.get(a[2]);
    Path certPath=Paths.get(a[3]), keyPath=Paths.get(a[4]);
    Certificate cert=PEMLoader.loadCert(certPath);
    PrivateKey key=PEMLoader.loadPrivateKey(keyPath);
    switch(a[0]) {
      case "enc-pass":
        Path middle=output.resolveSibling(output.getFileName()+".middle.ofd");
        try(OFDEncryptor e=new OFDEncryptor(source,middle)) {
          e.addUser(new UserPasswordEncryptor("alice","first-pass"));
          e.addUser(new UserPasswordEncryptor("bob","second-pass")); e.encrypt();
        }
        try(OFDIntegrity i=new OFDIntegrity(middle,output)) {i.protect(new GMProtectSigner(key,cert));}
        Files.delete(middle); break;
      case "enc-cert":
        Path mid=output.resolveSibling(output.getFileName()+".middle.ofd");
        try(OFDEncryptor e=new OFDEncryptor(source,mid)) {
          e.addUser(new UserCertEncryptor("alice",cert));
          e.addUser(new UserCertEncryptor("bob",cert)); e.encrypt();
        }
        try(OFDIntegrity i=new OFDIntegrity(mid,output)) {i.protect(new GMProtectSigner(key,cert));}
        Files.delete(mid); break;
      case "dec-pass":
        try(OFDDecryptor d=new OFDDecryptor(source,output)) {d.addUser(new UserPasswordDecryptor(a[5],a[6]));d.decrypt();}
        break;
      case "dec-cert":
        try(OFDDecryptor d=new OFDDecryptor(source,output)) {
          d.addUser(new UserCertDecryptor(a[5],(ECPrivateKeyParameters)ECUtil.generatePrivateKeyParameter(key)));d.decrypt();
        } break;
      case "verify":
        System.out.println(new OFDIntegrityVerifier().integrity(source,new GMProtectVerifier()));
        break;
      default: throw new IllegalArgumentException(a[0]);
    }
  }
}
