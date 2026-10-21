// Cross-implementation probe. Compile against the pinned org.ofdrw:ofdrw-sign:2.3.9
// runtime dependencies; inputs and outputs are supplied by the caller.
import java.io.ByteArrayInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.PrivateKey;
import java.security.cert.Certificate;
import org.ofdrw.core.signatures.SigType;
import org.ofdrw.gm.cert.PKCS12Tools;
import org.ofdrw.gm.ses.v5.SESeal;
import org.ofdrw.sign.signContainer.GBT35275DSContainer;
import org.ofdrw.sign.signContainer.SESV5Container;
import org.ofdrw.sign.verify.container.GBT35275ValidateContainer;
import org.ofdrw.sign.verify.container.SESV5ValidateContainer;

public final class JavaInterop {
    private static final String SM2_SM3 = "1.2.156.10197.1.501";
    private static final String PROPERTY = "/Doc_0/Signs/Sign_0/Signature.xml";

    public static void main(String[] args) throws Exception {
        if (args.length == 5 && args[0].equals("produce")) {
            Path p12 = Path.of(args[1]);
            Path sealPath = Path.of(args[2]);
            byte[] xml = Files.readAllBytes(Path.of(args[3]));
            Path output = Path.of(args[4]);
            Files.createDirectories(output);
            // These aliases/passwords belong solely to the publicly available upstream test fixture.
            PrivateKey key = PKCS12Tools.ReadPrvKey(p12, "private", "777777");
            Certificate certificate = PKCS12Tools.ReadUserCert(p12, "private", "777777");
            SESeal seal = SESeal.getInstance(Files.readAllBytes(sealPath));
            byte[] v5 = new SESV5Container(key, seal, certificate)
                .sign(new ByteArrayInputStream(xml), PROPERTY);
            byte[] gbt = new GBT35275DSContainer(certificate, key)
                .sign(new ByteArrayInputStream(xml), PROPERTY);
            Files.write(output.resolve("java-v5.dat"), v5);
            Files.write(output.resolve("java-gbt35275.dat"), gbt);
            Files.write(output.resolve("java-signer.cer"), certificate.getEncoded());
            Files.write(output.resolve("java-maker.cer"), seal.getCert().getOctets());
            Files.write(output.resolve("java-seal.esl"), seal.getEncoded("DER"));
            System.out.println("Java produced SES V5 and GB/T 35275 containers");
            return;
        }
        if (args.length == 4 && args[0].equals("verify")) {
            byte[] xml = Files.readAllBytes(Path.of(args[1]));
            byte[] v5 = Files.readAllBytes(Path.of(args[2]));
            byte[] gbt = Files.readAllBytes(Path.of(args[3]));
            new SESV5ValidateContainer().validate(SigType.Seal, SM2_SM3, xml, v5);
            new GBT35275ValidateContainer().validate(SigType.Sign, SM2_SM3, xml, gbt);
            System.out.println("Java verified SES V5 and GB/T 35275 containers");
            return;
        }
        throw new IllegalArgumentException("Expected produce p12 seal xml outdir or verify xml v5 gbt");
    }
}
