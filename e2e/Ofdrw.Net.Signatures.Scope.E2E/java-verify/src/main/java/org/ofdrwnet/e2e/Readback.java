package org.ofdrwnet.e2e;

import java.nio.charset.StandardCharsets;
import java.nio.file.Paths;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import org.dom4j.DocumentHelper;
import org.ofdrw.core.signatures.Signatures;
import org.ofdrw.core.signatures.sig.Parameter;
import org.ofdrw.core.signatures.sig.Parameters;
import org.ofdrw.reader.OFDReader;

public final class Readback {
    public static void main(String[] args) throws Exception {
        if (args.length != 1) throw new IllegalArgumentException("Expected signed OFD path");
        try (OFDReader reader = new OFDReader(Paths.get(args[0]));
             ZipFile zip = new ZipFile(args[0])) {
            Signatures signatures = reader.getDefaultSignatures();
            if (signatures == null) throw new IllegalStateException("Java reader found no Signatures.xml");
            List<org.ofdrw.core.signatures.Signature> records = signatures.getSignatures();
            if (records.size() != 2 || !"1".equals(records.get(1).getRelative()))
                throw new IllegalStateException("Java model rejected the Relative relation");
            String location = "Doc_0/Signs/Sign_1/Signature.xml";
            ZipEntry entry = zip.getEntry(location);
            if (entry == null) throw new IllegalStateException("Missing " + location);
            byte[] bytes = new byte[(int) entry.getSize()];
            try (java.io.InputStream stream = zip.getInputStream(entry)) {
                int position = 0;
                while (position < bytes.length) {
                    int count = stream.read(bytes, position, bytes.length - position);
                    if (count < 0) throw new IllegalStateException("Short Signature.xml");
                    position += count;
                }
            }
            org.ofdrw.core.signatures.sig.Signature signature =
                new org.ofdrw.core.signatures.sig.Signature(
                    DocumentHelper.parseText(new String(bytes, StandardCharsets.UTF_8)).getRootElement());
            Parameters parameters = signature.getSignedInfo().getParameters();
            Parameter purpose = parameters == null ? null : parameters.getParameter("purpose");
            if (purpose == null || !"package-consumer".equals(purpose.getValue()) ||
                    !"text".equals(purpose.getType()))
                throw new IllegalStateException("Java model rejected Parameters");
            System.out.println("Java OFDRW 2.4.0 read two signatures, Relative=1, and purpose Parameter");
        }
    }
}
