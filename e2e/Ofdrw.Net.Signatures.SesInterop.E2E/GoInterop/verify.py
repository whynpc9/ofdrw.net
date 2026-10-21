"""Independent GB/T 35275 narrow-profile probe; use task-local test samples.

Run with: python3 verify.py net-Signature.xml net-gbt35275.dat java-Signature.xml java-gbt35275.dat
Go parses the container and checks its attached SM3 digest. OpenSSL verifies
the extracted SM2 signatures. Neither step establishes certificate trust.
"""

import pathlib
import subprocess
import sys


def main() -> None:
    if len(sys.argv) != 5:
        raise SystemExit("expected .NET XML/container and Java XML/container pairs")
    paths = [str(pathlib.Path(name).resolve()) for name in sys.argv[1:]]
    subprocess.run(["go", "run", ".", *paths], cwd=pathlib.Path(__file__).parent, check=True)
    for name in (paths[1], paths[3]):
        pubkey = name + ".pub.pem"
        subprocess.run(
            ["openssl", "x509", "-inform", "DER", "-in", name + ".cert",
             "-pubkey", "-noout", "-out", pubkey],
            check=True,
        )
        subprocess.run(
            ["openssl", "pkeyutl", "-verify", "-pubin", "-inkey", pubkey,
             "-rawin", "-digest", "sm3", "-in", name + ".content",
             "-sigfile", name + ".sig", "-pkeyopt", "distid:1234567812345678"],
            check=True,
        )


if __name__ == "__main__":
    main()
