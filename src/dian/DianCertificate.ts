import fs from "fs";
import { createHash } from "crypto";
import forge from "node-forge";
import { DianCertificateError } from "./errors/DianError.js";

export interface DianCertificateMaterial {
  certificatePem: string;
  privateKeyPem: string;
  certificateBase64: string;
  certificateThumbprint: string;
}

export class DianCertificate {
  constructor(
    private readonly p12Path: string,
    private readonly p12Password: string
  ) {}

  load(): DianCertificateMaterial {
    if (!this.p12Path) {
      throw new DianCertificateError("No se recibió ruta de certificado P12");
    }

    if (!fs.existsSync(this.p12Path)) {
      throw new DianCertificateError("No se encontró el archivo de certificado P12", undefined, {
        p12Path: this.p12Path,
      });
    }

    try {
      const p12Buffer = fs.readFileSync(this.p12Path);
      const p12Asn1 = forge.asn1.fromDer(
        forge.util.createBuffer(p12Buffer.toString("binary"))
      );

      const p12 = forge.pkcs12.pkcs12FromAsn1(p12Asn1, false, this.p12Password);

      const certBags = p12.getBags({ bagType: forge.pki.oids.certBag })[forge.pki.oids.certBag] ?? [];
      if (certBags.length === 0 || !certBags[0].cert) {
        throw new DianCertificateError("El archivo P12 no contiene certificado X509");
      }

      const shroudedKeyBags =
        p12.getBags({ bagType: forge.pki.oids.pkcs8ShroudedKeyBag })[
          forge.pki.oids.pkcs8ShroudedKeyBag
        ] ?? [];
      const keyBags = p12.getBags({ bagType: forge.pki.oids.keyBag })[forge.pki.oids.keyBag] ?? [];

      const privateKeyBag = shroudedKeyBags[0]?.key ? shroudedKeyBags[0] : keyBags[0];
      if (!privateKeyBag?.key) {
        throw new DianCertificateError("El archivo P12 no contiene llave privada");
      }

      const certificatePem = forge.pki.certificateToPem(certBags[0].cert);
      const privateKeyPem = forge.pki.privateKeyToPem(privateKeyBag.key);

      const certDerBytes = forge.asn1
        .toDer(forge.pki.certificateToAsn1(certBags[0].cert))
        .getBytes();
      const certBuffer = Buffer.from(certDerBytes, "binary");
      const certificateBase64 = certBuffer.toString("base64");
      const certificateThumbprint = createHash("sha1")
        .update(certBuffer)
        .digest("hex")
        .toUpperCase();

      return {
        certificatePem,
        privateKeyPem,
        certificateBase64,
        certificateThumbprint,
      };
    } catch (error) {
      if (error instanceof DianCertificateError) {
        throw error;
      }

      throw new DianCertificateError(
        "No se pudo leer o descifrar el certificado P12. Verifica ruta y contraseña.",
        error,
        { p12Path: this.p12Path }
      );
    }
  }
}
