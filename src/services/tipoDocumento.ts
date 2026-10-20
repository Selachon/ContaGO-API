// Clase de documento DIAN (misma lógica que categoria_documento en motor/script_compras.py).
/** Clase de documento DIAN, para decidir con qué tipo de comprobante de Siigo se causa. */
export interface Ci2Categoria {
  clave: string;
  etiqueta: string;
}

const sinTildes = (s: string) => s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();

export function categoriaDocumento(documentType: string, isCreditNote: boolean): Ci2Categoria {
  const t = sinTildes(String(documentType ?? "")).trim();
  if (isCreditNote || t.includes("nota de credito") || t.includes("nota credito")) return { clave: "NC", etiqueta: "Nota crédito" };
  if (t.includes("nota de debito") || t.includes("nota debito")) return { clave: "ND", etiqueta: "Nota débito" };
  if (t.includes("documento soporte")) return { clave: "DS", etiqueta: "Documento soporte" };
  if (t.includes("equivalente") && t.includes("pos")) return { clave: "DEPOS", etiqueta: "Documento equivalente POS" };
  if (t.includes("equivalente") && t.includes("transporte")) return { clave: "DET", etiqueta: "Documento equivalente transporte" };
  if (!t || t.includes("factura")) return { clave: "FE", etiqueta: "Factura electrónica" };
  // Cualquier otro tipo se pregunta con su nombre tal cual lo reporta la DIAN.
  const original = String(documentType).trim();
  return { clave: `OTRO ${original.replace(/[.$]/g, "").toUpperCase()}`, etiqueta: original };
}
