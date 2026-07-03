import * as XLSX from "xlsx";
import JSZip from "jszip";

export type DianTaxRow = {
  iva: number;
  ica: number;
  ic: number;
  inc: number;
  timbre: number;
  bolsas: number;
  inCarbono: number;
  inCombustibles: number;
  icDatos: number;
  icl: number;
  inpp: number;
  ibua: number;
  icui: number;
  reteIva: number;
  reteRenta: number;
  reteIca: number;
};

// Si el buffer es un ZIP que contiene un .xlsx adentro, lo extrae y devuelve ese buffer.
// El listado del portal DIAN a veces viene empaquetado en ZIP.
async function resolveBuffer(buffer: Buffer): Promise<Buffer> {
  const isZip = buffer[0] === 0x50 && buffer[1] === 0x4b;
  if (!isZip) return buffer;

  // Intentar abrir como ZIP contenedor (no como XLSX, que también es PK internamente)
  try {
    const zip = await JSZip.loadAsync(buffer);
    const xlsxEntry = Object.values(zip.files).find(
      (f) => !f.dir && /\.(xlsx|xls|xlsm)$/i.test(f.name)
    );
    if (xlsxEntry) {
      const extracted = await xlsxEntry.async("nodebuffer");
      return extracted as Buffer;
    }
  } catch {
    // No era un ZIP contenedor válido — tratar directamente como xlsx
  }
  return buffer;
}

export async function parseDianListing(buffer: Buffer): Promise<Map<string, DianTaxRow>> {
  const xlsxBuf = await resolveBuffer(buffer);
  const wb = XLSX.read(xlsxBuf, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) return new Map();
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 }) as unknown[][];

  if (!rows.length) return new Map();

  const headers = rows[0] as string[];
  const idx = (name: string) => headers.findIndex((h) => String(h).trim() === name);
  const cufeIdx = headers.findIndex((h) => String(h).includes("CUFE"));
  if (cufeIdx < 0) return new Map();

  const colMap = {
    iva: idx("IVA"),
    ica: idx("ICA"),
    ic: idx("IC"),
    inc: idx("INC"),
    timbre: idx("Timbre"),
    bolsas: idx("INC Bolsas"),
    inCarbono: idx("IN Carbono"),
    inCombustibles: idx("IN Combustibles"),
    icDatos: idx("IC Datos"),
    icl: idx("ICL"),
    inpp: idx("INPP"),
    ibua: idx("IBUA"),
    icui: idx("ICUI"),
    reteIva: idx("Rete IVA"),
    reteRenta: idx("Rete Renta"),
    reteIca: idx("Rete ICA"),
  };

  const map = new Map<string, DianTaxRow>();
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] as unknown[];
    const cufe = String(row[cufeIdx] || "").trim();
    if (!cufe || cufe.length < 64) continue;

    const n = (col: number) => (col >= 0 ? Number(row[col]) || 0 : 0);
    map.set(cufe, {
      iva: n(colMap.iva),
      ica: n(colMap.ica),
      ic: n(colMap.ic),
      inc: n(colMap.inc),
      timbre: n(colMap.timbre),
      bolsas: n(colMap.bolsas),
      inCarbono: n(colMap.inCarbono),
      inCombustibles: n(colMap.inCombustibles),
      icDatos: n(colMap.icDatos),
      icl: n(colMap.icl),
      inpp: n(colMap.inpp),
      ibua: n(colMap.ibua),
      icui: n(colMap.icui),
      reteIva: n(colMap.reteIva),
      reteRenta: n(colMap.reteRenta),
      reteIca: n(colMap.reteIca),
    });
  }
  return map;
}
