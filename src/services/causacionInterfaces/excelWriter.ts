/**
 * Escritura de las interfaces de importación de Siigo a archivos .xlsx.
 */
import ExcelJS from "exceljs";
import {
  COLUMNAS_SIIGO,
  COLUMNAS_INFORME,
  type FilaSiigo,
  type FilaInforme,
} from "./types.js";

/** Vuelca filas a una hoja "Sheet1" y devuelve el .xlsx como Buffer. */
async function escribirHoja(
  columnas: readonly string[],
  filas: Array<Record<string, unknown>>
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");
  ws.addRow([...columnas]);
  for (const fila of filas) {
    ws.addRow(
      columnas.map((col) => {
        const v = fila[col];
        return v === undefined || v === null ? "" : v;
      })
    );
  }
  const buffer = await wb.xlsx.writeBuffer();
  return buffer as unknown as Buffer;
}

/** Genera el archivo de la interfaz contable Siigo. */
export function escribirInterfazSiigo(filas: FilaSiigo[]): Promise<Buffer> {
  return escribirHoja(COLUMNAS_SIIGO, filas);
}

/** Genera el archivo del informe auxiliar de retenciones. */
export function escribirInformeRetenciones(filas: FilaInforme[]): Promise<Buffer> {
  return escribirHoja(COLUMNAS_INFORME, filas);
}
