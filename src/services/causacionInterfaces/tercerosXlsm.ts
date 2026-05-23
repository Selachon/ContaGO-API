/**
 * Escritura del archivo de terceros sobre la plantilla `.xlsm` de Siigo.
 *
 * exceljs no preserva el proyecto VBA al escribir, así que se manipula el
 * `.xlsm` a nivel de ZIP con jszip: se rellena la hoja "Plantilla", se elimina
 * la hoja "Paises" y se conserva `vbaProject.bin` y el resto del paquete.
 *
 * Equivale a `escribir_en_plantilla` de script_terceros.py (load_workbook con
 * keep_vba=True, escribir filas, remover hoja "Paises").
 */
import JSZip from "jszip";
import { CausacionInterfacesError, COLUMNAS_TERCEROS, type FilaTercero } from "./types.js";
import { cargarLibro } from "./excelReader.js";
import { normalizarTexto } from "./normalize.js";

/** Convierte un índice de columna (1-based) a letra de columna Excel. */
function columnaLetra(n: number): string {
  let s = "";
  let v = n;
  while (v > 0) {
    const resto = (v - 1) % 26;
    s = String.fromCharCode(65 + resto) + s;
    v = Math.floor((v - 1) / 26);
  }
  return s;
}

/** Escapa texto para contenido XML. */
function escaparXml(texto: string): string {
  return texto.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Construye el XML de una fila de la hoja, con celdas como cadenas en línea. */
function filaXml(r: number, celdas: Array<{ col: number; valor: string }>, nCols: number): string {
  const cuerpo = celdas
    .map(
      (c) =>
        `<c r="${columnaLetra(c.col)}${r}" t="inlineStr"><is><t xml:space="preserve">` +
        `${escaparXml(c.valor)}</t></is></c>`
    )
    .join("");
  return `<row r="${r}" spans="1:${nCols}">${cuerpo}</row>`;
}

/** Elimina la primera coincidencia de un patrón en un texto XML. */
function quitar(xml: string, patron: RegExp): string {
  return xml.replace(patron, "");
}

/**
 * Rellena la plantilla `.xlsm` con los terceros y devuelve el archivo final.
 * Conserva las macros (vbaProject.bin) y elimina la hoja "Paises".
 */
export async function escribirTercerosXlsm(
  plantilla: Buffer,
  terceros: FilaTercero[]
): Promise<Buffer> {
  // ---- 1. Encabezado real de la hoja "Plantilla" ---------------------------
  const wb = await cargarLibro(plantilla);
  const hoja = wb.getWorksheet("Plantilla");
  if (!hoja) {
    throw new CausacionInterfacesError(
      'La plantilla de terceros no contiene la hoja "Plantilla".',
      422,
      "plantilla_sin_hoja"
    );
  }
  const nCols = Math.max(hoja.columnCount, COLUMNAS_TERCEROS.length);
  const encabezados: string[] = [];
  const filaEncabezado = hoja.getRow(1);
  for (let c = 1; c <= nCols; c++) {
    const v = filaEncabezado.getCell(c).value;
    encabezados.push(v === null || v === undefined ? "" : String(v).trim());
  }
  const columnaPorNombre = new Map<string, number>();
  encabezados.forEach((h, i) => {
    if (h) columnaPorNombre.set(normalizarTexto(h), i + 1);
  });

  // ---- 2. Construcción del nuevo <sheetData> -------------------------------
  const filas: string[] = [];
  filas.push(
    filaXml(
      1,
      encabezados.map((h, i) => ({ col: i + 1, valor: h })).filter((c) => c.valor !== ""),
      nCols
    )
  );
  terceros.forEach((tercero, idx) => {
    const celdas: Array<{ col: number; valor: string }> = [];
    for (const clave of COLUMNAS_TERCEROS) {
      const valor = tercero[clave];
      if (valor === undefined || valor === "") continue;
      const col = columnaPorNombre.get(normalizarTexto(clave));
      if (col) celdas.push({ col, valor: String(valor) });
    }
    filas.push(filaXml(idx + 2, celdas, nCols));
  });
  const sheetData = `<sheetData>${filas.join("")}</sheetData>`;
  const dimension = `<dimension ref="A1:${columnaLetra(nCols)}${terceros.length + 1}"/>`;

  // ---- 3. Cirugía sobre el paquete .xlsm -----------------------------------
  const zip = await JSZip.loadAsync(plantilla);

  const sheet1Path = "xl/worksheets/sheet1.xml";
  const sheet1 = await zip.file(sheet1Path)?.async("string");
  if (!sheet1) {
    throw new CausacionInterfacesError(
      "La plantilla de terceros tiene una estructura inesperada.",
      422,
      "plantilla_estructura_invalida"
    );
  }
  let nuevoSheet1 = sheet1.replace(/<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/, sheetData);
  nuevoSheet1 = nuevoSheet1.replace(/<dimension ref="[^"]*"\s*\/>/, dimension);
  zip.file(sheet1Path, nuevoSheet1);

  // Elimina la hoja "Paises" del libro.
  const workbook = await zip.file("xl/workbook.xml")?.async("string");
  if (workbook) {
    zip.file("xl/workbook.xml", quitar(workbook, /<sheet name="Paises"[^>]*\/>/));
  }
  const rels = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (rels) {
    let nuevoRels = quitar(rels, /<Relationship [^>]*Target="worksheets\/sheet2\.xml"[^>]*\/>/);
    nuevoRels = quitar(nuevoRels, /<Relationship [^>]*Target="calcChain\.xml"[^>]*\/>/);
    zip.file("xl/_rels/workbook.xml.rels", nuevoRels);
  }
  const contentTypes = await zip.file("[Content_Types].xml")?.async("string");
  if (contentTypes) {
    let nuevoCt = quitar(
      contentTypes,
      /<Override PartName="\/xl\/worksheets\/sheet2\.xml"[^>]*\/>/
    );
    nuevoCt = quitar(nuevoCt, /<Override PartName="\/xl\/calcChain\.xml"[^>]*\/>/);
    zip.file("[Content_Types].xml", nuevoCt);
  }

  // Elimina archivos huérfanos: la hoja "Paises" y la cadena de cálculo
  // (calcChain queda obsoleta al sustituir las fórmulas por valores).
  zip.remove("xl/worksheets/sheet2.xml");
  if (zip.file("xl/calcChain.xml")) zip.remove("xl/calcChain.xml");

  const salida = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    mimeType: "application/vnd.ms-excel.sheet.macroEnabled.12",
  });
  return salida as Buffer;
}
