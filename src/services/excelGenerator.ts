import JSZip from "jszip";
import fs from "fs";
import path from "path";
import type { InvoiceData } from "../types/dianExcel.js";

function colLetter(n: number): string {
  let r = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    r = String.fromCharCode(65 + rem) + r;
    n = Math.floor((n - 1) / 26);
  }
  return r;
}

function xmlEsc(s: string): string {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function sortInvoicesByDate(invoices: InvoiceData[]): InvoiceData[] {
  return [...invoices].sort((a, b) => {
    const da = a.issueDateISO || "9999-12-31";
    const db = b.issueDateISO || "9999-12-31";
    return da.localeCompare(db);
  });
}

function normalizeNit(nit: string | null | undefined): string {
  const raw = (nit || "").trim();
  if (!raw) return "N/A";
  return raw.replace(/[^0-9A-Za-z]/g, "").toUpperCase() || "N/A";
}

function resolveTemplatePath(): string {
  const configured = process.env.DIAN_EXCEL_TEMPLATE_PATH?.trim() || "templates/dian-excel-template.xlsx";
  const candidates: string[] = [];
  if (path.isAbsolute(configured)) candidates.push(configured);
  else candidates.push(path.join(process.cwd(), configured));
  candidates.push(path.join(process.cwd(), "templates", "dian-excel-template.xlsx"));
  candidates.push(path.join(path.dirname(new URL(import.meta.url).pathname), "../../templates/dian-excel-template.xlsx"));
  candidates.push("/app/templates/dian-excel-template.xlsx");
  for (const c of candidates) {
    const n = path.normalize(c);
    if (fs.existsSync(n)) return n;
  }
  throw new Error(`No se encontró la plantilla Excel. Revisa DIAN_EXCEL_TEMPLATE_PATH. cwd=${process.cwd()} configurado=${configured}`);
}

// Build an inline-string cell. Returns empty string if value is null/undefined/empty.
function txtCell(ref: string, value: string | null | undefined, style?: number): string {
  if (value == null || value === "") return "";
  const s = style !== undefined ? ` s="${style}"` : "";
  return `<c r="${ref}"${s} t="inlineStr"><is><t>${xmlEsc(String(value))}</t></is></c>`;
}

// Build a numeric cell. Returns empty string if value is null/undefined/NaN.
function numCell(ref: string, value: number | null | undefined, style?: number): string {
  if (value == null || !isFinite(value)) return "";
  const s = style !== undefined ? ` s="${style}"` : "";
  return `<c r="${ref}"${s}><v>${value}</v></c>`;
}

// Extract rows 1 and 2 from a sheet's sheetData XML verbatim.
function extractTemplateRows(sheetXml: string): string {
  const sdStart = sheetXml.indexOf("<sheetData>") + "<sheetData>".length;
  const sdEnd = sheetXml.indexOf("</sheetData>");
  const sd = sheetXml.substring(sdStart, sdEnd);
  const r2idx = sd.indexOf('<row r="2"');
  if (r2idx === -1) return sd.substring(0, sd.indexOf("</row>") + 6);
  const r2end = sd.indexOf("</row>", r2idx) + 6;
  return sd.substring(0, r2end);
}

// Replace the sheetData section in a sheet XML and optionally inject a <hyperlinks> block.
function patchSheetXml(
  sheetXml: string,
  newSheetData: string,
  hyperlinksXml: string | null,
  lastRow: number,
  endCol: string
): string {
  // Update <dimension>
  let out = sheetXml.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:${endCol}${lastRow}"/>`);

  // Replace sheetData block
  const sdStart = out.indexOf("<sheetData>");
  const sdEnd = out.indexOf("</sheetData>") + "</sheetData>".length;
  out = out.substring(0, sdStart) + newSheetData + out.substring(sdEnd);

  // Insert <hyperlinks> before <pageMargins> (OOXML spec order: position 19, before margins at 21)
  if (hyperlinksXml) {
    out = out.replace("<pageMargins", hyperlinksXml + "<pageMargins");
  }

  return out;
}

// Update the ref= attributes in a table XML to point to the new last row.
function patchTableRef(tableXml: string, endCol: string, lastRow: number): string {
  const newRef = `A2:${endCol}${Math.max(3, lastRow)}`;
  return tableXml.replace(/ref="A2:[A-Z]+\d+"/g, `ref="${newRef}"`);
}

// Style indices added to styles.xml (appended after the 7 existing xfs, indices 0-6)
const STYLE_CURRENCY = 7; // numFmtId 164 → #,##0.00
const STYLE_PERCENT = 8;  // numFmtId 165 → 0.00"%"

export async function generateExcelFile(
  invoices: InvoiceData[],
  outputPath: string,
  includeDriveColumn: boolean,
  isSentDocuments: boolean = false
): Promise<void> {
  const templatePath = resolveTemplatePath();
  const zip = await JSZip.loadAsync(fs.readFileSync(templatePath));

  // ── 1. Patch styles.xml ──────────────────────────────────────────────────
  let stylesXml = await zip.file("xl/styles.xml")!.async("string");

  // Inject <numFmts> before <fonts> (template has no numFmts section)
  stylesXml = stylesXml.replace(
    "<fonts",
    '<numFmts count="2">' +
      '<numFmt numFmtId="164" formatCode="#,##0.00"/>' +
      '<numFmt numFmtId="165" formatCode="0.00&quot;%&quot;"/>' +
      "</numFmts><fonts"
  );

  // Append 2 new xf entries to cellXfs and update the count
  stylesXml = stylesXml
    .replace('<cellXfs count="7">', '<cellXfs count="9">')
    .replace(
      "</cellXfs>",
      '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
        '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
        "</cellXfs>"
    );

  zip.file("xl/styles.xml", stylesXml);

  // ── 2. Sort invoices ─────────────────────────────────────────────────────
  const sorted = sortInvoicesByDate(invoices);

  // ── 3. Sheet 1 — Facturas DIAN ───────────────────────────────────────────
  // Fixed column positions (1-indexed, matching template table columns A–U)
  // A=ID  B=Tipo doc  C=Num Factura  D=NIT Emisor  E=Razon Social
  // F=Fecha  G=Concepto  H=Forma pago  I=Subtotal  J=Desc detalle
  // K=Recargo detalle  L=IVA  M=INC  N=Bolsas  O=ICUI  P=IC
  // Q=Desc Global  R=Recargo Global  S=Valor total  T=Enlace  U=CUFE

  let sheet1Rows: string[] = [];
  let sheet1DriveRels: Array<{ id: string; url: string }> = [];
  let sheet1Hyperlinks: string[] = [];
  let rIdCounter = 2; // rId1 is already used by the table relationship

  const docNumberToRow = new Map<string, number>();

  sorted.forEach((inv, idx) => {
    const rowNum = 3 + idx;
    docNumberToRow.set(inv.docNumber, rowNum);

    const partyNit = isSentDocuments ? inv.receiverNit : inv.issuerNit;
    const partyName = isSentDocuments ? inv.receiverName : inv.issuerName;

    const taxes: Record<string, number> = {};
    for (const t of inv.taxes || []) if (t.taxName) taxes[t.taxName] = t.amount;

    let c = "";
    c += numCell(`A${rowNum}`, idx + 1);
    c += txtCell(`B${rowNum}`, inv.documentType);
    c += txtCell(`C${rowNum}`, inv.docNumber);
    c += txtCell(`D${rowNum}`, partyNit);
    c += txtCell(`E${rowNum}`, partyName);
    c += txtCell(`F${rowNum}`, inv.issueDate);
    c += txtCell(`G${rowNum}`, inv.concepts);
    c += txtCell(`H${rowNum}`, inv.paymentMethod || "N/A");
    c += numCell(`I${rowNum}`, typeof inv.subtotal === "number" ? inv.subtotal : 0, STYLE_CURRENCY);
    c += numCell(`J${rowNum}`, inv.discount || 0, STYLE_CURRENCY);
    c += numCell(`K${rowNum}`, inv.surcharge || 0, STYLE_CURRENCY);
    c += numCell(`L${rowNum}`, taxes["IVA"] ?? 0, STYLE_CURRENCY);
    c += numCell(`M${rowNum}`, taxes["INC"] ?? 0, STYLE_CURRENCY);
    c += numCell(`N${rowNum}`, taxes["Bolsas"] ?? 0, STYLE_CURRENCY);
    c += numCell(`O${rowNum}`, taxes["ICUI"] ?? 0, STYLE_CURRENCY);
    c += numCell(`P${rowNum}`, taxes["IC"] ?? 0, STYLE_CURRENCY);
    c += numCell(`Q${rowNum}`, 0, STYLE_CURRENCY);
    c += numCell(`R${rowNum}`, 0, STYLE_CURRENCY);
    c += numCell(`S${rowNum}`, typeof inv.total === "number" ? inv.total : 0, STYLE_CURRENCY);

    if (includeDriveColumn && inv.driveUrl && !inv.driveUrl.includes("ERROR")) {
      const relId = `rId${rIdCounter++}`;
      sheet1DriveRels.push({ id: relId, url: inv.driveUrl });
      sheet1Hyperlinks.push(`<hyperlink ref="T${rowNum}" r:id="${relId}" display="Ver factura"/>`);
      c += txtCell(`T${rowNum}`, "Ver factura");
    }

    if (includeDriveColumn) c += txtCell(`U${rowNum}`, inv.cufe);
    else c += txtCell(`T${rowNum}`, inv.cufe);

    sheet1Rows.push(`<row r="${rowNum}">${c}</row>`);
  });

  const sheet1LastRow = Math.max(3, 2 + sorted.length);

  let sheet1Xml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
  const s1TemplateRows = extractTemplateRows(sheet1Xml);
  const s1SheetData = `<sheetData>${s1TemplateRows}${sheet1Rows.join("")}</sheetData>`;
  const s1Hyperlinks =
    sheet1Hyperlinks.length > 0
      ? `<hyperlinks>${sheet1Hyperlinks.join("")}</hyperlinks>`
      : null;
  sheet1Xml = patchSheetXml(sheet1Xml, s1SheetData, s1Hyperlinks, sheet1LastRow, includeDriveColumn ? "U" : "T");
  zip.file("xl/worksheets/sheet1.xml", sheet1Xml);

  if (sheet1DriveRels.length > 0) {
    let rels1 = await zip.file("xl/worksheets/_rels/sheet1.xml.rels")!.async("string");
    const newRels = sheet1DriveRels
      .map(
        (r) =>
          `<Relationship Id="${r.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xmlEsc(r.url)}" TargetMode="External"/>`
      )
      .join("");
    rels1 = rels1.replace("</Relationships>", newRels + "</Relationships>");
    zip.file("xl/worksheets/_rels/sheet1.xml.rels", rels1);
  }

  // ── 4. Sheet 2 — Detallado ───────────────────────────────────────────────
  // A=Item  B=Num Factura  C=Concepto  D=Cantidad  E=Base impuesto
  // F=Desc detalle  G=Recargo detalle  H=IVA  I=%IVA  J=INC  K=%INC
  // L=Bolsas  M=%Bolsas  N=ICUI  O=%ICUI  P=IC  Q=%IC  R=Precio unitario

  let sheet2Rows: string[] = [];
  let sheet2Hyperlinks: string[] = [];
  let detRow = 3;

  for (const inv of sorted) {
    const mainRow = docNumberToRow.get(inv.docNumber) || 3;
    for (const li of inv.lineItems || []) {
      const rowNum = detRow++;
      const td: Record<string, { amount: number; percent: number }> = {};
      for (const t of li.taxes || []) if (t.taxName) td[t.taxName] = { amount: t.amount, percent: t.percent };

      sheet2Hyperlinks.push(
        `<hyperlink ref="B${rowNum}" location="'Facturas DIAN'!C${mainRow}" display="${xmlEsc(inv.docNumber)}"/>`
      );

      const totalTax = (li.taxes || []).reduce((s, t) => s + t.amount, 0);

      let c = "";
      c += numCell(`A${rowNum}`, li.lineNumber);
      c += txtCell(`B${rowNum}`, inv.docNumber);
      c += txtCell(`C${rowNum}`, li.description);
      c += numCell(`D${rowNum}`, li.quantity);
      c += numCell(`E${rowNum}`, li.totalUnitPrice, STYLE_CURRENCY);
      c += numCell(`F${rowNum}`, li.discount, STYLE_CURRENCY);
      c += numCell(`G${rowNum}`, li.surcharge, STYLE_CURRENCY);
      c += numCell(`H${rowNum}`, td["IVA"]?.amount ?? 0, STYLE_CURRENCY);
      c += numCell(`I${rowNum}`, td["IVA"]?.percent ?? 0, STYLE_PERCENT);
      c += numCell(`J${rowNum}`, td["INC"]?.amount ?? 0, STYLE_CURRENCY);
      c += numCell(`K${rowNum}`, td["INC"]?.percent ?? 0, STYLE_PERCENT);
      c += numCell(`L${rowNum}`, td["Bolsas"]?.amount ?? 0, STYLE_CURRENCY);
      c += numCell(`M${rowNum}`, td["Bolsas"]?.percent ?? 0, STYLE_PERCENT);
      c += numCell(`N${rowNum}`, td["ICUI"]?.amount ?? 0, STYLE_CURRENCY);
      c += numCell(`O${rowNum}`, td["ICUI"]?.percent ?? 0, STYLE_PERCENT);
      c += numCell(`P${rowNum}`, td["IC"]?.amount ?? 0, STYLE_CURRENCY);
      c += numCell(`Q${rowNum}`, td["IC"]?.percent ?? 0, STYLE_PERCENT);
      c += numCell(`R${rowNum}`, li.totalUnitPrice + totalTax, STYLE_CURRENCY);

      sheet2Rows.push(`<row r="${rowNum}">${c}</row>`);
    }
  }

  const sheet2LastRow = Math.max(3, detRow - 1);

  let sheet2Xml = await zip.file("xl/worksheets/sheet2.xml")!.async("string");
  const s2TemplateRows = extractTemplateRows(sheet2Xml);
  const s2SheetData = `<sheetData>${s2TemplateRows}${sheet2Rows.join("")}</sheetData>`;
  const s2Hyperlinks =
    sheet2Hyperlinks.length > 0
      ? `<hyperlinks>${sheet2Hyperlinks.join("")}</hyperlinks>`
      : null;
  sheet2Xml = patchSheetXml(sheet2Xml, s2SheetData, s2Hyperlinks, sheet2LastRow, "R");
  zip.file("xl/worksheets/sheet2.xml", sheet2Xml);

  // ── 5. Sheet 3 — Datos de terceros ───────────────────────────────────────
  // A=NIT  B=Razon Social  C=Nombre Comercial  D=Resp Tributaria
  // E=Pais  F=Departamento  G=Municipio/Ciudad  H=Direccion  I=Telefono  J=Correo

  let sheet3Rows: string[] = [];

  type ThirdPartyRow = {
    nit: string;
    name: string;
    commercial: string;
    taxResp: string;
    country: string;
    dept: string;
    city: string;
    addr: string;
    phone: string;
    email: string;
  };

  const thirdPartiesByNit = new Map<string, ThirdPartyRow>();

  for (const inv of sorted) {
    const p: ThirdPartyRow = isSentDocuments
      ? {
          nit: inv.receiverNit,
          name: inv.receiverName,
          commercial: inv.receiverCommercialName || "N/A",
          taxResp: inv.receiverTaxResponsibility || "N/A",
          country: inv.receiverCountry || "N/A",
          dept: inv.receiverDepartment || "N/A",
          city: inv.receiverCity || "N/A",
          addr: inv.receiverAddress || "N/A",
          phone: inv.receiverPhone || "N/A",
          email: inv.receiverEmail || "N/A",
        }
      : {
          nit: inv.issuerNit,
          name: inv.issuerName,
          commercial: inv.issuerCommercialName || "N/A",
          taxResp: inv.issuerTaxResponsibility || "N/A",
          country: inv.issuerCountry || "N/A",
          dept: inv.issuerDepartment || "N/A",
          city: inv.issuerCity || "N/A",
          addr: inv.issuerAddress || "N/A",
          phone: inv.issuerPhone || "N/A",
          email: inv.issuerEmail || "N/A",
        };

    const nitKey = normalizeNit(p.nit);
    if (!thirdPartiesByNit.has(nitKey)) {
      thirdPartiesByNit.set(nitKey, p);
    }
  }

  const uniqueThirdParties = Array.from(thirdPartiesByNit.values());

  uniqueThirdParties.forEach((p, idx) => {
    const rowNum = 3 + idx;
    let c = "";
    c += txtCell(`A${rowNum}`, p.nit);
    c += txtCell(`B${rowNum}`, p.name);
    c += txtCell(`C${rowNum}`, p.commercial);
    c += txtCell(`D${rowNum}`, p.taxResp);
    c += txtCell(`E${rowNum}`, p.country);
    c += txtCell(`F${rowNum}`, p.dept);
    c += txtCell(`G${rowNum}`, p.city);
    c += txtCell(`H${rowNum}`, p.addr);
    c += txtCell(`I${rowNum}`, p.phone);
    c += txtCell(`J${rowNum}`, p.email);
    sheet3Rows.push(`<row r="${rowNum}">${c}</row>`);
  });

  const sheet3LastRow = Math.max(3, 2 + uniqueThirdParties.length);

  let sheet3Xml = await zip.file("xl/worksheets/sheet3.xml")!.async("string");
  const s3TemplateRows = extractTemplateRows(sheet3Xml);
  const s3SheetData = `<sheetData>${s3TemplateRows}${sheet3Rows.join("")}</sheetData>`;
  sheet3Xml = patchSheetXml(sheet3Xml, s3SheetData, null, sheet3LastRow, "J");
  zip.file("xl/worksheets/sheet3.xml", sheet3Xml);

  // ── 6. Patch table refs ──────────────────────────────────────────────────
  zip.file("xl/tables/table1.xml", patchTableRef(await zip.file("xl/tables/table1.xml")!.async("string"), includeDriveColumn ? "U" : "T", sheet1LastRow));
  zip.file("xl/tables/table2.xml", patchTableRef(await zip.file("xl/tables/table2.xml")!.async("string"), "R", sheet2LastRow));
  zip.file("xl/tables/table3.xml", patchTableRef(await zip.file("xl/tables/table3.xml")!.async("string"), "J", sheet3LastRow));

  // ── 7. Write output ──────────────────────────────────────────────────────
  const buf = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  fs.writeFileSync(outputPath, buf);
}

export function generateExcelFilename(startDate?: string, endDate?: string, prefix: string = "Facturas DIAN"): string {
  const fmt = (d: string): string => {
    const [year, month, day] = d.split("-");
    const months = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
    return `${months[parseInt(month) - 1]} ${parseInt(day)} ${year}`;
  };
  const start = startDate ? fmt(startDate) : "Inicio";
  const end = endDate ? fmt(endDate) : "Fin";
  return `${prefix} ${start} - ${end}.xlsx`;
}
