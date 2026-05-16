import { describe, it } from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import { extractThirdPartyCufesFromExcel } from "./dianThirdParties.js";

async function buildDianReport(rows: Array<Record<string, string>>): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Reporte");

  worksheet.addRow([
    "Grupo",
    "CUFE",
    "Tipo de documento",
    "Nombre Emisor",
    "NIT Emisor",
    "Nombre Receptor",
    "NIT Receptor",
  ]);

  for (const row of rows) {
    worksheet.addRow([
      row.grupo,
      row.cufe,
      row.tipoDocumento,
      row.nombreEmisor,
      row.nitEmisor,
      row.nombreReceptor,
      row.nitReceptor,
    ]);
  }

  return Buffer.from(await workbook.xlsx.writeBuffer());
}

describe("extractThirdPartyCufesFromExcel", () => {
  it("does not collapse sent support documents under the company NIT", async () => {
    const buffer = await buildDianReport([
      {
        grupo: "Emitidos",
        cufe: "ABCDEF00000000000001",
        tipoDocumento: "Documento soporte",
        nombreEmisor: "EMPRESA SAS",
        nitEmisor: "900123456",
        nombreReceptor: "PROVEEDOR UNO",
        nitReceptor: "111111111",
      },
      {
        grupo: "Emitidos",
        cufe: "ABCDEF00000000000002",
        tipoDocumento: "Documento soporte",
        nombreEmisor: "EMPRESA SAS",
        nitEmisor: "900123456",
        nombreReceptor: "PROVEEDOR DOS",
        nitReceptor: "222222222",
      },
    ]);

    const { cufesByNit, totalCount } = await extractThirdPartyCufesFromExcel(buffer, "900123456");

    assert.equal(totalCount, 2);
    assert.deepEqual(cufesByNit, {
      "111111111": { cufe: "ABCDEF00000000000001", direction: "sent" },
      "222222222": { cufe: "ABCDEF00000000000002", direction: "sent" },
    });
  });

  it("keeps the standard sent invoice receiver rule", async () => {
    const buffer = await buildDianReport([
      {
        grupo: "Emitidos",
        cufe: "ABCDEF10000000000001",
        tipoDocumento: "Factura electrónica",
        nombreEmisor: "EMPRESA SAS",
        nitEmisor: "900123456",
        nombreReceptor: "CLIENTE SAS",
        nitReceptor: "333333333",
      },
    ]);

    const { cufesByNit } = await extractThirdPartyCufesFromExcel(buffer, "900123456");

    assert.deepEqual(cufesByNit, {
      "333333333": { cufe: "ABCDEF10000000000001", direction: "sent" },
    });
  });

  it("extracts unique third parties from mixed reports (sent and received)", async () => {
    const companyNit = "900123456";
    const cufe1 = "a".repeat(96);
    const cufe2 = "b".repeat(96);
    const cufe3 = "c".repeat(96);
    
    const buffer = await buildDianReport([
      {
        grupo: "Emitidos",
        cufe: cufe1,
        tipoDocumento: "Factura electrónica",
        nombreEmisor: "EMPRESA SAS",
        nitEmisor: companyNit,
        nombreReceptor: "CLIENTE A",
        nitReceptor: "111111111",
      },
      {
        grupo: "Recibidos",
        cufe: cufe2,
        tipoDocumento: "Factura electrónica",
        nombreEmisor: "PROVEEDOR B",
        nitEmisor: "222222222",
        nombreReceptor: "EMPRESA SAS",
        nitReceptor: companyNit,
      },
      {
        grupo: "Emitidos",
        cufe: "d".repeat(96), // Misma empresa A, otro CUFE
        tipoDocumento: "Factura electrónica",
        nombreEmisor: "EMPRESA SAS",
        nitEmisor: companyNit,
        nombreReceptor: "CLIENTE A",
        nitReceptor: "111111111",
      },
      {
        grupo: "unknown", // Sin grupo pero detectable por NIT
        cufe: cufe3,
        tipoDocumento: "Factura electrónica",
        nombreEmisor: "PROVEEDOR C",
        nitEmisor: "333333333",
        nombreReceptor: "EMPRESA SAS",
        nitReceptor: companyNit,
      },
    ]);

    const { cufesByNit, totalCount } = await extractThirdPartyCufesFromExcel(buffer, companyNit);

    // Total de filas con CUFE es 4
    assert.equal(totalCount, 4);
    
    // Terceros únicos detectados: Cliente A (111...), Prov B (222...), Prov C (333...)
    const uniqueNits = Object.keys(cufesByNit);
    assert.equal(uniqueNits.length, 3);
    
    // Cliente A (Emitido)
    assert.equal(cufesByNit["111111111"].cufe, cufe1);
    assert.equal(cufesByNit["111111111"].direction, "sent");

    // Proveedor B (Recibido)
    assert.equal(cufesByNit["222222222"].cufe, cufe2);
    assert.equal(cufesByNit["222222222"].direction, "received");

    // Proveedor C (unknown group, emisor != company) -> received
    assert.equal(cufesByNit["333333333"].cufe, cufe3);
    assert.equal(cufesByNit["333333333"].direction, "received");
  });
});
