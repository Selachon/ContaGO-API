import { describe, it } from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import {
  CausationError,
  ensurePdfExtension,
  findUniqueMatchFromExcel,
  findUniqueMatchFromRows,
  parseExcelDateToFolders,
} from "./causationService.js";

async function buildExcelBuffer(rows: Array<{ date: unknown; driveLink: string; reference: string }>): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Datos");

  rows.forEach((row, idx) => {
    const rowIndex = idx + 1;
    sheet.getCell(rowIndex, 2).value = row.date as ExcelJS.CellValue;
    sheet.getCell(rowIndex, 12).value = row.driveLink;
    sheet.getCell(rowIndex, 24).value = row.reference;
  });

  const bytes = await workbook.xlsx.writeBuffer();
  return Buffer.from(bytes);
}

describe("causationService", () => {
  it("finds exact reference match case-insensitive without extension", async () => {
    const excel = await buildExcelBuffer([
      {
        date: "2026-04-10",
        driveLink: "https://drive.google.com/file/d/1AbcDefGhIjKlMnOpQrStUvWxYz12345/view",
        reference: "DS-1001",
      },
    ]);

    const match = await findUniqueMatchFromExcel(excel, "ds-1001.PDF");
    assert.equal(match.matchedRow, 1);
    assert.equal(match.reference, "DS-1001");
  });

  it("throws clear error when no match exists", async () => {
    const excel = await buildExcelBuffer([
      { date: "2026-04-10", driveLink: "https://drive.google.com/file/d/1AbcDefGhIjKlMnOpQrStUvWxYz12345/view", reference: "FC-1" },
    ]);

    await assert.rejects(() => findUniqueMatchFromExcel(excel, "DS-NO-EXISTE.pdf"), (err: unknown) => {
      assert.ok(err instanceof CausationError);
      assert.equal(err.code, "reference_not_found");
      return true;
    });
  });

  it("throws clear error when multiple matches exist", async () => {
    const excel = await buildExcelBuffer([
      { date: "2026-04-10", driveLink: "https://drive.google.com/file/d/1AbcDefGhIjKlMnOpQrStUvWxYz12345/view", reference: "DS-1" },
      { date: "2026-04-11", driveLink: "https://drive.google.com/file/d/1AbcDefGhIjKlMnOpQrStUvWxYz12345/view", reference: "ds-1.pdf" },
    ]);

    await assert.rejects(() => findUniqueMatchFromExcel(excel, "DS-1.pdf"), (err: unknown) => {
      assert.ok(err instanceof CausationError);
      assert.equal(err.code, "multiple_references_found");
      return true;
    });
  });

  it("throws clear error when column L is empty", async () => {
    const excel = await buildExcelBuffer([{ date: "2026-04-10", driveLink: "", reference: "DS-300" }]);

    await assert.rejects(() => findUniqueMatchFromExcel(excel, "DS-300"), (err: unknown) => {
      assert.ok(err instanceof CausationError);
      assert.equal(err.code, "empty_drive_link");
      return true;
    });
  });

  it("parses year and month folder from column B date", () => {
    const parsed = parseExcelDateToFolders("2025-01-15");
    assert.equal(parsed.year, "2025");
    assert.equal(parsed.monthName, "01-Enero");
  });

  it("throws clear error when column B date is invalid", () => {
    assert.throws(() => parseExcelDateToFolders("fecha invalida"), (err: unknown) => {
      assert.ok(err instanceof CausationError);
      assert.equal(err.code, "invalid_date_column_b");
      return true;
    });
  });

  it("ensures final file name has pdf extension", () => {
    assert.equal(ensurePdfExtension("DS-500"), "DS-500.pdf");
    assert.equal(ensurePdfExtension("DS-500.pdf"), "DS-500.pdf");
  });

  it("findUniqueMatchFromRows finds exact case-insensitive reference", () => {
    const match = findUniqueMatchFromRows(
      [
        {
          rowNumber: 2,
          dateValue: "2025-10-01",
          driveLink: "https://drive.google.com/file/d/1AbcDefGhIjKlMnOpQrStUvWxYz12345/view",
          reference: "DS-ABC-01",
        },
      ],
      "ds-abc-01.pdf"
    );

    assert.equal(match.matchedRow, 2);
    assert.equal(match.reference, "DS-ABC-01");
  });

  it("findUniqueMatchFromRows throws when no rows match", () => {
    assert.throws(
      () =>
        findUniqueMatchFromRows(
          [
            {
              rowNumber: 2,
              dateValue: "2025-10-01",
              driveLink: "https://drive.google.com/file/d/1AbcDefGhIjKlMnOpQrStUvWxYz12345/view",
              reference: "FC-1",
            },
          ],
          "DS-99.pdf"
        ),
      (err: unknown) => {
        assert.ok(err instanceof CausationError);
        assert.equal(err.code, "reference_not_found");
        return true;
      }
    );
  });

  it("findUniqueMatchFromRows throws when multiple rows match", () => {
    assert.throws(
      () =>
        findUniqueMatchFromRows(
          [
            {
              rowNumber: 2,
              dateValue: "2025-10-01",
              driveLink: "https://drive.google.com/file/d/1AbcDefGhIjKlMnOpQrStUvWxYz12345/view",
              reference: "DS-1",
            },
            {
              rowNumber: 3,
              dateValue: "2025-10-01",
              driveLink: "https://drive.google.com/file/d/1AbcDefGhIjKlMnOpQrStUvWxYz12345/view",
              reference: "ds-1.pdf",
            },
          ],
          "DS-1"
        ),
      (err: unknown) => {
        assert.ok(err instanceof CausationError);
        assert.equal(err.code, "multiple_references_found");
        return true;
      }
    );
  });

  it("findUniqueMatchFromRows throws when drive link is empty", () => {
    assert.throws(
      () =>
        findUniqueMatchFromRows(
          [
            {
              rowNumber: 2,
              dateValue: "2025-10-01",
              driveLink: "",
              reference: "DS-1",
            },
          ],
          "DS-1"
        ),
      (err: unknown) => {
        assert.ok(err instanceof CausationError);
        assert.equal(err.code, "empty_drive_link");
        return true;
      }
    );
  });
});
