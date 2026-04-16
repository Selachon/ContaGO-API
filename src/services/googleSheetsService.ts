import { google } from "googleapis";
import { CausationError } from "./causationService.js";

export interface RegistroCuentaCobroRow {
  rowNumber: number;
  dateValue: unknown;
  driveLink: string;
  reference: string;
}

function getRequiredEnv(name: string): string {
  const value = (process.env[name] || "").trim();
  if (!value) {
    throw new CausationError(`Falta variable de entorno ${name}`, 500, "missing_env_var", { name });
  }
  return value;
}

function getServiceAccountCredentials(): { client_email: string; private_key: string } {
  const clientEmail = getRequiredEnv("GOOGLE_DRIVE_CLIENT_EMAIL");
  const rawPrivateKey = getRequiredEnv("GOOGLE_DRIVE_PRIVATE_KEY");
  const privateKey = rawPrivateKey.replace(/\\n/g, "\n");

  return {
    client_email: clientEmail,
    private_key: privateKey,
  };
}

function getSpreadsheetConfig(): { spreadsheetId: string; gid: number } {
  const spreadsheetId = getRequiredEnv("GOOGLE_SHEETS_CAUSATION_SPREADSHEET_ID");
  const gidRaw = getRequiredEnv("GOOGLE_SHEETS_CAUSATION_GID");
  const gid = Number(gidRaw);

  if (!Number.isInteger(gid) || gid < 0) {
    throw new CausationError("GOOGLE_SHEETS_CAUSATION_GID inválido", 500, "invalid_sheets_gid", {
      gid: gidRaw,
    });
  }

  return { spreadsheetId, gid };
}

function normalizeCell(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

export async function readRegistroCuentasCobroRows(): Promise<{
  spreadsheetId: string;
  gid: string;
  rows: RegistroCuentaCobroRow[];
}> {
  const { spreadsheetId, gid } = getSpreadsheetConfig();
  const credentials = getServiceAccountCredentials();

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
  });

  const sheets = google.sheets({ version: "v4", auth });

  try {
    const metadata = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "sheets(properties(sheetId,title))",
    });

    const targetSheet = metadata.data.sheets?.find((sheet) => sheet.properties?.sheetId === gid);
    if (!targetSheet?.properties?.title) {
      throw new CausationError(
        "No se encontró la pestaña configurada por gid en el Registro de Cuentas de Cobro",
        422,
        "sheet_gid_not_found",
        { spreadsheetId, gid }
      );
    }

    const sheetTitle = targetSheet.properties.title.replace(/'/g, "''");
    const range = `'${sheetTitle}'!A:X`;

    const valuesRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });

    const rows = valuesRes.data.values || [];
    const parsed: RegistroCuentaCobroRow[] = [];

    rows.forEach((row, idx) => {
      const rowNumber = idx + 1;
      const dateValue = row[1] ?? ""; // B
      const driveLink = normalizeCell(row[11]); // L
      const reference = normalizeCell(row[23]); // X

      parsed.push({ rowNumber, dateValue, driveLink, reference });
    });

    return {
      spreadsheetId,
      gid: String(gid),
      rows: parsed,
    };
  } catch (error) {
    if (error instanceof CausationError) throw error;

    throw new CausationError("Error leyendo Google Sheet de Registro de Cuentas de Cobro", 502, "google_sheet_read_failed", {
      spreadsheetId,
      gid,
      cause: error instanceof Error ? error.message : String(error),
    });
  }
}
