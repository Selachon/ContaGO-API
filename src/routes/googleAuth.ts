import { Router, Request, Response } from "express";
import crypto from "crypto";
import { requireAuth } from "../middleware/auth.js";
import {
  getAuthUrl,
  exchangeCodeForTokens,
  encryptDriveTokens,
  getGoogleUserEmail,
  getOrCreateFolder,
  revokeAccess,
} from "../services/googleDrive.js";
import {
  getUserGoogleDrive,
  getUserGoogleDriveById,
  getUserGoogleDrives,
  updateUserGoogleDrive,
  removeUserGoogleDrive,
  removeUserGoogleDriveById,
  setSelectedUserGoogleDrive,
} from "../services/database.js";

const router = Router();

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * GET /auth/google/status
 * Retorna el estado de vinculacion de Google Drive del usuario autenticado.
 */
router.get("/status", requireAuth, async (req: Request, res: Response) => {
  try {
    const userId = req.user!.userId;
    const drives = await getUserGoogleDrives(userId);
    const selected = await getUserGoogleDrive(userId);

    res.json({
      connected: drives.length > 0,
      folderName: selected?.folder_name || null,
      connectedAt: selected?.connected_at || null,
      selectedConnectionId: selected?.connection_id || null,
      connections: drives.map((d) => ({
        connectionId: d.connection_id,
        email: d.user_email,
        folderName: d.folder_name,
        connectedAt: d.connected_at,
      })),
    });
  } catch (err) {
    console.error("Error verificando estado de Google Drive:", err);
    res.status(500).json({ status: "error", detalle: "Error verificando conexión" });
  }
});

/**
 * GET /auth/google/authorize
 * Redirige al consentimiento de Google OAuth.
 */
router.get("/authorize", requireAuth, (req: Request, res: Response) => {
  try {
    // El userId viaja en state para vincular el callback con el usuario correcto.
    const state = Buffer.from(JSON.stringify({ userId: req.user!.userId })).toString("base64");
    const authUrl = getAuthUrl(state);

    res.redirect(authUrl);
  } catch (err) {
    console.error("Error generando URL de autorización:", err);
    res.status(500).send(`
      <html>
        <body style="font-family: system-ui; padding: 40px; text-align: center;">
          <h2>Error</h2>
          <p>No se pudo iniciar la autorización con Google.</p>
          <p>Por favor, cierre esta ventana e intente nuevamente.</p>
          <script>setTimeout(() => window.close(), 3000);</script>
        </body>
      </html>
    `);
  }
});

/**
 * POST /auth/google/authorize-url
 * Devuelve la URL OAuth para abrir en popup sin exponer token de sesion en query string.
 */
router.post("/authorize-url", requireAuth, (req: Request, res: Response) => {
  try {
    const state = Buffer.from(JSON.stringify({ userId: req.user!.userId })).toString("base64");
    const authUrl = getAuthUrl(state);
    res.json({ ok: true, authUrl });
  } catch (err) {
    console.error("Error generando URL de autorización:", err);
    res.status(500).json({ ok: false, message: "No se pudo iniciar la autorización con Google." });
  }
});

/**
 * GET /auth/google/callback
 * Callback OAuth: valida state, intercambia codigo y guarda tokens cifrados.
 */
router.get("/callback", async (req: Request, res: Response) => {
  const { code, state, error } = req.query;

  // La UI del popup se renderiza con mensaje escapado para evitar inyeccion HTML.
  const sendResponse = (success: boolean, message: string) => {
    const safeMessage = escapeHtml(message || "");
    const color = success ? "#22c55e" : "#ef4444";
    const bgColor = success ? "#f0fdf4" : "#fef2f2";
    const icon = success ? "✓" : "✗";
    
    res.send(`<!DOCTYPE html>
<html>
<head>
  <title>${success ? "Conectado" : "Error"} - ContaGO</title>
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: system-ui, -apple-system, sans-serif;
      display: flex;
      align-items: center;
      justify-content: center;
      min-height: 100vh;
      margin: 0;
      background: ${bgColor};
    }
    .container {
      text-align: center;
      padding: 48px 40px;
      background: white;
      border-radius: 16px;
      box-shadow: 0 4px 24px rgba(0,0,0,0.1);
      max-width: 420px;
      margin: 20px;
    }
    .icon {
      width: 72px;
      height: 72px;
      border-radius: 50%;
      background: ${color};
      color: white;
      font-size: 36px;
      display: flex;
      align-items: center;
      justify-content: center;
      margin: 0 auto 24px;
    }
    h2 { color: #1f2937; margin: 0 0 12px; font-size: 24px; }
    p { color: #6b7280; margin: 0 0 24px; line-height: 1.5; }
    .note { 
      font-size: 14px; 
      color: #9ca3af; 
      background: #f9fafb;
      padding: 12px 16px;
      border-radius: 8px;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="icon">${icon}</div>
    <h2>${success ? "Google Drive Conectado" : "Error de Conexión"}</h2>
    <p>${safeMessage}</p>
    <div class="note">Ya puedes cerrar esta ventana</div>
  </div>
</body>
</html>`);
  };

  // Errores devueltos por Google (cancelacion o denegacion de permisos).
  if (error) {
    console.error("Error de Google OAuth:", error);
    return sendResponse(false, "La autorización fue cancelada o denegada.");
  }

  // Validaciones basicas para evitar callbacks incompletos.
  if (!code || typeof code !== "string") {
    return sendResponse(false, "No se recibió código de autorización.");
  }

  if (!state || typeof state !== "string") {
    return sendResponse(false, "Sesión inválida. Por favor, intente nuevamente.");
  }

  try {
    // Recupera el userId enviado al iniciar OAuth.
    const stateData = JSON.parse(Buffer.from(state, "base64").toString("utf8"));
    const userId = stateData.userId;

    if (!userId) {
      return sendResponse(false, "Sesión inválida.");
    }

    // Intercambia codigo por tokens de acceso y refresh.
    const { accessToken, refreshToken, expiryDate } = await exchangeCodeForTokens(code);

    // Se guarda email para referencia operativa en soporte.
    const userEmail = await getGoogleUserEmail(accessToken);

    // Nunca persistir tokens en texto plano.
    const encryptedTokens = encryptDriveTokens(accessToken, refreshToken);

    // Configuracion base de la integracion por usuario.
    const driveConfig = {
      connection_id: crypto.randomUUID(),
      encrypted_access_token: encryptedTokens.encrypted_access_token,
      encrypted_refresh_token: encryptedTokens.encrypted_refresh_token,
      token_expiry: new Date(expiryDate).toISOString(),
      folder_id: "",
      folder_name: "ContaGO Facturas",
      connected_at: new Date().toISOString(),
      last_used: new Date().toISOString(),
      user_email: userEmail,
    };

    // Crear/obtener carpeta tambien valida la vigencia de tokens.
    const folderId = await getOrCreateFolder(driveConfig);
    driveConfig.folder_id = folderId;

    // Persistir integracion activa.
    await updateUserGoogleDrive(userId, driveConfig);

    console.log(`Google Drive conectado para usuario ${userId}`);
    return sendResponse(true, "Tu cuenta de Google Drive ha sido vinculada exitosamente.");

  } catch (err) {
    const error = err as Error;
    console.error("Error en callback de Google:", error);
    console.error("Stack:", error.stack);
    return sendResponse(false, `Error: ${error.message || "Error desconocido"}`);
  }
});

/**
 * POST /auth/google/disconnect
 * Revoca acceso en Google y elimina la vinculacion local.
 */
router.post("/disconnect", requireAuth, async (req: Request, res: Response) => {
  try {
    const userId = req.user!.userId;
    const { connectionId } = req.body as { connectionId?: string };
    const driveConfig = connectionId
      ? await getUserGoogleDriveById(userId, connectionId)
      : await getUserGoogleDrive(userId);

    if (!driveConfig) {
      return res.status(400).json({ status: "error", detalle: "Google Drive no está conectado" });
    }

    // Revocar primero para evitar dejar permisos activos en Google.
    await revokeAccess(driveConfig);

    // Borrar configuracion local despues de revocar.
    if (connectionId) {
      await removeUserGoogleDriveById(userId, connectionId);
    } else {
      await removeUserGoogleDrive(userId);
    }

    console.log(`Google Drive desconectado para usuario ${userId}`);
    res.json({ status: "ok", message: "Google Drive desconectado exitosamente" });

  } catch (err) {
    console.error("Error desconectando Google Drive:", err);
    res.status(500).json({ status: "error", detalle: "Error al desconectar" });
  }
});

router.post("/select", requireAuth, async (req: Request, res: Response) => {
  try {
    const userId = req.user!.userId;
    const { connectionId } = req.body as { connectionId?: string };
    if (!connectionId) {
      return res.status(400).json({ status: "error", detalle: "connectionId es requerido" });
    }
    const ok = await setSelectedUserGoogleDrive(userId, connectionId);
    if (!ok) {
      return res.status(404).json({ status: "error", detalle: "Conexión no encontrada" });
    }
    return res.json({ status: "ok" });
  } catch (err) {
    console.error("Error seleccionando Google Drive:", err);
    return res.status(500).json({ status: "error", detalle: "Error al seleccionar conexión" });
  }
});

export default router;
