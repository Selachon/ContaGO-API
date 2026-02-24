import { Router, Request, Response } from "express";
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
  updateUserGoogleDrive,
  removeUserGoogleDrive,
} from "../services/database.js";

const router = Router();

/**
 * GET /auth/google/status
 * Verifica si el usuario tiene Google Drive vinculado
 */
router.get("/status", requireAuth, async (req: Request, res: Response) => {
  try {
    const userId = req.user!.userId;
    const driveConfig = await getUserGoogleDrive(userId);

    res.json({
      connected: !!driveConfig,
      folderName: driveConfig?.folder_name || null,
      connectedAt: driveConfig?.connected_at || null,
    });
  } catch (err) {
    console.error("Error verificando estado de Google Drive:", err);
    res.status(500).json({ status: "error", detalle: "Error verificando conexión" });
  }
});

/**
 * GET /auth/google/authorize
 * Redirige al usuario a la pantalla de consentimiento de Google
 */
router.get("/authorize", requireAuth, (req: Request, res: Response) => {
  try {
    // Guardar userId en state para recuperarlo en callback
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
 * GET /auth/google/callback
 * Callback de Google OAuth, intercambia código por tokens
 */
router.get("/callback", async (req: Request, res: Response) => {
  const { code, state, error } = req.query;

  // HTML de respuesta (se mostrará en el popup)
  const sendResponse = (success: boolean, message: string) => {
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
    <p>${message}</p>
    <div class="note">Ya puedes cerrar esta ventana</div>
  </div>
</body>
</html>`);
  };

  // Error de Google
  if (error) {
    console.error("Error de Google OAuth:", error);
    return sendResponse(false, "La autorización fue cancelada o denegada.");
  }

  // Validar código y state
  if (!code || typeof code !== "string") {
    return sendResponse(false, "No se recibió código de autorización.");
  }

  if (!state || typeof state !== "string") {
    return sendResponse(false, "Sesión inválida. Por favor, intente nuevamente.");
  }

  try {
    // Decodificar state para obtener userId
    const stateData = JSON.parse(Buffer.from(state, "base64").toString("utf8"));
    const userId = stateData.userId;

    if (!userId) {
      return sendResponse(false, "Sesión inválida.");
    }

    // Intercambiar código por tokens
    const { accessToken, refreshToken, expiryDate } = await exchangeCodeForTokens(code);

    // Obtener email del usuario de Google
    const userEmail = await getGoogleUserEmail(accessToken);

    // Encriptar tokens
    const encryptedTokens = encryptDriveTokens(accessToken, refreshToken);

    // Crear configuración inicial
    const driveConfig = {
      encrypted_access_token: encryptedTokens.encrypted_access_token,
      encrypted_refresh_token: encryptedTokens.encrypted_refresh_token,
      token_expiry: new Date(expiryDate).toISOString(),
      folder_id: "",
      folder_name: "ContaGO Facturas",
      connected_at: new Date().toISOString(),
      last_used: new Date().toISOString(),
      user_email: userEmail,
    };

    // Obtener o crear carpeta (esto también valida que los tokens funcionan)
    const folderId = await getOrCreateFolder(driveConfig);
    driveConfig.folder_id = folderId;

    // Guardar en base de datos
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
 * Desconecta Google Drive del usuario
 */
router.post("/disconnect", requireAuth, async (req: Request, res: Response) => {
  try {
    const userId = req.user!.userId;
    const driveConfig = await getUserGoogleDrive(userId);

    if (!driveConfig) {
      return res.status(400).json({ status: "error", detalle: "Google Drive no está conectado" });
    }

    // Revocar tokens en Google
    await revokeAccess(driveConfig);

    // Eliminar de la base de datos
    await removeUserGoogleDrive(userId);

    console.log(`Google Drive desconectado para usuario ${userId}`);
    res.json({ status: "ok", message: "Google Drive desconectado exitosamente" });

  } catch (err) {
    console.error("Error desconectando Google Drive:", err);
    res.status(500).json({ status: "error", detalle: "Error al desconectar" });
  }
});

export default router;
