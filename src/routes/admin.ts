import { Router, Request, Response } from "express";
import { ObjectId } from "mongodb";
import {
  listUsers,
  getUserById,
  updateUser,
  suspendUser,
  reactivateUser,
  logAdminAction,
  getAuditLogs,
} from "../services/adminService.js";
import { requireAuth } from "../middleware/auth.js";

const router = Router();

// Todas las rutas admin requieren autenticacion y rol admin.
router.use(requireAuth);
router.use((req: Request, res: Response, next) => {
  if (!req.user?.isAdmin) {
    return res.status(403).json({ ok: false, message: "Acceso denegado. Se requiere rol de administrador." });
  }
  next();
});

// ============================================
// GET /admin/users - Listar usuarios con paginacion y filtros
// ============================================
router.get("/users", async (req: Request, res: Response) => {
  try {
    const page = Math.max(1, parseInt(req.query.page as string) || 1);
    const limit = Math.min(100, Math.max(1, parseInt(req.query.limit as string) || 20));
    const search = (req.query.search as string)?.trim() || "";
    const status = req.query.status as string | undefined;
    const tool = req.query.tool as string | undefined;

    const result = await listUsers({ page, limit, search, status, tool });

    res.json({
      ok: true,
      users: result.users,
      pagination: {
        page,
        limit,
        total: result.total,
        totalPages: Math.ceil(result.total / limit),
      },
    });
  } catch (err) {
    console.error("[Admin] Error listando usuarios:", err);
    res.status(500).json({ ok: false, message: "Error interno al listar usuarios" });
  }
});

// ============================================
// GET /admin/users/:id - Detalle de un usuario
// ============================================
router.get("/users/:id", async (req: Request, res: Response) => {
  try {
    const { id } = req.params;

    if (!ObjectId.isValid(id)) {
      return res.status(400).json({ ok: false, message: "ID de usuario invalido" });
    }

    const user = await getUserById(id);
    if (!user) {
      return res.status(404).json({ ok: false, message: "Usuario no encontrado" });
    }

    res.json({ ok: true, user });
  } catch (err) {
    console.error("[Admin] Error obteniendo usuario:", err);
    res.status(500).json({ ok: false, message: "Error interno al obtener usuario" });
  }
});

// ============================================
// PATCH /admin/users/:id - Editar usuario
// ============================================
router.patch("/users/:id", async (req: Request, res: Response) => {
  try {
    const { id } = req.params;
    const actorId = req.user!.userId;

    if (!ObjectId.isValid(id)) {
      return res.status(400).json({ ok: false, message: "ID de usuario invalido" });
    }

    const allowedFields = ["name", "nits", "purchasedTools", "isAdmin"];
    const updates: Record<string, unknown> = {};
    const before: Record<string, unknown> = {};

    // Obtener estado actual para diff de auditoria
    const currentUser = await getUserById(id);
    if (!currentUser) {
      return res.status(404).json({ ok: false, message: "Usuario no encontrado" });
    }

    // Validar y recoger campos permitidos
    const currentUserObj = currentUser as unknown as Record<string, unknown>;
    for (const field of allowedFields) {
      if (req.body[field] !== undefined) {
        before[field] = currentUserObj[field];

        if (field === "name") {
          const name = req.body.name?.trim();
          if (!name || name.length < 2) {
            return res.status(400).json({ ok: false, message: "El nombre debe tener al menos 2 caracteres" });
          }
          updates.name = name;
        } else if (field === "nits") {
          const nits = Array.isArray(req.body.nits) ? req.body.nits : [];
          updates.nits = [...new Set(nits.filter((n: unknown) => typeof n === "string" && n.trim()).map((n: string) => n.trim()))];
        } else if (field === "purchasedTools") {
          const tools = Array.isArray(req.body.purchasedTools) ? req.body.purchasedTools : [];
          updates.purchasedTools = [...new Set(tools.filter((t: unknown) => typeof t === "string" && t.trim()).map((t: string) => t.trim()))];
        } else if (field === "isAdmin") {
          // Evitar que admin se quite a si mismo el rol
          if (id === actorId && req.body.isAdmin === false) {
            return res.status(400).json({ ok: false, message: "No puedes quitarte el rol de administrador a ti mismo" });
          }
          updates.is_admin = Boolean(req.body.isAdmin);
        }
      }
    }

    if (Object.keys(updates).length === 0) {
      return res.status(400).json({ ok: false, message: "No se proporcionaron campos validos para actualizar" });
    }

    const success = await updateUser(id, updates);
    if (!success) {
      return res.status(500).json({ ok: false, message: "Error al actualizar usuario" });
    }

    // Registrar auditoria
    await logAdminAction({
      actorId,
      action: "update_user",
      targetUserId: id,
      before,
      after: updates,
    });

    const updatedUser = await getUserById(id);
    res.json({ ok: true, user: updatedUser, message: "Usuario actualizado correctamente" });
  } catch (err) {
    console.error("[Admin] Error actualizando usuario:", err);
    res.status(500).json({ ok: false, message: "Error interno al actualizar usuario" });
  }
});

// ============================================
// POST /admin/users/:id/suspend - Suspender usuario
// ============================================
router.post("/users/:id/suspend", async (req: Request, res: Response) => {
  try {
    const { id } = req.params;
    const actorId = req.user!.userId;
    const reason = (req.body.reason as string)?.trim() || "Sin motivo especificado";

    if (!ObjectId.isValid(id)) {
      return res.status(400).json({ ok: false, message: "ID de usuario invalido" });
    }

    // No permitir auto-suspension
    if (id === actorId) {
      return res.status(400).json({ ok: false, message: "No puedes suspenderte a ti mismo" });
    }

    const user = await getUserById(id);
    if (!user) {
      return res.status(404).json({ ok: false, message: "Usuario no encontrado" });
    }

    if (user.status === "suspended") {
      return res.status(400).json({ ok: false, message: "El usuario ya esta suspendido" });
    }

    const success = await suspendUser(id);
    if (!success) {
      return res.status(500).json({ ok: false, message: "Error al suspender usuario" });
    }

    await logAdminAction({
      actorId,
      action: "suspend_user",
      targetUserId: id,
      reason,
    });

    res.json({ ok: true, message: "Usuario suspendido correctamente" });
  } catch (err) {
    console.error("[Admin] Error suspendiendo usuario:", err);
    res.status(500).json({ ok: false, message: "Error interno al suspender usuario" });
  }
});

// ============================================
// POST /admin/users/:id/reactivate - Reactivar usuario
// ============================================
router.post("/users/:id/reactivate", async (req: Request, res: Response) => {
  try {
    const { id } = req.params;
    const actorId = req.user!.userId;
    const reason = (req.body.reason as string)?.trim() || "Sin motivo especificado";

    if (!ObjectId.isValid(id)) {
      return res.status(400).json({ ok: false, message: "ID de usuario invalido" });
    }

    const user = await getUserById(id);
    if (!user) {
      return res.status(404).json({ ok: false, message: "Usuario no encontrado" });
    }

    if (user.status === "active") {
      return res.status(400).json({ ok: false, message: "El usuario ya esta activo" });
    }

    const success = await reactivateUser(id);
    if (!success) {
      return res.status(500).json({ ok: false, message: "Error al reactivar usuario" });
    }

    await logAdminAction({
      actorId,
      action: "reactivate_user",
      targetUserId: id,
      reason,
    });

    res.json({ ok: true, message: "Usuario reactivado correctamente" });
  } catch (err) {
    console.error("[Admin] Error reactivando usuario:", err);
    res.status(500).json({ ok: false, message: "Error interno al reactivar usuario" });
  }
});

// ============================================
// GET /admin/audit - Logs de auditoria
// ============================================
router.get("/audit", async (req: Request, res: Response) => {
  try {
    const page = Math.max(1, parseInt(req.query.page as string) || 1);
    const limit = Math.min(100, Math.max(1, parseInt(req.query.limit as string) || 20));
    const targetUserId = req.query.userId as string | undefined;
    const action = req.query.action as string | undefined;

    const result = await getAuditLogs({ page, limit, targetUserId, action });

    res.json({
      ok: true,
      logs: result.logs,
      pagination: {
        page,
        limit,
        total: result.total,
        totalPages: Math.ceil(result.total / limit),
      },
    });
  } catch (err) {
    console.error("[Admin] Error obteniendo logs de auditoria:", err);
    res.status(500).json({ ok: false, message: "Error interno al obtener logs" });
  }
});

export default router;
