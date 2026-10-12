/**
 * Búsqueda global (Ctrl/Cmd + K). Resultados agrupados por dominio y siempre
 * acotados al tenant activo.
 */
import { COL, col, fromDocs, scoped } from "./db.js";
import { documentLabel } from "./documents.js";
import type { ErpDocument, ErpPayment, ErpSupplier } from "./types.js";

export interface SearchHit {
  group: "documents" | "suppliers" | "payments";
  id: string;
  title: string;
  subtitle?: string;
  meta?: string;
  href: string;
}

export async function globalSearch(tenantId: string, term: string): Promise<SearchHit[]> {
  const q = String(term || "").trim();
  if (q.length < 2) return [];
  const rx = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");
  const numeric = Number(q.replace(/[^\d]/g, ""));

  const [documents, suppliers, payments] = await Promise.all([
    col(COL.documents)
      .find(
        scoped(tenantId, {
          $or: [
            { supplierName: rx }, { supplierNit: rx }, { documentNumber: rx },
            { cufe: rx }, { description: rx }, { "siigo.siigoName": rx },
            ...(Number.isFinite(numeric) && numeric > 0 ? [{ total: numeric }] : []),
          ],
        })
      )
      .sort({ issueDate: -1 })
      .limit(8)
      .toArray(),
    col(COL.suppliers)
      .find(scoped(tenantId, { $or: [{ name: rx }, { identification: rx }] }))
      .limit(5)
      .toArray(),
    col(COL.payments)
      .find(
        scoped(tenantId, {
          $or: [
            { reference: rx }, { "siigo.siigoName": rx }, { "allocations.documentLabel": rx },
            ...(Number.isFinite(numeric) && numeric > 0 ? [{ value: numeric }] : []),
          ],
        })
      )
      .sort({ paidAt: -1 })
      .limit(5)
      .toArray(),
  ]);

  const hits: SearchHit[] = [];

  for (const d of fromDocs<ErpDocument>(documents)) {
    hits.push({
      group: "documents",
      id: d.id,
      title: `${documentLabel(d)} — ${d.supplierName}`,
      subtitle: d.description || undefined,
      meta: `${d.issueDate} · ${d.total.toLocaleString("es-CO")} COP`,
      href: `/documentos/${d.id}`,
    });
  }
  for (const s of fromDocs<ErpSupplier>(suppliers)) {
    hits.push({
      group: "suppliers",
      id: s.id,
      title: s.name,
      subtitle: `NIT ${s.identification}`,
      meta: s.documentCount ? `${s.documentCount} documentos` : undefined,
      href: `/documentos?supplierNit=${s.identification}`,
    });
  }
  for (const p of fromDocs<ErpPayment>(payments)) {
    hits.push({
      group: "payments",
      id: p.id,
      title: `Pago ${p.value.toLocaleString("es-CO")} COP`,
      subtitle: p.siigo?.siigoName || p.reference || undefined,
      meta: `${p.paidAt} · ${p.allocations.length} documento(s)`,
      href: `/pagos/${p.id}`,
    });
  }

  return hits;
}
