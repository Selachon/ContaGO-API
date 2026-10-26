/** Cliente de ContaGO Contable para facturar las ventas del portal (clave de servicio propia; el navegador nunca habla con Contable). */
export class ContableError extends Error { constructor(m: string, public status = 502, public codigo?: string) { super(m); } }

async function llamar<T>(path: string, init: { method?: string; body?: unknown } = {}): Promise<T> {
  const key = process.env.CONTABLE_PORTAL_KEY;
  const base = (process.env.CONTABLE_API_URL || "").replace(/\/+$/, "");
  if (!key || !base) throw new ContableError("La facturación de ventas no está configurada.", 503);
  let res: Response;
  try {
    res = await fetch(`${base}/api/servicio/portal${path}`, {
      method: init.method || "GET",
      headers: { "x-service-key": key, ...(init.body !== undefined ? { "content-type": "application/json" } : {}) },
      body: init.body !== undefined ? JSON.stringify(init.body) : undefined,
      signal: AbortSignal.timeout(150_000),
    });
  } catch {
    throw new ContableError("No se pudo conectar con el software contable. La venta queda pendiente de facturar.", 502);
  }
  const data: any = await res.json().catch(() => ({}));
  if (!res.ok || data.ok === false) throw new ContableError(data.message || "Error del software contable", res.status >= 400 ? res.status : 502, data.codigo);
  return data as T;
}

export const ventaKey = (t: { userId: string; toolId: string; paymentId: string }) => `${t.userId}|${t.toolId}|${t.paymentId}`;
export const prepararVentaContable = (t: { userId: string; toolId: string; paymentId: string }, doc?: { docType?: string; document?: string }) =>
  llamar<any>(`/preparar?venta=${encodeURIComponent(ventaKey(t))}${doc?.document ? `&docType=${encodeURIComponent(doc.docType || "")}&document=${encodeURIComponent(doc.document)}` : ""}`);
export const facturarVentaContable = (t: { userId: string; toolId: string; paymentId: string }, cliente: unknown) => llamar<any>("/facturar", { method: "POST", body: { venta: ventaKey(t), cliente } });
export const lugaresContable = () => llamar<any>("/lugares");

/** PDF de la factura de una venta (binario). */
export async function pdfContable(docId: string): Promise<{ data: Buffer; tipo: string }> {
  const key = process.env.CONTABLE_PORTAL_KEY;
  const base = (process.env.CONTABLE_API_URL || "").replace(/\/+$/, "");
  if (!key || !base) throw new ContableError("La facturación de ventas no está configurada.", 503);
  let res: Response;
  try { res = await fetch(`${base}/api/servicio/portal/pdf?doc=${encodeURIComponent(docId)}`, { headers: { "x-service-key": key }, signal: AbortSignal.timeout(60_000) }); }
  catch { throw new ContableError("No se pudo conectar con el software contable.", 502); }
  if (!res.ok) { const d: any = await res.json().catch(() => ({})); throw new ContableError(d.message || "No se pudo descargar el PDF.", res.status); }
  return { data: Buffer.from(await res.arrayBuffer()), tipo: res.headers.get("content-type") || "application/pdf" };
}
