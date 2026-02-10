/**
 * Sanitiza un nombre de archivo removiendo caracteres inválidos
 */
export function sanitizeFilename(name: string | null | undefined): string {
  if (!name) return "unknown";
  
  // Normalizar unicode y remover diacríticos
  let sanitized = name.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  
  // Remover saltos de línea y tabs
  sanitized = sanitized.replace(/[\r\n\t]+/g, " ");
  
  // Remover caracteres inválidos para nombres de archivo
  sanitized = sanitized.replace(/[\\/:*?"<>|]+/g, " ");
  
  // Colapsar espacios múltiples
  sanitized = sanitized.replace(/\s+/g, " ").trim();
  
  // Limitar longitud
  return sanitized.slice(0, 200);
}
