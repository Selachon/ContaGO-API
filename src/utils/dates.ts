const SPANISH_MONTHS = [
  "Ene", "Feb", "Mar", "Abr", "May", "Jun",
  "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"
];

/**
 * Convierte una fecha ISO (YYYY-MM-DD) a formato español (Ene 01 2024)
 */
export function formatSpanishLabel(dateStr: string | null | undefined): string | null {
  if (!dateStr) return null;
  
  const normalized = dateStr.replace(/\//g, "-");
  
  try {
    const [year, month, day] = normalized.split("-").map(Number);
    if (!year || !month || !day) return null;
    
    return `${SPANISH_MONTHS[month - 1]} ${day.toString().padStart(2, "0")} ${year}`;
  } catch {
    return null;
  }
}
