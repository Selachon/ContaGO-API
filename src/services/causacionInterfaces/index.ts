/**
 * Motor "Causación vía Interfaces".
 *
 * Convierte archivos DIAN en interfaces de importación de Siigo. Fase 1:
 * proceso de compras. Próximas fases: ventas y terceros.
 */
export { runCompras, dividirEnPartes } from "./compras.js";
export { runVentas, analyzeVentas, extraerPrefijoConsecutivo } from "./ventas.js";
export { runTerceros, calcularDigitoVerificacion } from "./terceros.js";
export { escribirInterfazSiigo, escribirInformeRetenciones } from "./excelWriter.js";
export { escribirTercerosXlsm } from "./tercerosXlsm.js";
export {
  CausacionInterfacesError,
  COLUMNAS_SIIGO,
  COLUMNAS_INFORME,
  COLUMNAS_TERCEROS,
  type CeldaValor,
  type ComprasInput,
  type ComprasResultado,
  type VentasInput,
  type VentasResultado,
  type VentasAnalisis,
  type TercerosInput,
  type TercerosResultado,
  type FilaSiigo,
  type FilaInforme,
  type FilaTercero,
} from "./types.js";
