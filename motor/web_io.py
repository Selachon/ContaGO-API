"""
Helper de I/O para el modo --web del motor.

Objetivo: NO tocar la logica de calculo. El modo consola (sin --web) queda
exactamente igual que el original; el modo --web solo cambia los BORDES:
  - inputs de configuracion -> JSON de parametros (--params)
  - ruta del DIAN, parametrizaciones y salida -> argumentos
  - errores bloqueantes (input + exit) -> linea JSON + exit(2)
  - exito -> linea JSON con resumen + archivos

Uso en cada script del motor:

    import web_io
    CTX = web_io.init()
    WEB = CTX.web
    PARAMS = CTX.params
    ...
    if WEB:
        web_io.error("proveedores_faltantes", "Faltan proveedores", sorted(faltantes))
    print("FALTAN PROVEEDORES POR PARAMETRIZAR:")   # <- camino consola intacto
    ...
"""
import argparse
import json
import os
import sys


class _Ctx:
    def __init__(self):
        self.web = False
        self.params = {}
        self.dian = None       # DIAN principal (compras/ventas usan uno)
        self.dians = []        # lista de DIAN (terceros puede usar varios)
        self.out = None
        self.param_compras = None
        self.param_ventas = None
        self.impuestos = None
        self.plantilla_terceros = None
        self.solo_prefijos = False


def init():
    ctx = _Ctx()
    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument("--web", action="store_true")
    ap.add_argument("--dian", action="append")
    ap.add_argument("--params")
    ap.add_argument("--out")
    ap.add_argument("--param-compras")
    ap.add_argument("--param-ventas")
    ap.add_argument("--impuestos")
    ap.add_argument("--plantilla-terceros")
    ap.add_argument("--prefijos", action="store_true")
    args, rest = ap.parse_known_args()

    ctx.web = args.web
    ctx.dians = args.dian or []
    ctx.dian = ctx.dians[0] if ctx.dians else None
    ctx.out = args.out
    ctx.param_compras = args.param_compras
    ctx.param_ventas = args.param_ventas
    ctx.impuestos = args.impuestos
    ctx.plantilla_terceros = args.plantilla_terceros
    ctx.solo_prefijos = args.prefijos

    if args.params and os.path.exists(args.params):
        with open(args.params, encoding="utf-8") as fh:
            ctx.params = json.load(fh)

    # Preserva compatibilidad con el modo consola (argv[1] = nombre sugerido).
    sys.argv = [sys.argv[0]] + rest
    return ctx


def _emit(obj):
    print(json.dumps(obj, ensure_ascii=False))


def error(tipo, mensaje, detalle=None):
    """Modo web: emite JSON de error y termina con codigo 2."""
    _emit({"status": "error", "tipo": tipo, "mensaje": mensaje,
           "detalle": detalle if detalle is not None else []})
    sys.exit(2)


def ok(resumen, archivos):
    """Modo web: emite JSON de exito."""
    _emit({"status": "ok", "resumen": resumen, "archivos": archivos})


def prefijos(facturas, ds, nc):
    """Modo web fase 1 (ventas): emite los prefijos detectados y termina."""
    _emit({"status": "prefijos",
           "facturas": list(facturas), "ds": list(ds), "nc": list(nc)})
    sys.exit(0)
