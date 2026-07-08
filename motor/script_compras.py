import pandas as pd
import os
import unicodedata
import sys
from datetime import datetime

import web_io

CTX = web_io.init()
WEB = CTX.web
PARAMS = CTX.params

# === CONFIG POR EMPRESA (unificacion) ===
# Modo de manejo de obsequios cuyo IVA no coincide con el descuento detalle:
#   "error"        -> comportamiento Contago/ABC/Soft: detiene el proceso (bloqueante)
#   "contabilizar" -> comportamiento ElToro: avisa y contabiliza IVA vs CxP directamente
OBSEQUIOS_MODE = os.environ.get("CONTAGO_OBSEQUIOS_MODE", "error").strip().lower()


def renombrar_si_existe(df, mapeo):
    renames = {}
    for actual in df.columns:
        clave = str(actual).strip()
        if clave in mapeo:
            renames[actual] = mapeo[clave]
    if renames:
        df.rename(columns=renames, inplace=True)


def normalizar_texto(texto):
    texto = str(texto).strip().lower()
    texto = texto.replace("�", "u")
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.split())


def cargar_hoja_dian_con_header_dinamico(path, hoja):
    bruto = pd.read_excel(path, sheet_name=hoja, header=None)

    hoja_norm = normalizar_texto(hoja)
    if "detallado" in hoja_norm:
        candidatos = ["numero factura", "base del impuesto"]
    else:
        candidatos = ["tipo documento", "nit emisor"]
    fila_header = None

    for i in range(min(30, len(bruto))):
        valores = [normalizar_texto(v) for v in bruto.iloc[i].tolist() if pd.notna(v)]
        unidos = " | ".join(valores)
        if all(c in unidos for c in candidatos):
            fila_header = i
            break

    if fila_header is None:
        # Compatibilidad con formato anterior
        return pd.read_excel(path, sheet_name=hoja, header=1)

    headers = [str(x).strip() if pd.notna(x) else "" for x in bruto.iloc[fila_header].tolist()]
    df = bruto.iloc[fila_header + 1 :].copy()
    df.columns = headers
    df = df.loc[:, [c for c in df.columns if str(c).strip() != ""]]
    return df.reset_index(drop=True)

def seleccionar_archivo_entrada(nombre_sugerido=""):
    entrada_dir = "Entrada"
    if not os.path.isdir(entrada_dir):
        print("No existe la carpeta Entrada.")
        input("Presiona Enter para cerrar...")
        exit()

    archivos_xlsx = [
        f for f in os.listdir(entrada_dir)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    ]

    if not archivos_xlsx:
        print("No hay archivos .xlsx en la carpeta Entrada.")
        input("Presiona Enter para cerrar...")
        exit()

    if nombre_sugerido:
        nombre = nombre_sugerido.strip()
        if not nombre.lower().endswith(".xlsx"):
            nombre += ".xlsx"

        exacto = os.path.join(entrada_dir, nombre)
        if os.path.exists(exacto):
            print(f"Archivo DIAN usado: {nombre}")
            return exacto

        nombre_l = nombre.lower()
        parciales = [f for f in archivos_xlsx if nombre_l in f.lower()]
        if len(parciales) == 1:
            print(f"Archivo detectado automáticamente: {parciales[0]}")
            return os.path.join(entrada_dir, parciales[0])

    archivo_reciente = max(
        archivos_xlsx,
        key=lambda f: os.path.getmtime(os.path.join(entrada_dir, f))
    )
    print(f"Archivo DIAN usado automáticamente (más reciente): {archivo_reciente}")
    return os.path.join(entrada_dir, archivo_reciente)


if WEB:
    archivo_dian = CTX.dian
    archivo_param = CTX.param_compras
    archivo_impuestos = CTX.impuestos
    tipo_comprobante_compras = str(PARAMS.get("tipo_comprobante", ""))
    consecutivo_inicial_compras = int(PARAMS.get("consecutivo_inicial", 1))
    tipo_comprobante_nc = str(PARAMS.get("tipo_comprobante_nc", ""))
    consecutivo_inicial_nc = int(PARAMS.get("consecutivo_inicial_nc", 1))
else:
    archivo_sugerido = str(sys.argv[1]).strip() if len(sys.argv) > 1 else ""
    archivo_dian = seleccionar_archivo_entrada(archivo_sugerido)
    archivo_param = "Parametrizacion_Siigo_Compras.xlsx"
    archivo_impuestos = os.path.join("..", "Tabla maestro impuestos.xlsx")

    tipo_comprobante_compras = input("Tipo de comprobante para compras: ")
    consecutivo_inicial_compras = int(input("Consecutivo inicial compras: "))

    tipo_comprobante_nc = input("Tipo de comprobante para notas crédito: ")
    consecutivo_inicial_nc = int(input("Consecutivo inicial notas crédito: "))

facturas = cargar_hoja_dian_con_header_dinamico(archivo_dian, "Facturas DIAN")
detalle = cargar_hoja_dian_con_header_dinamico(archivo_dian, "Detallado")
param = pd.read_excel(archivo_param, sheet_name="Proveedores")
impuestos = pd.read_excel(archivo_impuestos)

facturas.columns = facturas.columns.astype(str).str.strip()
detalle.columns = detalle.columns.astype(str).str.strip()
param.columns = param.columns.astype(str).str.strip()
impuestos.columns = impuestos.columns.astype(str).str.strip()

# Ajuste de formato DIAN nuevo -> formato interno esperado
renombrar_si_existe(
    facturas,
    {
        "No.": "No",
        "Número factura": "Numero Factura",
        "Tipo documento": "Tipo de documento",
        "Número factura": "Numero Factura",
        "NIT emisor": "NIT Emisor",
        "Razón social emisor": "Razon Social Emisor",
        "Fecha": "Fecha de emision",
        "Subtotal": "Subtotal antes de impuestos",
        "Descuento": "Descuento detalle",
        "Recargo": "Recargo detalle",
        "Total": "Valor total",
        "Razón Social Emisor": "Razon Social Emisor",
    },
)

# Segunda pasada: renombre robusto por nombre normalizado (soporta tildes/codificación dañada)
mapeo_normalizado_facturas = {
    "tipo documento": "Tipo de documento",
    "numero factura": "Numero Factura",
    "nit emisor": "NIT Emisor",
    "razon social emisor": "Razon Social Emisor",
    "fecha": "Fecha de emision",
    "subtotal": "Subtotal antes de impuestos",
    "descuento": "Descuento detalle",
    "recargo": "Recargo detalle",
    "total": "Valor total",
}

renames_fact = {}
for actual in facturas.columns:
    clave_norm = normalizar_texto(actual)
    if clave_norm in mapeo_normalizado_facturas:
        renames_fact[actual] = mapeo_normalizado_facturas[clave_norm]
if renames_fact:
    facturas.rename(columns=renames_fact, inplace=True)

renombrar_si_existe(
    detalle,
    {
        "Número Factura": "Numero Factura",
        "Número factura": "Numero Factura",
    },
)

mapeo_normalizado_detalle = {
    "numero factura": "Numero Factura",
}

renames_det = {}
for actual in detalle.columns:
    clave_norm = normalizar_texto(actual)
    if clave_norm in mapeo_normalizado_detalle:
        renames_det[actual] = mapeo_normalizado_detalle[clave_norm]
if renames_det:
    detalle.rename(columns=renames_det, inplace=True)

col_nit = "NIT Emisor"
col_factura = "Numero Factura"
col_total = "Valor total"
col_fecha = "Fecha de emision"
col_tipo_documento = "Tipo de documento"

col_subtotal_facturas = "Subtotal antes de impuestos"
col_base_detalle = "Base del impuesto"
col_descuento = "Descuento detalle"
col_recargo = "Recargo detalle"


def limpiar_nit(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    if texto.lower() == "nan":
        return ""

    if texto.endswith(".0"):
        texto = texto[:-2]

    return texto


def limpiar_cuenta(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    if texto.lower() == "nan":
        return ""

    if "-" in texto:
        texto = texto.split("-")[0].strip()

    if texto.endswith(".0"):
        texto = texto[:-2]

    return texto


def es_nota_credito(tipo_documento):
    texto_norm = normalizar_texto(tipo_documento)
    return (
        "nota de credito" in texto_norm
        or "nota credito" in texto_norm
        or texto_norm.startswith("nc")
    )


def abreviar_tipo_documento(tipo_documento):
    texto = str(tipo_documento).strip().upper()

    if "DOCUMENTO EQUIVALENTE POS" in texto:
        return "DEPOS"

    if "FACTURA ELECTRÓNICA" in texto or "FACTURA ELECTRONICA" in texto:
        return "FE"

    if "NOTA DE CRÉDITO" in texto or "NOTA DE CREDITO" in texto:
        return "NC"

    return texto[:10]


def redondear(valor):
    return int(round(float(valor), 0))


def normalizar_tarifa_porcentaje(valor):
    if pd.isna(valor):
        return 0.0

    texto = str(valor).strip().replace("%", "").replace(",", ".")
    t = pd.to_numeric(texto, errors="coerce")
    if pd.isna(t):
        return 0.0
    t = float(t)
    if 0 < t <= 1:
        return t * 100
    return t


def buscar_impuesto(tipo, tarifa):
    tarifa = float(tarifa)
    tipo = str(tipo).upper().strip()

    if tipo == "INC":
        filtro = impuestos[
            impuestos["Nombre"].astype(str).str.upper().str.contains("IMPOCONSUMO", na=False)
            & (impuestos["Tarifa"].round(4) == round(tarifa, 4))
        ]
    else:
        filtro = impuestos[
            impuestos["Tipo de impuesto"].astype(str).str.upper().str.contains(tipo, na=False)
            & (impuestos["Tarifa"].round(4) == round(tarifa, 4))
        ]

    if filtro.empty:
        if WEB:
            web_io.error("impuesto_no_encontrado",
                         f"No se encontró el impuesto {tipo} con tarifa {tarifa} en la tabla maestro",
                         [{"tipo": tipo, "tarifa": tarifa}])
        print(f"No encontré impuesto: {tipo} tarifa {tarifa}")
        input("Corrige la tabla maestro de impuestos y presiona Enter...")
        exit()

    fila = filtro.iloc[0]

    return {
        "codigo": fila["Código"],
        "cuenta": limpiar_cuenta(fila["Compras"]),
        "base_minima": fila["Base"]
    }


def validar_cuenta_obsequio(row):
    cuenta = limpiar_cuenta(row.get("Cuenta_ingreso_obsequios", ""))

    if cuenta == "" or cuenta == "0":
        if WEB:
            web_io.error("cuenta_obsequio_faltante",
                         "Hay obsequios pero falta Cuenta_ingreso_obsequios en la parametrización",
                         [{"proveedor": str(row.get("Proveedor", "")), "nit": str(row[col_nit]),
                           "factura": str(row[col_factura])}])
        print("ERROR: Hay obsequios, pero falta Cuenta_ingreso_obsequios en parametrización.")
        print("Proveedor:", row.get("Proveedor", ""))
        print("NIT:", row[col_nit])
        print("Factura:", row[col_factura])
        input("Corrige la parametrización y presiona Enter...")
        exit()

    return cuenta


def validar_cuenta_otros_impuestos(row):
    cuenta = limpiar_cuenta(row.get("Cuenta_otros_impuestos", ""))

    if cuenta == "" or cuenta == "0":
        if WEB:
            web_io.error("cuenta_otros_faltante",
                         "Hay otros impuestos pero falta Cuenta_otros_impuestos en la parametrización",
                         [{"proveedor": str(row.get("Proveedor", "")), "nit": str(row[col_nit]),
                           "factura": str(row[col_factura])}])
        print("ERROR: Hay otros impuestos, pero falta Cuenta_otros_impuestos en parametrización.")
        print("Proveedor:", row.get("Proveedor", ""))
        print("NIT:", row[col_nit])
        print("Factura:", row[col_factura])
        input("Corrige la parametrización y presiona Enter...")
        exit()

    return cuenta


columnas_facturas_obligatorias = [
    col_nit,
    col_factura,
    col_fecha,
    col_tipo_documento,
    col_subtotal_facturas,
    col_descuento,
    col_total
]

for col in columnas_facturas_obligatorias:
    if col not in facturas.columns:
        if WEB:
            web_io.error("columna_faltante", f"Falta la columna '{col}' en la hoja Facturas DIAN", [col])
        print(f"Falta la columna '{col}' en la hoja Facturas DIAN.")
        input("Presiona Enter para cerrar...")
        exit()

columnas_detalle_obligatorias = [
    col_factura,
    col_base_detalle,
    "IVA",
    "% IVA"
]

for col in columnas_detalle_obligatorias:
    if col not in detalle.columns:
        if WEB:
            web_io.error("columna_faltante", f"Falta la columna '{col}' en la hoja Detallado", [col])
        print(f"Falta la columna '{col}' en la hoja Detallado.")
        input("Presiona Enter para cerrar...")
        exit()

columnas_param_obligatorias = [
    col_nit,
    "Cuenta_gasto",
    "Cuenta_por_pagar",
    "Cuenta_otros_impuestos"
]

for col in columnas_param_obligatorias:
    if col not in param.columns:
        if WEB:
            web_io.error("columna_faltante", f"Falta la columna '{col}' en la parametrización de compras (hoja Proveedores)", [col])
        print(f"Falta la columna '{col}' en Parametrizacion_Siigo_Compras.xlsx, hoja Proveedores.")
        input("Presiona Enter para cerrar...")
        exit()

columnas_impuestos_obligatorias = [
    "Código",
    "Tarifa",
    "Compras",
    "Base",
    "Tipo de impuesto",
    "Nombre"
]

for col in columnas_impuestos_obligatorias:
    if col not in impuestos.columns:
        if WEB:
            web_io.error("columna_faltante", f"Falta la columna '{col}' en la Tabla maestro de impuestos", [col])
        print(f"Falta la columna '{col}' en Tabla maestro impuestos.xlsx")
        input("Presiona Enter para cerrar...")
        exit()

facturas[col_nit] = facturas[col_nit].apply(limpiar_nit)
facturas[col_factura] = facturas[col_factura].astype(str).str.strip()
detalle[col_factura] = detalle[col_factura].astype(str).str.strip()
param[col_nit] = param[col_nit].apply(limpiar_nit)

for col in [col_subtotal_facturas, col_descuento, col_recargo, col_total]:
    if col not in facturas.columns:
        facturas[col] = 0
    facturas[col] = pd.to_numeric(facturas[col], errors="coerce").fillna(0)

for col in [
    col_base_detalle,
    "IVA",
    "% IVA",
    "INC",
    "% INC",
    "Bolsas",
    "% Bolsas",
    "ICUI",
    "% ICUI",
    "IC",
    "% IC",
    "ICL",
    "% ICL",
    "IC Porcentual",
    "IBUA",
    "ADV"
]:
    if col not in detalle.columns:
        detalle[col] = 0
    detalle[col] = pd.to_numeric(detalle[col], errors="coerce").fillna(0)

detalle["% IVA"] = detalle["% IVA"].apply(normalizar_tarifa_porcentaje)

for col in [
    "Tarifa_retefuente",
    "Tarifa_reteica",
    "Tarifa_reteiva",
    "Base_minima_reteica_especial"
]:
    if col not in param.columns:
        param[col] = 0
    param[col] = pd.to_numeric(param[col], errors="coerce").fillna(0)

if "Cuenta_ingreso_obsequios" not in param.columns:
    param["Cuenta_ingreso_obsequios"] = ""

impuestos["Tarifa"] = pd.to_numeric(impuestos["Tarifa"], errors="coerce").fillna(0)
impuestos["Base"] = pd.to_numeric(impuestos["Base"], errors="coerce").fillna(0)

nits_facturas = set(facturas[col_nit]) - {""}
nits_param = set(param[col_nit]) - {""}

faltantes = nits_facturas - nits_param

if faltantes:
    if WEB:
        web_io.error("proveedores_faltantes",
                     "Faltan proveedores por parametrizar", sorted(faltantes))
    print("FALTAN PROVEEDORES POR PARAMETRIZAR:")
    for nit in sorted(faltantes):
        print(nit)
    input("Corrige eso antes de continuar...")
    exit()

if not WEB:
    print("Todos los proveedores están parametrizados")

columnas_siigo = [
    "Tipo de comprobante",
    "Consecutivo comprobante",
    "Fecha de elaboración",
    "Sigla moneda",
    "Tasa de cambio",
    "Código cuenta contable",
    "Identificación tercero",
    "Sucursal",
    "Código producto",
    "Código de bodega",
    "Acción",
    "Cantidad producto",
    "Prefijo",
    "Consecutivo",
    "No. cuota",
    "Fecha vencimiento",
    "Código impuesto",
    "Código grupo activo fijo",
    "Código activo fijo",
    "Descripción",
    "Código centro/subcentro de costos",
    "Débito",
    "Crédito",
    "Observaciones",
    "Base gravable libro compras/ventas",
    "Base exenta libro compras/ventas",
    "Mes de cierre"
]

lineas = []
informe_retenciones = []


def agregar_linea(row, cuenta, valor, naturaleza_compra, descripcion, codigo_impuesto=""):
    valor = redondear(valor)

    if valor == 0:
        return

    es_nc = row["Es_Nota_Credito"]

    if naturaleza_compra == "D":
        debito = 0 if es_nc else valor
        credito = valor if es_nc else 0
    else:
        debito = valor if es_nc else 0
        credito = 0 if es_nc else valor

    descripcion = str(descripcion).strip()[:50]

    lineas.append({
        "Tipo de comprobante": row["Tipo_Comprobante_Siigo"],
        "Consecutivo comprobante": row["Consecutivo_Siigo"],
        "Fecha de elaboración": row[col_fecha],
        "Sigla moneda": "",
        "Tasa de cambio": "",
        "Código cuenta contable": limpiar_cuenta(cuenta),
        "Identificación tercero": row[col_nit],
        "Sucursal": "",
        "Código producto": "",
        "Código de bodega": "",
        "Acción": "",
        "Cantidad producto": "",
        "Prefijo": "",
        "Consecutivo": "",
        "No. cuota": "",
        "Fecha vencimiento": "",
        "Código impuesto": codigo_impuesto,
        "Código grupo activo fijo": "",
        "Código activo fijo": "",
        "Descripción": descripcion,
        "Código centro/subcentro de costos": "",
        "Débito": debito,
        "Crédito": credito,
        "Observaciones": "",
        "Base gravable libro compras/ventas": "",
        "Base exenta libro compras/ventas": "",
        "Mes de cierre": ""
    })


def dividir_salida_en_partes_por_comprobante(df_salida, max_lineas=500):
    if df_salida.empty:
        return []

    partes = []
    actual = []
    lineas_actual = 0

    agrupado = df_salida.groupby(["Tipo de comprobante", "Consecutivo comprobante"], sort=False)

    for _, grupo in agrupado:
        n = len(grupo)

        if lineas_actual > 0 and (lineas_actual + n) > max_lineas:
            partes.append(pd.concat(actual, ignore_index=True))
            actual = []
            lineas_actual = 0

        actual.append(grupo)
        lineas_actual += n

    if actual:
        partes.append(pd.concat(actual, ignore_index=True))

    return partes


# Excluir líneas cuyo Concepto sea literalmente "IVA": algunos proveedores
# (ej. ULIFE) codifican el IVA como línea de detalle separada con base=0,
# lo que el motor confundiría con un obsequio.
_concepto_es_iva = (
    detalle["Concepto"].astype(str).str.strip().str.upper() == "IVA"
    if "Concepto" in detalle.columns
    else pd.Series(False, index=detalle.index)
)
detalle["Es_Obsequio"] = (
    (detalle[col_base_detalle] == 0)
    & (detalle["IVA"] > 0)
    & (~_concepto_es_iva)
)

iva_obsequios = (
    detalle[detalle["Es_Obsequio"]]
    .groupby(col_factura)["IVA"]
    .sum()
    .reset_index()
)
iva_obsequios.rename(columns={"IVA": "IVA_OBSEQUIOS"}, inplace=True)

iva_obsequios_19 = (
    detalle[(detalle["Es_Obsequio"]) & (detalle["% IVA"].round(4) == 19.0)]
    .groupby(col_factura)["IVA"]
    .sum()
    .reset_index()
)
iva_obsequios_19.rename(columns={"IVA": "IVA_OBSEQUIOS_19"}, inplace=True)

iva_obsequios_5 = (
    detalle[(detalle["Es_Obsequio"]) & (detalle["% IVA"].round(4) == 5.0)]
    .groupby(col_factura)["IVA"]
    .sum()
    .reset_index()
)
iva_obsequios_5.rename(columns={"IVA": "IVA_OBSEQUIOS_5"}, inplace=True)

detalle_sin_obsequios = detalle[~detalle["Es_Obsequio"]].copy()

iva_19 = (
    detalle_sin_obsequios[detalle_sin_obsequios["% IVA"].round(4) == 19.0]
    .groupby(col_factura)["IVA"]
    .sum()
    .reset_index()
)
iva_19.rename(columns={"IVA": "IVA_19"}, inplace=True)

iva_5 = (
    detalle_sin_obsequios[detalle_sin_obsequios["% IVA"].round(4) == 5.0]
    .groupby(col_factura)["IVA"]
    .sum()
    .reset_index()
)
iva_5.rename(columns={"IVA": "IVA_5"}, inplace=True)

# Líneas con IVA positivo pero %IVA=0 (ej. Colanta). Infiere la tasa del cociente.
_lineas_iva_inf = detalle_sin_obsequios[
    (detalle_sin_obsequios["% IVA"].round(4) == 0)
    & (detalle_sin_obsequios["IVA"] > 0)
    & (detalle_sin_obsequios[col_base_detalle] > 0)
].copy()

if len(_lineas_iva_inf) > 0:
    _lineas_iva_inf["_tasa"] = (
        (_lineas_iva_inf["IVA"] / _lineas_iva_inf[col_base_detalle] * 100).round(0).astype(int)
    )
    _iva_inf_19 = (
        _lineas_iva_inf[_lineas_iva_inf["_tasa"] == 19]
        .groupby(col_factura)["IVA"].sum().reset_index()
    )
    _iva_inf_19.rename(columns={"IVA": "IVA_INF_19"}, inplace=True)
    _iva_inf_5 = (
        _lineas_iva_inf[_lineas_iva_inf["_tasa"] == 5]
        .groupby(col_factura)["IVA"].sum().reset_index()
    )
    _iva_inf_5.rename(columns={"IVA": "IVA_INF_5"}, inplace=True)
else:
    _iva_inf_19 = pd.DataFrame(columns=[col_factura, "IVA_INF_19"]).astype({"IVA_INF_19": float})
    _iva_inf_5 = pd.DataFrame(columns=[col_factura, "IVA_INF_5"]).astype({"IVA_INF_5": float})

inc = (
    detalle_sin_obsequios
    .groupby(col_factura)["INC"]
    .sum()
    .reset_index()
)
inc.rename(columns={"INC": "INC_DET"}, inplace=True)

otros_impuestos_cols = ["Bolsas", "ICUI", "IC", "IC Porcentual", "ICL", "IBUA", "ADV"]
detalle_sin_obsequios["Otros"] = detalle_sin_obsequios[otros_impuestos_cols].fillna(0).sum(axis=1)

otros = (
    detalle_sin_obsequios
    .groupby(col_factura)["Otros"]
    .sum()
    .reset_index()
)

df = facturas.merge(iva_19, on=col_factura, how="left")
df = df.merge(iva_5, on=col_factura, how="left")
df = df.merge(inc, on=col_factura, how="left")
df = df.merge(otros, on=col_factura, how="left")
df = df.merge(iva_obsequios, on=col_factura, how="left")
df = df.merge(iva_obsequios_19, on=col_factura, how="left")
df = df.merge(iva_obsequios_5, on=col_factura, how="left")
df = df.merge(param, on=col_nit, how="left")
df = df.merge(_iva_inf_19, on=col_factura, how="left")
df = df.merge(_iva_inf_5, on=col_factura, how="left")

for col in df.columns:
    if pd.api.types.is_numeric_dtype(df[col]):
        df[col] = df[col].fillna(0)
    else:
        df[col] = df[col].fillna("")

errores_obsequios = []

for _, row in df.iterrows():
    iva_obsequio = redondear(row["IVA_OBSEQUIOS"])
    descuento = redondear(row[col_descuento])

    if iva_obsequio > 0:
        if abs(iva_obsequio - descuento) > 1:
            errores_obsequios.append({
                "Factura": row[col_factura],
                "NIT": row[col_nit],
                "Proveedor": row.get("Razon Social Emisor", row.get("Proveedor", "")),
                "IVA obsequios detectado": iva_obsequio,
                "Descuento detalle": descuento,
                "Diferencia": iva_obsequio - descuento
            })

if errores_obsequios:
    if OBSEQUIOS_MODE == "contabilizar":
        if not WEB:
            print("AVISO: Facturas con IVA obsequios sin descuento detalle (se contabilizan IVA vs CxP directamente):")
            for error in errores_obsequios:
                print(error)
    else:
        if WEB:
            web_io.error("obsequios_descuadre",
                         "Hay diferencias entre IVA obsequios y Descuento detalle", errores_obsequios)
        print("ERROR: Hay diferencias entre IVA obsequios y Descuento detalle.")
        for error in errores_obsequios:
            print(error)
        input("Revisa el reporte antes de continuar...")
        exit()

df["Es_Nota_Credito"] = df[col_tipo_documento].apply(es_nota_credito)

contador_compras = consecutivo_inicial_compras
contador_nc = consecutivo_inicial_nc

tipos_siigo = []
consecutivos_siigo = []

for _, row in df.iterrows():
    if row["Es_Nota_Credito"]:
        tipos_siigo.append(tipo_comprobante_nc)
        consecutivos_siigo.append(contador_nc)
        contador_nc += 1
    else:
        tipos_siigo.append(tipo_comprobante_compras)
        consecutivos_siigo.append(contador_compras)
        contador_compras += 1

df["Tipo_Comprobante_Siigo"] = tipos_siigo
df["Consecutivo_Siigo"] = consecutivos_siigo

for _, row in df.iterrows():
    factura = row[col_factura]
    proveedor = row.get("Razon Social Emisor", row.get("Proveedor", ""))
    tipo_doc = row[col_tipo_documento]
    tipo_doc_abreviado = abreviar_tipo_documento(tipo_doc)

    descuento_header = redondear(row[col_descuento])
    base_gasto = redondear(row[col_subtotal_facturas]) - descuento_header

    iva_19_valor = row["IVA_19"] + row["IVA_INF_19"]
    iva_5_valor = row["IVA_5"] + row["IVA_INF_5"]
    inc_valor = row.get("INC_DET", 0)
    otros_valor = row["Otros"]

    iva_obsequios_valor = row["IVA_OBSEQUIOS"]
    iva_obsequios_19_valor = row["IVA_OBSEQUIOS_19"]
    iva_obsequios_5_valor = row["IVA_OBSEQUIOS_5"]

    iva_total = iva_19_valor + iva_5_valor

    retefuente_valor = 0
    reteica_valor = 0
    reteiva_valor = 0

    base_retefuente = base_gasto
    base_reteica = base_gasto
    base_reteiva = iva_total

    base_minima_retefuente = 0
    base_minima_reteica = 0
    base_minima_reteiva = 0

    descripcion = f"{tipo_doc_abreviado} {factura} {proveedor}"

    if base_gasto > 0:
        agregar_linea(row, row["Cuenta_gasto"], base_gasto, "D", descripcion)

    if iva_19_valor > 0:
        imp = buscar_impuesto("IVA", 19)
        agregar_linea(row, imp["cuenta"], iva_19_valor, "D", f"IVA19 {descripcion}", imp["codigo"])

    if iva_5_valor > 0:
        imp = buscar_impuesto("IVA", 5)
        agregar_linea(row, imp["cuenta"], iva_5_valor, "D", f"IVA5 {descripcion}", imp["codigo"])

    if inc_valor > 0:
        imp = buscar_impuesto("INC", 8)
        agregar_linea(row, imp["cuenta"], inc_valor, "D", f"IMPOCONSUMO {descripcion}", imp["codigo"])

    if otros_valor > 0:
        cuenta_otros = validar_cuenta_otros_impuestos(row)
        agregar_linea(row, cuenta_otros, otros_valor, "D", f"OTROS IMP {descripcion}")

    if iva_obsequios_19_valor > 0:
        imp = buscar_impuesto("IVA", 19)
        agregar_linea(row, imp["cuenta"], iva_obsequios_19_valor, "D", f"IVA19 OBSEQUIO {descripcion}", imp["codigo"])

    if iva_obsequios_5_valor > 0:
        imp = buscar_impuesto("IVA", 5)
        agregar_linea(row, imp["cuenta"], iva_obsequios_5_valor, "D", f"IVA5 OBSEQUIO {descripcion}", imp["codigo"])

    if OBSEQUIOS_MODE == "contabilizar":
        descuento_fila = redondear(row[col_descuento])
        if iva_obsequios_valor > 0 and abs(iva_obsequios_valor - descuento_fila) <= 1:
            # Caso normal: descuento cubre el obsequio -> credito a cuenta ingreso obsequios
            cuenta_obsequios = validar_cuenta_obsequio(row)
            agregar_linea(row, cuenta_obsequios, iva_obsequios_valor, "C", f"OBSEQUIO {descripcion}")
        # Si descuento = 0: solo queda el debito del IVA; el balance cierra contra cuenta_por_pagar
    else:
        if iva_obsequios_valor > 0:
            cuenta_obsequios = validar_cuenta_obsequio(row)
            agregar_linea(row, cuenta_obsequios, iva_obsequios_valor, "C", f"OBSEQUIO {descripcion}")

    if row["Tarifa_retefuente"] > 0:
        imp = buscar_impuesto("RETEFUENTE", row["Tarifa_retefuente"])
        base_minima_retefuente = imp["base_minima"]

        if base_retefuente >= base_minima_retefuente:
            retefuente_valor = base_retefuente * row["Tarifa_retefuente"] / 100
            agregar_linea(row, imp["cuenta"], retefuente_valor, "C", f"RF {descripcion}", imp["codigo"])

    if row["Tarifa_reteica"] > 0:
        imp = buscar_impuesto("RETEICA", row["Tarifa_reteica"])

        base_minima_reteica = row["Base_minima_reteica_especial"]

        if base_minima_reteica == 0:
            base_minima_reteica = imp["base_minima"]

        if base_reteica >= base_minima_reteica:
            reteica_valor = base_reteica * row["Tarifa_reteica"] / 1000
            agregar_linea(row, imp["cuenta"], reteica_valor, "C", f"RICA {descripcion}", imp["codigo"])

    if row["Tarifa_reteiva"] > 0 and base_reteiva > 0:
        imp = buscar_impuesto("RETEIVA", row["Tarifa_reteiva"])
        base_minima_reteiva = imp["base_minima"]

        if base_reteiva >= base_minima_reteiva:
            reteiva_valor = base_reteiva * row["Tarifa_reteiva"] / 100
            agregar_linea(row, imp["cuenta"], reteiva_valor, "C", f"RIVA {descripcion}", imp["codigo"])

    debito_actual = 0
    credito_actual = 0

    for linea in lineas:
        if (
            linea["Tipo de comprobante"] == row["Tipo_Comprobante_Siigo"]
            and linea["Consecutivo comprobante"] == row["Consecutivo_Siigo"]
        ):
            debito_actual += linea["Débito"]
            credito_actual += linea["Crédito"]

    diferencia = debito_actual - credito_actual

    if row["Es_Nota_Credito"]:
        agregar_linea(row, row["Cuenta_por_pagar"], diferencia * -1, "C", descripcion)
    else:
        agregar_linea(row, row["Cuenta_por_pagar"], diferencia, "C", descripcion)

    informe_retenciones.append({
        "Tipo documento": tipo_doc_abreviado,
        "Número factura": factura,
        "Fecha": row[col_fecha],
        "NIT proveedor": row[col_nit],
        "Proveedor": proveedor,

        "Base gasto": redondear(base_gasto),
        "IVA 19": redondear(iva_19_valor),
        "IVA 5": redondear(iva_5_valor),
        "IVA obsequios": redondear(iva_obsequios_valor),
        "IVA total base ReteIVA": redondear(base_reteiva),

        "Tarifa retefuente": row["Tarifa_retefuente"],
        "Base mínima retefuente": redondear(base_minima_retefuente),
        "Base usada retefuente": redondear(base_retefuente),
        "Retefuente aplicada": redondear(retefuente_valor),

        "Tarifa ReteICA": row["Tarifa_reteica"],
        "Base mínima ReteICA": redondear(base_minima_reteica),
        "Base usada ReteICA": redondear(base_reteica),
        "ReteICA aplicada": redondear(reteica_valor),

        "Tarifa ReteIVA": row["Tarifa_reteiva"],
        "Base mínima ReteIVA": redondear(base_minima_reteiva),
        "Base usada ReteIVA": redondear(base_reteiva),
        "ReteIVA aplicada": redondear(reteiva_valor),

        "Total retenciones": (
            redondear(retefuente_valor)
            + redondear(reteica_valor)
            + redondear(reteiva_valor)
        ),

        "Tipo comprobante Siigo": row["Tipo_Comprobante_Siigo"],
        "Consecutivo Siigo": row["Consecutivo_Siigo"]
    })

salida = pd.DataFrame(lineas, columns=columnas_siigo)

revision = salida.groupby(
    ["Tipo de comprobante", "Consecutivo comprobante"]
)[["Débito", "Crédito"]].sum().reset_index()

revision["Diferencia"] = revision["Débito"] - revision["Crédito"]

descuadrados = revision[revision["Diferencia"] != 0]
cuadra = descuadrados.empty

if not WEB:
    if not cuadra:
        print("Hay comprobantes descuadrados:")
        print(descuadrados)
        input("Revisa antes de importar. Presiona Enter para cerrar...")
    else:
        print("Todos los comprobantes cuadran.")

if WEB:
    carpeta_proceso = CTX.out
else:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    carpeta_proceso = os.path.join("Salida", f"Proceso_generado_{timestamp}")
os.makedirs(carpeta_proceso, exist_ok=True)

salida.to_excel(os.path.join(carpeta_proceso, "Salida_Siigo_Compras.xlsx"), index=False)

partes_salida = dividir_salida_en_partes_por_comprobante(salida, max_lineas=500)
for i, parte in enumerate(partes_salida, start=1):
    nombre_parte = os.path.join(carpeta_proceso, f"Salida_Siigo_Compras_Parte_{i:03d}.xlsx")
    parte.to_excel(nombre_parte, index=False)

informe = pd.DataFrame(informe_retenciones)
informe.to_excel(os.path.join(carpeta_proceso, "Informe_Retenciones_Compras.xlsx"), index=False)

if WEB:
    archivos = [{"tipo": "salida", "path": os.path.join(carpeta_proceso, "Salida_Siigo_Compras.xlsx")},
                {"tipo": "informe", "path": os.path.join(carpeta_proceso, "Informe_Retenciones_Compras.xlsx")}]
    for i in range(1, len(partes_salida) + 1):
        archivos.append({"tipo": "parte",
                         "path": os.path.join(carpeta_proceso, f"Salida_Siigo_Compras_Parte_{i:03d}.xlsx")})
    web_io.ok({"comprobantes": int(revision.shape[0]), "lineas": int(len(salida)),
               "cuadra": bool(cuadra), "partes": int(len(partes_salida)),
               "descuadrados": descuadrados.to_dict("records") if not cuadra else []},
              archivos)
else:
    print("")
    print(f"Carpeta del proceso: {carpeta_proceso}")
    print("Archivo generado: Salida\\Proceso_generado_YYYYMMDD_HHMMSS\\Salida_Siigo_Compras.xlsx")
    print(f"Partes generadas (máx 500 líneas sin partir comprobantes): {len(partes_salida)}")
    print("Informe generado: Salida\\Proceso_generado_YYYYMMDD_HHMMSS\\Informe_Retenciones_Compras.xlsx")
    input("Presiona Enter para cerrar...")
