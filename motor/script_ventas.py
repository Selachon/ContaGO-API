import pandas as pd
import os
import re
import sys
import unicodedata
from datetime import datetime

import web_io

CTX = web_io.init()
WEB = CTX.web
PARAMS = CTX.params


def renombrar_si_existe(df, mapeo):
    renames = {}
    for actual in df.columns:
        clave = str(actual).strip()
        if clave in mapeo:
            renames[actual] = mapeo[clave]
    if renames:
        df.rename(columns=renames, inplace=True)


def usar_columna_preferida(df, destino, preferida, alternativa=None):
    if preferida in df.columns:
        if destino in df.columns:
            base = df[preferida]
            if alternativa and alternativa in df.columns:
                base = base.fillna(df[alternativa])
            df[destino] = base
        else:
            df[destino] = df[preferida]


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
        return pd.read_excel(path, sheet_name=hoja, header=1)

    headers = [str(x).strip() if pd.notna(x) else "" for x in bruto.iloc[fila_header].tolist()]
    df = bruto.iloc[fila_header + 1 :].copy()
    df.columns = headers
    df = df.loc[:, [c for c in df.columns if str(c).strip() != ""]]
    return df.reset_index(drop=True)

# =========================
# INPUTS
# =========================

if WEB:
    archivo_dian = CTX.dian
    archivo_param = CTX.param_ventas
    archivo_param_compras = CTX.param_compras
    archivo_impuestos = CTX.impuestos
else:
    if len(sys.argv) > 1 and str(sys.argv[1]).strip() != "":
        archivo_dian = str(sys.argv[1]).strip()
    else:
        archivo_dian = ""
    archivo_param = "Parametrizacion_Siigo_Ventas.xlsx"
    archivo_param_compras = os.path.join("..", "Compras", "Parametrizacion_Siigo_Compras.xlsx")
    archivo_impuestos = os.path.join("..", "Tabla maestro impuestos.xlsx")

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


if not WEB:
    archivo_dian = seleccionar_archivo_entrada(archivo_dian)

def extraer_prefijo_y_consecutivo(numero_factura):
    texto = str(numero_factura).strip()
    m = re.match(r"^([A-Za-z]+)\s*[-_\/]?\s*(\d+)$", texto)
    if m:
        return m.group(1).upper(), int(m.group(2))

    pref = ""
    cons = ""
    for ch in texto:
        if ch.isalpha() and cons == "":
            pref += ch
        elif ch.isdigit():
            cons += ch

    if pref and cons:
        return pref.upper(), int(cons)

    return "SINPREF", 0

# =========================
# CARGA DE ARCHIVOS
# =========================

facturas = cargar_hoja_dian_con_header_dinamico(archivo_dian, "Facturas DIAN")
detalle = cargar_hoja_dian_con_header_dinamico(archivo_dian, "Detallado")
param = pd.read_excel(archivo_param, sheet_name="Proveedores")
param_compras = pd.read_excel(archivo_param_compras, sheet_name="Proveedores")
impuestos = pd.read_excel(archivo_impuestos)

facturas.columns = facturas.columns.astype(str).str.strip()
detalle.columns = detalle.columns.astype(str).str.strip()
param.columns = param.columns.astype(str).str.strip()
param_compras.columns = param_compras.columns.astype(str).str.strip()
impuestos.columns = impuestos.columns.astype(str).str.strip()

facturas["NIT_Emisor_Original"] = facturas.get("NIT Emisor", "")
facturas["Razon_Emisor_Original"] = facturas.get("Razón Social Emisor", facturas.get("Razon Social Emisor", ""))

# Ajuste de formato DIAN nuevo -> formato interno esperado
usar_columna_preferida(
    facturas,
    "NIT Emisor",
    "NIT Receptor",
    "NIT Emisor",
)

usar_columna_preferida(
    facturas,
    "Razon Social Emisor",
    "Razón Social Receptor",
    "Razón Social Emisor",
)

renombrar_si_existe(
    facturas,
    {
        "Tipo documento": "Tipo de documento",
        "Número factura": "Numero Factura",
        "Fecha": "Fecha de emision",
        "Subtotal": "Subtotal antes de impuestos",
        "Descuento": "Descuento detalle",
        "Total": "Valor total",
        "Razón Social Emisor": "Razon Social Emisor",
    },
)

mapeo_normalizado_facturas = {
    "tipo documento": "Tipo de documento",
    "numero factura": "Numero Factura",
    "nit emisor": "NIT Emisor",
    "nit receptor": "NIT Receptor",
    "razon social emisor": "Razon Social Emisor",
    "razon social receptor": "Razón Social Receptor",
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
    },
)

renames_det = {}
for actual in detalle.columns:
    if normalizar_texto(actual) == "numero factura":
        renames_det[actual] = "Numero Factura"
if renames_det:
    detalle.rename(columns=renames_det, inplace=True)

renombrar_si_existe(
    param,
    {
        "NIT Receptor": "NIT Emisor",
        "Cuenta_ingreso": "Cuenta_gasto",
        "Cuenta_por_cobrar": "Cuenta_por_pagar",
        "Cuenta_devolucion_ingreso": "Cuenta_ingreso_obsequios",
    },
)

renombrar_si_existe(
    param_compras,
    {
        "NIT Receptor": "NIT Emisor",
    },
)

# =========================
# COLUMNAS
# =========================

col_nit = "NIT Emisor"
col_factura = "Numero Factura"
col_total = "Valor total"
col_fecha = "Fecha de emision"
col_tipo_documento = "Tipo de documento"

col_subtotal_facturas = "Subtotal antes de impuestos"
col_base_detalle = "Base del impuesto"

col_descuento = "Descuento detalle"
col_recargo = "Recargo detalle"

# =========================
# FUNCIONES
# =========================

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
    texto = normalizar_texto(tipo_documento)
    return "nota de credito" in texto or "nota credito" in texto


def es_documento_soporte(tipo_documento):
    return "documento soporte" in normalizar_texto(tipo_documento)


def obtener_valor_param(row, campo):
    if row.get("Es_Documento_Soporte", False):
        return row.get(f"PC_{campo}", "")
    return row.get(f"PV_{campo}", "")


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


def obtener_texto_campo(row, *campos):
    for campo in campos:
        if campo in row.index:
            valor = row[campo]
            if isinstance(valor, pd.Series):
                valor = valor.iloc[0] if not valor.empty else ""
            if pd.isna(valor):
                continue
            texto = str(valor).strip()
            if texto and texto.lower() != "nan":
                return texto
    return ""


def normalizar_tarifa_porcentaje(valor):
    t = float(valor)
    if 0 < t <= 1:
        return t * 100
    return t


def buscar_impuesto(tipo, tarifa, es_nc=False, es_documento_soporte=False):
    tarifa = normalizar_tarifa_porcentaje(tarifa)
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

    if es_documento_soporte:
        cuenta_base = "Compras"
    else:
        cuenta_base = "Devolución ventas" if es_nc and "Devolución ventas" in fila.index else "Ventas"

    return {
        "codigo": fila["Código"],
        "cuenta": limpiar_cuenta(fila[cuenta_base]),
        "base_minima": fila["Base"]
    }


def validar_cuenta_obsequio(row):
    cuenta = limpiar_cuenta(obtener_valor_param(row, "Cuenta_ingreso_obsequios"))

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
    cuenta = limpiar_cuenta(obtener_valor_param(row, "Cuenta_otros_impuestos"))

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


# =========================
# VALIDAR COLUMNAS NECESARIAS
# =========================

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
            web_io.error("columna_faltante", f"Falta la columna '{col}' en la parametrización de ventas (hoja Proveedores)", [col])
        print(f"Falta la columna '{col}' en Parametrizacion_Siigo_Ventas.xlsx, hoja Proveedores.")
        input("Presiona Enter para cerrar...")
        exit()

columnas_impuestos_obligatorias = [
    "Código",
    "Tarifa",
    "Ventas",
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

# =========================
# LIMPIEZA BÁSICA
# =========================

facturas[col_nit] = facturas[col_nit].apply(limpiar_nit)
facturas[col_factura] = facturas[col_factura].astype(str).str.strip()

detalle[col_factura] = detalle[col_factura].astype(str).str.strip()

param[col_nit] = param[col_nit].apply(limpiar_nit)
param_compras[col_nit] = param_compras[col_nit].apply(limpiar_nit)

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
    "% ICL"
]:
    if col not in detalle.columns:
        detalle[col] = 0
    detalle[col] = pd.to_numeric(detalle[col], errors="coerce").fillna(0)

for col in [
    "Tarifa_retefuente",
    "Tarifa_reteica",
    "Tarifa_reteiva",
    "Base_minima_reteica_especial"
]:
    if col not in param.columns:
        param[col] = 0
    param[col] = pd.to_numeric(param[col], errors="coerce").fillna(0)

for col in [
    "Tarifa_retefuente",
    "Tarifa_reteica",
    "Tarifa_reteiva",
    "Base_minima_reteica_especial"
]:
    if col not in param_compras.columns:
        param_compras[col] = 0
    param_compras[col] = pd.to_numeric(param_compras[col], errors="coerce").fillna(0)

if "Cuenta_ingreso_obsequios" not in param.columns:
    param["Cuenta_ingreso_obsequios"] = ""

if "Cuenta_ingreso_obsequios" not in param_compras.columns:
    param_compras["Cuenta_ingreso_obsequios"] = ""

impuestos["Tarifa"] = pd.to_numeric(impuestos["Tarifa"], errors="coerce").fillna(0)
impuestos["Base"] = pd.to_numeric(impuestos["Base"], errors="coerce").fillna(0)

# =========================
# VALIDAR PROVEEDORES
# =========================

nits_facturas = set(
    facturas.loc[
        ~facturas[col_tipo_documento].apply(es_documento_soporte),
        col_nit
    ]
) - {""}
nits_param = set(param[col_nit]) - {""}

faltantes = nits_facturas - nits_param

if faltantes:
    if WEB:
        web_io.error("proveedores_faltantes",
                     "Faltan proveedores por parametrizar (excluye Documento Soporte)",
                     sorted(faltantes))
    print("FALTAN PROVEEDORES POR PARAMETRIZAR (excluye Documento Soporte):")
    for nit in sorted(faltantes):
        print(nit)
    input("Corrige eso antes de continuar...")
    exit()

if not WEB:
    print("Todos los proveedores están parametrizados")

# =========================
# COLUMNAS SIIGO
# =========================

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


# =========================
# DETECTAR OBSEQUIOS
# =========================

detalle["Es_Obsequio"] = (
    (detalle[col_base_detalle] == 0)
    & (detalle["IVA"] > 0)
)

iva_obsequios = (
    detalle[detalle["Es_Obsequio"]]
    .groupby(col_factura)["IVA"]
    .sum()
    .reset_index()
)
iva_obsequios.rename(columns={"IVA": "IVA_OBSEQUIOS"}, inplace=True)

iva_obsequios_19 = (
    detalle[(detalle["Es_Obsequio"]) & (detalle["% IVA"] == 19)]
    .groupby(col_factura)["IVA"]
    .sum()
    .reset_index()
)
iva_obsequios_19.rename(columns={"IVA": "IVA_OBSEQUIOS_19"}, inplace=True)

iva_obsequios_5 = (
    detalle[(detalle["Es_Obsequio"]) & (detalle["% IVA"] == 5)]
    .groupby(col_factura)["IVA"]
    .sum()
    .reset_index()
)
iva_obsequios_5.rename(columns={"IVA": "IVA_OBSEQUIOS_5"}, inplace=True)

detalle_sin_obsequios = detalle[~detalle["Es_Obsequio"]].copy()

# =========================
# AGRUPAR IMPUESTOS DEL DETALLE
# =========================

iva_19 = (
    detalle_sin_obsequios[
        detalle_sin_obsequios["% IVA"].apply(normalizar_tarifa_porcentaje).round(4) == 19
    ]
    .groupby(col_factura)["IVA"]
    .sum()
    .reset_index()
)
iva_19.rename(columns={"IVA": "IVA_19"}, inplace=True)

iva_5 = (
    detalle_sin_obsequios[
        detalle_sin_obsequios["% IVA"].apply(normalizar_tarifa_porcentaje).round(4) == 5
    ]
    .groupby(col_factura)["IVA"]
    .sum()
    .reset_index()
)
iva_5.rename(columns={"IVA": "IVA_5"}, inplace=True)

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

# =========================
# UNIR INFORMACIÓN
# =========================

df = facturas.merge(iva_19, on=col_factura, how="left")
df = df.merge(iva_5, on=col_factura, how="left")
df = df.merge(inc, on=col_factura, how="left")
df = df.merge(otros, on=col_factura, how="left")
df = df.merge(iva_obsequios, on=col_factura, how="left")
df = df.merge(iva_obsequios_19, on=col_factura, how="left")
df = df.merge(iva_obsequios_5, on=col_factura, how="left")

df["Es_Documento_Soporte"] = df[col_tipo_documento].apply(es_documento_soporte)
df["NIT_Contable"] = df[col_nit].apply(limpiar_nit)
df.loc[df["Es_Documento_Soporte"], "NIT_Contable"] = df.loc[df["Es_Documento_Soporte"], "NIT_Emisor_Original"].apply(limpiar_nit)
df[col_nit] = df["NIT_Contable"]

param_v = param.add_prefix("PV_").rename(columns={"PV_NIT Emisor": "NIT_Contable"})
param_c = param_compras.add_prefix("PC_").rename(columns={"PC_NIT Emisor": "NIT_Contable"})

df = df.merge(param_v, on="NIT_Contable", how="left")
df = df.merge(param_c, on="NIT_Contable", how="left")

faltantes_soporte = set(df.loc[df["Es_Documento_Soporte"], "NIT_Contable"]) - set(param_compras[col_nit])
faltantes_soporte = {x for x in faltantes_soporte if x}
if faltantes_soporte:
    if WEB:
        web_io.error("terceros_soporte_faltantes",
                     "Faltan terceros de Documento Soporte en la parametrización de compras",
                     sorted(faltantes_soporte))
    print("FALTAN TERCEROS DE DOCUMENTO SOPORTE EN PARAMETRIZACION DE COMPRAS:")
    for nit in sorted(faltantes_soporte):
        print(nit)
    input("Corrige eso antes de continuar...")
    exit()

for col in df.columns:
    if pd.api.types.is_numeric_dtype(df[col]):
        df[col] = df[col].fillna(0)
    else:
        df[col] = df[col].fillna("")

# =========================
# VALIDAR OBSEQUIOS VS DESCUENTO DETALLE
# =========================

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
    if WEB:
        web_io.error("obsequios_descuadre",
                     "Hay diferencias entre IVA obsequios y Descuento detalle", errores_obsequios)
    print("ERROR: Hay diferencias entre IVA obsequios y Descuento detalle.")
    for error in errores_obsequios:
        print(error)
    input("Revisa el reporte antes de continuar...")
    exit()

# =========================
# TIPO Y CONSECUTIVO SEGÚN DOCUMENTO
# =========================

df["Es_Nota_Credito"] = df[col_tipo_documento].apply(es_nota_credito)

df["Prefijo_Factura"] = df[col_factura].apply(lambda x: extraer_prefijo_y_consecutivo(x)[0])
df["Consecutivo_Factura"] = df[col_factura].apply(lambda x: extraer_prefijo_y_consecutivo(x)[1])

prefijos_factura = sorted(df.loc[(~df["Es_Nota_Credito"]) & (~df["Es_Documento_Soporte"]), "Prefijo_Factura"].unique().tolist())
prefijos_nc = sorted(df.loc[df["Es_Nota_Credito"], "Prefijo_Factura"].unique().tolist())
prefijos_ds = sorted(df.loc[df["Es_Documento_Soporte"], "Prefijo_Factura"].unique().tolist())

mapa_tipo_factura = {}
mapa_tipo_nc = {}
mapa_tipo_ds = {}

if WEB:
    # Fase 1: si solo se piden los prefijos detectados, se emiten y se termina.
    if CTX.solo_prefijos:
        web_io.prefijos(prefijos_factura, prefijos_ds, prefijos_nc)
    # Fase 2: el mapeo prefijo -> tipo de comprobante llega en los parámetros.
    mapa_tipo_factura = {str(k): str(v) for k, v in PARAMS.get("tipos_factura", {}).items()}
    mapa_tipo_ds = {str(k): str(v) for k, v in PARAMS.get("tipos_ds", {}).items()}
    mapa_tipo_nc = {str(k): str(v) for k, v in PARAMS.get("tipos_nc", {}).items()}
else:
    print("\nConfiguración de tipo de comprobante por prefijo (VENTAS):")
    for pref in prefijos_factura:
        mapa_tipo_factura[pref] = input(f"Tipo de comprobante para facturas prefijo {pref}: ").strip()

    if prefijos_ds:
        print("\nConfiguración de tipo de comprobante por prefijo (DOCUMENTO SOPORTE):")
        for pref in prefijos_ds:
            mapa_tipo_ds[pref] = input(f"Tipo de comprobante para DS prefijo {pref}: ").strip()

    if prefijos_nc:
        print("\nConfiguración de tipo de comprobante por prefijo (NOTAS CRÉDITO):")
        for pref in prefijos_nc:
            mapa_tipo_nc[pref] = input(f"Tipo de comprobante para NC prefijo {pref}: ").strip()

tipos_siigo = []
consecutivos_siigo = []

for _, row in df.iterrows():
    pref = row["Prefijo_Factura"]
    cons = row["Consecutivo_Factura"]

    if row["Es_Documento_Soporte"]:
        tipos_siigo.append(mapa_tipo_ds.get(pref, mapa_tipo_factura.get(pref, "")))
        consecutivos_siigo.append(cons)
    elif row["Es_Nota_Credito"]:
        tipos_siigo.append(mapa_tipo_nc.get(pref, mapa_tipo_factura.get(pref, "")))
        consecutivos_siigo.append(cons)
    else:
        tipos_siigo.append(mapa_tipo_factura.get(pref, ""))
        consecutivos_siigo.append(cons)

df["Tipo_Comprobante_Siigo"] = tipos_siigo
df["Consecutivo_Siigo"] = consecutivos_siigo

# =========================
# GENERAR ASIENTOS
# =========================

for _, row in df.iterrows():
    factura = row[col_factura]
    receptor = obtener_texto_campo(row, "Razón Social Receptor", "Razon Social Receptor", "Cliente", "Proveedor")
    concepto = obtener_texto_campo(row, "Concepto")
    tipo_doc = row[col_tipo_documento]
    tipo_doc_abreviado = abreviar_tipo_documento(tipo_doc)

    base_gasto = row[col_subtotal_facturas]

    iva_19_valor = row["IVA_19"]
    iva_5_valor = row["IVA_5"]
    inc_valor = row.get("INC_DET", 0)
    otros_valor = pd.to_numeric(row["Otros"], errors="coerce") or 0

    iva_obsequios_valor = row["IVA_OBSEQUIOS"]
    iva_obsequios_19_valor = row["IVA_OBSEQUIOS_19"]
    iva_obsequios_5_valor = row["IVA_OBSEQUIOS_5"]

    iva_total = iva_19_valor + iva_5_valor

    retefuente_valor = 0
    reteica_valor = 0
    reteiva_valor = 0

    tarifa_retefuente = pd.to_numeric(obtener_valor_param(row, "Tarifa_retefuente"), errors="coerce")
    tarifa_reteica = pd.to_numeric(obtener_valor_param(row, "Tarifa_reteica"), errors="coerce")
    tarifa_reteiva = pd.to_numeric(obtener_valor_param(row, "Tarifa_reteiva"), errors="coerce")
    base_minima_reteica_especial = pd.to_numeric(obtener_valor_param(row, "Base_minima_reteica_especial"), errors="coerce")

    tarifa_retefuente = 0 if pd.isna(tarifa_retefuente) else float(tarifa_retefuente)
    tarifa_reteica = 0 if pd.isna(tarifa_reteica) else float(tarifa_reteica)
    tarifa_reteiva = 0 if pd.isna(tarifa_reteiva) else float(tarifa_reteiva)
    base_minima_reteica_especial = 0 if pd.isna(base_minima_reteica_especial) else float(base_minima_reteica_especial)

    base_retefuente = base_gasto
    base_reteica = base_gasto
    base_reteiva = iva_total

    base_minima_retefuente = 0
    base_minima_reteica = 0
    base_minima_reteiva = 0

    descripcion = f"{tipo_doc_abreviado} {factura} {receptor} {concepto}".strip()

    naturaleza_base = "D" if row["Es_Documento_Soporte"] else "C"

    if base_gasto > 0:
        agregar_linea(row, obtener_valor_param(row, "Cuenta_gasto"), base_gasto, naturaleza_base, descripcion)

    if iva_19_valor > 0:
        imp = buscar_impuesto("IVA", 19, row["Es_Nota_Credito"], row.get("Es_Documento_Soporte", False))
        agregar_linea(row, imp["cuenta"], iva_19_valor, naturaleza_base, f"IVA19 {descripcion}", imp["codigo"])

    if iva_5_valor > 0:
        imp = buscar_impuesto("IVA", 5, row["Es_Nota_Credito"], row.get("Es_Documento_Soporte", False))
        agregar_linea(row, imp["cuenta"], iva_5_valor, naturaleza_base, f"IVA5 {descripcion}", imp["codigo"])

    if inc_valor > 0:
        imp = buscar_impuesto("INC", 8, row["Es_Nota_Credito"], row.get("Es_Documento_Soporte", False))
        agregar_linea(row, imp["cuenta"], inc_valor, naturaleza_base, f"IMPOCONSUMO {descripcion}", imp["codigo"])

    if otros_valor > 0:
        cuenta_otros = validar_cuenta_otros_impuestos(row)
        agregar_linea(row, cuenta_otros, otros_valor, naturaleza_base, f"OTROS IMP {descripcion}")

    if iva_obsequios_19_valor > 0:
        imp = buscar_impuesto("IVA", 19, row["Es_Nota_Credito"], row.get("Es_Documento_Soporte", False))
        agregar_linea(
            row,
            imp["cuenta"],
            iva_obsequios_19_valor,
            naturaleza_base,
            f"IVA19 OBSEQUIO {descripcion}",
            imp["codigo"]
        )

    if iva_obsequios_5_valor > 0:
        imp = buscar_impuesto("IVA", 5, row["Es_Nota_Credito"], row.get("Es_Documento_Soporte", False))
        agregar_linea(
            row,
            imp["cuenta"],
            iva_obsequios_5_valor,
            naturaleza_base,
            f"IVA5 OBSEQUIO {descripcion}",
            imp["codigo"]
        )

    if iva_obsequios_valor > 0:
        cuenta_obsequios = validar_cuenta_obsequio(row)
        agregar_linea(
            row,
            cuenta_obsequios,
            iva_obsequios_valor,
            naturaleza_base,
            f"OBSEQUIO {descripcion}"
        )

    if tarifa_retefuente > 0:
        imp = buscar_impuesto("RETEFUENTE", tarifa_retefuente, row["Es_Nota_Credito"], row.get("Es_Documento_Soporte", False))
        base_minima_retefuente = imp["base_minima"]

        if base_retefuente >= base_minima_retefuente:
            retefuente_valor = base_retefuente * tarifa_retefuente / 100
            agregar_linea(row, imp["cuenta"], retefuente_valor, "C", f"RF {descripcion}", imp["codigo"])

    if tarifa_reteica > 0:
        imp = buscar_impuesto("RETEICA", tarifa_reteica, row["Es_Nota_Credito"], row.get("Es_Documento_Soporte", False))

        base_minima_reteica = base_minima_reteica_especial

        if base_minima_reteica == 0:
            base_minima_reteica = imp["base_minima"]

        if base_reteica >= base_minima_reteica:
            reteica_valor = base_reteica * tarifa_reteica / 1000
            agregar_linea(row, imp["cuenta"], reteica_valor, "C", f"RICA {descripcion}", imp["codigo"])

    if tarifa_reteiva > 0 and base_reteiva > 0:
        imp = buscar_impuesto("RETEIVA", tarifa_reteiva, row["Es_Nota_Credito"], row.get("Es_Documento_Soporte", False))
        base_minima_reteiva = imp["base_minima"]

        if base_reteiva >= base_minima_reteiva:
            reteiva_valor = base_reteiva * tarifa_reteiva / 100
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

    diferencia = redondear(abs(debito_actual - credito_actual))

    if diferencia > 0:
        descripcion_corta = str(descripcion).strip()[:50]

        if debito_actual < credito_actual:
            debito_cxc = diferencia
            credito_cxc = 0
        else:
            debito_cxc = 0
            credito_cxc = diferencia

        lineas.append({
            "Tipo de comprobante": row["Tipo_Comprobante_Siigo"],
            "Consecutivo comprobante": row["Consecutivo_Siigo"],
            "Fecha de elaboración": row[col_fecha],
            "Sigla moneda": "",
            "Tasa de cambio": "",
            "Código cuenta contable": limpiar_cuenta(obtener_valor_param(row, "Cuenta_por_pagar")),
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
            "Código impuesto": "",
            "Código grupo activo fijo": "",
            "Código activo fijo": "",
            "Descripción": descripcion_corta,
            "Código centro/subcentro de costos": "",
            "Débito": debito_cxc,
            "Crédito": credito_cxc,
            "Observaciones": "",
            "Base gravable libro compras/ventas": "",
            "Base exenta libro compras/ventas": "",
            "Mes de cierre": ""
        })

    informe_retenciones.append({
        "Tipo documento": tipo_doc_abreviado,
        "Número factura": factura,
        "Fecha": row[col_fecha],
        "NIT proveedor": row[col_nit],
        "Proveedor": receptor,

        "Base gasto": redondear(base_gasto),
        "IVA 19": redondear(iva_19_valor),
        "IVA 5": redondear(iva_5_valor),
        "IVA obsequios": redondear(iva_obsequios_valor),
        "IVA total base ReteIVA": redondear(base_reteiva),

        "Tarifa retefuente": tarifa_retefuente,
        "Base mínima retefuente": redondear(base_minima_retefuente),
        "Base usada retefuente": redondear(base_retefuente),
        "Retefuente aplicada": redondear(retefuente_valor),

        "Tarifa ReteICA": tarifa_reteica,
        "Base mínima ReteICA": redondear(base_minima_reteica),
        "Base usada ReteICA": redondear(base_reteica),
        "ReteICA aplicada": redondear(reteica_valor),

        "Tarifa ReteIVA": tarifa_reteiva,
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

# =========================
# EXPORTAR Y VALIDAR CUADRE
# =========================

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

archivo_salida = os.path.join(carpeta_proceso, "Salida_Siigo_Ventas.xlsx")
archivo_informe = os.path.join(carpeta_proceso, "Informe_Retenciones_Ventas.xlsx")

salida.to_excel(archivo_salida, index=False)

partes_salida = dividir_salida_en_partes_por_comprobante(salida, max_lineas=500)
for i, parte in enumerate(partes_salida, start=1):
    nombre_parte = os.path.join(carpeta_proceso, f"Salida_Siigo_Ventas_Parte_{i:03d}.xlsx")
    parte.to_excel(nombre_parte, index=False)

informe = pd.DataFrame(informe_retenciones)
informe.to_excel(archivo_informe, index=False)

if WEB:
    archivos = [{"tipo": "salida", "path": archivo_salida},
                {"tipo": "informe", "path": archivo_informe}]
    for i in range(1, len(partes_salida) + 1):
        archivos.append({"tipo": "parte",
                         "path": os.path.join(carpeta_proceso, f"Salida_Siigo_Ventas_Parte_{i:03d}.xlsx")})
    web_io.ok({"comprobantes": int(revision.shape[0]), "lineas": int(len(salida)),
               "cuadra": bool(cuadra), "partes": int(len(partes_salida)),
               "descuadrados": descuadrados.to_dict("records") if not cuadra else []},
              archivos)
else:
    print("")
    print(f"Carpeta del proceso: {carpeta_proceso}")
    print(f"Archivo generado: {archivo_salida}")
    print(f"Partes generadas (máx 500 líneas sin partir comprobantes): {len(partes_salida)}")
    print(f"Informe generado: {archivo_informe}")
    input("Presiona Enter para cerrar...")
