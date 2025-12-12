# ia_arca_emitidos.py
# Conversión de ARCA "Emitidos" -> Formato Holistor (HWVta1modelo)
# AIE San Justo

import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

# --- Rutas de assets ---
HERE = Path(__file__).parent
LOGO = HERE / "logo_aie.png"
FAVICON = HERE / "favicon-aie.ico"

st.set_page_config(
    page_title="ARCA Emitidos → Formato Holistor",
    page_icon=str(FAVICON) if FAVICON.exists() else None,
    layout="centered",
)

if LOGO.exists():
    st.image(str(LOGO), width=180)

st.title("ARCA Emitidos → Formato Holistor")

st.write(
    "Subí el Excel original descargado de **ARCA** "
    "(Libro IVA Digital - Ventas/Emitidos) y descargá un archivo "
    "listo para importar en **Holistor** (según HWVta1modelo)."
)

uploaded = st.file_uploader(
    "Subí el archivo de ARCA (.xlsx)",
    type=["xlsx"],
)

def pick_col(df: pd.DataFrame, *candidates: str) -> str:
    """Devuelve el primer nombre de columna existente en df, según candidatos."""
    cols = set(df.columns)
    for c in candidates:
        if c in cols:
            return c
    raise KeyError(f"No se encontró ninguna de estas columnas: {candidates}")

def map_cpbte_letra(concepto: str):
    """
    Devuelve (Cpbte, Letra) según el texto 'Tipo' de ARCA.
      - Cpbte: F / NC / ND / R
      - Letra: A / B / C (con caso especial heredado: si empieza con '8 ' => letra B)
    """
    concepto = str(concepto).strip()

    # Cpbte
    if "Nota de Crédito" in concepto:
        cpbte = "NC"
    elif "Nota de Débito" in concepto:
        cpbte = "ND"
    elif "Recibo" in concepto:
        cpbte = "R"
    elif "Factura" in concepto:
        cpbte = "F"
    else:
        cpbte = ""

    # Letra
    if concepto.startswith("8 "):
        letra = "B"
    else:
        letra = concepto[-1] if concepto else ""

    return cpbte, letra

def tipo_doc_holistor(v) -> int:
    """
    Tipo Doc:
      - CUIT -> 80
      - DNI  -> 96
      - Si ya viene numérico, lo devuelve como int.
    """
    if pd.isna(v):
        return 0
    s = str(v).strip().upper()

    # Si ya es número (ej: 80 / 96)
    try:
        return int(float(s))
    except Exception:
        pass

    if "CUIT" in s:
        return 80
    if "DNI" in s:
        return 96
    return 0

def get_num(row, col) -> float:
    """Devuelve número limpio (NaN -> 0)."""
    v = row.get(col, 0)
    if pd.isna(v):
        return 0.0
    try:
        return float(v)
    except Exception:
        return 0.0

if uploaded is None:
    st.stop()

# --- LECTURA DEL EXCEL DE ARCA ---
# header=1 porque la fila 2 del archivo suele tener los encabezados reales
df = pd.read_excel(uploaded, sheet_name=0, header=1)

# --- Columnas esperadas (ARCA Emitidos) ---
COL_FECHA = pick_col(df, "Fecha")
COL_TIPO_ARCA = pick_col(df, "Tipo")
COL_PV = pick_col(df, "Punto de Venta", "Pto. Vta.", "Pto Vta")
COL_NRO_DESDE = pick_col(df, "Número Desde", "Numero Desde")
# COL_NRO_HASTA = pick_col(df, "Número Hasta", "Numero Hasta")  # no se usa

COL_TIPO_DOC_REC = pick_col(df, "Tipo Doc. Receptor", "Tipo Doc Receptor")
COL_NRO_DOC_REC = pick_col(df, "Nro. Doc. Receptor", "Nro Doc Receptor", "Nro Doc.", "Nro. Doc.")
COL_NOM_REC = pick_col(df, "Denominación Receptor", "Denominacion Receptor", "Razón Social Receptor", "Razon Social Receptor")

# Montos (ignoramos 0%, 2.5%, 5%)
COL_IVA_105 = pick_col(df, "IVA 10,5%")
COL_NETO_105 = pick_col(df, "Neto Grav. IVA 10,5%")
COL_IVA_21 = pick_col(df, "IVA 21%")
COL_NETO_21 = pick_col(df, "Neto Grav. IVA 21%")
COL_IVA_27 = pick_col(df, "IVA 27%")
COL_NETO_27 = pick_col(df, "Neto Grav. IVA 27%")

COL_NETO_NG = pick_col(df, "Neto No Gravado")
COL_EXENTAS = pick_col(df, "Op. Exentas")
COL_OTROS = pick_col(df, "Otros Tributos")
COL_TOTAL = pick_col(df, "Imp. Total", "Importe Total", "Imp Total")

registros = []

for _, row in df.iterrows():
    concepto = str(row.get(COL_TIPO_ARCA, "")).strip()
    if not concepto:
        continue

    cpbte, letra = map_cpbte_letra(concepto)
    es_nc = (cpbte == "NC")

    # Signo:
    # - NC: negativo
    # - resto: positivo
    def s(valor: float) -> float:
        if valor == 0:
            return 0.0
        return -abs(valor) if es_nc else abs(valor)

    # Importes relevantes del comprobante (para filtrar filas sin montos)
    exng_val = s(get_num(row, COL_NETO_NG) + get_num(row, COL_EXENTAS))
    otros_val = s(get_num(row, COL_OTROS))
    total_val = s(get_num(row, COL_TOTAL))

    # Si no hay montos relevantes (incluyendo alícuotas), ignorar fila
    netos_ivas = [
        s(get_num(row, COL_NETO_105)), s(get_num(row, COL_IVA_105)),
        s(get_num(row, COL_NETO_21)),  s(get_num(row, COL_IVA_21)),
        s(get_num(row, COL_NETO_27)),  s(get_num(row, COL_IVA_27)),
    ]
    if (
        exng_val == 0
        and otros_val == 0
        and total_val == 0
        and all(v == 0 for v in netos_ivas)
    ):
        continue

    # Base (modelo)
    base = {
        "Fecha dd/mm/aaaa": row.get(COL_FECHA),
        "Cpbte": cpbte,                 # F / NC / ND
        "Tipo": letra,                  # A / B / C
        "Suc.": row.get(COL_PV),
        "Número": row.get(COL_NRO_DESDE),
        "Razón Social o Denominación Cliente": row.get(COL_NOM_REC),
        "Tipo Doc.": tipo_doc_holistor(row.get(COL_TIPO_DOC_REC)),
        "CUIT": row.get(COL_NRO_DOC_REC),
        "Domicilio": "",
        "C.P.": "",
        "Pcia": "",
        "Cond Fisc": "RI" if letra == "A" else "MT",
        "Cód. Neto": "",                # manual
        "Cód. NG/EX": "",               # manual
        "Cód. P/R": "",                 # manual
        "Pcia P/R": "",
    }

    filas_comp = []

    # Alícuotas consideradas: 10,5% / 21% / 27%
    aliquotas = [
        (10.5, COL_NETO_105, COL_IVA_105),
        (21.0, COL_NETO_21, COL_IVA_21),
        (27.0, COL_NETO_27, COL_IVA_27),
    ]

    for aliq_val, col_neto, col_iva in aliquotas:
        neto = s(get_num(row, col_neto))
        iva = s(get_num(row, col_iva))

        # Ignorar filas sin montos por alícuota
        if neto == 0 and iva == 0:
            continue

        rec = base.copy()
        rec["Neto Gravado"] = neto
        rec["Alíc."] = aliq_val
        rec["IVA Liquidado"] = iva
        rec["IVA Débito"] = iva
        rec["Conceptos NG/EX"] = 0.0
        rec["Perc./Ret."] = 0.0
        filas_comp.append(rec)

    # Asignar NG/EX y Otros una sola vez (en la 1ra fila del comprobante)
    if filas_comp:
        if exng_val != 0 or otros_val != 0:
            filas_comp[0]["Conceptos NG/EX"] = exng_val
            filas_comp[0]["Perc./Ret."] = otros_val
    else:
        # Caso sin alícuotas:
        # - si hay NG/EX u Otros: los volcamos
        # - si no, pero hay Total (típico caso con solo Imp. Total): mandamos Total a Conceptos NG/EX
        rec = base.copy()
        rec["Neto Gravado"] = 0.0
        rec["Alíc."] = 0.0
        rec["IVA Liquidado"] = 0.0
        rec["IVA Débito"] = 0.0

        if exng_val != 0 or otros_val != 0:
            rec["Conceptos NG/EX"] = exng_val
            rec["Perc./Ret."] = otros_val
        else:
            rec["Conceptos NG/EX"] = total_val
            rec["Perc./Ret."] = 0.0

        filas_comp.append(rec)

    # Total: recalculado (consistente con Recibidos)
    for rec in filas_comp:
        rec["Total"] = (
            float(rec.get("Neto Gravado", 0) or 0)
            + float(rec.get("IVA Liquidado", 0) or 0)
            + float(rec.get("Conceptos NG/EX", 0) or 0)
            + float(rec.get("Perc./Ret.", 0) or 0)
        )
        registros.append(rec)

if not registros:
    st.error("No se encontraron comprobantes con importes.")
    st.stop()

cols_salida = [
    "Fecha dd/mm/aaaa",
    "Cpbte",
    "Tipo",
    "Suc.",
    "Número",
    "Razón Social o Denominación Cliente",
    "Tipo Doc.",
    "CUIT",
    "Domicilio",
    "C.P.",
    "Pcia",
    "Cond Fisc",
    "Cód. Neto",
    "Neto Gravado",
    "Alíc.",
    "IVA Liquidado",
    "IVA Débito",
    "Cód. NG/EX",
    "Conceptos NG/EX",
    "Cód. P/R",
    "Perc./Ret.",
    "Pcia P/R",
    "Total",
]

salida = pd.DataFrame(registros)[cols_salida]

st.subheader("Vista previa de la salida")
st.dataframe(salida.head(50))

# --- GENERAR EXCEL PARA DESCARGAR ---
buffer = BytesIO()
with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
    salida.to_excel(writer, sheet_name="Salida", index=False)

    workbook = writer.book
    worksheet = writer.sheets["Salida"]

    # Formato montos: miles + 2 decimales
    money_format = workbook.add_format({"num_format": "#,##0.00"})

    col_idx = {name: i for i, name in enumerate(salida.columns)}

    # Columnas de importes
    for nombre in ["Neto Gravado", "IVA Liquidado", "IVA Débito", "Conceptos NG/EX", "Perc./Ret.", "Total"]:
        j = col_idx[nombre]
        worksheet.set_column(j, j, 16, money_format)

    # Formato especial para Alíc.: 2 enteros y 3 decimales (ej. 21,000)
    aliq_format = workbook.add_format({"num_format": "00.000"})
    j_aliq = col_idx["Alíc."]
    worksheet.set_column(j_aliq, j_aliq, 8, aliq_format)

    # Ajustes de ancho básicos
    worksheet.set_column(col_idx["Razón Social o Denominación Cliente"], col_idx["Razón Social o Denominación Cliente"], 42)
    worksheet.set_column(col_idx["CUIT"], col_idx["CUIT"], 14)
    worksheet.set_column(col_idx["Cpbte"], col_idx["Cpbte"], 6)
    worksheet.set_column(col_idx["Tipo"], col_idx["Tipo"], 6)
    worksheet.set_column(col_idx["Suc."], col_idx["Suc."], 8)
    worksheet.set_column(col_idx["Número"], col_idx["Número"], 12)

buffer.seek(0)

st.download_button(
    "📥 Descargar Excel procesado",
    data=buffer,
    file_name="Emitidos_salida.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# --- Footer ---
st.markdown(
    "<br><hr style='opacity:0.3'><div style='text-align:center; "
    "font-size:12px; color:#6b7280;'>"
    "© AIE – Herramienta para uso interno | Developer Alfonso Alderete"
    "</div>",
    unsafe_allow_html=True,
)
