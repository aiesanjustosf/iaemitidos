# ia_arca_emitidos.py
# ARCA Emitidos (XLSX o CSV) -> Formato Holistor (HWVta1modelo)
# AIE San Justo

import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
import csv
import re

# ---------------- Matriz interna (CSV) ----------------
# Código -> (Cpbte, Letra)
TIPOS_COMP = {
    "1": ("F", "A"),
    "2": ("ND", "A"),
    "3": ("NC", "A"),
    "4": ("R", "A"),
    # "5": ("", ""),  # NOTAS DE VENTA AL CONTADO A (sin mapeo)

    "6": ("F", "B"),
    "7": ("ND", "B"),
    "8": ("NC", "B"),
    "9": ("R", "B"),

    "11": ("F", "C"),
    "12": ("ND", "C"),
    "13": ("NC", "C"),
    "15": ("R", "C"),

    "51": ("F", "M"),
    "52": ("ND", "M"),
    "53": ("NC", "M"),
    "54": ("R", "M"),

    # FCE / MiPyME
    "201": ("FP", "A"),
    "202": ("NP", "A"),
    "203": ("PC", "A"),

    "206": ("FP", "B"),
    "207": ("NP", "B"),
    "208": ("PC", "B"),
    "211": ("FP", "C"),
    "212": ("NP", "C"),
    "213": ("PC", "C"),
}

# Créditos => importes negativos
CREDITOS = {"NC", "PC"}

# ---------------- Paths / assets ----------------
HERE = Path(__file__).parent

def first_existing(paths):
    for p in paths:
        if p.exists():
            return p
    return None

LOGO_PATH = first_existing([HERE / "logo_aie.png", HERE / "assets" / "logo_aie.png"])
FAVICON_PATH = first_existing([HERE / "favicon-aie.ico", HERE / "assets" / "favicon-aie.ico"])

# ---------------- UI ----------------
st.set_page_config(
    page_title="ARCA Emitidos → Formato Holistor",
    page_icon=str(FAVICON_PATH) if FAVICON_PATH else None,
    layout="centered",
)

if LOGO_PATH:
    st.image(str(LOGO_PATH), width=180)

st.title("ARCA Emitidos → Formato Holistor")

uploaded = st.file_uploader("Subí ARCA Emitidos (.xlsx o .csv)", type=["xlsx", "csv"])
if uploaded is None:
    st.stop()

# ---------------- Helpers ----------------
def sniff_delimiter(text: str) -> str:
    try:
        d = csv.Sniffer().sniff(text[:5000], delimiters=";,|\t")
        return d.delimiter
    except Exception:
        return ";"

def read_arca(file) -> tuple[pd.DataFrame, str]:
    name = (file.name or "").lower()
    if name.endswith(".csv"):
        raw = file.getvalue().decode("utf-8", errors="replace")
        sep = sniff_delimiter(raw)
        df = pd.read_csv(BytesIO(file.getvalue()), sep=sep, dtype=str, encoding="utf-8")
        return df, "csv"
    try:
        return pd.read_excel(file, sheet_name=0, header=1, dtype=object), "xlsx"
    except Exception:
        return pd.read_excel(file, sheet_name=0, header=0, dtype=object), "xlsx"

def pick_col(df: pd.DataFrame, *cands: str) -> str:
    cols = set(df.columns)
    for c in cands:
        if c in cols:
            return c
    raise KeyError(f"No se encontró ninguna de estas columnas: {cands}")

def parse_amount(v) -> float:
    if v is None:
        return 0.0
    if isinstance(v, float) and pd.isna(v):
        return 0.0
    if isinstance(v, (int, float)) and not pd.isna(v):
        return float(v)
    s = str(v).strip()
    if not s:
        return 0.0
    s = s.replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def digits_only(v) -> str:
    if v is None:
        return ""
    return re.sub(r"\D+", "", str(v))

def tipo_doc(v) -> int:
    """
    CSV suele venir 80/96 numérico.
    XLSX puede venir como texto 'CUIT'/'DNI' o numérico.
    """
    if v is None:
        return 0
    if isinstance(v, float) and pd.isna(v):
        return 0
    s = str(v).strip().upper()
    if not s:
        return 0
    try:
        return int(float(s))
    except Exception:
        pass
    if "CUIT" in s:
        return 80
    if "DNI" in s:
        return 96
    return 0

def format_ddmmyyyy(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    s = str(v).strip()
    if not s:
        return ""
    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(dt):
        return s
    return dt.strftime("%d/%m/%Y")

def map_tipo_from_text(desc: str) -> tuple[str, str]:
    """Para XLSX cuando viene texto tipo '1 - Factura A' / 'Nota de Crédito B'."""
    s = str(desc or "").strip()
    su = s.upper()

    if "NOTA DE CRÉDITO" in su or "NOTA DE CREDITO" in su:
        t = "NC"
    elif "NOTA DE DÉBITO" in su or "NOTA DE DEBITO" in su:
        t = "ND"
    elif "RECIBO" in su:
        t = "R"
    elif "FACTURA" in su:
        t = "F"
    else:
        t = ""

    letra = s[-1].upper() if s else ""
    if letra not in ("A", "B", "C", "M"):
        letra = ""

    # regla heredada
    if s.startswith("8 ") and s.strip().upper().endswith("C"):
        letra = "B"

    return t, letra

def decode_csv_tipo(tipo_comp_raw: str) -> tuple[str, str]:
    """CSV trae código."""
    k = str(tipo_comp_raw).strip()
    try:
        k = str(int(float(k)))
    except Exception:
        pass
    return TIPOS_COMP.get(k, ("", ""))

# ---------------- Main ----------------
df, kind = read_arca(uploaded)

# Columnas base
COL_FECHA = pick_col(df, "Fecha de Emisión", "Fecha", "Fecha de Emision")
COL_TIPO_COMP = pick_col(df, "Tipo de Comprobante", "Tipo")
COL_PV = pick_col(df, "Punto de Venta", "Pto. Vta.", "Pto Vta", "Punto Venta")
COL_NRO_DESDE = pick_col(df, "Número Desde", "Numero Desde")

COL_TIPO_DOC_REC = pick_col(df, "Tipo Doc. Receptor", "Tipo Doc Receptor")
COL_NRO_DOC_REC = pick_col(df, "Nro. Doc. Receptor", "Nro Doc Receptor", "Nro Doc.", "Nro. Doc.")
COL_NOM_REC = pick_col(df, "Denominación Receptor", "Denominacion Receptor")

# Montos (solo 10,5 / 21 / 27)
COL_IVA_105 = pick_col(df, "IVA 10,5%")
COL_NETO_105 = pick_col(df, "Imp. Neto Gravado IVA 10,5%", "Neto Grav. IVA 10,5%")
COL_IVA_21 = pick_col(df, "IVA 21%")
COL_NETO_21 = pick_col(df, "Imp. Neto Gravado IVA 21%", "Neto Grav. IVA 21%")
COL_IVA_27 = pick_col(df, "IVA 27%")
COL_NETO_27 = pick_col(df, "Imp. Neto Gravado IVA 27%", "Neto Grav. IVA 27%")

COL_NETO_NG = pick_col(df, "Imp. Neto No Gravado", "Neto No Gravado")
COL_EXENTAS = pick_col(df, "Imp. Op. Exentas", "Op. Exentas")
COL_OTROS = pick_col(df, "Otros Tributos")
COL_TOTAL = pick_col(df, "Imp. Total")

registros = []

for _, row in df.iterrows():
    tipo_comp_raw = row.get(COL_TIPO_COMP, "")
    if tipo_comp_raw is None or (isinstance(tipo_comp_raw, str) and not tipo_comp_raw.strip()):
        continue

    # Tipo comprobante + letra
    if kind == "csv":
        cpbte, letra = decode_csv_tipo(tipo_comp_raw)
    else:
        cpbte, letra = map_tipo_from_text(tipo_comp_raw)

    es_credito = (cpbte in CREDITOS)

    def sg(x: float) -> float:
        if x == 0:
            return 0.0
        return -abs(x) if es_credito else abs(x)

    # Tipo doc / nro doc
    tdoc = tipo_doc(row.get(COL_TIPO_DOC_REC))
    nro_doc = digits_only(row.get(COL_NRO_DOC_REC))

    # CUIT salida + Condición fiscal (sin MT)
    cuit_out = nro_doc
    cond_fisc = ""

    if letra == "A" and tdoc == 80:
        cond_fisc = "RI"

    elif letra == "B" and tdoc == 80:
        cond_fisc = "EX"

    elif letra == "B" and tdoc == 96 and nro_doc:
        dni8 = nro_doc.zfill(8)
        cuit_out = f"00-{dni8}-0"
        cond_fisc = "CF"

    # Importes
    exng_val = sg(parse_amount(row.get(COL_NETO_NG)) + parse_amount(row.get(COL_EXENTAS)))
    otros_val = sg(parse_amount(row.get(COL_OTROS)))
    total_val = sg(parse_amount(row.get(COL_TOTAL)))

    netos_ivas = [
        sg(parse_amount(row.get(COL_NETO_105))), sg(parse_amount(row.get(COL_IVA_105))),
        sg(parse_amount(row.get(COL_NETO_21))),  sg(parse_amount(row.get(COL_IVA_21))),
        sg(parse_amount(row.get(COL_NETO_27))),  sg(parse_amount(row.get(COL_IVA_27))),
    ]

    # Ignorar filas sin montos
    if exng_val == 0 and otros_val == 0 and total_val == 0 and all(v == 0 for v in netos_ivas):
        continue

    base = {
        "Fecha dd/mm/aaaa": format_ddmmyyyy(row.get(COL_FECHA)),
        "Cpbte": cpbte,
        "Tipo": letra,
        "Suc.": row.get(COL_PV),
        "Número": row.get(COL_NRO_DESDE),
        "Razón Social o Denominación Cliente": row.get(COL_NOM_REC),
        "Tipo Doc.": tdoc,
        "CUIT": cuit_out,
        "Domicilio": "",
        "C.P.": "",
        "Pcia": "",
        "Cond Fisc": cond_fisc,
        "Cód. Neto": "",
        "Cód. NG/EX": "",
        "Cód. P/R": "",
        "Pcia P/R": "",
    }

    filas_comp = []
    aliquotas = [
        (10.5, COL_NETO_105, COL_IVA_105),
        (21.0, COL_NETO_21, COL_IVA_21),
        (27.0, COL_NETO_27, COL_IVA_27),
    ]

    for aliq_val, col_neto, col_iva in aliquotas:
        neto = sg(parse_amount(row.get(col_neto)))
        iva = sg(parse_amount(row.get(col_iva)))
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

    # NG/EX y Otros: en la 1ra fila
    if filas_comp:
        if exng_val != 0 or otros_val != 0:
            filas_comp[0]["Conceptos NG/EX"] = exng_val
            filas_comp[0]["Perc./Ret."] = otros_val
    else:
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

# ---------------- Export ----------------
buffer = BytesIO()
with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
    salida.to_excel(writer, sheet_name="Salida", index=False)

    wb = writer.book
    ws = writer.sheets["Salida"]

    money_fmt = wb.add_format({"num_format": "#,##0.00"})
    aliq_fmt = wb.add_format({"num_format": "00.000"})

    col_idx = {c: i for i, c in enumerate(salida.columns)}

    ws.set_column(col_idx["Fecha dd/mm/aaaa"], col_idx["Fecha dd/mm/aaaa"], 12)
    ws.set_column(col_idx["Cpbte"], col_idx["Cpbte"], 6)
    ws.set_column(col_idx["Tipo"], col_idx["Tipo"], 6)
    ws.set_column(col_idx["Suc."], col_idx["Suc."], 8)
    ws.set_column(col_idx["Número"], col_idx["Número"], 12)
    ws.set_column(col_idx["Razón Social o Denominación Cliente"], col_idx["Razón Social o Denominación Cliente"], 42)
    ws.set_column(col_idx["CUIT"], col_idx["CUIT"], 16)

    for nombre in ["Neto Gravado", "IVA Liquidado", "IVA Débito", "Conceptos NG/EX", "Perc./Ret.", "Total"]:
        ws.set_column(col_idx[nombre], col_idx[nombre], 16, money_fmt)

    ws.set_column(col_idx["Alíc."], col_idx["Alíc."], 8, aliq_fmt)

buffer.seek(0)

st.download_button(
    "📥 Descargar Excel procesado",
    data=buffer,
    file_name="Emitidos_salida.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.markdown(
    "<br><hr style='opacity:0.3'><div style='text-align:center; font-size:12px; color:#6b7280;'>"
    "© AIE – Herramienta para uso interno | Developer Alfonso Alderete"
    "</div>",
    unsafe_allow_html=True,
)
