# ia_arca_emitidos.py
# Conversión ARCA "Emitidos" (XLSX o CSV) -> Formato Holistor (HWVta1modelo)
# AIE San Justo

import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
import csv
import re
import zipfile
import xml.etree.ElementTree as ET

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
    "Subí el archivo original de **ARCA** (Libro IVA Digital - Ventas/Emitidos) en **XLSX o CSV**.\n\n"
    "- Si subís **CSV**, también subí **TABLAARCA** para decodificar el **Tipo de Comprobante** (código) "
    "a **Cpbte (F/NC/ND/R)** y **Tipo (A/B/C)**.\n"
    "- En el CSV, **Tipo Doc. Receptor** ya viene como **80/96**, se toma directo."
)

uploaded = st.file_uploader("Subí ARCA Emitidos (.xlsx o .csv)", type=["xlsx", "csv"])
tablaarca_file = st.file_uploader("Subí TABLAARCA (.xlsx o .csv)", type=["xlsx", "csv"])

# ---------------- Helpers ----------------

def sniff_delimiter(text: str) -> str:
    try:
        d = csv.Sniffer().sniff(text[:5000], delimiters=";,|\t")
        return d.delimiter
    except Exception:
        return ";"  # típico ARCA

def read_arca_file(file) -> tuple[pd.DataFrame, str]:
    name = (file.name or "").lower()
    if name.endswith(".csv"):
        raw = file.getvalue().decode("utf-8", errors="replace")
        sep = sniff_delimiter(raw)
        df = pd.read_csv(BytesIO(file.getvalue()), sep=sep, dtype=str, encoding="utf-8")
        return df, "csv"

    # XLSX: ARCA suele tener encabezado real en fila 2
    try:
        return pd.read_excel(file, sheet_name=0, header=1), "xlsx"
    except Exception:
        return pd.read_excel(file, sheet_name=0, header=0), "xlsx"

def pick_col(df: pd.DataFrame, *candidates: str) -> str:
    cols = set(df.columns)
    for c in candidates:
        if c in cols:
            return c
    raise KeyError(f"No se encontró ninguna de estas columnas: {candidates}")

def parse_amount(v) -> float:
    """Convierte montos ARCA (ej. '3.679,34') a float."""
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
    # 3.679,34 -> 3679.34
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def tipo_doc_numeric(v) -> int:
    """
    En CSV ARCA viene numérico (80/96). Tomarlo directo.
    Si llega texto no numérico (por XLSX u otro), intenta parsear; si falla devuelve 0.
    """
    if v is None:
        return 0
    if isinstance(v, float) and pd.isna(v):
        return 0
    s = str(v).strip()
    if not s:
        return 0
    try:
        return int(float(s))
    except Exception:
        return 0

# --- TABLAARCA: lectura robusta (incluye tu TABLAARCACGT.xlsx que puede no abrir con openpyxl) ---

def _strip_ns(tag: str) -> str:
    return tag.split("}", 1)[1] if "}" in tag else tag

def read_tablaarca_any(file) -> pd.DataFrame:
    """
    Lee TABLAARCA desde XLSX/CSV. Para XLSX "no estándar", parsea sheet1.xml + sharedStrings.xml.
    Espera columnas: Código | Descripción | Tipo | Letra
    """
    name = (file.name or "").lower()

    if name.endswith(".csv"):
        raw = file.getvalue().decode("utf-8", errors="replace")
        sep = sniff_delimiter(raw)
        return pd.read_csv(BytesIO(file.getvalue()), sep=sep, dtype=str, encoding="utf-8")

    # 1) intento normal
    try:
        return pd.read_excel(file, sheet_name=0, dtype=str)
    except Exception:
        pass

    # 2) fallback XML (xlsx)
    content = file.getvalue()
    zf = zipfile.ZipFile(BytesIO(content))

    shared = []
    sst_xml = zf.read("xl/sharedStrings.xml").decode("utf-8", errors="ignore")
    sst_root = ET.fromstring(sst_xml)
    for si in list(sst_root):
        if _strip_ns(si.tag) != "si":
            continue
        texts = []
        for ch in si.iter():
            if _strip_ns(ch.tag) == "t":
                texts.append(ch.text or "")
        shared.append("".join(texts))

    sheet_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8", errors="ignore")
    root = ET.fromstring(sheet_xml)
    ns = {"s": root.tag.split("}")[0].strip("{")}
    sheetData = root.find("s:sheetData", ns)

    def col_letter_to_index(col: str) -> int:
        idx = 0
        for ch in col:
            idx = idx * 26 + (ord(ch) - 64)
        return idx - 1

    def cell_ref_to_rc(r: str):
        m = re.match(r"([A-Z]+)(\d+)", r)
        return int(m.group(2)) - 1, col_letter_to_index(m.group(1))

    rows = []
    for r in sheetData.findall("s:row", ns):
        cells = {}
        for c in r.findall("s:c", ns):
            ref = c.attrib.get("r", "")
            _, col_idx = cell_ref_to_rc(ref)
            t = c.attrib.get("t")
            v = c.find("s:v", ns)
            if v is None:
                continue
            val = v.text
            if t == "s":
                val = shared[int(val)]
            cells[col_idx] = val

        if not cells:
            continue

        # Tomamos columnas A-D
        rows.append([cells.get(i, "") for i in range(4)])

    # Buscar header "Código" / "Descripción" / "Tipo" / "Letra"
    header_i = None
    for i, r in enumerate(rows):
        if [x.strip().upper() for x in r[:4]] == ["CÓDIGO", "DESCRIPCIÓN", "TIPO", "LETRA"] or \
           [x.strip().upper() for x in r[:4]] == ["CODIGO", "DESCRIPCIÓN", "TIPO", "LETRA"] or \
           [x.strip().upper() for x in r[:4]] == ["CÓDIGO", "DESCRIPCION", "TIPO", "LETRA"] or \
           [x.strip().upper() for x in r[:4]] == ["CODIGO", "DESCRIPCION", "TIPO", "LETRA"]:
            header_i = i
            break

    if header_i is None:
        # fallback: asumir 2da fila como header (como en tu archivo)
        header_i = 1

    data = rows[header_i + 1 :]
    out = pd.DataFrame(data, columns=["Código", "Descripción", "Tipo", "Letra"])
    return out

def build_tabla_lookup(tabla: pd.DataFrame) -> dict:
    """
    Lookup por código: { '1': ('F','A'), '3': ('NC','A'), ... }
    """
    # normalizar nombres posibles
    cols = {c.strip().lower(): c for c in tabla.columns}
    col_cod = cols.get("código") or cols.get("codigo") or cols.get("cód.") or cols.get("cod.")
    col_tipo = cols.get("tipo")
    col_letra = cols.get("letra")

    if not col_cod or not col_tipo or not col_letra:
        raise ValueError("TABLAARCA no tiene columnas esperadas: Código/Tipo/Letra")

    lk = {}
    for _, r in tabla.iterrows():
        k = str(r.get(col_cod, "")).strip()
        t = str(r.get(col_tipo, "")).strip().upper()
        l = str(r.get(col_letra, "")).strip().upper()
        if not k or not t:
            continue
        try:
            k = str(int(float(k)))
        except Exception:
            pass

        # Normalizar recibos
        if t == "RC":
            t = "R"

        lk[k] = (t, l)
    return lk

def map_tipo_from_xlsx_text(desc: str) -> tuple[str, str]:
    """
    Fallback si no hay TABLAARCA: parsea texto tipo '1 - Factura A'
    """
    s = str(desc or "").strip()

    # Cpbte
    su = s.upper()
    if "NOTA DE CREDITO" in su or "NOTA DE CRÉDITO" in su:
        t = "NC"
    elif "NOTA DE DEBITO" in su or "NOTA DE DÉBITO" in su:
        t = "ND"
    elif "RECIBO" in su:
        t = "R"
    elif "FACTURA" in su:
        t = "F"
    else:
        t = ""

    # Letra
    letra = s[-1].upper() if s else ""
    if letra not in ("A", "B", "C"):
        letra = ""

    # compatibilidad con tu regla heredada
    if s.startswith("8 ") and s.strip().upper().endswith("C"):
        letra = "B"

    return t, letra

# ---------------- Main ----------------

if uploaded is None:
    st.stop()

df, kind = read_arca_file(uploaded)

tabla_lookup = {}
if tablaarca_file is not None:
    tabla_df = read_tablaarca_any(tablaarca_file)
    tabla_lookup = build_tabla_lookup(tabla_df)

# --- Columnas ARCA (CSV / XLSX) ---
COL_FECHA = pick_col(df, "Fecha de Emisión", "Fecha", "Fecha de Emision")
COL_TIPO_COMP = pick_col(df, "Tipo de Comprobante", "Tipo")
COL_PV = pick_col(df, "Punto de Venta", "Pto. Vta.", "Pto Vta", "Punto Venta")
COL_NRO_DESDE = pick_col(df, "Número Desde", "Numero Desde")
# COL_NRO_HASTA = pick_col(df, "Número Hasta", "Numero Hasta")  # no se usa

COL_TIPO_DOC_REC = pick_col(df, "Tipo Doc. Receptor", "Tipo Doc Receptor")
COL_NRO_DOC_REC = pick_col(df, "Nro. Doc. Receptor", "Nro Doc Receptor", "Nro Doc.", "Nro. Doc.")
COL_NOM_REC = pick_col(df, "Denominación Receptor", "Denominacion Receptor")

# Montos (ignoramos 0%, 2,5% y 5% por pedido)
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

# Si es CSV, exigimos TABLAARCA para decodificar el código del comprobante
if kind == "csv" and not tabla_lookup:
    st.error("El archivo es CSV y no se cargó TABLAARCA. Subí TABLAARCA para decodificar 'Tipo de Comprobante'.")
    st.stop()

registros = []

for _, row in df.iterrows():
    tipo_comp_raw = row.get(COL_TIPO_COMP, "")
    if tipo_comp_raw is None or (isinstance(tipo_comp_raw, str) and not tipo_comp_raw.strip()):
        continue

    # Decodificación de tipo/ letra
    cpbte, letra = "", ""

    if kind == "csv":
        # 'Tipo de Comprobante' es código
        k = str(tipo_comp_raw).strip()
        try:
            k = str(int(float(k)))
        except Exception:
            pass
        if k in tabla_lookup:
            cpbte, letra = tabla_lookup[k]
        else:
            cpbte, letra = "", ""
    else:
        # XLSX: puede venir "1 - Factura A"
        s = str(tipo_comp_raw).strip()
        # Si trae código al inicio y hay tabla, usamos tabla
        m = re.match(r"^\s*(\d+)\s*-", s)
        if m and tabla_lookup:
            k = m.group(1)
            cpbte, letra = tabla_lookup.get(k, ("", ""))
        else:
            cpbte, letra = map_tipo_from_xlsx_text(s)

    es_nc = (cpbte == "NC")

    def sg(valor: float) -> float:
        if valor == 0:
            return 0.0
        return -abs(valor) if es_nc else abs(valor)

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
        "Fecha dd/mm/aaaa": row.get(COL_FECHA),
        "Cpbte": cpbte,                  # F / NC / ND / R
        "Tipo": letra,                   # A / B / C
        "Suc.": row.get(COL_PV),
        "Número": row.get(COL_NRO_DESDE),
        "Razón Social o Denominación Cliente": row.get(COL_NOM_REC),
        "Tipo Doc.": tipo_doc_numeric(row.get(COL_TIPO_DOC_REC)),   # CSV ya trae 80/96
        "CUIT": row.get(COL_NRO_DOC_REC),
        "Domicilio": "",
        "C.P.": "",
        "Pcia": "",
        "Cond Fisc": "RI" if letra == "A" else "MT",
        "Cód. Neto": "",     # manual
        "Cód. NG/EX": "",    # manual
        "Cód. P/R": "",      # manual
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

    # Asignar NG/EX y Otros una sola vez (en la 1ra fila)
    if filas_comp:
        if exng_val != 0 or otros_val != 0:
            filas_comp[0]["Conceptos NG/EX"] = exng_val
            filas_comp[0]["Perc./Ret."] = otros_val
    else:
        # Sin alícuotas: volcamos NG/EX y Otros; si solo hay Total, va a NG/EX
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

# --- Generar Excel para descargar ---
buffer = BytesIO()
with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
    salida.to_excel(writer, sheet_name="Salida", index=False)

    workbook = writer.book
    worksheet = writer.sheets["Salida"]

    money_format = workbook.add_format({"num_format": "#,##0.00"})
    aliq_format = workbook.add_format({"num_format": "00.000"})

    col_idx = {name: i for i, name in enumerate(salida.columns)}

    # Montos
    for nombre in ["Neto Gravado", "IVA Liquidado", "IVA Débito", "Conceptos NG/EX", "Perc./Ret.", "Total"]:
        j = col_idx[nombre]
        worksheet.set_column(j, j, 16, money_format)

    # Alícuota
    worksheet.set_column(col_idx["Alíc."], col_idx["Alíc."], 8, aliq_format)

    # Anchos básicos
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
