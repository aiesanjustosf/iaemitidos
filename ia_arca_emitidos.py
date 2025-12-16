# ia_arca_emitidos.py
# ARCA Emitidos (XLSX o CSV) + Ventas Pastor Chess (XLSX) -> Formato Holistor (HWVta1modelo)
# AIE San Justo

import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
import csv
import re
from datetime import date, datetime

# ---------------- Matriz interna (ARCA CSV) ----------------
TIPOS_COMP = {
    "1": ("F", "A"),
    "2": ("ND", "A"),
    "3": ("NC", "A"),
    "4": ("R", "A"),
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
CREDITOS_ARCA = {"NC", "PC"}  # crédito => negativo

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
    page_title="Emitidos → Formato Holistor",
    page_icon=str(FAVICON_PATH) if FAVICON_PATH else None,
    layout="centered",
)

if LOGO_PATH:
    st.image(str(LOGO_PATH), width=180)

st.title("Emitidos → Formato Holistor")

fuente = st.radio(
    "Fuente de datos",
    ["ARCA Emitidos (XLSX/CSV)", "Ventas Pastor Chess (XLSX)"],
    horizontal=True,
)

# ---------------- Helpers comunes ----------------
def sniff_delimiter(text: str) -> str:
    try:
        d = csv.Sniffer().sniff(text[:5000], delimiters=";,|\t")
        return d.delimiter
    except Exception:
        return ";"


def pick_col(df: pd.DataFrame, *cands: str) -> str:
    cols = list(df.columns)
    colset = set(cols)
    for c in cands:
        if c in colset:
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


def fecha_out(v) -> str:
    """
    Devuelve SIEMPRE texto DD/MM/AAAA.
    - Si viene DD/MM/AAAA: lo deja igual.
    - Si viene ISO YYYY-MM-DD o YYYY-MM-DD HH:MM:SS: lo convierte a DD/MM/AAAA.
    - Si viene date/datetime: lo formatea DD/MM/AAAA.
    - Si viene vacío: "".
    """
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""

    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")

    s = str(v).strip()
    if not s:
        return ""

    if re.fullmatch(r"\d{2}[/-]\d{2}[/-]\d{4}", s):
        return s.replace("-", "/")

    if re.match(r"^\d{4}-\d{2}-\d{2}(\s+\d{2}:\d{2}:\d{2})?$", s):
        dt = pd.to_datetime(s, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%d/%m/%Y")

    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if not pd.isna(dt):
        return dt.strftime("%d/%m/%Y")

    return s


# ---------------- ARCA: helpers ----------------
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


def map_tipo_from_text(desc: str) -> tuple[str, str]:
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

    # compat recibidos (si apareciera el caso)
    if s.startswith("8 ") and s.strip().upper().endswith("C"):
        letra = "B"

    return t, letra


def decode_csv_tipo(tipo_comp_raw: str) -> tuple[str, str]:
    k = str(tipo_comp_raw).strip()
    try:
        k = str(int(float(k)))
    except Exception:
        pass
    return TIPOS_COMP.get(k, ("", ""))


def process_arca(uploaded) -> tuple[pd.DataFrame, list[str]]:
    df, kind = read_arca(uploaded)
    warnings: list[str] = []

    COL_FECHA = pick_col(df, "Fecha de Emisión", "Fecha", "Fecha de Emision")
    COL_TIPO_COMP = pick_col(df, "Tipo de Comprobante", "Tipo")
    COL_PV = pick_col(df, "Punto de Venta", "Pto. Vta.", "Pto Vta", "Punto Venta")
    COL_NRO_DESDE = pick_col(df, "Número Desde", "Numero Desde")

    COL_TIPO_DOC_REC = pick_col(df, "Tipo Doc. Receptor", "Tipo Doc Receptor")
    COL_NRO_DOC_REC = pick_col(df, "Nro. Doc. Receptor", "Nro Doc Receptor", "Nro Doc.", "Nro. Doc.")
    COL_NOM_REC = pick_col(df, "Denominación Receptor", "Denominacion Receptor")

    # Solo alícuotas consideradas (ignoramos 0%, 2.5%, 5%)
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

        if kind == "csv":
            cpbte, letra = decode_csv_tipo(tipo_comp_raw)
        else:
            cpbte, letra = map_tipo_from_text(tipo_comp_raw)

        es_credito = (cpbte in CREDITOS_ARCA)

        def sg(x: float) -> float:
            if x == 0:
                return 0.0
            return -abs(x) if es_credito else abs(x)

        tdoc = tipo_doc(row.get(COL_TIPO_DOC_REC))
        nro_doc = digits_only(row.get(COL_NRO_DOC_REC))

        # Condición fiscal (sin MT)
        cuit_out = nro_doc
        cond_fisc = ""

        # Reglas:
        # A 80 -> RI
        # B 80 -> EX
        # B 96 -> CF + CUIT armado "00-XXXXXXXX-0"
        if letra == "A" and tdoc == 80:
            cond_fisc = "RI"
        elif letra == "B" and tdoc == 80:
            cond_fisc = "EX"
        elif letra == "B" and tdoc == 96 and nro_doc:
            dni8 = nro_doc.zfill(8)
            cuit_out = f"00-{dni8}-0"
            cond_fisc = "CF"

        exng_val = sg(parse_amount(row.get(COL_NETO_NG)) + parse_amount(row.get(COL_EXENTAS)))
        otros_val = sg(parse_amount(row.get(COL_OTROS)))
        total_val = sg(parse_amount(row.get(COL_TOTAL)))

        netos_ivas = [
            sg(parse_amount(row.get(COL_NETO_105))), sg(parse_amount(row.get(COL_IVA_105))),
            sg(parse_amount(row.get(COL_NETO_21))),  sg(parse_amount(row.get(COL_IVA_21))),
            sg(parse_amount(row.get(COL_NETO_27))),  sg(parse_amount(row.get(COL_IVA_27))),
        ]

        if exng_val == 0 and otros_val == 0 and total_val == 0 and all(v == 0 for v in netos_ivas):
            continue

        base = {
            "Fecha dd/mm/aaaa": fecha_out(row.get(COL_FECHA)),
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
        raise ValueError("No se encontraron comprobantes con importes.")

    cols_salida = [
        "Fecha dd/mm/aaaa", "Cpbte", "Tipo", "Suc.", "Número",
        "Razón Social o Denominación Cliente", "Tipo Doc.", "CUIT",
        "Domicilio", "C.P.", "Pcia", "Cond Fisc",
        "Cód. Neto", "Neto Gravado", "Alíc.",
        "IVA Liquidado", "IVA Débito",
        "Cód. NG/EX", "Conceptos NG/EX",
        "Cód. P/R", "Perc./Ret.", "Pcia P/R",
        "Total",
    ]

    salida = pd.DataFrame(registros)[cols_salida]
    return salida, warnings


# ---------------- Pastor Chess ----------------
PASTOR_SKIP = {
    "ANTICIPO",
    "COMODATO ACTIVOS FIJOS",
    "SOBRANTE DE LIQUIDACION",
    "RECIBO",
}

PASTOR_PERCEP_MAP = [
    ("Percepción 3337", "PV07"),
    ("Percepción 5329", "PV07"),
    ("Percepción 212", "PV06"),
    ("I.I.B.B(SANTA FE)", "PV06"),
    ("I.I.B.B(FORMOSA)", "PV06"),
]


def process_pastor(uploaded) -> tuple[pd.DataFrame, list[str]]:
    warnings: list[str] = []
    df = pd.read_excel(uploaded, sheet_name=0, header=0, dtype=object)

    # columnas base (por nombre real del archivo Pastor Chess)
    COL_FECHA = pick_col(df, "Fecha Comprobante")
    COL_DESC_COMP = pick_col(df, "Descripcion Comprobante", "Descripción Comprobante")
    COL_LETRA = pick_col(df, "Letra")
    COL_SUC = pick_col(df, "Serie \\ Punto de venta", "Serie / Punto de venta", "Serie / Punto de Venta")
    COL_NUM = pick_col(df, "Numero", "Número")
    COL_RS = pick_col(df, "Razon Social", "Razón Social")
    COL_TDOC = pick_col(df, "Tipo Id", "Tipo ID", "Tipo Id.")
    COL_DDOC = pick_col(df, "Descripción Id", "Descripcion Id")
    COL_NDOC = pick_col(df, "Identificador")
    COL_PCIA = pick_col(df, "Provincia")
    COL_COND = pick_col(df, "Tipo IVA", "Condición fiscal", "Condicion fiscal", "Cond Fisc")

    COL_NETO = pick_col(df, "Subtotal Neto")
    COL_IVA = pick_col(df, "I.V.A", "IVA")
    COL_TOTAL = pick_col(df, "Subtotal Final", "Total")

    registros = []

    for i, row in df.iterrows():
        desc = str(row.get(COL_DESC_COMP, "") or "").strip().upper()
        if not desc:
            continue

        # excluir tipos
        if desc in PASTOR_SKIP:
            continue

        # cpbte
        if desc == "FACTURA":
            cpbte = "F"
        elif desc == "NOTA DE CREDITO":
            cpbte = "NC"
        elif desc == "NOTA DE DEBITO":
            cpbte = "ND"
        else:
            # no procesar otros
            continue

        es_credito = (cpbte == "NC")

        def sg(x: float) -> float:
            if x == 0:
                return 0.0
            return -abs(x) if es_credito else abs(x)

        letra = str(row.get(COL_LETRA, "") or "").strip().upper()
        suc = row.get(COL_SUC)
        nro = row.get(COL_NUM)

        rs = row.get(COL_RS)

        tdoc = tipo_doc(row.get(COL_TDOC))
        nro_doc = digits_only(row.get(COL_NDOC))

        # DNI => 00-XXXXXXXX-0 (8 dígitos fijos)
        cuit_out = nro_doc
        if tdoc == 96 and nro_doc:
            dni8 = nro_doc.zfill(8)
            cuit_out = f"00-{dni8}-0"

        # Condición fiscal
        cond = str(row.get(COL_COND, "") or "").strip().upper()
        if cond == "MT":
            cond = "MTD"

        pcia = str(row.get(COL_PCIA, "") or "").strip()

        neto = sg(parse_amount(row.get(COL_NETO)))
        iva = sg(parse_amount(row.get(COL_IVA)))
        total_origen = sg(parse_amount(row.get(COL_TOTAL)))

        # chequeo IVA 21%
        if neto != 0:
            esperado = round(abs(neto) * 0.21, 2)
            if round(abs(iva), 2) not in (esperado, round(esperado + 0.01, 2), round(esperado - 0.01, 2)):
                warnings.append(
                    f"Fila {i+2}: IVA no cuadra con 21% (Neto={neto:,.2f} / IVA={iva:,.2f} / Esperado≈{sg(esperado):,.2f})."
                )

        # percepciones presentes
        percs = []
        for col_name, cod_pr in PASTOR_PERCEP_MAP:
            if col_name in df.columns:
                val = sg(parse_amount(row.get(col_name)))
                if val != 0:
                    percs.append((cod_pr, val, col_name))

        # Si no hay montos (ni neto/iva ni percepciones), omitir
        if neto == 0 and iva == 0 and not percs:
            continue

        base = {
            "Fecha dd/mm/aaaa": fecha_out(row.get(COL_FECHA)),
            "Cpbte": cpbte,
            "Tipo": letra,
            "Suc.": suc,
            "Número": nro,
            "Razón Social o Denominación Cliente": rs,
            "Tipo Doc.": tdoc,
            "CUIT": cuit_out,
            "Domicilio": "",
            "C.P.": "",
            "Pcia": pcia,
            "Cond Fisc": cond,
            "Cód. Neto": "135",
            "Cód. NG/EX": "",
            "Cód. P/R": "",
            "Pcia P/R": "",
        }

        # Armado de líneas:
        # - 0 percepciones: una línea neto/iva
        # - 1 percepción: va en la MISMA línea neto/iva
        # - >1 percepciones: primera en la línea principal; resto en líneas extra sin neto/iva
        lineas = []

        # Línea principal
        main = base.copy()
        main["Neto Gravado"] = neto
        main["Alíc."] = 21.0  # 21,000
        main["IVA Liquidado"] = iva
        main["IVA Débito"] = iva
        main["Conceptos NG/EX"] = 0.0
        main["Perc./Ret."] = 0.0

        if percs:
            cod_pr0, val0, _ = percs[0]
            main["Cód. P/R"] = cod_pr0
            main["Perc./Ret."] = val0

        # Total línea principal (suma de componentes)
        main["Total"] = (
            float(main["Neto Gravado"] or 0)
            + float(main["IVA Liquidado"] or 0)
            + float(main["Conceptos NG/EX"] or 0)
            + float(main["Perc./Ret."] or 0)
        )
        lineas.append(main)

        # Líneas extra para percepciones adicionales
        if len(percs) > 1:
            for cod_pr, val, _colname in percs[1:]:
                extra = base.copy()
                extra["Cód. Neto"] = ""         # no repetir
                extra["Neto Gravado"] = 0.0
                extra["Alíc."] = 0.0
                extra["IVA Liquidado"] = 0.0
                extra["IVA Débito"] = 0.0
                extra["Conceptos NG/EX"] = 0.0
                extra["Cód. P/R"] = cod_pr
                extra["Perc./Ret."] = val
                extra["Total"] = float(val or 0)
                lineas.append(extra)

        # Check total vs origen (solo informativo)
        total_calc = sum(float(x.get("Total", 0) or 0) for x in lineas)
        if total_origen != 0 and round(total_calc, 2) != round(total_origen, 2):
            warnings.append(
                f"Fila {i+2}: Total origen ({total_origen:,.2f}) != Total calculado ({total_calc:,.2f})."
            )

        registros.extend(lineas)

    if not registros:
        raise ValueError("No se encontraron comprobantes con importes (Pastor Chess).")

    cols_salida = [
        "Fecha dd/mm/aaaa", "Cpbte", "Tipo", "Suc.", "Número",
        "Razón Social o Denominación Cliente", "Tipo Doc.", "CUIT",
        "Domicilio", "C.P.", "Pcia", "Cond Fisc",
        "Cód. Neto", "Neto Gravado", "Alíc.",
        "IVA Liquidado", "IVA Débito",
        "Cód. NG/EX", "Conceptos NG/EX",
        "Cód. P/R", "Perc./Ret.", "Pcia P/R",
        "Total",
    ]

    salida = pd.DataFrame(registros)[cols_salida]
    return salida, warnings


# ---------------- Ejecutar según fuente ----------------
if fuente.startswith("ARCA"):
    uploaded = st.file_uploader("Subí ARCA Emitidos (.xlsx o .csv)", type=["xlsx", "csv"], key="arca_upl")
    if uploaded is None:
        st.stop()
    try:
        salida, warns = process_arca(uploaded)
        nombre_salida = "Emitidos_salida.xlsx"
    except Exception as e:
        st.error(str(e))
        st.stop()
else:
    uploaded = st.file_uploader("Subí Ventas Pastor Chess (.xlsx)", type=["xlsx"], key="pastor_upl")
    if uploaded is None:
        st.stop()
    try:
        salida, warns = process_pastor(uploaded)
        nombre_salida = "PastorChess_salida.xlsx"
    except Exception as e:
        st.error(str(e))
        st.stop()

# ---------------- Preview ----------------
st.subheader("Vista previa de la salida")
st.dataframe(salida.head(50))

if warns:
    st.warning("Se detectaron advertencias (no bloquean la salida).")
    st.write("\n".join(warns[:50]))
    if len(warns) > 50:
        st.write(f"... y {len(warns) - 50} más.")

# ---------------- Export ----------------
buffer = BytesIO()
with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
    salida.to_excel(writer, sheet_name="Salida", index=False)

    wb = writer.book
    ws = writer.sheets["Salida"]

    money_fmt = wb.add_format({"num_format": "#,##0.00"})
    aliq_fmt = wb.add_format({"num_format": "00.000"})
    text_fmt = wb.add_format({"num_format": "@"})  # texto

    col_idx = {c: i for i, c in enumerate(salida.columns)}

    # FECHA como texto para que Holistor no invierta día/mes
    ws.set_column(col_idx["Fecha dd/mm/aaaa"], col_idx["Fecha dd/mm/aaaa"], 12, text_fmt)

    # anchos
    ws.set_column(col_idx["Cpbte"], col_idx["Cpbte"], 6)
    ws.set_column(col_idx["Tipo"], col_idx["Tipo"], 6)
    ws.set_column(col_idx["Suc."], col_idx["Suc."], 10)
    ws.set_column(col_idx["Número"], col_idx["Número"], 12)
    ws.set_column(col_idx["Razón Social o Denominación Cliente"], col_idx["Razón Social o Denominación Cliente"], 42)
    ws.set_column(col_idx["CUIT"], col_idx["CUIT"], 16)

    # importes
    for nombre in ["Neto Gravado", "IVA Liquidado", "IVA Débito", "Conceptos NG/EX", "Perc./Ret.", "Total"]:
        if nombre in col_idx:
            ws.set_column(col_idx[nombre], col_idx[nombre], 16, money_fmt)

    ws.set_column(col_idx["Alíc."], col_idx["Alíc."], 8, aliq_fmt)

buffer.seek(0)

st.download_button(
    "📥 Descargar Excel procesado",
    data=buffer,
    file_name=nombre_salida,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.markdown(
    "<br><hr style='opacity:0.3'><div style='text-align:center; font-size:12px; color:#6b7280;'>"
    "© AIE – Herramienta para uso interno | Developer Alfonso Alderete"
    "</div>",
    unsafe_allow_html=True,
)
