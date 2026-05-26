import streamlit as st
import pandas as pd
import re
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Procesador de Pedidos",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  [data-testid="stSidebar"] { background: #0f172a; }
  [data-testid="stSidebar"] * { color: #cbd5e1 !important; }
  [data-testid="stSidebar"] .stRadio label { color: #f1f5f9 !important; font-weight: 500; }
  .process-badge {
    display: inline-block; padding: 4px 14px; border-radius: 20px;
    font-size: 12px; font-weight: 600; letter-spacing: .04em;
  }
  .badge-a { background: #dbeafe; color: #1e40af; }
  .badge-b { background: #fce7f3; color: #9d174d; }
  .stat-card {
    background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px;
    padding: 14px 18px; text-align: center;
  }
  .stat-num { font-size: 28px; font-weight: 700; }
  .stat-lbl { font-size: 12px; color: #64748b; margin-top: 2px; }
  div[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
# Maps country name -> ordered list of inter tarifa sheets to search (first match wins).
# For IT/DE/FR there are two sheets each (ES-XX local + XX-XX native) — both searched in order.
# For BE/NL/PL/SE each country has its own independent sheet; if the file uses a combined
# sheet (e.g. "BE - NL") it is tried as fallback so both naming conventions work.
COUNTRY_SHEETS = {
    "Francia":      ["ES-FR", "FR-FR"],
    "France":       ["ES-FR", "FR-FR"],
    "Italia":       ["ES-IT", "IT-IT"],
    "Italy":        ["ES-IT", "IT-IT"],
    "Alemania":     ["ES-DE", "DE-DE"],
    "Germany":      ["ES-DE", "DE-DE"],
    "Portugal":     ["PT"],
    "Bélgica":      ["BE", "BE - NL"],
    "Belgium":      ["BE", "BE - NL"],
    "Países Bajos": ["NL", "BE - NL"],
    "Netherlands":  ["NL", "BE - NL"],
    "Polonia":      ["PL", "PL-SE"],
    "Poland":       ["PL", "PL-SE"],
    "Suecia":       ["SE", "PL-SE"],
    "Sweden":       ["SE", "PL-SE"],
}
# Keep old name as alias for any remaining references
COUNTRY_SHEET = {k: v[0] for k, v in COUNTRY_SHEETS.items()}
SPAIN = {"España", "Spain", "ES"}
NAC_ORDER = ["T_MIR", "T_AMZ", "T_C4", "T_MM", "T_PRIV"]

# ── Email: marketplace → responsable ──────────────────────────────────────────
# Clave en minúsculas, debe coincidir con el valor de la columna Marketplace.
MARKETPLACE_EMAILS = {
    "worten (beezup)":  "responsable.worten@empresa.com",
    "aurgi":            "responsable.aurgi@empresa.com",
    "b2x-85-shein":     "responsable.shein@empresa.com",
    "amazon":           "responsable.amazon@empresa.com",
    "miravia":          "responsable.miravia@empresa.com",
}

def get_email_for_marketplace(marketplace: str) -> str:
    return MARKETPLACE_EMAILS.get(str(marketplace).strip().lower(), "")

def send_order_email(destinatario, asunto, df_lineas, remitente, password):
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = asunto
        msg["From"]    = remitente
        msg["To"]      = destinatario

        STATUS_COLOR_MAP = {
            "✅ OK":           "#065f46",
            "🟡 EN MÍNIMO":    "#92400e",
            "🔴 BAJO MÍNIMO":  "#991b1b",
            "❌ NO EN TARIFA": "#831843",
            "⚠️ SIN PRECIO":   "#92400e",
        }
        th_style = "padding:7px 10px;background:#1B2A4A;color:#fff;font-family:Arial,sans-serif;font-size:12px;text-align:left;"
        td_style = "padding:6px 10px;border:1px solid #e2e8f0;font-family:Arial,sans-serif;font-size:13px;"
        cabeceras = ["Pedido","Fecha","Marketplace","País","SKU","Cant","Precio (€)","PVP Mín (€)","Estado","Dif vs Mín (€)"]
        ths = "".join(f"<th style='{th_style}'>{h}</th>" for h in cabeceras)

        filas_html = ""
        for i, (_, row) in enumerate(df_lineas.iterrows()):
            estado = str(row.get("Estado", ""))
            color_est = STATUS_COLOR_MAP.get(estado, "#1e293b")
            bg = "#f8fafc" if i % 2 == 0 else "#ffffff"
            def td(v, bold=False, color=""):
                s = td_style + (f"font-weight:600;" if bold else "") + (f"color:{color};" if color else "")
                return f"<td style='{s}'>{v}</td>"
            filas_html += f"<tr style='background:{bg}'>"
            filas_html += td(row.get("Pedido","")) + td(row.get("Fecha","")) + td(row.get("Marketplace",""))
            filas_html += td(row.get("País","")) + td(row.get("SKU","")) + td(row.get("Cant",""))
            filas_html += td(f"{row.get('Precio Pedido (€)','')} €") + td(f"{row.get('PVP Mín (€)','')} €")
            filas_html += td(estado, bold=True, color=color_est) + td(f"{row.get('Dif vs Mín (€)','')} €")
            filas_html += "</tr>"

        html = f"""<html><body style="font-family:Arial,sans-serif;color:#1e293b;">
        <h2 style="color:#1B2A4A;">📦 Revisión de pedidos — Análisis de Tarifa</h2>
        <p>Se han detectado pedidos que requieren tu atención:</p>
        <table style="border-collapse:collapse;width:100%;">
          <thead><tr>{ths}</tr></thead><tbody>{filas_html}</tbody>
        </table>
        <br><p style="font-size:12px;color:#64748b;">Generado automáticamente por <strong>Procesador de Pedidos</strong>.</p>
        </body></html>"""
        msg.attach(MIMEText(html, "html", "utf-8"))

        # Adjunto Excel
        excel_bytes = build_excel([("Líneas pedido", df_lineas, None)])
        part = MIMEBase("application", "octet-stream")
        part.set_payload(excel_bytes)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", 'attachment; filename="lineas_pedido.xlsx"')
        msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(remitente, password)
            server.sendmail(remitente, destinatario, msg.as_string())

        return True, f"✅ Email enviado correctamente a **{destinatario}**"
    except Exception as e:
        return False, f"❌ Error al enviar: {str(e)}"

# ── Helpers ───────────────────────────────────────────────────────────────────

def parse_price(val):
    if pd.isna(val) or val == "" or val == 0:
        return None
    s = str(val).replace("€", "").replace(",", ".").strip()
    try:
        return float(s)
    except:
        return None

def normalize_sku(sku):
    """Strip leading S/s and leading zeros: S08299->8299, 01338->1338, A01_EU01_xxx stays"""
    s = str(sku).strip()
    if s.startswith(("A0", "A7", "A1")):  # A01_EU01_xxx format — keep as-is
        return s
    s = re.sub(r"^[Ss]0*", "", s)   # strip leading S + zeros
    s = s.lstrip("0") or s           # strip remaining leading zeros
    return s

def split_skus(sku_val):
    """Split space-separated multi-SKU cell into list of individual SKUs"""
    raw = str(sku_val).strip()
    if not raw or raw in ("nan", ""):
        return []
    return [s.strip() for s in raw.split() if s.strip()]

def expand_multi_sku_rows(df):
    """
    For rows where SKU column contains multiple space-separated SKUs,
    create one row per SKU. Price and Cant are kept as-is per line
    (informative — full order price shown on each line).
    Returns expanded DataFrame with new columns _sku_norm and _multi_flag.
    """
    expanded = []
    for _, row in df.iterrows():
        skus = split_skus(row["Sku"])
        n = len(skus)
        for sku in skus:
            new_row = row.copy()
            new_row["_sku_orig"] = row["Sku"]
            new_row["Sku"] = sku
            new_row["_sku_norm"] = normalize_sku(sku)
            new_row["_multi_flag"] = n > 1
            new_row["_total_skus"] = n
            expanded.append(new_row)
    return pd.DataFrame(expanded)

@st.cache_data(show_spinner=False)
def load_tarifa(file_bytes, filename):
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)

def build_sheet_lookup(df, sheet_name):
    """Build {sku_str: {min, pub, sheet}} lookup for a SINGLE tarifa sheet"""
    lookup = {}
    ref_cols = [c for c in df.columns if str(c).strip() == "REFERENCIA"]
    if not ref_cols:
        return lookup
    ref_col = ref_cols[0]
    for _, r in df.iterrows():
        key = str(r[ref_col]).strip()
        if key:
            try:
                lookup[key] = {
                    "min": float(r.get("PVP MIN.", 0)),
                    "pub": float(r.get("PVP PUB.", 0)),
                    "sheet": sheet_name,
                }
            except:
                pass
    return lookup

def build_lookup(sheets, order):
    """Build merged {sku_str: {min, pub, sheet}} from multiple sheets (national use)"""
    lookup = {}
    for sh in order:
        df = sheets.get(sh)
        if df is None:
            continue
        for key, val in build_sheet_lookup(df, sh).items():
            if key not in lookup:   # first sheet in order wins
                lookup[key] = val
    return lookup

def get_pvp(sku_norm, pais, nac_lookup, inter_lookups):
    """
    Find PVP for a normalized SKU.
    - Spain: national tarifa only (T_MIR → T_AMZ → T_C4 → ...).
    - Other countries: search each sheet in COUNTRY_SHEETS list in order.
      Italia → [ES-IT, IT-IT], Francia → [ES-FR, FR-FR], Alemania → [ES-DE, DE-DE].
      If not found in any of the country's sheets → ❌ NO EN TARIFA.
      No cross-country fallback — a Portuguese price is never used for an Italian order.
    """
    if pais in SPAIN:
        return nac_lookup.get(sku_norm)

    sheets_for_country = COUNTRY_SHEETS.get(pais)
    if sheets_for_country:
        for sheet_key in sheets_for_country:
            lkp = inter_lookups.get(sheet_key, {})
            if sku_norm in lkp:
                return lkp[sku_norm]
        # SKU not in any sheet for this country
        return None

    # Country not mapped → try national as best guess
    return nac_lookup.get(sku_norm)

def analyze_row(row, nac_lookup, inter_lookups):
    sku_norm = row["_sku_norm"]
    pais = str(row.get("País", "")).strip()
    price = parse_price(row.get("Pedido.1"))

    pvp = get_pvp(sku_norm, pais, nac_lookup, inter_lookups)

    if pvp is None:
        return "❌ NO EN TARIFA", None, None, None, None, "CANCELAR"
    
    pvp_min = pvp["min"]
    pvp_pub = pvp["pub"]
    tarifa_sheet = pvp["sheet"]
    tipo = "Nacional" if pais in SPAIN else f"Internacional"

    if price is None:
        return "⚠️ SIN PRECIO", pvp_min, pvp_pub, None, None, tarifa_sheet
    
    diff_min = round(price - pvp_min, 2)
    diff_pub = round(price - pvp_pub, 2)

    if price < pvp_min:
        status = "🔴 BAJO MÍNIMO"
    elif price == pvp_min:
        status = "🟡 EN MÍNIMO"
    else:
        status = "✅ OK"

    return status, pvp_min, pvp_pub, diff_min, diff_pub, tarifa_sheet

def clean_hoja1(df_raw):
    """Delete col A (unnamed:0) and filler row (row with '--')"""
    df = df_raw.copy()
    # Drop first column if it looks like index/bool
    first_col = df.columns[0]
    if "unnamed" in str(first_col).lower() or str(df[first_col].iloc[0]).lower() in ("--", "false", "true"):
        df = df.drop(columns=[first_col])
    # Drop filler rows (cells containing only '-' or '--')
    mask = df.apply(lambda row: all(str(v).strip() in ("--", "-", "nan", "") for v in row), axis=1)
    df = df[~mask].reset_index(drop=True)
    return df

def check_duplicates_hoja1(df):
    """Check duplicates on col C (mail) + col O (Marketplace Order ID), exclude '--'"""
    cols = list(df.columns)
    if len(cols) < 15:
        return pd.DataFrame()
    col_c = cols[2]
    col_o = cols[14]
    valid = df[(df[col_o].astype(str) != "--") & df[col_o].notna() & (df[col_o].astype(str) != "nan")]
    key_count = valid.groupby([col_c, col_o]).size()
    dupes_keys = key_count[key_count > 1].index
    if len(dupes_keys) == 0:
        return pd.DataFrame()
    mask = valid.apply(lambda r: (r[col_c], r[col_o]) in dupes_keys, axis=1)
    return valid[mask].copy()

def check_duplicates_miravia(df_h1):
    """Check duplicates on Combination col (idx 13 after clean), extracted 13 digits"""
    cols = list(df_h1.columns)
    if len(cols) < 14:
        return pd.DataFrame(), {}
    col_comb = cols[13]
    extracted = df_h1[col_comb].fillna("").astype(str).str.strip().apply(
        lambda x: x[-13:] if len(x) >= 13 and x not in ["nan", "--", ""] else ""
    )
    df_h1 = df_h1.copy()
    df_h1["_id_extracted"] = extracted
    valid = df_h1[df_h1["_id_extracted"] != ""]
    id_counts = valid["_id_extracted"].value_counts()
    dup_ids = id_counts[id_counts > 1].index
    dupes = valid[valid["_id_extracted"].isin(dup_ids)].copy()
    return dupes, dict(zip(df_h1.index, extracted))

def cross_miravia_cancelados(df_h1, df_cancelados):
    """Cross Miravia Combination IDs with ES cancelados file col B"""
    cols_h1 = list(df_h1.columns)
    col_comb = cols_h1[13] if len(cols_h1) > 13 else None
    if col_comb is None or df_cancelados is None:
        return pd.DataFrame()
    
    extracted = df_h1[col_comb].fillna("").astype(str).str.strip().apply(
        lambda x: x[-13:] if len(x) >= 13 else ""
    )
    
    cols_can = list(df_cancelados.columns)
    col_b = cols_can[1] if len(cols_can) > 1 else cols_can[0]
    canceled_set = set(df_cancelados[col_b].astype(str).str.strip())
    
    matches = []
    for idx, (_, row) in enumerate(df_h1.iterrows()):
        mid = extracted.iloc[idx] if idx < len(extracted) else ""
        if mid and mid in canceled_set:
            es_row = df_cancelados[df_cancelados[col_b].astype(str).str.strip() == mid]
            matches.append({
                "ID Pedido": row.get("ID", ""),
                "Combination": row.get(col_comb, ""),
                "ID Extraído": mid,
                "SKU": row.get("SKU", ""),
                "Cliente": row.get("Cliente", ""),
                "Total": row.get("Total", ""),
                "Estado ES": es_row["Estado"].values[0] if not es_row.empty else "",
                "Motivo": es_row["Motivo de devolución: No se ha entregado al comprador"].values[0] if not es_row.empty else "",
            })
    return pd.DataFrame(matches)

# ── Excel export ──────────────────────────────────────────────────────────────
def build_excel(sections):
    """sections = list of (sheet_name, df, summary_text)"""
    wb = Workbook()
    HDR_FILL = PatternFill("solid", start_color="1B2A4A")
    HDR_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    HDR_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
    STATUS_COLOR = {
        "✅ OK": "27AE60", "🟡 EN MÍNIMO": "F39C12", "🔴 BAJO MÍNIMO": "E74C3C",
        "❌ NO EN TARIFA": "C0392B", "⚠️ SIN PRECIO": "F39C12", "CANCELAR": "C0392B",
    }

    first = True
    for sheet_name, df, summary in sections:
        ws = wb.active if first else wb.create_sheet()
        ws.title = sheet_name[:31]
        first = False

        start = 1
        if summary:
            ws.cell(1, 1, summary).font = Font(name="Arial", bold=True, size=11,
                color="27AE60" if "Sin " in summary or "✅" in summary else "C0392B")
            start = 3

        if df is None or len(df) == 0:
            continue

        for ci, col in enumerate(df.columns, 1):
            c = ws.cell(start, ci, str(col))
            c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = HDR_ALIGN

        for ri, (_, row) in enumerate(df.iterrows(), 1):
            for ci, val in enumerate(row.values, 1):
                ws.cell(start + ri, ci, val).font = Font(name="Arial", size=10)
            # Color Estado column if present
            if "Estado" in df.columns:
                est_ci = list(df.columns).index("Estado") + 1
                status_val = str(row.get("Estado", ""))
                for k, color in STATUS_COLOR.items():
                    if k in status_val:
                        c = ws.cell(start + ri, est_ci)
                        c.fill = PatternFill("solid", start_color=color)
                        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
                        break

        for col in ws.columns:
            ws.column_dimensions[get_column_letter(col[0].column)].width = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 📦 Procesador Pedidos")
    st.markdown("---")

    proceso = st.radio(
        "**Proceso**",
        ["🛒  Pago aceptado", "🏪  Pago aceptado Miravia"],
        key="proceso_radio",
    )
    es_miravia = "Miravia" in proceso

    st.markdown("---")
    st.markdown("### Tarifas (obligatorio)")
    nac_file = st.file_uploader("Tarifa Nacional (.xlsx)", type="xlsx", key="nac")
    inter_file = st.file_uploader("Tarifa Internacional (.xlsx)", type="xlsx", key="inter")

    st.markdown("---")
    if es_miravia:
        st.markdown("### Ficheros Miravia")
        miravia_file = st.file_uploader("PagoAceptadoMiravia.xlsx", type="xlsx", key="miravia")
        cancelados_file = st.file_uploader("Fichero ES... cancelados (.xlsx)", type="xlsx", key="cancelados")
        libro_file = None
    else:
        st.markdown("### Fichero de pedidos")
        libro_file = st.file_uploader(
            "Rentabilidad (.xlsx)\n*(Turaco / Jabiru — Hoja1 + Hoja2)*",
            type="xlsx", key="libro"
        )
        miravia_file = cancelados_file = None

    st.markdown("---")
    run_btn = st.button("▶ Procesar", type="primary", use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
st.title("📦 Procesador de Pedidos")

if not run_btn and "results" not in st.session_state:
    st.info("👈 Selecciona el proceso y sube los ficheros en el panel lateral. Luego pulsa **Procesar**.")
    st.markdown("""
    #### Proceso A — Pago aceptado
    - **Hoja1**: limpieza (borra col A + fila --), chequeo de duplicados  
    - **Hoja2 / Rentabilidad**: análisis tarifa con expansión de multi-SKU  
    - SKUs con varios productos separados por espacio → **una línea por SKU**  
    - Normalización automática de SKU: `S08299` → `8299`, `01338` → `1338`

    #### Proceso B — Pago aceptado Miravia
    - Todo lo anterior  
    - **Cruce cancelados**: col O Combination (`Arise - XXXXXXXXXXXXX`) → extrae 13 dígitos → cruza con col B del fichero ES  
    - **Duplicados Combination** en la columna extraída
    """)
    st.stop()

# ── Si se pulsa Procesar: validar, calcular y guardar en session_state ─────────
if run_btn:
    errors = []
    if not nac_file:   errors.append("Tarifa Nacional")
    if not inter_file: errors.append("Tarifa Internacional")
    if not es_miravia and not libro_file:   errors.append("Fichero Rentabilidad")
    if es_miravia     and not miravia_file: errors.append("PagoAceptadoMiravia.xlsx")
    if errors:
        st.error(f"Faltan ficheros: {', '.join(errors)}")
        st.stop()

    # Tarifas
    with st.spinner("Cargando tarifas..."):
        nac_sheets   = load_tarifa(nac_file.read(),   nac_file.name)
        inter_sheets = load_tarifa(inter_file.read(), inter_file.name)
        nac_lookup   = build_lookup(nac_sheets, NAC_ORDER)
        inter_lookups = {sh: build_sheet_lookup(df, sh) for sh, df in inter_sheets.items()}

    if not es_miravia:
        with st.spinner("Procesando..."):
            libro_bytes  = libro_file.read()
            libro_sheets = pd.read_excel(io.BytesIO(libro_bytes), sheet_name=None)
            hoja1_raw    = libro_sheets.get("Hoja1", pd.DataFrame())
            hoja2_raw    = libro_sheets.get("Hoja2", pd.DataFrame())
            sheet_names  = list(libro_sheets.keys())
            if "Hoja2" not in libro_sheets and len(sheet_names) > 1:
                hoja2_raw = libro_sheets[sheet_names[1]]

            h1_clean  = clean_hoja1(hoja1_raw) if not hoja1_raw.empty else pd.DataFrame()
            dupes_h1  = check_duplicates_hoja1(h1_clean) if not h1_clean.empty else pd.DataFrame()

            df_tarifa   = pd.DataFrame()
            multi_count = 0
            if not hoja2_raw.empty:
                df_expanded = expand_multi_sku_rows(hoja2_raw)
                results = []
                for _, row in df_expanded.iterrows():
                    status, pvp_min, pvp_pub, diff_min, diff_pub, tarifa_sheet = analyze_row(row, nac_lookup, inter_lookups)
                    results.append({
                        "Pedido": row.get("Pedido",""), "Fecha": row.get("Fecha",""),
                        "Marketplace": row.get("Marketplace",""), "Id Marketplace": row.get("Id Marketplace",""),
                        "País": row.get("País",""), "SKU Original": row.get("_sku_orig", row.get("Sku","")),
                        "SKU": row.get("Sku",""), "SKU Norm.": row.get("_sku_norm",""),
                        "Multi-SKU": "✔" if row.get("_multi_flag") else "",
                        "Cant": row.get("Cant",""), "Precio Pedido (€)": parse_price(row.get("Pedido.1")),
                        "Hoja Tarifa": tarifa_sheet, "PVP Mín (€)": pvp_min, "PVP Pub (€)": pvp_pub,
                        "Dif vs Mín (€)": diff_min, "Dif vs Pub (€)": diff_pub, "Estado": status,
                    })
                df_tarifa   = pd.DataFrame(results)
                multi_count = int(df_expanded["_multi_flag"].sum())

        st.session_state["results"] = {
            "tipo": "A", "nac_count": len(nac_lookup),
            "inter_count": sum(len(v) for v in inter_lookups.values()),
            "h1_clean": h1_clean, "dupes_h1": dupes_h1,
            "df_tarifa": df_tarifa, "multi_count": multi_count,
        }

    else:  # Miravia
        with st.spinner("Procesando Miravia..."):
            mir_bytes    = miravia_file.read()
            mir_sheets   = pd.read_excel(io.BytesIO(mir_bytes), sheet_name=None)
            hoja1_raw    = mir_sheets.get("Hoja1", pd.DataFrame())
            hoja2_raw    = mir_sheets.get("Hoja2", pd.DataFrame())

            df_cancelados = None
            if cancelados_file:
                can_bytes  = cancelados_file.read()
                can_sheets = pd.read_excel(io.BytesIO(can_bytes), sheet_name=None)
                df_cancelados = list(can_sheets.values())[0]

            h1_clean       = clean_hoja1(hoja1_raw) if not hoja1_raw.empty else pd.DataFrame()
            dupes_comb, id_map = check_duplicates_miravia(h1_clean) if not h1_clean.empty else (pd.DataFrame(), {})
            df_cancel_match = cross_miravia_cancelados(h1_clean, df_cancelados) if df_cancelados is not None else pd.DataFrame()

            df_tarifa   = pd.DataFrame()
            multi_count = 0
            if not hoja2_raw.empty:
                df_expanded = expand_multi_sku_rows(hoja2_raw)
                results = []
                for _, row in df_expanded.iterrows():
                    status, pvp_min, pvp_pub, diff_min, diff_pub, tarifa_sheet = analyze_row(row, nac_lookup, inter_lookups)
                    results.append({
                        "Pedido": row.get("Pedido",""), "SKU Original": row.get("_sku_orig",""),
                        "SKU": row.get("Sku",""), "SKU Norm.": row.get("_sku_norm",""),
                        "Multi-SKU": "✔" if row.get("_multi_flag") else "",
                        "País": row.get("País",""), "Precio Pedido (€)": parse_price(row.get("Pedido.1")),
                        "Hoja Tarifa": tarifa_sheet, "PVP Mín (€)": pvp_min, "PVP Pub (€)": pvp_pub,
                        "Dif vs Mín (€)": diff_min, "Estado": status,
                    })
                df_tarifa   = pd.DataFrame(results)
                multi_count = int(df_expanded["_multi_flag"].sum())

        st.session_state["results"] = {
            "tipo": "B", "nac_count": len(nac_lookup),
            "inter_count": sum(len(v) for v in inter_lookups.values()),
            "h1_clean": h1_clean, "dupes_comb": dupes_comb, "id_map": id_map,
            "df_cancel_match": df_cancel_match, "df_cancelados": df_cancelados,
            "cancelados_provided": cancelados_file is not None,
            "df_tarifa": df_tarifa, "multi_count": multi_count,
        }

# ── Leer resultados de session_state y mostrar ────────────────────────────────
if "results" not in st.session_state:
    st.stop()

R = st.session_state["results"]

st.success(f"✅ Tarifas cargadas — Nacional: {R['nac_count']:,} SKUs | Internacional: {R['inter_count']:,} refs")

# ── Widget email reutilizable ─────────────────────────────────────────────────
def email_widget(df_tarifa, key_prefix):
    """Muestra selector de líneas + formulario de envío de email."""
    st.markdown("---")
    st.markdown("### 📧 Enviar líneas de pedido por email")

    opciones = df_tarifa.apply(
        lambda r: f"Pedido {r.get('Pedido','')} · {r.get('Marketplace', r.get('País',''))} · SKU {r.get('SKU','')} · {r.get('Estado','')}",
        axis=1
    ).tolist()

    seleccionadas = st.multiselect(
        "Selecciona las líneas a enviar:",
        options=list(range(len(df_tarifa))),
        format_func=lambda i: opciones[i],
        key=f"{key_prefix}_rows",
    )

    if not seleccionadas:
        st.info("Selecciona una o varias líneas de la tabla para enviar por email.")
        return

    df_sel = df_tarifa.iloc[seleccionadas].reset_index(drop=True)

    # Autocompletar email si todas las líneas son del mismo marketplace
    mps = df_sel["Marketplace"].dropna().unique().tolist() if "Marketplace" in df_sel.columns else []
    email_auto = get_email_for_marketplace(mps[0]) if len(mps) == 1 else ""
    asunto_auto = f"Revisión pedidos — {', '.join(mps)}" if mps else "Revisión pedidos"

    col1, col2 = st.columns([2, 1])
    with col1:
        destinatario = st.text_input(
            "📬 Email destinatario",
            value=email_auto,
            placeholder="responsable@empresa.com",
            help="Se autocompleta según el marketplace. Editable libremente.",
            key=f"{key_prefix}_dest",
        )
    with col2:
        asunto = st.text_input("✏️ Asunto", value=asunto_auto, key=f"{key_prefix}_asunto")

    st.caption(f"Se enviarán **{len(df_sel)}** línea(s) con adjunto Excel.")

    remitente = st.secrets.get("EMAIL_REMITENTE", "")
    password  = st.secrets.get("EMAIL_PASSWORD",  "")

    if not remitente or not password:
        st.warning("⚠️ Configura `EMAIL_REMITENTE` y `EMAIL_PASSWORD` en los Secrets de Streamlit.")
    elif not destinatario:
        st.warning("⚠️ Introduce un email destinatario.")
    else:
        if st.button("📤 Enviar email", type="primary", key=f"{key_prefix}_send"):
            ok, msg_result = send_order_email(destinatario, asunto, df_sel, remitente, password)
            if ok:
                st.success(msg_result)
            else:
                st.error(msg_result)

# ─────────────────────────────────────────────────────────────────────────────
def color_status(val):
    colors = {
        "✅ OK":           "background-color:#d1fae5;color:#065f46",
        "🟡 EN MÍNIMO":    "background-color:#fef3c7;color:#92400e",
        "🔴 BAJO MÍNIMO":  "background-color:#fee2e2;color:#991b1b",
        "❌ NO EN TARIFA": "background-color:#fce7f3;color:#831843",
        "⚠️ SIN PRECIO":   "background-color:#fef3c7;color:#92400e",
    }
    return colors.get(val, "")

# ═══════════════════════════════════════════════════════════════════════════════
# DISPLAY A
# ═══════════════════════════════════════════════════════════════════════════════
if R["tipo"] == "A":
    df_tarifa   = R["df_tarifa"]
    h1_clean    = R["h1_clean"]
    dupes_h1    = R["dupes_h1"]
    multi_count = R["multi_count"]

    st.markdown("---")
    st.markdown("## 🛒 Pago aceptado — Resultados")

    ok     = (df_tarifa["Estado"] == "✅ OK").sum() if not df_tarifa.empty else 0
    warn   = (df_tarifa["Estado"].str.startswith("🔴").sum() + df_tarifa["Estado"].str.startswith("🟡").sum()) if not df_tarifa.empty else 0
    cancel = df_tarifa["Estado"].str.startswith("❌").sum() if not df_tarifa.empty else 0

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("📋 Hoja1 filas",    len(h1_clean))
    col2.metric("⚠️ Duplicados",     len(dupes_h1))
    col3.metric("✅ OK tarifa",       ok)
    col4.metric("🔴 Bajo mínimo",    warn)
    col5.metric("❌ No en tarifa",   cancel)

    tab1, tab2, tab3 = st.tabs(["💰 Análisis Tarifa", "🔍 Duplicados Hoja1", "📄 Hoja1 limpia"])

    with tab1:
        if not df_tarifa.empty:
            if multi_count > 0:
                st.info(f"🔀 Se han expandido pedidos multi-SKU: **{multi_count}** líneas generadas por separación de SKUs en la misma celda")
            st.dataframe(df_tarifa.style.map(color_status, subset=["Estado"]), use_container_width=True, hide_index=True)
            email_widget(df_tarifa, "procA")
        else:
            st.info("No hay datos en Hoja2 para analizar.")

    with tab2:
        if len(dupes_h1) > 0:
            st.warning(f"⚠️ Se encontraron **{len(dupes_h1)}** filas duplicadas (col C mail + col O Marketplace Order ID)")
            st.dataframe(dupes_h1, use_container_width=True, hide_index=True)
        else:
            st.success("✅ Sin duplicados en Hoja1")

    with tab3:
        if not h1_clean.empty:
            st.dataframe(h1_clean, use_container_width=True, hide_index=True)
        else:
            st.info("No hay datos en Hoja1.")

    st.markdown("---")
    sections = [
        ("Análisis Tarifa", df_tarifa, None),
        ("Duplicados Hoja1", dupes_h1 if len(dupes_h1) > 0 else None,
         "Sin duplicados encontrados" if len(dupes_h1) == 0 else f"⚠️ {len(dupes_h1)} duplicados detectados"),
        ("Hoja1 limpia", h1_clean, None),
    ]
    excel_bytes = build_excel([(s, d, t) for s, d, t in sections if d is not None or t is not None])
    st.download_button("⬇️ Descargar Excel completo", data=excel_bytes,
        file_name=f"PagoAceptado_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

# ═══════════════════════════════════════════════════════════════════════════════
# DISPLAY B — Miravia
# ═══════════════════════════════════════════════════════════════════════════════
else:
    df_tarifa        = R["df_tarifa"]
    h1_clean         = R["h1_clean"]
    dupes_comb       = R["dupes_comb"]
    id_map           = R["id_map"]
    df_cancel_match  = R["df_cancel_match"]
    df_cancelados    = R["df_cancelados"]
    multi_count      = R["multi_count"]
    cancelados_prov  = R["cancelados_provided"]

    st.markdown("---")
    st.markdown("## 🏪 Pago aceptado Miravia — Resultados")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📋 Pedidos Miravia",    len(h1_clean))
    col2.metric("🔴 Cancelados match",   len(df_cancel_match))
    col3.metric("⚠️ Dupl. Combination",  len(dupes_comb))
    col4.metric("💰 Tarifa analizados",  len(df_tarifa))

    tabs = st.tabs(["🔴 Cancelados", "⚠️ Duplicados Combination", "💰 Análisis Tarifa", "📄 Hoja1 limpia"])

    with tabs[0]:
        if not cancelados_prov:
            st.info("No se ha subido el fichero ES de cancelados.")
        elif len(df_cancel_match) > 0:
            st.error(f"🔴 **{len(df_cancel_match)}** pedidos Miravia encontrados en el fichero de cancelados")
            st.dataframe(df_cancel_match, use_container_width=True, hide_index=True)
        else:
            n_ids = sum(1 for v in id_map.values() if v)
            n_can = len(df_cancelados) if df_cancelados is not None else 0
            st.success(f"✅ Ningún pedido Miravia en cancelados ({n_ids} IDs cruzados contra {n_can:,} cancelados)")

    with tabs[1]:
        if len(dupes_comb) > 0:
            st.warning(f"⚠️ **{len(dupes_comb)}** filas con ID Combination duplicado")
            st.dataframe(dupes_comb, use_container_width=True, hide_index=True)
        else:
            st.success("✅ Sin IDs duplicados en columna Combination")

    with tabs[2]:
        if not df_tarifa.empty:
            if multi_count > 0:
                st.info(f"🔀 **{multi_count}** líneas generadas por expansión de pedidos multi-SKU")
            st.dataframe(df_tarifa.style.map(color_status, subset=["Estado"]), use_container_width=True, hide_index=True)
            email_widget(df_tarifa, "procB")
        else:
            st.info("Sin datos en Hoja2 para análisis de tarifa (habitual en Miravia).")

    with tabs[3]:
        st.dataframe(h1_clean, use_container_width=True, hide_index=True)

    st.markdown("---")
    sections = [
        ("Cancelados Match",      df_cancel_match if len(df_cancel_match) > 0 else None,
         "Sin cancelados encontrados" if len(df_cancel_match) == 0 else None),
        ("Duplicados Combination", dupes_comb if len(dupes_comb) > 0 else None,
         "Sin duplicados en Combination" if len(dupes_comb) == 0 else None),
        ("Análisis Tarifa",       df_tarifa if not df_tarifa.empty else None, None),
        ("Hoja1 limpia",          h1_clean, None),
    ]
    excel_bytes = build_excel([(s, d, t) for s, d, t in sections if d is not None or t is not None])
    st.download_button("⬇️ Descargar Excel completo", data=excel_bytes,
        file_name=f"PagoAceptadoMiravia_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
