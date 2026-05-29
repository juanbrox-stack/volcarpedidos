import streamlit as st
import pandas as pd
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ── Page config ───────────────────────────────────────────────────────────────
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

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

# Maps marketplace name (substring match) -> national tarifa sheet to search FIRST
# The rest of NAC_ORDER is searched as fallback if SKU not found in primary sheet
MARKETPLACE_NAC_SHEET = {
    "carrefour":  "T_C4",
    "amazon":     "T_AMZ",
    "mediamarkt": "T_MM",
    "privalia":   "T_PRIV",
    # Mirakl/BeezUP/Showroomprive/default -> T_MIR (first in NAC_ORDER)
}

def nac_order_for_marketplace(marketplace: str) -> list:
    """Return NAC_ORDER reordered so the channel-specific sheet is searched first."""
    mkt_lower = str(marketplace).lower()
    primary = None
    for keyword, sheet in MARKETPLACE_NAC_SHEET.items():
        if keyword in mkt_lower:
            primary = sheet
            break
    if primary is None:
        return NAC_ORDER  # default: T_MIR first
    return [primary] + [s for s in NAC_ORDER if s != primary]
SPAIN = {"España", "Spain", "ES"}
NAC_ORDER = ["T_MIR", "T_AMZ", "T_C4", "T_MM", "T_PRIV"]

# ── Email helpers ──────────────────────────────────────────────────────────────

def load_remitentes(file_bytes):
    """Load remitentes Excel: expects columns Canal, Email (and optionally Nombre)"""
    df = pd.read_excel(io.BytesIO(file_bytes))
    # Normalise column names
    df.columns = [c.strip() for c in df.columns]
    return df

def get_remitente(remitentes_df, canal):
    """Find email for a canal (case-insensitive substring match)"""
    if remitentes_df is None or remitentes_df.empty:
        return None, None
    canal_lower = str(canal).lower()
    email_col = next((c for c in remitentes_df.columns if "mail" in c.lower()), None)
    canal_col = next((c for c in remitentes_df.columns if "canal" in c.lower() or "channel" in c.lower()), None)
    nombre_col = next((c for c in remitentes_df.columns if "nombre" in c.lower() or "name" in c.lower()), None)
    if not email_col or not canal_col:
        return None, None
    for _, row in remitentes_df.iterrows():
        if str(row[canal_col]).lower() in canal_lower or canal_lower in str(row[canal_col]).lower():
            email = str(row[email_col]).strip()
            nombre = str(row[nombre_col]).strip() if nombre_col else canal
            return email, nombre
    return None, None

def send_cancel_email(smtp_server, smtp_port, smtp_user, smtp_pass,
                      to_email, canal_nombre, pedidos_list):
    """Send cancellation notification email"""
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Pedidos a cancelar — {canal_nombre}"
    msg["From"] = smtp_user
    msg["To"] = to_email

    rows_html = "".join(
        f"""<tr style="background:{'#f8fafc' if i%2==0 else '#fff'}">
            <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0">{p.get('Pedido','')}</td>
            <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0">{p.get('Id Marketplace','')}</td>
            <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0"><code>{p.get('SKU Original','')}</code></td>
            <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0">{p.get('País','')}</td>
        </tr>"""
        for i, p in enumerate(pedidos_list)
    )

    html = f"""
    <html><body style="font-family:Arial,sans-serif;color:#1e293b;max-width:700px;margin:0 auto">
    <div style="background:#1B2A4A;padding:20px 28px;border-radius:10px 10px 0 0">
        <h2 style="color:#fff;margin:0;font-size:20px">📦 Pedidos a cancelar</h2>
        <p style="color:#93c5fd;margin:6px 0 0;font-size:14px">Canal: <b>{canal_nombre}</b></p>
    </div>
    <div style="padding:20px 28px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 10px 10px">
        <p style="color:#64748b;font-size:14px">Los siguientes pedidos deben ser cancelados por precio por debajo de tarifa mínima o SKU no encontrado en tarifa:</p>
        <table style="width:100%;border-collapse:collapse;font-size:13px">
            <thead>
                <tr style="background:#1B2A4A;color:#fff">
                    <th style="padding:10px 12px;text-align:left">Pedido</th>
                    <th style="padding:10px 12px;text-align:left">ID Marketplace</th>
                    <th style="padding:10px 12px;text-align:left">SKU</th>
                    <th style="padding:10px 12px;text-align:left">País</th>
                </tr>
            </thead>
            <tbody>{rows_html}</tbody>
        </table>
        <p style="margin-top:20px;font-size:12px;color:#94a3b8">Generado automáticamente por Procesador de Pedidos Turaco</p>
    </div>
    </body></html>"""

    msg.attach(MIMEText(html, "html"))
    with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_user, to_email, msg.as_string())

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
    """Build per-sheet lookup {sheet_name: {sku: {min, pub, sheet}}} for national tarifa.
    Keeps each sheet isolated so marketplace-aware ordering works correctly."""
    result = {}
    for sh in order:
        df = sheets.get(sh)
        if df is None:
            continue
        result[sh] = build_sheet_lookup(df, sh)
    return result

def get_pvp(sku_norm, pais, nac_lookup, inter_lookups, marketplace=""):
    """
    Find PVP for a normalized SKU.
    - Spain: searches national sheets ordered by marketplace channel:
        Carrefour → T_C4 first, Amazon → T_AMZ first, Mediamarkt → T_MM first,
        Privalia → T_PRIV first, others → T_MIR first. Remaining sheets as fallback.
    - Other countries: search each sheet in COUNTRY_SHEETS list in order.
        Italia → [ES-IT, IT-IT], Francia → [ES-FR, FR-FR], Alemania → [ES-DE, DE-DE].
        If not found in any of the country's sheets → ❌ NO EN TARIFA.
        No cross-country fallback.
    """
    if pais in SPAIN:
        order = nac_order_for_marketplace(marketplace)
        for sh in order:
            lkp = nac_lookup.get(sh, {})
            if sku_norm in lkp:
                return lkp[sku_norm]
        return None

    sheets_for_country = COUNTRY_SHEETS.get(pais)
    if sheets_for_country:
        for sheet_key in sheets_for_country:
            lkp = inter_lookups.get(sheet_key, {})
            if sku_norm in lkp:
                return lkp[sku_norm]
        return None

    # Country not mapped → try national as best guess
    order = nac_order_for_marketplace(marketplace)
    for sh in order:
        lkp = nac_lookup.get(sh, {})
        if sku_norm in lkp:
            return lkp[sku_norm]
    return None

def analyze_row(row, nac_lookup, inter_lookups):
    sku_norm = row["_sku_norm"]
    pais = str(row.get("País", "")).strip()
    price = parse_price(row.get("Pedido.1"))
    marketplace = str(row.get("Marketplace", ""))

    pvp = get_pvp(sku_norm, pais, nac_lookup, inter_lookups, marketplace)

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
    st.markdown("### 📧 Remitentes y email")
    remitentes_file = st.file_uploader(
        "Fichero remitentes (.xlsx) — Columnas: Canal, Email, Nombre",
        type="xlsx", key="remitentes"
    )
    with st.expander("⚙️ Config. SMTP (para envío de emails)", expanded=False):
        smtp_server = st.text_input("Servidor SMTP", value="smtp.gmail.com", key="smtp_srv")
        smtp_port   = st.number_input("Puerto", value=465, key="smtp_port")
        smtp_user   = st.text_input("Usuario (email remitente)", key="smtp_user")
        smtp_pass   = st.text_input("Contraseña / App password", type="password", key="smtp_pass")

    st.markdown("---")
    run_btn = st.button("▶ Procesar", type="primary", use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
st.title("📦 Procesador de Pedidos")

if not run_btn:
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


# ── Read SMTP config from session_state (set by sidebar widgets) ──────────────
smtp_server = st.session_state.get("smtp_srv", "smtp.gmail.com")
smtp_port   = int(st.session_state.get("smtp_port", 465))
smtp_user   = st.session_state.get("smtp_user", "")
smtp_pass   = st.session_state.get("smtp_pass", "")

# ── Load remitentes ──────────────────────────────────────────────────────────
remitentes_df = None
if remitentes_file:
    try:
        remitentes_df = load_remitentes(remitentes_file.read())
        st.sidebar.success(f"✅ Remitentes: {len(remitentes_df)} canales cargados")
    except Exception as e:
        st.sidebar.warning(f"⚠️ No se pudo leer remitentes: {e}")

# ── Validate uploads ──────────────────────────────────────────────────────────
errors = []
if not nac_file: errors.append("Tarifa Nacional")
if not inter_file: errors.append("Tarifa Internacional")
if not es_miravia and not libro_file: errors.append("Fichero Rentabilidad")
if es_miravia and not miravia_file: errors.append("PagoAceptadoMiravia.xlsx")

if errors:
    st.error(f"Faltan ficheros: {', '.join(errors)}")
    st.stop()

# ── Load tarifas ──────────────────────────────────────────────────────────────
with st.spinner("Cargando tarifas..."):
    nac_sheets = load_tarifa(nac_file.read(), nac_file.name)
    inter_sheets = load_tarifa(inter_file.read(), inter_file.name)

    # Build lookups
    nac_lookup = build_lookup(nac_sheets, NAC_ORDER)
    # Each inter sheet gets its own isolated lookup — never mixed
    inter_lookups = {sh: build_sheet_lookup(df, sh) for sh, df in inter_sheets.items()}

st.success(f"✅ Tarifas cargadas — Nacional: {sum(len(v) for v in nac_lookup.values()):,} SKUs | Internacional: {sum(len(v) for v in inter_lookups.values()):,} refs")

# ═══════════════════════════════════════════════════════════════════════════════
# PROCESO A — Pago aceptado
# ═══════════════════════════════════════════════════════════════════════════════
if not es_miravia:
    with st.spinner("Procesando..."):
        libro_bytes = libro_file.read()
        libro_sheets = pd.read_excel(io.BytesIO(libro_bytes), sheet_name=None)

        hoja1_raw = libro_sheets.get("Hoja1", pd.DataFrame())
        hoja2_raw = libro_sheets.get("Hoja2", pd.DataFrame())
        # Detect alternate sheet names (Rentabilidad might have different names)
        sheet_names = list(libro_sheets.keys())
        if "Hoja2" not in libro_sheets and len(sheet_names) > 1:
            hoja2_raw = libro_sheets[sheet_names[1]]

        # ── Clean Hoja1
        h1_clean = clean_hoja1(hoja1_raw) if not hoja1_raw.empty else pd.DataFrame()

        # ── Duplicates Hoja1
        dupes_h1 = check_duplicates_hoja1(h1_clean) if not h1_clean.empty else pd.DataFrame()

        # ── Expand multi-SKU and analyse tarifa
        if not hoja2_raw.empty:
            df_expanded = expand_multi_sku_rows(hoja2_raw)
            results = []
            for _, row in df_expanded.iterrows():
                status, pvp_min, pvp_pub, diff_min, diff_pub, tarifa_sheet = analyze_row(
                    row, nac_lookup, inter_lookups
                )
                results.append({
                    "Pedido": row.get("Pedido", ""),
                    "Fecha": row.get("Fecha", ""),
                    "Marketplace": row.get("Marketplace", ""),
                    "Id Marketplace": row.get("Id Marketplace", ""),
                    "País": row.get("País", ""),
                    "SKU Original": row.get("_sku_orig", row.get("Sku", "")),
                    "SKU": row.get("Sku", ""),
                    "SKU Norm.": row.get("_sku_norm", ""),
                    "Multi-SKU": "✔" if row.get("_multi_flag") else "",
                    "Cant": row.get("Cant", ""),
                    "Precio Pedido (€)": parse_price(row.get("Pedido.1")),
                    "Hoja Tarifa": tarifa_sheet,
                    "PVP Mín (€)": pvp_min,
                    "PVP Pub (€)": pvp_pub,
                    "Dif vs Mín (€)": diff_min,
                    "Dif vs Pub (€)": diff_pub,
                    "Estado": status,
                })
            df_tarifa = pd.DataFrame(results)
            multi_count = df_expanded["_multi_flag"].sum()
        else:
            df_tarifa = pd.DataFrame()
            multi_count = 0

    # ── Display
    st.markdown("---")
    st.markdown("## 🛒 Pago aceptado — Resultados")

    ok = (df_tarifa["Estado"] == "✅ OK").sum() if not df_tarifa.empty else 0
    warn = df_tarifa["Estado"].str.startswith("🔴").sum() + df_tarifa["Estado"].str.startswith("🟡").sum() if not df_tarifa.empty else 0
    cancel = df_tarifa["Estado"].str.startswith("❌").sum() if not df_tarifa.empty else 0

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("📋 Hoja1 filas", len(h1_clean))
    col2.metric("⚠️ Duplicados", len(dupes_h1))
    col3.metric("✅ OK tarifa", ok)
    col4.metric("🔴 Bajo mínimo", warn)
    col5.metric("❌ No en tarifa", cancel)

    # ── Tabs: Análisis Tarifa / Duplicados / Hoja1 limpia ─────────────────────
    tab1, tab2, tab3 = st.tabs(["💰 Análisis Tarifa", "🔍 Duplicados Hoja1", "📄 Hoja1 limpia"])

    with tab1:
        if not df_tarifa.empty:
            if multi_count > 0:
                st.info(f"🔀 Se han expandido pedidos multi-SKU: **{multi_count}** líneas generadas por separación de SKUs en la misma celda")

            def color_status(val):
                colors = {
                    "✅ OK": "background-color:#d1fae5;color:#065f46",
                    "🟡 EN MÍNIMO": "background-color:#fef3c7;color:#92400e",
                    "🔴 BAJO MÍNIMO": "background-color:#fee2e2;color:#991b1b",
                    "❌ NO EN TARIFA": "background-color:#fce7f3;color:#831843",
                    "⚠️ SIN PRECIO": "background-color:#fef3c7;color:#92400e",
                }
                return colors.get(val, "")

            styled = df_tarifa.style.map(color_status, subset=["Estado"])
            st.dataframe(styled, use_container_width=True, hide_index=True)
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

    # ── Cancelados tracker (ALWAYS visible outside tabs) ───────────────────────
    df_cancel_rows = df_tarifa[df_tarifa["Estado"].str.startswith("❌")].copy() if not df_tarifa.empty else pd.DataFrame()
    if not df_cancel_rows.empty:
        st.markdown("---")
        st.markdown("### 📋 Gestión de cancelados")
        st.caption("Marca cada cancelado como enviado al canal. Los enviados quedan sombreados.")

        if "cancelados_enviados_a" not in st.session_state:
            st.session_state["cancelados_enviados_a"] = set()

        marketplaces = df_cancel_rows["Marketplace"].unique().tolist()

        col_mark, col_clear = st.columns([1, 4])
        with col_mark:
            if st.button("✅ Marcar todos", key="mark_all_a"):
                st.session_state["cancelados_enviados_a"] = set(df_cancel_rows.index.tolist())
                st.rerun()
        with col_clear:
            if st.button("↩ Limpiar marcas", key="clear_a"):
                st.session_state["cancelados_enviados_a"] = set()
                st.rerun()

        # Send by channel buttons
        if smtp_user and smtp_pass:
            st.markdown("**📧 Enviar cancelados por canal:**")
            btn_cols = st.columns(min(len(marketplaces), 4))
            for ci, mkt in enumerate(marketplaces):
                mkt_rows = df_cancel_rows[df_cancel_rows["Marketplace"] == mkt]
                email, nombre = get_remitente(remitentes_df, mkt) if remitentes_df is not None else (None, None)
                with btn_cols[ci % 4]:
                    disabled = email is None
                    help_txt = f"→ {email}" if email else "Sin email en remitentes para este canal"
                    if st.button(f"✉️ {mkt[:20]}", key=f"send_a_{ci}", disabled=disabled, help=help_txt):
                        try:
                            send_cancel_email(smtp_server, int(smtp_port), smtp_user, smtp_pass,
                                              email, nombre or mkt, mkt_rows.to_dict("records"))
                            for idx in mkt_rows.index:
                                st.session_state["cancelados_enviados_a"].add(idx)
                            st.success(f"✅ Email enviado a {email} ({len(mkt_rows)} pedidos)")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Error: {e}")
        else:
            st.info("💡 Configura SMTP en el sidebar y sube el fichero de remitentes para enviar emails por canal.")

        # Per-row cards
        for i, (idx, row) in enumerate(df_cancel_rows.iterrows()):
            enviado = idx in st.session_state["cancelados_enviados_a"]
            bg = "#f0fdf4" if enviado else "#fff1f2"
            border = "#86efac" if enviado else "#fca5a5"
            mkt = row.get("Marketplace", "")
            email, _ = get_remitente(remitentes_df, mkt) if remitentes_df is not None else (None, None)
            email_tag = f'<small style="color:#64748b"> → {email}</small>' if email else ""
            st.markdown(
                f"""<div style="background:{bg};border:1.5px solid {border};border-radius:8px;
                padding:10px 16px;margin-bottom:4px;display:flex;align-items:center;gap:12px;">
                <span style="font-size:13px;color:#374151;">
                <b>Pedido {row.get('Pedido','')}</b> &nbsp;·&nbsp;
                {mkt}{email_tag} &nbsp;·&nbsp;
                {row.get('Id Marketplace','')} &nbsp;·&nbsp;
                SKU <code>{row.get('SKU Original','')}</code> &nbsp;·&nbsp;
                {row.get('País','')}
                </span>
                <span style="margin-left:auto;font-weight:600;color:{'#15803d' if enviado else '#9f1239'}">
                {'✅ Enviado' if enviado else '⬜ Pendiente'}
                </span></div>""",
                unsafe_allow_html=True
            )
            checked = st.checkbox("Marcar como enviado", value=enviado, key=f"cancel_a_{idx}_{i}")
            if checked != enviado:
                if checked:
                    st.session_state["cancelados_enviados_a"].add(idx)
                else:
                    st.session_state["cancelados_enviados_a"].discard(idx)
                st.rerun()

    # ── Export
    st.markdown("---")
    sections = [
        ("Análisis Tarifa", df_tarifa, None),
        ("Duplicados Hoja1",
         dupes_h1 if len(dupes_h1) > 0 else None,
         "Sin duplicados encontrados" if len(dupes_h1) == 0 else f"⚠️ {len(dupes_h1)} duplicados detectados"),
        ("Hoja1 limpia", h1_clean, None),
    ]
    excel_bytes = build_excel([(s, d, t) for s, d, t in sections if d is not None or t is not None])
    st.download_button(
        "⬇️ Descargar Excel completo",
        data=excel_bytes,
        file_name=f"PagoAceptado_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

# ═══════════════════════════════════════════════════════════════════════════════
# PROCESO B — Pago aceptado Miravia
# ═══════════════════════════════════════════════════════════════════════════════
else:
    with st.spinner("Procesando Miravia..."):
        mir_bytes = miravia_file.read()
        mir_sheets = pd.read_excel(io.BytesIO(mir_bytes), sheet_name=None)
        hoja1_raw = mir_sheets.get("Hoja1", pd.DataFrame())
        hoja2_raw = mir_sheets.get("Hoja2", pd.DataFrame())

        # Load cancelados if provided
        df_cancelados = None
        if cancelados_file:
            can_bytes = cancelados_file.read()
            can_sheets = pd.read_excel(io.BytesIO(can_bytes), sheet_name=None)
            df_cancelados = list(can_sheets.values())[0]

        # Clean Hoja1
        h1_clean = clean_hoja1(hoja1_raw) if not hoja1_raw.empty else pd.DataFrame()

        # Duplicates Combination
        dupes_comb, id_map = check_duplicates_miravia(h1_clean) if not h1_clean.empty else (pd.DataFrame(), {})

        # Cross-reference cancelados
        df_cancel_match = cross_miravia_cancelados(h1_clean, df_cancelados) if df_cancelados is not None else pd.DataFrame()

        # Tarifa analysis (Hoja2 if exists)
        df_tarifa = pd.DataFrame()
        multi_count = 0
        if not hoja2_raw.empty:
            df_expanded = expand_multi_sku_rows(hoja2_raw)
            results = []
            for _, row in df_expanded.iterrows():
                status, pvp_min, pvp_pub, diff_min, diff_pub, tarifa_sheet = analyze_row(
                    row, nac_lookup, inter_lookups
                )
                results.append({
                    "Pedido": row.get("Pedido", ""), "SKU Original": row.get("_sku_orig", ""),
                    "SKU": row.get("Sku", ""), "SKU Norm.": row.get("_sku_norm", ""),
                    "Multi-SKU": "✔" if row.get("_multi_flag") else "",
                    "País": row.get("País", ""),
                    "Precio Pedido (€)": parse_price(row.get("Pedido.1")),
                    "Hoja Tarifa": tarifa_sheet,
                    "PVP Mín (€)": pvp_min, "PVP Pub (€)": pvp_pub,
                    "Dif vs Mín (€)": diff_min, "Estado": status,
                })
            df_tarifa = pd.DataFrame(results)
            multi_count = df_expanded["_multi_flag"].sum()

    # ── Display
    st.markdown("---")
    st.markdown("## 🏪 Pago aceptado Miravia — Resultados")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📋 Pedidos Miravia", len(h1_clean))
    col2.metric("🔴 Cancelados match", len(df_cancel_match))
    col3.metric("⚠️ Dupl. Combination", len(dupes_comb))
    col4.metric("💰 Tarifa analizados", len(df_tarifa))

    tabs = st.tabs(["⚠️ Duplicados Combination", "💰 Análisis Tarifa", "📄 Hoja1 limpia"])

    with tabs[0]:
        if len(dupes_comb) > 0:
            st.warning(f"⚠️ **{len(dupes_comb)}** filas con ID Combination duplicado")
            st.dataframe(dupes_comb, use_container_width=True, hide_index=True)
        else:
            st.success("✅ Sin IDs duplicados en columna Combination")

    with tabs[1]:
        if not df_tarifa.empty:
            if multi_count > 0:
                st.info(f"🔀 **{multi_count}** líneas generadas por expansión de pedidos multi-SKU")

            def color_status(val):
                colors = {
                    "✅ OK": "background-color:#d1fae5;color:#065f46",
                    "🟡 EN MÍNIMO": "background-color:#fef3c7;color:#92400e",
                    "🔴 BAJO MÍNIMO": "background-color:#fee2e2;color:#991b1b",
                    "❌ NO EN TARIFA": "background-color:#fce7f3;color:#831843",
                    "⚠️ SIN PRECIO": "background-color:#fef3c7;color:#92400e",
                }
                return colors.get(val, "")

            styled = df_tarifa.style.map(color_status, subset=["Estado"])
            st.dataframe(styled, use_container_width=True, hide_index=True)
        else:
            st.info("Sin datos en Hoja2 para análisis de tarifa (habitual en Miravia).")

    with tabs[2]:
        st.dataframe(h1_clean, use_container_width=True, hide_index=True)

    # ── Cancelados tracker (ALWAYS visible outside tabs) ───────────────────────
    st.markdown("---")
    st.markdown("### 📋 Gestión de cancelados Miravia")

    if cancelados_file is None:
        st.info("💡 Sube el fichero ES de cancelados en el sidebar para cruzar pedidos.")
    elif len(df_cancel_match) == 0:
        n_ids = sum(1 for v in id_map.values() if v)
        n_can = len(df_cancelados) if df_cancelados is not None else 0
        st.success(f"✅ Ningún pedido Miravia encontrado en cancelados ({n_ids} IDs cruzados contra {n_can:,} cancelados)")
    else:
        st.error(f"🔴 **{len(df_cancel_match)}** pedidos Miravia encontrados en el fichero de cancelados")

        if "cancelados_enviados_b" not in st.session_state:
            st.session_state["cancelados_enviados_b"] = set()

        col_mark, col_clear = st.columns([1, 4])
        with col_mark:
            if st.button("✅ Marcar todos", key="mark_all_b"):
                st.session_state["cancelados_enviados_b"] = set(range(len(df_cancel_match)))
                st.rerun()
        with col_clear:
            if st.button("↩ Limpiar marcas", key="clear_b"):
                st.session_state["cancelados_enviados_b"] = set()
                st.rerun()

        # Email send button
        if smtp_user and smtp_pass:
            email_mir, nombre_mir = get_remitente(remitentes_df, "Miravia") if remitentes_df is not None else (None, None)
            disabled_mir = email_mir is None
            help_mir = f"→ {email_mir}" if email_mir else "Sin email para Miravia en remitentes"
            lbl_mir = f"✉️ Enviar a Miravia ({email_mir})" if email_mir else "✉️ Enviar a Miravia"
            if st.button(lbl_mir, key="send_b_miravia", disabled=disabled_mir, help=help_mir):
                try:
                    pedidos_list = [
                        {"Pedido": r.get("ID Pedido",""), "Id Marketplace": r.get("ID Extraído",""),
                         "SKU Original": r.get("SKU",""), "País": ""}
                        for _, r in df_cancel_match.iterrows()
                    ]
                    send_cancel_email(smtp_server, int(smtp_port), smtp_user, smtp_pass,
                                      email_mir, nombre_mir or "Miravia", pedidos_list)
                    st.session_state["cancelados_enviados_b"] = set(range(len(df_cancel_match)))
                    st.success(f"✅ Email enviado a {email_mir} ({len(pedidos_list)} pedidos)")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error al enviar: {e}")
        else:
            st.info("💡 Configura SMTP en el sidebar y sube el fichero de remitentes para enviar emails.")

        # Per-row cards
        for i, (_, row) in enumerate(df_cancel_match.iterrows()):
            enviado = i in st.session_state["cancelados_enviados_b"]
            bg = "#f0fdf4" if enviado else "#fff1f2"
            border = "#86efac" if enviado else "#fca5a5"
            st.markdown(
                f"""<div style="background:{bg};border:1.5px solid {border};border-radius:8px;
                padding:10px 16px;margin-bottom:4px;display:flex;align-items:center;gap:12px;">
                <span style="font-size:13px;color:#374151;">
                <b>Pedido {row.get('ID Pedido','')}</b> &nbsp;·&nbsp;
                Combination: <code>{row.get('Combination','')}</code> &nbsp;·&nbsp;
                ID: <code>{row.get('ID Extraído','')}</code> &nbsp;·&nbsp;
                SKU: <code>{row.get('SKU','')}</code> &nbsp;·&nbsp;
                {row.get('Cliente','')}
                </span>
                <span style="margin-left:auto;font-weight:600;color:{'#15803d' if enviado else '#9f1239'}">
                {'✅ Enviado' if enviado else '⬜ Pendiente'}
                </span></div>""",
                unsafe_allow_html=True
            )
            checked = st.checkbox("Marcar como enviado", value=enviado, key=f"cancel_b_{i}")
            if checked != enviado:
                if checked:
                    st.session_state["cancelados_enviados_b"].add(i)
                else:
                    st.session_state["cancelados_enviados_b"].discard(i)
                st.rerun()

    # ── Export
    st.markdown("---")
    sections = [
        ("Cancelados Match", df_cancel_match if len(df_cancel_match) > 0 else None,
         "Sin cancelados encontrados" if len(df_cancel_match) == 0 else None),
        ("Duplicados Combination", dupes_comb if len(dupes_comb) > 0 else None,
         "Sin duplicados en Combination" if len(dupes_comb) == 0 else None),
        ("Análisis Tarifa", df_tarifa if not df_tarifa.empty else None, None),
        ("Hoja1 limpia", h1_clean, None),
    ]
    excel_bytes = build_excel([(s, d, t) for s, d, t in sections if d is not None or t is not None])
    st.download_button(
        "⬇️ Descargar Excel completo",
        data=excel_bytes,
        file_name=f"PagoAceptadoMiravia_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
