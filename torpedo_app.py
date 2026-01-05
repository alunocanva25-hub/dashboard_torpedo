import streamlit as st
import pandas as pd
import plotly.express as px
import requests, re
from io import BytesIO
from datetime import date, timedelta
import streamlit.components.v1 as components

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors


# ======================================================
# CONFIG
# ======================================================
st.set_page_config(page_title="TORPEDO SEMANAL – Produtividade", layout="wide")


# ======================================================
# CSS
# ======================================================
st.markdown("""
<style>
:root{
  --bg:#6fa6d6;
  --card:#b9d3ee;
  --ink:#0b2b45;
  --red:#9b0d0d;
  --stroke: rgba(10,40,70,0.30);
  --stroke2: rgba(10,40,70,0.22);
  --glass: rgba(255,255,255,0.55);
  --glass2: rgba(255,255,255,0.35);
}

.stApp { background: var(--bg); }
.block-container{ padding-top: 0.6rem; max-width: 1500px; }

/* Cards */
.card{
  background: var(--card);
  border: 2px solid var(--stroke);
  border-radius: 18px;
  padding: 14px 16px;
  box-shadow: 0 10px 18px rgba(0,0,0,0.18);
  margin-bottom: 14px;
  text-align: center;
}
.card-title{
  font-weight: 950;
  color:var(--ink);
  font-size: 13px;
  text-transform: uppercase;
  margin-bottom: 10px;
  letter-spacing: .3px;
  text-align: center;
}

/* KPI */
.kpi-row{
  display:flex;
  justify-content:space-between;
  align-items:flex-end;
  gap: 10px;
}
.kpi-big{
  font-size: 42px;
  font-weight: 950;
  color:var(--red);
  line-height: 1.0;
}

/* Topbar */
.topbar{
  background: var(--glass2);
  border: 2px solid var(--stroke2);
  border-radius: 18px;
  padding: 10px 14px;
  display:flex;
  justify-content:space-between;
  align-items:center;
  margin-bottom: 10px;
}
.brand{ display:flex; align-items:center; gap:12px; }
.brand-badge{
  width:46px; height:46px; border-radius: 14px;
  background: var(--glass);
  border: 2px solid var(--stroke2);
  display:flex; align-items:center; justify-content:center;
  font-weight: 950; color:var(--ink);
}
.brand-text .t1{ font-weight:950; color:var(--ink); line-height:1.1; }
.brand-text .t2{ font-weight:800; color:var(--ink); opacity:.85; font-size:12px; }
.right-note{ text-align:right; font-weight:950; color:var(--ink); }
.right-note small{ font-weight:800; opacity:.9; font-size:12px; }

/* Botões */
div.stButton > button{
  border-radius: 10px;
  font-weight: 900;
  border: 2px solid var(--stroke2);
  background: rgba(255,255,255,0.45);
  color:var(--ink);
  padding: .35rem .7rem;
}
div.stButton > button:hover{
  background: rgba(255,255,255,0.65);
  border-color: rgba(10,40,70,0.35);
}

/* Segmented (padrão) */
div[data-baseweb="segmented-control"]{
  background: rgba(255,255,255,0.35);
  border: 2px solid rgba(10,40,70,0.22);
  border-radius: 14px;
  padding: 6px;
}
div[data-baseweb="segmented-control"] span{
  font-weight: 900 !important;
  color: #0b2b45 !important;
}
div[data-baseweb="segmented-control"] div[aria-checked="true"]{
  background: #0b2b45 !important;
  border-radius: 10px !important;
}
div[data-baseweb="segmented-control"] div[aria-checked="true"] span{
  color: #ffffff !important;
}

/* ===== LOCALIDADE: 1 LINHA + SCROLL (NÃO QUEBRA) ===== */
.seg-local div[data-baseweb="segmented-control"]{
  width: 100% !important;
  overflow-x: auto !important;
  overflow-y: hidden !important;
  -webkit-overflow-scrolling: touch !important;
  white-space: nowrap !important;
}
.seg-local div[data-baseweb="segmented-control"] > div,
.seg-local div[data-baseweb="segmented-control"] > div > div,
.seg-local div[data-baseweb="segmented-control"] [role="radiogroup"]{
  display: inline-flex !important;
  flex-wrap: nowrap !important;
  width: max-content !important;
  gap: 6px !important;
}
.seg-local div[data-baseweb="segmented-control"] [role="radio"]{
  flex: 0 0 auto !important;
  white-space: nowrap !important;
  min-width: 72px !important;
  padding: 6px 10px !important;
}
.seg-local div[data-baseweb="segmented-control"]::-webkit-scrollbar{
  height: 6px;
}
.seg-local div[data-baseweb="segmented-control"]::-webkit-scrollbar-thumb{
  background: rgba(11,43,69,0.35);
  border-radius: 999px;
}

/* Tabelas estilo torpedo */
.tblwrap{
  border-radius: 14px;
  overflow: hidden;
  border: 1px solid rgba(10,40,70,0.25);
  background: rgba(255,255,255,0.35);
}
.tblhead{
  padding: 9px 10px;
  font-weight: 950;
  color: white;
  text-align: left;
  letter-spacing: .2px;
}
.head-blue{ background:#1F77B4; }
.head-green{ background:#2CA02C; }
.head-yellow{ background:#F1C40F; color:#1b1b1b; }

.tbl{ width: 100%; border-collapse: collapse; }
.tbl td{
  border-top: 1px solid rgba(10,40,70,0.18);
  padding: 8px 10px;
  font-size: 13px;
  color:#0b2b45;
  font-weight: 900;
}
.col-date{ width: 120px; }
.col-dow{ width: 62px; text-align:center; }

/* LOGIN */
.login-top-space { height: 18px; }
</style>
""", unsafe_allow_html=True)


# ======================================================
# CONSTANTES
# ======================================================
DOW_PT = {0: "SEG", 1: "TER", 2: "QUA", 3: "QUI", 4: "SEX", 5: "SÁB", 6: "DOM"}

OPCOES_DEMANDA = [
    "BAIXA DE LAUDO",
    "VIAGEM AO IMMETRO-PA",
    "ACOMPANHAMENTO APCL/APJL",
    "APOIO VERIFICAÇÃO/TRIAGEM",
    "ANEXO DE AR",
    "DIGITALIZAÇÃO DE AR",
    "RENOMEAÇÃO DE AR",
    "ANEXO/DIGITALIZAÇÃO/RENOMEAR - AR",
]


# ======================================================
# HELPERS
# ======================================================
def fmt_int(n: int) -> str:
    return f"{int(n):,}".replace(",", ".")

def monday_of_week(d: date) -> date:
    return d - timedelta(days=d.weekday())

def _extrair_drive_id(url: str):
    m = re.search(r"[?&]id=([a-zA-Z0-9-_]+)", url)
    if m:
        return m.group(1)
    m = re.search(r"/file/d/([a-zA-Z0-9-_]+)", url)
    if m:
        return m.group(1)
    return None

def _drive_direct_download(url: str) -> str:
    did = _extrair_drive_id(url)
    if did:
        return f"https://drive.google.com/uc?id={did}"
    return url

def _bytes_is_html(raw: bytes) -> bool:
    head = raw[:800].lstrip().lower()
    return head.startswith(b"<!doctype html") or b"<html" in head

def _bytes_is_xlsx(raw: bytes) -> bool:
    return raw[:2] == b"PK"

@st.cache_data(ttl=600, show_spinner="🔄 Carregando base (XLSX/CSV)...")
def carregar_base(url_original: str) -> pd.DataFrame:
    url = _drive_direct_download(url_original)
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    raw = r.content

    if _bytes_is_html(raw):
        raise RuntimeError("URL retornou HTML (provável permissão/link). No Drive: 'Qualquer pessoa com o link' (Visualizador).")

    if _bytes_is_xlsx(raw):
        return pd.read_excel(BytesIO(raw), sheet_name=0, engine="openpyxl")

    for enc in ["utf-8-sig", "utf-8", "cp1252", "latin1"]:
        try:
            return pd.read_csv(BytesIO(raw), sep=None, engine="python", encoding=enc)
        except UnicodeDecodeError:
            continue
    return pd.read_csv(BytesIO(raw), sep=None, engine="python", encoding="utf-8", encoding_errors="replace")

def validar_estrutura_posicional(df: pd.DataFrame):
    if df is None or df.empty:
        st.error("Base vazia.")
        st.stop()
    if len(df.columns) < 8:
        st.error("A base precisa ter pelo menos até a coluna H (8 colunas).")
        st.stop()

def normalize_colab_series(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.upper().str.strip()
    return s.replace({"": None, "NAN": None, "NONE": None, "NULL": None, "-": None})

def html_torpedo_table(title: str, head_class: str, df_rows: pd.DataFrame) -> str:
    linhas = []
    for _, r in df_rows.iterrows():
        dstr = pd.to_datetime(r["Data"]).strftime("%d/%m/%Y")
        linhas.append(
            f"<tr>"
            f"<td class='col-date'>{dstr}</td>"
            f"<td class='col-dow'>{r['DOW']}</td>"
            f"<td>{r['Demanda']}</td>"
            f"</tr>"
        )
    body = "\n".join(linhas)
    return f"""
    <div class="tblwrap">
      <div class="tblhead {head_class}">{title}</div>
      <table class="tbl"><tbody>{body}</tbody></table>
    </div>
    """

def achar_coluna_por_nome(df: pd.DataFrame, nomes_possiveis: list[str]):
    cols = [str(c).upper().strip() for c in df.columns]
    for i, c in enumerate(cols):
        for n in nomes_possiveis:
            if n in c:
                return df.columns[i]
    return None


# ======================================================
# DONUT
# ======================================================
def donut_colaborador_acumulado(df_base: pd.DataFrame, ano_ref: int | None):
    if df_base.empty:
        return None, 0

    base = df_base.copy()
    if ano_ref is not None:
        base = base[base[COL_DATA].dt.year == int(ano_ref)].copy()

    base = base.dropna(subset=["_COLAB_"]).copy()
    if base.empty:
        return None, 0

    dados = (
        base.groupby("_COLAB_")["_QTD_"]
        .sum()
        .reset_index()
        .rename(columns={"_COLAB_": "Colaborador", "_QTD_": "Notas"})
        .sort_values("Notas", ascending=False)
    )

    total = int(dados["Notas"].sum())

    fig = px.pie(
        dados,
        names="Colaborador",
        values="Notas",
        hole=0.65,
        template="plotly_white"
    )
    fig.update_layout(
        height=320,
        margin=dict(l=10, r=10, t=55, b=10),
        legend_title_text="",
        title=f"NOTAS ATENDIDAS POR COLABORADOR - {ano_ref if ano_ref else ''}".strip()
    )
    fig.update_traces(textinfo="value")

    fig.add_annotation(
        x=0.5, y=0.5, xref="paper", yref="paper",
        text=f"<b>{fmt_int(total)}</b><br><span style='font-size:11px'>TOTAL</span>",
        showarrow=False
    )

    return fig, total


# ======================================================
# LOGIN
# ======================================================
def tela_login():
    st.markdown("""
    <style>
      .login-frame { width: min(980px, 96vw); margin: 0 auto; }
      .login-card{
        border-radius: 26px; overflow: hidden;
        border: 2px solid rgba(10,40,70,0.25);
        box-shadow: 0 16px 40px rgba(0,0,0,0.22);
        background: rgba(255,255,255,0.40);
        backdrop-filter: blur(8px);
      }
      .login-header{ padding: 22px; background: #2f6f97; color: white; }
      .login-header .h1{ font-size: 34px; font-weight: 950; }
      .login-header .h2{ font-size: 18px; font-weight: 800; opacity: .95; }
      .login-body{ padding: 18px 22px; background: rgba(255,255,255,0.22); }

      div[data-testid="stTextInput"] label{ display:none !important; }
      div[data-testid="stTextInput"] input{
        width: 100% !important; height: 56px !important;
        border-radius: 10px !important;
        border: 2px solid rgba(10,40,70,0.25) !important;
        background: rgba(20,20,25,0.88) !important;
        padding: 0 16px !important;
        font-weight: 900 !important;
        font-size: 20px !important;
        color: #ffffff !important;
        box-sizing: border-box !important;
      }
      div[data-testid="stTextInput"] input::placeholder{ color: rgba(255,255,255,0.65); }

      .login-btns div.stButton > button{
        width: 100%; height: 56px;
        border-radius: 12px;
        border: 2px solid rgba(10,40,70,0.22);
        background: rgba(255,255,255,0.35);
        color: #0b2b45;
        font-weight: 950;
        font-size: 18px;
      }
      .login-btns div.stButton > button:hover{ background: rgba(255,255,255,0.55); }

      .login-note{
        margin-top: 10px;
        font-size: 12px;
        font-weight: 900;
        color: rgba(11,43,69,0.85);
        text-align: center;
      }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<div class='login-top-space'></div>", unsafe_allow_html=True)

    _, center, _ = st.columns([1, 6, 1])
    with center:
        st.markdown("<div class='login-frame'>", unsafe_allow_html=True)

        st.markdown("""
        <div class="login-card">
          <div class="login-header">
            <div class="h1">🔐 Acesso Restrito</div>
            <div class="h2">Torpedo Semanal • Produtividade</div>
          </div>
          <div class="login-body">
        """, unsafe_allow_html=True)

        st.markdown("<div style='margin-top:30px'></div>", unsafe_allow_html=True)
        st.text_input("", key="login_usuario", placeholder="Digite seu usuário")
        st.text_input("", key="login_senha", type="password", placeholder="Digite sua senha")

        b1, b2 = st.columns(2, gap="medium")
        with b1:
            entrar = st.button("Entrar")
        with b2:
            limpar = st.button("Limpar")

        st.markdown("""<div class="login-note">✅ Segurança via <b>st.secrets</b></div>""", unsafe_allow_html=True)
        st.markdown("""</div></div>""", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    if limpar:
        st.session_state["login_usuario"] = ""
        st.session_state["login_senha"] = ""
        st.rerun()

    if entrar:
        try:
            if (
                st.session_state.get("login_usuario") == st.secrets["auth"]["usuario"]
                and st.session_state.get("login_senha") == st.secrets["auth"]["senha"]
            ):
                st.session_state["logado"] = True
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos")
        except Exception:
            st.error("Secrets não configurado no Streamlit Cloud.")


# ======================================================
# SESSÃO
# ======================================================
if "logado" not in st.session_state:
    st.session_state["logado"] = False

if not st.session_state["logado"]:
    tela_login()
    st.stop()


# ======================================================
# TOPO
# ======================================================
st.markdown("""
<div class="topbar">
  <div class="brand">
    <div class="brand-badge">3C</div>
    <div class="brand-text">
      <div class="t1">TORPEDO SEMANAL – PRODUTIVIDADE</div>
      <div class="t2">Barras diárias por colaborador + Donut acumulado</div>
    </div>
  </div>
  <div class="right-note">
    BASE DRIVE (XLSX)<br>
    <small>Colab = H • Notas = B • Tipo = C • Localidade = D • Data (baixa) = E</small>
  </div>
</div>
""", unsafe_allow_html=True)


# ======================================================
# ATUALIZAR BASE
# ======================================================
colA, colB = st.columns([1, 6])
with colA:
    if st.button("🔄 Atualizar base"):
        st.cache_data.clear()
        st.rerun()
with colB:
    st.caption("Use quando atualizar o arquivo no Drive (XLSX).")


# ======================================================
# CARREGAMENTO
# ======================================================
URL_BASE = "https://drive.google.com/uc?id=1VadynN01W4mNRLfq8ABZAaQP8Sfim5tb"

try:
    df = carregar_base(URL_BASE)
    validar_estrutura_posicional(df)
except Exception as e:
    st.error(str(e))
    st.stop()


# ======================================================
# MAPEAMENTO FIXO (B/C/D/E/H)
# ======================================================
COL_NOTAS = df.columns[1]   # B
COL_TIPO  = df.columns[2]   # C
COL_LOCAL = df.columns[3]   # D
COL_DATA  = df.columns[4]   # E
COL_COLAB = df.columns[7]   # H

df = df.copy()
df[COL_DATA] = pd.to_datetime(df[COL_DATA], errors="coerce", dayfirst=True)
df = df.dropna(subset=[COL_DATA]).copy()

df["_COLAB_"] = normalize_colab_series(df[COL_COLAB])
df["_TIPO_"]  = df[COL_TIPO].astype(str).str.upper().str.strip()
df["_LOCAL_"] = df[COL_LOCAL].astype(str).str.upper().str.strip()
df["_NOTA_ID_"] = df[COL_NOTAS].astype(str).str.strip()
df["_QTD_"] = 1

COL_RESULTADO = achar_coluna_por_nome(df, ["RESULTADO", "SITUA", "STATUS", "PARECER"])
df["_RES_"] = df[COL_RESULTADO].astype(str).str.upper().str.strip() if COL_RESULTADO else ""


# ======================================================
# SELETORES (Ano • Período • Calendário • Semana)
# ======================================================
anos_disponiveis = sorted(df[COL_DATA].dropna().dt.year.unique().astype(int).tolist())
c_sel1, c_sel2, c_sel3, c_sel4 = st.columns([1.0, 1.3, 2.2, 2.0], gap="medium")

with c_sel1:
    if anos_disponiveis:
        ano_sel = st.selectbox("Ano", options=anos_disponiveis, index=len(anos_disponiveis) - 1, key="ano_sel")
    else:
        ano_sel = None
        st.selectbox("Ano", options=["—"], index=0, key="ano_sel_disabled", disabled=True)

with c_sel2:
    modo_periodo = st.segmented_control(
        "Período",
        options=["Semanal", "Mensal"],
        default=st.session_state.get("modo_periodo", "Semanal"),
        key="modo_periodo",
    )

df_ano = df if (ano_sel is None) else df[df[COL_DATA].dt.year == int(ano_sel)].copy()
ano_txt = str(ano_sel) if ano_sel else "—"

if not df_ano.empty and df_ano[COL_DATA].notna().any():
    _min_d = df_ano[COL_DATA].min().date()
    _max_d = df_ano[COL_DATA].max().date()
else:
    _min_d = date.today()
    _max_d = date.today()

saved = st.session_state.get("range_calendario", None)
if isinstance(saved, (tuple, list)) and len(saved) == 2:
    s_ini, s_fim = saved
else:
    s_ini, s_fim = _min_d, _max_d

if not isinstance(s_ini, date): s_ini = _min_d
if not isinstance(s_fim, date): s_fim = _max_d

s_ini = max(_min_d, min(_max_d, s_ini))
s_fim = max(_min_d, min(_max_d, s_fim))
if s_fim < s_ini:
    s_ini, s_fim = s_fim, s_ini

with c_sel3:
    data_ini, data_fim = st.date_input(
        "Filtro por calendário (início/fim)",
        value=(s_ini, s_fim),
        min_value=_min_d,
        max_value=_max_d,
        key="range_calendario",
    )

with c_sel4:
    semana_sel = None
    if modo_periodo == "Semanal" and (ano_sel is not None) and (not df_ano.empty):
        semanas_disp = sorted(df_ano[COL_DATA].dropna().dt.isocalendar().week.unique().astype(int).tolist())
        opcoes_sem = ["Todas"] + [f"S{w:02d}" for w in semanas_disp]
        semana_sel = st.selectbox("Semana (S01..S53)", opcoes_sem, index=0, key="semana_sel")
    else:
        st.selectbox("Semana (S01..S53)", ["—"], index=0, key="semana_sel_disabled", disabled=True)

if modo_periodo == "Semanal" and semana_sel and semana_sel != "Todas" and (ano_sel is not None):
    w = int(str(semana_sel).replace("S", ""))
    try:
        data_ini = date.fromisocalendar(int(ano_sel), w, 1)
        data_fim = date.fromisocalendar(int(ano_sel), w, 5)
        st.caption(f"Semana {semana_sel}: {data_ini.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')} (seg–sex)")
    except ValueError:
        st.warning("Semana inválida para este ano (ISO). Usando o filtro por calendário.")

df_periodo = df_ano.copy()
if not df_periodo.empty and df_periodo[COL_DATA].notna().any():
    _dini = pd.to_datetime(data_ini)
    _dfim = pd.to_datetime(data_fim) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)
    df_periodo = df_periodo[(df_periodo[COL_DATA] >= _dini) & (df_periodo[COL_DATA] <= _dfim)].copy()


# ======================================================
# FILTROS (Localidade / Tipo / Controles do gráfico NO LUGAR do colaborador opcional)
# ======================================================
locais = sorted([x for x in df_periodo["_LOCAL_"].dropna().unique().tolist() if str(x).strip()])
tipos  = sorted([x for x in df_periodo["_TIPO_"].dropna().unique().tolist() if str(x).strip()])
op_local_tabs = ["TOTAL"] + locais

c_loc, c_tipo, c_ctrl = st.columns([2.2, 1.6, 2.2], gap="medium")

with c_loc:
    st.caption("Localidade")
    st.markdown("<div class='seg-local'>", unsafe_allow_html=True)
    local_tab = st.segmented_control(
        label="",
        options=op_local_tabs,
        default=st.session_state.get("local_tab", "TOTAL"),
        key="local_tab",
    )
    st.markdown("</div>", unsafe_allow_html=True)

with c_tipo:
    tipo_sel = st.multiselect("Tipo de nota", options=tipos, default=[])

with c_ctrl:
    # Placeholder: os controles reais serão renderizados depois do df_semana existir
    st.caption("Colaboradores (para o gráfico) / Visual")
    ctrl_box = st.container()

# aplica filtros (somente Localidade + Tipo)
df_filtro = df_periodo.copy()
if local_tab and local_tab != "TOTAL":
    df_filtro = df_filtro[df_filtro["_LOCAL_"] == str(local_tab).upper().strip()]
if tipo_sel:
    df_filtro = df_filtro[df_filtro["_TIPO_"].isin([str(s).upper().strip() for s in tipo_sel])]


# ======================================================
# RANGE (seg–sex) + acumulados
# ======================================================
if modo_periodo == "Semanal":
    mon = monday_of_week(data_ini)
    week_start = pd.to_datetime(mon)
    week_end = pd.to_datetime(mon + timedelta(days=4))
else:
    week_start = pd.to_datetime(data_ini)
    week_end = pd.to_datetime(data_fim)

df_semana = df_filtro[(df_filtro[COL_DATA] >= week_start) & (df_filtro[COL_DATA] <= week_end)].copy()
df_semana["DOW_NUM"] = df_semana[COL_DATA].dt.weekday
df_semana["DOW"] = df_semana["DOW_NUM"].map(DOW_PT)
df_semana = df_semana[df_semana["DOW_NUM"].between(0, 4)].copy()

total_periodo = int(df_semana["_QTD_"].sum())
total_ano = int(df_filtro[df_filtro[COL_DATA].dt.year == int(ano_sel)]["_QTD_"].sum()) if ano_sel else int(df_filtro["_QTD_"].sum())

colabs_disp = sorted(normalize_colab_series(df_semana["_COLAB_"]).dropna().unique().tolist()) if not df_semana.empty else []

# ======================================================
# CONTROLES DO GRÁFICO (renderiza aqui, SEM mexer no session_state depois)
# ======================================================
default_colabs = colabs_disp[:6] if len(colabs_disp) > 6 else colabs_disp
# garante defaults só se a key ainda não existir (antes do widget)
if "colabs_graf" not in st.session_state:
    st.session_state["colabs_graf"] = default_colabs
if "modo_barra" not in st.session_state:
    st.session_state["modo_barra"] = "Lado a lado"

# se existir seleção antiga, filtra para não quebrar (apenas antes do widget)
st.session_state["colabs_graf"] = [c for c in st.session_state["colabs_graf"] if c in colabs_disp]

with ctrl_box:
    if df_semana.empty:
        st.info("Sem dados no período.")
    else:
        st.multiselect(
            "Colaboradores (para o gráfico)",
            options=colabs_disp,
            default=st.session_state["colabs_graf"],
            key="colabs_graf",
        )
        st.selectbox(
            "Visual",
            ["Lado a lado", "Empilhado"],
            index=0 if st.session_state.get("modo_barra", "Lado a lado") == "Lado a lado" else 1,
            key="modo_barra",
        )


# ======================================================
# LINHA PRINCIPAL: BARRAS + DONUT + RESUMO
# ======================================================
row_main = st.columns([2.2, 1.3, 1.0], gap="medium")

# ---- 1) BARRAS
with row_main[0]:
    st.markdown('<div class="card"><div class="card-title">PRODUTIVIDADE DIÁRIA — POR COLABORADOR (SEG–SEX)</div>', unsafe_allow_html=True)

    if df_semana.empty:
        st.info("Sem dados no período selecionado.")
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        colabs_sel = st.session_state.get("colabs_graf", default_colabs)
        modo_barra = st.session_state.get("modo_barra", "Lado a lado")

        base = df_semana.copy()
        if colabs_sel:
            base = base[base["_COLAB_"].isin([str(x).upper().strip() for x in colabs_sel])].copy()

        tmp = (
            base.groupby(["DOW_NUM", "DOW", "_COLAB_"], as_index=False)["_QTD_"]
            .sum()
            .rename(columns={"_QTD_": "Notas"})
        )

        dias_df = pd.DataFrame({"DOW_NUM":[0,1,2,3,4], "DOW":["SEG","TER","QUA","QUI","SEX"]})
        col_df = pd.DataFrame({"_COLAB_": [str(x).upper().strip() for x in (colabs_sel if colabs_sel else colabs_disp)]})

        grid = dias_df.assign(_k=1).merge(col_df.assign(_k=1), on="_k").drop(columns="_k")
        tmp = grid.merge(tmp, on=["DOW_NUM","DOW","_COLAB_"], how="left").fillna({"Notas":0})
        tmp["Notas"] = tmp["Notas"].astype(int)
        tmp = tmp.sort_values(["DOW_NUM", "_COLAB_"])

        barmode = "group" if modo_barra == "Lado a lado" else "stack"

        fig_bar = px.bar(
            tmp,
            x="DOW",
            y="Notas",
            color="_COLAB_",
            barmode=barmode,
            text="Notas",
            template="plotly_white",
            category_orders={"DOW": ["SEG","TER","QUA","QUI","SEX"]},
        )
        fig_bar.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=10, b=10),
            legend_title_text="Colaborador"
        )
        fig_bar.update_traces(textposition="outside", cliponaxis=False)

        total_graf = int(tmp["Notas"].sum())
        fig_bar.add_annotation(
            xref="paper", yref="paper",
            x=0.99, y=1.12,
            text=f"<b>TOTAL SEMANAL: {fmt_int(total_graf)}</b>",
            showarrow=False,
            align="right"
        )

        st.plotly_chart(fig_bar, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

# ---- 2) DONUT
with row_main[1]:
    st.markdown('<div class="card"><div class="card-title">ACUMULADO POR COLABORADOR</div>', unsafe_allow_html=True)

    fig_donut, _total_donut = donut_colaborador_acumulado(df_filtro, int(ano_sel) if ano_sel else None)
    if fig_donut is None:
        st.info("Sem dados para o acumulado por colaborador.")
    else:
        st.plotly_chart(fig_donut, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

# ---- 3) RESUMO
with row_main[2]:
    periodo_txt = f"{week_start.strftime('%d/%m/%Y')} a {week_end.strftime('%d/%m/%Y')}"
    st.markdown(
        f"""
        <div class="card">
          <div class="card-title">RESUMO</div>
          <div class="kpi-row">
            <div class="kpi-big">{fmt_int(total_periodo)}</div>
          </div>
          <div style="margin-top:10px; font-weight:950; color:#0b2b45;">
            {periodo_txt}
          </div>
          <div style="margin-top:10px; background:rgba(255,255,255,0.35); border:1px solid rgba(10,40,70,0.22); border-radius:12px; padding:10px;">
            <div style="font-weight:900; color:#0b2b45; font-size:12px;">TOTAL NO ANO</div>
            <div style="font-weight:950; color:#9b0d0d; font-size:22px; line-height:1.0;">{fmt_int(total_ano)}</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ======================================================
# DEMANDA MANUAL
# ======================================================
def tabela_para_colaborador_manual(nome_colab: str, week_start_ts: pd.Timestamp) -> pd.DataFrame:
    semana_key = week_start_ts.strftime("%Y-%m-%d")
    key_base = f"demanda|{semana_key}|{nome_colab}"

    if "demanda_manual" not in st.session_state:
        st.session_state["demanda_manual"] = {}

    dias = []
    for i in range(5):
        d = (week_start_ts + pd.Timedelta(days=i)).date()
        dow = ["SEG", "TER", "QUA", "QUI", "SEX"][i]
        dias.append((d, dow))

    for d, _dow in dias:
        k = f"{key_base}|{d.isoformat()}"
        if k not in st.session_state["demanda_manual"]:
            st.session_state["demanda_manual"][k] = "-"

    st.markdown(
        "<div style='text-align:left; font-weight:950; color:#0b2b45; margin: 4px 0 10px 0;'>"
        "Selecione a demanda por dia</div>",
        unsafe_allow_html=True
    )

    for d, dow in dias:
        k = f"{key_base}|{d.isoformat()}"
        cA, cB = st.columns([1, 2.2], gap="small")
        with cA:
            st.markdown(
                f"<div style='font-weight:950;color:#0b2b45'>{dow} • {d.strftime('%d/%m')}</div>",
                unsafe_allow_html=True
            )
        with cB:
            opts = ["-"] + OPCOES_DEMANDA
            atual = st.session_state["demanda_manual"].get(k, "-")
            idx = opts.index(atual) if atual in opts else 0
            escolha = st.selectbox(label="", options=opts, index=idx, key=f"sb_{k}")
            st.session_state["demanda_manual"][k] = escolha

    rows = []
    for d, dow in dias:
        k = f"{key_base}|{d.isoformat()}"
        rows.append({"Data": pd.to_datetime(d), "DOW": dow, "Demanda": st.session_state["demanda_manual"][k]})

    return pd.DataFrame(rows)


# ======================================================
# 3 TABELAS
# ======================================================
st.markdown('<div class="card"><div class="card-title">DEMANDA DE APOIO (MANUAL)</div>', unsafe_allow_html=True)

if st.button("🧹 Limpar demandas desta semana"):
    semana_key = week_start.strftime("%Y-%m-%d")
    if "demanda_manual" in st.session_state:
        apagar = [k for k in list(st.session_state["demanda_manual"].keys()) if f"demanda|{semana_key}|" in k]
        for k in apagar:
            del st.session_state["demanda_manual"][k]
    st.rerun()

pessoas = sorted(normalize_colab_series(df_filtro["_COLAB_"]).dropna().unique().tolist())

top3_tbl = []
if not df_semana.empty:
    top3_tbl = (
        df_semana.dropna(subset=["_COLAB_"])
        .groupby("_COLAB_")["_QTD_"]
        .sum()
        .sort_values(ascending=False)
        .head(3)
        .index.tolist()
    )
if len(top3_tbl) < 3:
    resto = [p for p in pessoas if p not in top3_tbl]
    top3_tbl = (top3_tbl + resto)[:3]

s1, s2, s3 = st.columns(3, gap="large")
with s1:
    colab1 = st.selectbox("Tabela 1", options=pessoas, index=(pessoas.index(top3_tbl[0]) if top3_tbl and top3_tbl[0] in pessoas else 0), key="t1")
with s2:
    colab2 = st.selectbox("Tabela 2", options=pessoas, index=(pessoas.index(top3_tbl[1]) if len(top3_tbl) > 1 and top3_tbl[1] in pessoas else 0), key="t2")
with s3:
    colab3 = st.selectbox("Tabela 3", options=pessoas, index=(pessoas.index(top3_tbl[2]) if len(top3_tbl) > 2 and top3_tbl[2] in pessoas else 0), key="t3")

tcols = st.columns(3, gap="large")
cores = ["head-blue", "head-green", "head-yellow"]
selecionados = [colab1, colab2, colab3]

rendered_tables = {}
for i in range(3):
    nome = selecionados[i]
    with tcols[i]:
        df_tbl = tabela_para_colaborador_manual(nome, week_start)
        rendered_tables[nome] = df_tbl
        st.markdown(html_torpedo_table(f"DEMANDA DE APOIO – {nome}", cores[i], df_tbl), unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)


# ======================================================
# PDF
# ======================================================
def gerar_pdf_torpedo(ano_ref, periodo_txt, total_periodo, total_ano, tabelas_dict):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    elementos = []

    elementos.append(Paragraph(f"<b>TORPEDO SEMANAL – PRODUTIVIDADE ({ano_ref})</b>", styles["Title"]))
    elementos.append(Spacer(1, 10))
    elementos.append(Paragraph(f"<b>Período:</b> {periodo_txt}", styles["Normal"]))
    elementos.append(Paragraph(f"<b>Total do período:</b> {total_periodo}", styles["Normal"]))
    elementos.append(Paragraph(f"<b>Total no ano:</b> {total_ano}", styles["Normal"]))
    elementos.append(Spacer(1, 12))

    elementos.append(Paragraph("<b>Demandas (seg–sex)</b>", styles["Heading2"]))
    elementos.append(Spacer(1, 6))

    for nome, df_tbl in (tabelas_dict or {}).items():
        elementos.append(Paragraph(f"<b>{nome}</b>", styles["Heading3"]))
        data = [["Data", "Dia", "Demanda"]] + [
            [pd.to_datetime(r["Data"]).strftime("%d/%m/%Y"), r["DOW"], r["Demanda"]]
            for _, r in df_tbl.iterrows()
        ]
        t = Table(data, repeatRows=1, colWidths=[75, 35, 360])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
            ("FONT", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,-1), 8),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
        ]))
        elementos.append(t)
        elementos.append(Spacer(1, 10))

    doc.build(elementos)
    buffer.seek(0)
    return buffer

periodo_txt_pdf = f"{week_start.strftime('%d/%m/%Y')} a {week_end.strftime('%d/%m/%Y')}"
pdf_buffer = gerar_pdf_torpedo(
    ano_ref=ano_txt,
    periodo_txt=periodo_txt_pdf,
    total_periodo=total_periodo,
    total_ano=total_ano,
    tabelas_dict=rendered_tables
)

st.download_button(
    label="📄 Exportar PDF (Torpedo)",
    data=pdf_buffer,
    file_name=f"Torpedo_Semanal_{ano_txt}_{week_start.strftime('%Y%m%d')}.pdf",
    mime="application/pdf"
)


# ======================================================
# PRINT PARA PDF
# ======================================================
st.markdown("""
<style>
@media print {
  header, footer, [data-testid="stSidebar"], [data-testid="stToolbar"] { display: none !important; }
  .no-print { display: none !important; }
  .block-container { max-width: 100% !important; padding: 0 !important; }
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="no-print">', unsafe_allow_html=True)
st.button("🖨️ Exportar dashboard (print em PDF)")
st.markdown('</div>', unsafe_allow_html=True)

components.html(
    """
    <script>
      const btns = window.parent.document.querySelectorAll('button');
      const target = Array.from(btns).find(b => b.innerText.trim() === '🖨️ Exportar dashboard (print em PDF)');
      if (target && !target.dataset.printBound) {
        target.dataset.printBound = "1";
        target.addEventListener('click', () => {
          window.parent.focus();
          window.parent.print();
        });
      }
    </script>
    """,
    height=0
)
