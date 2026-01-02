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
# CSS (visual + LOGIN moderno centralizado)
# ======================================================
st.markdown("""
<style>
:root{
  --bg:#6fa6d6;
  --card:#b9d3ee;
  --ink:#0b2b45;
  --ink2:#0B2A47;
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
.kpi-mini{ text-align:center; }
.kpi-mini .lbl{
  font-weight:900; color:var(--ink); font-size:12px; text-transform:uppercase;
}
.kpi-mini .val{
  font-weight:950; color:var(--red); font-size:26px; line-height: 1.0;
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

/* Segmented */
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

/* LOGIN moderno */
.login-wrap{
  min-height: 100vh;
  display:flex;
  align-items:center;
  justify-content:center;
  padding: 18px;
}
.login-shell{
  width: min(460px, 96vw);
  border-radius: 22px;
  overflow: hidden;
  border: 2px solid rgba(10,40,70,0.22);
  box-shadow: 0 16px 40px rgba(0,0,0,0.22);
  background: rgba(255,255,255,0.55);
  backdrop-filter: blur(8px);
}
.login-header{
  padding: 18px 18px 14px 18px;
  background: linear-gradient(135deg, rgba(11,43,69,0.92), rgba(31,119,180,0.92));
  color: white;
}
.login-header .h1{ font-size: 18px; font-weight: 950; letter-spacing:.3px; }
.login-header .h2{ font-size: 12px; font-weight: 800; opacity:.9; margin-top: 4px; }
.login-body{
  padding: 16px 18px 18px 18px;
}
.login-chip{
  display:inline-flex;
  align-items:center;
  gap:8px;
  padding: 8px 10px;
  border-radius: 14px;
  border: 1px solid rgba(10,40,70,0.18);
  background: rgba(255,255,255,0.45);
  color: #0b2b45;
  font-weight: 900;
  font-size: 12px;
  margin-bottom: 12px;
}
.small-muted{ font-size: 12px; font-weight: 800; color: rgba(11,43,69,0.85); }
</style>
""", unsafe_allow_html=True)

# ======================================================
# LOGIN (centralizado + moderno)
# ======================================================
def tela_login():
    st.markdown("""
    <div class="login-wrap">
      <div class="login-shell">
        <div class="login-header">
          <div class="h1">🔐 Acesso Restrito</div>
          <div class="h2">Torpedo Semanal • Produtividade</div>
        </div>
        <div class="login-body">
          <div class="login-chip">✅ Segurança via <b>st.secrets</b></div>
    """, unsafe_allow_html=True)

    usuario = st.text_input("Usuário", key="login_usuario", placeholder="Digite seu usuário")
    senha = st.text_input("Senha", type="password", key="login_senha", placeholder="Digite sua senha")

    c1, c2 = st.columns(2)
    with c1:
        entrar = st.button("Entrar", use_container_width=True)
    with c2:
        limpar = st.button("Limpar", use_container_width=True)

    if limpar:
        st.session_state["login_usuario"] = ""
        st.session_state["login_senha"] = ""
        st.rerun()

    if entrar:
        try:
            if usuario == st.secrets["auth"]["usuario"] and senha == st.secrets["auth"]["senha"]:
                st.session_state["logado"] = True
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos")
        except Exception:
            st.error("Secrets não configurado. Verifique [auth] usuario/senha no Streamlit Cloud.")

    st.markdown("""
        <div class="small-muted" style="margin-top:10px;">
          Se precisar trocar senha/usuário, ajuste em <b>Secrets</b> do Streamlit Cloud.
        </div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

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
      <div class="t2">Gráfico por colaborador + 3 tabelas (seg–sex) no padrão da referência</div>
    </div>
  </div>
  <div class="right-note">
    BASE DRIVE (XLSX)<br>
    <small>Colab = H • Notas = B • Tipo = C • Localidade = D • Data = E</small>
  </div>
</div>
""", unsafe_allow_html=True)

# ======================================================
# HELPERS (Drive XLSX/CSV)
# ======================================================
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

    # fallback CSV
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

DOW_PT = {0: "SEG", 1: "TER", 2: "QUA", 3: "QUI", 4: "SEX", 5: "SÁB", 6: "DOM"}

def monday_of_week(d: date) -> date:
    return d - timedelta(days=d.weekday())

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

def compor_demanda_do_dia(g: pd.DataFrame) -> str:
    # Sua base não tem "texto demanda", então sintetiza por TIPO + LOCALIDADE
    tipos_top = g["_TIPO_"].value_counts().head(2).index.tolist()
    locs_top = g["_LOCAL_"].value_counts().head(2).index.tolist()
    parts = []
    if tipos_top:
        parts.append("TIPO: " + ", ".join([str(x) for x in tipos_top if str(x).strip()]))
    if locs_top:
        parts.append("LOCAL: " + ", ".join([str(x) for x in locs_top if str(x).strip()]))
    return " | ".join(parts) if parts else "-"

def normalize_colab_series(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.upper().str.strip()
    s = s.replace({"": None, "NAN": None, "NONE": None, "NULL": None, "-": None})
    return s

# ======================================================
# BOTÃO ATUALIZAR BASE
# ======================================================
colA, colB = st.columns([1, 6])
with colA:
    if st.button("🔄 Atualizar base"):
        st.cache_data.clear()
        st.rerun()
with colB:
    st.caption("Use quando atualizar o arquivo no Drive (XLSX).")

# ======================================================
# CARREGAMENTO (XLSX no Drive) — LINK FIXO NO SCRIPT
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
# B=1 C=2 D=3 E=4 H=7
# ======================================================
COL_NOTAS = df.columns[1]   # B (NOTAS)
COL_TIPO  = df.columns[2]   # C (TIPO)
COL_LOCAL = df.columns[3]   # D (LOCALIDADE)
COL_DATA  = df.columns[4]   # E (DATA DA BAIXA)
COL_COLAB = df.columns[7]   # H (COLABORADORES)

df = df.copy()
df[COL_DATA] = pd.to_datetime(df[COL_DATA], errors="coerce", dayfirst=True)
df = df.dropna(subset=[COL_DATA]).copy()

df["_COLAB_"] = normalize_colab_series(df[COL_COLAB])
df["_TIPO_"]  = df[COL_TIPO].astype(str).str.upper().str.strip()
df["_LOCAL_"] = df[COL_LOCAL].astype(str).str.upper().str.strip()
df["_NOTAS_"] = pd.to_numeric(df[COL_NOTAS], errors="coerce").fillna(0).astype(int)

# ======================================================
# SELETORES (Ano • Período • Calendário • Semana ISO)
# ======================================================
anos_disponiveis = sorted(df[COL_DATA].dropna().dt.year.unique().astype(int).tolist())
ano_padrao = anos_disponiveis[-1] if anos_disponiveis else None

c_sel1, c_sel2, c_sel3, c_sel4 = st.columns([1.0, 1.3, 2.2, 2.0], gap="medium")

with c_sel1:
    ano_sel = st.selectbox(
        "Ano",
        options=anos_disponiveis if anos_disponiveis else ["—"],
        index=(len(anos_disponiveis) - 1) if anos_disponiveis else 0,
        key="ano_sel",
    )
    if ano_sel == "—":
        ano_sel = None

with c_sel2:
    modo_periodo = st.segmented_control(
        "Período",
        options=["Semanal", "Mensal"],
        default=st.session_state.get("modo_periodo", "Semanal"),
        key="modo_periodo",
    )

df_ano = df if ano_sel is None else df[df[COL_DATA].dt.year == int(ano_sel)].copy()
ano_txt = str(ano_sel) if ano_sel else "—"

if not df_ano.empty and df_ano[COL_DATA].notna().any():
    _min_d = df_ano[COL_DATA].min().date()
    _max_d = df_ano[COL_DATA].max().date()
else:
    _min_d = date.today()
    _max_d = date.today()

with c_sel3:
    data_ini, data_fim = st.date_input(
        "Filtro por calendário (início/fim)",
        value=(st.session_state.get("data_ini", _min_d), st.session_state.get("data_fim", _max_d)),
        min_value=_min_d,
        max_value=_max_d,
        key="range_calendario",
    )

with c_sel4:
    semana_sel = None
    if modo_periodo == "Semanal" and not df_ano.empty and ano_sel is not None:
        semanas_disp = sorted(df_ano[COL_DATA].dropna().dt.isocalendar().week.unique().astype(int).tolist())
        opcoes_sem = ["Todas"] + [f"S{w:02d}" for w in semanas_disp]
        semana_sel = st.selectbox("Semana (S01..S53)", opcoes_sem, index=0, key="semana_sel")

# aplica semana ISO seg(1) a sex(5)
if modo_periodo == "Semanal" and semana_sel and semana_sel != "Todas" and ano_sel is not None:
    w = int(str(semana_sel).replace("S", ""))
    try:
        data_ini = date.fromisocalendar(int(ano_sel), w, 1)
        data_fim = date.fromisocalendar(int(ano_sel), w, 5)
        st.caption(f"Semana {semana_sel}: {data_ini.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')} (seg–sex)")
    except ValueError:
        st.warning("Semana inválida para este ano (ISO). Usando o filtro por calendário.")

# aplica filtro calendário (inclusive)
df_periodo = df_ano.copy()
if not df_periodo.empty and df_periodo[COL_DATA].notna().any():
    _dini = pd.to_datetime(data_ini)
    _dfim = pd.to_datetime(data_fim) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)
    df_periodo = df_periodo[(df_periodo[COL_DATA] >= _dini) & (df_periodo[COL_DATA] <= _dfim)].copy()

# ======================================================
# FILTROS (Localidade / Tipo)
# ======================================================
locais = sorted([x for x in df_periodo["_LOCAL_"].dropna().unique().tolist() if str(x).strip() != ""])
tipos  = sorted([x for x in df_periodo["_TIPO_"].dropna().unique().tolist() if str(x).strip() != ""])

f1, f2 = st.columns([1.4, 1.4], gap="medium")
with f1:
    local_sel = st.multiselect("Localidade", options=locais, default=[])
with f2:
    tipo_sel = st.multiselect("Tipo de nota", options=tipos, default=[])

df_filtro = df_periodo.copy()
if local_sel:
    df_filtro = df_filtro[df_filtro["_LOCAL_"].isin([str(s).upper().strip() for s in local_sel])]
if tipo_sel:
    df_filtro = df_filtro[df_filtro["_TIPO_"].isin([str(s).upper().strip() for s in tipo_sel])]

# ======================================================
# RANGE seg–sex + acumulados
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
df_semana = df_semana[df_semana["DOW_NUM"].between(0, 4)]  # seg–sex

total_periodo = int(df_semana["_NOTAS_"].sum())
total_ano = int(df[df[COL_DATA].dt.year == int(ano_sel)]["_NOTAS_"].sum()) if ano_sel else int(df["_NOTAS_"].sum())

# ======================================================
# Colaboradores no gráfico (MOSTRAR TODOS)
# ======================================================
collabs_all = (
    normalize_colab_series(df_filtro["_COLAB_"])
    .dropna()
    .unique()
    .tolist()
)
collabs_all = sorted(collabs_all)

top3_semana = []
if not df_semana.empty:
    top3_semana = (
        df_semana.groupby("_COLAB_")["_NOTAS_"]
        .sum()
        .sort_values(ascending=False)
        .head(3)
        .index
        .tolist()
    )

col_g1, col_g2 = st.columns([2.4, 1.0], gap="medium")
with col_g2:
    selected_collabs = st.multiselect(
        "Colaboradores no gráfico",
        options=collabs_all,
        default=top3_semana if top3_semana else (collabs_all[:3] if len(collabs_all) >= 3 else collabs_all)
    )

# ======================================================
# Gráfico de linha (seg–sex) por colaborador
# ======================================================
def grafico_linha_colab(df_sem: pd.DataFrame, selected: list[str]):
    if df_sem.empty or not selected:
        return None

    base_days = pd.DataFrame({"DOW_NUM": [0, 1, 2, 3, 4], "DOW": ["SEG", "TER", "QUA", "QUI", "SEX"]})
    out = []
    for c in selected:
        tmp = (
            df_sem[df_sem["_COLAB_"] == c]
            .groupby(["DOW_NUM", "DOW"], as_index=False)["_NOTAS_"]
            .sum()
        )
        tmp = base_days.merge(tmp, on=["DOW_NUM", "DOW"], how="left")
        tmp["_NOTAS_"] = tmp["_NOTAS_"].fillna(0).astype(int)
        tmp["Colaborador"] = c
        tmp = tmp.rename(columns={"_NOTAS_": "Notas"})
        out.append(tmp)

    plot_df = pd.concat(out, ignore_index=True).sort_values("DOW_NUM")
    fig = px.line(plot_df, x="DOW", y="Notas", color="Colaborador", markers=True, template="plotly_white")
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=35, b=10), legend_title_text="")
    return fig

# ======================================================
# BLOCO PRINCIPAL (gráfico + KPI)
# ======================================================
with col_g1:
    st.markdown('<div class="card"><div class="card-title">PRODUTIVIDADE (SEG–SEX) — POR COLABORADOR</div>', unsafe_allow_html=True)
    fig_line = grafico_linha_colab(df_semana, selected_collabs)
    if fig_line is None:
        st.info("Sem dados no período selecionado ou nenhum colaborador selecionado.")
    else:
        st.plotly_chart(fig_line, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

with col_g2:
    periodo_txt = f"{week_start.strftime('%d/%m/%Y')} a {week_end.strftime('%d/%m/%Y')}"
    st.markdown(
        f"""
        <div class="card">
          <div class="card-title">RESUMO</div>
          <div class="kpi-row">
            <div class="kpi-big">{str(total_periodo)}</div>
            <div class="kpi-mini">
              <div class="lbl">TOTAL PERÍODO</div>
              <div class="val">{str(total_periodo)}</div>
            </div>
          </div>
          <div style="margin-top:10px; font-weight:950; color:#0b2b45;">
            {periodo_txt}
          </div>
          <div style="margin-top:10px; display:flex; justify-content:space-between; gap:8px;">
            <div style="flex:1; background:rgba(255,255,255,0.35); border:1px solid rgba(10,40,70,0.22); border-radius:12px; padding:10px;">
              <div style="font-weight:900; color:#0b2b45; font-size:12px;">TOTAL NO ANO</div>
              <div style="font-weight:950; color:#9b0d0d; font-size:22px; line-height:1.0;">{str(total_ano)}</div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ======================================================
# 3 TABELAS (seg–sex) — estilo torpedo
# ======================================================
st.markdown('<div class="card"><div class="card-title">TABELAS (3) — TORPEDO SEMANAL (SEG–SEX)</div>', unsafe_allow_html=True)

base_days = pd.DataFrame({"Data": pd.to_datetime([week_start + pd.Timedelta(days=i) for i in range(5)])})
base_days["DOW"] = base_days["Data"].dt.weekday.map(DOW_PT)

pessoas = collabs_all[:]  # todos colaboradores (sem cortar)

def tabela_para_colaborador(nome_colab: str) -> pd.DataFrame:
    df_c = df_semana[df_semana["_COLAB_"] == nome_colab].copy()
    if df_c.empty:
        out = base_days.copy()
        out["Demanda"] = "-"
        return out

    df_c["Data"] = df_c[COL_DATA].dt.normalize()
    agg = df_c.groupby("Data", as_index=False).apply(lambda g: compor_demanda_do_dia(g)).reset_index()
    agg = agg.rename(columns={0: "Demanda"})[["Data", "Demanda"]]
    out = base_days.merge(agg, on="Data", how="left")
    out["Demanda"] = out["Demanda"].fillna("-")
    return out

# default das tabelas = top 3 do período (se tiver), senão primeiros 3 da lista completa
top3_tbl = top3_semana[:] if len(top3_semana) == 3 else (pessoas[:3] if len(pessoas) >= 3 else pessoas)

s1, s2, s3 = st.columns(3, gap="large")
with s1:
    colab1 = st.selectbox("Tabela 1", options=pessoas, index=(pessoas.index(top3_tbl[0]) if top3_tbl else 0) if pessoas else 0, key="t1")
with s2:
    colab2 = st.selectbox("Tabela 2", options=pessoas, index=(pessoas.index(top3_tbl[1]) if len(top3_tbl) > 1 else 0) if pessoas else 0, key="t2")
with s3:
    colab3 = st.selectbox("Tabela 3", options=pessoas, index=(pessoas.index(top3_tbl[2]) if len(top3_tbl) > 2 else 0) if pessoas else 0, key="t3")

tcols = st.columns(3, gap="large")
cores = ["head-blue", "head-green", "head-yellow"]
selecionados = [colab1, colab2, colab3]

rendered_tables = {}
for i in range(3):
    nome = selecionados[i]
    df_tbl = tabela_para_colaborador(nome)
    rendered_tables[nome] = df_tbl
    with tcols[i]:
        st.markdown(html_torpedo_table(f"DEMANDA DE APOIO – {nome}", cores[i], df_tbl), unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)

# ======================================================
# PDF (relatório do torpedo)
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

    elementos.append(Paragraph("<b>Tabelas (seg–sex)</b>", styles["Heading2"]))
    elementos.append(Spacer(1, 6))

    for nome, df_tbl in (tabelas_dict or {}).items():
        elementos.append(Paragraph(f"<b>{nome}</b>", styles["Heading3"]))
        data = [["Data", "Dia", "Resumo"]] + [
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
