import io
import re
import datetime as dt
import pandas as pd
import streamlit as st
import plotly.express as px
import requests

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# ======================================================
# CONFIG
# ======================================================
st.set_page_config(page_title="Torpedo Semanal", layout="wide")

# ======================================================
# CSS (Dashboard + Login)
# ======================================================
CSS = """
<style>
.stApp { background: #EAF2FB; }
.block-container{ padding-top: 0.8rem; max-width: 1500px; }

.titlebar{
  border-radius: 18px;
  padding: 14px 18px;
  background: #DDEBFA;
  border: 2px solid rgba(10,40,70,0.18);
  box-shadow: 0 10px 25px rgba(10,40,70,0.08);
  margin-bottom: 14px;
}

.card{
  border-radius: 18px;
  padding: 14px 16px;
  background: #FFFFFF;
  border: 2px solid rgba(10,40,70,0.12);
  box-shadow: 0 10px 25px rgba(10,40,70,0.08);
}

.badge{
  display:inline-block;
  padding: 10px 14px;
  border-radius: 16px;
  background: #0B2A47;
  color: #fff;
  font-weight: 800;
  text-align:center;
  min-width: 140px;
}

.small{ font-size: 12px; opacity: 0.85; }
.section-title{ font-weight: 900; font-size: 18px; margin: 0 0 6px 0; }

.tablewrap{
  border-radius: 16px;
  overflow: hidden;
  border: 2px solid rgba(10,40,70,0.12);
}
.tblhead{
  padding: 10px 12px;
  font-weight: 900;
  color: #0B2A47;
}
.demotable{
  width: 100%;
  border-collapse: collapse;
  background: #fff;
}
.demotable td{
  border-top: 1px solid rgba(10,40,70,0.12);
  padding: 10px 10px;
  font-size: 14px;
}
.col-date{ width: 140px; font-weight: 700; }
.col-dow{ width: 80px; font-weight: 900; text-align:center; }

.head-blue{ background:#1F77B4; color:#fff; }
.head-green{ background:#2CA02C; color:#fff; }
.head-yellow{ background:#F1C40F; color:#1B1B1B; }

/* LOGIN */
.login-wrap{
  min-height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
}
.login-card{
  width: min(430px, 92vw);
  background: #FFFFFF;
  border: 2px solid rgba(10,40,70,0.12);
  border-radius: 18px;
  box-shadow: 0 10px 25px rgba(10,40,70,0.12);
  padding: 22px 20px;
}
.login-title{
  font-size: 22px;
  font-weight: 900;
  color: #0B2A47;
  text-align: center;
  margin-bottom: 6px;
}
.login-sub{
  font-size: 13px;
  opacity: .8;
  text-align: center;
  margin-bottom: 14px;
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

DOW_PT = {0: "SEG", 1: "TER", 2: "QUA", 3: "QUI", 4: "SEX", 5: "SÁB", 6: "DOM"}

# ======================================================
# LOGIN (IMPLANTADO + CENTRALIZADO)
# ======================================================
def tela_login():
    st.markdown("""
    <div class="login-wrap">
      <div class="login-card">
        <div class="login-title">🔐 Acesso Restrito</div>
        <div class="login-sub">Informe suas credenciais para continuar</div>
    """, unsafe_allow_html=True)

    usuario = st.text_input("Usuário", key="login_usuario")
    senha = st.text_input("Senha", type="password", key="login_senha")

    col1, col2 = st.columns(2)
    with col1:
        entrar = st.button("Entrar", use_container_width=True)
    with col2:
        limpar = st.button("Limpar", use_container_width=True)

    if limpar:
        st.session_state["login_usuario"] = ""
        st.session_state["login_senha"] = ""
        st.rerun()

    if entrar:
        if usuario == st.secrets["auth"]["usuario"] and senha == st.secrets["auth"]["senha"]:
            st.session_state["logado"] = True
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos")

    st.markdown("</div></div>", unsafe_allow_html=True)

if "logado" not in st.session_state:
    st.session_state["logado"] = False
if not st.session_state["logado"]:
    tela_login()
    st.stop()

# ======================================================
# FUNÇÕES (PADRÃO DO OUTRO PROJETO)
# ======================================================
def normalizar_nome_col(c: str) -> str:
    c = str(c).strip().upper()
    c = re.sub(r"\s+", " ", c)
    return c

def achar_coluna(df: pd.DataFrame, candidatos: list[str]) -> str:
    """
    Encontra a primeira coluna que bata com qualquer item em candidatos,
    comparando por 'contém' (tolerante) no nome normalizado.
    """
    cols_norm = {normalizar_nome_col(c): c for c in df.columns}
    # match exato primeiro
    for cand in candidatos:
        cand_n = normalizar_nome_col(cand)
        if cand_n in cols_norm:
            return cols_norm[cand_n]
    # match por "contém"
    for cand in candidatos:
        cand_n = normalizar_nome_col(cand)
        for cnorm, coriginal in cols_norm.items():
            if cand_n in cnorm:
                return coriginal
    raise ValueError(f"Não encontrei coluna. Tentei: {candidatos}")

def validar_estrutura(df: pd.DataFrame):
    if df is None or df.empty:
        raise ValueError("A base carregou vazia. Verifique o link/arquivo.")
    if len(df.columns) < 2:
        raise ValueError("A base tem poucas colunas. Verifique se o XLSX está correto.")

@st.cache_data(show_spinner=False, ttl=300)
def carregar_base(url_drive_uc: str, sheet_name: str | int | None = 0) -> pd.DataFrame:
    """
    Baixa XLSX do Drive via link 'uc?id=' e lê com pandas.read_excel().
    """
    try:
        r = requests.get(url_drive_uc, timeout=35)
        r.raise_for_status()
    except Exception as e:
        raise RuntimeError(f"Falha ao baixar o XLSX do Drive: {e}")

    content = io.BytesIO(r.content)
    try:
        df = pd.read_excel(content, sheet_name=sheet_name, engine="openpyxl")
    except Exception as e:
        raise RuntimeError(f"Falha ao ler XLSX (aba={sheet_name}): {e}")

    return df

def week_monday(d: dt.date) -> dt.date:
    return d - dt.timedelta(days=d.weekday())

def safe_upper(x):
    if pd.isna(x):
        return x
    return str(x).strip().upper()

def html_demanda_table(title: str, color_class: str, df_rows: pd.DataFrame) -> str:
    rows_html = ""
    for _, r in df_rows.iterrows():
        dstr = pd.to_datetime(r["Data"]).strftime("%d/%m/%Y")
        rows_html += f"""
          <tr>
            <td class="col-date">{dstr}</td>
            <td class="col-dow">{r["DOW"]}</td>
            <td>{r["Demanda"]}</td>
          </tr>
        """
    return f"""
      <div class="tablewrap">
        <div class="tblhead {color_class}">{title}</div>
        <table class="demotable">
          <tbody>{rows_html}</tbody>
        </table>
      </div>
    """

def make_excel_bytes(resumo: pd.DataFrame,
                     prod_semana: pd.DataFrame,
                     demandas_semana: pd.DataFrame,
                     acumulado_ano: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumo.to_excel(writer, index=False, sheet_name="RESUMO")
        prod_semana.to_excel(writer, index=False, sheet_name="PROD_SEMANA")
        demandas_semana.to_excel(writer, index=False, sheet_name="DEMANDAS_SEMANA")
        acumulado_ano.to_excel(writer, index=False, sheet_name="ACUMULADO_ANO")
    return output.getvalue()

def make_pdf_bytes(title: str,
                   periodo: str,
                   total_semanal: int,
                   total_ano: int,
                   prod_semana_table: pd.DataFrame,
                   demandas_tables: dict[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=28, leftMargin=28, topMargin=28, bottomMargin=28)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    story.append(Paragraph(periodo, styles["Normal"]))
    story.append(Spacer(1, 10))
    story.append(Paragraph(f"<b>Total semanal:</b> {total_semanal}", styles["Normal"]))
    story.append(Paragraph(f"<b>Total no ano:</b> {total_ano}", styles["Normal"]))
    story.append(Spacer(1, 14))

    story.append(Paragraph("<b>Produtividade da semana (SEG a SEX)</b>", styles["Heading2"]))
    story.append(Spacer(1, 6))

    if prod_semana_table is None or prod_semana_table.empty:
        prod_semana_table = pd.DataFrame(columns=["Data", "DOW", "Colaborador", "Notas"])

    tdata = [list(prod_semana_table.columns)] + prod_semana_table.values.tolist()
    t = Table(tdata, hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0B2A47")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#AAB7C4")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F3F7FB")]),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
    ]))
    story.append(t)
    story.append(Spacer(1, 14))

    story.append(Paragraph("<b>Demandas de apoio</b>", styles["Heading2"]))
    story.append(Spacer(1, 6))

    for nome, df_dem in (demandas_tables or {}).items():
        story.append(Paragraph(f"<b>{nome}</b>", styles["Heading3"]))

        if df_dem is None or df_dem.empty:
            df_dem = pd.DataFrame(columns=["Data", "DOW", "Demanda"])

        tdata2 = [["Data", "Dia", "Demanda"]] + [
            [pd.to_datetime(r["Data"]).strftime("%d/%m/%Y"), r["DOW"], r["Demanda"]]
            for _, r in df_dem.iterrows()
        ]
        t2 = Table(tdata2, hAlign="LEFT", colWidths=[80, 40, 360])
        t2.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F77B4")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#AAB7C4")),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F3F7FB")]),
            ("FONTSIZE", (0,0), (-1,-1), 9),
        ]))
        story.append(t2)
        story.append(Spacer(1, 10))

    doc.build(story)
    return buffer.getvalue()

# ======================================================
# DEFAULTS (SECRETS)
# ======================================================
URL_BASE_DEFAULT = ""
DEFAULT_YEAR = dt.date.today().year
DEFAULT_SHEET = 0

if "torpedo" in st.secrets:
    URL_BASE_DEFAULT = st.secrets["torpedo"].get("url_base", URL_BASE_DEFAULT)
    DEFAULT_YEAR = int(st.secrets["torpedo"].get("default_year", DEFAULT_YEAR))
    DEFAULT_SHEET = st.secrets["torpedo"].get("sheet_name", DEFAULT_SHEET)

# ======================================================
# HEADER
# ======================================================
st.markdown(
    """
    <div class="titlebar">
      <div style="font-size:28px;font-weight:900;color:#0B2A47;">TORPEDO SEMANAL</div>
      <div class="small">Base XLSX no Drive (igual ao outro projeto) + Login + Export PDF/Excel</div>
    </div>
    """,
    unsafe_allow_html=True
)

# ======================================================
# SIDEBAR
# ======================================================
with st.sidebar:
    st.header("⚙️ Configurações")

    if st.button("🚪 Sair"):
        st.session_state["logado"] = False
        st.rerun()

    st.divider()
    st.subheader("📦 Base (XLSX no Drive)")
    URL_BASE = st.text_input("URL (uc?id=...)", value=URL_BASE_DEFAULT)
    sheet_name = st.text_input("Aba (nome ou índice)", value=str(DEFAULT_SHEET))

    load_btn = st.button("🔄 Carregar Base", use_container_width=True)

    st.divider()
    st.subheader("📅 Semana")
    ref_date = st.date_input("Escolha uma data da semana", value=dt.date.today())
    monday = week_monday(ref_date)
    week_start = pd.to_datetime(monday)
    week_end = pd.to_datetime(monday + dt.timedelta(days=4))
    periodo_str = f"Semana: {week_start.strftime('%d/%m/%Y')} a {week_end.strftime('%d/%m/%Y')}"

    ano_acumulado = st.number_input("Ano do acumulado", value=DEFAULT_YEAR, step=1)

    st.divider()
    st.subheader("👥 Tabelas (3 colaboradores)")
    default_people = ["DAYVISON", "MATHEUS", "REDILENA"]
    people_tables = st.multiselect("Escolha 3", options=default_people, default=default_people)

# ======================================================
# CARREGAMENTO BASE
# ======================================================
df = None
if load_btn:
    try:
        if not URL_BASE.strip():
            raise ValueError("Informe a URL do Drive (uc?id=...).")

        # sheet_name pode ser número (índice) ou string (nome)
        sh = sheet_name.strip()
        sh_arg = int(sh) if sh.isdigit() else sh

        df = carregar_base(URL_BASE.strip(), sheet_name=sh_arg)
        validar_estrutura(df)
        st.success("Base carregada com sucesso.")
    except Exception as e:
        st.error(str(e))

# se não clicou carregar ainda, tenta manter em cache via session
if "df_base" not in st.session_state:
    st.session_state["df_base"] = None
if df is not None:
    st.session_state["df_base"] = df

df = st.session_state["df_base"]

if df is None:
    st.info("Carregue a base no menu lateral para iniciar.")
    st.stop()

# ======================================================
# MAPEAR COLUNAS (Torpedo)
# ======================================================
# Mantive o padrão de achar_coluna. Ajuste os candidatos conforme sua base real.
COL_DATA = achar_coluna(df, ["DATA", "DT", "DIA", "DATA ATENDIMENTO", "DT ATEND"])
COL_COLAB = achar_coluna(df, ["COLABORADOR", "USUARIO", "RESPONSAVEL", "NOME", "COLAB"])
COL_NOTAS = achar_coluna(df, ["NOTAS", "QTD", "QTDE", "QUANTIDADE", "TOTAL", "ATENDIDAS", "NOTAS ATENDIDAS", "NOTA"])
# DEMANDA pode não existir; se não existir, o dashboard ainda roda
COL_DEMANDA = None
try:
    COL_DEMANDA = achar_coluna(df, ["DEMANDA", "APOIO", "OBS", "OBSERVACAO", "DESCRICAO", "MOTIVO"])
except Exception:
    COL_DEMANDA = None

# Normaliza
df[COL_DATA] = pd.to_datetime(df[COL_DATA], errors="coerce", dayfirst=True)
df = df.dropna(subset=[COL_DATA]).copy()
df["_COLAB_"] = df[COL_COLAB].astype(str).str.upper().str.strip()
df["_NOTAS_"] = pd.to_numeric(df[COL_NOTAS], errors="coerce").fillna(0).astype(int)

if COL_DEMANDA:
    df["_DEMANDA_"] = df[COL_DEMANDA].fillna("-").astype(str)
else:
    df["_DEMANDA_"] = "-"

# ======================================================
# FILTROS / CÁLCULOS
# ======================================================
df_semana = df[(df[COL_DATA] >= week_start) & (df[COL_DATA] <= week_end)].copy()
df_semana["DOW_NUM"] = df_semana[COL_DATA].dt.weekday
df_semana["DOW"] = df_semana["DOW_NUM"].map(DOW_PT)
df_semana = df_semana[df_semana["DOW_NUM"].between(0, 4)]

total_semanal = int(df_semana["_NOTAS_"].sum())

df_ano = df[df[COL_DATA].dt.year == int(ano_acumulado)].copy()
total_ano = int(df_ano["_NOTAS_"].sum())

# colaborador no gráfico: top 6 da semana
collabs = []
if not df_semana.empty:
    collabs = df_semana.groupby("_COLAB_")["_NOTAS_"].sum().sort_values(ascending=False).head(6).index.tolist()
elif not df_ano.empty:
    collabs = df_ano["_COLAB_"].dropna().unique().tolist()

with st.sidebar:
    st.divider()
    st.subheader("📈 Gráfico")
    selected_collabs = st.multiselect(
        "Colaboradores no gráfico",
        options=sorted(set(collabs)) if collabs else [],
        default=collabs[:3] if collabs else []
    )

# ======================================================
# DASHBOARD (Layout)
# ======================================================
col_left, col_right = st.columns([1.15, 0.85], gap="large")

with col_left:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown(f'<div class="section-title">PRODUTIVIDADE SEMANAL</div><div class="small">{periodo_str}</div>', unsafe_allow_html=True)

    if df_semana.empty or not selected_collabs:
        st.info("Sem dados da semana ou nenhum colaborador selecionado.")
    else:
        base_days = pd.DataFrame({"DOW_NUM": [0,1,2,3,4], "DOW": ["SEG","TER","QUA","QUI","SEX"]})
        out = []
        for c in selected_collabs:
            tmp = df_semana[df_semana["_COLAB_"] == c].groupby(["DOW_NUM","DOW"], as_index=False)["_NOTAS_"].sum()
            tmp = base_days.merge(tmp, on=["DOW_NUM","DOW"], how="left")
            tmp["_NOTAS_"] = tmp["_NOTAS_"].fillna(0).astype(int)
            tmp["Colaborador"] = c
            tmp.rename(columns={"_NOTAS_": "Notas"}, inplace=True)
            out.append(tmp)

        plot_df = pd.concat(out, ignore_index=True).sort_values("DOW_NUM")

        fig = px.line(plot_df, x="DOW", y="Notas", color="Colaborador", markers=True)
        fig.update_layout(margin=dict(l=10,r=10,t=10,b=10), legend_title_text="", yaxis_title="Notas", xaxis_title="")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

with col_right:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown(f'<div class="section-title">ACUMULADO DE NOTAS ATENDIDAS POR COLABORADOR - {int(ano_acumulado)}</div>', unsafe_allow_html=True)

    top_row = st.columns([0.52, 0.48], gap="medium")
    with top_row[0]:
        st.markdown(f'<div class="badge">{total_semanal}<div class="small">TOTAL SEMANAL</div></div>', unsafe_allow_html=True)
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        st.markdown(f'<div class="badge">{total_ano}<div class="small">TOTAL</div></div>', unsafe_allow_html=True)

    with top_row[1]:
        if df_ano.empty:
            st.info("Sem dados no ano selecionado.")
        else:
            pie = df_ano.groupby("_COLAB_", as_index=False)["_NOTAS_"].sum().sort_values("_NOTAS_", ascending=False)
            pie.rename(columns={"_COLAB_": "Colaborador", "_NOTAS_": "Notas"}, inplace=True)
            fig2 = px.pie(pie, values="Notas", names="Colaborador", hole=0.60)
            fig2.update_layout(margin=dict(l=10,r=10,t=10,b=10), legend_title_text="")
            st.plotly_chart(fig2, use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ======================================================
# DEMANDAS (3 tabelas)
# ======================================================
st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
st.markdown("<div class='card'>", unsafe_allow_html=True)
st.markdown(f"<div class='section-title'>DEMANDAS DE APOIO (SEMANA)</div><div class='small'>{periodo_str}</div>", unsafe_allow_html=True)

base = pd.DataFrame({"Data": pd.to_datetime([week_start + pd.Timedelta(days=i) for i in range(5)])})
base["DOW"] = base["Data"].dt.weekday.map(DOW_PT)

df_dem_semana = df[(df[COL_DATA] >= week_start) & (df[COL_DATA] <= week_end)].copy()
df_dem_semana["Data"] = df_dem_semana[COL_DATA].dt.normalize()

cols = st.columns(3, gap="medium")
color_classes = ["head-blue", "head-green", "head-yellow"]
rendered = {}

if len(people_tables) == 0:
    st.info("Selecione os colaboradores das tabelas.")
else:
    for i in range(min(3, len(people_tables))):
        nome = safe_upper(people_tables[i])
        df_show = base.copy()

        tmp = df_dem_semana[df_dem_semana["_COLAB_"] == nome][["Data","_DEMANDA_"]].copy()
        if not tmp.empty:
            tmp = tmp.groupby("Data", as_index=False)["_DEMANDA_"].agg(
                lambda x: " | ".join([t for t in x if str(t).strip() and str(t).strip() != "-"])
            )
            tmp.rename(columns={"_DEMANDA_": "Demanda"}, inplace=True)
            df_show = df_show.merge(tmp, on="Data", how="left")

        df_show["Demanda"] = df_show["Demanda"].fillna("-")
        rendered[nome] = df_show

        with cols[i]:
            st.markdown(
                html_demanda_table(f"DEMANDA DE APOIO - {nome}", color_classes[i], df_show),
                unsafe_allow_html=True
            )

st.markdown("</div>", unsafe_allow_html=True)

# ======================================================
# EXPORTS
# ======================================================
st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
st.markdown("<div class='card'>", unsafe_allow_html=True)
st.markdown("<div class='section-title'>EXPORTAR RELATÓRIO</div>", unsafe_allow_html=True)

resumo_df = pd.DataFrame([{
    "Semana_inicio": week_start.strftime("%d/%m/%Y"),
    "Semana_fim": week_end.strftime("%d/%m/%Y"),
    "Ano_acumulado": int(ano_acumulado),
    "Total_semanal": total_semanal,
    "Total_ano": total_ano,
    "URL_BASE": URL_BASE
}])

prod_semana_export = df_semana.copy()
if not prod_semana_export.empty:
    prod_semana_export = prod_semana_export[[COL_DATA, "DOW", "_COLAB_", "_NOTAS_"]].copy()
    prod_semana_export.rename(columns={COL_DATA: "Data", "_COLAB_": "Colaborador", "_NOTAS_": "Notas"}, inplace=True)
    prod_semana_export["Data"] = pd.to_datetime(prod_semana_export["Data"]).dt.strftime("%d/%m/%Y")

dem_export = df_dem_semana.copy()
if not dem_export.empty:
    dem_export = dem_export[[COL_DATA, "_COLAB_", "_DEMANDA_"]].copy()
    dem_export.rename(columns={COL_DATA: "Data", "_COLAB_": "Colaborador", "_DEMANDA_": "Demanda"}, inplace=True)
    dem_export["Data"] = pd.to_datetime(dem_export["Data"]).dt.strftime("%d/%m/%Y")

acum_export = df_ano.groupby("_COLAB_", as_index=False)["_NOTAS_"].sum().sort_values("_NOTAS_", ascending=False)
acum_export.rename(columns={"_COLAB_": "Colaborador", "_NOTAS_": "Notas"}, inplace=True)

c1, c2 = st.columns(2, gap="medium")

with c1:
    xlsx_bytes = make_excel_bytes(
        resumo=resumo_df,
        prod_semana=prod_semana_export if not prod_semana_export.empty else pd.DataFrame(),
        demandas_semana=dem_export if not dem_export.empty else pd.DataFrame(),
        acumulado_ano=acum_export if not acum_export.empty else pd.DataFrame()
    )
    st.download_button(
        "📗 Baixar Excel (Relatório)",
        data=xlsx_bytes,
        file_name=f"torpedo_semanal_{week_start.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with c2:
    pdf_bytes = make_pdf_bytes(
        title="TORPEDO SEMANAL",
        periodo=periodo_str,
        total_semanal=total_semanal,
        total_ano=total_ano,
        prod_semana_table=prod_semana_export if not prod_semana_export.empty else pd.DataFrame(),
        demandas_tables=rendered if rendered else {}
    )
    st.download_button(
        "📄 Baixar PDF (Relatório)",
        data=pdf_bytes,
        file_name=f"torpedo_semanal_{week_start.strftime('%Y%m%d')}.pdf",
        mime="application/pdf",
        use_container_width=True
    )

st.markdown("</div>", unsafe_allow_html=True)

st.caption("Observação: se sua base não tiver a coluna de DEMANDA/APOIO/OBS, as tabelas vão ficar com '-'.")
