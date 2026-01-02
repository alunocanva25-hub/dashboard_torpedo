import io
import datetime as dt
import pandas as pd
import streamlit as st
import plotly.express as px

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors


# ======================================================
# CONFIG
# ======================================================
st.set_page_config(page_title="Torpedo Produtividade Semanal", layout="wide")

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

.small{
  font-size: 12px;
  opacity: 0.85;
}

.section-title{
  font-weight: 900;
  font-size: 18px;
  margin: 0 0 6px 0;
}

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

.demotable td, .demotable th{
  border-top: 1px solid rgba(10,40,70,0.12);
  padding: 10px 10px;
  font-size: 14px;
}

.col-date{ width: 140px; font-weight: 700; }
.col-dow{ width: 80px; font-weight: 900; text-align:center; }

.head-blue{ background:#1F77B4; color:#fff; }
.head-green{ background:#2CA02C; color:#fff; }
.head-yellow{ background:#F1C40F; color:#1B1B1B; }

</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# ======================================================
# HELPERS
# ======================================================
DOW_PT = {0: "SEG", 1: "TER", 2: "QUA", 3: "QUI", 4: "SEX", 5: "SÁB", 6: "DOM"}

def parse_any_date(s: pd.Series) -> pd.Series:
    # tenta inferir datas pt-BR e ISO
    return pd.to_datetime(s, errors="coerce", dayfirst=True)

def week_monday(d: dt.date) -> dt.date:
    return d - dt.timedelta(days=d.weekday())

def safe_upper(x):
    if pd.isna(x):
        return x
    return str(x).strip().upper()

def read_table(uploaded_file: st.runtime.uploaded_file_manager.UploadedFile) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(uploaded_file)
    raise ValueError("Formato não suportado. Use CSV ou Excel.")

def build_week_days(monday: dt.date):
    return [monday + dt.timedelta(days=i) for i in range(5)]  # SEG..SEX

def html_demanda_table(title: str, color_class: str, df_rows: pd.DataFrame) -> str:
    # df_rows: Data(date), DOW(str), Demanda(str)
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
          <tbody>
            {rows_html}
          </tbody>
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
                   demandas_tables: dict) -> bytes:
    # demandas_tables: {nome: df(Data,DOW,Demanda)}
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

    # tabela produtividade
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

    for nome, df_dem in demandas_tables.items():
        story.append(Paragraph(f"<b>{nome}</b>", styles["Heading3"]))
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
# SIDEBAR
# ======================================================
st.markdown(
    """
    <div class="titlebar">
      <div style="font-size:28px;font-weight:900;color:#0B2A47;">TORPEDO PRODUTIVIDADE SEMANAL</div>
      <div class="small">Painel semanal com produtividade por colaborador + demandas de apoio</div>
    </div>
    """,
    unsafe_allow_html=True
)

with st.sidebar:
    st.header("⚙️ Configurações")

    up_prod = st.file_uploader("📄 Upload PRODUTIVIDADE (CSV/Excel)", type=["csv", "xlsx", "xls"])
    up_dem = st.file_uploader("📄 Upload DEMANDAS (CSV/Excel)", type=["csv", "xlsx", "xls"])

    st.divider()
    st.subheader("Semana de referência")
    ref_date = st.date_input("Escolha uma data da semana", value=dt.date.today())
    monday = week_monday(ref_date)
    week_days = build_week_days(monday)

    ano_acumulado = st.number_input("Ano do acumulado", value=dt.date.today().year, step=1)

    st.divider()
    st.subheader("Tabelas de Demanda")
    default_people = ["DAYVISON", "MATHEUS", "REDILENA"]
    people_tables = st.multiselect("Colaboradores nas 3 tabelas", options=default_people, default=default_people)

    st.caption("Dica: deixe 3 nomes (como na imagem).")


# ======================================================
# LOAD DATA
# ======================================================
prod = pd.DataFrame(columns=["Data", "Colaborador", "Notas"])
dem = pd.DataFrame(columns=["Data", "Colaborador", "Demanda"])

errors = []

if up_prod is not None:
    try:
        prod = read_table(up_prod).copy()
    except Exception as e:
        errors.append(f"Erro ao ler PRODUTIVIDADE: {e}")

if up_dem is not None:
    try:
        dem = read_table(up_dem).copy()
    except Exception as e:
        errors.append(f"Erro ao ler DEMANDAS: {e}")

if errors:
    for e in errors:
        st.error(e)

# Normaliza colunas
if not prod.empty:
    needed = {"Data", "Colaborador", "Notas"}
    if not needed.issubset(set(prod.columns)):
        st.warning(f"Produtividade: colunas obrigatórias faltando. Precisa ter: {sorted(list(needed))}")
    else:
        prod["Data"] = parse_any_date(prod["Data"])
        prod["Colaborador"] = prod["Colaborador"].apply(safe_upper)
        prod["Notas"] = pd.to_numeric(prod["Notas"], errors="coerce").fillna(0).astype(int)
        prod = prod.dropna(subset=["Data"])

if not dem.empty:
    needed2 = {"Data", "Colaborador", "Demanda"}
    if not needed2.issubset(set(dem.columns)):
        st.warning(f"Demandas: colunas obrigatórias faltando. Precisa ter: {sorted(list(needed2))}")
    else:
        dem["Data"] = parse_any_date(dem["Data"])
        dem["Colaborador"] = dem["Colaborador"].apply(safe_upper)
        dem["Demanda"] = dem["Demanda"].fillna("-").astype(str)
        dem = dem.dropna(subset=["Data"])


# ======================================================
# FILTERS / CALCS
# ======================================================
week_start = pd.to_datetime(monday)
week_end = pd.to_datetime(monday + dt.timedelta(days=4))  # sexta
periodo_str = f"Semana: {week_start.strftime('%d/%m/%Y')} a {week_end.strftime('%d/%m/%Y')}"

prod_semana = prod[(prod["Data"] >= week_start) & (prod["Data"] <= week_end)].copy() if not prod.empty else prod.copy()

# Cria dia da semana pt
if not prod_semana.empty:
    prod_semana["DOW_NUM"] = prod_semana["Data"].dt.weekday
    prod_semana["DOW"] = prod_semana["DOW_NUM"].map(DOW_PT)
    prod_semana = prod_semana[prod_semana["DOW_NUM"].between(0,4)]  # SEG..SEX

# Totais
total_semanal = int(prod_semana["Notas"].sum()) if not prod_semana.empty else 0

acumulado_ano = prod.copy()
if not acumulado_ano.empty:
    acumulado_ano = acumulado_ano[acumulado_ano["Data"].dt.year == int(ano_acumulado)]
total_ano = int(acumulado_ano["Notas"].sum()) if not acumulado_ano.empty else 0

# Colaboradores para gráfico: top 3 da semana (se existir), senão todos do acumulado
collabs_week = []
if not prod_semana.empty:
    collabs_week = (
        prod_semana.groupby("Colaborador")["Notas"].sum().sort_values(ascending=False).head(5).index.tolist()
    )
else:
    if not acumulado_ano.empty:
        collabs_week = acumulado_ano["Colaborador"].dropna().unique().tolist()

# Sidebar selector (somente se tiver dados)
if collabs_week:
    with st.sidebar:
        st.subheader("Gráfico de linhas")
        selected_collabs = st.multiselect("Colaboradores no gráfico", options=sorted(set(collabs_week)), default=collabs_week[:3])
else:
    selected_collabs = []


# ======================================================
# LAYOUT
# ======================================================
col_left, col_right = st.columns([1.15, 0.85], gap="large")

with col_left:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown(f'<div class="section-title">PRODUTIVIDADE SEMANAL</div><div class="small">{periodo_str}</div>', unsafe_allow_html=True)

    if prod_semana.empty or not selected_collabs:
        st.info("Envie o arquivo de produtividade e selecione colaboradores para exibir o gráfico.")
    else:
        # garante todos os dias SEG..SEX apareçam por colaborador
        base_days = pd.DataFrame({
            "DOW_NUM": [0,1,2,3,4],
            "DOW": ["SEG","TER","QUA","QUI","SEX"]
        })
        out = []
        for c in selected_collabs:
            tmp = prod_semana[prod_semana["Colaborador"] == c].groupby(["DOW_NUM","DOW"], as_index=False)["Notas"].sum()
            tmp = base_days.merge(tmp, on=["DOW_NUM","DOW"], how="left")
            tmp["Notas"] = tmp["Notas"].fillna(0).astype(int)
            tmp["Colaborador"] = c
            out.append(tmp)
        plot_df = pd.concat(out, ignore_index=True)

        fig = px.line(
            plot_df.sort_values("DOW_NUM"),
            x="DOW",
            y="Notas",
            color="Colaborador",
            markers=True
        )
        fig.update_layout(
            margin=dict(l=10,r=10,t=10,b=10),
            legend_title_text="",
            yaxis_title="Notas",
            xaxis_title=""
        )
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
        if acumulado_ano.empty:
            st.info("Envie o arquivo de produtividade para montar o acumulado.")
        else:
            pie = acumulado_ano.groupby("Colaborador", as_index=False)["Notas"].sum()
            pie = pie.sort_values("Notas", ascending=False)
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

if len(people_tables) == 0:
    st.info("Selecione ao menos 1 colaborador para exibir as tabelas.")
else:
    # monta tabela base com SEG..SEX
    base = pd.DataFrame({"Data": pd.to_datetime(week_days)})
    base["DOW"] = base["Data"].dt.weekday.map(DOW_PT)
    base = base[base["Data"].dt.weekday.between(0,4)].copy()

    cols = st.columns(3, gap="medium")
    color_classes = ["head-blue", "head-green", "head-yellow"]

    demandas_semana = dem[(dem["Data"] >= week_start) & (dem["Data"] <= week_end)].copy() if not dem.empty else dem.copy()

    rendered = {}
    for i in range(min(3, len(people_tables))):
        nome = safe_upper(people_tables[i])
        df_show = base.copy()
        if not demandas_semana.empty:
            tmp = demandas_semana[demandas_semana["Colaborador"] == nome][["Data","Demanda"]].copy()
            tmp = tmp.groupby("Data", as_index=False)["Demanda"].agg(lambda x: " | ".join([t for t in x if str(t).strip()]))
            df_show = df_show.merge(tmp, on="Data", how="left")
        df_show["Demanda"] = df_show["Demanda"].fillna("-")
        rendered[nome] = df_show

        with cols[i]:
            html = html_demanda_table(f"DEMANDA DE APOIO - {nome}", color_classes[i], df_show)
            st.markdown(html, unsafe_allow_html=True)

    # se selecionaram mais de 3, mostra extra abaixo
    if len(people_tables) > 3:
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        st.caption("Colaboradores extras selecionados (abaixo):")
        for nome in people_tables[3:]:
            nome = safe_upper(nome)
            df_show = base.copy()
            if not demandas_semana.empty:
                tmp = demandas_semana[demandas_semana["Colaborador"] == nome][["Data","Demanda"]].copy()
                tmp = tmp.groupby("Data", as_index=False)["Demanda"].agg(lambda x: " | ".join([t for t in x if str(t).strip()]))
                df_show = df_show.merge(tmp, on="Data", how="left")
            df_show["Demanda"] = df_show["Demanda"].fillna("-")
            rendered[nome] = df_show
            st.dataframe(df_show, use_container_width=True, hide_index=True)

st.markdown("</div>", unsafe_allow_html=True)


# ======================================================
# EXPORTS
# ======================================================
st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
st.markdown("<div class='card'>", unsafe_allow_html=True)
st.markdown("<div class='section-title'>EXPORTAR RELATÓRIO</div>", unsafe_allow_html=True)

# Dataframes para export
resumo_df = pd.DataFrame([{
    "Semana_inicio": week_start.strftime("%d/%m/%Y"),
    "Semana_fim": week_end.strftime("%d/%m/%Y"),
    "Ano_acumulado": int(ano_acumulado),
    "Total_semanal": total_semanal,
    "Total_ano": total_ano,
}])

prod_semana_export = prod_semana.copy()
if not prod_semana_export.empty:
    prod_semana_export = prod_semana_export[["Data","DOW","Colaborador","Notas"]].sort_values(["Data","Colaborador"])
    prod_semana_export["Data"] = prod_semana_export["Data"].dt.strftime("%d/%m/%Y")

dem_export = dem.copy()
if not dem_export.empty:
    dem_export = dem_export[(dem_export["Data"] >= week_start) & (dem_export["Data"] <= week_end)].copy()
    dem_export["Data"] = dem_export["Data"].dt.strftime("%d/%m/%Y")

acum_export = acumulado_ano.copy()
if not acum_export.empty:
    acum_export = acum_export.groupby("Colaborador", as_index=False)["Notas"].sum().sort_values("Notas", ascending=False)

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
    # PDF com tabelas (sem gráficos)
    demandas_tables = {}
    if "rendered" in locals():
        for k, dfv in rendered.items():
            demandas_tables[k] = dfv.copy()

    pdf_bytes = make_pdf_bytes(
        title="TORPEDO PRODUTIVIDADE SEMANAL",
        periodo=periodo_str,
        total_semanal=total_semanal,
        total_ano=total_ano,
        prod_semana_table=(prod_semana_export if not prod_semana_export.empty else pd.DataFrame(columns=["Data","DOW","Colaborador","Notas"])),
        demandas_tables=(demandas_tables if demandas_tables else {})
    )
    st.download_button(
        "📄 Baixar PDF (Relatório)",
        data=pdf_bytes,
        file_name=f"torpedo_semanal_{week_start.strftime('%Y%m%d')}.pdf",
        mime="application/pdf",
        use_container_width=True
    )

st.markdown("</div>", unsafe_allow_html=True)

st.caption("⚠️ Observação crítica: se o seu time preenche manualmente e esquece dias, o gráfico pode enganar (picos artificiais). O ideal é padronizar: todo dia registrar 0 quando não houver atendimento.")
