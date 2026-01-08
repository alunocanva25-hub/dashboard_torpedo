# ======================================================
# FILTROS (✅ Layout igual ao print)
#   - ESQUERDA: só “vazio” (não mexe no restante do layout)
#   - DIREITA (topo): Tipo de nota + Localidade lado a lado
#   - DIREITA (abaixo): Colaboradores (para o gráfico) / Visual
# ======================================================
locais = sorted([x for x in df_periodo["_LOCAL_"].dropna().unique().tolist() if str(x).strip()])
tipos  = sorted([x for x in df_periodo["_TIPO_"].dropna().unique().tolist() if str(x).strip()])

op_local_tabs = ["TOTAL"] + locais
op_tipo_tabs  = ["TOTAL"] + tipos

c_left, c_right = st.columns([2.4, 2.6], gap="large")

with c_left:
    # Mantém o “respiro” do layout do lado esquerdo, como no print
    st.markdown("<div style='height:52px'></div>", unsafe_allow_html=True)

with c_right:
    # Linha: Tipo de nota | Localidade
    r1, r2 = st.columns([1.1, 1.9], gap="medium")

    with r1:
        st.caption("Tipo de nota")
        st.markdown("<div class='seg-local'>", unsafe_allow_html=True)
        tipo_tab = st.segmented_control(
            label="",
            options=op_tipo_tabs,
            default=st.session_state.get("tipo_tab", "TOTAL"),
            key="tipo_tab",
        )
        st.markdown("</div>", unsafe_allow_html=True)

    with r2:
        st.caption("Localidade")
        st.markdown("<div class='seg-local'>", unsafe_allow_html=True)
        local_tab = st.segmented_control(
            label="",
            options=op_local_tabs,
            default=st.session_state.get("local_tab", "TOTAL"),
            key="local_tab",
        )
        st.markdown("</div>", unsafe_allow_html=True)

    # Abaixo: controles do gráfico
    st.caption("Colaboradores (para o gráfico) / Visual")
    ctrl_box = st.container()

# aplica filtros (Tipo + Localidade)
df_filtro = df_periodo.copy()
if tipo_tab and tipo_tab != "TOTAL":
    df_filtro = df_filtro[df_filtro["_TIPO_"] == str(tipo_tab).upper().strip()]
if local_tab and local_tab != "TOTAL":
    df_filtro = df_filtro[df_filtro["_LOCAL_"] == str(local_tab).upper().strip()]
