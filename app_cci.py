import datetime
import io
import os
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# DIRETÓRIOS E ARQUIVOS NUVEM & LOCAL
# ---------------------------------------------------------
PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))

CAMINHO_HISTORICO_XLS = os.path.join(PASTA_PROJETO, "historico_real.xls")
CAMINHO_HISTORICO_XLSX = os.path.join(PASTA_PROJETO, "historico_real.xlsx")

if os.path.exists(CAMINHO_HISTORICO_XLS):
    CAMINHO_HISTORICO = CAMINHO_HISTORICO_XLS
elif os.path.exists(CAMINHO_HISTORICO_XLSX):
    CAMINHO_HISTORICO = CAMINHO_HISTORICO_XLSX
else:
    CAMINHO_HISTORICO = None

CAMINHO_ENTRADA = os.path.join(PASTA_PROJETO, "cotacoes_semanais.xlsx")
CAMINHO_CUSTO_SISTEMA = os.path.join(
    PASTA_PROJETO, "custo_atual_sistema.xlsx"
)

# ---------------------------------------------------------
# FUNÇÃO PARA CALCULAR VARIAÇÃO DE REFERÊNCIA DIRETO DO HISTÓRICO REAL
# ---------------------------------------------------------
@st.cache_data(ttl=3600)
def obter_variacoes_historico_real(caminho_file):
    variacoes = {}
    if caminho_file and os.path.exists(caminho_file):
        try:
            xls = pd.ExcelFile(caminho_file)
            mapa_abas = {
                "Frango Congelado": ("Frango Congelado", "À vista R$"),
                "Boi Gordo": ("Boi ", "À vista R$"),
                "Ovo": ("Ovos", "Branco"),
                "Suíno Vivo": ("Suino Vivo", "SP"),
            }

            for comm, (aba_nome, col_p) in mapa_abas.items():
                if aba_nome in xls.sheet_names:
                    df = pd.read_excel(caminho_file, sheet_name=aba_nome, skiprows=1)
                    df.columns = [str(c).strip() for c in df.columns]
                    col_dt = next((c for c in df.columns if "data" in c.lower()), df.columns[0])
                    df[col_dt] = pd.to_datetime(df[col_dt], format="%d/%m/%Y", errors="coerce")
                    
                    df["Preco_Limpo"] = pd.to_numeric(df[col_p], errors="coerce")
                    df = df.dropna(subset=[col_dt, "Preco_Limpo"]).sort_values(by=col_dt)
                    
                    if not df.empty:
                        max_d = df[col_dt].max()
                        
                        # 1. Variação Mensal (Últimos 30d vs 30d Anteriores)
                        df_u30 = df[df[col_dt] >= (max_d - pd.Timedelta(days=30))]
                        df_p30 = df[(df[col_dt] < (max_d - pd.Timedelta(days=30))) & (df[col_dt] >= (max_d - pd.Timedelta(days=60)))]
                        
                        m_u30 = df_u30["Preco_Limpo"].mean()
                        m_p30 = df_p30["Preco_Limpo"].mean()
                        v_mes = ((m_u30 - m_p30) / m_p30 * 100) if (pd.notnull(m_p30) and m_p30 > 0) else 0.0
                        
                        # 2. Variação Semanal (Última Sexta vs Sexta Anterior)
                        df_sextas = df[df[col_dt].dt.weekday == 4]
                        if len(df_sextas) >= 2:
                            p_sex1 = df_sextas.iloc[-1]["Preco_Limpo"]
                            p_sex2 = df_sextas.iloc[-2]["Preco_Limpo"]
                            v_sem = ((p_sex1 - p_sex2) / p_sex2 * 100) if p_sex2 > 0 else 0.0
                        else:
                            v_sem = 0.0
                        
                        variacoes[comm] = {
                            "Var_%_Ref_Mercado_Semanal": v_sem,
                            "Var_%_Ref_Mercado_Mensal": v_mes
                        }
        except Exception:
            pass

    # Defaults para commodities sem histórico
    commodities_padrao = ["Frango Congelado", "Boi Gordo", "Ovo", "Suíno Vivo", "FLV Geral"]
    for c in commodities_padrao:
        if c not in variacoes:
            variacoes[c] = {"Var_%_Ref_Mercado_Semanal": 0.0, "Var_%_Ref_Mercado_Mensal": 0.0}

    df_res = pd.DataFrame.from_dict(variacoes, orient="index").reset_index()
    df_res.columns = ["Produto", "Var_%_Ref_Mercado_Semanal", "Var_%_Ref_Mercado_Mensal"]
    return df_res

# ---------------------------------------------------------
# CONFIGURAÇÃO DE TELA
# ---------------------------------------------------------
st.set_page_config(
    page_title="Ferramenta de Negociação Cencosud",
    page_icon="🥩",
    layout="wide",
)

# CABEÇALHO EXECUTIVO
col_logo1, col_logo2 = st.columns([1, 4])
with col_logo1:
    caminho_logo_local = os.path.join(PASTA_PROJETO, "logo_viva.png")
    if os.path.exists(caminho_logo_local):
        st.image(caminho_logo_local, width=150)
    else:
        st.image(
            "https://raw.githubusercontent.com/ygpassos-max/Cencosud/main/logo_viva.png",
            width=150,
        )

with col_logo2:
    st.title("🛒 Ferramenta de Negociação Cencosud")
    st.caption("Programa VIVA Perecíveis • Inteligência Comercial e Monitoramento de Custos")

st.divider()

aba_entrada, aba_matriz, aba_historico = st.tabs([
    "📝 Entrada Comprador (Semanal)",
    "📊 Matriz de Decisão & Spreads",
    "📈 Inteligência & Histórico Temporal",
])

# CATALOGO ORDENADO ALFABETICAMENTE
ESTRUTURA_PRODUTOS = {
    "Aves": sorted(["Filé de Peito", "Sobrecoxa"]),
    "Bovino": sorted(["Alcatra", "Contra Filé", "Coxão Mole", "Dianteiro"]),
    "FLV": sorted([
        "Banana",
        "Batata",
        "Cebola",
        "Maçã Gala",
        "Ovos c/20",
        "Tomate",
    ]),
    "Suíno": sorted(["Carré Suíno"]),
}

FORNECEDORES_POR_CATEGORIA = {
    "Aves": sorted(["BRF", "Genérico", "Seara", "SSA"]),
    "Bovino": sorted(["Genérico", "JBS"]),
    "FLV": sorted([
        "Distribuidor",
        "Granja Faria",
        "Importação",
        "Mantiqueira",
        "Produtor",
    ]),
    "Suíno": sorted(["BRF", "Genérico", "Seara"]),
}

REFERENCIA_COMMODITY = {
    "Filé de Peito": "Frango Congelado",
    "Sobrecoxa": "Frango Congelado",
    "Alcatra": "Boi Gordo",
    "Contra Filé": "Boi Gordo",
    "Coxão Mole": "Boi Gordo",
    "Dianteiro": "Boi Gordo",
    "Carré Suíno": "Suíno Vivo",
    "Ovos c/20": "Ovo",
    "Banana": "FLV Geral",
    "Batata": "FLV Geral",
    "Cebola": "FLV Geral",
    "Maçã Gala": "FLV Geral",
    "Tomate": "FLV Geral",
}

BANDEIRAS_CENCOSUD = sorted([
    "Bretas",
    "Gbarbosa - BA",
    "Gbarbosa - SE",
    "Giga",
    "Prezunic",
])

# ---------------------------------------------------------
# ABA 1: FORMULÁRIO DO COMPRADOR
# ---------------------------------------------------------
with aba_entrada:
    st.subheader("Alimentação Semanal de Compras")

    col_top1, col_top2 = st.columns(2)
    with col_top1:
        bandeira_sel = st.selectbox(
            "Bandeira Cencosud:", options=BANDEIRAS_CENCOSUD, key="sel_bandeira"
        )
        categoria_sel = st.selectbox(
            "Categoria do Produto:",
            options=sorted(list(ESTRUTURA_PRODUTOS.keys())),
            key="sel_categoria",
        )

    with col_top2:
        fornecedores_disponiveis = FORNECEDORES_POR_CATEGORIA.get(
            categoria_sel, sorted(["Genérico"])
        )
        fornecedor_sel = st.selectbox(
            "Fornecedor:", options=fornecedores_disponiveis, key="sel_fornecedor"
        )

        itens_disponiveis = ESTRUTURA_PRODUTOS[categoria_sel]
        produto_sel = st.selectbox(
            "Produto / Item Negociado:",
            options=itens_disponiveis,
            key="sel_produto",
        )

    with st.form(key="form_comprador_dinamico"):
        c_form1, c_form2 = st.columns(2)
        with c_form1:
            custo_pago = st.number_input(
                "Cotação (R$):",
                min_value=0.0,
                value=15.00,
                step=0.10,
                format="%.2f",
            )
        with c_form2:
            data_negocio = st.date_input(
                "Data Cotação:", value=datetime.date.today()
            )

        commodity_ref = REFERENCIA_COMMODITY.get(
            produto_sel, "FLV Geral"
        )
        st.info(
            f"💡 **Índice de Referência:** O item **{produto_sel}** será"
            f" comparado com a variação da commodity **{commodity_ref}** na"
            " Matriz de Decisão."
        )

        btn_salvar = st.form_submit_button(
            label="💾 Registrar Cotação de Compra"
        )

    if btn_salvar:
        novo_registro = {
            "Data_Compra": data_negocio.strftime("%Y-%m-%d"),
            "Bandeira": bandeira_sel,
            "Categoria": categoria_sel,
            "Produto": produto_sel,
            "Commodity_Referencia": commodity_ref,
            "Cotacao_Cencosud": custo_pago,
            "Fornecedor": fornecedor_sel,
            "Data_Captura": datetime.date.today().strftime("%Y-%m-%d"),
        }

        if os.path.exists(CAMINHO_ENTRADA):
            df_existente = pd.read_excel(CAMINHO_ENTRADA)
            df_atualizado = pd.concat(
                [df_existente, pd.DataFrame([novo_registro])], ignore_index=True
            )
        else:
            df_atualizado = pd.DataFrame([novo_registro])

        df_atualizado.to_excel(CAMINHO_ENTRADA, index=False)
        st.success(
            f"✅ Cotação de **{produto_sel}** ({fornecedor_sel}) registrada com"
            " sucesso!"
        )

    st.divider()

    if os.path.exists(CAMINHO_ENTRADA):
        df_historico_cotacoes = pd.read_excel(CAMINHO_ENTRADA)
        df_historico_cotacoes = df_historico_cotacoes.iloc[::-1].reset_index(
            drop=True
        )

        cols_view = [
            "Data_Compra",
            "Bandeira",
            "Categoria",
            "Produto",
            "Cotacao_Cencosud",
            "Fornecedor",
        ]
        cols_existentes = [
            c for c in cols_view if c in df_historico_cotacoes.columns
        ]

        with st.expander(
            "📋 Ver Cotações Registradas na Semana (Mais recentes primeiro)"
        ):
            st.dataframe(
                df_historico_cotacoes[cols_existentes].style.format({
                    "Cotacao_Cencosud": "R$ {:.2f}"
                }, na_rep="-"),
                use_container_width=True,
            )

# ---------------------------------------------------------
# ABA 2: MATRIZ DE DECISÃO (AUTO-CALCULADA DO HISTÓRICO REAL E COMPACTA)
# ---------------------------------------------------------
with aba_matriz:
    st.subheader("📊 Matriz Comercial: Custo Atual vs Cotação vs Variação de Mercado")

    f1, f2, f3 = st.columns(3)
    with f1:
        bandeira_filtro = st.selectbox(
            "🔍 Bandeira Cencosud:",
            options=["Todas as Bandeiras"] + BANDEIRAS_CENCOSUD,
        )
    with f2:
        categoria_filtro = st.selectbox(
            "🔍 Categoria:",
            options=["Todas as Categorias"] + sorted(list(ESTRUTURA_PRODUTOS.keys())),
        )
    with f3:
        todos_fornecedores = sorted(list(set([
            forn for sub in FORNECEDORES_POR_CATEGORIA.values() for forn in sub
        ])))
        fornecedor_filtro = st.selectbox(
            "🔍 Comparar Fornecedor:",
            options=["Todos os Fornecedores"] + todos_fornecedores,
        )

    # 1. Puxa as Variações Reais do Histórico da Aba 3
    df_ref_mkt = obter_variacoes_historico_real(CAMINHO_HISTORICO)

    # 2. Carrega Cotações Semanalmente
    if os.path.exists(CAMINHO_ENTRADA):
        df_cotacoes = pd.read_excel(CAMINHO_ENTRADA)
        if "Custo_Pago_Cencosud" in df_cotacoes.columns and "Cotacao_Cencosud" not in df_cotacoes.columns:
            df_cotacoes["Cotacao_Cencosud"] = df_cotacoes["Custo_Pago_Cencosud"]
    else:
        df_cotacoes = pd.DataFrame(
            columns=["Bandeira", "Categoria", "Produto", "Cotacao_Cencosud", "Fornecedor"]
        )

    # 3. Custos em Sistema
    todos_itens = sorted([
        item for sublista in ESTRUTURA_PRODUTOS.values() for item in sublista
    ])
    df_custo_sis = (
        pd.read_excel(CAMINHO_CUSTO_SISTEMA)
        if os.path.exists(CAMINHO_CUSTO_SISTEMA)
        else pd.DataFrame({
            "Bandeira": [
                b for b in BANDEIRAS_CENCOSUD for _ in range(len(todos_itens))
            ],
            "Produto": todos_itens * len(BANDEIRAS_CENCOSUD),
            "Custo_Atual_Sistema": [12.50] * (len(todos_itens) * len(BANDEIRAS_CENCOSUD)),
        })
    )

    # Filtros
    if bandeira_filtro != "Todas as Bandeiras":
        df_custo_f = df_custo_sis[df_custo_sis["Bandeira"] == bandeira_filtro].copy()
        df_cot_f = df_cotacoes[df_cotacoes["Bandeira"] == bandeira_filtro].copy()
    else:
        df_custo_f = (
            df_custo_sis.groupby("Produto", as_index=False)["Custo_Atual_Sistema"]
            .mean()
        )
        df_custo_f["Bandeira"] = "Todas"
        df_cot_f = df_cotacoes.copy()

    if categoria_filtro != "Todas as Categorias":
        itens_cat = ESTRUTURA_PRODUTOS[categoria_filtro]
        df_custo_f = df_custo_f[df_custo_f["Produto"].isin(itens_cat)]
        df_cot_f = df_cot_f[df_cot_f["Produto"].isin(itens_cat)]

    if fornecedor_filtro != "Todos os Fornecedores":
        df_cot_f = df_cot_f[df_cot_f["Fornecedor"] == fornecedor_filtro]

    if not df_cot_f.empty and "Cotacao_Cencosud" in df_cot_f.columns:
        df_cot_f_agrupado = (
            df_cot_f.groupby(["Produto", "Fornecedor"], as_index=False)[
                "Cotacao_Cencosud"
            ].mean()
        )
    else:
        df_cot_f_agrupado = pd.DataFrame(
            columns=["Produto", "Fornecedor", "Cotacao_Cencosud"]
        )

    # Cruzamento de tabelas
    df_matriz = pd.merge(df_custo_f, df_cot_f_agrupado, on="Produto", how="left")

    df_matriz["Commodity_Referencia"] = df_matriz["Produto"].map(
        lambda p: REFERENCIA_COMMODITY.get(p, "FLV Geral")
    )

    # Cruza com Variações Reais do Mercado
    df_matriz = pd.merge(
        df_matriz,
        df_ref_mkt,
        left_on="Commodity_Referencia",
        right_on="Produto",
        how="left",
        suffixes=("", "_Mkt"),
    )

    if "Cotacao_Cencosud" not in df_matriz.columns:
        df_matriz["Cotacao_Cencosud"] = None

    # Cálculos
    df_matriz["Var_%_Cotacao_vs_Custo"] = (
        (df_matriz["Cotacao_Cencosud"] - df_matriz["Custo_Atual_Sistema"])
        / df_matriz["Custo_Atual_Sistema"]
    ) * 100

    df_matriz["Spread_BRL"] = (
        df_matriz["Cotacao_Cencosud"] - df_matriz["Custo_Atual_Sistema"]
    )

    def diagnostico(row):
        if pd.isna(row["Cotacao_Cencosud"]):
            return "⚪ Sem Cotação"
        var_cot = row["Var_%_Cotacao_vs_Custo"]
        var_mkt = row.get("Var_%_Ref_Mercado_Semanal", 0)

        if pd.isna(var_mkt):
            var_mkt = 0

        if var_cot > (var_mkt + 1.5):
            return "🔴 ALERTA"
        elif var_cot < var_mkt:
            return "🟢 EXCELENTE"
        else:
            return "🟡 ALINHADO"

    df_matriz["Diagnostico_CCI"] = df_matriz.apply(diagnostico, axis=1)

    st.divider()

    # CARDS DE PERFORMANCE
    df_cotadas = df_matriz.dropna(subset=["Cotacao_Cencosud"])
    total_cotados = len(df_cotadas)
    qtd_alertas = len(
        df_matriz[df_matriz["Diagnostico_CCI"].str.contains("🔴", na=False)]
    )
    saldo_spread = df_cotadas["Spread_BRL"].sum() if not df_cotadas.empty else 0.0

    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("Itens Cotados na Rodada", f"{total_cotados} SKUs")
    mc2.metric(
        "Saldo do Spread (R$)",
        f"R$ {saldo_spread:,.2f}",
        delta=f"{saldo_spread:+.2f} BRL",
        delta_color="inverse" if saldo_spread > 0 else "normal",
    )
    mc3.metric(
        "Alertas de Negociação",
        f"{qtd_alertas} itens",
        delta="Revisar Cotação" if qtd_alertas > 0 else "OK",
        delta_color="inverse",
    )

    st.divider()

    # DICIONÁRIO E CONFIGURAÇÃO COMPACTA PARA NÃO TER BARRA DE ROLAGEM HORIZONTAL
    dicionario_colunas = {
        "Bandeira": "Bandeira",
        "Produto": "Produto",
        "Fornecedor": "Fornecedor",
        "Custo_Atual_Sistema": "Custo Atual",
        "Cotacao_Cencosud": "Cotação",
        "Spread_BRL": "Spread R$",
        "Var_%_Cotacao_vs_Custo": "Var. Cotação",
        "Commodity_Referencia": "Referência",
        "Var_%_Ref_Mercado_Semanal": "Var. Ref. Sem.",
        "Var_%_Ref_Mercado_Mensal": "Var. Ref. Mês",
        "Diagnostico_CCI": "Alerta",
    }

    cols_ordem_exata = [
        "Bandeira",
        "Produto",
        "Fornecedor",
        "Custo_Atual_Sistema",
        "Cotacao_Cencosud",
        "Spread_BRL",
        "Var_%_Cotacao_vs_Custo",
        "Commodity_Referencia",
        "Var_%_Ref_Mercado_Semanal",
        "Var_%_Ref_Mercado_Mensal",
        "Diagnostico_CCI",
    ]

    for c in cols_ordem_exata:
        if c not in df_matriz.columns:
            df_matriz[c] = None

    df_exibicao = df_matriz[cols_ordem_exata].rename(columns=dicionario_colunas)

    # Exibição compacta e 100% responsiva na largura da tela
    st.dataframe(
        df_exibicao.style.format({
            "Custo Atual": "R$ {:.2f}",
            "Cotação": "R$ {:.2f}",
            "Spread R$": "R$ {:.2f}",
            "Var. Cotação": "{:.2f}%",
            "Var. Ref. Sem.": "{:.2f}%",
            "Var. Ref. Mês": "{:.2f}%",
        }, na_rep="-"),
        use_container_width=True,
        column_config={
            "Bandeira": st.column_config.TextColumn("Bandeira", width="small"),
            "Produto": st.column_config.TextColumn("Produto", width="medium"),
            "Fornecedor": st.column_config.TextColumn("Fornecedor", width="small"),
            "Custo Atual": st.column_config.NumberColumn("Custo Atual", width="small"),
            "Cotação": st.column_config.NumberColumn("Cotação", width="small"),
            "Spread R$": st.column_config.NumberColumn("Spread R$", width="small"),
            "Var. Cotação": st.column_config.NumberColumn("Var. Cotação", width="small"),
            "Referência": st.column_config.TextColumn("Referência", width="small"),
            "Var. Ref. Sem.": st.column_config.NumberColumn("Var. Ref. Sem.", width="small"),
            "Var. Ref. Mês": st.column_config.NumberColumn("Var. Ref. Mês", width="small"),
            "Alerta": st.column_config.TextColumn("Alerta", width="small"),
        }
    )

    # EXPORTAÇÃO EXCEL
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_exibicao.to_excel(writer, index=False, sheet_name="Matriz_Negociacao")
    excel_data = output.getvalue()

    st.download_button(
        label="📥 Exportar Matriz de Decisão em Excel",
        data=excel_data,
        file_name=f"Matriz_Negociacao_Cencosud_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------------------------------------------------------
# ABA 3: HISTÓRICO TEMPORAL
# ---------------------------------------------------------
with aba_historico:
    st.subheader("📈 Análise Executiva e Tendência Histórica Mensal")

    if CAMINHO_HISTORICO and os.path.exists(CAMINHO_HISTORICO):
        try:
            xls = pd.ExcelFile(CAMINHO_HISTORICO)
            abas_disponiveis = sorted(xls.sheet_names)

            prod_selecionado = st.selectbox(
                "🔍 Selecione a Categoria de Produto:",
                options=abas_disponiveis,
            )

            df_p = pd.read_excel(
                CAMINHO_HISTORICO, sheet_name=prod_selecionado, skiprows=1
            )
            df_p.columns = [str(c).strip() for c in df_p.columns]

            col_data = next(
                (c for c in df_p.columns if "data" in c.lower()), df_p.columns[0]
            )
            df_p[col_data] = pd.to_datetime(
                df_p[col_data], format="%d/%m/%Y", errors="coerce"
            )

            aba_str = str(prod_selecionado).strip().lower()

            if "ovos" in aba_str or "ovo" in aba_str:
                col_preco = "Branco"
                st.info("ℹ️ Exibindo cotação **CIF Região Grande SP** para Ovos.")
            elif "suino" in aba_str:
                col_preco = "SP"
                st.info("ℹ️ Exibindo cotação oficial para **Suíno Vivo (SP)**.")
            elif "boi" in aba_str:
                col_preco = "À vista R$"
                st.info("ℹ️ Exibindo cotação do **Boi Gordo em Reais (R$)**.")
            else:
                col_preco = "À vista R$"
                st.info(f"ℹ️ Exibindo cotação de **{prod_selecionado} em Reais (R$)**.")

            df_p["Preco_Limpo"] = pd.to_numeric(
                df_p[col_preco], errors="coerce"
            )
            df_p = df_p.dropna(subset=[col_data, "Preco_Limpo"]).sort_values(
                by=col_data
            )

            max_d = df_p[col_data].max()

            df_ultimo_mes = df_p[df_p[col_data] >= (max_d - pd.Timedelta(days=30))]
            custo_medio_ultimo_mes = df_ultimo_mes["Preco_Limpo"].mean()

            df_mes_anterior = df_p[
                (df_p[col_data] < (max_d - pd.Timedelta(days=30)))
                & (df_p[col_data] >= (max_d - pd.Timedelta(days=60)))
            ]
            custo_medio_mes_anterior = df_mes_anterior["Preco_Limpo"].mean()

            if pd.notnull(custo_medio_mes_anterior) and custo_medio_mes_anterior > 0:
                var_perc_mensal = (
                    (custo_medio_ultimo_mes - custo_medio_mes_anterior)
                    / custo_medio_mes_anterior
                ) * 100
            else:
                var_perc_mensal = 0.0

            df_sextas = df_p[df_p[col_data].dt.weekday == 4]
            if len(df_sextas) >= 2:
                preco_ultima_sexta = df_sextas.iloc[-1]["Preco_Limpo"]
                dt_sexta = df_sextas.iloc[-1][col_data].strftime("%d/%m/%Y")
                preco_sexta_anterior = df_sextas.iloc[-2]["Preco_Limpo"]

                var_perc_semanal = (
                    (preco_ultima_sexta - preco_sexta_anterior)
                    / preco_sexta_anterior
                ) * 100
            elif not df_sextas.empty:
                preco_ultima_sexta = df_sextas.iloc[-1]["Preco_Limpo"]
                dt_sexta = df_sextas.iloc[-1][col_data].strftime("%d/%m/%Y")
                var_perc_semanal = 0.0
            else:
                preco_ultima_sexta = df_p.iloc[-1]["Preco_Limpo"]
                dt_sexta = df_p.iloc[-1][col_data].strftime("%d/%m/%Y")
                var_perc_semanal = 0.0

            st.divider()

            k1, k2 = st.columns(2)
            k1.metric(
                label="Custo Médio (Último Mês)",
                value=f"R$ {custo_medio_ultimo_mes:,.2f}",
                delta=f"{var_perc_mensal:+.2f}% vs mês anterior",
            )
            k2.metric(
                label=f"Fechamento Sexta-feira ({dt_sexta})",
                value=f"R$ {preco_ultima_sexta:,.2f}",
                delta=f"{var_perc_semanal:+.2f}% vs semana anterior",
            )

            st.divider()
            st.subheader(f"📉 Curva de Tendência Mensal - {prod_selecionado}")

            df_5y = df_p[
                df_p[col_data] >= (max_d - pd.DateOffset(years=5))
            ].copy()

            df_5y["Periodo"] = df_5y[col_data].dt.to_period("M")
            df_chart_mes = (
                df_5y.groupby("Periodo")["Preco_Limpo"].mean().reset_index()
            )
            df_chart_mes["Data_Mensal"] = df_chart_mes[
                "Periodo"
            ].dt.to_timestamp()

            df_final_chart = df_chart_mes.set_index("Data_Mensal")[
                ["Preco_Limpo"]
            ]
            df_final_chart.columns = ["Preço Médio (R$)"]

            st.line_chart(df_final_chart, use_container_width=True)

            with st.expander("📋 Ver Dados Brutos em Tabela"):
                st.dataframe(
                    df_p[[col_data, col_preco]], use_container_width=True
                )

        except Exception as e:
            st.error(f"❌ Erro ao ler o histórico: {e}")
    else:
        st.warning(
            "⚠️ Garanta que o arquivo **historico_real.xls** está no GitHub."
        )
