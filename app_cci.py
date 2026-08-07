import datetime
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
CAMINHO_REFERENCIA_MERCADO = os.path.join(
    PASTA_PROJETO, "referencia_mercado.xlsx"
)
CAMINHO_CUSTO_SISTEMA = os.path.join(
    PASTA_PROJETO, "custo_atual_sistema.xlsx"
)

# ---------------------------------------------------------
# CONFIGURAÇÃO DE TELA
# ---------------------------------------------------------
st.set_page_config(
    page_title="Cencosud Commodity Intelligence (CCI)",
    page_icon="🥩",
    layout="wide",
)

st.title("🥩 Cencosud Commodity Intelligence (CCI)")
st.caption("Plataforma de Inteligência Comercial e Análise de Tendências")

aba_entrada, aba_matriz, aba_historico = st.tabs([
    "📝 Entrada Comprador (Semanal)",
    "📊 Matriz de Decisão & Spreads",
    "📈 Inteligência & Histórico Temporal",
])

PRODUTOS_OFICIAIS = [
    "Frango Congelado",
    "Boi Gordo",
    "Ovo",
    "Suíno Vivo",
]

BANDEIRAS_CENCOSUD = [
    "Prezunic",
    "Bretas",
    "Gbarbosa - BA",
    "Gbarbosa - SE",
    "Giga",
]

# ---------------------------------------------------------
# ABA 1: FORMULÁRIO DO COMPRADOR
# ---------------------------------------------------------
with aba_entrada:
    st.subheader("Alimentação Semanal de Compras")
    with st.form(key="form_comprador"):
        col1, col2 = st.columns(2)
        with col1:
            bandeira_sel = st.selectbox(
                "Bandeira Cencosud:", options=BANDEIRAS_CENCOSUD
            )
            produto_sel = st.selectbox("Produto:", options=PRODUTOS_OFICIAIS)
        with col2:
            custo_pago = st.number_input(
                "Cotação / Novo Custo Pago (R$):",
                min_value=0.0,
                value=15.0,
                step=0.10,
                format="%.2f",
            )
            fornecedor = st.text_input(
                "Fornecedor:", value="Fornecedor Parceiro"
            )

        data_negocio = st.date_input(
            "Data da Compra:", value=datetime.date.today()
        )
        btn_salvar = st.form_submit_button(
            label="💾 Registrar Cotação de Compra"
        )

    if btn_salvar:
        novo_registro = {
            "Data_Compra": data_negocio.strftime("%Y-%m-%d"),
            "Bandeira": bandeira_sel,
            "Produto": produto_sel,
            "Cotacao_Cencosud": custo_pago,
            "Fornecedor": fornecedor,
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
        st.success("✅ Cotação registrada com sucesso!")

    st.divider()
    if os.path.exists(CAMINHO_ENTRADA):
        st.subheader("📋 Cotações Registradas nesta Semana")
        st.dataframe(pd.read_excel(CAMINHO_ENTRADA), use_container_width=True)

# ---------------------------------------------------------
# ABA 2: MATRIZ DE DECISÃO
# ---------------------------------------------------------
with aba_matriz:
    st.subheader("📊 Matriz Comercial: Custo Atual vs Cotação vs Variação de Mercado")

    c_filt1, c_filt2 = st.columns([1, 2])
    with c_filt1:
        bandeira_filtro = st.selectbox(
            "🔍 Filtrar por Bandeira Cencosud:",
            options=["Todas as Bandeiras"] + BANDEIRAS_CENCOSUD,
        )

    if os.path.exists(CAMINHO_REFERENCIA_MERCADO):
        df_ref_mkt = pd.read_excel(CAMINHO_REFERENCIA_MERCADO)
    else:
        df_ref_mkt = pd.DataFrame({
            "Produto": PRODUTOS_OFICIAIS,
            "Var_%_Ref_Mercado_Semanal": [1.50, -0.80, 2.10, 0.00],
            "Var_%_Ref_Mercado_Mensal": [3.20, -1.50, 4.00, 1.20],
        })

    if os.path.exists(CAMINHO_ENTRADA):
        df_cotacoes = pd.read_excel(CAMINHO_ENTRADA)
        if "Custo_Pago_Cencosud" in df_cotacoes.columns and "Cotacao_Cencosud" not in df_cotacoes.columns:
            df_cotacoes["Cotacao_Cencosud"] = df_cotacoes["Custo_Pago_Cencosud"]
    else:
        df_cotacoes = pd.DataFrame(columns=["Bandeira", "Produto", "Cotacao_Cencosud"])

    df_custo_sis = (
        pd.read_excel(CAMINHO_CUSTO_SISTEMA)
        if os.path.exists(CAMINHO_CUSTO_SISTEMA)
        else pd.DataFrame({
            "Bandeira": [b for b in BANDEIRAS_CENCOSUD for _ in PRODUTOS_OFICIAIS],
            "Produto": PRODUTOS_OFICIAIS * len(BANDEIRAS_CENCOSUD),
            "Custo_Atual_Sistema": [
                7.20, 225.00, 110.00, 7.00,
                7.10, 224.00, 109.00, 6.95,
                7.30, 226.00, 111.00, 7.05,
                7.25, 225.50, 110.50, 7.00,
                7.15, 223.50, 108.50, 6.90,
            ],
        })
    )

    if bandeira_filtro != "Todas as Bandeiras":
        df_custo_f = df_custo_sis[df_custo_sis["Bandeira"] == bandeira_filtro]
        df_cot_f = df_cotacoes[df_cotacoes["Bandeira"] == bandeira_filtro]
    else:
        df_custo_f = df_custo_sis.groupby("Produto", as_index=False)["Custo_Atual_Sistema"].mean()
        if not df_cotacoes.empty and "Cotacao_Cencosud" in df_cotacoes.columns:
            df_cot_f = df_cotacoes.groupby("Produto", as_index=False)["Cotacao_Cencosud"].mean()
        else:
            df_cot_f = pd.DataFrame(columns=["Produto", "Cotacao_Cencosud"])

    df_matriz = pd.merge(df_custo_f, df_cot_f, on="Produto", how="left")
    df_matriz = pd.merge(df_matriz, df_ref_mkt, on="Produto", how="left")

    if "Cotacao_Cencosud" not in df_matriz.columns:
        df_matriz["Cotacao_Cencosud"] = None

    df_matriz["Var_%_Cotacao_vs_Custo"] = (
        (df_matriz["Cotacao_Cencosud"] - df_matriz["Custo_Atual_Sistema"])
        / df_matriz["Custo_Atual_Sistema"]
    ) * 100

    df_matriz["Spread_BRL"] = (
        df_matriz["Cotacao_Cencosud"] - df_matriz["Custo_Atual_Sistema"]
    )

    def diagnostico(row):
        if pd.isna(row["Cotacao_Cencosud"]):
            return "⚪ Sem Cotação na Semana"
        var_cot = row["Var_%_Cotacao_vs_Custo"]
        var_mkt = row.get("Var_%_Ref_Mercado_Semanal", 0)
        
        if pd.isna(var_mkt):
            var_mkt = 0

        if var_cot > (var_mkt + 1.5):
            return "🔴 ALERTA: Alta acima do mercado"
        elif var_cot < var_mkt:
            return "🟢 EXCELENTE: Preço abaixo do mercado"
        else:
            return "🟡 DENTRO DA META: Alinhado ao mercado"

    df_matriz["Diagnostico_CCI"] = df_matriz.apply(diagnostico, axis=1)

    st.divider()

    cols_exibir = [
        "Produto",
        "Custo_Atual_Sistema",
        "Cotacao_Cencosud",
        "Spread_BRL",
        "Var_%_Cotacao_vs_Custo",
        "Var_%_Ref_Mercado_Semanal",
        "Var_%_Ref_Mercado_Mensal",
        "Diagnostico_CCI",
    ]

    if "Bandeira" in df_matriz.columns:
        cols_exibir.insert(0, "Bandeira")

    cols_existentes = [c for c in cols_exibir if c in df_matriz.columns]

    st.dataframe(
        df_matriz[cols_existentes].style.format({
            "Custo_Atual_Sistema": "R$ {:.2f}",
            "Cotacao_Cencosud": "R$ {:.2f}",
            "Spread_BRL": "R$ {:.2f}",
            "Var_%_Cotacao_vs_Custo": "{:.2f}%",
            "Var_%_Ref_Mercado_Semanal": "{:.2f}%",
            "Var_%_Ref_Mercado_Mensal": "{:.2f}%",
        }, na_rep="-"),
        use_container_width=True,
    )

    st.divider()

    with st.expander("⚙️ Painel do Gestor: Atualizar Variação da Referência de Mercado"):
        df_editor_ref = st.data_editor(
            df_ref_mkt, num_rows="fixed", key="editor_matriz_ref"
        )
        if st.button("💾 Salvar Variações de Mercado"):
            df_editor_ref.to_excel(CAMINHO_REFERENCIA_MERCADO, index=False)
            st.success("✅ Variações de Mercado salvas!")

# ---------------------------------------------------------
# ABA 3: HISTÓRICO TEMPORAL COM VARIAÇÕES MENSAL E SEMANAL
# ---------------------------------------------------------
with aba_historico:
    st.subheader("📈 Análise Executiva e Tendência Histórica Mensal")

    if CAMINHO_HISTORICO and os.path.exists(CAMINHO_HISTORICO):
        try:
            xls = pd.ExcelFile(CAMINHO_HISTORICO)
            abas_disponiveis = xls.sheet_names

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

            # ---------------------------------------------------------
            # CÁLCULOS DOS KPIS COM VARIAÇÃO % (MENSAL E SEMANAL)
            # ---------------------------------------------------------
            max_d = df_p[col_data].max()

            # 1. Custo Médio Último Mês vs Mês Anterior
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

            # 2. Fechamento Última Sexta vs Sexta Anterior
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

            # EXIBIÇÃO EM 2 CARDS EXECUTIVOS COM DELTAS
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
