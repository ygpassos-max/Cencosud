import datetime
import os
import pandas as pd
import streamlit as st

PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))

CAMINHO_HISTORICO_XLSX = os.path.join(PASTA_PROJETO, "historico_real.xlsx")
CAMINHO_HISTORICO_XLS = os.path.join(PASTA_PROJETO, "historico_real.xls")

if os.path.exists(CAMINHO_HISTORICO_XLSX):
    CAMINHO_HISTORICO = CAMINHO_HISTORICO_XLSX
elif os.path.exists(CAMINHO_HISTORICO_XLS):
    CAMINHO_HISTORICO = CAMINHO_HISTORICO_XLS
else:
    CAMINHO_HISTORICO = None

CAMINHO_ENTRADA = os.path.join(PASTA_PROJETO, "cotacoes_semanais.xlsx")
CAMINHO_REFERENCIA_MERCADO = os.path.join(
    PASTA_PROJETO, "referencia_mercado.xlsx"
)

st.set_page_config(
    page_title="Cencosud Commodity Intelligence (CCI)",
    page_icon="🥩",
    layout="wide",
)

st.title("🥩 Cencosud Commodity Intelligence (CCI)")
st.caption("Plataforma de Inteligência Comercial e Análise de Tendências")

aba_entrada, aba_analise, aba_historico = st.tabs([
    "📝 Entrada Comprador (Semanal)",
    "📊 Matriz de Referência de Mercado",
    "📈 Inteligência & Histórico Temporal",
])

PRODUTOS_OFICIAIS = ["Frango Congelado", "Boi Gordo", "Ovo", "Suíno Vivo"]
BANDEIRAS_CENCOSUD = [
    "Prezunic",
    "Bretas",
    "Gbarbosa - BA",
    "Gbarbosa - SE",
    "Giga",
]

# ---------------------------------------------------------
# ABA 1: FORMULÁRIO COMPRADOR
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
                "Novo Custo Pago Cencosud (R$):",
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
            "Custo_Pago_Cencosud": custo_pago,
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
        st.dataframe(pd.read_excel(CAMINHO_ENTRADA), use_container_width=True)

# ---------------------------------------------------------
# ABA 2: MATRIZ DE MERCADO
# ---------------------------------------------------------
with aba_analise:
    st.subheader("Gestão do Preço Referência de Mercado (CEPEA/CEAGESP)")
    df_ref_atual = (
        pd.read_excel(CAMINHO_REFERENCIA_MERCADO)
        if os.path.exists(CAMINHO_REFERENCIA_MERCADO)
        else pd.DataFrame({
            "Produto": PRODUTOS_OFICIAIS,
            "Preco_Mercado_Atual": [7.50, 230.00, 115.00, 7.20],
            "Unidade_Medida": [
                "R$ / Kg",
                "R$ / @",
                "R$ / Caixa CIF SP",
                "R$ / Kg SP",
            ],
            "Preco_Mercado_Semana_Anterior": [7.40, 228.00, 112.00, 7.10],
            "Preco_Mercado_Mes_Anterior": [7.20, 225.00, 110.00, 6.90],
        })
    )
    df_editor_ref = st.data_editor(
        df_ref_atual, num_rows="fixed", key="editor_ref_mercado"
    )
    if st.button("💾 Atualizar Referência de Mercado"):
        df_editor_ref.to_excel(CAMINHO_REFERENCIA_MERCADO, index=False)
        st.success("✅ Tabela atualizada!")

# ---------------------------------------------------------
# ABA 3: HISTÓRICO COM SELETOR & TENDÊNCIA DOS ÚLTIMOS 5 ANOS
# ---------------------------------------------------------
with aba_historico:
    st.subheader("📈 Análise Executiva e Tendência Histórica")

    if CAMINHO_HISTORICO and os.path.exists(CAMINHO_HISTORICO):
        try:
            df_raw = pd.read_excel(CAMINHO_HISTORICO)
            df_raw.columns = [str(c).strip() for c in df_raw.columns]

            # Identifica Coluna de Data
            col_data = next(
                (
                    c
                    for c in df_raw.columns
                    if "data" in c.lower() or "date" in c.lower()
                ),
                df_raw.columns[0],
            )
            df_raw[col_data] = pd.to_datetime(df_raw[col_data], errors="coerce")
            df_raw = df_raw.dropna(subset=[col_data]).sort_values(by=col_data)

            # Identifica se o arquivo tem estrutura em Linhas (Coluna Produto) ou Colunas por Produto
            col_prod = next(
                (
                    c
                    for c in df_raw.columns
                    if "prod" in c.lower() or "item" in c.lower()
                ),
                None,
            )

            # 1. TRATAMENTO CASO O ARQUIVO TENHA UMA COLUNA 'PRODUTO'
            if col_prod:
                lista_produtos = [
                    str(p).strip() for p in df_raw[col_prod].dropna().unique()
                ]
                prod_sel_h = st.selectbox(
                    "🔍 Selecione o Produto:", options=lista_produtos
                )
                df_p = df_raw[df_raw[col_prod] == prod_sel_h].copy()

                col_valor = next(
                    (
                        c
                        for c in df_p.columns
                        if "preco" in c.lower()
                        or "custo" in c.lower()
                        or "valor" in c.lower()
                        or "brl" in c.lower()
                    ),
                    df_p.select_dtypes(include=["number"]).columns[0],
                )
                df_p["Preco_Limpo"] = pd.to_numeric(
                    df_p[col_valor], errors="coerce"
                )

            # 2. TRATAMENTO CASO CADA PRODUTO SEJA UMA COLUNA SEPARADA
            else:
                cols_numericas = df_raw.select_dtypes(
                    include=["number", "object"]
                ).columns
                cols_filtradas = [
                    c
                    for c in cols_numericas
                    if c != col_data and "unnamed" not in c.lower()
                ]
                prod_sel_h = st.selectbox(
                    "🔍 Selecione o Produto:", options=cols_filtradas
                )

                df_p = pd.DataFrame({
                    col_data: df_raw[col_data],
                    "Preco_Limpo": pd.to_numeric(
                        df_raw[prod_sel_h]
                        .astype(str)
                        .str.replace("R$", "", regex=False)
                        .str.replace(".", "", regex=False)
                        .str.replace(",", ".", regex=False)
                        .str.strip(),
                        errors="coerce",
                    ),
                })

            df_p = df_p.dropna(subset=["Preco_Limpo", col_data])

            # Filtro dos Últimos 5 Anos de Histórico
            data_limite_5anos = df_p[col_data].max() - pd.DateOffset(years=5)
            df_5anos = df_p[df_p[col_data] >= data_limite_5anos].copy()

            # REGRA DE REGIONALIZAÇÃO/MOEDA
            p_lower = str(prod_sel_h).lower()
            if "frango" in p_lower or "boi" in p_lower:
                st.info(
                    f"ℹ️ Exibindo cotação oficial em **Reais (R$)** para"
                    f" **{prod_sel_h}**."
                )
            elif "ovo" in p_lower:
                st.info(
                    "ℹ️ Exibindo cotação **CIF Região Grande SP** para Ovos."
                )
            elif "suino" in p_lower:
                st.info("ℹ️ Exibindo cotação para **Suíno Vivo (SP)**.")

            # CÁLCULO DOS CARDS EXECUTIVOS
            # 1. Custo Médio do Último Mês (últimos 30 dias da base)
            max_data = df_5anos[col_data].max()
            df_30d = df_5anos[
                df_5anos[col_data] >= (max_data - pd.Timedelta(days=30))
            ]
            custo_medio_30d = df_30d["Preco_Limpo"].mean()

            # 2. Preço da Última Sexta-feira Registrada
            df_sextas = df_5anos[df_5anos[col_data].dt.weekday == 4]
            if not df_sextas.empty:
                preco_ultima_sexta = df_sextas.iloc[-1]["Preco_Limpo"]
                data_sexta_str = df_sextas.iloc[-1][col_data].strftime(
                    "%d/%m/%Y"
                )
            else:
                preco_ultima_sexta = df_5anos.iloc[-1]["Preco_Limpo"]
                data_sexta_str = df_5anos.iloc[-1][col_data].strftime(
                    "%d/%m/%Y"
                )

            st.divider()
            k1, k2, k3 = st.columns(3)
            k1.metric(
                label="Custo Médio (Último Mês)",
                value=(
                    f"R$ {custo_medio_30d:,.2f}"
                    if pd.notnull(custo_medio_30d)
                    else "N/A"
                ),
            )
            k2.metric(
                label=f"Fechamento Sexta-feira ({data_sexta_str})",
                value=(
                    f"R$ {preco_ultima_sexta:,.2f}"
                    if pd.notnull(preco_ultima_sexta)
                    else "N/A"
                ),
            )
            k3.metric(
                label="Volume da Série (5 Anos)",
                value=f"{len(df_5anos)} registros",
            )

            st.divider()

            # GRÁFICO LIMPO DE TENDÊNCIA CONTÍNUA (ÚLTIMOS 5 ANOS)
            st.subheader(
                f"📉 Curva de Tendência Histórica (Últimos 5 Anos) -"
                f" {prod_sel_h}"
            )

            # Agrupamento Mensal para remover poluição visual diária
            df_5anos["Ano_Mês"] = df_5anos[col_data].dt.to_period("M")
            df_tendencia = (
                df_5anos.groupby("Ano_Mês")["Preco_Limpo"]
                .mean()
                .reset_index()
            )
            df_tendencia["Data"] = df_tendencia["Ano_Mês"].dt.to_timestamp()

            df_chart = df_tendencia.set_index("Data")[["Preco_Limpo"]]
            df_chart.columns = ["Preço Médio (R$)"]

            st.line_chart(df_chart, use_container_width=True)

            with st.expander("📋 Ver Tabela de Dados"):
                st.dataframe(
                    df_5anos[[col_data, "Preco_Limpo"]],
                    use_container_width=True,
                )

        except Exception as e:
            st.error(f"❌ Erro ao estruturar o histórico: {e}")
    else:
        st.warning(
            "⚠️ O arquivo de histórico não foi localizado no repositório."
        )
