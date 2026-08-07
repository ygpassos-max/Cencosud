import datetime
import os
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# DIRETÓRIO DE ARMAZENAMENTO COMPATÍVEL COM NUVEM & LOCAL
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# CONFIGURAÇÃO DE TELA STREAMLIT
# ---------------------------------------------------------
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

    with st.form(key="form_comprador_cencosud"):
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

        btn_salvar_comprador = st.form_submit_button(
            label="💾 Registrar Cotação de Compra"
        )

    if btn_salvar_comprador:
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
        st.subheader("📋 Cotações da Semana")
        df_compras = pd.read_excel(CAMINHO_ENTRADA)
        st.dataframe(df_compras, use_container_width=True)

# ---------------------------------------------------------
# ABA 2: MATRIZ DE REFERÊNCIA DE MERCADO
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

    if st.button("💾 Atualizar Tabela de Referência de Mercado"):
        df_editor_ref.to_excel(CAMINHO_REFERENCIA_MERCADO, index=False)
        st.success("✅ Referência de mercado atualizada com sucesso!")

# ---------------------------------------------------------
# ABA 3: INTELIGÊNCIA & HISTÓRICO TEMPORAL (CORRIGIDO)
# ---------------------------------------------------------
with aba_historico:
    st.subheader("📈 Análise de Histórico e Tendência Anual")

    if CAMINHO_HISTORICO and os.path.exists(CAMINHO_HISTORICO):
        try:
            df_h = pd.read_excel(CAMINHO_HISTORICO)

            if df_h.empty:
                st.warning("⚠️ O arquivo de histórico está vazio.")
            else:
                df_h.columns = [str(col).strip() for col in df_h.columns]

                # Localização Inteligente de Colunas
                col_data = next(
                    (
                        c
                        for c in df_h.columns
                        if "data" in c.lower() or "date" in c.lower()
                    ),
                    df_h.columns[0],
                )
                
                col_prod = next(
                    (
                        c
                        for c in df_h.columns
                        if "prod" in c.lower() or "item" in c.lower()
                    ),
                    None,
                )

                # Busca Coluna de Preço
                cols_preco_possiveis = [
                    c
                    for c in df_h.columns
                    if "preco" in c.lower()
                    or "custo" in c.lower()
                    or "valor" in c.lower()
                    or "brl" in c.lower()
                ]

                if cols_preco_possiveis:
                    col_preco = cols_preco_possiveis[0]
                else:
                    col_preco = df_h.columns[1] if len(df_h.columns) > 1 else df_h.columns[0]

                # Tratamento de Formato de Preço (se vier como texto "15,50")
                if df_h[col_preco].dtype == object:
                    df_h[col_preco] = (
                        df_h[col_preco]
                        .astype(str)
                        .str.replace("R$", "", regex=False)
                        .str.replace(".", "", regex=False)
                        .str.replace(",", ".", regex=False)
                        .str.strip()
                    )
                df_h[col_preco] = pd.to_numeric(df_h[col_preco], errors="coerce")

                # Tratamento de Data
                df_h[col_data] = pd.to_datetime(df_h[col_data], errors="coerce")
                df_h = df_h.dropna(subset=[col_data, col_preco]).sort_values(
                    by=col_data
                )

                col_esq_filt, col_dir_info = st.columns([1, 2])

                with col_esq_filt:
                    if col_prod and df_h[col_prod].nunique() > 0:
                        produtos_disponiveis = list(df_h[col_prod].dropna().unique())
                        prod_selecionado = st.selectbox(
                            "Selecione o Produto para Análise:",
                            options=produtos_disponiveis,
                        )
                        df_filtrado = df_h[df_h[col_prod] == prod_selecionado]
                    else:
                        prod_selecionado = "Geral"
                        df_filtrado = df_h

                with col_dir_info:
                    p_str = str(prod_selecionado).lower()
                    if "frango" in p_str or "boi" in p_str:
                        st.info(f"ℹ️ Exibindo cotação oficial em **Reais (R$)** para **{prod_selecionado}**.")
                    elif "ovo" in p_str:
                        st.info("ℹ️ Exibindo cotação **CIF Região Grande SP** para Ovos.")
                    elif "suino" in p_str:
                        st.info("ℹ️ Exibindo cotação para **Suíno Vivo (SP)**.")
                    else:
                        st.info("ℹ️ Exibindo série histórica do produto selecionado.")

                if not df_filtrado.empty:
                    # Filtro da Última Sexta-feira Registrada
                    df_sextas = df_filtrado[df_filtrado[col_data].dt.weekday == 4]

                    if not df_sextas.empty:
                        ultimo_preco_sexta = df_sextas.iloc[-1][col_preco]
                        data_ultima_sexta = df_sextas.iloc[-1][col_data].strftime("%d/%m/%Y")
                    else:
                        ultimo_preco_sexta = df_filtrado.iloc[-1][col_preco]
                        data_ultima_sexta = df_filtrado.iloc[-1][col_data].strftime("%d/%m/%Y")

                    # Cálculo do Custo Médio do Último Mês
                    ultima_data_base = df_filtrado[col_data].max()
                    data_limite_30d = ultima_data_base - pd.Timedelta(days=30)
                    df_ultimo_mes = df_filtrado[df_filtrado[col_data] >= data_limite_30d]
                    media_custo_ultimo_mes = df_ultimo_mes[col_preco].mean()

                    st.divider()

                    # CARDS DE KPIS HISTÓRICOS
                    m1, m2, m3 = st.columns(3)
                    m1.metric(
                        label="Custo Médio (Último Mês)",
                        value=f"R$ {media_custo_ultimo_mes:,.2f}" if pd.notnull(media_custo_ultimo_mes) else "N/A",
                    )
                    m2.metric(
                        label=f"Preço Última Sexta-feira ({data_ultima_sexta})",
                        value=f"R$ {ultimo_preco_sexta:,.2f}" if pd.notnull(ultimo_preco_sexta) else "N/A",
                    )
                    m3.metric(
                        label="Total de Registros",
                        value=f"{len(df_filtrado)} cotações",
                    )

                    st.divider()

                    # GRÁFICO DE TENDÊNCIA DE ANOS
                    st.subheader(f"📊 Curva Múltipla de Tendência Anual - {prod_selecionado}")

                    df_grafico = df_filtrado.copy()
                    df_grafico["Ano"] = df_grafico[col_data].dt.year
                    df_grafico["Mês_Dia"] = df_grafico[col_data].dt.strftime("%m-%d")

                    df_pivot = df_grafico.pivot_table(
                        index="Mês_Dia",
                        columns="Ano",
                        values=col_preco,
                        aggfunc="mean",
                    )
                    st.line_chart(df_pivot, use_container_width=True)

                    with st.expander("📋 Ver Dados Históricos Detalhados em Tabela"):
                        cols_mostrar = [c for c in [col_data, col_prod, col_preco] if c in df_filtrado.columns]
                        st.dataframe(df_filtrado[cols_mostrar], use_container_width=True)
                else:
                    st.warning("⚠️ Nenhum registro encontrado para este produto.")

        except Exception as e:
            st.error(f"❌ Erro ao processar a base de histórico: {e}")
    else:
        st.warning("⚠️ O arquivo **historico_real.xls** precisa estar presente no repositório do GitHub.")
