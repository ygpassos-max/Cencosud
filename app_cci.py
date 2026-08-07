import datetime
import os
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# DIRETÓRIO DE ARMAZENAMENTO COMPATÍVEL COM NUVEM & LOCAL
# ---------------------------------------------------------
PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))

# Aceita tanto .xlsx quanto .xls
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
st.caption("Plataforma de Inteligência Comercial e Variação de Custos")

aba_entrada, aba_analise = st.tabs(
    ["📝 Entrada & Indicadores Semanais", "📊 Matriz de Mercado & Histórico Real"]
)

PRODUTOS_OFICIAIS = ["Filé de Peito", "Ovo c/ 20", "Alcatra", "Contra Filé"]
BANDEIRAS_CENCOSUD = [
    "Prezunic",
    "Bretas",
    "Gbarbosa - BA",
    "Gbarbosa - SE",
    "Giga",
]

# ---------------------------------------------------------
# ABA 1: FORMULÁRIO + KPIS + VARIAÇÃO
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
        df_compras = pd.read_excel(CAMINHO_ENTRADA)

        if os.path.exists(CAMINHO_REFERENCIA_MERCADO):
            df_ref = pd.read_excel(CAMINHO_REFERENCIA_MERCADO)
        else:
            df_ref = pd.DataFrame({
                "Produto": PRODUTOS_OFICIAIS,
                "Preco_Mercado_Atual": [14.50, 8.20, 38.00, 42.00],
                "Preco_Mercado_Semana_Anterior": [14.20, 8.50, 37.50, 41.00],
                "Preco_Mercado_Mes_Anterior": [13.80, 8.00, 36.00, 40.00],
            })

        df_cruzado = pd.merge(df_compras, df_ref, on="Produto", how="left")

        df_cruzado["Var_%_Mercado_Semanal"] = (
            (
                df_cruzado["Preco_Mercado_Atual"]
                - df_cruzado["Preco_Mercado_Semana_Anterior"]
            )
            / df_cruzado["Preco_Mercado_Semana_Anterior"]
        ) * 100

        df_cruzado["Var_%_Mercado_Mensal"] = (
            (
                df_cruzado["Preco_Mercado_Atual"]
                - df_cruzado["Preco_Mercado_Mes_Anterior"]
            )
            / df_cruzado["Preco_Mercado_Mes_Anterior"]
        ) * 100

        df_cruzado = df_cruzado.sort_values(
            by=["Produto", "Bandeira", "Data_Compra"]
        )
        df_cruzado["Ultimo_Custo_Sistema"] = df_cruzado.groupby(
            ["Produto", "Bandeira"]
        )["Custo_Pago_Cencosud"].shift(1)

        df_cruzado["Ultimo_Custo_Sistema"] = df_cruzado[
            "Ultimo_Custo_Sistema"
        ].fillna(df_cruzado["Custo_Pago_Cencosud"])

        df_cruzado["Var_%_Custo_Cencosud"] = (
            (
                df_cruzado["Custo_Pago_Cencosud"]
                - df_cruzado["Ultimo_Custo_Sistema"]
            )
            / df_cruzado["Ultimo_Custo_Sistema"]
        ) * 100

        df_cruzado["Spread_%"] = (
            (
                df_cruzado["Custo_Pago_Cencosud"]
                - df_cruzado["Preco_Mercado_Atual"]
            )
            / df_cruzado["Preco_Mercado_Atual"]
        ) * 100

        df_cruzado["Diferenca_Total_BRL"] = (
            df_cruzado["Custo_Pago_Cencosud"]
            - df_cruzado["Preco_Mercado_Atual"]
        )

        st.subheader("📈 Painel Executivo de Desempenho")
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)

        total_cotacoes = len(df_cruzado)
        spread_medio = df_cruzado["Spread_%"].mean()
        var_mercado_semanal_media = df_cruzado["Var_%_Mercado_Semanal"].mean()
        economia_estimada = df_cruzado["Diferenca_Total_BRL"].sum()

        kpi1.metric(
            label="Total de Cotações", value=f"{total_cotacoes} compras"
        )
        kpi2.metric(
            label="Spread Médio vs Mercado",
            value=f"{spread_medio:.2f}%",
            delta=f"{spread_medio:.2f}%",
            delta_color="inverse",
        )
        kpi3.metric(
            label="Var. Média Mercado (7d)",
            value=f"{var_mercado_semanal_media:.2f}%",
        )
        kpi4.metric(
            label="Impacto / Desvio Total BRL",
            value=f"R$ {economia_estimada:,.2f}",
            delta=(
                "Economia"
                if economia_estimada < 0
                else "Custo Acima do Mercado"
            ),
            delta_color="inverse",
        )

        st.divider()
        st.subheader("📋 Acompanhamento Semanal de Variações de Preço")

        cols_exibir = [
            "Data_Compra",
            "Bandeira",
            "Produto",
            "Custo_Pago_Cencosud",
            "Ultimo_Custo_Sistema",
            "Var_%_Custo_Cencosud",
            "Preco_Mercado_Atual",
            "Var_%_Mercado_Semanal",
            "Var_%_Mercado_Mensal",
            "Spread_%",
            "Fornecedor",
        ]

        st.dataframe(
            df_cruzado[cols_exibir].style.format({
                "Custo_Pago_Cencosud": "R$ {:.2f}",
                "Ultimo_Custo_Sistema": "R$ {:.2f}",
                "Preco_Mercado_Atual": "R$ {:.2f}",
                "Var_%_Custo_Cencosud": "{:.2f}%",
                "Var_%_Mercado_Semanal": "{:.2f}%",
                "Var_%_Mercado_Mensal": "{:.2f}%",
                "Spread_%": "{:.2f}%",
            }),
            use_container_width=True,
        )

# ---------------------------------------------------------
# ABA 2: REFERÊNCIA DE MERCADO & CORREÇÃO DO HISTÓRICO
# ---------------------------------------------------------
with aba_analise:
    st.subheader("Gestão de Referência de Mercado e Histórico Real")

    col_esq, col_dir = st.columns(2)

    with col_esq:
        st.write("**1. Tabela de Preço Referência de Mercado (CEPEA/CEAGESP)**")

        df_ref_atual = (
            pd.read_excel(CAMINHO_REFERENCIA_MERCADO)
            if os.path.exists(CAMINHO_REFERENCIA_MERCADO)
            else pd.DataFrame({
                "Produto": PRODUTOS_OFICIAIS,
                "Preco_Mercado_Atual": [14.50, 8.20, 38.00, 42.00],
                "Preco_Mercado_Semana_Anterior": [14.20, 8.50, 37.50, 41.00],
                "Preco_Mercado_Mes_Anterior": [13.80, 8.00, 36.00, 40.00],
            })
        )

        df_editor_ref = st.data_editor(
            df_ref_atual, num_rows="fixed", key="editor_ref_mercado"
        )

        if st.button("💾 Atualizar Tabela de Referência de Mercado"):
            df_editor_ref.to_excel(CAMINHO_REFERENCIA_MERCADO, index=False)
            st.success("✅ Referência de mercado atualizada com sucesso!")

    with col_dir:
        st.write("**2. Leitura da Base Histórica Real**")
        if CAMINHO_HISTORICO and os.path.exists(CAMINHO_HISTORICO):
            try:
                df_hist = pd.read_excel(CAMINHO_HISTORICO)

                st.success(f"✅ Histórico lido ({len(df_hist)} registros).")
                st.dataframe(df_hist, height=220, use_container_width=True)

                col_num = df_hist.select_dtypes(
                    include=["float", "int"]
                ).columns
                if len(col_num) > 0:
                    st.write("**Evolução Histórica:**")
                    st.line_chart(df_hist[col_num[0]])

            except Exception as e:
                st.error(
                    f"❌ Erro ao ler o arquivo de histórico: {e}. Verifique se"
                    " a extensão está correta."
                )
        else:
            st.warning(
                "⚠️ Arquivo de histórico não localizado na raiz do repositório."
            )
