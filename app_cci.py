import datetime
import os
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# DIRETÓRIO DE ARMAZENAMENTO COMPATÍVEL COM NUVEM & LOCAL
# ---------------------------------------------------------
PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))

CAMINHO_HISTORICO = os.path.join(PASTA_PROJETO, "historico_real.xlsx")
CAMINHO_ENTRADA = os.path.join(PASTA_PROJETO, "cotacoes_semanais.xlsx")
CAMINHO_REFERENCIA_MERCADO = os.path.join(
    PASTA_PROJETO, "referencia_mercado.xlsx"
)
CAMINHO_SAIDA_CONSOLIDADA = os.path.join(
    PASTA_PROJETO, "CCI_Matriz_Decisao_Consolidada.xlsx"
)

# ---------------------------------------------------------
# INTERFACE STREAMLIT
# ---------------------------------------------------------
st.set_page_config(
    page_title="Cencosud Commodity Intelligence (CCI)",
    page_icon="🥩",
    layout="wide",
)

st.title("🥩 Cencosud Commodity Intelligence (CCI)")
st.caption("Painel Online de Cotações e Inteligência Comercial")

aba_entrada, aba_analise = st.tabs(
    ["📝 Entrada do Comprador (Semanal)", "📊 Matriz de Decisão & Histórico Real"]
)

PRODUTOS_OFICIAIS = [
    "Filé de Peito",
    "Ovo c/ 20",
    "Alcatra",
    "Contra Filé",
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
    st.write(
        "Selecione a bandeira, o produto e informe os dados negociados nesta"
        " semana."
    )

    with st.form(key="form_comprador_cencosud"):
        col1, col2 = st.columns(2)

        with col1:
            bandeira_sel = st.selectbox(
                "Bandeira Cencosud:", options=BANDEIRAS_CENCOSUD
            )
            produto_sel = st.selectbox("Produto:", options=PRODUTOS_OFICIAIS)

        with col2:
            custo_pago = st.number_input(
                "Custo Pago Cencosud (R$):",
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
        st.success("✅ Cotação registrada com sucesso na nuvem!")

    st.divider()
    if os.path.exists(CAMINHO_ENTRADA):
        st.subheader("📋 Cotações Registradas nesta Semana")
        st.dataframe(pd.read_excel(CAMINHO_ENTRADA), use_container_width=True)

# ---------------------------------------------------------
# ABA 2: PAINEL DE INTELIGÊNCIA & HISTÓRICO
# ---------------------------------------------------------
with aba_analise:
    st.subheader("Painel Geral de Inteligência Comercial")

    col_esq, col_dir = st.columns(2)

    with col_esq:
        st.write("**1. Referência Semanal de Mercado (CEPEA/CEAGESP)**")

        df_ref_atual = (
            pd.read_excel(CAMINHO_REFERENCIA_MERCADO)
            if os.path.exists(CAMINHO_REFERENCIA_MERCADO)
            else pd.DataFrame({
                "Produto": PRODUTOS_OFICIAIS,
                "Preco_Mercado_Referencia_BRL": [14.50, 8.20, 38.00, 42.00],
            })
        )

        df_editor_ref = st.data_editor(
            df_ref_atual, num_rows="fixed", key="editor_ref_mercado"
        )

        if st.button("💾 Atualizar Referência de Mercado"):
            df_editor_ref.to_excel(CAMINHO_REFERENCIA_MERCADO, index=False)
            st.success("✅ Tabela de Preços de Referência atualizada!")

    with col_dir:
        st.write("**2. Base Histórica Real**")
        if os.path.exists(CAMINHO_HISTORICO):
            try:
                df_hist = pd.read_excel(CAMINHO_HISTORICO)
                st.success(
                    f"✅ Histórico Real carregado ({len(df_hist)} linhas)."
                )
                st.dataframe(df_hist, height=200, use_container_width=True)
            except Exception as e:
                st.error(f"❌ Erro ao ler o arquivo: {e}")
        else:
            st.warning(
                "⚠️ Suba o arquivo **historico_real.xlsx** no seu repositório"
                " do GitHub."
            )

    st.divider()

    if os.path.exists(CAMINHO_ENTRADA):
        st.subheader("📊 Cruzamento: Custo Pago vs. Preço Mercado")
        df_compras = pd.read_excel(CAMINHO_ENTRADA)

        if os.path.exists(CAMINHO_REFERENCIA_MERCADO):
            df_ref = pd.read_excel(CAMINHO_REFERENCIA_MERCADO)
            df_cruzado = pd.merge(df_compras, df_ref, on="Produto", how="left")

            if "Preco_Mercado_Referencia_BRL" in df_cruzado.columns:
                df_cruzado["Spread_BRL"] = (
                    df_cruzado["Custo_Pago_Cencosud"]
                    - df_cruzado["Preco_Mercado_Referencia_BRL"]
                )
                df_cruzado["Spread_%"] = (
                    df_cruzado["Spread_BRL"]
                    / df_cruzado["Preco_Mercado_Referencia_BRL"]
                ) * 100

                def acao(row):
                    if row["Spread_%"] > 3.0:
                        return "⚠️ NEGOCIAR: Custo >3% acima do mercado"
                    elif row["Spread_%"] < 0:
                        return "✅ EXCELENTE: Preço abaixo do mercado"
                    else:
                        return "➡️ DENTRO DA META: Alinhado ao benchmark"

                df_cruzado["Recomendacao_CCI"] = df_cruzado.apply(acao, axis=1)
                st.dataframe(df_cruzado, use_container_width=True)
