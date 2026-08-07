import datetime
import os
import pandas as pd
import streamlit as st

PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))

CAMINHO_HISTORICO = os.path.join(PASTA_PROJETO, "historico_real.xlsx")
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
# ABA 3: HISTÓRICO DIRETO E PERFEITO
# ---------------------------------------------------------
with aba_historico:
    st.subheader("📈 Análise Executiva e Tendência Histórica")

    if os.path.exists(CAMINHO_HISTORICO):
        try:
            df_h = pd.read_excel(CAMINHO_HISTORICO)

            # Padroniza nomes das colunas
            df_h.columns = [str(c).strip().capitalize() for c in df_h.columns]

            if (
                "Data" in df_h.columns
                and "Produto" in df_h.columns
                and "Preco" in df_h.columns
            ):
                df_h["Data"] = pd.to_datetime(df_h["Data"], errors="coerce")
                df_h["Preco"] = pd.to_numeric(df_h["Preco"], errors="coerce")
                df_h = df_h.dropna().sort_values(by="Data")

                # Seletor de Produto
                prods_unicos = list(df_h["Produto"].unique())
                prod_selecionado = st.selectbox(
                    "🔍 Selecione o Produto:", options=prods_unicos
                )

                df_p = df_h[df_h["Produto"] == prod_selecionado].copy()

                # Mensagens de Regra do Produto
                p_str = str(prod_selecionado).lower()
                if "frango" in p_str or "boi" in p_str:
                    st.info(
                        f"ℹ️ Exibindo cotação oficial em **Reais (R$)** para"
                        f" **{prod_selecionado}**."
                    )
                elif "ovo" in p_str:
                    st.info(
                        "ℹ️ Exibindo cotação **CIF Região Grande SP** para"
                        " Ovos."
                    )
                elif "suino" in p_str:
                    st.info("ℹ️ Exibindo cotação para **Suíno Vivo (SP)**.")

                # Cálculos dos KPIs
                max_d = df_p["Data"].max()
                df_30d = df_p[df_p["Data"] >= (max_d - pd.Timedelta(days=30))]
                custo_medio_30d = df_30d["Preco"].mean()

                # Fechamento de Sexta
                df_sextas = df_p[df_p["Data"].dt.weekday == 4]
                if not df_sextas.empty:
                    preco_sexta = df_sextas.iloc[-1]["Preco"]
                    dt_sexta = df_sextas.iloc[-1]["Data"].strftime("%d/%m/%Y")
                else:
                    preco_sexta = df_p.iloc[-1]["Preco"]
                    dt_sexta = df_p.iloc[-1]["Data"].strftime("%d/%m/%Y")

                st.divider()
                k1, k2, k3 = st.columns(3)
                k1.metric(
                    label="Custo Médio (Último Mês)",
                    value=f"R$ {custo_medio_30d:,.2f}",
                )
                k2.metric(
                    label=f"Fechamento Sexta-feira ({dt_sexta})",
                    value=f"R$ {preco_sexta:,.2f}",
                )
                k3.metric(
                    label="Volume de Registros",
                    value=f"{len(df_p)} cotações",
                )

                st.divider()

                # Gráfico de Tendência (Últimos 5 Anos agrupado por Mês)
                st.subheader(
                    f"📉 Curva de Tendência Mensal -"
                    f" {prod_selecionado}"
                )

                df_5y = df_p[
                    df_p["Data"] >= (max_d - pd.DateOffset(years=5))
                ].copy()
                df_5y["Ano_Mes"] = df_5y["Data"].dt.to_period("M")

                df_chart_data = (
                    df_5y.groupby("Ano_Mes")["Preco"].mean().reset_index()
                )
                df_chart_data["Data_Plot"] = df_chart_data[
                    "Ano_Mes"
                ].dt.to_timestamp()

                df_final_chart = df_chart_data.set_index("Data_Plot")[["Preco"]]
                df_final_chart.columns = ["Preço Médio (R$)"]

                st.line_chart(df_final_chart, use_container_width=True)

                with st.expander("📋 Ver Tabela de Dados"):
                    st.dataframe(
                        df_p[["Data", "Produto", "Preco"]],
                        use_container_width=True,
                    )

            else:
                st.error(
                    "❌ As colunas da planilha precisam se chamar exatamente:"
                    " **Data**, **Produto** e **Preco**."
                )

        except Exception as e:
            st.error(f"❌ Erro ao ler a planilha: {e}")
    else:
        st.warning(
            "⚠️ Suba o arquivo **historico_real.xlsx** com as colunas Data,"
            " Produto e Preco no GitHub."
        )
