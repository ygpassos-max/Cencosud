import datetime
import io
import os
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

# ---------------------------------------------------------
# LINK DA SUA PLANILHA DO GOOGLE SHEETS (COLE SEU LINK AQUI)
# ---------------------------------------------------------
URL_GOOGLE_SHEETS = "https://docs.google.com/spreadsheets/d/1RzJ41JQQPM6ibJrnxmaJNq5JCZiROXAemsaxBvcQV2I/edit?usp=sharing"

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

CAMINHO_CUSTO_SISTEMA = os.path.join(
    PASTA_PROJETO, "custo_atual_sistema.xlsx"
)

# ---------------------------------------------------------
# LEITURA E GRAVAÇÃO NO GOOGLE SHEETS
# ---------------------------------------------------------
def carregar_cotacoes_google():
    """Carrega as cotações diretamente da Planilha do Google Sheets."""
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(spreadsheet=URL_GOOGLE_SHEETS, ttl=0)
        if df is not None and not df.empty:
            df.columns = [str(c).strip() for c in df.columns]
            return df
    except Exception as e:
        pass
    return pd.DataFrame(
        columns=[
            "Data_Compra",
            "Bandeira",
            "Categoria",
            "Produto",
            "Commodity_Referencia",
            "Cotacao_Cencosud",
            "Fornecedor",
            "Data_Captura",
        ]
    )

def salvar_cotacao_google(novo_registro):
    """Adiciona uma nova cotação na Planilha do Google Sheets."""
    try:
        df_atual = carregar_cotacoes_google()
        df_novo = pd.DataFrame([novo_registro])
        df_final = pd.concat([df_atual, df_novo], ignore_index=True)
        
        conn = st.connection("gsheets", type=GSheetsConnection)
        conn.update(spreadsheet=URL_GOOGLE_SHEETS, data=df_final)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"❌ Erro ao salvar no Google Sheets: {e}")
        return False

# ---------------------------------------------------------
# EXTRAI VALORES DE REFERÊNCIA HISTÓRICOS DE MERCADO
# ---------------------------------------------------------
@st.cache_data(ttl=3600)
def obter_precos_referencia_historico(caminho_file):
    precos_ref = {}
    if caminho_file and os.path.exists(caminho_file):
        try:
            xls = pd.ExcelFile(caminho_file)
            mapa_referencias = {
                "Frango Congelado": ["Frango Congelado"],
                "Boi Gordo": ["Boi ", "Boi"],
                "Ovo": ["Ovos", "Ovo"],
                "Suíno Vivo": ["Suino Vivo", "Suino"],
                "Banana": ["Banana"],
                "Batata": ["Batata"],
                "Cebola": ["Cebola"],
                "Maçã Gala": ["Maça", "Maca", "Maçã"],
                "Tomate": ["Tomate"],
            }

            for comm_ref, abas_possiveis in mapa_referencias.items():
                df_acumulado = pd.DataFrame()
                for aba in abas_possiveis:
                    if aba in xls.sheet_names:
                        df_teste = pd.read_excel(
                            caminho_file, sheet_name=aba, nrows=2
                        )
                        cols_teste = [str(c).strip() for c in df_teste.columns]

                        if "Dia" in cols_teste and "Mês" in cols_teste:
                            df = pd.read_excel(caminho_file, sheet_name=aba)
                            df.columns = [str(c).strip() for c in df.columns]

                            col_p = next(
                                (
                                    c
                                    for c in ["Preço k", "Preco k", "PreçoC"]
                                    if c in df.columns
                                ),
                                df.columns[-1],
                            )

                            df["Data_Tmp"] = pd.to_datetime(
                                df["Ano"].astype(str)
                                + "-"
                                + df["Mês"].astype(str)
                                + "-"
                                + df["Dia"].astype(str),
                                errors="coerce",
                            )
                            df["Preco_Limpo"] = pd.to_numeric(
                                df[col_p].astype(str).str.replace(",", "."),
                                errors="coerce",
                            )
                            df_sub = df[["Data_Tmp", "Preco_Limpo"]].dropna()
                        else:
                            df = pd.read_excel(
                                caminho_file, sheet_name=aba, skiprows=1
                            )
                            df.columns = [str(c).strip() for c in df.columns]

                            col_dt = next(
                                (c for c in df.columns if "data" in c.lower()),
                                df.columns[0],
                            )
                            col_p = (
                                "Branco"
                                if "ovos" in aba.lower()
                                else (
                                    "SP"
                                    if "suino" in aba.lower()
                                    else "À vista R$"
                                )
                            )

                            df["Data_Tmp"] = pd.to_datetime(
                                df[col_dt], format="%d/%m/%Y", errors="coerce"
                            )
                            df["Preco_Limpo"] = pd.to_numeric(
                                df[col_p], errors="coerce"
                            )
                            df_sub = df[["Data_Tmp", "Preco_Limpo"]].dropna()

                        df_acumulado = pd.concat([df_acumulado, df_sub])

                if not df_acumulado.empty:
                    df_acumulado = df_acumulado.sort_values(by="Data_Tmp")
                    max_d = df_acumulado["Data_Tmp"].max()

                    df_u30 = df_acumulado[
                        df_acumulado["Data_Tmp"] >= (max_d - pd.Timedelta(days=30))
                    ]
                    preco_mes_ant = (
                        df_u30["Preco_Limpo"].mean()
                        if not df_u30.empty
                        else df_acumulado["Preco_Limpo"].mean()
                    )

                    df_sextas = df_acumulado[
                        df_acumulado["Data_Tmp"].dt.weekday == 4
                    ]
                    if not df_sextas.empty:
                        preco_sem_ant = df_sextas.iloc[-1]["Preco_Limpo"]
                    else:
                        preco_sem_ant = df_acumulado.iloc[-1]["Preco_Limpo"]

                    precos_ref[comm_ref] = {
                        "Preco_Ref_Semana_Anterior": preco_sem_ant,
                        "Preco_Ref_Mes_Anterior": preco_mes_ant,
                    }

        except Exception:
            pass

    for c in [
        "Frango Congelado",
        "Boi Gordo",
        "Ovo",
        "Suíno Vivo",
        "Banana",
        "Batata",
        "Cebola",
        "Maçã Gala",
        "Tomate",
    ]:
        if c not in precos_ref:
            precos_ref[c] = {
                "Preco_Ref_Semana_Anterior": None,
                "Preco_Ref_Mes_Anterior": None,
            }

    df_res = pd.DataFrame.from_dict(precos_ref, orient="index").reset_index()
    df_res.columns = [
        "Produto",
        "Preco_Ref_Semana_Anterior",
        "Preco_Ref_Mes_Anterior",
    ]
    return df_res


# ---------------------------------------------------------
# CONFIGURAÇÃO DE TELA E TÍTULO
# ---------------------------------------------------------
st.set_page_config(
    page_title="Ferramenta de Negociação Cencosud",
    page_icon="🥩",
    layout="wide",
)

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
    st.caption(
        "Programa VIVA Perecíveis • Inteligência Comercial e Monitoramento de"
        " Custos"
    )

st.divider()

aba_entrada, aba_matriz, aba_historico = st.tabs([
    "📝 Entrada Comprador (Semanal)",
    "📊 Matriz de Decisão & Spreads",
    "📈 Inteligência & Histórico Temporal",
])

# CATÁLOGO ORDENADO
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
    "Banana": "Banana",
    "Batata": "Batata",
    "Cebola": "Cebola",
    "Maçã Gala": "Maçã Gala",
    "Tomate": "Tomate",
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
                "Cotação (R$/kg):",
                min_value=0.0,
                value=5.00,
                step=0.05,
                format="%.2f",
                help="Informe o valor cotado em R$/kg para padrão universal.",
            )
        with c_form2:
            data_negocio = st.date_input(
                "Data Cotação:", value=datetime.date.today()
            )

        commodity_ref = REFERENCIA_COMMODITY.get(
            produto_sel, produto_sel
        )
        st.info(
            f"💡 **Padrão de Entrada:** Cotação calculada em **R$/kg**. O item"
            f" **{produto_sel}** será comparado com a referência"
            f" **{commodity_ref}** na Matriz."
        )

        btn_salvar = st.form_submit_button(
            label="💾 Registrar Cotação no Google Sheets"
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

        if salvar_cotacao_google(novo_registro):
            st.success(
                f"✅ Cotação de **{produto_sel}** ({fornecedor_sel}) em R$/kg"
                " gravada com sucesso no Google Sheets!"
            )

    st.divider()

    df_historico_cotacoes = carregar_cotacoes_google()
    if not df_historico_cotacoes.empty:
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
            "📋 Ver Cotações Registradas (Google Sheets Nuvem)"
        ):
            st.dataframe(
                df_historico_cotacoes[cols_existentes].style.format({
                    "Cotacao_Cencosud": "R$ {:.2f}/kg"
                }, na_rep="-"),
                use_container_width=True,
            )

# ---------------------------------------------------------
# ABA 2: MATRIZ DE DECISÃO
# ---------------------------------------------------------
with aba_matriz:
    st.subheader("📊 Matriz Comercial: Custo Atual vs Cotação vs Variação de Mercado")

    df_precos_ref = obter_precos_referencia_historico(CAMINHO_HISTORICO)
    df_cotacoes = carregar_cotacoes_google()

    if (
        "Custo_Pago_Cencosud" in df_cotacoes.columns
        and "Cotacao_Cencosud" not in df_cotacoes.columns
    ):
        df_cotacoes["Cotacao_Cencosud"] = df_cotacoes["Custo_Pago_Cencosud"]

    datas_disponiveis = (
        sorted(df_cotacoes["Data_Compra"].astype(str).unique(), reverse=True)
        if not df_cotacoes.empty and "Data_Compra" in df_cotacoes.columns
        else []
    )

    f1, f2, f3, f4 = st.columns(4)
    with f1:
        opcao_rodada = st.selectbox(
            "📅 Rodada / Data Cotação:",
            options=["Última Rodada Registrada"] + datas_disponiveis,
        )
    with f2:
        bandeira_filtro = st.selectbox(
            "🔍 Bandeira Cencosud:",
            options=["Todas as Bandeiras"] + BANDEIRAS_CENCOSUD,
        )
    with f3:
        categoria_filtro = st.selectbox(
            "🔍 Categoria:",
            options=["Todas as Categorias"] + sorted(list(ESTRUTURA_PRODUTOS.keys())),
        )
    with f4:
        todos_fornecedores = sorted(
            list(
                set([
                    forn
                    for sub in FORNECEDORES_POR_CATEGORIA.values()
                    for forn in sub
                ])
            )
        )
        fornecedor_filtro = st.selectbox(
            "🔍 Comparar Fornecedor:",
            options=["Todos os Fornecedores"] + todos_fornecedores,
        )

    if not df_cotacoes.empty and "Data_Compra" in df_cotacoes.columns:
        if opcao_rodada == "Última Rodada Registrada":
            ultima_dt = df_cotacoes["Data_Compra"].astype(str).max()
            df_cotacoes_f = df_cotacoes[
                df_cotacoes["Data_Compra"].astype(str) == ultima_dt
            ].copy()
        else:
            df_cotacoes_f = df_cotacoes[
                df_cotacoes["Data_Compra"].astype(str) == opcao_rodada
            ].copy()
    else:
        df_cotacoes_f = df_cotacoes.copy()

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
            "Custo_Atual_Sistema": [12.50]
            * (len(todos_itens) * len(BANDEIRAS_CENCOSUD)),
        })
    )

    if bandeira_filtro != "Todas as Bandeiras":
        df_custo_f = df_custo_sis[
            df_custo_sis["Bandeira"] == bandeira_filtro
        ].copy()
        df_cot_f = df_cotacoes_f[
            df_cotacoes_f["Bandeira"] == bandeira_filtro
        ].copy()
    else:
        df_custo_f = (
            df_custo_sis.groupby("Produto", as_index=False)[
                "Custo_Atual_Sistema"
            ].mean()
        )
        df_custo_f["Bandeira"] = "Todas"
        df_cot_f = df_cotacoes_f.copy()

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

    df_matriz = pd.merge(df_custo_f, df_cot_f_agrupado, on="Produto", how="left")
    df_matriz = df_matriz.dropna(subset=["Cotacao_Cencosud"]).copy()

    df_matriz["Commodity_Referencia"] = df_matriz["Produto"].map(
        lambda p: REFERENCIA_COMMODITY.get(p, p)
    )

    df_matriz = pd.merge(
        df_matriz,
        df_precos_ref,
        left_on="Commodity_Referencia",
        right_on="Produto",
        how="left",
        suffixes=("", "_Ref"),
    )

    if not df_matriz.empty:
        df_matriz["Var_%_Cotacao_vs_Custo"] = (
            (df_matriz["Cotacao_Cencosud"] - df_matriz["Custo_Atual_Sistema"])
            / df_matriz["Custo_Atual_Sistema"]
        ) * 100

        df_matriz["Var_%_Ref_Mercado_Semanal"] = (
            (df_matriz["Cotacao_Cencosud"] - df_matriz["Preco_Ref_Semana_Anterior"])
            / df_matriz["Preco_Ref_Semana_Anterior"]
        ) * 100

        df_matriz["Var_%_Ref_Mercado_Mensal"] = (
            (df_matriz["Cotacao_Cencosud"] - df_matriz["Preco_Ref_Mes_Anterior"])
            / df_matriz["Preco_Ref_Mes_Anterior"]
        ) * 100

        df_matriz["Spread_BRL"] = (
            df_matriz["Cotacao_Cencosud"] - df_matriz["Custo_Atual_Sistema"]
        )

        def diagnostico(row):
            if pd.isna(row["Cotacao_Cencosud"]):
                return "⚪ Sem Cotação"
            var_sem = row.get("Var_%_Ref_Mercado_Semanal", 0)

            if pd.isna(var_sem):
                var_sem = 0

            if var_sem > 1.5:
                return "🔴 ALERTA"
            elif var_sem < 0:
                return "🟢 EXCELENTE"
            else:
                return "🟡 ALINHADO"

        df_matriz["Diagnostico_CCI"] = df_matriz.apply(diagnostico, axis=1)

    st.divider()

    total_cotados = len(df_matriz)
    qtd_alertas = (
        len(df_matriz[df_matriz["Diagnostico_CCI"].str.contains("🔴", na=False)])
        if not df_matriz.empty
        else 0
    )
    saldo_spread = (
        df_matriz["Spread_BRL"].sum() if not df_matriz.empty else 0.0
    )

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

    if not df_matriz.empty:
        dicionario_colunas = {
            "Bandeira": "Bandeira",
            "Produto": "Produto",
            "Fornecedor": "Fornecedor",
            "Custo_Atual_Sistema": "Custo Atual",
            "Cotacao_Cencosud": "Cotação (R$/kg)",
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

        df_exibicao = df_matriz[cols_ordem_exata].rename(
            columns=dicionario_colunas
        )

        def highlight_alertas(s):
            if "🔴" in str(s["Alerta"]):
                return ["background-color: #fce8e6"] * len(s)
            elif "🟢" in str(s["Alerta"]):
                return ["background-color: #e6f4ea"] * len(s)
            return [""] * len(s)

        df_styled = df_exibicao.style.apply(highlight_alertas, axis=1).format({
            "Custo Atual": "R$ {:.2f}",
            "Cotação (R$/kg)": "R$ {:.2f}",
            "Spread R$": "R$ {:.2f}",
            "Var. Cotação": "{:.2f}%",
            "Var. Ref. Sem.": "{:.2f}%",
            "Var. Ref. Mês": "{:.2f}%",
        }, na_rep="-")

        st.dataframe(
            df_styled,
            use_container_width=True,
            column_config={
                "Bandeira": st.column_config.TextColumn(
                    "Bandeira", width="small"
                ),
                "Produto": st.column_config.TextColumn("Produto", width="medium"),
                "Fornecedor": st.column_config.TextColumn(
                    "Fornecedor", width="small"
                ),
                "Custo Atual": st.column_config.NumberColumn(
                    "Custo Atual", width="small"
                ),
                "Cotação (R$/kg)": st.column_config.NumberColumn(
                    "Cotação (R$/kg)", width="small"
                ),
                "Spread R$": st.column_config.NumberColumn(
                    "Spread R$", width="small"
                ),
                "Var. Cotação": st.column_config.NumberColumn(
                    "Var. Cotação", width="small"
                ),
                "Referência": st.column_config.TextColumn(
                    "Referência", width="small"
                ),
                "Var. Ref. Sem.": st.column_config.NumberColumn(
                    "Var. Ref. Sem.", width="small"
                ),
                "Var. Ref. Mês": st.column_config.NumberColumn(
                    "Var. Ref. Mês", width="small"
                ),
                "Alerta": st.column_config.TextColumn("Alerta", width="small"),
            },
        )

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_exibicao.to_excel(
                writer, index=False, sheet_name="Matriz_Negociacao"
            )
        excel_data = output.getvalue()

        st.download_button(
            label="📥 Exportar Matriz de Decisão em Excel",
            data=excel_data,
            file_name=f"Matriz_Negociacao_Cencosud_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("ℹ️ Nenhum item cotado para a rodada e filtros selecionados.")

# ---------------------------------------------------------
# ABA 3: HISTÓRICO TEMPORAL
# ---------------------------------------------------------
with aba_historico:
    st.subheader("📈 Análise Executiva e Tendência Histórica Mensal")

    if CAMINHO_HISTORICO and os.path.exists(CAMINHO_HISTORICO):
        try:
            xls = pd.ExcelFile(CAMINHO_HISTORICO)
            todas_abas = sorted(xls.sheet_names)

            abas_flv = [
                a
                for a in todas_abas
                if a in [
                    "Banana",
                    "Batata",
                    "Cebola",
                    "Maça",
                    "Maca",
                    "Tomate",
                    "Ovos",
                    "Ovo",
                ]
            ]
            abas_proteinas = [a for a in todas_abas if a not in abas_flv]

            col_macro, col_prod = st.columns(2)

            with col_macro:
                macro_sel = st.selectbox(
                    "🔍 Filtrar por Categoria:",
                    options=["Proteínas", "FLV"],
                    key="sel_macro_aba3",
                )

            options_produtos = (
                abas_proteinas if macro_sel == "Proteínas" else abas_flv
            )

            if not options_produtos:
                options_produtos = todas_abas

            with col_prod:
                prod_selecionado = st.selectbox(
                    "🔍 Selecione o Produto:",
                    options=options_produtos,
                    key="sel_prod_aba3",
                )

            df_teste = pd.read_excel(
                CAMINHO_HISTORICO, sheet_name=prod_selecionado, nrows=2
            )
            cols_teste = [str(c).strip() for c in df_teste.columns]

            if "Dia" in cols_teste and "Mês" in cols_teste:
                df_p = pd.read_excel(
                    CAMINHO_HISTORICO, sheet_name=prod_selecionado
                )
                df_p.columns = [str(c).strip() for c in df_p.columns]

                col_preco = next(
                    (
                        c
                        for c in ["Preço k", "Preco k", "PreçoC"]
                        if c in df_p.columns
                    ),
                    df_p.columns[-1],
                )
                df_p[col_preco] = (
                    df_p[col_preco].astype(str).str.replace(",", ".")
                )

                df_p["Data_Tmp"] = pd.to_datetime(
                    df_p["Ano"].astype(str)
                    + "-"
                    + df_p["Mês"].astype(str)
                    + "-"
                    + df_p["Dia"].astype(str),
                    errors="coerce",
                )
                col_data = "Data_Tmp"
                st.info(
                    f"ℹ️ Exibindo cotação de **{prod_selecionado}** (Base"
                    " Hortifruti/FLV em R$/kg)."
                )
            else:
                df_p = pd.read_excel(
                    CAMINHO_HISTORICO, sheet_name=prod_selecionado, skiprows=1
                )
                df_p.columns = [str(c).strip() for c in df_p.columns]

                col_data = next(
                    (c for c in df_p.columns if "data" in c.lower()),
                    df_p.columns[0],
                )
                df_p[col_data] = pd.to_datetime(
                    df_p[col_data], format="%d/%m/%Y", errors="coerce"
                )

                aba_str = str(prod_selecionado).strip().lower()
                if "ovos" in aba_str or "ovo" in aba_str:
                    col_preco = "Branco"
                    st.info(
                        "ℹ️ Exibindo cotação **CIF Região Grande SP** para Ovos"
                        " (FLV)."
                    )
                elif "suino" in aba_str:
                    col_preco = "SP"
                    st.info(
                        "ℹ️ Exibindo cotação oficial para **Suíno Vivo"
                        " (SP)**."
                    )
                elif "boi" in aba_str:
                    col_preco = "À vista R$"
                    st.info(
                        "ℹ️ Exibindo cotação do **Boi Gordo em Reais (R$)**."
                    )
                else:
                    col_preco = "À vista R$"
                    st.info(
                        f"ℹ️ Exibindo cotação de **{prod_selecionado} em Reais"
                        " (R$/kg)**."
                    )

            df_p["Preco_Limpo"] = pd.to_numeric(
                df_p[col_preco], errors="coerce"
            )
            df_p = df_p.dropna(subset=[col_data, "Preco_Limpo"]).sort_values(
                by=col_data
            )

            max_d = df_p[col_data].max()

            df_ultimo_mes = df_p[
                df_p[col_data] >= (max_d - pd.Timedelta(days=30))
            ]
            custo_medio_ultimo_mes = df_ultimo_mes["Preco_Limpo"].mean()

            df_mes_anterior = df_p[
                (df_p[col_data] < (max_d - pd.Timedelta(days=30)))
                & (df_p[col_data] >= (max_d - pd.Timedelta(days=60)))
            ]
            custo_medio_mes_anterior = df_mes_anterior["Preco_Limpo"].mean()

            if (
                pd.notnull(custo_medio_mes_anterior)
                and custo_medio_mes_anterior > 0
            ):
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
                label="Custo Médio Mercado (Último Mês)",
                value=f"R$ {custo_medio_ultimo_mes:,.2f}/kg",
                delta=f"{var_perc_mensal:+.2f}% vs mês anterior",
            )
            k2.metric(
                label=f"Fechamento Sexta-feira ({dt_sexta})",
                value=f"R$ {preco_ultima_sexta:,.2f}/kg",
                delta=f"{var_perc_semanal:+.2f}% vs semana anterior",
            )

            st.divider()
            st.subheader(
                f"📉 Curva Comparativa: Mercado vs Custo Sistema Cencosud -"
                f" {prod_selecionado}"
            )

            df_5y = df_p[
                df_p[col_data] >= (max_d - pd.DateOffset(years=5))
            ].copy()
            df_5y["Periodo"] = df_5y[col_data].dt.to_period("M")
            df_mkt_mes = (
                df_5y.groupby("Periodo")["Preco_Limpo"].mean().reset_index()
            )
            df_mkt_mes["Data_Mensal"] = df_mkt_mes["Periodo"].dt.to_timestamp()
            df_mkt_mes = df_mkt_mes.set_index("Data_Mensal")[["Preco_Limpo"]]
            df_mkt_mes.columns = ["Mercado Oficial (R$/kg)"]

            if os.path.exists(CAMINHO_CUSTO_SISTEMA):
                df_sis = pd.read_excel(CAMINHO_CUSTO_SISTEMA)
                if (
                    not df_sis.empty
                    and "Produto" in df_sis.columns
                    and "Custo_Atual_Sistema" in df_sis.columns
                ):
                    df_sis_prod = df_sis[
                        df_sis["Produto"] == prod_selecionado
                    ].copy()
                    if not df_sis_prod.empty:
                        col_dt_sis = next(
                            (c for c in df_sis_prod.columns if "data" in c.lower()),
                            None,
                        )
                        if col_dt_sis:
                            df_sis_prod["Data_Dt"] = pd.to_datetime(
                                df_sis_prod[col_dt_sis], errors="coerce"
                            )
                        else:
                            df_sis_prod["Data_Dt"] = pd.to_datetime(
                                datetime.date.today()
                            )

                        df_sis_prod["Periodo"] = df_sis_prod[
                            "Data_Dt"
                        ].dt.to_period("M")
                        df_cenc_mes = (
                            df_sis_prod.groupby("Periodo")[
                                "Custo_Atual_Sistema"
                            ]
                            .mean()
                            .reset_index()
                        )
                        df_cenc_mes["Data_Mensal"] = df_cenc_mes[
                            "Periodo"
                        ].dt.to_timestamp()
                        df_cenc_mes = df_cenc_mes.set_index("Data_Mensal")[
                            ["Custo_Atual_Sistema"]
                        ]
                        df_cenc_mes.columns = ["Cencosud - Custo Sistema (R$/kg)"]

                        df_chart_duplo = df_mkt_mes.join(
                            df_cenc_mes, how="left"
                        )
                    else:
                        df_chart_duplo = df_mkt_mes
                else:
                    df_chart_duplo = df_mkt_mes
            else:
                df_chart_duplo = df_mkt_mes

            st.line_chart(df_chart_duplo, use_container_width=True)

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
