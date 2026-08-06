import datetime
import os
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# CONFIGURAÇÃO DE CAMINHOS E DIRETÓRIO (PASTA NA ÁREA DE TRABALHO)
# ---------------------------------------------------------
USER_PROFILE = os.environ.get('USERPROFILE', '')
PASTA_PROJETO = os.path.join(USER_PROFILE, 'Desktop', 'Projeto Custo')

# Garante que a pasta 'Projeto Custo' exista na Área de Trabalho
if not os.path.exists(PASTA_PROJETO):
    os.makedirs(PASTA_PROJETO)

CAMINHO_HISTORICO = os.path.join(PASTA_PROJETO, 'historico_real.xlsx')
CAMINHO_ENTRADA = os.path.join(PASTA_PROJETO, 'cotacoes_semanais.xlsx')
CAMINHO_SAIDA_CONSOLIDADA = os.path.join(
    PASTA_PROJETO, 'CCI_Matriz_Decisao_Consolidada.xlsx'
)

# ---------------------------------------------------------
# INTERFACE STREAMLIT
# ---------------------------------------------------------
st.set_page_config(
    page_title='Cencosud Commodity Intelligence (CCI)',
    page_icon='🥩',
    layout='wide',
)

st.title('🥩 Cencosud Commodity Intelligence (CCI)')
st.caption(
    'Painel de Entrada de Cotações Semanais e Análise de Margem - Direitório:'
    f' {PASTA_PROJETO}'
)

# Aba inicial para Compradores e Aba de Inteligência
aba_entrada, aba_analise = st.tabs(
    ['📝 Entrada do Comprador (Semanal)', '📊 Matriz de Decisão & Histórico']
)

# Lista Oficial dos 4 Itens Prioritários
PRODUTOS_OFICIAIS = [
    'Filé de Peito',
    'Ovo c/ 20',
    'Alcatra',
    'Contra Filé',
]

UFS_CENCOSUD = ['SP', 'BA', 'PE', 'RJ', 'SE', 'CE', 'GO']

# ---------------------------------------------------------
# ABA 1: FORMULÁRIO INTERATIVO DO COMPRADOR
# ---------------------------------------------------------
with aba_entrada:
    st.subheader('Alimentação Semanal de Preços pelo Comprador')
    st.write(
        'Insira os valores negociados ou cotações de referência da semana e'
        ' clique em **Salvar Cotação**.'
    )

    with st.form(key='form_cotacao_comprador'):
        col1, col2, col3 = st.columns(3)

        with col1:
            produto_sel = st.selectbox(
                'Selecione o Produto:', options=PRODUTOS_OFICIAIS
            )
            uf_sel = st.selectbox('Estado (UF Cencosud):', options=UFS_CENCOSUD)

        with col2:
            preco_mercado = st.number_input(
                'Preço Referência Mercado / CEPEA (R$):',
                min_value=0.0,
                value=10.0,
                step=0.10,
                format='%.2f',
            )
            custo_pago = st.number_input(
                'Custo Pago Cencosud (R$):',
                min_value=0.0,
                value=10.0,
                step=0.10,
                format='%.2f',
            )

        with col3:
            fornecedor = st.text_input(
                'Nome do Fornecedor:', value='Fornecedor Padrão'
            )
            data_referencia = st.date_input(
                'Data da Negociação:', value=datetime.date.today()
            )

        btn_salvar = st.form_submit_button(label='💾 Salvar Cotação na Base')

    if btn_salvar:
        # Estrutura a nova linha
        novo_registro = {
            'Data_Cotacao': data_referencia.strftime('%Y-%m-%d'),
            'Produto': produto_sel,
            'UF_Cencosud': uf_sel,
            'Preco_Mercado_BRL': preco_mercado,
            'Custo_Pago_Cencosud': custo_pago,
            'Fornecedor': fornecedor,
            'Spread_BRL': custo_pago - preco_mercado,
            'Spread_%': (
                ((custo_pago - preco_mercado) / preco_mercado) * 100
                if preco_mercado > 0
                else 0
            ),
        }

        # Carrega dados existentes ou cria novo arquivo
        if os.path.exists(CAMINHO_ENTRADA):
            df_existente = pd.read_excel(CAMINHO_ENTRADA)
            df_atualizado = pd.concat(
                [df_existente, pd.DataFrame([novo_registro])], ignore_index=True
            )
        else:
            df_atualizado = pd.DataFrame([novo_registro])

        # Salva na pasta Projeto Custo na Área de Trabalho
        df_atualizado.to_excel(CAMINHO_ENTRADA, index=False)
        st.success(
            f'✅ Cotação de **{produto_sel}** para **{uf_sel}** salva com sucesso'
            f" em '{CAMINHO_ENTRADA}'!"
        )

    # Exibe tabela das cotações da semana registradas até agora
    if os.path.exists(CAMINHO_ENTRADA):
        st.divider()
        st.subheader('📋 Cotações Salvas na Semana')
        df_semana = pd.read_excel(CAMINHO_ENTRADA)
        st.dataframe(df_semana, use_container_width=True)

# ---------------------------------------------------------
# ABA 2: CRUZAMENTO COM DADOS HISTÓRICOS RETAIL
# ---------------------------------------------------------
with aba_analise:
    st.subheader('Cruzamento com a Base Histórica')

    # Verifica se o arquivo do histórico real baixado por você está na pasta
    if os.path.exists(CAMINHO_HISTORICO):
        df_hist_real = pd.read_excel(CAMINHO_HISTORICO)
        st.success(
            f"✅ Base histórica real carregada com sucesso de '{CAMINHO_HISTORICO}'!"
        )
        st.dataframe(df_hist_real.head(10), use_container_width=True)
    else:
        st.warning(
            f"⚠️ Cole o arquivo Excel do seu histórico real dentro da pasta **'Projeto Custo'** na Área de Trabalho com o nome exato: **'historico_real.xlsx'**."
        )

    if os.path.exists(CAMINHO_ENTRADA) and os.path.exists(CAMINHO_HISTORICO):
        if st.button('🔄 Gerar Consolidação Completa do CCI'):
            df_cot = pd.read_excel(CAMINHO_ENTRADA)
            df_hist = pd.read_excel(CAMINHO_HISTORICO)

            with pd.ExcelWriter(
                CAMINHO_SAIDA_CONSOLIDADA, engine='openpyxl'
            ) as writer:
                df_cot.to_excel(
                    writer, sheet_name='Cotacoes_Comprador', index=False
                )
                df_hist.to_excel(
                    writer, sheet_name='Historico_Real', index=False
                )

            st.balloons()
            st.success(
                f'🎉 Relatório Consolidado salvo em: {CAMINHO_SAIDA_CONSOLIDADA}'
            )