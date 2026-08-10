import datetime
import io
import os
import pandas as pd
import requests
import streamlit as st

# COLE AQUI A URL GERADA NA IMPLANTAÇÃO DO GOOGLE APPS SCRIPT
URL_WEBHOOK_GOOGLE = "https://script.google.com/macros/s/AKfycbyEWMGiJCp8NRhOcLFGoWJbmQidglG7NTFjXyn_eJbWaqrkiMkf1cxxplcN5Dw5jOzT/exec"
URL_PLANILHA_LEITURA = (
    "https://docs.google.com/spreadsheets/d/1RzJ41JQQPM6ibJrnxmaJNq5JCZiROXAemsaxBvcQV2I/edit?usp=sharing"  # Link normal do Google Sheets
)


def carregar_cotacoes_google():
    """Lê as cotações públicas exportadas em CSV."""
    try:
        if "spreadsheets/d/" in URL_PLANILHA_LEITURA:
            sheet_id = URL_PLANILHA_LEITURA.split("spreadsheets/d/")[1].split(
                "/"
            )[0]
            url_csv = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
            df = pd.read_csv(url_csv)
            return df
    except Exception:
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
    """Envia o novo registro via Webhook seguro para o Google Sheets."""
    try:
        response = requests.post(URL_WEBHOOK_GOOGLE, json=novo_registro)
        if response.status_code == 200:
            st.cache_data.clear()
            return True
        else:
            st.error(f"❌ Resposta do servidor: {response.status_code}")
            return False
    except Exception as e:
        st.error(f"❌ Erro de conexão: {e}")
        return False
