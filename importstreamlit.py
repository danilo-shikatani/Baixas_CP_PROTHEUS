import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Conversor de arquivo para baixa contas a pagar PROTHEUS", layout="centered")

st.title("📄 Conversor de Excel para baixas via planilha Contas a pagar protheus")
st.markdown("Envie um arquivo Excel Funçoes contas a pagar para gerar o arquivo .CSV ou .TXT formatado.")

# Upload do arquivo
arquivo = st.file_uploader("Selecione o arquivo Excel (.xlsx)", type=["xlsx"])

# Parâmetros fixos
with st.form("parametros"):
    st.subheader("🔧 Parâmetros fixos")
    dt_baixa = st.text_input("Data de Baixa", "02/07/2025")
    motivo = st.text_input("Motivo de baixa", "DEB")
    banco = st.text_input("Banco", "033")
    agencia = st.text_input("Agência", "3409")
    conta = st.text_input("Conta", "130067894")
    historico = st.text_input("Histórico", "BX MANUAL TXT")
    processar = st.form_submit_button("✅ Processar arquivo")

# Processamento
if processar and arquivo is not None:
    try:

        

        df = pd.read_excel(arquivo, skiprows=1, engine='openpyxl',
        dtype={'No. Titulo': str, 'Parcela': str, 'Filial': str})
        df = df.astype(str).replace({'nan': '', 'NaN': '', 'None': ''})
        df.columns = df.columns.str.strip()
        
        df_selecionado = df[['Filial', 'Prefixo', 'No. Titulo', 'Parcela', 'Tipo', 'Fornecedor', 'Loja']].copy()
        df_selecionado.columns = ['E1_FILIAL','E1_PREFIXO','E1_NUM','E1_PARCELA','E1_TIPO','E1_CLIENTE','E1_LOJA']

        df_selecionado['E1_FILIAL'] = df_selecionado['E1_FILIAL'].str[:4]
        df_selecionado['E1_NUM'] = df_selecionado['E1_NUM'].str.replace('.0', '', regex=False).str.zfill(9)
        df_selecionado['E1_PREFIXO'] = df_selecionado['E1_PREFIXO'].str.replace('.0', '', regex=False).str.zfill(3) 
        df_selecionado['E1_PARCELA'] = df_selecionado['E1_PARCELA'].str.replace('.0', '', regex=False).str.zfill(2)
        
        df_selecionado['DT_BAIXA'] = dt_baixa
        df_selecionado['MOTIVO'] = motivo
        df_selecionado['BANCO'] = banco
        df_selecionado['AGENCIA'] = agencia
        df_selecionado['CONTA'] = conta
        df_selecionado['HISTORICO'] = historico
        
        df_selecionado.replace({'nan': '', 'NaN': '', 'None': ''}, inplace=True)
        
        txt = df_selecionado.to_csv(index=False, sep='|', header=False).encode('latin1')

        st.success("✅ Arquivo processado com sucesso!")
        st.download_button(
        label="⬇️ Baixar TXT formatado",
        data=txt,
        file_name="resultado.txt",
        mime="text/plain"
)

    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")
