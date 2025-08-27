import streamlit as st
import pandas as pd
import unicodedata
import re

st.set_page_config(page_title="Conversor de arquivo para baixa contas a pagar PROTHEUS", layout="centered")

st.title("📄 Conversor de Excel para baixas via planilha Contas a pagar Protheus")
st.markdown("Envie um arquivo Excel de **Funções Contas a Pagar** para gerar o arquivo `.TXT` formatado.")

# ---------- Utilidades ----------
def normalize_label(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.lower().strip()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^a-z0-9 ]", "", s)
    s = s.replace(" ", "")
    return s

# sinônimos possíveis (normalizados) -> papel esperado
SINONIMOS = {
    # Filial
    "filial": "filial",
    "idfilial": "filial",
    "e1filial": "filial",
    "empresa": "filial",

    # Prefixo
    "prefixo": "prefixo",
    "e1prefixo": "prefixo",
    "serie": "prefixo",

    # Número do título
    "notitulo": "num_titulo",
    "notitulo": "num_titulo",
    "numerotitulo": "num_titulo",
    "numerodotitulo": "num_titulo",
    "titulo": "num_titulo",
    "documento": "num_titulo",
    "e1num": "num_titulo",
    "nrtitulo": "num_titulo",
    "nrotitulo": "num_titulo",

    # Parcela
    "parcela": "parcela",
    "e1parcela": "parcela",

    # Tipo
    "tipo": "tipo",
    "tipotitulo": "tipo",
    "e1tipo": "tipo",

    # Fornecedor
    "fornecedor": "fornecedor",
    "codfornecedor": "fornecedor",
    "idfornecedor": "fornecedor",
    "fornec": "fornecedor",
    "cnpjfornecedor": "fornecedor",

    # Loja
    "loja": "loja",
    "lojafornecedor": "loja",
    "lojaFornecedor": "loja",
    "e1loja": "loja",
}

PAPEIS = ["filial", "prefixo", "num_titulo", "parcela", "tipo", "fornecedor", "loja"]

# Dicionário dos motivos para o selectbox
motivos_dict = {
    "NOR": "NORMAL",
    "DAC": "DACAO",
    "DEB": "DEBITO CC",
    "LIQ": "LIQUIDACAO",
    "CEC": "COMP CARTE",
    "FAT": "FATURAS",
    "RES": "RESIDUO",
    "CAN": "CANCELAMEN",
    "STP": "SUBSTPR",
    "CMP": "COMPENSACA",
    "CNF": "CANCELA NF",
    "LOJ": "OUTRA LOJA",
    "BFT": "BAIXA FAT.",
    "TRO": "TROCO",
    "MPR": "MAIS PRAZO",
    "OFF": "+NEGOCIOS",
    "DIS": "DISTRATO",
    "CDD": "CESS.DIREI",
    "PIX": "PIX_MANUAL",
    "SER": "SERASA",
    "PER": "PERDA",
}

# ---------- Upload ----------
arquivo = st.file_uploader("Selecione o arquivo Excel (.xlsx)", type=["xlsx"])

with st.form("parametros"):
    st.subheader("🔧 Parâmetros fixos")
    header_row = st.number_input("Linha do cabeçalho (1 = primeira linha da planilha)", min_value=1, max_value=200, value=1, step=1)

    dt_baixa = st.text_input("Data de Baixa", "02/07/2025")

    motivo_descricao = st.selectbox(
        "Motivo de baixa",
        options=[f"{k} - {v}" for k, v in motivos_dict.items()],
        index=list(motivos_dict.keys()).index("DEB")
    )
    motivo = motivo_descricao.split(" - ")[0]

    banco = st.text_input("Banco", "033")
    agencia = st.text_input("Agência", "3409")
    conta = st.text_input("Conta", "130067894")
    historico = st.text_input("Histórico", "BX MANUAL TXT")

    processar = st.form_submit_button("✅ Processar arquivo")

if processar and arquivo is None:
    st.error("Envie um arquivo antes de processar.")

# ---------- Processamento ----------
if processar and arquivo is not None:
    try:
        # Lê a planilha usando a linha informada como cabeçalho
        df = pd.read_excel(arquivo, header=header_row - 1, engine="openpyxl")
        df.columns = [c.strip() for c in df.columns.astype(str)]
        st.caption("🧭 Colunas detectadas:")
        st.write(list(df.columns))

        # Tenta mapear automaticamente
        col_norm = {c: normalize_label(c) for c in df.columns}
        guess_map = {papel: None for papel in PAPEIS}
        for original, norm in col_norm.items():
            papel = SINONIMOS.get(norm)
            if papel and guess_map.get(papel) is None:
                guess_map[papel] = original

        st.divider()
        st.subheader("🧩 Mapeamento de colunas (ajuste se necessário)")

        # Permite ajuste manual via selects, pré-preenchidos com o chute
        cols_disp = ["-- selecionar --"] + list(df.columns)
        sel_filial     = st.selectbox("Coluna Filial", cols_disp, index=(cols_disp.index(guess_map["filial"]) if guess_map["filial"] in cols_disp else 0))
        sel_prefixo    = st.selectbox("Coluna Prefixo", cols_disp, index=(cols_disp.index(guess_map["prefixo"]) if guess_map["prefixo"] in cols_disp else 0))
        sel_numtitulo  = st.selectbox("Coluna Nº do Título", cols_disp, index=(cols_disp.index(guess_map["num_titulo"]) if guess_map["num_titulo"] in cols_disp else 0))
        sel_parcela    = st.selectbox("Coluna Parcela", cols_disp, index=(cols_disp.index(guess_map["parcela"]) if guess_map["parcela"] in cols_disp else 0))
        sel_tipo       = st.selectbox("Coluna Tipo", cols_disp, index=(cols_disp.index(guess_map["tipo"]) if guess_map["tipo"] in cols_disp else 0))
        sel_fornecedor = st.selectbox("Coluna Fornecedor", cols_disp, index=(cols_disp.index(guess_map["fornecedor"]) if guess_map["fornecedor"] in cols_disp else 0))
        sel_loja       = st.selectbox("Coluna Loja", cols_disp, index=(cols_disp.index(guess_map["loja"]) if guess_map["loja"] in cols_disp else 0))

        selecionadas = {
            "E1_FILIAL": sel_filial if sel_filial != "-- selecionar --" else None,
            "E1_PREFIXO": sel_prefixo if sel_prefixo != "-- selecionar --" else None,
            "E1_NUM": sel_numtitulo if sel_numtitulo != "-- selecionar --" else None,
            "E1_PARCELA": sel_parcela if sel_parcela != "-- selecionar --" else None,
            "E1_TIPO": sel_tipo if sel_tipo != "-- selecionar --" else None,
            "E1_CLIENTE": sel_fornecedor if sel_fornecedor != "-- selecionar --" else None,
            "E1_LOJA": sel_loja if sel_loja != "-- selecionar --" else None,
        }

        faltando = [k for k, v in selecionadas.items() if v is None]
        if faltando:
            st.error(f"Mapeie todas as colunas obrigatórias. Faltando: {', '.join(faltando)}")
            st.stop()

        # Seleciona/renomeia
        df_sel = df[[selecionadas["E1_FILIAL"], selecionadas["E1_PREFIXO"], selecionadas["E1_NUM"],
                     selecionadas["E1_PARCELA"], selecionadas["E1_TIPO"], selecionadas["E1_CLIENTE"],
                     selecionadas["E1_LOJA"]]].copy()
        df_sel.columns = ["E1_FILIAL","E1_PREFIXO","E1_NUM","E1_PARCELA","E1_TIPO","E1_CLIENTE","E1_LOJA"]

        # Normalizações esperadas
        df_sel = df_sel.astype(str).replace({'nan': '', 'NaN': '', 'None': ''})

        df_sel["E1_FILIAL"]   = df_sel["E1_FILIAL"].str.strip().str[:4]
        df_sel["E1_NUM"]      = df_sel["E1_NUM"].str.replace(".0", "", regex=False).str.replace(",", "", regex=False).str.strip().str.zfill(9)
        df_sel["E1_PREFIXO"]  = df_sel["E1_PREFIXO"].str.replace(".0", "", regex=False).str.strip().str.zfill(3)
        df_sel["E1_PARCELA"]  = df_sel["E1_PARCELA"].str.replace(".0", "", regex=False).str.strip().str.zfill(2)
        df_sel["E1_TIPO"]     = df_sel["E1_TIPO"].str.strip()
        df_sel["E1_CLIENTE"]  = df_sel["E1_CLIENTE"].str.strip()
        df_sel["E1_LOJA"]     = df_sel["E1_LOJA"].str.replace(".0", "", regex=False).str.strip().str.zfill(2)

        # Parâmetros fixos
        df_sel["DT_BAIXA"]  = dt_baixa
        df_sel["MOTIVO"]    = motivo
        df_sel["BANCO"]     = banco
        df_sel["AGENCIA"]   = agencia
        df_sel["CONTA"]     = conta
        df_sel["HISTORICO"] = historico

        # Exporta TXT
        txt = df_sel.to_csv(index=False, sep="|", header=False).encode("latin1", errors="replace")

        st.success("✅ Arquivo processado com sucesso!")
        st.download_button(
            label="⬇️ Baixar TXT formatado",
            data=txt,
            file_name="resultado.txt",
            mime="text/plain"
        )

        with st.expander("Pré-visualização dos primeiros registros"):
            st.dataframe(df_sel.head(50))

    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")

