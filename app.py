import streamlit as st
import pandas as pd
from io import BytesIO

def process_excel(uploaded_file):
    """Processa o arquivo Excel conforme especificações do usuário"""
    df = pd.read_excel(uploaded_file)

    # 1. Renomear e criar colunas conforme especificado
    df_new = pd.DataFrame()
    df_new["UPC CODE"] = df["UPC"]
    df_new["STYLE"] = df["PRODUCT"]
    df_new["COLOR DESCRIPTION"] = df["UPPER "]
    df_new["SIZE"] = df["SIZE"]
    df_new["Nº DO PO"] = df["P.O. NBR"]
    df_new["QTD"] = df[" RFID LABEL QTY"]
    df_new["COD DA ETQ"] = ""
    df_new["VALOR DO FILTRO"] = 1
    
    # PREFIXO DA EMP: 6 primeiros caracteres da coluna UPC + zero na frente
    df_new["PREFIXO DA EMP"] = "0" + df["UPC"].astype(str).str[:6]
    
    # ITEM DE REF: 6 últimos caracteres excluído o último da coluna UPC + zero na frente
    df_new["ITEM DE REF"] = "0" + df["UPC"].astype(str).str[6:-1]
    
    df_new["SERIAL"] = ""

    # 2. Replicar linhas de acordo com a quantidade da coluna QTD
    df_expanded = df_new.loc[df_new.index.repeat(df_new["QTD"])].reset_index(drop=True)

    # 3. Formatar todas as colunas para texto
    for col in df_expanded.columns:
        df_expanded[col] = df_expanded[col].astype(str)

    return df_expanded

# Interface do Streamlit
st.set_page_config(page_title="Processador de Etiquetas", layout="wide")
st.title("📊 Processador de Planilhas Excel")
st.markdown("""
Esta aplicação processa uma planilha Excel para extrair e transformar dados, gerando uma planilha única pronta para a criação de etiquetas.

**Instruções:**
1. Faça o upload do seu arquivo Excel (formato `.xlsm` ou `.xlsx`).
2. Aguarde o processamento dos dados.
3. Faça o download da planilha gerada.
""")

st.divider()

uploaded_file = st.file_uploader("Escolha um arquivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        st.success("✅ Arquivo carregado com sucesso!")
        
        # Processar o arquivo
        with st.spinner("Processando dados..."):
            df_processed = process_excel(uploaded_file)
        
        st.success(f"✅ Processamento concluído! Total de linhas geradas: {len(df_processed)}")
        
        # Mostrar pré-visualização
        st.write("### Pré-visualização dos dados processados:")
        st.dataframe(df_processed.head(20))
        
        # Gerar arquivo Excel para download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_processed.to_excel(writer, index=False, sheet_name='Sheet1')
        
        output.seek(0)
        
        # Botão de download
        st.download_button(
            label="⬇️ Baixar planilha processada (.xlsx)",
            data=output,
            file_name="planilha_processada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {str(e)}")
        st.write("Verifique se o arquivo possui as colunas necessárias:")
        st.write("- UPC")
        st.write("- PRODUCT")
        st.write("- UPPER ")
        st.write("- SIZE")
        st.write("- P.O. NBR")
        st.write("- RFID LABEL QTY")
else:
    st.info("👆 Faça upload de um arquivo Excel para começar.")
