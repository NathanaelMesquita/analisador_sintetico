import pandas as pd
import streamlit as st
from io import BytesIO

# Função para criar a planilha Excel com o conteúdo das sheets
def create_excel_report(file1, file2):
    # Carrega as planilhas do usuário
    excel_file1 = pd.ExcelFile(file1)
    excel_file2 = pd.ExcelFile(file2)

    # Cria um buffer de bytes para armazenar o arquivo Excel na memória
    output = BytesIO()
    
    # Cria um objeto ExcelWriter para salvar múltiplas sheets no arquivo
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Itera sobre as sheets dos dois arquivos e adiciona ao writer
        for idx, excel_file in enumerate([excel_file1, excel_file2], start=1):
            for sheet_name in excel_file.sheet_names:
                # Carrega os dados da sheet em um DataFrame
                df = excel_file.parse(sheet_name)

                # Escreve cada DataFrame como uma nova sheet no arquivo Excel
                writer.sheet_names = excel_file.sheet_names
                sheet_title = f"{sheet_name}_file{idx}"
                df.to_excel(writer, sheet_name=sheet_title, index=False)

    # Reposiciona o ponteiro do buffer para o início
    output.seek(0)
    
    return output

# Interface do Streamlit
st.title("Gerador de Relatório de Análise em Excel")
st.write("Faça o upload de dois arquivos Excel para gerar um novo arquivo com as sheets combinadas.")

# Upload dos arquivos
file1 = st.file_uploader("Upload do primeiro arquivo Excel", type="xlsx")
file2 = st.file_uploader("Upload do segundo arquivo Excel", type="xlsx")

# Botão para gerar e baixar o relatório
if file1 and file2:
    output = create_excel_report(file1, file2)
    
    st.download_button(
        label="Baixar Relatório de Análise",
        data=output,
        file_name="Relatorio_Analise.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
