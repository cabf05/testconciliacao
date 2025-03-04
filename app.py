import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO

# Função para extrair transações do PDF
def extract_transactions(pdf_document):
    transactions = []
    summary_data = []
    
    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        text = page.get_text("text")

        # Regex para capturar as informações principais da transação
        match = re.search(
            r"Empresa:\s*(.*?)\s*\|\s*CNPJ: .*?\n"
            r"Nome do favorecido:\s*(.*?)\n"
            r".*?Data da operação:\s*(\d{2}/\d{2}/\d{4}) - \d{2}h\d{2}\n"
            r"N° de controle:\s*(\d+)\s*\|",  # Captura o número do documento (comprovante)
            text, re.DOTALL
        )
        
        match_valor = re.search(r"Valor\s*R\$\s*([\d,.]+)", text)  # Captura o valor da transação
        
        if match and match_valor:
            pagador = match.group(1).strip()
            favorecido = match.group(2).strip()
            data = match.group(3)
            numero_documento = match.group(4)  # Número do comprovante
            valor = match_valor.group(1).replace(",", ".")

            file_name = f"{pagador.replace(' ', '_')}_para_{favorecido.replace(' ', '_')}_{data}_R${valor}.pdf"
            transactions.append((page_num, file_name))

            summary_data.append({
                "Empresa": pagador,
                "Fornecedor": favorecido,
                "Data da Operação": data,
                "Valor": float(valor.replace("R$ ", "").replace(",", ".")),  # Convertendo para número
                "Número do Documento": numero_documento,
                "Arquivo PDF": file_name
            })

    return transactions, summary_data

# Interface com Streamlit
st.title("🔎 Sistema de Conciliação de Pagamentos")
st.subheader("Faça upload de um ou mais arquivos PDF contendo comprovantes.")

uploaded_files = st.file_uploader("Selecione os arquivos", type="pdf", accept_multiple_files=True)

if uploaded_files:
    all_transactions = []
    all_summary_data = []
    
    for uploaded_file in uploaded_files:
        st.write(f"📄 Processando: {uploaded_file.name}...")
        
        pdf_document = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        transactions, summary_data = extract_transactions(pdf_document)
        
        if transactions:
            st.success(f"✅ {len(transactions)} comprovante(s) encontrado(s) em {uploaded_file.name}.")
            all_transactions.append((pdf_document, transactions))
            all_summary_data.extend(summary_data)
        else:
            st.warning(f"⚠ Nenhum comprovante encontrado em {uploaded_file.name}.")
    
    if all_summary_data:
        df_comprovantes = pd.DataFrame(all_summary_data)
        
        st.subheader("📊 Resumo das Transações Bancárias")
        st.dataframe(df_comprovantes)  # Exibe a tabela
        
        csv_comprovantes = df_comprovantes.to_csv(index=False, sep=";").encode()
        st.download_button("📥 Baixar Tabela Resumo (CSV)", data=csv_comprovantes, file_name="resumo_transacoes.csv", mime="text/csv")

        # Upload da planilha de contas a pagar
        st.subheader("📂 Faça upload da planilha de Contas a Pagar")
        contas_pagar_file = st.file_uploader("Selecione o arquivo CSV", type="csv")
        
        if contas_pagar_file:
            df_contas_pagar = pd.read_csv(contas_pagar_file, sep=";", dtype=str)
            
            # Convertendo valores para float para facilitar a comparação
            df_contas_pagar["Valor"] = df_contas_pagar["Valor"].str.replace("R$ ", "").str.replace(",", ".").astype(float)

            st.subheader("📋 Resumo da Planilha de Contas a Pagar")
            st.dataframe(df_contas_pagar)

            # Fazer a conciliação
            df_conciliado = df_contas_pagar.merge(df_comprovantes, on=["Empresa", "Fornecedor", "Valor"], how="left")
            
            # Criar lista de registros sem correspondência
            df_contas_sem_pagamento = df_conciliado[df_conciliado["Número do Documento"].isna()]
            df_pagamentos_sem_conta = df_comprovantes[~df_comprovantes["Número do Documento"].isin(df_conciliado["Número do Documento"])]

            st.subheader("✅ Tabela Conciliada")
            st.dataframe(df_conciliado)

            csv_conciliado = df_conciliado.to_csv(index=False, sep=";").encode()
            st.download_button("📥 Baixar Tabela Conciliada (CSV)", data=csv_conciliado, file_name="tabela_conciliada.csv", mime="text/csv")

            st.subheader("⚠ Contas a Pagar SEM Comprovante")
            st.dataframe(df_contas_sem_pagamento)

            csv_contas_sem = df_contas_sem_pagamento.to_csv(index=False, sep=";").encode()
            st.download_button("📥 Baixar Contas SEM Comprovante (CSV)", data=csv_contas_sem, file_name="contas_sem_comprovante.csv", mime="text/csv")

            st.subheader("⚠ Pagamentos SEM Correspondência com Contas a Pagar")
            st.dataframe(df_pagamentos_sem_conta)

            csv_pagamentos_sem = df_pagamentos_sem_conta.to_csv(index=False, sep=";").encode()
            st.download_button("📥 Baixar Pagamentos SEM Contas a Pagar (CSV)", data=csv_pagamentos_sem, file_name="pagamentos_sem_conta.csv", mime="text/csv")
