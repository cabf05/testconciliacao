import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO

# Importações para fuzzy matching
from fuzzywuzzy import fuzz as fuzzywuzzy_fuzz
from rapidfuzz import fuzz as rapidfuzz_fuzz

# --- FUNÇÕES DE EXTRAÇÃO E SALVAMENTO DE COMPROVANTES ---

def extract_transactions(pdf_document):
    """
    Extrai os dados dos comprovantes de cada página do PDF.
    Retorna uma lista de tuplas (número da página, nome do arquivo) e 
    uma lista de dicionários para o resumo.
    A extração é feita com buscas individuais para cada campo, tornando o processo mais robusto.
    """
    transactions = []
    summary_data = []
    
    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        text = page.get_text("text")
        
        # Extração individual dos campos usando padrões não gananciosos:
        data_operacao_match = re.search(r"Data da operação:\s*(\d{2}/\d{2}/\d{4})", text)
        documento_match = re.search(r"Documento:\s*(\d+)", text)
        empresa_match = re.search(r"Empresa:\s*(.*?)\s*\|", text)
        favorecido_match = re.search(r"Nome do favorecido:\s*(.*?)\n", text)
        valor_match = re.search(r"Valor\s*R\$\s*([\d.,]+)", text)
        
        if data_operacao_match and documento_match and empresa_match and favorecido_match and valor_match:
            data_operacao = data_operacao_match.group(1).strip()
            numero_documento = documento_match.group(1).strip()
            empresa = empresa_match.group(1).strip()
            fornecedor = favorecido_match.group(1).strip()
            # Remove separador de milhar e troca vírgula decimal
            valor_str = valor_match.group(1).strip().replace(".", "").replace(",", ".")
            try:
                valor = float(valor_str)
            except Exception:
                valor = 0.0
            
            file_name = f"{empresa.replace(' ', '_')}_para_{fornecedor.replace(' ', '_')}_{data_operacao.replace('/', '-')}_R${valor:.2f}.pdf"
            
            transactions.append((page_num, file_name))
            summary_data.append({
                "Empresa": empresa,
                "Fornecedor": fornecedor,
                "Data da Operação": data_operacao,
                "Valor": valor,
                "Número do Documento": numero_documento,
                "Arquivo PDF": file_name
            })
    
    return transactions, summary_data

def save_transaction_pdfs(pdf_document, transactions):
    """
    Salva cada página (comprovante) em um novo PDF mantendo o layout original.
    Retorna uma lista de tuplas (nome do arquivo, bytes do PDF).
    """
    files = []
    for page_num, file_name in transactions:
        pdf_writer = fitz.open()
        pdf_writer.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)
        pdf_bytes = pdf_writer.write()
        files.append((file_name, pdf_bytes))
    return files

# --- FUNÇÕES PARA PADRONIZAÇÃO E CONCILIAÇÃO ---

def standardize_data(df, columns):
    """Padroniza as colunas: converte para texto em minúsculo e remove espaços extras."""
    for col in columns:
        df[col] = df[col].astype(str).str.lower().str.strip()
    return df

def fuzzy_merge(df_contas, df_comprovantes, method="fuzzywuzzy", threshold=90):
    """
    Realiza a correspondência fuzzy entre a planilha de contas a pagar e os comprovantes.
    Filtra os candidatos com base no valor (Valor_std) e compara as colunas Empresa e Fornecedor.
    Retorna um DataFrame com os dados da conta mais os dados do comprovante (quando a similaridade for ≥ threshold).
    """
    matched_rows = []
    for idx, conta in df_contas.iterrows():
        # Filtra os candidatos que tenham o mesmo valor (Valor_std)
        candidates = df_comprovantes[df_comprovantes["Valor_std"] == conta["Valor_std"]]
        if candidates.empty:
            row = conta.to_dict()
            row.update({"Número do Documento": None, "Data da Operação": None, "Arquivo PDF": None, "Fuzzy Score": None})
            matched_rows.append(row)
        else:
            found_match = False
            for j, comp in candidates.iterrows():
                if method == "fuzzywuzzy":
                    score_empresa = fuzzywuzzy_fuzz.token_set_ratio(conta["Empresa"], comp["Empresa"])
                    score_fornecedor = fuzzywuzzy_fuzz.token_set_ratio(conta["Fornecedor"], comp["Fornecedor"])
                elif method == "rapidfuzz":
                    score_empresa = rapidfuzz_fuzz.token_set_ratio(conta["Empresa"], comp["Empresa"])
                    score_fornecedor = rapidfuzz_fuzz.token_set_ratio(conta["Fornecedor"], comp["Fornecedor"])
                score = (score_empresa + score_fornecedor) / 2
                if score >= threshold:
                    found_match = True
                    row = conta.to_dict()
                    row.update({
                        "Número do Documento": comp["Número do Documento"],
                        "Data da Operação": comp["Data da Operação"],
                        "Arquivo PDF": comp["Arquivo PDF"],
                        "Fuzzy Score": score
                    })
                    matched_rows.append(row)
            if not found_match:
                row = conta.to_dict()
                row.update({"Número do Documento": None, "Data da Operação": None, "Arquivo PDF": None, "Fuzzy Score": None})
                matched_rows.append(row)
    return pd.DataFrame(matched_rows)

def resolve_duplicates(df):
    """
    Para contas a pagar que possuem mais de um comprovante vinculado, 
    exibe uma interface interativa para o usuário escolher qual opção manter.
    Considera 'Código' como identificador único da conta a pagar.
    """
    resolved_rows = []
    grouped = df.groupby('Código')
    for codigo, group in grouped:
        if len(group) == 1:
            resolved_rows.append(group)
        else:
            st.write(f"Para a conta a pagar com Código {codigo}, foram encontrados múltiplos comprovantes:")
            options = {}
            for idx, row in group.iterrows():
                option_str = f"Doc: {row['Número do Documento']} | Arquivo: {row['Arquivo PDF']} | Data Operação: {row['Data da Operação']}"
                options[option_str] = idx
            chosen = st.selectbox(f"Selecione o comprovante para a conta com Código {codigo}:", list(options.keys()), key=f"select_{codigo}")
            chosen_idx = options[chosen]
            resolved_rows.append(group.loc[[chosen_idx]])
    if resolved_rows:
        resolved_df = pd.concat(resolved_rows, ignore_index=True)
    else:
        resolved_df = pd.DataFrame()
    return resolved_df

# --- INTERFACE DO APLICATIVO ---

st.title("🔎 Sistema de Conciliação de Pagamentos - Versão com Múltiplos Métodos de Correspondência")

# 1. Upload dos PDFs com comprovantes
st.subheader("Upload de PDFs de Comprovantes Bancários")
uploaded_files = st.file_uploader("Selecione um ou mais arquivos PDF", type="pdf", accept_multiple_files=True)

all_transactions = []
all_summary_data = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.write(f"Processando: {uploaded_file.name} ...")
        pdf_document = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        transactions, summary_data = extract_transactions(pdf_document)
        if transactions:
            st.success(f"{len(transactions)} comprovante(s) encontrados em {uploaded_file.name}.")
            all_transactions.append((pdf_document, transactions))
            all_summary_data.extend(summary_data)
        else:
            st.warning(f"Nenhum comprovante encontrado em {uploaded_file.name}.")

if all_summary_data:
    df_comprovantes = pd.DataFrame(all_summary_data)
    st.subheader("Resumo dos Comprovantes Bancários")
    st.dataframe(df_comprovantes)
    csv_comprovantes = df_comprovantes.to_csv(index=False, sep=";").encode()
    st.download_button(
        "Baixar Resumo dos Comprovantes (CSV)",
        data=csv_comprovantes,
        file_name="resumo_comprovantes.csv",
        mime="text/csv",
        key="download_csv_comprovantes"
    )

    # 2. Upload da planilha de contas a pagar
    st.subheader("Upload da Planilha de Contas a Pagar")
    contas_file = st.file_uploader("Selecione o arquivo CSV da planilha de Contas a Pagar", type="csv", key="contas")
    
    if contas_file:
        # Se o CSV usar vírgula como separador (como no seu arquivo), ajuste o sep para ","
        df_contas = pd.read_csv(contas_file, sep=",", dtype=str)
        # Verifica as colunas obrigatórias
        required_cols = ["Empresa", "Fornecedor", "Data Vencimento", "Valor", "Código"]
        if not all(col in df_contas.columns for col in required_cols):
            st.error("A planilha de contas a pagar deve conter as colunas: Empresa, Fornecedor, Data Vencimento, Valor e Código.")
        else:
            # Padroniza e converte os dados para facilitar a conciliação
            df_contas_std = df_contas.copy()
            df_contas_std = standardize_data(df_contas_std, ["Empresa", "Fornecedor"])
            # Garanta que a coluna 'Código' não possua espaços extras
            df_contas_std["Código"] = df_contas_std["Código"].astype(str).str.strip()
            df_contas_std["Valor"] = df_contas_std["Valor"].str.replace(r"r\$\s*", "", regex=True).str.replace(",", ".").astype(float)
            df_contas_std["Valor_std"] = df_contas_std["Valor"].round(2)
            
            st.subheader("Resumo da Planilha de Contas a Pagar")
            st.dataframe(df_contas_std)

            # 3. Escolha do método de correspondência
            match_method = st.selectbox("Selecione o método de correspondência:", options=["Padrão", "Fuzzy Wuzzy", "RapidFuzz"])
            
            if match_method == "Padrão":
                # Utilize os nomes das colunas conforme estão: "Empresa", "Fornecedor" e "Valor_std"
                df_conciliado = pd.merge(
                    df_contas_std,
                    df_comprovantes_std,
                    left_on=["Empresa", "Fornecedor", "Valor_std"],
                    right_on=["Empresa", "Fornecedor", "Valor_std"],
                    how="left",
                    suffixes=("_conta", "_comprovante")
                )
            else:
                threshold = st.slider("Defina o limiar para correspondência fuzzy:", min_value=50, max_value=100, value=90)
                if match_method == "Fuzzy Wuzzy":
                    df_conciliado = fuzzy_merge(df_contas_std, df_comprovantes_std, method="fuzzywuzzy", threshold=threshold)
                elif match_method == "RapidFuzz":
                    df_conciliado = fuzzy_merge(df_contas_std, df_comprovantes_std, method="rapidfuzz", threshold=threshold)
            
            st.subheader("Tabela Conciliada Inicial")
            st.dataframe(df_conciliado)
            csv_conciliado = df_conciliado.to_csv(index=False, sep=";").encode()
            st.download_button(
                "Baixar Tabela Conciliada Inicial (CSV)",
                data=csv_conciliado,
                file_name="tabela_conciliada_inicial.csv",
                mime="text/csv",
                key="download_conciliado_inicial"
            )
            
            # 4. Resolução de duplicidades (se houver mais de um comprovante para o mesmo 'Código')
            duplicate_mask = df_conciliado.duplicated(subset=["Código"], keep=False)
            df_duplicates = df_conciliado[duplicate_mask].sort_values("Código")
            if not df_duplicates.empty:
                st.subheader("Resolver Duplicidades")
                df_resolvido = resolve_duplicates(df_duplicates)
            else:
                df_resolvido = df_conciliado.copy()
            
            st.subheader("Tabela Conciliada Final (após resolução)")
            st.dataframe(df_resolvido)
            csv_conciliado_final = df_resolvido.to_csv(index=False, sep=";").encode()
            st.download_button(
                "Baixar Tabela Conciliada Final (CSV)",
                data=csv_conciliado_final,
                file_name="tabela_conciliada_final.csv",
                mime="text/csv",
                key="download_conciliado_final"
            )
            
            # 5. Listar contas a pagar sem comprovante e comprovantes sem vínculo
            df_contas_sem_comprovante = df_conciliado[df_conciliado["Número do Documento"].isna()]
            st.subheader("Contas a Pagar SEM Comprovante")
            st.dataframe(df_contas_sem_comprovante)
            csv_contas_sem = df_contas_sem_comprovante.to_csv(index=False, sep=";").encode()
            st.download_button(
                "Baixar Contas SEM Comprovante (CSV)",
                data=csv_contas_sem,
                file_name="contas_sem_comprovante.csv",
                mime="text/csv",
                key="download_contas_sem"
            )
            
            linked_doc_numbers = df_resolvido["Número do Documento"].dropna().unique()
            df_receipts_sem_conta = df_comprovantes[~df_comprovantes["Número do Documento"].isin(linked_doc_numbers)]
            st.subheader("Comprovantes SEM Correspondência com Contas a Pagar")
            st.dataframe(df_receipts_sem_conta)
            csv_receipts_sem = df_receipts_sem_conta.to_csv(index=False, sep=";").encode()
            st.download_button(
                "Baixar Comprovantes SEM Contas (CSV)",
                data=csv_receipts_sem,
                file_name="comprovantes_sem_conta.csv",
                mime="text/csv",
                key="download_comprovantes_sem_conta"
            )

# 6. Download dos PDFs individuais dos comprovantes
if all_transactions:
    st.subheader("Download dos Comprovantes Individuais (PDF)")
    pdf_index = 0
    for pdf_document, transactions in all_transactions:
        for file_name, pdf_bytes in save_transaction_pdfs(pdf_document, transactions):
            st.download_button(
                label=f"Baixar {file_name}",
                data=pdf_bytes,
                file_name=file_name,
                mime="application/pdf",
                key=f"download_pdf_{pdf_index}"
            )
            pdf_index += 1
