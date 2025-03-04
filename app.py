import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO

# Importações para fuzzy matching
from fuzzywuzzy import fuzz as fuzzywuzzy_fuzz
from rapidfuzz import fuzz as rapidfuzz_fuzz

# --- EXTRAÇÃO DE COMPROVANTES ---

def extract_transactions(pdf_document):
    """
    Extrai os dados dos comprovantes de cada página do PDF.
    Retorna uma lista de tuplas (número da página, nome do arquivo) e uma lista de dicionários para o resumo.
    """
    transactions = []
    summary_data = []
    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        text = page.get_text("text")
        # Extração individual usando regex não gananciosa
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

# --- PADRONIZAÇÃO E CONCILIAÇÃO ---

def standardize_data(df, columns):
    """Converte as colunas para minúsculas e remove espaços em branco."""
    for col in columns:
        df[col] = df[col].astype(str).str.lower().str.strip()
    return df

def fuzzy_merge(df_contas, df_comprovantes, method="fuzzywuzzy", threshold=90):
    """
    Realiza a correspondência fuzzy entre a planilha de contas a pagar e os comprovantes.
    Filtra candidatos com base em Valor_std e compara as colunas Empresa e Fornecedor.
    """
    matched_rows = []
    for idx, conta in df_contas.iterrows():
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

def resolve_ambiguous_receipts(df):
    """
    Cria duas novas colunas:
      - "Possível_Cod_Comprovante": igual ao valor atual de "Número do Documento"
      - "Cod_Comprovante": que será preenchido com o código real de vinculação, conforme decisão do usuário.
      
    Se um mesmo código (não nulo) aparecer em várias linhas, é considerada ambiguidade.
    Para cada código ambíguo, o usuário escolhe a conta correta para aquele comprovante.
    Para as linhas que não forem escolhidas, "Cod_Comprovante" fica vazia.
    """
    df = df.copy()
    # Cria a coluna de possível código a partir do merge
    df["Possível_Cod_Comprovante"] = df["Número do Documento"]
    # Inicializa a coluna final igual à possível
    df["Cod_Comprovante"] = df["Possível_Cod_Comprovante"]
    
    # Identifica códigos ambíguos (não nulos) que aparecem em mais de uma linha
    ambig_codes = df["Possível_Cod_Comprovante"].dropna().value_counts()
    ambig_codes = ambig_codes[ambig_codes > 1].index.tolist()
    
    for code in ambig_codes:
        group = df[df["Possível_Cod_Comprovante"] == code]
        st.write(f"Ambiguidade para o comprovante {code}:")
        options = {}
        # Exiba informações úteis para a escolha (por exemplo, Código da conta, Empresa, Fornecedor, Data Vencimento)
        for idx, row in group.iterrows():
            option_str = (f"Código da conta: {row['Código']} | Empresa: {row['Empresa'].title()} | "
                          f"Fornecedor: {row['Fornecedor'].title()} | Data Vencimento: {row['Data Vencimento']}")
            options[option_str] = idx
        chosen_option = st.selectbox(f"Selecione a conta correta para o comprovante {code}:", list(options.keys()), key=f"amb_{code}")
        chosen_idx = options[chosen_option]
        # Para todas as linhas do grupo que NÃO foram escolhidas, zere a coluna Cod_Comprovante
        for idx in group.index:
            if idx != chosen_idx:
                df.at[idx, "Cod_Comprovante"] = ""
    return df

# --- FLUXO PRINCIPAL DO APP ---

st.title("🔎 Conciliação de Pagamentos")

# 1. Upload dos PDFs de comprovantes
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
    # DataFrame dos comprovantes extraídos
    df_comprovantes = pd.DataFrame(all_summary_data)
    st.subheader("Resumo dos Comprovantes Bancários")
    st.dataframe(df_comprovantes)
    csv_comprovantes = df_comprovantes.to_csv(index=False, sep=";").encode()
    st.download_button("Baixar Resumo dos Comprovantes (CSV)",
                       data=csv_comprovantes,
                       file_name="resumo_comprovantes.csv",
                       mime="text/csv",
                       key="download_csv_comprovantes")
    
    # Versão padronizada dos comprovantes
    df_comprovantes_std = df_comprovantes.copy()
    df_comprovantes_std = standardize_data(df_comprovantes_std, ["Empresa", "Fornecedor"])
    df_comprovantes_std["Valor_std"] = df_comprovantes_std["Valor"].round(2)
    
    # 2. Upload da planilha de contas a pagar
    st.subheader("Upload da Planilha de Contas a Pagar")
    contas_file = st.file_uploader("Selecione o arquivo CSV da planilha de Contas a Pagar", type="csv", key="contas")
    
    if contas_file:
        df_contas = pd.read_csv(contas_file, sep=",", dtype=str)
        required_cols = ["Empresa", "Fornecedor", "Data Vencimento", "Valor", "Código"]
        if not all(col in df_contas.columns for col in required_cols):
            st.error("A planilha de contas a pagar deve conter as colunas: Empresa, Fornecedor, Data Vencimento, Valor e Código.")
        else:
            df_contas_std = df_contas.copy()
            df_contas_std = standardize_data(df_contas_std, ["Empresa", "Fornecedor"])
            df_contas_std["Código"] = df_contas_std["Código"].astype(str).str.strip()
            df_contas_std["Valor"] = df_contas_std["Valor"].str.replace(r"r\$\s*", "", regex=True).str.replace(",", ".").astype(float)
            df_contas_std["Valor_std"] = df_contas_std["Valor"].round(2)
            df_contas_std["Data Vencimento"] = pd.to_datetime(df_contas_std["Data Vencimento"], dayfirst=True, errors="coerce")
            st.subheader("Resumo da Planilha de Contas a Pagar")
            st.dataframe(df_contas_std)
            
            # 3. Seleção do método de correspondência
            match_method = st.selectbox("Selecione o método de correspondência:", options=["Padrão", "Fuzzy Wuzzy", "RapidFuzz"])
            if match_method == "Padrão":
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
            
            # Converter "Data da Operação" para datetime (se aplicável)
            df_conciliado["Data da Operação"] = pd.to_datetime(df_conciliado["Data da Operação"], dayfirst=True, errors="coerce")
            df_conciliado["Data_Match"] = df_conciliado.apply(
                lambda row: row["Data Vencimento"] == row["Data da Operação"]
                if pd.notnull(row["Data Vencimento"]) and pd.notnull(row["Data da Operação"]) else False, axis=1
            )
            
            # 4. Na conciliação inicial, crie a coluna "Possível_Cod_Comprovante"
            # Essa coluna será igual a "Número do Documento" (pode haver duplicidade)
            df_conciliado["Possível_Cod_Comprovante"] = df_conciliado["Número do Documento"]
            # Crie também a coluna "Cod_Comprovante", que inicialmente recebe o mesmo valor
            df_conciliado["Cod_Comprovante"] = df_conciliado["Possível_Cod_Comprovante"]
            
            # Se houver ambiguidade (ou seja, o mesmo comprovante para mais de uma conta), resolva:
            if df_conciliado["Possível_Cod_Comprovante"].dropna().duplicated().any():
                st.write("Foram detectadas ambiguidades na vinculação dos comprovantes. Por favor, resolva:")
                df_conciliado_final = resolve_ambiguous_receipts(df_conciliado)
            else:
                df_conciliado_final = df_conciliado.copy()
            
            st.subheader("Tabela de Conciliação Final")
            st.dataframe(df_conciliado_final)
            csv_conciliado = df_conciliado_final.to_csv(index=False, sep=";").encode()
            st.download_button("Baixar Tabela de Conciliação Final (CSV)",
                               data=csv_conciliado,
                               file_name="tabela_conciliada_final.csv",
                               mime="text/csv",
                               key="download_conciliado_final")
            
            # 5. Contas a Pagar sem Conciliação: linhas em que "Cod_Comprovante" está vazio
            df_contas_sem = df_conciliado_final[
                df_conciliado_final["Cod_Comprovante"].isna() |
                (df_conciliado_final["Cod_Comprovante"].astype(str).str.strip() == "")
            ]
            st.subheader("Contas a Pagar sem Conciliação")
            if df_contas_sem.empty:
                st.warning("🚨 Nenhuma conta a pagar sem conciliação encontrada.")
            else:
                st.dataframe(df_contas_sem)
                csv_contas_sem = df_contas_sem.to_csv(index=False, sep=";").encode()
                st.download_button("Baixar Contas a Pagar sem Conciliação (CSV)",
                                   data=csv_contas_sem,
                                   file_name="contas_sem_conciliacao.csv",
                                   mime="text/csv",
                                   key="download_contas_sem")
            
            # 6. Comprovantes sem Conciliação: comprovantes extraídos que não foram vinculados a nenhuma conta
            linked_codes = df_conciliado_final["Cod_Comprovante"].dropna().unique()
            df_receipts_sem = df_comprovantes[~df_comprovantes["Número do Documento"].isin(linked_codes)]
            st.subheader("Comprovantes sem Conciliação")
            st.dataframe(df_receipts_sem)
            csv_receipts_sem = df_receipts_sem.to_csv(index=False, sep=";").encode()
            st.download_button("Baixar Comprovantes sem Conciliação (CSV)",
                               data=csv_receipts_sem,
                               file_name="comprovantes_sem_conciliacao.csv",
                               mime="text/csv",
                               key="download_comprovantes_sem")
            
# 7. Download dos PDFs individuais dos comprovantes
if all_transactions:
    st.subheader("Download dos Comprovantes Individuais (PDF)")
    pdf_index = 0
    for pdf_document, transactions in all_transactions:
        for file_name, pdf_bytes in save_transaction_pdfs(pdf_document, transactions):
            st.download_button(label=f"Baixar {file_name}",
                               data=pdf_bytes,
                               file_name=file_name,
                               mime="application/pdf",
                               key=f"download_pdf_{pdf_index}")
            pdf_index += 1
