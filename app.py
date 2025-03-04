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
    Retorna uma lista de tuplas (número da página, nome do arquivo) e uma lista de dicionários para o resumo.
    A extração é feita com buscas individuais para cada campo, tornando o processo mais robusto.
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

# --- FUNÇÕES DE PADRONIZAÇÃO E CONCILIAÇÃO ---

def standardize_data(df, columns):
    """Padroniza as colunas: converte para texto em minúsculo e remove espaços extras."""
    for col in columns:
        df[col] = df[col].astype(str).str.lower().str.strip()
    return df

def fuzzy_merge(df_contas, df_comprovantes, method="fuzzywuzzy", threshold=90):
    """
    Realiza a correspondência fuzzy entre a planilha de contas a pagar e os comprovantes.
    Compara as colunas Empresa e Fornecedor e filtra candidatos com base no Valor_std.
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

def resolve_ambiguities(df):
    """
    Resolve ambiguidades:
    - Para contas a pagar (identificadas por "Código") que receberam mais de um comprovante.
    - Para comprovantes (identificados por "Número do Documento") associados a várias contas.
    O usuário deverá escolher a correspondência correta para cada grupo ambíguo.
    """
    resolved_rows = []
    # Resolver por "Código" (conta com múltiplos comprovantes)
    grouped_codigo = df.groupby("Código")
    for codigo, group in grouped_codigo:
        if len(group) == 1:
            resolved_rows.append(group)
        else:
            st.write(f"Ambiguidade para a conta com Código {codigo}:")
            options = {}
            for idx, row in group.iterrows():
                # Formatação das datas (caso sejam válidas)
                try:
                    date_op_str = pd.to_datetime(row["Data da Operação"], dayfirst=True).strftime("%d/%m/%Y")
                except Exception:
                    date_op_str = "N/A"
                try:
                    date_ven_str = pd.to_datetime(row["Data Vencimento"], dayfirst=True).strftime("%d/%m/%Y")
                except Exception:
                    date_ven_str = "N/A"
                option_str = f"Doc: {row['Número do Documento']} | Data Op: {date_op_str} | Data Ven: {date_ven_str}"
                options[option_str] = idx
            chosen = st.selectbox(f"Selecione o comprovante correto para a conta {codigo}:", list(options.keys()), key=f"select_codigo_{codigo}")
            chosen_idx = options[chosen]
            resolved_rows.append(group.loc[[chosen_idx]])
    df_resolved = pd.concat(resolved_rows, ignore_index=True)
    # Resolver por "Número do Documento" (um comprovante associado a várias contas)
    duplicated_doc = df_resolved[df_resolved["Número do Documento"].notna() & df_resolved.duplicated(subset=["Número do Documento"], keep=False)]
    if not duplicated_doc.empty:
        grouped_doc = duplicated_doc.groupby("Número do Documento")
        for doc, group in grouped_doc:
            if len(group) > 1:
                st.write(f"Ambiguidade para o comprovante {doc} associado a várias contas:")
                options = {}
                for idx, row in group.iterrows():
                    try:
                        date_op_str = pd.to_datetime(row["Data da Operação"], dayfirst=True).strftime("%d/%m/%Y")
                    except Exception:
                        date_op_str = "N/A"
                    try:
                        date_ven_str = pd.to_datetime(row["Data Vencimento"], dayfirst=True).strftime("%d/%m/%Y")
                    except Exception:
                        date_ven_str = "N/A"
                    option_str = f"Código: {row['Código']} | Data Ven: {date_ven_str} | Data Op: {date_op_str}"
                    options[option_str] = idx
                chosen = st.selectbox(f"Selecione a conta correta para o comprovante {doc}:", list(options.keys()), key=f"select_doc_{doc}")
                chosen_idx = options[chosen]
                # Remove outras linhas com este comprovante
                df_resolved = df_resolved.drop(group.index.difference([chosen_idx]))
    return df_resolved

# --- INTERFACE DO APLICATIVO ---

st.title("🔎 Conciliação de Pagamentos - Verificação de Datas e Ambiguidades")

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
        # Ajuste o separador conforme necessário (aqui usamos vírgula)
        df_contas = pd.read_csv(contas_file, sep=",", dtype=str)
        required_cols = ["Empresa", "Fornecedor", "Data Vencimento", "Valor", "Código"]
        if not all(col in df_contas.columns for col in required_cols):
            st.error("A planilha de contas a pagar deve conter as colunas: Empresa, Fornecedor, Data Vencimento, Valor e Código.")
        else:
            # Padronização e conversão dos dados
            df_contas_std = df_contas.copy()
            df_contas_std = standardize_data(df_contas_std, ["Empresa", "Fornecedor"])
            df_contas_std["Código"] = df_contas_std["Código"].astype(str).str.strip()
            df_contas_std["Valor"] = df_contas_std["Valor"].str.replace(r"r\$\s*", "", regex=True).str.replace(",", ".").astype(float)
            df_contas_std["Valor_std"] = df_contas_std["Valor"].round(2)
            df_contas_std["Data Vencimento"] = pd.to_datetime(df_contas_std["Data Vencimento"], dayfirst=True, errors="coerce")
            st.subheader("Resumo da Planilha de Contas a Pagar")
            st.dataframe(df_contas_std)
            
            # 3. Seleção do método de conciliação
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
            
            # Converter "Data da Operação" para datetime
            df_conciliado["Data da Operação"] = pd.to_datetime(df_conciliado["Data da Operação"], dayfirst=True, errors="coerce")
            # Criar coluna que verifica se as datas coincidem
            df_conciliado["Data_Match"] = df_conciliado.apply(
                lambda row: row["Data Vencimento"] == row["Data da Operação"]
                if pd.notnull(row["Data Vencimento"]) and pd.notnull(row["Data da Operação"]) else False, axis=1
            )
            
            # 4. Resolver ambiguidades (mesmo comprovante para mais de uma conta ou vice-versa)
            duplicate_mask_codigo = df_conciliado.duplicated(subset=["Código"], keep=False)
            duplicate_mask_doc = df_conciliado["Número do Documento"].notna() & df_conciliado.duplicated(subset=["Número do Documento"], keep=False)
            if duplicate_mask_codigo.any() or duplicate_mask_doc.any():
                st.write("Foram encontradas ambiguidades na conciliação. Por favor, resolva:")
                df_conciliado_final = resolve_ambiguities(df_conciliado)
            else:
                df_conciliado_final = df_conciliado.copy()
            
            st.subheader("Tabela Conciliada Inicial")
            st.dataframe(df_conciliado_final)
            csv_conciliado = df_conciliado_final.to_csv(index=False, sep=";").encode()
            st.download_button("Baixar Tabela Conciliada Inicial (CSV)",
                               data=csv_conciliado,
                               file_name="tabela_conciliada_inicial.csv",
                               mime="text/csv",
                               key="download_conciliado_inicial")
            
            # 5. Listar contas a pagar sem comprovante (verifica nulos ou strings vazias)
            df_contas_sem_comprovante = df_conciliado_final[
    df_conciliado_final["Número do Documento"].isna() | 
    (df_conciliado_final["Número do Documento"].astype(str).str.strip() == "nan") |
    (df_conciliado_final["Número do Documento"].astype(str).str.strip() == "") |
    (df_conciliado_final["Número do Documento"].astype(str).str.strip().isna())
]
            st.subheader("📌 Contas a Pagar SEM Comprovante")
            if df_contas_sem_comprovante.empty:
            st.warning("🚨 Nenhuma conta a pagar sem comprovante foi encontrada.")
else:
    st.dataframe(df_contas_sem_comprovante)
    csv_contas_sem = df_contas_sem_comprovante.to_csv(index=False, sep=";").encode()
    st.download_button("Baixar Contas SEM Comprovante (CSV)",
                       data=csv_contas_sem,
                       file_name="contas_sem_comprovante.csv",
                       mime="text/csv",
                       key="download_contas_sem")

            # 6. Listar comprovantes sem vínculo
            linked_doc_numbers = df_conciliado_final["Número do Documento"].dropna().unique()
            df_receipts_sem_conta = df_comprovantes[~df_comprovantes["Número do Documento"].isin(linked_doc_numbers)]
            st.subheader("Comprovantes SEM Correspondência com Contas a Pagar")
            st.dataframe(df_receipts_sem_conta)
            csv_receipts_sem = df_receipts_sem_conta.to_csv(index=False, sep=";").encode()
            st.download_button("Baixar Comprovantes SEM Contas (CSV)",
                               data=csv_receipts_sem,
                               file_name="comprovantes_sem_conta.csv",
                               mime="text/csv",
                               key="download_comprovantes_sem_conta")

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
