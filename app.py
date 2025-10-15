import pandas as pd
import streamlit as st
import io

# --- Função de Leitura Robusta (com correção de encoding) ---
def carregar_arquivo(uploaded_file):
    """
    Tenta ler um arquivo como Excel (.xlsx, .xls) ou CSV, lidando com
    erros comuns de codificação de caracteres.
    """
    if uploaded_file is None:
        return None
    try:
        # Tenta ler como Excel primeiro (geralmente lida bem com encodings)
        df = pd.read_excel(uploaded_file, dtype=str)
        return df
    except Exception:
        try:
            # Se falhar, reseta e tenta como CSV usando 'latin-1' como fallback
            uploaded_file.seek(0)
            # Tenta detectar o separador (vírgula ou ponto e vírgula)
            preview = uploaded_file.read(1024).decode('latin-1')
            sep = ';' if preview.count(';') > preview.count(',') else ','
            uploaded_file.seek(0)
            # Usa encoding='latin-1' para evitar o erro 'utf-8'
            df = pd.read_csv(uploaded_file, sep=sep, dtype=str, encoding='latin-1', on_bad_lines='warn')
            return df
        except Exception as e:
            st.error(f"Não foi possível ler o arquivo {uploaded_file.name}. Verifique o formato. Erro final: {e}")
            return None

# --- Configuração da Página ---
st.set_page_config(
    page_title="Processador de Bases",
    page_icon="✅",
    layout="wide",
)

st.title("🚀 Processador e Validador de Bases")
st.write("Faça o upload das bases PAINEL, EDUCAPI e COMERCIAL nos formatos .xlsx, .xls ou .csv.")

# --- Upload dos Arquivos ---
st.header("1. Importação das Bases")
col1, col2, col3 = st.columns(3)

with col1:
    painel_file = st.file_uploader("Base PAINEL", type=["xlsx", "xls", "csv"])
with col2:
    educapi_file = st.file_uploader("Base EDUCAPI", type=["xlsx", "xls", "csv"])
with col3:
    comercial_file = st.file_uploader("Base COMERCIAL", type=["xlsx", "xls", "csv"])

# --- Botão de Processamento e Lógica Principal ---
if st.button("Processar Bases", type="primary", use_container_width=True):
    if painel_file:
        with st.spinner("Carregando e processando arquivos..."):
            # Carregar DataFrames usando a função corrigida
            df_painel = carregar_arquivo(painel_file)
            df_educapi = carregar_arquivo(educapi_file) if educapi_file else pd.DataFrame({'E': []})
            df_comercial = carregar_arquivo(comercial_file) if comercial_file else pd.DataFrame({'E': []})

            if df_painel is None:
                st.error("Falha ao ler a base PAINEL. O processamento foi interrompido.")
                st.stop()

            colunas_necessarias = ['L', 'C', 'H']
            if not all(col in df_painel.columns for col in colunas_necessarias):
                st.error(f"Erro: A base PAINEL deve conter as colunas: {', '.join(colunas_necessarias)}.")
                st.stop()
            
            # --- Regras de Validação ---
            educapi_cpfs = set(df_educapi['E'].str.strip()) if 'E' in df_educapi.columns and not df_educapi.empty else set()
            comercial_cpfs = set(df_comercial['E'].str.strip()) if 'E' in df_comercial.columns and not df_comercial.empty else set()

            df_painel['VALIDAÇÃO ESTADO/STATUS'] = df_painel['L'].apply(
                lambda x: 'Matricula Liberada SP' if str(x).strip().lower() == 'são paulo' else 'Matricula Liberada'
            )
            df_painel['STATUS VALIDAÇÃO'] = df_painel.apply(
                lambda row: 'OK' if str(row['VALIDAÇÃO ESTADO/STATUS']).strip() == str(row['C']).strip() else 'CORRIGIR', axis=1
            )
            def verificar_cpf(cpf):
                cpf_str = str(cpf).strip()
                if cpf_str in educapi_cpfs: return 'Matricula Liberada EDUCAPI'
                if cpf_str in comercial_cpfs: return 'Matricula Liberada SPE'
                return ''
            df_painel['PROCV VALIDAÇÃO'] = df_painel['H'].apply(verificar_cpf)
            def status_final_validacao(row):
                c, m, o = str(row['C']).strip(), str(row['VALIDAÇÃO ESTADO/STATUS']).strip(), str(row['PROCV VALIDAÇÃO']).strip()
                if not o: return 'VERIFICAR'
                if c == m == o: return 'OK'
                if m == o and c != m: return 'STATUS CADASTRADO DE FORMA INCORRETA'
                return 'VERIFICAR'
            df_painel['STATUS FINAL'] = df_painel.apply(status_final_validacao, axis=1)

        st.success("Bases processadas com sucesso!")

        # --- Visualização e Download ---
        st.header("2. Resultados")
        st.dataframe(df_painel)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_painel.to_excel(writer, index=False, sheet_name='Painel_Processado')
        st.download_button(
            label="📥 Baixar Resultado Principal (Excel)", data=output.getvalue(), file_name='PAINEL_Processado_Final.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True
        )

        # --- Exportação de Linhas sem PK ---
        linhas_sem_pk = df_painel[df_painel['H'].isnull() | (df_painel['H'].str.strip() == '')]
        if not linhas_sem_pk.empty:
            st.warning(f"Foram encontradas {len(linhas_sem_pk)} linhas sem CPF (PK).")
            output_sem_pk = io.BytesIO()
            with pd.ExcelWriter(output_sem_pk, engine='xlsxwriter') as writer:
                linhas_sem_pk.to_excel(writer, index=False, sheet_name='Linhas_Sem_PK')
                resumo = linhas_sem_pk.describe(include='all').transpose().reset_index().rename(columns={'index': 'Coluna'})
                resumo.to_excel(writer, index=False, sheet_name='Resumo_Geral')
            st.download_button(
                label=f"📥 Baixar Relatório de {len(linhas_sem_pk)} Linhas Sem PK (Excel)", data=output_sem_pk.getvalue(),
                file_name='Relatorio_Linhas_Sem_PK.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True
            )
    else:
        st.error("Por favor, faça o upload da base PAINEL para continuar.")
