import pandas as pd
import streamlit as st
import io

# --- Configuração da Página ---
st.set_page_config(
    page_title="Processador de Bases",
    page_icon="✅",
    layout="wide",
)

st.title("🚀 Processador e Validador de Bases")
st.write("Faça o upload das bases PAINEL, EDUCAPI e COMERCIAL para realizar as validações.")

# --- Upload dos Arquivos ---
st.header("1. Importação das Bases")
col1, col2, col3 = st.columns(3)

with col1:
    painel_file = st.file_uploader("Base PAINEL", type=["xlsx"])
with col2:
    educapi_file = st.file_uploader("Base EDUCAPI", type=["xlsx"])
with col3:
    comercial_file = st.file_uploader("Base COMERCIAL", type=["xlsx"])

# --- Botão de Processamento e Lógica Principal ---
if st.button("Processar Bases", type="primary", use_container_width=True):
    if painel_file:
        try:
            # Carregar DataFrames
            df_painel = pd.read_excel(painel_file, dtype={'H': str})
            df_educapi = pd.read_excel(educapi_file, dtype={'E': str}) if educapi_file else pd.DataFrame({'E': []})
            df_comercial = pd.read_excel(comercial_file, dtype={'E': str}) if comercial_file else pd.DataFrame({'E': []})
            
            # Otimização com sets para busca rápida
            educapi_cpfs = set(df_educapi['E'].str.strip())
            comercial_cpfs = set(df_comercial['E'].str.strip())

            # --- Regras de Validação ---
            with st.spinner("Aplicando regras de validação..."):
                # REGRA 1 (Coluna M)
                df_painel['VALIDAÇÃO ESTADO/STATUS'] = df_painel['L'].apply(
                    lambda x: 'Matricula Liberada SP' if str(x).strip().lower() == 'são paulo' else 'Matricula Liberada'
                )

                # REGRA 2 (Coluna N)
                df_painel['STATUS VALIDAÇÃO'] = df_painel.apply(
                    lambda row: 'OK' if str(row['VALIDAÇÃO ESTADO/STATUS']).strip() == str(row['C']).strip() else 'CORRIGIR',
                    axis=1
                )
                
                # REGRA 3 (Coluna O)
                def verificar_cpf(cpf):
                    cpf_str = str(cpf).strip()
                    if cpf_str in educapi_cpfs: return 'Matricula Liberada EDUCAPI'
                    if cpf_str in comercial_cpfs: return 'Matricula Liberada SPE'
                    return ''
                df_painel['PROCV VALIDAÇÃO'] = df_painel['H'].apply(verificar_cpf)

                # REGRA 4 (Coluna P)
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

            # Converter DataFrame para Excel em memória para download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_painel.to_excel(writer, index=False, sheet_name='Painel_Processado')
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Baixar Resultado Principal (Excel)",
                data=processed_data,
                file_name='PAINEL_Processado_Final.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True
            )

            # --- Exportação de Linhas sem PK ---
            linhas_sem_pk = df_painel[df_painel['H'].isnull() | (df_painel['H'].str.strip() == '')]
            if not linhas_sem_pk.empty:
                st.warning(f"Foram encontradas {len(linhas_sem_pk)} linhas sem CPF (PK).")
                
                output_sem_pk = io.BytesIO()
                with pd.ExcelWriter(output_sem_pk, engine='xlsxwriter') as writer:
                    linhas_sem_pk.to_excel(writer, index=False, sheet_name='Linhas_Sem_PK')
                    # Adiciona a aba de resumo
                    resumo = linhas_sem_pk.describe(include='all').transpose().reset_index().rename(columns={'index': 'Coluna'})
                    resumo.to_excel(writer, index=False, sheet_name='Resumo_Geral')
                
                sem_pk_data = output_sem_pk.getvalue()
                st.download_button(
                    label=f"📥 Baixar Relatório de {len(linhas_sem_pk)} Linhas Sem PK (Excel)",
                    data=sem_pk_data,
                    file_name='Relatorio_Linhas_Sem_PK.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"Ocorreu um erro durante o processamento: {e}")
    else:
        st.error("Por favor, faça o upload da base PAINEL para continuar.")