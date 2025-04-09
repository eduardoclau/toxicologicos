import streamlit as st
import pandas as pd

st.set_page_config(page_title="Filtro de Empregados", layout="wide")
st.title("🔍 Filtro de Empregados por Cidade e Cargo")

# Função para processar a planilha
def carregar_dados_com_filtro(arquivo, cidade=None, cargo=None):
    xls = pd.ExcelFile(arquivo)
    df = xls.parse(xls.sheet_names[0], header=None)

    dataframes = []
    empresa_atual = None
    start_idx = None

    for idx, row in df.iterrows():
        if row[0] == "Empresa":
            empresa_atual = row[1]
            continue
        elif isinstance(row[0], str) and row[0].startswith("Total de empregados"):
            if start_idx is not None:
                bloco = df.iloc[start_idx:idx]
                bloco.columns = header
                bloco = bloco.reset_index(drop=True)
                bloco["empresa"] = empresa_atual
                dataframes.append(bloco)
            start_idx = None
        elif row[0] == "Matrícula":
            header = row.tolist()
            start_idx = idx + 1

    if not dataframes:
        return pd.DataFrame()  # retorna vazio se não houver dados

    df_final = pd.concat(dataframes, ignore_index=True)

    # Aplica filtros
    if cidade:
        df_final = df_final[
            df_final["Cidade de Atuação"].str.contains(cidade, case=False, na=False)
        ]
    if cargo:
        df_final = df_final[
            df_final["Cargo"].str.contains(cargo, case=False, na=False)
        ]

    return df_final

# Upload do arquivo
arquivo = st.file_uploader("📁 Envie a planilha .xlsx", type="xlsx")

if arquivo:
    try:
        df_completo = carregar_dados_com_filtro(arquivo)

        if df_completo.empty:
            st.warning("❗ Nenhum dado encontrado na planilha.")
        else:
            cidades = sorted(df_completo["Cidade de Atuação"].dropna().unique())
            cargos = sorted(df_completo["Cargo"].dropna().unique())

            # Seletores
            cidade_sel = st.selectbox("🌆 Selecione a Cidade de Atuação", [""] + cidades)
            cargo_sel = st.selectbox("💼 Selecione o Cargo", [""] + cargos)

            # Filtrando com base na seleção
            df_filtrado = carregar_dados_com_filtro(arquivo, cidade=cidade_sel if cidade_sel else None,
                                                    cargo=cargo_sel if cargo_sel else None)

            st.markdown(f"### 📋 Resultado ({len(df_filtrado)} registros encontrados)")
            st.dataframe(df_filtrado[["empresa", "Empregado", "Cargo", "Cidade de Atuação"]])

            # Botão para baixar resultado
            csv = df_filtrado.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Baixar CSV com resultado",
                data=csv,
                file_name="empregados_filtrados.csv",
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
