import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Filtro de Empregados", layout="wide")
st.title("🔍 Filtro de Empregados por Cidade e Cargo")

# Função para processar a planilha e aplicar filtros
def carregar_dados_com_filtro(arquivo, cidades=None, cargo=None):
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
        return pd.DataFrame()

    df_final = pd.concat(dataframes, ignore_index=True)

    # Aplica filtros
    if cidades:
        df_final = df_final[
            df_final["Cidade de Atuação"].str.upper().isin([c.upper() for c in cidades])
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

            # Filtros
            cidades_sel = st.multiselect("🌆 Selecione uma ou mais Cidades", cidades)
            cargo_sel = st.selectbox("💼 Selecione o Cargo", [""] + cargos)

            # Aplica os filtros
            df_filtrado = carregar_dados_com_filtro(
                arquivo,
                cidades=cidades_sel if cidades_sel else None,
                cargo=cargo_sel if cargo_sel else None
            )

            st.markdown(f"### 📋 Resultado ({len(df_filtrado)} registros encontrados)")

            colunas_exibidas = [
                "empresa",
                "Empregado",
                "CPF",
                "Data Nascimento",
                "Cargo",
                "Cidade de Atuação"
            ]

            st.dataframe(df_filtrado[colunas_exibidas])

            # Exportar como Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_filtrado[colunas_exibidas].to_excel(writer, index=False, sheet_name='Empregados Filtrados')
                writer.save()
                output.seek(0)

            st.download_button(
                label="📥 Baixar Excel com resultado",
                data=output,
                file_name="empregados_filtrados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
