import pandas as pd
import streamlit as st

# =============================
# FUNÇÕES AUXILIARES
# =============================
def limpar_colunas(df):
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace("\n", " ")
        .str.replace("\r", " ")
    )
    return df

def formatar_ordens(ordens):
    ordens = sorted(ordens)
    if len(ordens) == 1:
        return f"Ordem {ordens[0]}"
    texto = ", ".join(str(o) for o in ordens[:-1])
    texto += f" e {ordens[-1]}"
    return f"Ordens para esta parada: {texto}"

# =============================
# PROCESSAMENTO PRINCIPAL
# =============================
def processar_dataframe(df):
    df = limpar_colunas(df)

    # Conversões seguras
    df["Sequence"] = df["Sequence"].astype(int)
    df["Stop"] = df["Stop"].astype(int)

    # Coluna de endereço REAL do arquivo
    col_endereco = "Unnamed: 4"

    # Agrupamento por endereço
    agrupado = (
        df.groupby(col_endereco, as_index=False)
        .agg({
            "Sequence": list,
            "Stop": list,
            "Bairro": "first",
            "City": "first",
            "Zipcode/Postal code": "first"
        })
    )

    agrupado["Parada_num"] = agrupado["Stop"].apply(min)
    agrupado["Observações"] = agrupado["Sequence"].apply(formatar_ordens)
    agrupado["Total de Pacotes"] = agrupado["Sequence"].apply(
        lambda x: f"{len(x)} pacotes"
    )

    # =============================
    # DATAFRAME FINAL PARA O CIRCUIT
    # =============================
    saida = pd.DataFrame()
    saida["Parada"] = agrupado["Parada_num"].apply(lambda x: f"Parada {x}")
    saida["Address Line"] = agrupado[col_endereco]
    saida["Secondary Address Line"] = agrupado["Bairro"]
    saida["City"] = agrupado["City"]
    saida["State"] = "São Paulo"
    saida["Zip Code"] = agrupado["Zipcode/Postal code"]
    saida["Total de Pacotes"] = agrupado["Total de Pacotes"]
    saida["Observações"] = agrupado["Observações"]

    return saida

# =============================
# STREAMLIT UI
# =============================
st.set_page_config(
    page_title="Conversor Shopee → Circuit",
    layout="centered"
)

st.title("📦 Conversor de Rotas para Circuit")

arquivo = st.file_uploader("Selecione o arquivo Excel", type=["xlsx"])

if arquivo:
    try:
        df_original = pd.read_excel(arquivo)

        st.success("Arquivo carregado com sucesso")
        st.write("Colunas detectadas:")
        st.write(list(df_original.columns))

        if st.button("🚀 Processar arquivo"):
            df_saida = processar_dataframe(df_original)

            st.success("Arquivo processado com sucesso")
            st.dataframe(df_saida)

            output_file = "rota_circuit.xlsx"
            df_saida.to_excel(output_file, index=False)

            with open(output_file, "rb") as f:
                st.download_button(
                    "⬇️ Baixar Excel",
                    data=f,
                    file_name="rota_circuit.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
