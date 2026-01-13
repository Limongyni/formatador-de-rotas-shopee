import pandas as pd
import re
import unicodedata
import streamlit as st

# =============================
# FUNÇÕES DE SUPORTE
# =============================
def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = unicodedata.normalize("NFKD", texto)
    texto = texto.encode("ASCII", "ignore").decode("ASCII")
    texto = texto.lower().strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto

def limpar_colunas(df):
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace("\n", " ")
        .str.replace("\r", " ")
    )
    return df

def encontrar_coluna(df, termos):
    for termo in termos:
        for col in df.columns:
            if termo.lower() in col.lower():
                return col
    raise KeyError(f"Coluna não encontrada. Esperado algo como: {termos}")

def extrair_rua_numero(endereco):
    if pd.isna(endereco):
        return None, None

    match = re.search(r"^(.*?),\s*(\d+)", endereco)
    if match:
        return match.group(1).strip(), match.group(2)
    return endereco.strip(), None

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

    col_endereco = encontrar_coluna(df, ["destination", "address"])
    col_sequence = encontrar_coluna(df, ["sequence"])
    col_stop = encontrar_coluna(df, ["stop"])
    col_city = encontrar_coluna(df, ["city"])
    col_bairro = encontrar_coluna(df, ["bairro", "neighborhood"])
    col_cep = encontrar_coluna(df, ["zip", "postal"])

    df[["Rua", "Numero"]] = df[col_endereco].apply(
        lambda x: pd.Series(extrair_rua_numero(x))
    )

    df["Rua_norm"] = df["Rua"].apply(normalizar_texto)
    df["Numero_norm"] = df["Numero"].astype(str)

    df["Sequence_num"] = df[col_sequence].astype(int)
    df["Stop_num"] = df[col_stop].astype(int)

    agrupado = (
        df.groupby(["Rua_norm", "Numero_norm"], as_index=False)
        .agg({
            "Rua": "first",
            "Numero": "first",
            col_bairro: "first",
            col_city: "first",
            col_cep: "first",
            "Sequence_num": list,
            "Stop_num": list
        })
    )

    agrupado["Stop_final"] = agrupado["Stop_num"].apply(min)
    agrupado["Observações"] = agrupado["Sequence_num"].apply(formatar_ordens)
    agrupado["Total de Pacotes"] = agrupado["Sequence_num"].apply(
        lambda x: f"{len(x)} pacotes"
    )

    saida = pd.DataFrame()
    saida["Stop Name"] = agrupado["Stop_final"].apply(lambda x: f"Parada {x}")
    saida["Address"] = agrupado["Rua"] + ", " + agrupado["Numero"]
    saida["Secondary Address Line"] = agrupado[col_bairro]
    saida["City"] = agrupado[col_city]
    saida["State"] = "São Paulo"
    saida["Zip Code"] = agrupado[col_cep]
    saida["Observações"] = agrupado["Observações"]
    saida["Total de Pacotes"] = agrupado["Total de Pacotes"]

    return saida

# =============================
# INTERFACE STREAMLIT
# =============================
st.set_page_config(
    page_title="Agrupador de Paradas - Circuit",
    layout="centered"
)

st.title("📦 Agrupador de Paradas por Endereço")
st.write("Compatível com desktop e mobile.")

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

            output_file = "saida_circuit.xlsx"
            df_saida.to_excel(output_file, index=False)

            with open(output_file, "rb") as f:
                st.download_button(
                    "⬇️ Baixar Excel",
                    data=f,
                    file_name="saida_circuit.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
