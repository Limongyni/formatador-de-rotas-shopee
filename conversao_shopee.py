import pandas as pd
import re
import unicodedata
import streamlit as st

# =============================
# FUNÇÕES AUXILIARES
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

def encontrar_coluna_endereco(df):
    """
    Procura a coluna de endereço de forma robusta:
    - Destination Address
    - destination address
    - Unnamed: X (caso padrão Shopee)
    """
    for col in df.columns:
        col_norm = col.lower().strip()
        if "destination" in col_norm and "address" in col_norm:
            return col
    # fallback para padrão Shopee
    if "Unnamed: 4" in df.columns:
        return "Unnamed: 4"
    raise KeyError("Coluna de endereço não encontrada")

def extrair_rua_numero(endereco):
    if pd.isna(endereco):
        return None, None

    rua_match = re.search(r"^(.*?),\s*\d+", endereco)
    numero_match = re.search(r",\s*(\d+)", endereco)

    rua = rua_match.group(1).strip() if rua_match else endereco
    numero = numero_match.group(1) if numero_match else None

    return rua, numero

def formatar_ordens(lista_ordens):
    lista_ordens = sorted(lista_ordens)

    if len(lista_ordens) == 1:
        return f"Ordem {lista_ordens[0]}"

    texto = ", ".join(str(o) for o in lista_ordens[:-1])
    texto += f" e {lista_ordens[-1]}"
    return f"Ordens para esta parada: {texto}"

# =============================
# PROCESSAMENTO
# =============================
def processar_dataframe(df):
    df = limpar_colunas(df).copy()

    # 🔑 detectar coluna de endereço corretamente
    col_endereco = encontrar_coluna_endereco(df)

    df[["Rua", "Numero"]] = df[col_endereco].apply(
        lambda x: pd.Series(extrair_rua_numero(x))
    )

    df["Rua_norm"] = df["Rua"].apply(normalizar_texto)
    df["Numero_norm"] = df["Numero"].astype(str)

    df["Sequence_num"] = df["Sequence"].astype(int)
    df["Stop_num"] = df["Stop"].astype(int)

    agrupado = (
        df.groupby(["Rua_norm", "Numero_norm"], as_index=False)
        .agg({
            "Rua": "first",
            "Numero": "first",
            "Bairro": "first",
            "City": "first",
            "Zipcode/Postal code": "first",
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
    saida["Address Line"] = agrupado["Rua"] + ", " + agrupado["Numero"]
    saida["Secondary Address Line"] = agrupado["Bairro"]
    saida["City"] = agrupado["City"]
    saida["State"] = "São Paulo"
    saida["Zip Code"] = agrupado["Zipcode/Postal code"]
    saida["Observações"] = agrupado["Observações"]
    saida["Total de Pacotes"] = agrupado["Total de Pacotes"]

    return saida

# =============================
# STREAMLIT UI
# =============================
st.set_page_config(
    page_title="Agrupador de Paradas - Circuit",
    layout="centered"
)

st.title("📦 Agrupador de Paradas por Endereço")

arquivo = st.file_uploader("Selecione o arquivo Excel", type=["xlsx"])

if arquivo:
    try:
        df_original = pd.read_excel(arquivo)

        st.success("Arquivo carregado com sucesso!")
        st.write("Colunas detectadas:")
        st.write(list(df_original.columns))

        if st.button("🚀 Processar arquivo"):
            df_saida = processar_dataframe(df_original)

            st.success("Arquivo processado com sucesso!")
            st.dataframe(df_saida)

            output_path = "saida_circuit.xlsx"
            df_saida.to_excel(output_path, index=False)

            with open(output_path, "rb") as f:
                st.download_button(
                    label="⬇️ Baixar arquivo Excel",
                    data=f,
                    file_name="saida_circuit.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")

