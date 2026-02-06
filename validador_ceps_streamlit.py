import streamlit as st
import pandas as pd

st.set_page_config(page_title="Validação de CEP", layout="centered")

st.title("📍 Validação de CEP - F TGT DF")
st.write("Faça upload dos arquivos para verificar se os CEPs pertencem à base.")

# Upload dos arquivos
arquivo_base = st.file_uploader(
    "📄 Upload do arquivo CEP Base",
    type=["xlsx"]
)

arquivo_logradouros = st.file_uploader(
    "📄 Upload do arquivo Logradouros",
    type=["xlsx"]
)

if arquivo_base and arquivo_logradouros:
    try:
        base_df = pd.read_excel(arquivo_base)
        logradouros_df = pd.read_excel(arquivo_logradouros)

        # Garantir tipo inteiro
        for df in [base_df, logradouros_df]:
            df["CEP inicial"] = df["CEP inicial"].astype(int)
            df["CEP final"] = df["CEP final"].astype(int)

        def pertence_f_tgt(cep_ini, cep_fim, base):
            return (
                (base["CEP inicial"] <= cep_fim) &
                (base["CEP final"] >= cep_ini)
            ).any()

        # Aplicar regra
        logradouros_df["Pertence_F_TGT_DF"] = logradouros_df.apply(
            lambda row: pertence_f_tgt(
                row["CEP inicial"],
                row["CEP final"],
                base_df
            ),
            axis=1
        )

        st.success("✅ Processo finalizado com sucesso!")

        # Mostrar prévia
        st.subheader("🔎 Prévia do resultado")
        st.dataframe(logradouros_df.head(20))

        # Botão de download
        output = logradouros_df.to_excel(
            index=False,
            engine="openpyxl"
        )

        st.download_button(
            label="⬇️ Baixar arquivo com resultado",
            data=output,
            file_name="LOGRADOURO_COM_RESULTADO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Erro ao processar os arquivos: {e}")
else:
    st.info("⬆️ Envie os dois arquivos para iniciar o processamento.")