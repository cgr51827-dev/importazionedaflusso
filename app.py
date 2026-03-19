import streamlit as st
import pandas as pd
import io
import zipfile

st.title("Compilatore BCC")

uploaded_file = st.file_uploader("Carica file flusso", type=["xlsx", "xlsm"])

# funzione per leggere colonne Excel (A, J, AF, BA ecc.)
def col(df, lettera):
    idx = 0
    for c in lettera:
        idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return df.iloc[:, idx - 1]

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.success("File caricato!")

    # =========================
    # IMPORT STANDARD
    # =========================
    def genera_import_standard(df):
        out = pd.DataFrame()
        out["A"] = col(df, "J")
        out["D"] = col(df, "Q")
        out["J"] = col(df, "K")
        out["K"] = col(df, "N")
        out["L"] = col(df, "L")
        out["M"] = col(df, "M")
        out["N"] = col(df, "AF")
        out["Q"] = col(df, "AJ")
        out["R"] = col(df, "AN")
        out["U"] = col(df, "E")
        return out

    # =========================
    # SALDO
    # =========================
    def genera_saldo(df):
        out = pd.DataFrame()
        out["B"] = col(df, "E")
        out["C"] = pd.to_numeric(col(df, "AB"), errors="coerce")
        return out

    # =========================
    # ALTRI DATI (zeri iniziali)
    # =========================
    def genera_altri_dati(df):
        out = pd.DataFrame()
        out["A"] = col(df, "E").astype(str)
        out["B"] = col(df, "IP").astype(str)
        out["C"] = col(df, "IM").astype(str)
        out["D"] = col(df, "IR").astype(str)
        out["E"] = col(df, "IN").astype(str)
        out["F"] = col(df, "IO").astype(str)
        return out

    # =========================
    # RECAPITI TELEFONICI H:R
    # =========================
    def recapiti(df):
        colonne = ["R","S","T","BA","BB","BS","BT","BU","BC"]

        righe = []

        for i in range(len(df)):
            valori = []

            for c in colonne:
                val = str(col(df, c).iloc[i]).strip()
                if val and val != "nan":
                    valori.append(val)

            # compattazione senza buchi
            while len(valori) < 11:
                valori.append("")

            nuova = {"B": str(col(df, "E").iloc[i])}

            colonne_out = list("HIJKLMNOPQR")

            for j, c in enumerate(colonne_out):
                nuova[c] = valori[j]

            righe.append(nuova)

        return pd.DataFrame(righe)

    # =========================
    # MAIL CLIENTI
    # =========================
    def mail_clienti(df):
        out = pd.DataFrame()
        out["A"] = col(df,"E")
        out["B"] = col(df,"J")
        out["C"] = col(df,"AP")
        out["D"] = col(df,"GN")
        out["E"] = col(df,"GO")
        out["F"] = col(df,"IN")
        out["G"] = col(df,"IO")
        return out

    # =========================
    # MAIL BANCHE
    # =========================
    def mail_banche(df):
        out = pd.DataFrame()
        out["A"] = col(df,"E")
        out["B"] = col(df,"J")
        out["C"] = col(df,"AP")
        out["D"] = col(df,"GN")
        out["E"] = col(df,"GO")
        out["F"] = col(df,"IR")
        return out

    # =========================
    # GENERAZIONE ZIP
    # =========================
    if st.button("Genera e scarica tutto"):
        st.success("File pronti!")

        files = {
            "import_standard.xlsx": genera_import_standard(df),
            "recapiti.xlsx": recapiti(df),
            "saldo.xlsx": genera_saldo(df),
            "altri_dati.xlsx": genera_altri_dati(df),
            "mail_clienti.xlsx": mail_clienti(df),
            "mail_banche.xlsx": mail_banche(df),
        }

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as z:
            for nome, dataframe in files.items():
                excel_buffer = io.BytesIO()
                dataframe.to_excel(excel_buffer, index=False)
                z.writestr(nome, excel_buffer.getvalue())

        st.download_button(
            label="Scarica TUTTI i file (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="compilatore_bcc.zip",
            mime="application/zip"
        )
