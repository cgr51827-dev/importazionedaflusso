import streamlit as st
import pandas as pd
import io

st.title("Compilatore BCC")

uploaded_file = st.file_uploader("Carica file flusso", type=["xlsx", "xlsm"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.success("File caricato!")

def col(df, lettera):
    idx = 0
    for c in lettera:
        idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return df.iloc[:, idx - 1]

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
    # ALTRI DATI
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
    # RECAPITI TELEFONICI
    # =========================
    def recapiti(df):
        colonne = ["R","S","T","BA","BB","BS","BT","BU","BC"]

        def get_col(df, name):
            idx = sum([(ord(c)-64)*(26**i) for i,c in enumerate(reversed(name))]) - 1
            return df.iloc[:, idx]

        righe = []

        for i in range(len(df)):
            valori = []

            for c in colonne:
                val = str(get_col(df, c).iloc[i]).strip()
                if val and val != "nan":
                    valori.append(val)

            while len(valori) < 11:
                valori.append("")

            nuova = {"B": str(col(df,"E").iloc[i])}

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

    def crea_excel(df_out):
        buffer = io.BytesIO()
        df_out.to_excel(buffer, index=False)
        return buffer.getvalue()

    if st.button("Genera file"):
        st.success("File pronti!")

        st.download_button("IMPORT STANDARD", crea_excel(genera_import_standard(df)), "import.xlsx")
        st.download_button("RECAPITI", crea_excel(recapiti(df)), "recapiti.xlsx")
        st.download_button("SALDO", crea_excel(genera_saldo(df)), "saldo.xlsx")
        st.download_button("ALTRI DATI", crea_excel(genera_altri_dati(df)), "altri.xlsx")
        st.download_button("MAIL CLIENTI", crea_excel(mail_clienti(df)), "mail_clienti.xlsx")
        st.download_button("MAIL BANCHE", crea_excel(mail_banche(df)), "mail_banche.xlsx")
