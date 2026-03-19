import streamlit as st
import pandas as pd
import io

st.title("Compilatore BCC")

uploaded_file = st.file_uploader("Carica file flusso", type=["xlsx", "xlsm"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.success("File caricato!")

    # =========================
    # IMPORT STANDARD
    # =========================
    def genera_import_standard(df):
        out = pd.DataFrame()
        out["A"] = df["J"]
        out["D"] = df["Q"]
        out["J"] = df["K"]
        out["K"] = df["N"]
        out["L"] = df["L"]
        out["M"] = df["M"]
        out["N"] = df["AF"]
        out["Q"] = df["AJ"]
        out["R"] = df["AN"]
        out["U"] = df["E"]
        return out

    # =========================
    # SALDO
    # =========================
    def genera_saldo(df):
        out = pd.DataFrame()
        out["B"] = df["E"]
        out["C"] = pd.to_numeric(df["AB"], errors="coerce")
        return out

    # =========================
    # ALTRI DATI (zeri iniziali)
    # =========================
    def genera_altri_dati(df):
        out = pd.DataFrame()
        out["A"] = df["E"].astype(str)
        out["B"] = df["IP"].astype(str)
        out["C"] = df["IM"].astype(str)
        out["D"] = df["IR"].astype(str)
        out["E"] = df["IN"].astype(str)
        out["F"] = df["IO"].astype(str)
        return out

    # =========================
    # RECAPITI TELEFONICI (H:R)
    # =========================
    def recapiti(df):
        colonne = ["R","S","T","BA","BB","BS","BT","BU","BC"]

        righe = []

        for _, row in df.iterrows():
            valori = []

            for col in colonne:
                val = str(row.get(col, "")).strip()
                if val and val != "nan":
                    valori.append(val)

            # compattazione
            while len(valori) < 11:
                valori.append("")

            nuova_riga = {
                "B": str(row["E"]),
            }

            colonne_output = list("HIJKLMNOPQR")

            for i, col in enumerate(colonne_output):
                nuova_riga[col] = valori[i]

            righe.append(nuova_riga)

        return pd.DataFrame(righe)

    # =========================
    # MAIL CLIENTI
    # =========================
    def mail_clienti(df):
        out = pd.DataFrame()
        out["A"] = df["E"]
        out["B"] = df["J"]
        out["C"] = df["AP"]
        out["D"] = df["GN"]
        out["E"] = df["GO"]
        out["F"] = df["IN"]
        out["G"] = df["IO"]
        return out

    # =========================
    # MAIL BANCHE
    # =========================
    def mail_banche(df):
        out = pd.DataFrame()
        out["A"] = df["E"]
        out["B"] = df["J"]
        out["C"] = df["AP"]
        out["D"] = df["GN"]
        out["E"] = df["GO"]
        out["F"] = df["IR"]
        return out

    def crea_excel(df_out):
        buffer = io.BytesIO()
        df_out.to_excel(buffer, index=False)
        return buffer.getvalue()

    if st.button("Genera file"):
        st.success("File pronti!")

        st.download_button(
            "Scarica IMPORT STANDARD",
            crea_excel(genera_import_standard(df)),
            "import_standard.xlsx"
        )

        st.download_button(
            "Scarica RECAPITI TELEFONICI",
            crea_excel(recapiti(df)),
            "recapiti.xlsx"
        )

        st.download_button(
            "Scarica SALDO",
            crea_excel(genera_saldo(df)),
            "saldo.xlsx"
        )

        st.download_button(
            "Scarica ALTRI DATI",
            crea_excel(genera_altri_dati(df)),
            "altri_dati.xlsx"
        )

        st.download_button(
            "Scarica MAIL CLIENTI",
            crea_excel(mail_clienti(df)),
            "mail_clienti.xlsx"
        )

        st.download_button(
            "Scarica MAIL BANCHE",
            crea_excel(mail_banche(df)),
            "mail_banche.xlsx"
        )
