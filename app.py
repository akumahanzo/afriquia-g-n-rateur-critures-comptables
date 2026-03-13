import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.title("Traitement carte Afriquia")

uploaded_file = st.file_uploader("Importer le fichier Excel", type=["xlsx"])

if uploaded_file:

    xl = pd.ExcelFile(uploaded_file)

    if 'factures' not in xl.sheet_names or 'details' not in xl.sheet_names:
        st.error("Le fichier doit contenir les feuilles 'factures' et 'details'")
    else:

        df_factures = pd.read_excel(xl, sheet_name='factures', dtype=str)
        df_details = pd.read_excel(xl, sheet_name='details', dtype=str)

        df_factures.columns = df_factures.columns.str.strip().str.lower()
        df_details.columns = df_details.columns.str.strip().str.lower()

        df_details = df_details.rename(columns={'code carte':'carte'})

        df_merge = pd.merge(
            df_factures,
            df_details,
            on='carte',
            how='left'
        )

        st.write("Aperçu des données")
        st.dataframe(df_merge)

        buffer = BytesIO()

        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_merge.to_excel(writer,index=False)

        st.download_button(
            label="Télécharger fichier Excel",
            data=buffer.getvalue(),
            file_name="resultat.xlsx"
        )
