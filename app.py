import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.title("Traitement carte Afriquia - Génération écritures comptables")

uploaded_file = st.file_uploader(
    "Importer le fichier Excel (factures + details)",
    type=["xlsx"]
)

if uploaded_file:

    xl = pd.ExcelFile(uploaded_file)

    if 'factures' not in xl.sheet_names or 'details' not in xl.sheet_names:
        st.error("Le fichier doit contenir les feuilles 'factures' et 'details'")
        st.stop()

    df_factures = pd.read_excel(xl, sheet_name='factures', dtype=str)
    df_details = pd.read_excel(xl, sheet_name='details', dtype=str)

    df_factures.columns = df_factures.columns.str.strip().str.lower()
    df_details.columns = df_details.columns.str.strip().str.lower()

    df_details = df_details.rename(columns={'code carte': 'carte'})
    df_details = df_details.drop_duplicates(subset=['carte'])

    df_factures['carte'] = df_factures['carte'].astype(str)
    df_details['carte'] = df_details['carte'].astype(str)

    df_merge = pd.merge(
        df_factures,
        df_details,
        on='carte',
        how='left'
    )

    df_merge['montant transaction ttc'] = pd.to_numeric(
        df_merge['montant transaction ttc'],
        errors='coerce'
    ).fillna(0)

    df_merge['type_produit'] = df_merge['produit'].str.lower().apply(
        lambda x: 'peage' if 'peage' in str(x) else 'gazoil'
    )

    df_grouped = df_merge.groupby(
        ['carte', 'salarie', 'modalite', 'code affaire', 'type_produit']
    )['montant transaction ttc'].sum().reset_index()

    lignes = []

    mois_precedent = datetime.now().replace(day=1) - pd.Timedelta(days=1)
    date_compta = mois_precedent.date()

    total = df_grouped['montant transaction ttc'].sum()

    for carte in df_grouped['carte'].unique():

        df_carte = df_grouped[df_grouped['carte'] == carte]

        peage = df_carte[df_carte['type_produit'] == 'peage']

        if not peage.empty:

            r = peage.iloc[0]

            lignes.append({
                "Type compte": "Compte général",
                "N° compte": "614310000",
                "Description": f"Afriquia Carte {carte} - PEAGE",
                "MODALITÉ Code": r["modalite"],
                "Salarie Code": r["salarie"],
                "Affaire Code": r["code affaire"],
                "Montant débit": r["montant transaction ttc"],
                "Montant crédit": 0
            })

        gazoil = df_carte[df_carte['type_produit'] == 'gazoil']

        if not gazoil.empty:

            r = gazoil.iloc[0]

            lignes.append({
                "Type compte": "Compte général",
                "N° compte": "612515000",
                "Description": f"Afriquia Carte {carte} - GAZOIL",
                "MODALITÉ Code": r["modalite"],
                "Salarie Code": r["salarie"],
                "Affaire Code": r["code affaire"],
                "Montant débit": r["montant transaction ttc"],
                "Montant crédit": 0
            })

    lignes.append({
        "Type compte": "Fournisseur",
        "N° compte": "F2A021",
        "Description": "Fournisseur Afriquia",
        "Montant débit": 0,
        "Montant crédit": total
    })

    df_final = pd.DataFrame(lignes)

    st.subheader("Aperçu des écritures")
    st.dataframe(df_final)

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="ecritures comptables")

    st.download_button(
        label="Télécharger fichier écritures comptables",
        data=buffer.getvalue(),
        file_name="ecritures_comptables_afriquia.xlsx"
    )
