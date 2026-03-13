import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.title("Traitement carte Afriquia - T2S - CM")

uploaded_file = st.file_uploader(
    "Importer fichier Excel à traiter",
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

    df_factures['date facture'] = pd.to_datetime(
        df_factures['date facture'],
        errors='coerce'
    )

    cartes_factures = set(df_factures['carte'].unique())
    cartes_details = set(df_details['carte'].unique())

    cartes_manquantes = cartes_factures - cartes_details

    df_nouvelles_cartes = pd.DataFrame({
        'carte': list(cartes_manquantes),
        'salarie': [''] * len(cartes_manquantes),
        'modalite': [''] * len(cartes_manquantes),
        'code affaire': [''] * len(cartes_manquantes)
    })

    df_factures_joint = df_factures[df_factures['carte'].isin(cartes_details)]

    df_merge = pd.merge(
        df_factures_joint,
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

    lignes_charge = []

    total_fournisseur = df_grouped['montant transaction ttc'].sum()

    mois_precedent = datetime.now().replace(day=1) - pd.Timedelta(days=1)
    date_compta = mois_precedent.date()

    for carte in df_grouped['carte'].unique():

        df_carte = df_grouped[df_grouped['carte'] == carte]

        peage = df_carte[df_carte['type_produit'] == 'peage']

        if not peage.empty:

            r = peage.iloc[0]

            lignes_charge.append({
                'Nom de la feuille': '',
                'Date comptabilisation': date_compta,
                'Date TVA': date_compta,
                'Type compte': 'Compte général',
                'N° compte': '614310000',
                'N° document': '',
                'Description': f"Afriquia Carte {carte} - PÉAGE",
                'MODALITÉ Code': r['modalite'],
                'Salarie Code': r['salarie'],
                'Affaire Code': r['code affaire'],
                'Division Code': '',
                'Montant débit': r['montant transaction ttc'],
                'Montant crédit': 0,
                'Montant': r['montant transaction ttc']
            })

        gazoil = df_carte[df_carte['type_produit'] == 'gazoil']

        if not gazoil.empty:

            montant = gazoil['montant transaction ttc'].sum()

            r = gazoil.iloc[0]

            lignes_charge.append({
                'Nom de la feuille': '',
                'Date comptabilisation': date_compta,
                'Date TVA': date_compta,
                'Type compte': 'Compte général',
                'N° compte': '612515000',
                'N° document': '',
                'Description': f"Afriquia Carte {carte} - GAZOIL",
                'MODALITÉ Code': r['modalite'],
                'Salarie Code': r['salarie'],
                'Affaire Code': r['code affaire'],
                'Division Code': '',
                'Montant débit': montant,
                'Montant crédit': 0,
                'Montant': montant
            })

    lignes_charge.append({
        'Nom de la feuille': '',
        'Date comptabilisation': date_compta,
        'Date TVA': date_compta,
        'Type compte': 'Fournisseur',
        'N° compte': 'F2A021',
        'N° document': '',
        'Description': "Fournisseur Afriquia",
        'MODALITÉ Code': '',
        'Salarie Code': '',
        'Affaire Code': '',
        'Division Code': 'MULTIDIVISION',
        'Montant débit': 0,
        'Montant crédit': total_fournisseur,
        'Montant': -total_fournisseur
    })

    df_final = pd.DataFrame(lignes_charge)

    colonnes_restantes = [
        'Groupe compta. produit TVA','Code devise','N° doc. externe','Type document',
        'Groupe de comptabilisation','Groupe compta. marché TVA',"Date d'échéance",
        'Trans. tripartite UE','Type compta. TVA','Groupe compta. marché',
        'Groupe compta. produit','Prorata de déduction','Différence prorata TVA',
        'N° compte TVA Correctif','Payment Method','% prorata de déduction',
        'Montant DS','Type compte, contrepartie','Nom du compte',
        'N° compte contrepartie','Type compta. contrepartie',
        'Groupe compta. marché contr.','Groupe compta. produit contr.',
        'Code échelonnement','Correction','Commentaire','Lettre Sage',
        'Advance','Nature Code','Recouvreur Code','Commercial Code',
        'Type client Code','NumberOfJournalRecords','Débit total',
        'Crédit total','Solde','Solde final'
    ]

    for col in colonnes_restantes:
        if col not in df_final.columns:
            df_final[col] = ''

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="écritures comptables")
        df_nouvelles_cartes.to_excel(writer, index=False, sheet_name="nouvelles cartes")

    st.subheader("Aperçu écritures comptables")
    st.dataframe(df_final)

    st.download_button(
        "Télécharger fichier Excel",
        buffer.getvalue(),
        "ecritures_comptables_afriquia.xlsx"
    )
