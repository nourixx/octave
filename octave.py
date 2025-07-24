import pandas as pd
import streamlit as st
from PIL import Image
import unicodedata
import re
import openpyxl
from openpyxl.utils import get_column_letter
st.set_page_config(page_title="OCTAVE", layout="wide")
PASSWORD = "adasa2024"

# Stocker l'état d'authentification
if "auth_ok" not in st.session_state:
    st.session_state.auth_ok = False

# Si non authentifié, afficher le champ
if not st.session_state.auth_ok:
    password = st.text_input("🔒 Entrez le mot de passe :", type="password")
    
    if password == PASSWORD:
        st.session_state.auth_ok = True
        st.rerun()  # ✅ Version stable dans Streamlit 1.45.1
    elif password != "":
        st.error("❌ Mot de passe incorrect.")
    else:
        st.info("🔐 Veuillez entrer le mot de passe.")
    
    st.stop()

# ✅ L'utilisateur est validé, l'app peut s'afficher normalement
st.success("🔓 Accès autorisé. Bienvenue dans Octave !")
# ========== logo =========

logo_path = "loogo.png"  
try:
    logo = Image.open(logo_path)
    st.image(logo, width=150)  
except FileNotFoundError:
    st.warning("Logo non trouvé, vérifie le nom du fichier.")

st.cache_data.clear()
st.title("BIENVENUE SUR OCTAVE")

tab1, tab2, tab_analyse = st.tabs([
    "💶 Échéances par mois",
    "📍 Totaux par site",
    "📊 Analyse fichier analytique"
])



# ====== Fonctions utilitaires ======
def normaliser_colonnes(df):
    df.columns = [unicodedata.normalize('NFKD', str(c)).encode('ascii', 'ignore').decode('utf-8').strip().upper().replace(" ", "_") for c in df.columns]
    return df

def get_column_name(columns, target_keywords):
    norm = lambda s: re.sub(r'[^a-z]', '', unicodedata.normalize('NFKD', str(s).lower()))
    for col in columns:
        if any(norm(k) in norm(col) for k in target_keywords):
            return col
    return None

# ===== Onglet 2 : Totaux par site =====
with tab2:
    st.subheader("Totaux calculés par site")

    file2 = st.file_uploader("📂 charger le fichier y_PCA ", type="xlsx", key="resultat")

    if file2:
        df_raw = pd.read_excel(file2, sheet_name="Résultat")
        df_totaux = df_raw[df_raw["Unnamed: 1"] == "Total ADASA"].copy()

        df_totaux = df_totaux.rename(columns={
            'Données': 'Q',
            'Unnamed: 17': 'R',
            'Unnamed: 18': 'S',
            'Unnamed: 19': 'T',
            'Unnamed: 20': 'U',
            'Unnamed: 21': 'V',
            'Unnamed: 22': 'W',
            'Unnamed: 0': 'NOM_SITE'
        })

        for col in ['Q', 'R', 'S', 'T', 'U', 'V', 'W']:
            df_totaux[col] = pd.to_numeric(df_totaux[col], errors='coerce').fillna(0)

        df_totaux["TOTAL_SITE_CALCULE"] = (
            df_totaux["Q"] - df_totaux["R"] + df_totaux["S"] + df_totaux["T"]
            - df_totaux["U"] - df_totaux["V"] + df_totaux["W"]
        )

        site_names = []
        for idx in df_totaux.index:
            nom_site = None
            i = idx - 1
            while i >= 0:
                val = df_raw.at[i, 'Unnamed: 0']
                if pd.notna(val):
                    nom_site = val
                    break
                i -= 1
            site_names.append(nom_site)

        df_totaux["NOM_SITE"] = site_names
        df_totaux["NOM_SITE"] = df_totaux["NOM_SITE"].apply(
            lambda x: "FRANCAS (tous sites)" if isinstance(x, str) and x.upper().startswith("FRANCAS") else x
        )
        df_totaux = df_totaux.groupby("NOM_SITE", as_index=False)["TOTAL_SITE_CALCULE"].sum()

        st.dataframe(df_totaux)

        ca_totale = df_totaux["TOTAL_SITE_CALCULE"].sum()
        st.metric("💰 Chiffre d'affaires total", f"{ca_totale:,.2f} €")

        import plotly.express as px
        fig = px.pie(
            df_totaux,
            names="NOM_SITE",
            values="TOTAL_SITE_CALCULE",
            title="Répartition du CA par site",
            hole=0.4
        )
        st.plotly_chart(fig, use_container_width=True)
# ===== Onglet 1 : Échéances par mois =====
with tab1:
    st.subheader("Analyse des échéances mensuelles")

    file1 = st.file_uploader("📂 Charger le fichier y_Pilotage", type="xlsx", key="echeance")

    def calculer_echeances_par_mois(df):
        df = normaliser_colonnes(df)

        col_nom = get_column_name(df.columns, ['nom apprenant', 'nom_prenom_apprenant'])
        col_debut_contrat = get_column_name(df.columns, ['date debut contrat'])
        col_fin_contrat = get_column_name(df.columns, ['date fin contrat'])
        col_ordre = get_column_name(df.columns, ['numero ordre echeance'])
        col_debut_echeance = get_column_name(df.columns, ['date debut echeance'])
        col_fin_echeance = get_column_name(df.columns, ['date fin echeance'])
        col_valeur = get_column_name(df.columns, ['montant echeance', 'valeur echeance'])

        colonnes = [col_nom, col_debut_contrat, col_fin_contrat, col_ordre, col_debut_echeance, col_fin_echeance, col_valeur]

        if None in colonnes:
            raise ValueError("Le fichier ne contient pas toutes les colonnes nécessaires.")

        df_filtered = df[colonnes].copy()
        df_filtered.columns = [
            'NOM_PRENOM_APPRENANT', 'DATE_DEBUT_CONTRAT', 'DATE_FIN_CONTRAT',
            'NUMERO_ORDRE_ECHEANCE', 'DATE_DEBUT_ECHEANCE', 'DATE_FIN_ECHEANCE', 'VALEUR_ECHEANCE']

        df_filtered['DATE_FIN_CONTRAT'] = pd.to_datetime(df_filtered['DATE_FIN_CONTRAT'], errors='coerce')
        df_filtered['DATE_FIN_ECHEANCE'] = pd.to_datetime(df_filtered['DATE_FIN_ECHEANCE'], errors='coerce')
        df_filtered['DATE_DEBUT_ECHEANCE'] = pd.to_datetime(df_filtered['DATE_DEBUT_ECHEANCE'], errors='coerce')

        count_echeances = df_filtered.groupby('NOM_PRENOM_APPRENANT')['NUMERO_ORDRE_ECHEANCE'].count()
        apprenants_2_echeances = count_echeances[count_echeances == 2].index

        def choisir_date(row):
            if row['NOM_PRENOM_APPRENANT'] in apprenants_2_echeances and row['NUMERO_ORDRE_ECHEANCE'] == 2:
                return row['DATE_FIN_ECHEANCE']
            else:
                return row['DATE_DEBUT_ECHEANCE']

        df_filtered['DATE'] = df_filtered.apply(choisir_date, axis=1)
        df_filtered['Mois'] = df_filtered['DATE'].dt.to_period('M')
        monthly_totals = df_filtered.groupby('Mois')['VALEUR_ECHEANCE'].sum().reset_index()
        monthly_totals['Mois'] = monthly_totals['Mois'].apply(lambda p: p.to_timestamp().strftime('%B %Y'))

        nb_apprenants = df_filtered['NOM_PRENOM_APPRENANT'].nunique()
        nb_lignes = len(df_filtered)
        total_valeur = df_filtered['VALEUR_ECHEANCE'].sum()

        return monthly_totals, nb_apprenants, nb_lignes, total_valeur, df_filtered

    if file1:
        try:
            df1 = pd.read_excel(file1, sheet_name=0, header=1)
            res1, nb_app, nb_lignes, total, df_base = calculer_echeances_par_mois(df1)

            st.success("✅ Fichier traité avec succès !")
            st.write("📦 Nombre d'échéances :", nb_lignes)
            st.write("👨‍🎓 Nombre d'apprenants :", nb_app)
            st.write("💶 Montant total toutes échéances : {:,.2f} €".format(total))

            res1["ANNEE"] = res1["Mois"].apply(lambda x: int(x.split()[-1]) if isinstance(x, str) else None)
            annees_dispo = sorted(res1["ANNEE"].dropna().unique())
            annee_choisie = st.selectbox("📅 Filtrer le tableau par année", annees_dispo, index=len(annees_dispo)-1)

            res1_filtré = res1[res1["ANNEE"] == annee_choisie]
            # Ajouter une ligne "Total"
            total_general = res1_filtré["VALEUR_ECHEANCE"].sum()
            ligne_total = pd.DataFrame({ "Mois": ["TOTAL"],"VALEUR_ECHEANCE": [total_general],"ANNEE": [annee_choisie]})

            res1_filtré_total = pd.concat([res1_filtré, ligne_total], ignore_index=True)
            def color_ligne_total(row):
                if row["Mois"] == "TOTAL":
                    return ['background-color: #fff3b0'] * len(row)  # Jaune pâle
                else:
                    return [''] * len(row)

            styled_df = res1_filtré_total.drop(columns=["ANNEE"]).style.apply(color_ligne_total, axis=1)
            st.dataframe(styled_df, use_container_width=True, height=400)

            df_base["ANNEE"] = df_base["DATE"].dt.year
            df_base["MOIS"] = df_base["DATE"].dt.month

            df_filtre = df_base[df_base["ANNEE"] == annee_choisie]
            df_mensuel = df_filtre.groupby("MOIS")["VALEUR_ECHEANCE"].sum().reset_index()

            st.subheader(f"📊 Échéances par mois - année {annee_choisie}")
            st.bar_chart(df_mensuel.set_index("MOIS"))

        except Exception as e:
            st.error("Erreur : {}".format(e))

# ===== Onglet analyse fichier analytique =====
with tab_analyse:
    st.header("Budget du CFA")

    uploaded_file = st.file_uploader("📤 Charger la balance analytique", type=["xlsx"])

    if uploaded_file:
        try:
            # Chargement du fichier
            xls = pd.ExcelFile(uploaded_file)
            df = xls.parse(xls.sheet_names[0])  # première feuille automatiquement

            # Nettoyage : on garde que les lignes d'écriture (pas les totaux gris)
            df = df[df["Type"] == "Lignes d'écritures"]

            # Conversion des montants
            for col in ["Débit", "Crédit", "Solde progressif"]:
                df[col] = pd.to_numeric(df[col], errors="coerce")

            # Conversion de Compte général pour filtrage
            df["Compte général"] = pd.to_numeric(df["Compte général"], errors="coerce")

            # Filtrage spécifique pour CFA : ignorer les comptes de 611201 à 611215
            df = df[~((df["Code journal"] == "CFA") & (df["Compte général"].between(611201, 611215)))]

            # Agrégation par Code journal
            resume = df.groupby("Code journal")[["Débit", "Crédit", "Solde progressif"]].sum().reset_index()

            # Calcul du total global
            total_row = pd.DataFrame({
                "Code journal": ["TOTAL"],
                "Débit": [resume["Débit"].sum()],
                "Crédit": [resume["Crédit"].sum()],
                "Solde progressif": [resume["Solde progressif"].sum()]
            })

            resume = pd.concat([resume, total_row], ignore_index=True)

            # Fonction de style pour mettre en valeur la ligne TOTAL
            def highlight_total(row):
                if row["Code journal"] == "TOTAL":
                    return ["background-color: #ffe599; font-weight: bold"] * len(row)
                else:
                    return [""] * len(row)

            # Affichage
            st.dataframe(resume.style.apply(highlight_total, axis=1).format({
                "Débit": "{:,.2f}".format,
                "Crédit": "{:,.2f}".format,
                "Solde progressif": "{:,.2f}".format
            }), use_container_width=True)

        except Exception as e:
            st.error(f"Erreur lors de l'analyse : {e}")
    else:
        st.info("Veuillez charger un fichier Excel pour afficher les données.")



