import streamlit as st
import pdfplumber
import pandas as pd
import io

# --- Fonction pour convertir les quantités ---
def convertir_quantite(val):
    if not val:
        return 0
    try:
        val = val.replace(".", ",")
        return float(val.replace(",", "."))
    except:
        return 0

# --- Fonction pour extraire les produits d'un PDF ---
def extraire_produits_pdf(file):
    produits = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row and row[0] and row[0].isdigit() and len(row[0]) == 6:
                        code = row[0]
                        libelle = row[1]
                        quantite = convertir_quantite(row[6])
                        produits.append([code, libelle, quantite])
    return produits


# --- Fonction pour extraire unité/boite depuis le PDF ---
def extraire_unite_boite(file):
    data = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row and row[0] and row[0].isdigit():
                        code = row[0]
                        unite = int(row[1]) if row[1].isdigit() else 1
                        data.append([code, unite])
    return pd.DataFrame(data, columns=["Code", "Unités par Boîte"])


# --- PAGE CONFIG ---
st.set_page_config(page_title="Gestion des Bons de Commande - Imandis Trading",
                   page_icon="📦",
                   layout="wide")

# --- STYLE CSS PERSONNALISÉ ---
st.markdown("""
    <style>
    .main {
        background-color: #F5F5F5;
        padding: 20px;
    }
    .title {
        text-align: center;
        color: #B71C1C;
        font-size: 36px;
        font-weight: bold;
    }
    .subtitle {
        text-align: center;
        color: #757575;
        font-size: 20px;
        margin-bottom: 20px;
    }
    .stButton>button {
        background-color: #B71C1C;
        color: white;
        border-radius: 10px;
        padding: 0.6em 1.2em;
        font-weight: bold;
    }
    .stDownloadButton>button {
        background-color: #2E7D32;
        color: white;
        border-radius: 10px;
        padding: 0.6em 1.2em;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# --- HEADER ---
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.image("image.png", width=180)
    st.markdown("<div class='title'>📦 Gestion des Bons de Commande</div>", unsafe_allow_html=True)
    st.markdown("<div class='subtitle'>Imandis Trading</div>", unsafe_allow_html=True)

st.markdown("---")

# --- DESCRIPTION ---
st.info("**Importez vos bons de commande PDF** et obtenez automatiquement les totaux par produit, avec le nombre de boîtes calculé (formule Excel), prêts à exporter en Excel.")

# --- UPLOAD ---
uploaded_files = st.file_uploader(
    "📂 Sélectionnez vos bons de commande PDF",
    type="pdf",
    accept_multiple_files=True
)

unite_file = st.file_uploader(
    "📂 Sélectionnez le fichier PDF des unités par boîte",
    type="pdf"
)

if uploaded_files and unite_file:
    tous_produits = []

    with st.spinner("🔍 Extraction des données en cours..."):
        for file in uploaded_files:
            tous_produits.extend(extraire_produits_pdf(file))

    if tous_produits:
        df = pd.DataFrame(tous_produits, columns=["Code", "Libellé Produit", "Quantité Commandée (UC)"])
        totaux = df.groupby(["Code", "Libellé Produit"], as_index=False)["Quantité Commandée (UC)"].sum()

        # Extraction des unités par boîte
        df_unites = extraire_unite_boite(unite_file)

        # Fusion
        df_final = pd.merge(totaux, df_unites, on="Code", how="left")

        # Renommer la colonne Code
        df_final = df_final.rename(columns={"Code": "Code Article"})

        # --- Calcul du nombre de boîtes pour l'affichage web ---
        df_final["Nombre de Boîtes"] = df_final["Quantité Commandée (UC)"] / df_final["Unités par Boîte"]

        st.success("✅ Extraction terminée avec succès !")

        # Affichage tableau web
        with st.expander("📊 Voir les résultats détaillés", expanded=True):
            st.dataframe(df_final, use_container_width=True)

        # --- Export Excel avec FORMULE dans "Nombre de Boîtes" ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # On écrit sans la colonne calculée d'abord
            df_export = df_final[["Code Article", "Libellé Produit", "Quantité Commandée (UC)", "Unités par Boîte"]]
            df_export.to_excel(writer, index=False, sheet_name="Totaux Produits")

            workbook  = writer.book
            worksheet = writer.sheets["Totaux Produits"]

            # Insérer le titre de la colonne
            col_formula = df_export.shape[1]  # position de la nouvelle colonne
            worksheet.write(0, col_formula, "Nombre de Boîtes")

            # Ajouter la formule dans chaque ligne
            for row in range(1, len(df_export) + 1):
                worksheet.write_formula(row, col_formula, f"=C{row+1}/D{row+1}")

        output.seek(0)

        st.download_button(
            label="Télécharger le fichier Excel",
            data=output,
            file_name="totaux_commandes_boites.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠ Aucun produit trouvé dans les fichiers importés.")
