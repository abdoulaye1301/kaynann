import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
from openpyxl import load_workbook
import io

st.set_page_config(
    page_title="AFRIKA LEYRI", layout="wide", page_icon="ndao abdoulaye.png"
)
profil = Image.open("Logo Afrika Leyri.png")
st.logo(profil)


# --- Authentification simple ---
USER = "AFRIKA LEYRI"
PASSWORD = "Afrikaleyri2025"

if "authentifie" not in st.session_state:
    st.sidebar.session_state.authentifie = False

  
    # --- Navigation ---
#page = st.sidebar.radio("📁 Menu de navigation", ["Données", "Tableau de bord"])
# URL de récupération des données en CSV
donnee = pd.read_excel(f"https://kf.kobotoolbox.org/api/v2/assets/a3RSyGfABzRzmSL8qNVDg9/export-settings/esjBm7RAEkEEwJCzVDkgqZa/data.xlsx")


# Charger la feuille sélectionnée
nomscol1=["Nom de l'entreprise","Prenom & Nom répondant","Fonction du répondant",
        "Telephone répondant","Moyenne de Relance Effectuée", "Réaction globale",
        "Prix recommandez","Détails & Commentaires du prospect","Prochaine Action à Mener",
        "Date prévue pour la prochaine action"]
# Définir les chemins des fichiers source et destination
base1=donnee[nomscol1]
base1["Date"] = pd.to_datetime(donnee["_submission_time"])
base1["Agent"] = donnee["_submitted_by"].apply(lambda x: "NGOULLE THIOUNE" if x== "ngoulle_thioune" 
                                               else ("FATOU BINTOU DIALLO" if x=="fatou_bintou_diallo" 
                                                     else ("ADJAB LUCIDE ALAINA" if x=="adjab_lucide_alaina" 
                                                           else "SIMONE MANDIAME")))
nomscol=["Date","Agent","Nom de l'entreprise","Prenom & Nom répondant","Fonction du répondant",
        "Telephone répondant","Moyenne de Relance Effectuée", "Réaction globale",
        "Prix recommandez","Détails & Commentaires du prospect","Prochaine Action à Mener",
        "Date prévue pour la prochaine action"]
base=base1[nomscol]

# Définir les bornes du slider
base["Date"] = base["Date"].dt.date
min_date = min(base["Date"])
max_date = max(base["Date"])

# Slider Streamlit pour filtrer une plage de dates
colo=st.columns(2)
start_date, end_date = colo[0].slider(
    "Sélectionnez une plage de dates",
    min_value=min_date,
    max_value=max_date,
    value=(min_date, max_date),  # valeur par défaut (tout)
    format="DD/MM/YYYY"
)
agent_filter = colo[1].multiselect(
    "Sélectionnez le(s) nom(s) d'agent(s)",base["Agent"].unique()
)
# Filtrer les données selon la plage sélectionnée
base = base[(base["Date"] >= start_date) & (base["Date"] <= end_date)]
# Appliquer le filtre par agent
if agent_filter:
    base = base[base["Agent"].isin(agent_filter)]


def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Données')
    output.seek(0)
    return output

def login():
    st.sidebar.title("🔐 Connexion")
    with st.sidebar.form("login_form"):
        user = st.text_input("Nom d'utilisateur")
        pwd = st.text_input("Mot de passe", type="password")
        submit = st.form_submit_button("Connexion")
        if submit:
            if user == USER and pwd == PASSWORD:
                st.session_state.authentifie = True
                st.success("Bienvenue dans l'application Afrika Leyri !")
            else:
                st.error("❌ Identifiants incorrects")
                st.warning("Veuillez vous connecter pour accéder aux données.")
                st.stop()
# --- Onction de vuisualisation des données ---

def visualiser_donnees(base):
        # Agrégation par jour
    evolution = base.groupby("Date")
    
    st.markdown(f"<h2 style='text-align: center;'>!---------- EVOLUTION DES RELANCES ----------!</h4><br>", unsafe_allow_html=True)

    col= st.columns(5)
    col[0].metric("Nombre de relances totales", base["Telephone répondant"].nunique())
    col[1].metric("NGOULLE THIOUNE", (base["Agent"] == "NGOULLE THIOUNE").sum())
    col[2].metric("FATOU BINTOU DIALLO", (base["Agent"] == "FATOU BINTOU DIALLO").sum())
    col[3].metric("ADJAB LUCIDE ALAINA", (base["Agent"] == "ADJAB LUCIDE ALAINA").sum())
    col[4].metric("SIMONE MANDIAME", (base["Agent"] == "SIMONE MANDIAME").sum())

    
    # Afficher les résultats
    
    st.markdown(f"<h3 style='text-align: center;'>!---------- Visualisation des données ----------!</h4><br>", unsafe_allow_html=True)
    
    st.dataframe(base.sort_values(by=["Date","Agent"], ascending=False))

    # Téléchargement des données en format Excel
    excel_data = to_excel(base)
    col=st.columns(3)
    if start_date==end_date:
        col[0].download_button(
            label="📄 Télécharger les données",
            data=excel_data,
            file_name=f"Données KAYNANN du {start_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        col[1].button("🔄 Actualiser les données")
        col[2].button("Plus de détails", on_click=tableau_de_bord, args=(base,))
    else:
        col[0].download_button(
            label="📄 Télécharger les données",
            data=excel_data,
            file_name=f"Données KAYNANN du {start_date} au {end_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        col[1].button("🔄 Actualiser les données")
        col[2].button("Plus de détails", on_click=tableau_de_bord, args=(base,))
# --- Fonction de tableau de bord ---
def tableau_de_bord(base):
            # Agrégation par jour
    evolution = base.groupby("Date")
    

    st.subheader("📊 Évolution des ventes et installations des commerciaux")
    col= st.columns(4)
    col[0].metric("Nombre de relances totales", base["Telephone répondant"].nunique())
    col[1].metric("NGOULLE THIONE", (base["Agent"] == "NGOULLE THIONE").sum())
    #col[2].metric("🔢 Nombre de packs de 10000", int((base["Montant"] >= 10000).sum()))
    #col[3].metric("🔢 Nombre d'installations", int((base["Operation"] == "Installation").sum()))
    #colone= st.columns(3)
    #colone[0].metric("💴 CA Total Réalisé", f"{base["Montant"].sum():,.0f}".replace(",", " ")+" XOF")
    #colone[1].metric("💴 CA des packs de 5000", f"{base[base["Montant"] < 10000]["Montant"].sum():,.0f}".replace(",", " ")+" XOF")
    #colone[2].metric("💴 CA des packs de 10000", f"{base[base["Montant"] >= 10000]["Montant"].sum():,.0f}".replace(",", " ")+" XOF")

        # Evolution des ventes
    """evolution["Montant_affiche"] = evolution["Montant"].map(lambda x: f"{x:,.0f}".replace(",", " "))
    fig = px.line(evolution, x="Date", y="Montant",
                  text="Montant_affiche",
                    title="CA des ventes par jour",
                    markers=True)
    fig.update_traces(marker=dict(size=8, color="red", line=dict(width=2, color="DarkSlateGrey")),
                      textfont=dict(size=12, color="white"),
    textposition="top center")
    fig.update_layout(xaxis=dict(tickformat="%d-%m",
                      tickangle=-45, 
                      tickvals=base["Date"].unique()))
    st.plotly_chart(fig, use_container_width=True)


    # Représentation graphique avec plotly
    colon= st.columns(2)
    gra=px.histogram(pack, x="Numéro_Pack", y="Nombre de pack",
                    title="Nombre de ventes par pack")
    gra.update_traces(texttemplate='%{y}', textposition='auto')
    gra.update_layout( xaxis_title="Numéro de Pack",xaxis=dict(tickmode='linear',dtick=1), yaxis_title="Nombre de Packs")
    colon[0].plotly_chart(gra, use_container_width=True)


# graphique des operations
    colon[1].write("Répartition des opérations")
    colon[1].plotly_chart(px.pie(base, names="Operation"), use_container_width=True,title="Répartition des opérations")  





  

    # Performance des agents
    donnee_agre = base.groupby(["Prenom Nom","Operation"]).agg(
        {"Telephone_Client": "count", "Numéro_Pack": "count", "Montant": "sum"}
        ).reset_index()
    donnee_agre = donnee_agre.rename(
    columns={
        "Telephone_Client": "Nombre d'installations",
        "Numéro_Pack": "Nombre de Packs vendus",
        "Montant": "Montant",
    }
    )
    st.subheader("Récapitulatif des ventes et installations des commerciaux")
    st.dataframe(donnee_agre.sort_values(by=["Prenom Nom", "Montant"], ascending=False))
    # Performance des agents
    donnee_vente = base.groupby(["Zone"]).agg(
        {"Numéro_Pack": "count", "Montant": "sum"}
        ).reset_index()
    st.subheader("Récapitulatif des ventes par Zone")
    st.dataframe(donnee_vente.sort_values(by=["Numéro_Pack", "Montant"], ascending=False))
"""
  
# --- Page 1 : Visualisation simple ---
#if page == "Données":
visualiser_donnees(base)
#if not st.session_state.authentifie:
 #   login()
  #  if st.session_state.authentifie:
   #     visualiser_donnees(base)
#else:
 #   visualiser_donnees(base)
        


# --- Page 2 : Tableau de bord (protégé) ---
#elif page == "Tableau de bord":
 #   if not st.session_state.authentifie:
  #      login()
   #     if st.session_state.authentifie:
    #        tableau_de_bord(base)
    #else:
     #   tableau_de_bord(base)
        
        # --- Message de bienvenue si connecté --
           