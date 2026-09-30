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
#USER = "AFRIKA LEYRI"
PASSWORD = "Afr"

if "authentifie" not in st.session_state:
    st.session_state.authentifie = False

  
    # --- Navigation ---
page = st.sidebar.radio("📁 Menu de navigation", ["KAYNANN", "AFRIKA LEYRI"])
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
                                                           else ("ROUGUIATOU DANFACA" if x=="danfaca_rougui" 
                                                                 else "SIMONE MANDIAME"))))
nomscol=["Date","Agent","Nom de l'entreprise","Prenom & Nom répondant","Fonction du répondant",
        "Telephone répondant","Moyenne de Relance Effectuée", "Réaction globale",
        "Prix recommandez","Détails & Commentaires du prospect","Prochaine Action à Mener",
        "Date prévue pour la prochaine action"]
base=base1[nomscol]

# Définir les bornes du slider
base["Date"] = base["Date"].dt.date
base_kaynann=base.drop(columns=["Agent"])
min_date = min(base["Date"])
max_date = max(base["Date"])


def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Données')
    output.seek(0)
    return output

def login():
    #st.sidebar.title("🔐 Connexion")
    with st.sidebar.form("login_form"):
        #user = st.text_input("Nom d'utilisateur")
        pwd = st.text_input("Mot de passe", type="password")
        submit = st.form_submit_button("Connexion")
        if submit:
            if  pwd == PASSWORD:
                st.session_state.authentifie = True
                st.success("✅ Connexion réussie !")
            else:
                st.error("❌ Identifiants incorrects")
                st.warning("Veuillez vous connecter pour accéder aux données.")
                st.stop()
# --- Onction de vuisualisation des données ---

def kaynann(base):
    
    st.markdown(f"<h2 style='text-align: center;'>!---------- 📊 EVOLUTION DES RELANCES ----------!</h4><br>", unsafe_allow_html=True)
    col= st.columns(2)
    # Slider Streamlit pour filtrer une plage de dates
    start_date, end_date = col[0].slider(
        "Sélectionnez une plage de dates",
        min_value=min_date,
        max_value=max_date,
        value=(min_date, max_date),  # valeur par défaut (tout)
        format="DD/MM/YYYY"
    )
    # Filtrer les données selon la plage sélectionnée
    base = base[(base["Date"] >= start_date) & (base["Date"] <= end_date)]

    
    # Injection de style CSS pour centrer le contenu du widget metric dans cette colonne spécifique
    col[1].markdown(
        """
        <style>[data-testid="stMetric"] {text-align: center;}
        [data-testid="stMetricLabel"] {display: flex;justify-content: center;}
        [data-testid="stMetricValue"] {display: flex;justify-content: center;}
        </style>
        """,
        unsafe_allow_html=True
    )
    col[1].metric("Nombre de relances totales", base["Telephone répondant"].nunique())

      # Représentation graphique avec plotly

    #colon = st.columns(1)
    pa = base.groupby("Réaction globale").size().reset_index(name="Nombre de Prospects")

    # Utilisation de px.bar au lieu de px.histogram et passage de y="Nombre de Prospects"
    gra = px.bar(
        pa, 
        x="Réaction globale", 
        y="Nombre de Prospects",
        title="Nombre de prospects par réaction",
        color="Réaction globale"
    )
    # Centrer le titre du graphique (title_x=0.5)
    gra.update_layout(title_x=0.5)
    gra.update_traces(texttemplate='%{y}', 
                      textposition='auto',
                      textfont_size=16)
    gra.update_layout(
        xaxis_title="Réaction",
        yaxis_title="Nombre de Prospects",
        showlegend=False,
        yaxis=dict(
        showticklabels=False,  # Cache les chiffres de l'axe Y
        title=None             # Supprime le titre de l'axe Y ("Nombre de Prospects")
    )
    )

    st.plotly_chart(gra, use_container_width=True)
    # graphique des operations
    #colon[1].write("Répartition des opérations")
    #colon[1].plotly_chart(px.pie(base_kaynann, names="Operation"), use_container_width=True,title="Répartition des opérations")  
    
    
    
    # Afficher les résultats
    
    st.markdown(f"<h3 style='text-align: center;'>!---------- Visualisation des données ----------!</h4><br>", unsafe_allow_html=True)
    colonee= st.columns(3)
    agent_filter = colonee[1].multiselect(
            "Réaction du prospect",base["Réaction globale"].unique()
        )
    # Appliquer le filtre par agent
    if agent_filter:
        base = base[base["Réaction globale"].isin(agent_filter)]
    
    st.dataframe(base.sort_values(by=["Date"], ascending=False))

        #col[2].button("Plus de détails", on_click=tableau_de_bord, args=(base,))
# --- Fonction de tableau de bord ---
def afrikaleyri(base):
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
            # Agrégation par jour
    evolution = base.groupby("Date")
    
    st.markdown(f"<h2 style='text-align: center;'>!---------- 📊 EVOLUTION DES RELANCES ----------!</h4><br>", unsafe_allow_html=True)

    col= st.columns(6)
    # Injection de style CSS pour centrer le contenu du widget metric dans cette colonne spécifique
    col[0].markdown(
        """
        <style>[data-testid="stMetric"] {text-align: center;}
        [data-testid="stMetricLabel"] {display: flex;justify-content: center;}
        [data-testid="stMetricValue"] {display: flex;justify-content: center;}
        </style>
        """,
        unsafe_allow_html=True
    )
    col[0].metric("Nbr de relances totales", base["Telephone répondant"].nunique())
    col[1].markdown(
        """
        <style>[data-testid="stMetric"] {text-align: center;}
        [data-testid="stMetricLabel"] {display: flex;justify-content: center;}
        [data-testid="stMetricValue"] {display: flex;justify-content: center;}
        </style>
        """,
        unsafe_allow_html=True
    )
    col[1].metric("NGOULLE THIOUNE", (base["Agent"] == "NGOULLE THIOUNE").sum())
    col[2].markdown(
        """
        <style>[data-testid="stMetric"] {text-align: center;}
        [data-testid="stMetricLabel"] {display: flex;justify-content: center;}
        [data-testid="stMetricValue"] {display: flex;justify-content: center;}
        </style>
        """,
        unsafe_allow_html=True
    )
    col[2].metric("FATOU BINTOU DIALLO", (base["Agent"] == "FATOU BINTOU DIALLO").sum())
    col[3].markdown(
        """
        <style>[data-testid="stMetric"] {text-align: center;}
        [data-testid="stMetricLabel"] {display: flex;justify-content: center;}
        [data-testid="stMetricValue"] {display: flex;justify-content: center;}
        </style>
        """,
        unsafe_allow_html=True
    )
    col[3].metric("ADJAB LUCIDE ALAINA", (base["Agent"] == "ADJAB LUCIDE ALAINA").sum())
    col[4].markdown(
        """
        <style>[data-testid="stMetric"] {text-align: center;}
        [data-testid="stMetricLabel"] {display: flex;justify-content: center;}
        [data-testid="stMetricValue"] {display: flex;justify-content: center;}
        </style>
        """,
        unsafe_allow_html=True
    )
    col[4].metric("SIMONE MANDIAME", (base["Agent"] == "SIMONE MANDIAME").sum())
    col[5].markdown(
            """
            <style>[data-testid="stMetric"] {text-align: center;}
            [data-testid="stMetricLabel"] {display: flex;justify-content: center;}
            [data-testid="stMetricValue"] {display: flex;justify-content: center;}
            </style>
            """,
            unsafe_allow_html=True
        )
    col[5].metric("ROUGUIATOU DANFACA", (base["Agent"] == "ROUGUIATOU DANFACA").sum())

        # Représentation graphique avec plotly

    #colon = st.columns(1)
    pa = base.groupby("Réaction globale").size().reset_index(name="Nombre de Prospects")

    # Utilisation de px.bar au lieu de px.histogram et passage de y="Nombre de Prospects"
    gra = px.bar(
        pa, 
        x="Réaction globale", 
        y="Nombre de Prospects",
        title="Nombre de prospects par réaction",
        color="Réaction globale"
    )
    # Centrer le titre du graphique (title_x=0.5)
    gra.update_layout(title_x=0.5)
    gra.update_traces(texttemplate='%{y}', 
                        textposition='auto',
                        textfont_size=16)
    gra.update_layout(
        xaxis_title="Réaction",
        yaxis_title="Nombre de Prospects",
        showlegend=False,
        yaxis=dict(
        showticklabels=False,  # Cache les chiffres de l'axe Y
        title=None             # Supprime le titre de l'axe Y ("Nombre de Prospects")
    )
    )

    st.plotly_chart(gra, use_container_width=True)
    # graphique des operations
    #colon[1].write("Répartition des opérations")
    #colon[1].plotly_chart(px.pie(base_kaynann, names="Operation"), use_container_width=True,title="Répartition des opérations")  
        
        
    # Afficher les résultats
    
    st.markdown(f"<h3 style='text-align: center;'>!---------- Visualisation des données ----------!</h4><br>", unsafe_allow_html=True)
    colonee= st.columns(3)
    agent_filter = colonee[1].multiselect(
            "Réaction du prospect",base["Réaction globale"].unique()
        )
    # Appliquer le filtre par agent
    if agent_filter:
        base = base[base["Réaction globale"].isin(agent_filter)]
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
        #col[2].button("Plus de détails", on_click=tableau_de_bord, args=(base,))
    else:
        col[0].download_button(
            label="📄 Télécharger les données",
            data=excel_data,
            file_name=f"Données KAYNANN du {start_date} au {end_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        col[1].button("🔄 Actualiser les données")    

  
# --- Page 1 : KAYNANN ---
if page == "KAYNANN":
    kaynann(base_kaynann)  

# --- Page 2 : Tableau de bord (protégé) ---
elif page == "AFRIKA LEYRI":

    if not st.session_state.authentifie:
        st.warning("Veuillez vous connecter pour accéder aux données.")
        login()
        if st.session_state.authentifie:
            afrikaleyri(base)
    else:
        afrikaleyri(base)
