import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
from openpyxl import load_workbook
import io
import textwrap

st.set_page_config(
    page_title="AFRIKA LEYRI", layout="wide", page_icon="ndao abdoulaye.png"
)
profil = Image.open("Logo Afrika Leyri.png")
st.logo(profil)


# --- Authentification simple ---
#USER = "AFRIKA LEYRI"
PASSWORD = "AFR"

if "authentifie" not in st.session_state:
    st.session_state.authentifie = False

  
    # --- Navigation ---
page = st.sidebar.selectbox("📁 Menu de navigation", ["KAYNANN", "AFRIKA LEYRI"])

# URL de récupération des données B2B
donnee = pd.read_excel(f"https://kf.kobotoolbox.org/api/v2/assets/a3RSyGfABzRzmSL8qNVDg9/export-settings/esjBm7RAEkEEwJCzVDkgqZa/data.xlsx")
#st.write(donnee.columns.to_list())

# Charger la feuille sélectionnée
nomscol1=["Nom de l'entreprise","Precisez","Prenom & Nom répondant","Fonction du répondant",
        "Telephone répondant","Moyenne de Relance Effectuée", "Réaction globale",
        "Prix recommandez","Détails & Commentaires du prospect","Prochaine Action à Mener",
        "Date prévue pour la prochaine action"]
# Définir les chemins des fichiers source et destination
base1=donnee[nomscol1]
base1["Date"] = pd.to_datetime(donnee["Date de la propection"])
base1["Agent"] = donnee["_submitted_by"].apply(lambda x: "NGOULLE THIOUNE" if x== "ngoulle_thioune" 
                                               else ("FATOU BINTOU DIALLO" if x=="fatou_bintou_diallo" 
                                                     else ("ADJAB LUCIDE ALAINA" if x=="adjab_lucide_alaina" 
                                                           else ("ROUGUIATOU DANFACA" if x=="danfaca_rougui" 
                                                                 else "SIMONE MANDIAME"))))
nomscol=["Date","Agent","Nom de l'entreprise","Precisez","Prenom & Nom répondant","Fonction du répondant",
        "Telephone répondant","Moyenne de Relance Effectuée", "Réaction globale",
        "Prix recommandez","Détails & Commentaires du prospect","Prochaine Action à Mener",
        "Date prévue pour la prochaine action"]
base=base1[nomscol]

base["Date"] = base["Date"].dt.date
# Définir les bornes du slider
dates_valides = base["Date"].dropna()

if dates_valides.empty:
    st.error("Aucune date valide n'est disponible dans les données.")
    st.stop()

min_date = dates_valides.min()
max_date = dates_valides.max()




base_kaynann=base.drop(columns=["Agent"])
#min_date = min(base["Date"])
#max_date = max(base["Date"])

# =======================CONNEXION=====================================#
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
# ========================B2C====================================#
def kaynann_b2c():
    
    st.markdown(f"<h2 style='text-align: center;'>!---------- 📊 RAPPORT TERRAIN COMMERCIAL ----------!</h4><br>", unsafe_allow_html=True)
        
    # URL de récupération des données B2C
    donnee_muni_usine = pd.read_excel(f"https://kf.kobotoolbox.org/api/v2/assets/aWPJ3oxcYAQyuz94napK8F/export-settings/esxgJ3dtGxH6oAsC8ToDFcj/data.xlsx")
    #st.write(donnee_muni_usine.columns.to_list())
    #st.write(donnee_muni_usine.head())
    colonee_B2C=["Groupe","N° Muni-usine","Information à renseigner","Type utilisateur","Quantité vendue","Consommation d'eau (en litre)"
    ,"Difficultés rencontrés","Diagnostic maintenance","Photo maintenance_URL","Diagnostic hygiène"
    ,"Photo hygiène_URL","Recommandations","Nombre de personnes rencontré"]
    donnee_B2C=donnee_muni_usine[colonee_B2C]
    donnee_B2C["Début"] = donnee_muni_usine["start"].dt.time
    donnee_B2C["Fin"] = donnee_muni_usine["end"].dt.time
    donnee_B2C["Date"] = donnee_muni_usine["today"].dt.date
    donnee_B2C["Zone"] = donnee_muni_usine["username"].apply(lambda x: "RUFISQUE" if x== "aissatou__diouf" else "PARCELLES")
    donnee_B2C["Agent"] = donnee_muni_usine["username"].apply(lambda x: "AISSATOU DIOUF" if x== "aissatou__diouf" else "DAOUDA DIOUF")

   
    min_date_b2c = donnee_B2C["Date"].min()
    max_date_b2c = donnee_B2C["Date"].max()
    col= st.columns(2)
        # Slider Streamlit pour filtrer une plage de dates
    start_date_b2c, end_date_b2c = col[0].slider(
        "Sélectionnez une plage de dates",
        min_value=min_date_b2c,
        max_value=max_date_b2c,
        value=(min_date_b2c, max_date_b2c),  # valeur par défaut (tout)
        format="DD/MM/YYYY"
    )
    colo=col[1].columns(2)
    info=colo[0].selectbox(
        "Information mini-usine & utilisateur",["Tous"] + donnee_B2C["Information à renseigner"].dropna().unique().tolist()
    )
    zone=colo[1].selectbox(
            "Zone",["Tous"] + donnee_B2C["Zone"].dropna().unique().tolist()
        )
    Numero_muni_usine=donnee_B2C["N° Muni-usine"].dropna().unique().tolist()
    #=========donnee calcul conso===============
    nomscol_B2C_cons=["Date","Zone","Agent","Groupe","N° Muni-usine","Quantité vendue","Consommation d'eau (en litre)"]
    
    if info=="Machine":
        nomscol_B2C=["Date","Début","Fin","Zone","Agent","Groupe","N° Muni-usine","Quantité vendue","Consommation d'eau (en litre)",
                     "Diagnostic maintenance","Photo maintenance_URL","Diagnostic hygiène","Photo hygiène_URL"]
    elif info=="Utilisateur":
        nomscol_B2C=["Date","Début","Fin","Zone","Agent","N° Muni-usine",
                     "Type utilisateur","Nombre de personnes rencontré","Difficultés rencontrés",
                     "Recommandations"]
    else: 
        nomscol_B2C=["Date","Début","Fin","Zone","Agent","Groupe","N° Muni-usine","Information à renseigner",
                     "Type utilisateur","Quantité vendue","Consommation d'eau (en litre)",
                     "Diagnostic maintenance","Photo maintenance_URL","Diagnostic hygiène","Photo hygiène_URL",
                     "Nombre de personnes rencontré","Difficultés rencontrés",
                     "Recommandations"]
    # Filtrer les données selon la plage sélectionnée
    if info != "Tous" and zone != "Tous":
        donnee_B2C = donnee_B2C[(donnee_B2C["Information à renseigner"] == info) & (donnee_B2C["Zone"] == zone) & (donnee_B2C["Date"] >= start_date_b2c) & (donnee_B2C["Date"] <= end_date_b2c)]
    elif info != "Tous":
        donnee_B2C = donnee_B2C[(donnee_B2C["Information à renseigner"] == info) & (donnee_B2C["Date"] >= start_date_b2c) & (donnee_B2C["Date"] <= end_date_b2c)]
    elif zone != "Tous":
        donnee_B2C = donnee_B2C[(donnee_B2C["Zone"] == zone) & (donnee_B2C["Date"] >= start_date_b2c) & (donnee_B2C["Date"] <= end_date_b2c)]
    else:
        donnee_B2C = donnee_B2C[(donnee_B2C["Date"] >= start_date_b2c) & (donnee_B2C["Date"] <= end_date_b2c)]
    # Affichage du tableau avec les photos
    donnee_B2C_1=donnee_B2C[nomscol_B2C]

    if info!="Utilisateur":
        nonnee_cons=donnee_B2C[nomscol_B2C_cons]
        # Convertir la consommation en valeur numérique
        nonnee_cons["Consommation d'eau (en litre)"] = pd.to_numeric(
            nonnee_cons["Consommation d'eau (en litre)"],
            errors="coerce"
        )
     # Injection de style CSS pour centrer le contenu du widget metric dans cette colonne spécifique
    colonne= st.columns(2)
    colonne[1].markdown(
        """
        <style>[data-testid="stMetric"] {text-align: center;}
        [data-testid="stMetricLabel"] {display: flex;justify-content: center;}
        [data-testid="stMetricValue"] {display: flex;justify-content: center;}
        </style>
        """,
        unsafe_allow_html=True
    )
    
    if info!="Machine":
        colonne[1].metric("Nombre de personnes rencontrées", int(donnee_B2C_1["Nombre de personnes rencontré"].sum()))
    
    #===============================VOLUME DE LITRES D'EAU CONSOMMÉS PAR LES MINI-USINES========================#
    colonne_quantite= st.columns(2)

    #===============================VOLUME DE LITRES D'EAU CONSOMMÉS PAR LES MINI-USINES========================#
    if info!="Utilisateur":
        #colonne_quantite[0].markdown(f"<h5 style='text-align: center;'>QUANTITE D'EAU VENDU (EN LITRE)</h5><br>", unsafe_allow_html=True)
        
        donnee_agre = (
            nonnee_cons.groupby(["Zone","N° Muni-usine"])
            .agg({"Quantité vendue": "sum"})
            .reset_index()
        )
        donnee_ordre = donnee_agre.sort_values(by=["Zone","Quantité vendue"], ascending=False)
        donnee_calcu=donnee_ordre.copy()
        #donnee_ordre = donnee_ordre[donnee_ordre["Quantité vendue"] > 0]
        #colonne_quantite[0].write(donnee_ordre)


        #===========================COSOMATION CALCULER AVEC LA VOLUME==============================#

        colonne_quantite[0].markdown(f"<h5 style='text-align: center;'>QUANTITE D'EAU VENDUE & DEDUITE</h5><br>", unsafe_allow_html=True)
            

        donne_cal1 = nonnee_cons[
            pd.to_numeric(
                nonnee_cons["Consommation d'eau (en litre)"],
                errors="coerce"
            ).notna()
        ].copy()

        # Convertir les colonnes
        donne_cal1["Consommation d'eau (en litre)"] = pd.to_numeric(
            donne_cal1["Consommation d'eau (en litre)"],
            errors="coerce"
        )

        donne_cal1["Date"] = pd.to_datetime(
            donne_cal1["Date"], errors="coerce"
        )

        # Filtrer les muni-usines et la période sélectionnée
        donne_cal1 = donne_cal1[
            donne_cal1["N° Muni-usine"].isin(Numero_muni_usine)
            & (donne_cal1["Date"] >= pd.to_datetime(start_date_b2c))
            & (donne_cal1["Date"] <= pd.to_datetime(end_date_b2c))
        ].copy()

        # Trier les relevés par muni-usine et par date
        donne_cal1 = donne_cal1.sort_values(
            ["N° Muni-usine", "Date"]
        )

        # Récupérer le premier et le dernier relevé de chaque muni-usine
        premier_releve = donne_cal1.drop_duplicates(
            subset=["N° Muni-usine"],
            keep="first"
        ).set_index("N° Muni-usine")["Consommation d'eau (en litre)"]

        dernier_releve = donne_cal1.drop_duplicates(
            subset=["N° Muni-usine"],
            keep="last"
        ).set_index("N° Muni-usine")["Consommation d'eau (en litre)"]

        # Calculer le volume consommé pour chaque muni-usine
        volume = dernier_releve.subtract(premier_releve)

        # Affecter le résultat au numéro de muni-usine correspondant
        donnee_calcu["Volume initiale"] = (
                    donnee_calcu["N° Muni-usine"].map(premier_releve)
                )
        
        donnee_calcu["Volume finale"] = (
            donnee_calcu["N° Muni-usine"].map(dernier_releve)
        )
        donnee_calcu["Volume déduite"] = (
                    donnee_calcu["N° Muni-usine"].map(volume)
                )

        donnee_calcu = donnee_calcu[
            (donnee_calcu["Volume déduite"].fillna(0) != 0)
            | (donnee_calcu["Quantité vendue"].fillna(0) != 0)
        ]
        # Afficher le résultat
        colonne_quantite[0].write(donnee_calcu)


        #============GRAPHIQUE DU VOLUME DE LITRES D'EAU CONSOMMÉS PAR LES MINI-USINES========================#
        #gra_quantite = px.bar(
         #   donnee_ordre,
          #  x="N° Muni-usine",
           # y="Quantité vendue",
            #title="VOLUME DE LITRES D'EAU VENDUS",
            #color="Zone"
        #)
        #gra_quantite.update_layout(title_x=0.5)
        #st.plotly_chart(gra_quantite, use_container_width=True)
    #===============================NOMBRE DE PERSONNES RENCONTRÉES PAR TYPE D'UTILISATEUR========================#
    if info!="Machine":
        #colon = st.columns(1)
        pa_b2c = (
            donnee_B2C_1.groupby("Type utilisateur", as_index=False)
            ["Nombre de personnes rencontré"]
            .sum()
        )
        #Utilisation de px.bar au lieu de px.histogram et passage de y="Nombre de Prospects"
        gra_b2c = px.bar(
            pa_b2c, 
            x="Type utilisateur", 
            y="Nombre de personnes rencontré",
            title="NOMBRE DE PERSONNES RENCONTRÉS PAR TYPE D'UTILISATEUR",
            color="Type utilisateur",
        text="Nombre de personnes rencontré"
        )
        # Centrer le titre du graphique (title_x=0.5)
        gra_b2c.update_layout(title_x=0.5)
        gra_b2c.update_traces(texttemplate='%{y}', 
                            textposition='auto',
                            textfont_size=16)
        gra_b2c.update_layout(
            xaxis_title="Utilisateur",
            yaxis_title="Nombre de personnes rencontré",
            showlegend=False,
            yaxis=dict(
            showticklabels=False,  # Cache les chiffres de l'axe Y
            title=None             # Supprime le titre de l'axe Y ("Nombre de personnes rencontré")
        )
        )

        colonne_quantite[1].plotly_chart(gra_b2c, use_container_width=True)
        #------------AFFICHAGE DES COLONNES DE TEXTE AVEC RETOUR A LA LIGNE----------------#
    
    st.dataframe(
        donnee_B2C_1,
        column_config={
            "Photo hygiène_URL": st.column_config.ImageColumn(
                "Photo pour l'hygiène",
                help="Photo prise lors du diagnostic pour l'hygiène",
                width="medium"
            ),
            "Photo maintenance_URL": st.column_config.ImageColumn(
                "Photo pour la maintenance",
                help="Photo prise lors du diagnostic pour la maintenance",
                width="medium"
            ),
            "Difficultés rencontrés": st.column_config.TextColumn(
                            "Difficultés rencontrés",
                            width="large"
                        ),
                        "Recommandations": st.column_config.TextColumn(
                            "Recommandations",
                            width="large"
                        )
        },
        hide_index=True,
        use_container_width=True
    )
    # ==========================================#
    
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Données')
    output.seek(0)
    return output


# ======================B2B======================================#
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
    colonee= st.columns(4)
    reaction_filter = colonee[1].multiselect(
            "Réaction du prospect",base["Réaction globale"].unique()
        )
    # Appliquer le filtre par Réaction du prospect
    if reaction_filter:
        base = base[base["Réaction globale"].isin(reaction_filter)]
    
    action_filter = colonee[2].multiselect(
                "Prochaine Action à Mener",base["Prochaine Action à Mener"].unique()
            )
    # Appliquer le filtre par Prochaine Action à Mener
    if action_filter:
        base = base[base["Prochaine Action à Mener"].isin(action_filter)]

    action_filter = colonee[3].multiselect(
                    "Nom de l'entreprise",base["Nom de l'entreprise"].unique()
                )
    # Appliquer le filtre par Nom de l'entreprise
    if action_filter:
        base = base[base["Nom de l'entreprise"].isin(action_filter)]

    st.dataframe(base.sort_values(by=["Date"], ascending=False))
    
        #col[2].button("Plus de détails", on_click=tableau_de_bord, args=(base,))
# --- Fonction de tableau de bord ---
def afrikaleyri(base):
    
    st.markdown(f"<h2 style='text-align: center;'>!---------- 📊 EVOLUTION DES RELANCES ----------!</h4><br>", unsafe_allow_html=True)
    
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
    colonee= st.columns(4)
    reaction_filter = colonee[1].multiselect(
            "Réaction du prospect",base["Réaction globale"].unique()
        )
    action_filter = colonee[2].multiselect(
                    "Prochaine Action à Mener",base["Prochaine Action à Mener"].unique()
                )
    # Appliquer le filtre par Réaction du prospect
    if reaction_filter:
        base = base[base["Réaction globale"].isin(reaction_filter)]
    
    # Appliquer le filtre par Prochaine Action à Mener
    if action_filter:
        base = base[base["Prochaine Action à Mener"].isin(action_filter)]

    action_filter = colonee[3].multiselect(
                    "Nom de l'entreprise",base["Nom de l'entreprise"].unique()
                )
    # Appliquer le filtre par Nom de l'entreprise
    if action_filter:
        base = base[base["Nom de l'entreprise"].isin(action_filter)]
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
    sectio= st.sidebar.radio(
            "Équipe",("B2B", "B2C")
        )
    if sectio == "B2B":
        kaynann(base_kaynann)  
    elif sectio == "B2C":
        kaynann_b2c()

# --- Page 2 : Tableau de bord (protégé) ---
elif page == "AFRIKA LEYRI":

    if not st.session_state.authentifie:
        st.warning("Veuillez vous connecter pour accéder aux données.")
        login()
        if st.session_state.authentifie:
            sectio= st.sidebar.radio(
                    "Équipe",("B2B", "B2C")
                )
            if sectio == "B2B":
                afrikaleyri(base)
            else:
                kaynann_b2c()
    else:
        sectio= st.sidebar.radio(
                "Équipe",("B2B", "B2C")
            )
        if sectio == "B2B":
            afrikaleyri(base)
        else:
            kaynann_b2c()
