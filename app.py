import datetime as dt

import streamlit as st

import plan_ics as p

st.set_page_config(page_title="Plan d'entraînement → calendrier", page_icon="📅")
st.title("Plan d'entraînement → calendrier")
st.write("Dépose le fichier de ton entraîneur et récupère un fichier `.ics` "
         "à ouvrir sur ton téléphone ou ton ordinateur.")


@st.cache_data(show_spinner="Lecture du fichier…")
def lire_plan(contenu: bytes, nom: str):
    return p.parser_plan(p.lire_texte(contenu, nom))


fichier = st.file_uploader("Fichier du plan (.doc ou .docx)", type=["doc", "docx"])

col1, col2 = st.columns(2)
heure = col1.time_input("Heure de début par défaut", dt.time(18, 0))
duree = col2.number_input("Durée par défaut (minutes)", 15, 240, 60, step=15)
inclure_repos = st.checkbox("Inclure les jours de repos dans le calendrier", value=False)

if fichier is not None:
    try:
        plan = lire_plan(fichier.getvalue(), fichier.name)
    except Exception as e:
        st.error(f"Impossible de lire le fichier : {e}")
        st.stop()

    if not plan:
        st.warning("Aucune séance détectée. Vérifie que le fichier a le format habituel "
                   "(« Semaine du … » puis « Séance 1 (samedi) … »).")
        st.stop()

    evts = p.evenements(plan, inclure_repos, (heure.hour, heure.minute), int(duree))
    st.success(f"{len(evts)} événements prêts.")

    st.dataframe(
        [{"Date": e["debut"].strftime("%a %d/%m/%Y %H:%M"), "Titre": e["titre"]}
         for e in evts],
        use_container_width=True, hide_index=True,
    )

    st.download_button(
        "Télécharger le calendrier (.ics)",
        data=p.generer_ics(evts),
        file_name="plan_entrainement.ics",
        mime="text/calendar",
    )
    st.caption("Sur iPhone : ouvre le fichier téléchargé, puis touche « Ajouter tout ». "
               "Rien n'est conservé sur le serveur.")
