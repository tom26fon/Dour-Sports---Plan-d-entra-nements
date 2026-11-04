"""Logique : lecture du plan (.doc/.docx), détection des séances, génération du .ics."""
import io
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
from datetime import datetime, timedelta
from pathlib import Path


# ---------------------------------------------------------------- Lecture ---

def _convertir_doc_en_docx(src: Path) -> Path:
    """Convertit un .doc en .docx (Word sous Windows, sinon LibreOffice)."""
    dst = src.with_suffix(".docx")

    # Méthode 1 : Microsoft Word (Windows uniquement)
    if sys.platform == "win32":
        try:
            import pythoncom  # fourni avec pywin32
            import win32com.client
            pythoncom.CoInitialize()  # nécessaire car Streamlit utilise des threads
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            try:
                doc = word.Documents.Open(str(src))
                doc.SaveAs2(str(dst), FileFormat=16)  # 16 = .docx
                doc.Close()
            finally:
                word.Quit()
            if dst.exists():
                return dst
        except Exception:
            pass  # on essaie LibreOffice

    # Méthode 2 : LibreOffice (serveur Linux ou PC sans Word)
    soffice = (shutil.which("soffice") or shutil.which("libreoffice")
               or r"C:\Program Files\LibreOffice\program\soffice.exe")
    profil = Path(tempfile.gettempdir()) / "lo_profil_plan"
    try:
        subprocess.run(
            [soffice, "--headless", f"-env:UserInstallation={profil.as_uri()}",
             "--convert-to", "docx", "--outdir", str(src.parent), str(src)],
            check=True, capture_output=True, timeout=120,
        )
    except Exception as e:
        raise RuntimeError(
            "Conversion du .doc impossible (ni Word ni LibreOffice disponible). "
            "Enregistre le fichier en .docx et réessaie."
        ) from e
    if not dst.exists():
        raise RuntimeError("La conversion n'a produit aucun fichier .docx.")
    return dst


def lire_texte(contenu: bytes, nom_fichier: str) -> str:
    """Renvoie le texte brut d'un .doc ou .docx fourni sous forme d'octets."""
    from docx import Document  # pip install python-docx

    ext = Path(nom_fichier).suffix.lower()
    # Dossier temporaire supprimé automatiquement : rien n'est conservé.
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / f"plan{ext}"
        src.write_bytes(contenu)
        if ext == ".doc":
            src = _convertir_doc_en_docx(src)
        elif ext != ".docx":
            raise ValueError(f"Format non supporté : {ext}")
        doc = Document(str(src))
        texte = "\n".join(p.text for p in doc.paragraphs)
    return texte.replace("\x0b", "\n").replace("\xa0", " ")


# ---------------------------------------------------------------- Parsing ---

RE_SEMAINE = re.compile(
    r"Semaine du \w+ (\d{2}/\d{2}/\d{4}) au \w+ (\d{2}/\d{2}/\d{4})", re.I
)
RE_SEANCE = re.compile(r"^\s*S[ée]ance\s+(\d+)\s*\((\w+)\)\s*(.*)$", re.I)


def parser_plan(texte: str) -> list[dict]:
    seances, semaine_debut, courante = [], None, None
    for ligne in texte.splitlines():
        m_sem = RE_SEMAINE.search(ligne)
        if m_sem:
            semaine_debut, courante = m_sem.group(1), None
            continue
        m_sea = RE_SEANCE.match(ligne)
        if m_sea and semaine_debut:
            courante = {
                "semaine_debut": semaine_debut,
                "numero": int(m_sea.group(1)),
                "jour": m_sea.group(2).lower(),
                "lignes": [m_sea.group(3).strip()],
            }
            seances.append(courante)
            continue
        if courante is not None and ligne.strip():
            courante["lignes"].append(ligne.strip())
        elif not ligne.strip():
            courante = None
    for s in seances:
        s["description"] = "\n".join(l for l in s.pop("lignes") if l)
    return seances


# Règles de détection : la PREMIÈRE qui correspond gagne, donc l'ordre compte.
REGLES_TYPE = [
    (r"\bcross\b",             "Cross"),
    (r"\bds\b",                "Entraînement piste"),   # Ds = Dour Sports
    (r"\bc[ôo]tes?\b",         "Séance côtes"),
    (r"\bfartlek\b",           "Fartlek"),
    (r"endurance\s+intensive", "Endurance intensive"),
    (r"endurance\s+lente",     "Endurance lente"),
    (r"r[ée]g[ée]n[ée]ration", "Régénération"),
    (r"\bfooting\b",           "Footing"),
    (r"\brepos\b",             "Repos"),
]


def titre_seance(description: str) -> str:
    texte = description.lower()
    for motif, titre in REGLES_TYPE:
        if re.search(motif, texte):
            return titre
    premiere = description.splitlines()[0] if description else "Séance"
    return premiere if len(premiere) <= 50 else premiere[:47] + "..."


def detecter_heure(description: str):
    m = re.search(r"\b(\d{1,2})\s*h\s*(\d{2})?\b", description)
    if m:
        return int(m.group(1)), int(m.group(2) or 0)
    return None


# -------------------------------------------------------------------- ICS ---

def _echapper(texte: str) -> str:
    return (texte.replace("\\", "\\\\").replace(";", "\\;")
                 .replace(",", "\\,").replace("\n", "\\n"))


def _plier(ligne: str) -> str:
    """Les lignes .ics ne doivent pas dépasser 75 octets."""
    if len(ligne.encode("utf-8")) <= 75:
        return ligne
    morceaux, courant = [], ""
    for car in ligne:
        if len((courant + car).encode("utf-8")) > 73:
            morceaux.append(courant)
            courant = car
        else:
            courant += car
    morceaux.append(courant)
    return "\r\n ".join(morceaux)


def _est_repos(description: str) -> bool:
    return description.strip().lower() == "repos"


def evenements(plan: list[dict], inclure_repos: bool, heure_defaut=(18, 0),
               duree_min: int = 60) -> list[dict]:
    """Calcule la liste des événements (date, titre, description)."""
    res = []
    for s in plan:
        if _est_repos(s["description"]) and not inclure_repos:
            continue
        jour = (datetime.strptime(s["semaine_debut"], "%d/%m/%Y")
                + timedelta(days=s["numero"] - 1))  # Séance 1 = samedi
        h, m = detecter_heure(s["description"]) or heure_defaut
        debut = jour.replace(hour=h, minute=m)
        res.append({
            "debut": debut,
            "fin": debut + timedelta(minutes=duree_min),
            "titre": titre_seance(s["description"]),
            "description": s["description"],
        })
    return res


def generer_ics(evts: list[dict]) -> bytes:
    lignes = [
        "BEGIN:VCALENDAR", "VERSION:2.0",
        "PRODID:-//Plan entrainement//FR", "CALSCALE:GREGORIAN",
        "X-WR-CALNAME:Plan d'entraînement",
    ]
    maintenant = datetime.now().strftime("%Y%m%dT%H%M%S")
    for e in evts:
        lignes += [
            "BEGIN:VEVENT",
            f"UID:{uuid.uuid4()}@plan-entrainement",
            f"DTSTAMP:{maintenant}",
            f"DTSTART:{e['debut'].strftime('%Y%m%dT%H%M%S')}",
            f"DTEND:{e['fin'].strftime('%Y%m%dT%H%M%S')}",
            f"SUMMARY:{_echapper(e['titre'])}",
            f"DESCRIPTION:{_echapper(e['description'])}",
            "END:VEVENT",
        ]
    lignes.append("END:VCALENDAR")
    return ("\r\n".join(_plier(l) for l in lignes) + "\r\n").encode("utf-8")
