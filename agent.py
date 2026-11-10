"""
Agent IA de complétion de fichiers Excel (prototype pédagogique).

Même principe que l'agent d'aide aux appels d'offres : l'agent ne remplit une case vide
que s'il trouve une SOURCE dans les fichiers Excel existants. Sinon, il laisse la case
vide et la signale "À vérifier". Il n'invente rien.

Il enchaîne 5 étapes, chacune confiée à un "rôle" :
  1. DÉTECTEUR    : repère les cases vides du fichier à compléter (règles Python, pas besoin d'IA)
  2. CHERCHEUR    : pour chaque case vide, cherche les lignes les plus proches dans les fichiers de
                    référence (RAG sur des lignes Excel). Si le même identifiant existe, on recopie directement.
  3. RÉDACTEUR    : propose une valeur en s'appuyant UNIQUEMENT sur ces lignes, avec ses sources
  4. VÉRIFICATEUR : un second appel au LLM relit la proposition : est-elle justifiée par les sources ?
  5. GARDE-FOUS   : règles Python (valeur autorisée, sources réelles, confiance, chiffres justifiés)
                    -> sinon "À vérifier par un humain"
Résultat : une COPIE du fichier complétée et colorée + un journal de chaque décision.
Le fichier d'origine n'est jamais modifié.

Lancement :
    python creer_donnees_exemple.py   -> crée des fichiers Excel fictifs pour tester
    python agent.py --demo            -> sans LLM ni internet (vote simple), pour tester la chaîne
    python agent.py                   -> avec le LLM (clé GROQ_API_KEY dans un fichier .env)
"""
from __future__ import annotations

import argparse
import json
import os
import re
from collections import defaultdict
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

DOSSIER = Path(__file__).parent
FICHIER_A_COMPLETER = DOSSIER / "data" / "a_completer" / "articles_a_completer.xlsx"
DOSSIER_REFERENCE = DOSSIER / "data" / "reference"
SORTIE = DOSSIER / "sorties"

CLE = "id_article"          # la clé primaire : la colonne qui identifie chaque ligne
# Colonnes que l'IA ne remplit JAMAIS seule (risque métier) : seule une correspondance exacte sur la clé
# peut les compléter. Sinon, l'IA fait une suggestion et un humain décide.
COLONNES_SENSIBLES = {"prix_unitaire"}
MODELE = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")

SEUIL_PERTINENCE = 0.10     # score minimal pour qu'une ligne de référence soit jugée utile
SEUIL_CONFIANCE = 0.60      # même seuil que dans l'agent appels d'offres
NB_LIGNES = 5               # nombre de lignes de référence données au rédacteur

SOURCE_LIGNE = "ligne à compléter"   # le rédacteur peut aussi s'appuyer sur la ligne elle-même

MOTS_VIDES = """le la les un une des du de d l à au aux et ou en dans sur pour par avec sans""".split()

VERT = PatternFill("solid", fgColor="C6EFCE")    # complété automatiquement
ORANGE = PatternFill("solid", fgColor="FFEB9C")  # suggestion à vérifier
ROUGE = PatternFill("solid", fgColor="FFC7CE")   # aucune information trouvée


def est_vide(valeur) -> bool:
    return valeur is None or (isinstance(valeur, float) and pd.isna(valeur)) or str(valeur).strip() == ""


def en_texte(valeur) -> str:
    """Écrit une valeur proprement (16.0 -> 16) pour la recherche et pour le LLM."""
    if isinstance(valeur, float) and valeur.is_integer():
        return str(int(valeur))
    return str(valeur)


# ---------------------------------------------------------------------------
# Accès au LLM : une seule fonction, qui renvoie toujours un dictionnaire
# ---------------------------------------------------------------------------
class LLM:
    def __init__(self, demo: bool):
        self.demo = demo
        self.client = None
        if not demo:
            from dotenv import load_dotenv
            from groq import Groq
            load_dotenv()
            self.client = Groq(api_key=os.environ["GROQ_API_KEY"])

    def json(self, consigne: str, contenu: str) -> dict:
        """Envoie une consigne + un contenu au LLM et récupère un JSON (ou {} en cas d'échec)."""
        try:
            reponse = self.client.chat.completions.create(
                model=MODELE,
                messages=[{"role": "system", "content": consigne},
                          {"role": "user", "content": contenu}],
                temperature=0.1,
                response_format={"type": "json_object"},
            )
            return json.loads(reponse.choices[0].message.content)
        except Exception as erreur:
            print(f"   [LLM] échec : {type(erreur).__name__} -> réponse de secours")
            return {}


# ---------------------------------------------------------------------------
# Étape 1 : DÉTECTEUR (pas besoin d'IA pour trouver une case vide)
# ---------------------------------------------------------------------------
def detecter_cases_vides(tableau: pd.DataFrame) -> list[dict]:
    cases = []
    for index, ligne in tableau.iterrows():
        for colonne in tableau.columns:
            if colonne != CLE and est_vide(ligne[colonne]):
                cases.append({"index": index, "ligne_excel": index + 2, "colonne": colonne,
                              "cle": ligne.get(CLE)})
    return cases


# ---------------------------------------------------------------------------
# Étape 2 : CHERCHEUR (un RAG sur des lignes Excel, avec TF-IDF comme dans l'agent appels d'offres)
# ---------------------------------------------------------------------------
class BaseDeReference:
    def __init__(self, dossier: Path):
        # Chaque ligne de chaque fichier devient un "passage" texte, avec sa source précise
        self.lignes, self.textes, self.textes_recherche, self.sources = [], [], [], []
        self.valeurs_autorisees = defaultdict(set)   # valeurs déjà vues dans chaque colonne
        self.colonnes_numeriques = set()
        for fichier in sorted(dossier.glob("*.xlsx")):
            tableau = pd.read_excel(fichier)
            for colonne in tableau.columns:
                if pd.api.types.is_numeric_dtype(tableau[colonne]):
                    self.colonnes_numeriques.add(colonne)
                else:
                    self.valeurs_autorisees[colonne].update(tableau[colonne].dropna().astype(str))
            for index, ligne in tableau.iterrows():
                donnees = {c: v for c, v in ligne.items() if not est_vide(v)}
                self.lignes.append(donnees)
                self.textes.append(self.decrire(donnees))
                self.textes_recherche.append(self.texte_recherche(donnees))
                self.sources.append(f"{fichier.name}, ligne {index + 2}")
        self.vectoriseur = TfidfVectorizer(stop_words=MOTS_VIDES)
        self.matrice = self.vectoriseur.fit_transform(self.textes_recherche)

    @staticmethod
    def decrire(donnees: dict) -> str:
        """Transforme une ligne en texte lisible pour le LLM (« colonne: valeur | ... »)."""
        return " | ".join(f"{c}: {en_texte(v)}" for c, v in donnees.items() if c != CLE)

    @staticmethod
    def texte_recherche(donnees: dict) -> str:
        """Texte utilisé pour la recherche : seulement les VALEURS.
        Les noms de colonnes sont identiques partout : ils feraient croire que toutes les lignes se ressemblent.
        On ignore aussi la clé : un identifiant comme ART-101 n'a pas de sens."""
        return " ".join(en_texte(v) for c, v in donnees.items() if c != CLE)

    def correspondance_exacte(self, cle, colonne: str) -> dict | None:
        """Si le même identifiant existe déjà avec la colonne remplie, on recopie (règle, pas d'IA)."""
        for donnees, source in zip(self.lignes, self.sources):
            if donnees.get(CLE) == cle and colonne in donnees:
                return {"valeur": donnees[colonne], "source": source}
        return None

    def chercher(self, ligne_connue: dict, colonne: str, k: int = NB_LIGNES) -> list[dict]:
        """Cherche les lignes les plus proches qui contiennent une valeur pour la colonne voulue."""
        question = self.texte_recherche(ligne_connue)
        scores = cosine_similarity(self.vectoriseur.transform([question]), self.matrice)[0]
        resultats = []
        for i in scores.argsort()[::-1]:
            if scores[i] < SEUIL_PERTINENCE or len(resultats) == k:
                break
            if colonne in self.lignes[i]:
                resultats.append({"source": self.sources[i], "ligne": self.lignes[i],
                                  "texte": self.textes[i], "score": round(float(scores[i]), 3)})
        return resultats


# ---------------------------------------------------------------------------
# Étape 3 : RÉDACTEUR
# ---------------------------------------------------------------------------
def proposer(llm: LLM, colonne: str, ligne_connue: dict, passages: list[dict],
             valeurs_autorisees: list[str], numerique: bool) -> dict:
    if not passages:
        return {"valeur": None, "sources": [], "justification": "Aucune ligne proche trouvée.", "confiance": 0.0}

    if llm.demo:  # sans IA : 2 règles simples
        # Règle 1 : si une valeur autorisée est écrite dans la ligne elle-même (ex. « titane »), on la prend
        texte_ligne = BaseDeReference.texte_recherche(ligne_connue).lower()
        for valeur in valeurs_autorisees:
            if valeur.lower() in texte_ligne:
                return {"valeur": valeur, "sources": [SOURCE_LIGNE],
                        "justification": "Valeur écrite dans la ligne à compléter (mode démo).", "confiance": 0.9}
        # Règle 2 : vote des lignes proches, pondéré par leur score de similarité
        votes, appuis = defaultdict(float), defaultdict(list)
        for p in passages:
            valeur = en_texte(p["ligne"][colonne])
            votes[valeur] += p["score"]
            appuis[valeur].append(p["source"])
        gagnante = max(votes, key=votes.get)
        return {"valeur": gagnante, "sources": appuis[gagnante],
                "justification": "Valeur la plus fréquente parmi les lignes proches (mode démo).",
                "confiance": round(votes[gagnante] / sum(votes.values()), 2)}

    consigne = (
        f"Tu complètes UNE case vide (colonne « {colonne} ») d'un fichier Excel industriel. "
        "Appuie-toi UNIQUEMENT sur la ligne à compléter et sur les lignes de référence fournies. "
        "Si l'information ne s'y trouve pas, réponds avec une valeur null. N'estime jamais un prix ou une quantité. "
        + (f"La valeur doit être l'une de celles-ci : {valeurs_autorisees}. " if valeurs_autorisees else "")
        + 'Réponds en JSON : {"valeur": "... ou null", '
        f'"sources": ["noms exacts des sources utilisées, ou \'{SOURCE_LIGNE}\'"], '
        '"justification": "une phrase", "confiance": nombre entre 0 et 1}'
    )
    contenu = (f"LIGNE À COMPLÉTER ({SOURCE_LIGNE}) : {BaseDeReference.decrire(ligne_connue)}\n"
               f"COLONNE À REMPLIR : {colonne}\n\nLIGNES DE RÉFÉRENCE :\n"
               + "\n".join(f"[{p['source']}] {p['texte']}" for p in passages))
    return llm.json(consigne, contenu)


# ---------------------------------------------------------------------------
# Étape 4 : VÉRIFICATEUR (un second "agent" relit le premier)
# ---------------------------------------------------------------------------
def verifier(llm: LLM, colonne: str, ligne_connue: dict, proposition: dict, passages: list[dict]) -> dict:
    if llm.demo or not passages:
        return {"justifiee": bool(passages), "commentaire": "vérification automatique non effectuée (mode démo)"}
    consigne = (
        "Tu es un relecteur exigeant. Vérifie si la valeur proposée pour la case vide est directement "
        "justifiée par la ligne à compléter et les lignes de référence (et pas seulement plausible). "
        'Réponds en JSON : {"justifiee": true ou false, "commentaire": "une phrase"}'
    )
    contenu = (f"LIGNE À COMPLÉTER : {BaseDeReference.decrire(ligne_connue)}\nCOLONNE : {colonne}\n"
               f"PROPOSITION : {json.dumps(proposition, ensure_ascii=False)}\nLIGNES DE RÉFÉRENCE :\n"
               + "\n".join(f"[{p['source']}] {p['texte']}" for p in passages))
    return llm.json(consigne, contenu)


# ---------------------------------------------------------------------------
# Étape 5 : GARDE-FOUS (règles Python : on ne fait jamais confiance aveuglément au LLM)
# ---------------------------------------------------------------------------
def garde_fous(colonne: str, proposition: dict, verification: dict, passages: list[dict], ligne_connue: dict,
               valeurs_autorisees: list[str], numerique: bool) -> dict:
    valeur = proposition.get("valeur")
    if not passages or est_vide(valeur) or str(valeur).lower() == "null":
        return {"statut": "Non trouvé", "valeur": None, "sources": "", "confiance": 0.0,
                "alertes": "aucune information dans les fichiers de référence", "commentaire": ""}

    alertes = []
    sources_autorisees = {p["source"] for p in passages} | {SOURCE_LIGNE}
    sources = [s for s in proposition.get("sources", []) if s in sources_autorisees]
    if not sources:
        alertes.append("aucune source valide citée")

    valeur = str(valeur).strip()
    if numerique:
        try:
            nombre = float(valeur.replace(",", "."))
        except ValueError:
            alertes.append("la colonne attend un nombre")
            nombre = None
        # Un chiffre n'est accepté que s'il est écrit dans la ligne elle-même (ex. « 25 mm » dans la désignation).
        # Sinon c'est une estimation : un humain doit décider (prix, quantités...).
        if nombre is not None:
            valeur = en_texte(nombre)
            if not re.search(rf"(?<![\d,.]){re.escape(valeur)}(?![\d,.])", BaseDeReference.decrire(ligne_connue)):
                alertes.append("chiffre non écrit dans la ligne : estimation à valider")
    elif valeurs_autorisees and valeur not in valeurs_autorisees:
        alertes.append(f"valeur inconnue des fichiers de référence ({valeur})")

    if colonne in COLONNES_SENSIBLES:
        alertes.append("colonne sensible : validation humaine obligatoire")

    try:
        confiance = float(proposition.get("confiance", 0))
    except (TypeError, ValueError):
        confiance = 0.0
    if confiance < SEUIL_CONFIANCE:
        alertes.append(f"confiance faible ({confiance:.2f})")
    if not verification.get("justifiee", False):
        alertes.append("proposition non justifiée selon le vérificateur")

    return {"statut": "À vérifier" if alertes else "Complété", "valeur": valeur,
            "sources": ", ".join(sources), "confiance": round(confiance, 2),
            "alertes": "; ".join(alertes) or "aucune",
            "commentaire": verification.get("commentaire", "") or proposition.get("justification", "")}


# ---------------------------------------------------------------------------
# Écriture des résultats : une copie colorée + un onglet Journal (le fichier d'origine n'est pas modifié)
# ---------------------------------------------------------------------------
def ecrire_resultats(journal: list[dict], tableau: pd.DataFrame) -> Path:
    SORTIE.mkdir(exist_ok=True)
    chemin = SORTIE / FICHIER_A_COMPLETER.name.replace(".xlsx", "_complete.xlsx")
    tableau.to_excel(chemin, index=False)

    classeur = load_workbook(chemin)
    feuille = classeur.active
    numero_colonne = {nom: i + 1 for i, nom in enumerate(tableau.columns)}
    for j in journal:
        cellule = feuille.cell(row=j["Ligne Excel"], column=numero_colonne[j["Colonne"]])
        if j["Statut"] == "Complété":
            cellule.fill = VERT
        elif j["Statut"] == "À vérifier":
            cellule.fill = ORANGE
            cellule.comment = Comment(f"Suggestion : {j['Valeur proposée']}\n{j['Alertes']}", "Agent IA")
        else:
            cellule.fill = ROUGE

    onglet = classeur.create_sheet("Journal")
    entetes = list(journal[0].keys()) if journal else []
    onglet.append(entetes)
    for j in journal:
        onglet.append([j[e] for e in entetes])
    classeur.save(chemin)
    return chemin


# ---------------------------------------------------------------------------
# L'ORCHESTRATEUR : il enchaîne les étapes
# ---------------------------------------------------------------------------
def lancer(demo: bool) -> None:
    llm = LLM(demo)
    base = BaseDeReference(DOSSIER_REFERENCE)
    tableau = pd.read_excel(FICHIER_A_COMPLETER).astype(object)
    print(f"Base de référence : {len(base.textes)} lignes lues dans {DOSSIER_REFERENCE.name}/")

    print("Étape 1 - Détection des cases vides...")
    cases = detecter_cases_vides(tableau)
    print(f"   {len(cases)} cases vides trouvées")

    journal = []
    for case in cases:
        index, colonne = case["index"], case["colonne"]
        # La ligne telle qu'elle est MAINTENANT : les cases déjà complétées aident les suivantes
        ligne_connue = {c: v for c, v in tableau.loc[index].items() if not est_vide(v)}
        numerique = colonne in base.colonnes_numeriques
        autorisees = sorted(base.valeurs_autorisees.get(colonne, []))
        print(f"\n{case['cle']} / {colonne}")

        exacte = base.correspondance_exacte(case["cle"], colonne)          # étape 2 (raccourci)
        if exacte:
            final = {"statut": "Complété", "valeur": en_texte(exacte["valeur"]), "sources": exacte["source"],
                     "confiance": 1.0, "alertes": "aucune",
                     "commentaire": f"même {CLE} trouvé dans les fichiers de référence"}
        else:
            passages = base.chercher(ligne_connue, colonne)                  # étape 2
            print(f"   Étape 2 - {len(passages)} ligne(s) proche(s) : {[p['source'] for p in passages]}")
            proposition = proposer(llm, colonne, ligne_connue, passages, autorisees, numerique)  # étape 3
            verification = verifier(llm, colonne, ligne_connue, proposition, passages)          # étape 4
            final = garde_fous(colonne, proposition, verification, passages, ligne_connue, autorisees, numerique)  # 5

        if final["statut"] == "Complété":   # on n'écrit dans le tableau que les valeurs sûres
            tableau.at[index, colonne] = float(final["valeur"]) if numerique else final["valeur"]
        print(f"   Résultat : {final['statut']} -> {final['valeur']}  (alertes : {final['alertes']})")

        journal.append({"Ligne Excel": case["ligne_excel"], CLE: case["cle"], "Colonne": colonne,
                        "Statut": final["statut"], "Valeur proposée": final["valeur"],
                        "Sources": final["sources"], "Confiance": final["confiance"],
                        "Alertes": final["alertes"], "Commentaire": final["commentaire"]})

    chemin = ecrire_resultats(journal, tableau)
    resume = pd.DataFrame(journal)["Statut"].value_counts().to_dict() if journal else {}
    print(f"\nTerminé. Synthèse : {resume}")
    print(f"Fichier créé : {chemin}  (onglet « Journal » = détail de chaque décision)")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--demo", action="store_true", help="tester sans LLM ni internet")
    lancer(parser.parse_args().demo)
