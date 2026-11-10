"""
Crée des fichiers Excel FICTIFS pour tester l'agent (inspirés d'un ERP de chaudronnerie inox).
  data/reference/   : les fichiers déjà complets, qui servent de base de connaissances
  data/a_completer/ : le fichier qui contient des cases vides
Lancement : python creer_donnees_exemple.py
"""
from pathlib import Path

import pandas as pd

DOSSIER = Path(__file__).parent / "data"

# 1) Catalogue des articles existants (complet)
catalogue = pd.DataFrame([
    # id_article, designation,                 famille,      matiere,    diametre_mm, atelier,         prix_unitaire
    ("ART-001", "Tube inox 10 mm",               "Tube",        "Inox 316L", 10, "Tuyauterie",     42.50),
    ("ART-002", "Tube inox 16 mm",               "Tube",        "Inox 316L", 16, "Tuyauterie",     58.90),
    ("ART-003", "Tube inox 20 mm",               "Tube",        "Inox 316L", 20, "Tuyauterie",     71.20),
    ("ART-004", "Tube titane 16 mm",             "Tube",        "Titane",    16, "Tuyauterie",    184.00),
    ("ART-005", "Coude inox 90° 10 mm",          "Coude",       "Inox 316L", 10, "Tuyauterie",     26.40),
    ("ART-006", "Coude inox 90° 16 mm",          "Coude",       "Inox 316L", 16, "Tuyauterie",     31.80),
    ("ART-007", "Coude inox 45° 20 mm",          "Coude",       "Inox 316L", 20, "Tuyauterie",     29.60),
    ("ART-008", "Bride inox 10 mm",              "Bride",       "Inox 304L", 10, "Chaudronnerie", 213.50),
    ("ART-009", "Bride inox 16 mm",              "Bride",       "Inox 304L", 16, "Chaudronnerie", 236.00),
    ("ART-010", "Bride inox 20 mm",              "Bride",       "Inox 304L", 20, "Chaudronnerie", 251.70),
    ("ART-011", "Collecteur inox 16 mm",         "Collecteur",  "Inox 316L", 16, "Soudure",       388.00),
    ("ART-012", "Collecteur inox 20 mm",         "Collecteur",  "Inox 316L", 20, "Soudure",       412.30),
    ("ART-013", "Collecteur titane 20 mm",       "Collecteur",  "Titane",    20, "Soudure",       905.00),
    ("ART-014", "Tôle formée inox 1,5 mm",       "Tôle formée", "Inox 304L", None, "Tôlerie",     118.40),
    ("ART-015", "Tôle formée inox 2 mm",         "Tôle formée", "Inox 304L", None, "Tôlerie",     134.90),
    ("ART-016", "Tôle formée titane 1,5 mm",     "Tôle formée", "Titane",    None, "Tôlerie",     402.00),
], columns=["id_article", "designation", "famille", "matiere", "diametre_mm", "atelier", "prix_unitaire"])

# 2) Une deuxième source : les règles d'affectation des ateliers (autre format, autre fichier)
regles_ateliers = pd.DataFrame([
    ("Tube", "Tuyauterie", "Découpe et cintrage des tubes"),
    ("Coude", "Tuyauterie", "Cintrage des coudes"),
    ("Bride", "Chaudronnerie", "Usinage et perçage des brides"),
    ("Collecteur", "Soudure", "Assemblage soudé de plusieurs tubes"),
    ("Tôle formée", "Tôlerie", "Découpe laser et pliage des tôles"),
], columns=["famille", "atelier", "operation_principale"])

# 3) Le fichier à compléter : des cases vides, avec des cas tests volontaires
a_completer = pd.DataFrame([
    # Cas 1 : article déjà connu (même id) -> complétion directe, sans IA
    ("ART-006", "Coude inox 90° 16 mm",    None,    "Inox 316L", 16,   None,          None),
    # Cas 2 : nouvel article proche d'articles existants -> l'IA doit trouver famille, matière, atelier
    ("ART-101", "Tube inox 25 mm",         None,    None,        None, None,          None),
    # Cas 3 : nouvel article en titane -> la matière doit être Titane (et pas Inox)
    ("ART-102", "Collecteur titane 16 mm", "Collecteur", None,   16,   None,          None),
    # Cas 4 : article sans équivalent dans la base -> rien ne doit être inventé
    ("ART-103", "Capteur de pression",     None,    None,        None, None,          None),
], columns=["id_article", "designation", "famille", "matiere", "diametre_mm", "atelier", "prix_unitaire"])

if __name__ == "__main__":
    (DOSSIER / "reference").mkdir(parents=True, exist_ok=True)
    (DOSSIER / "a_completer").mkdir(parents=True, exist_ok=True)
    catalogue.to_excel(DOSSIER / "reference" / "catalogue_articles.xlsx", index=False)
    regles_ateliers.to_excel(DOSSIER / "reference" / "regles_ateliers.xlsx", index=False)
    a_completer.to_excel(DOSSIER / "a_completer" / "articles_a_completer.xlsx", index=False)
    print("Fichiers d'exemple créés dans", DOSSIER)
