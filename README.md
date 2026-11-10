# Agent IA de complétion de fichiers Excel

Prototype d'**agent IA** (workflow agentique + RAG) qui complète les **cases vides** d'un fichier Excel
en s'appuyant sur des **fichiers Excel déjà existants** (catalogue, référentiels, historiques).
Même principe que mon agent d'aide aux appels d'offres : **pas de source, pas de réponse**.
Une case n'est remplie que si une source le justifie. Sinon, elle est signalée « À vérifier » ou « Non trouvé ».

> Toutes les données sont **fictives** : un petit catalogue d'articles de chaudronnerie inox (tubes, coudes,
> brides, collecteurs, tôles), proche de ce qu'on trouve dans un ERP de production comme Clipper.

## Comment ça marche

```
fichier à compléter ──► 1. Détecteur ──► liste des cases vides
                                            │ (pour chaque case vide)
fichiers de référence ─► 2. Chercheur ────► même identifiant ? -> on recopie (règle, sans IA)
                                            sinon : lignes les plus proches + sources (RAG)
                         3. Rédacteur ────► valeur proposée + sources + confiance
                         4. Vérificateur ─► « la valeur est-elle justifiée par les sources ? »
                         5. Garde-fous ───► sinon statut « À vérifier »
                                            │
             copie du fichier, colorée + onglet « Journal » (le fichier d'origine n'est jamais modifié)
```

| Étape | Rôle | Ce qu'il fait |
|---|---|---|
| 1 | Détecteur | Repère les cases vides (règle Python : pas besoin d'IA pour ça) |
| 2 | Chercheur (RAG) | Si le même `id_article` existe dans les références, il recopie la valeur. Sinon, il cherche les lignes les plus proches dans tous les fichiers de référence (TF-IDF + similarité cosinus) |
| 3 | Rédacteur | Le LLM propose une valeur **uniquement à partir de ces lignes**, avec ses sources (`fichier, ligne N`) |
| 4 | Vérificateur | Un second appel au LLM relit la proposition : est-elle justifiée par les sources ? |
| 5 | Garde-fous | Règles Python : valeur déjà connue dans la colonne, sources réellement fournies, confiance ≥ 0,60, chiffres présents dans la ligne, colonnes sensibles (prix) jamais remplies seules. Sinon : **« À vérifier »** |

**Les couleurs dans le fichier produit :**

- **Vert** : complété automatiquement, avec une source ;
- **Orange** : suggestion à valider (la valeur proposée est dans le commentaire de la cellule) ;
- **Rouge** : aucune information dans les fichiers de référence.

L'onglet **Journal** détaille chaque décision : valeur, sources, confiance, alertes.

## Technologies

Python · Groq API (modèle Llama 3.3 70B) · scikit-learn (TF-IDF) · pandas · openpyxl

## Lancer le projet

```bash
pip install -r requirements.txt

# 0) Créer les fichiers Excel fictifs de test
python creer_donnees_exemple.py

# 1) Tester toute la chaîne sans IA ni internet
python agent.py --demo

# 2) Version complète avec le LLM
cp .env.example .env        # puis coller sa clé Groq dans .env (https://console.groq.com/keys)
python agent.py
```

Le résultat est écrit dans `sorties/articles_a_completer_complete.xlsx`.

Pour l'utiliser sur d'autres fichiers : mettre les fichiers complets dans `data/reference/`, le fichier à compléter
dans `data/a_completer/`, puis adapter en haut de `agent.py` : `FICHIER_A_COMPLETER`, `CLE` (la colonne identifiant)
et `COLONNES_SENSIBLES`.

## Cas tests prévus dans les données

| Ligne | Situation | Attendu |
|---|---|---|
| ART-006 | Article déjà présent dans le catalogue | Famille, atelier et prix recopiés directement (confiance 1) |
| ART-101 | « Tube inox 25 mm », nouvel article proche des tubes existants | Famille Tube, matière Inox 316L, atelier Tuyauterie. Diamètre 25 lu dans la désignation (avec le LLM). Prix : **suggestion à valider** |
| ART-102 | « Collecteur titane 16 mm » | Matière **Titane** (pas Inox), atelier Soudure. Prix : suggestion à valider |
| ART-103 | « Capteur de pression », aucun équivalent | **Non trouvé** partout. L'agent ne doit **rien inventer** |

## Limites et pistes d'amélioration

- La recherche utilise **TF-IDF** (mots en commun). Des **embeddings** trouveraient aussi les synonymes
  (« conduite » / « tube »).
- Le mode `--demo` vote entre les lignes proches : il ne sait pas lire un chiffre dans une désignation
  (le diamètre de ART-101 reste donc « À vérifier » sans LLM).
- Avec de gros fichiers, il faudrait limiter les appels au LLM (par exemple ne traiter que certaines colonnes)
  et stocker les vecteurs dans une base vectorielle.
- En entreprise, les fichiers de référence viendraient directement de l'ERP (exports Clipper) ou de SharePoint,
  et l'outil pourrait être intégré à Excel via Power Automate ou Copilot Studio.

## Structure

```
agent.py                       # tout l'agent (5 étapes + orchestrateur)
creer_donnees_exemple.py       # crée les fichiers Excel fictifs
data/reference/                # les fichiers complets (base de connaissances)
data/a_completer/              # le fichier avec des cases vides
sorties/                       # le fichier complété + journal (créé au lancement)
requirements.txt
.env.example                   # modèle du fichier qui contient la clé API (le vrai .env n'est jamais publié)
```

Auteur : Dwayne Enam, étudiant ingénieur à l'EFREI Paris (Big Data & IA).
