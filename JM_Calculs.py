import pandas as pd
import os
from typing import Iterable, Literal

from moteur import verif_els, verif_elu, verif_feu, design_section

# Clefs des dictionnaires.
IDSECTION_KEYS = {
    'File', 'ID'
}
MATERIAUX_KEYS = {
    'fck', 'class_acier', 'fyk', 'Ef', 'sigma_fs', 'sigma_fu', 'carbone_feu',
}
GEOMETRIE_KEYS = {
    'h_dalle', 'b_dalle', 'As', 'dprim_s', 'ns',
}
RENFORTS_KEYS = {
    'Asr', 'dprim_sr', 'nsr', 'Af', 'dprim_f', 'nf',
}
EFFORTS_KEYS = {
    'm_els_1', 'm_els_2', 'm_elu', 'm_feu',
}

# Clefs [ Nombre entier]
INT_KEYS = {'ns', 'nsr', 'nf'}

# =====================================================================================



# =====================================================================================
# -----------------------------------------------------------------------------
# Fonction : excel_to_listofrowdicts(path)
# Objectif :
#   Lire un fichier Excel et renvoyer son contenu sous forme de liste de
#   dictionnaires, où chaque dictionnaire représente une ligne du tableau.
#
# Détails :
#   1. La fonction utilise la bibliothèque pandas (pd.read_excel) pour charger
#      la feuille de calcul spécifiée.
#   2. Chaque ligne du fichier Excel est ensuite convertie en dictionnaire,
#      avec les noms de colonnes comme clés.
#   3. Le résultat final est une liste de ces dictionnaires, facile à parcourir
#      et à transmettre à d’autres fonctions comme _build_input_dict().
#
# Exemple :
#   Si le fichier Excel contient :
#       | E      | fck | largeur | hauteur |
#       |--------|------|---------|---------|
#       | 210000 | 30   | 300     | 500     |
#       | 210000 | 35   | 250     | 400     |
#
#   _read_excel_with_pandas("exemple.xlsx") ➜ [
#       {'E': 210000, 'fck': 30, 'largeur': 300, 'hauteur': 500},
#       {'E': 210000, 'fck': 35, 'largeur': 250, 'hauteur': 400}
#   ]
#
# Paramètres :
#   path : str
#       Chemin complet du fichier Excel à lire.
#   sheet_name : int | str (optionnel)
#       Nom ou index de la feuille à charger (0 = première feuille).
#
# Retour :
#   list[dict]
#       Liste de dictionnaires représentant les lignes du fichier Excel.
# -----------------------------------------------------------------------------
def excel_to_listofrowdicts(path: str, sheet_name=0) -> list[dict] :
    """Lit un fichier Excel et renvoie une liste de lignes sous forme de dictionnaires."""
    df = pd.read_excel(path, sheet_name=sheet_name)
    df = df.where(pd.notnull(df), None)
    # Conversion du DataFrame en liste de dictionnaires
    # Chaque dictionnaire correspond à une ligne, avec {colonne: valeur}
    return [dict(r) for r in df.to_dict(orient='records')]

# -----------------------------------------------------------------------------
# Fonction : _to_number(v)
# Objectif :
#   Convertir une valeur brute (issue d’un fichier Excel ou d’une saisie utilisateur)
#   en un nombre (float ou int) proprement utilisable dans les calculs.
#
# Détails :
#   - Si la valeur est vide (None, "NaN", "null", etc.), la fonction renvoie None.
#   - Si c’est déjà un nombre (int ou float), la valeur est renvoyée telle quelle.
#   - Si c’est une chaîne, la fonction :
#       * supprime les espaces et espaces insécables,
#       * remplace les virgules par des points (ex : "3,14" → "3.14"),
#       * essaie ensuite de convertir la chaîne en float.
#   - Si la conversion échoue (par exemple pour du texte comme "Acier B500"),
#     la valeur d’origine est simplement renvoyée sans provoquer d’erreur.
#
# Exemple :
#   _to_number(" 1 200,5 ")  ->  1200.5
#   _to_number("NaN")        ->  None
#   _to_number("Aciers B500")->  "Aciers B500"
# -----------------------------------------------------------------------------
def _to_number(v):
    if v is None: # Si la cellule Excel était vide, pandas renvoie souvent None (ou NaN) --> Donc ici, on renvoie directement None (valeur vide propre).
        return None
    if isinstance(v, (int, float)): # Si la valeur est déjà numérique, pas besoin de conversion → on la retourne telle quelle --> Cela évite des traitements inutiles.
        return v
    s = str(v).strip() # Ici, la fonction corrige plusieurs cas très fréquents dans les fichiers Excel :
    if s == "" or s.lower() in {"nan", "none", "null"}: 
        return None
    # Ici, la fonction corrige plusieurs cas très fréquents dans les fichiers Excel 
    # (Espaces normaux -" 1 200 "                                      --> devient "1200" /
    #  Espaces insécables (\u00a0) - "1 200" (copié depuis Word/Excel) --> devient "1200" /
    # Virgule comme séparateur décimal - "3,14"                        -->devient "3.14")
    s = s.replace(" ", "").replace("\u00a0", "").replace(",", ".")
    try:   # On essaye de convertir en float. Si ça échoue (par exemple "Aciers B500" ou "N/A"), on retourne la valeur d’origine inchangée. Cela permet de ne pas bloquer le programme sur une cellule non numérique (comme un commentaire)
        return float(s)
    except Exception:
        return v


# -----------------------------------------------------------------------------
# Fonction : _build_input_dict(row)
# Objectif :
#   Transformer une ligne de données (par exemple issue d’un fichier Excel)
#   en un dictionnaire structuré contenant quatre sous-parties :
#       - 'materiaux'
#       - 'geometrie'
#       - 'renforts'
#       - 'efforts'
#
# Détails :
#   - La fonction parcourt chaque colonne de la ligne (clé/valeur du dictionnaire `row`).
#   - Chaque clé (nom de colonne) est comparée à des ensembles de référence
#     (MATERIAUX_KEYS, GEOMETRIE_KEYS, RENFORTS_KEYS, EFFORTS_KEYS)
#     pour savoir à quelle catégorie elle appartient.
#   - La valeur correspondante est convertie en nombre si possible
#     grâce à la fonction utilitaire _to_number().
#   - Les colonnes inconnues sont ignorées (cela permet d’avoir des colonnes
#     supplémentaires dans Excel sans provoquer d’erreur).
#   - Certaines clés définies comme entières (dans INT_KEYS) sont corrigées :
#     Excel renvoie parfois 4.0 au lieu de 4 → on convertit proprement en int.
#
# Résultat :
#   Retourne un dictionnaire structuré prêt à être utilisé pour les calculs :
#       {
#           'materiaux': {...},
#           'geometrie': {...},
#           'renforts': {...},
#           'efforts': {...}
#       }
#
# Exemple :
#   row = {
#       "E": 210000, "fyk": 500,
#       "largeur": 300, "hauteur": 500,
#       "nb_barres": 4.0, "Nx": 100
#   }
#
#   _build_input_dict(row) ➜
#       {
#           'materiaux': {'E': 210000, 'fyk': 500},
#           'geometrie': {'largeur': 300, 'hauteur': 500},
#           'renforts': {'nb_barres': 4},
#           'efforts': {'Nx': 100}
#       }
# -----------------------------------------------------------------------------
def _build_input_dict(row: dict) -> dict[str, dict]:
    """Retourne {'materiaux','geometrie','renforts','efforts'} à partir d'un dict de ligne."""

    # Initialisation des 4 sous-dictionnaires
    materiaux: dict = {}
    geometrie: dict = {}
    renforts: dict = {}
    efforts: dict = {}

    # Répartition de chaque colonne dans la bonne catégorie
    for k, v in row.items():
        key = k.strip()  # Supprime les espaces autour du nom de colonne
        val = _to_number(v) # Convertit la valeur en nombre si possible
        if key in MATERIAUX_KEYS:
            materiaux[key] = val
        elif key in GEOMETRIE_KEYS:
            geometrie[key] = val
        elif key in RENFORTS_KEYS:
            renforts[key] = val
        elif key in EFFORTS_KEYS:
            efforts[key] = val
        else:
            # Ignore les colonnes inconnues
            pass

    # Correction des valeurs entières (ex : 4.0 → 4)
    for k in list(geometrie.keys()):
        if k in INT_KEYS and geometrie[k] is not None:
            try:
                geometrie[k] = int(float(geometrie[k]))
            except Exception:
                pass
    for k in list(renforts.keys()):
        if k in INT_KEYS and renforts[k] is not None:
            try:
                renforts[k] = int(float(renforts[k]))
            except Exception:
                pass

    # Retour du dictionnaire structuré complet
    return {
        'materiaux': materiaux,
        'geometrie': geometrie,
        'renforts': renforts,
        'efforts': efforts,
    }

# -----------------------------------------------------------------------------
# Fonction : input_dicos_entrée(path, sheet_name=0)
# Objectif :
#   Charger un fichier Excel et transformer chaque ligne en un dictionnaire
#   structuré contenant les 4 groupes de données nécessaires :
#       - 'materiaux'
#       - 'geometrie'
#       - 'renforts'
#       - 'efforts'
#
# Détails :
#   1. La fonction lit le fichier Excel à l’aide de _read_excel_with_pandas(),
#      qui renvoie une liste de lignes sous forme de dictionnaires.
#   2. Chaque dictionnaire (chaque ligne Excel) est ensuite passé à la fonction
#      _build_input_dict(), qui trie et nettoie les colonnes pour créer
#      la structure standard attendue par les fonctions de calcul.
#   3. Le résultat final est une liste de dictionnaires formatés,
#      prête à être utilisée dans run_batch() ou d’autres traitements.
#
# Exemple :
#   input_dicos_entrée("calculs.xlsx") ➜ [
#       {
#           'materiaux': {...},
#           'geometrie': {...},
#           'renforts': {...},
#           'efforts': {...}
#       },
#       {
#           'materiaux': {...},
#           'geometrie': {...},
#           'renforts': {...},
#           'efforts': {...}
#       },
#       ...
#   ]
#
# Paramètres :
#   path : str
#       Chemin complet du fichier Excel à charger.
#   sheet_name : int | str (optionnel)
#       Nom ou index de la feuille Excel à lire (0 = première feuille).
#
# Retour :
#   list[dict[str, dict]]
#       Liste de dictionnaires structurés pour chaque ligne du fichier.
# -----------------------------------------------------------------------------
def input_dicos_entrée(path: str, sheet_name=0) -> list[dict[str, dict]]:
    """Charge un fichier Excel et normalise chaque ligne en 4 sous-dictionnaires."""
    
    # Étape 1 : Lire le fichier Excel et récupérer une liste de lignes
    # Chaque ligne est un dictionnaire {colonne: valeur}
    rows = excel_to_listofrowdicts(path, sheet_name=sheet_name)
    
    # Étape 2 : Pour chaque ligne, construire un dictionnaire structuré
    # grâce à la fonction _build_input_dict()
    # (list comprehension = boucle condensée en une seule ligne)
    return [_build_input_dict(r) for r in rows]

# -----------------------------------------------------------------------------
# Fonction : row_results(d, combs)
# Objectif :
#   Calculer, pour une seule ligne de données (un cas de calcul),
#   toutes les valeurs de résultats numériques correspondant aux
#   combinaisons demandées :
#       - ELS  → État Limite de Service
#       - ELU  → État Limite Ultime
#       - FEU  → Résistance au feu
#
# Détails :
#   1. La fonction reçoit un dictionnaire `d` déjà normalisé contenant
#      quatre sous-parties principales :
#         - 'materiaux' : propriétés des matériaux
#         - 'geometrie' : dimensions et formes
#         - 'renforts'  : armatures ou renforts
#         - 'efforts'   : efforts appliqués
#
#   2. Pour chaque combinaison demandée (par exemple 'els' et 'elu'),
#      la fonction appelle `design_section(m, g, r, e, mode)` :
#         - `mode = 'sls'` pour ELS (service)
#         - `mode = 'uls'` pour ELU (ultime)
#         - `mode = 'fire'` pour Feu
#
#   3. Les résultats retournés par `design_section()` sont des dictionnaires
#      contenant les valeurs de calculs (moments, contraintes, etc.).
#      Ces résultats sont ajoutés à un dictionnaire `out` avec un préfixe
#      spécifique :
#         - "els_" pour ELS
#         - "elu_" pour ELU
#         - "feu_" pour Feu
#
# Exemple :
#   d = {
#       'materiaux': {...},
#       'geometrie': {...},
#       'renforts': {...},
#       'efforts': {...}
#   }
#
#   row_results(d, ('els', 'elu')) ➜
#       {
#           'els_m1': 120.0, 'els_sigma_c1': 15.2, ...
#           'elu_m_ed': 140.0, 'elu_sigma_c': 28.4, ...
#       }
#
# Paramètres :
#   d : dict[str, dict]
#       Ligne de données normalisée avec les 4 sous-dictionnaires.
#   combs : Iterable[str]
#       Liste ou tuple des combinaisons à évaluer ('els', 'elu', 'feu').
#
# Retour :
#   dict[str, float]
#       Dictionnaire des résultats numériques pour cette ligne.
# -----------------------------------------------------------------------------




def row_results (d: dict[str, dict], combs) -> dict[str, float]:
    """Calcule les colonnes de résultats pour une ligne à l'aide de design_section()."""

    # Extraction des sous-dictionnaires depuis la ligne d'entrée
    m = d['materiaux']
    g = d['geometrie']
    r = d['renforts']
    e = d['efforts']

    # Dictionnaire des résultats à remplir
    out: dict[str, float] = {}

    # -------------------------------------------------------------------------
    # Calculs en État Limite de Service (ELS)
    # -------------------------------------------------------------------------
    if 'els' in combs:
         # Appel à la fonction de conception pour le mode "service" (sls)
        sls = design_section(m, g, r, e, 'sls')
        # Ajout des résultats dans le dictionnaire de sortie avec préfixe "els_"
        out.update({
            'els_m1': sls.get('m1'),
            'els_m2': sls.get('m2'),
            'els_sigma_c1': sls.get('sigma_c1'),
            'els_sigma_c2': sls.get('sigma_c2'),
            'els_sigma_s1': sls.get('sigma_s1'),
            'els_sigma_s2': sls.get('sigma_s2'),
            'els_sigma_sr2': sls.get('sigma_sr2'),
            'els_sigma_f2': sls.get('sigma_f2'),
        })

    # -------------------------------------------------------------------------
    # Calculs en État Limite Ultime (ELU)
    # -------------------------------------------------------------------------
    if 'elu' in combs:
        # Appel à la fonction de conception pour le mode "ultime" (uls)
        uls = design_section(m, g, r, e, 'uls')
        # Ajout des résultats correspondants avec préfixe "elu_"
        out.update({
            'elu_m_ed': uls.get('m_ed'),
            'elu_m_rd1': uls.get('m_rd1'),
            'elu_m_rd2': uls.get('m_rd2'),
            'elu_sigma_c': uls.get('sigma_c'),
            'elu_sigma_s': uls.get('sigma_s'),
            'elu_sigma_sr': uls.get('sigma_sr'),
            'elu_sigma_f': uls.get('sigma_f'),
        })

    # -------------------------------------------------------------------------
    # Calculs en situation d'Incendie (FEU)
    # -------------------------------------------------------------------------
    if 'feu' in combs:
        # Appel à la fonction de conception pour le mode "feu" (fire)
        feu = design_section(m, g, r, e, 'fire')
        # Ajout des résultats correspondants avec préfixe "feu_"
        out.update({
            'feu_m_ed': feu.get('m_ed'),
            'feu_m_rd1': feu.get('m_rd1'),
            'feu_m_rd2': feu.get('m_rd2'),
            'feu_sigma_c': feu.get('sigma_c'),
            'feu_sigma_s': feu.get('sigma_s'),
            'feu_sigma_sr': feu.get('sigma_sr'),
            'feu_sigma_f': feu.get('sigma_f'),
        })
    # Retourne le dictionnaire complet des résultats
    return out

# -----------------------------------------------------------------------------
# Fonction : rows_results(calculs, combs=("els","elu"), sheet_name=0)
# Objectif :
#   Calculer, pour chaque ligne d'un fichier Excel (ou d'un ensemble de données),
#   les résultats numériques associés aux combinaisons demandées :
#       - 'els' → État Limite de Service
#       - 'elu' → État Limite Ultime
#       - 'feu' → Résistance au feu
#
# Détails :
#   1. La fonction reçoit en entrée :
#        - soit un chemin vers un fichier Excel (.xlsx / .xls / .xlsm),
#        - soit un itérable déjà normalisé (liste de dictionnaires
#          contenant les quatre sous-parties : 'materiaux', 'geometrie',
#          'renforts', 'efforts').
#
#   2. Elle appelle `input_dicos_entrée()` pour charger et normaliser
#      les données d'entrée. Cette fonction transforme le tableau Excel
#      en une liste de dictionnaires structurés, prêts pour le calcul.
#
#   3. Pour chaque ligne de données, `row_results()` est appelée pour
#      effectuer les calculs et retourner un dictionnaire de résultats
#      (moments, contraintes, etc.) avec des clés préfixées selon la
#      combinaison ('els_', 'elu_', 'feu_').
#
#   4. La fonction renvoie enfin une liste de dictionnaires contenant les
#      résultats de calcul pour toutes les lignes du fichier.
#
# Exemple :
#   rows_results("calculs.xlsx", combs=("els","elu")) ➜ [
#       {'els_m1': 120.0, 'els_sigma_c1': 15.5, 'elu_m_ed': 160.0, ...},
#       {'els_m1': 98.3, 'els_sigma_c1': 14.1, 'elu_m_ed': 140.5, ...},
#       ...
#   ]
#
# Paramètres :
#   calculs : str
#       Chemin du fichier Excel à lire, ou itérable de lignes normalisées.
#   combs : Iterable[Literal['els', 'elu', 'feu']]
#       Combinaisons à calculer (par défaut : 'els' et 'elu').
#   sheet_name : int | str
#       Nom ou index de la feuille Excel à lire (0 = première feuille).
#
# Retour :
#   list[dict[str, float]]
#       Liste de dictionnaires contenant les résultats numériques
#       pour chaque ligne d'entrée.
# -----------------------------------------------------------------------------
def rows_results(
    calculs: str,
    combs: Iterable[Literal['els', 'elu', 'feu']] = ("els", "elu"),
    sheet_name=0,
) -> list[dict[str, float]]:
    """Calcule les colonnes de résultats pour chaque ligne d'un fichier ou d'une liste normalisée."""
    
    # -------------------------------------------------------------------------
    # 1) Charger et normaliser les données d'entrée depuis le fichier Excel
    #    La fonction input_dicos_entrée() retourne une liste de dictionnaires :
    #       [{'materiaux': {...}, 'geometrie': {...}, 'renforts': {...}, 'efforts': {...}}, ...]
    # -------------------------------------------------------------------------
    rows = input_dicos_entrée(calculs, sheet_name=sheet_name)

    # -------------------------------------------------------------------------
    # 2) Calculer les résultats pour chaque ligne en appelant row_results()
    #    Cette fonction calcule les valeurs ELS/ELU/FEU pour une ligne donnée.
    # -------------------------------------------------------------------------
    return [row_results(d, combs) for d in rows]

# =====================================================================================

# =====================================================================================



# -----------------------------------------------------------------------------
# Fonction : excel_results(path, out_path=None, sheet_name=0, combs=("els","elu"))
# Objectif :
#   Générer un fichier Excel de sortie contenant :
#     - les colonnes d'entrée (issues du fichier Excel source),
#     - + les colonnes de résultats calculées (ELS / ELU / FEU selon 'combs').
#
# Détails :
#   1) Lecture du fichier Excel d'entrée dans un DataFrame pandas (df_in).
#   2) Normalisation des lignes (dicts 'materiaux'/'geometrie'/'renforts'/'efforts')
#      via input_dicos_entrée(), puis calcul des résultats par ligne via row_results().
#   3) Construction d'un DataFrame des résultats (df_res) et concaténation
#      horizontale avec df_in → df_out.
#   4) Écriture de df_out vers un fichier Excel (out_path) ; si out_path n'est
#      pas fourni, on crée "<path>_results.xlsx".
#
# Paramètres :
#   path : str
#       Chemin du fichier Excel d'entrée (.xlsx/.xls/.xlsm).
#   out_path : str | None
#       Chemin du fichier Excel de sortie. Si None, suffixe "_results.xlsx".
#   sheet_name : int | str
#       Index (0 = première feuille) ou nom de la feuille à lire/écrire.
#   combs : tuple[str, ...] | list[str]
#       Combinaisons à calculer, par ex. ("els","elu") ou ("els","elu","feu").
#
# Retour :
#   str
#       Chemin du fichier Excel écrit.
#
# Remarques :
#   - Si certaines lignes ne produisent pas toutes les mêmes clés de résultats,
#     pandas remplira les valeurs manquantes avec NaN (comportement normal).
#   - Assure-toi que input_dicos_entrée() lit la même feuille/ordre que df_in,
#     pour garder l'alignement ligne-à-ligne lors de la concaténation.
# ----------------------------------------------------------------------------
def excel_results(path: str, out_path: str | None = None, sheet_name=0, combs=("els", "elu")) -> str:
    """Écrit un Excel de sortie = colonnes d'entrée + colonnes de résultats."""
  
    # 1) Lire le tableau d'entrée (Excel → DataFrame)
    df_in = pd.read_excel(path, sheet_name=sheet_name)

    # 2) Normaliser les lignes (→ 4 sous-dicts) et calculer les résultats
    #    input_dicos_entrée(path) renvoie une liste de lignes normalisées :
    #    [{'materiaux': {...}, 'geometrie': {...}, 'renforts': {...}, 'efforts': {...}}, ...]
    rows = input_dicos_entrée(path, sheet_name=sheet_name)

    #    Pour chaque ligne normalisée, row_results(...) renvoie un dict de résultats
    #    (clés préfixées : 'els_*', 'elu_*', 'feu_*' selon 'combs').
    results_rows = [row_results(d, combs) for d in rows]

    # 3) Construire le DataFrame des résultats et concaténer avec l'entrée
    df_res = pd.DataFrame(results_rows) # colonnes = toutes les clés rencontrées
    #    reset_index(drop=True) pour garantir l'alignement par index (0..n-1)
    df_out = pd.concat([df_in.reset_index(drop=True), df_res], axis=1)

    # 4) Déterminer le chemin de sortie si absent : "<path>_results.xlsx"
    if out_path is None:
        base, _ = os.path.splitext(path)
        out_path = base + "_results.xlsx"

    # 5) Écrire le fichier Excel final sur disque (une seule feuille : 'calculs')
    df_out.to_excel(out_path, index=False, sheet_name='calculs')
    # 6) Retourner le chemin du fichier écrit
    return out_path



def run_in_terminal(
    calculs: str,
    combs: Iterable[Literal['els', 'elu', 'feu']] = ("els", "elu"),
    sheet_name=0,
) -> None:
    dic_list = input_dicos_entrée(calculs, sheet_name=sheet_name)


    for i, d in enumerate(dic_list, start=1):
        m = d['materiaux']
        g = d['geometrie']
        r = d['renforts']
        e = d['efforts']
        print(f"\n===== Calcul #{i} =====")
        if 'els' in combs:
            verif_els(m, g, r, e)
        if 'elu' in combs:
            verif_elu(m, g, r, e)
        if 'feu' in combs:
            verif_feu(m, g, r, e)


if __name__ == "__main__":
    # Input file (CSV or XLSX)
    PATH = r"D:\Python\Citallios\Calculs\DataBase_template.xlsx"  # <-- edit me


    print (PATH)

    print (excel_to_listofrowdicts(PATH ,0))

    ROWS =excel_to_listofrowdicts(PATH ,0)
    ROWS_1 =ROWS[0]

    print (_build_input_dict(ROWS_1))

    print ( input_dicos_entrée(PATH,0))
    
    print (rows_results(PATH,("els", "elu"),0))

    # run_in_terminal(PATH,("els", "elu"),0)

    excel_results(PATH,None,0)

    print("END")
