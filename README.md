je crée mon environnement virtuel avec la commande `python -m venv venv`
j'ouvre le dossier du projet dans Visual Studio Code
j'ouvre le terminal intégré de VS Code
j'active l'environnement virtuel avec `.\venv\Scripts\Activate.ps1`
j'exécute `python --version` pour vérifier la version de Python
je vérifie que pip fonctionne avec `pip --version`
j'installe la bibliothèque nécessaire pour lire le fichier Excel avec `pip install openpyxl`
je vérifie que les fichiers de l'exercice se trouvent dans le dossier `donnees`
je vérifie que `donnees` contient le catalogue, les trois fichiers des enseignes et l'historique des prix
je crée le programme `analyser_marche.py` à la racine du projet
j'ajoute dans `analyser_marche.py` les bibliothèques nécessaires pour lire les fichiers CSV, JSON et Excel
j'ajoute les chemins vers les dossiers `donnees` et `resultats` pour retrouver les fichiers du projet
je crée le catalogue de référence avec les identifiants et les noms des produits
j'ajoute des alias pour reconnaître un même produit même si les enseignes n'écrivent pas son nom de la même façon
j'ajoute une fonction qui normalise les noms pour faciliter le rapprochement des produits
j'ajoute une fonction qui transforme les prix affichés en nombres que le programme peut comparer
j'ajoute le chargement du fichier CSV de Marché Lagune avec son séparateur point-virgule
j'ajoute le chargement du fichier JSON de Baobab Shop et le calcul du prix avec livraison
j'ajoute le chargement du classeur Excel du Comptoir Ivoire et l'application des remises
j'ajoute les contrôles de disponibilité pour exclure les produits en rupture ou sans stock
j'ajoute le nettoyage des doublons, des prix manquants et des prix aberrants
j'ajoute un journal qui décrit les anomalies trouvées et le traitement appliqué
j'ajoute le rapprochement des offres des trois enseignes avec les identifiants du catalogue
j'ajoute la comparaison du prix payé pour chaque produit disponible dans au moins deux enseignes
j'ajoute le calcul de l'enseigne la moins chère et des écarts en FCFA et en pourcentage
j'ajoute une synthèse qui compte les fois où chaque enseigne est la moins chère
j'ajoute le calcul de l'écart moyen par rapport au prix le plus bas et la liste des trois plus grands écarts
j'ajoute le calcul du panier de rentrée : P04, P12, deux P14, P15 et P28
j'ajoute le calcul du panier optimisé en choisissant le prix disponible le plus bas pour chaque article
j'ajoute le chargement de l'historique des prix et la comparaison avec la semaine du 24 au 30 novembre 2025
j'ajoute la mesure du temps de traitement de l'historique et l'explication de la complexité de l'index
j'ajoute un menu console avec huit options pour exécuter séparément les étapes de l'exercice
dans l'option 1, je charge les fichiers et j'affiche le nombre d'offres utilisables par enseigne
dans l'option 2, j'affiche les anomalies et j'exporte le journal de nettoyage
dans l'option 3, j'affiche le rapprochement des articles avec le catalogue
dans l'option 4, je compare les prix et j'exporte le fichier `comparatif.csv`
dans l'option 5, j'affiche la synthèse des prix les plus bas et des économies possibles
dans l'option 6, je calcule le prix du panier de rentrée dans chaque enseigne
dans l'option 7, je compare l'historique annuel à la semaine du Black Friday et j'exporte les résultats
dans l'option 8, je mesure le temps de traitement et j'affiche l'explication de la complexité
j'exécute le programme avec `python .\analyser_marche.py`
je choisis un numéro de 1 à 8 dans le menu pour lancer l'étape correspondante
j'appuie sur Entrée pour revenir au menu après l'affichage des résultats
je choisis 0 dans le menu pour quitter l'application
je retrouve les fichiers exportés dans le dossier `resultats`
j'ouvre `comparatif.csv` pour consulter les prix comparés
j'ouvre `journal_nettoyage.csv` pour consulter les anomalies et leur traitement
j'ouvre `historique_black_friday.csv` pour consulter les moyennes de prix historiques
les fichiers CSV sont enregistrés en UTF-8 avec BOM pour faciliter leur ouverture dans Excel
je garde en tête que les enseignes et les prix fournis sont fictifs et servent uniquement à cet exercice
