# Générateur export bulletin ITSI

Application web (100 % client, sans serveur) qui reformate un fichier
« Combine » (export de bulletins de paie, ex. Marie-Christine) au format
export bulletin standard : regroupement par département de tournage, mise
en page identique au modèle, et **formules de calcul conservées** (le
fichier généré reste un outil de travail modifiable — changer un nombre
d'heures recalcule automatiquement le montant en euros).

Tout se passe dans le navigateur : le fichier déposé n'est jamais envoyé à
un serveur.

## Utilisation

Ouvrir la page, déposer un fichier `.xlsx` (glisser-déposer ou bouton
« Choisir un fichier »), ajuster si besoin les champs Société / Production /
N° objet / IDCC, puis télécharger le fichier généré.

## Développement

```bash
npm install
npm run dev       # serveur de développement
npm run build     # build de production dans dist/
npm run preview   # sert le build de production en local
```

## Déploiement

Le déploiement sur GitHub Pages est automatique via
`.github/workflows/deploy.yml` à chaque push sur `main`. Si ce n'est pas
déjà fait, activer une fois GitHub Pages sur le repo :
**Settings → Pages → Source : GitHub Actions**.

## Logique de conversion

Le mapping des colonnes source → colonnes cible se fait par **nom de
colonne** (le libellé de la ligne d'en-tête du fichier déposé), jamais par
position : si l'ordre des colonnes change d'un export à l'autre, ou si des
colonnes sont ajoutées/absentes selon la production, la génération reste
correcte tant que l'intitulé de colonne ne change pas. La résolution du
mapping se fait dans `resolveColumnMapping()` (`src/generator.js`) :

- Toute colonne source dont le libellé (normalisé : minuscule, accents et
  unités "(h)"/"(€)" harmonisés) correspond à un libellé du format cible
  (`STANDARD_HEADERS`) est placée à la position standard correspondante,
  quelle que soit sa position dans le fichier source.
- Quelques champs "identité" portent un nom différent entre les deux
  formats (`IDENTITY_RENAME` : Code bulletin, Statut, Code contrat, Code
  employé, Matricule).
- Les colonnes de travail internes (Total base, MG, Ratio MG, Supp ap. MG,
  Total somme — qui servaient à répartir la Garantie Minimale entre deux
  bulletins d'un même salarié) sont détectées par motif (`DROPPED_PATTERNS`)
  et non reprises.
- Toute colonne source restante, sans équivalent dans le format cible (ex.
  Cachet grp./iso., Indem. Matériel, Indem. VHSS…), est ajoutée
  automatiquement en fin de tableau pour ne perdre aucune donnée — la liste
  de ces colonnes "extra" varie donc d'un fichier à l'autre.

Le classement des métiers par département (`METIER_TO_DEPT`) reste, lui,
une liste explicite à compléter si de nouveaux intitulés de poste
apparaissent ; un métier non reconnu est classé dans "AUTRES" et signalé à
l'utilisateur dans l'interface plutôt que de faire échouer la génération.
