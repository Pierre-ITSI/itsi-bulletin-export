# Générateur export bulletin ITSI

Application web (100 % client, sans serveur) qui reformate un export excel
extrait d'itsi-production au format bulletin/feuille standard : regroupement
par département de tournage, mise en page identique au modèle, et
**formules de calcul conservées** (le fichier généré reste un outil de
travail modifiable — changer un nombre d'heures ou un taux horaire recalcule
automatiquement le montant en euros).

Tout se passe dans le navigateur : le fichier déposé n'est jamais envoyé à
un serveur.

## Utilisation

Ouvrir la page, choisir le **type de fichier** déposé (« Combine » pour un
export de bulletins de paie, ex. Marie-Christine, ou « Feuille » pour un
export de relevés d'heures — cf. `Deux types de fichiers source` plus bas),
déposer le fichier `.xlsx` (glisser-déposer ou bouton « Choisir un
fichier »), ajuster si besoin les champs Société / Production / N° objet /
IDCC, puis télécharger le fichier généré.

## Deux types de fichiers source

Le sélecteur « Type de fichier » choisit entre deux générateurs, qui
partagent l'essentiel de leur logique (`src/generator.js`) mais ciblent des
formats source différents :

- **Combine** (`generate()` / `buildOutput()`) : export de bulletins de
  paie consolidés (une ligne par bulletin), regroupement par contrat (code
  contrat). Section Garantie Minimale, "Jours travaillés (dates)"/"Jour(s)
  travaillés", "Salaire brut" (remplace "Total somme"), "Coût employeur",
  "Salaire net imposable/net".
- **Feuille** (`generateSheet()` / `buildSheetOutput()`) : export de
  relevés d'heures hebdomadaires (`SheetExcelExport` côté production, une
  ligne par relevé/semaine), regroupement par **matricule** (pas de code
  contrat dans cet export). Pas de section Garantie Minimale ni de "Jour(s)
  travaillés" ; à la place, "Total somme" (calculée, même rôle que "Salaire
  brut" côté Combine) et "Total itsi" (recopiée telle quelle depuis le
  fichier source, total déjà calculé côté production).

Les deux partagent : la résolution de colonnes par nom (`resolveColumnMapping()`,
paramétrable par format cible), le classement des métiers par département
(`METIER_TO_DEPT`/`classifyMetier()` — même référentiel des deux côtés,
"Métier" désignant le même champ), le tri par matricule
(`compareMatricules()`), les sections de couleur, les sous-totaux, et le
calcul "taux horaire x coefficient x heures" pour les variables concernées
(cf. `Colonnes calculées` plus bas — indexé par libellé court côté Combine,
par code brut côté Feuille, `SHEET_HOUR_RATE_COEF`).

Le générateur "Feuille" a été construit uniquement à partir des colonnes
définies dans `SheetExcelExport::headings()`/`format()` (fourni en
conversation), **sans fichier d'export réel pour validation** — à tester
sur un vrai fichier avant usage en production.

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
  quelle que soit sa position dans le fichier source. **Seules les colonnes
  standard réellement présentes dans le fichier déposé sont écrites** — le
  tableau de sortie ne contient jamais de colonne vide "au cas où".
- Quelques champs "identité" portent un nom différent entre les deux
  formats (`IDENTITY_RENAME` : Code bulletin, Statut, Code contrat, Code
  employé, Matricule).
- **Toute colonne source restante**, sans équivalent dans le format cible
  (ex. Cachet grp./iso., Indem. Matériel, Indem. VHSS, ou les anciennes
  colonnes de travail Total base/MG/Ratio MG/Supp ap. MG), est ajoutée
  automatiquement en fin de tableau pour ne perdre aucune donnée — la liste
  de ces colonnes "extra" varie donc d'un fichier à l'autre.
- **Exception** : la colonne "Total somme" n'est **pas** reprise dans
  l'export — "Salaire brut" (calculée, voir plus bas) la remplace. C'est la
  seule colonne du fichier déposé qui n'apparaît jamais dans le tableau de
  sortie.

Le classement des métiers par département (`METIER_TO_DEPT`, `src/generator.js`)
est extrait directement des référentiels LoV `Job` / `Department` de la base
de données itsi-production (382 intitulés de poste → 23 départements), avec
les mêmes noms de département que ceux utilisés côté API/back-office (ex.
"Réalisation - Mise en scène / ADS", "Régie / PA (Locations-Unit-LO-Transport)",
"Autres"…). Un métier absent de cette table est classé dans "Autres" et
signalé à l'utilisateur dans l'interface plutôt que de faire échouer la
génération ; à distinguer d'un métier valide déjà rattaché à "Autres" dans le
référentiel (ex. stagiaire, documentaliste), qui n'est pas signalé.

**Ordre des départements** : les salariés sont triés par matricule avant
d'être regroupés par département (`compareMatricules()`), et les
départements sortent dans leur ordre de première apparition — sans liste
d'ordre codée en dur. Le matricule est attribué avec le code CNC du
département en préfixe, donc trier par matricule revient à trier par
département. Ce mécanisme reproduit volontairement celui
d'itsi-production (`ProjectJobController::index`, tri par référence puis
regroupement Laravel qui conserve l'ordre de première rencontre).

## Sous-totaux

Trois niveaux de sous-total sont générés, avec des formules SOMME qui ne se
recoupent jamais (pas de double comptage) :

- **Par contrat** (`Code contrat`) : chaque contrat (dès qu'il a un code) a
  sa ligne de sous-total, même s'il n'a qu'un seul bulletin sur la période —
  cela permet de repérer chaque contrat d'un coup d'œil dans le tableau.
- **Par département** : somme des sous-totaux contrat du département.
- **Total général** : somme des sous-totaux de département.

Chaque niveau de sous-total additionne **toutes les colonnes à l'exception
des informations contrat** (Nom, Taux horaire, dates…) : Jour(s) travaillés,
variables de paie (heures et euros), colonnes Garantie Minimale, et totaux
bruts (Salaire brut, Coût employeur…). `shouldSumForSubtotal()` dans
`src/generator.js` centralise cette règle.

Une colonne dont la formule référence la ligne "TOTAL" par salarié du
fichier source (ex. "Ratio MG", qui répartissait la Garantie Minimale entre
les bulletins d'un même salarié) ne peut pas être traduite telle quelle,
cette ligne n'existant plus dans le nouveau format : sa dernière valeur
calculée est alors figée à la place, et un message le signale dans
l'interface.

La colonne source "Jour(s) travaillés" (liste de dates séparées par ";")
donne lieu à **deux** colonnes dans l'export : "Jours travaillés (dates)",
juste avant, qui liste ces dates en clair (AAAA-MM-JJ séparées par ";",
même format que le fichier source) pour vérification,
et "Jour(s) travaillés" qui garde le nombre de jours (`countJours()`,
`formatWorkedDates()` dans `src/generator.js`) utilisé par les sous-totaux.
La colonne "dates" n'est jamais sommée (texte, pas nombre) : elle reste
vide sur les lignes de sous-total.

## Colonnes calculées "taux horaire x coefficient x heures"

Pour les variables de paie du référentiel LoV `Remuneration Category`
d'itsi-production dont le montant est un pur produit heures x taux x
coefficient (id 2 à 24, puis 40 à 46 — ex. "H. normales", "H. supp. 125%",
"Majo. jour 25%", "Retrait plafond majo."…), la colonne "(en €)" est
**calculée** par formule `Taux horaire x coefficient x colonne (en h)`
plutôt que recopiée/traduite depuis le fichier source. Le fichier source
fige en effet le taux horaire en valeur littérale directement dans sa
formule (ex. `=68.0408999*N2`) : modifier le taux horaire dans le fichier
généré n'y recalculait donc rien. La formule vivante corrige ce point pour
ces colonnes (`HOUR_RATE_COEF` dans `src/generator.js`).

Les indemnités à montant libre (transport, repas, cachets, défraiements…)
ne sont pas concernées et gardent leur comportement précédent (valeur ou
formule traduite depuis le fichier source).

## Mise en forme et ordre des colonnes

Les colonnes sont regroupées en 4 zones pastel, séparées par une bordure
verticale (`sectionForLabel()` / `columnPlan` dans `src/generator.js`) —
l'ordre des colonnes suit ce même regroupement, pas seulement leur couleur :

1. **Informations contrat** (bleu) : Code bulletin, Statut, Code contrat,
   Nom, Prénom, Métier, dates, taux horaire…
2. **Variables de paie** (vert) : Jours travaillés (dates), Jour(s)
   travaillés, toutes les colonnes heures/euros (H. normales, majorations,
   indemnités…), ainsi que les colonnes sans équivalent standard type
   Cachet/Déf./Indem.
3. **Garantie Minimale** (rose) : les anciennes colonnes de travail Total
   base/MG/Ratio MG/Supp ap. MG, juste avant les totaux.
4. **Totaux bruts** (violet), **toujours en toutes dernières colonnes** :
   Salaire brut, Coût employeur, Salaire net imposable, Salaire net, puis
   "Total indemnité (NS)" si le fichier déposé contient au moins une colonne
   "non soumise" (ex. "Indem. Matériel (NS)", "Déf. non soumis") — sinon
   cette dernière colonne n'apparaît pas.

`Salaire brut (en €)` est **calculée** (formule SOMME de toutes les colonnes
euros de la zone "variables de paie" de la ligne, **à l'exclusion des
colonnes "non soumises"**), et non recopiée depuis le fichier source :
modifier un nombre d'heures recalcule aussi le salaire brut, comme le reste
de l'outil. Elle remplace "Total somme", qui n'est plus reprise dans
l'export.

`Total indemnité (NS)` est également **calculée** : formule SOMME de
l'ensemble des colonnes "non soumises" en euros de la ligne — repérées soit
par l'abréviation "(NS)" (ex. "Indem. Matériel (NS)"), soit en toutes
lettres (ex. "Déf. non soumis") — via `isNsLabel()` / `nsSumCols` dans
`src/generator.js`, tenues à part de "Salaire brut".

Les 4 premières lignes (bandeau + en-têtes) sont figées pour rester visibles
au défilement.
