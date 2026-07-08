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

Le mapping des colonnes source → colonnes cible et le classement des
métiers par département sont définis dans `src/generator.js`
(`METIER_TO_DEPT`, `IDENTITY_MAP`, `CALC_MAP`, `EXTRA_COLS`) — à adapter si
de nouveaux intitulés de poste ou de nouvelles colonnes apparaissent dans
les fichiers source.

Quelques colonnes du fichier source n'ont pas d'équivalent dans le format
cible :

- Les colonnes de travail internes (Total base, MG, Ratio MG, Supp ap. MG,
  Total somme) servaient uniquement à répartir la Garantie Minimale entre
  deux bulletins d'un même salarié ; elles ne sont pas reprises.
- Les colonnes Déf. soumis/non soumis, Rep. continue, Indem. Matériel et
  Indem. MàL n'ont pas de colonne dédiée dans le format cible : elles sont
  ajoutées en fin de tableau pour ne perdre aucune donnée.
