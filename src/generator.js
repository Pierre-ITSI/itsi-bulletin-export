// Génère un fichier "export bulletin" (format ITSI / type TOURNAGE) à partir
// d'un fichier "Combine" (type Marie-Christine), en conservant les formules
// de calcul (taux horaire x heures, etc.) pour que le fichier reste un outil
// de travail modifiable.
//
// Portage JS (exceljs) du script generate_bulletin_export.py.

import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

// --------------------------------------------------------------------------
// 1. Correspondance des colonnes source -> colonnes cible, PAR NOM D'EN-TÊTE
// --------------------------------------------------------------------------
//
// Les fichiers "Combine" n'ont pas toujours les colonnes dans le même ordre
// ni le même jeu de colonnes (une production qui utilise des "cachets" ou
// des indemnités spécifiques ajoute/retire des colonnes). Le mapping se fait
// donc entièrement sur le LIBELLÉ de la ligne d'en-tête du fichier déposé,
// jamais sur la position (lettre de colonne) : si l'ordre change, tout
// continue à fonctionner tant que l'intitulé de colonne reste le même.

// Quelques champs "identité" portent un nom différent entre le fichier
// source et le format cible ; tous les autres (Nom, Prénom, Métier, Date de
// début, Date de fin, Taux horaire, Jour(s) travaillés, Abattement, et
// toutes les paires heures/euros) sont retrouvés par correspondance directe
// de libellé (voir normalizeLabel) sans avoir besoin d'être listés ici.
const IDENTITY_RENAME = {
  "code bulletin": "Code bulletin (itsi)",
  statut: "Statut (itsi)",
  "code contrat": "Code contrat (itsi)",
  "code employe": "Code salarié projet",
  matricule: "Matricule salarié production (itsi)",
};

const STANDARD_HEADERS = [
  "Code bulletin (itsi)", "Statut (itsi)", "Code contrat (itsi)",
  "Code salarié projet", "Matricule salarié production (itsi)", "Nom",
  "Prénom", "Abattement", "Métier", "Date de début", "Date de fin",
  "Taux horaire", "Jour(s) travaillés", "H. normales (en h)",
  "H. normales (en €)", "H. supp. 125% (en h)", "H. supp. 125% (en €)",
  "H. supp. 150% (en h)", "H. supp. 150% (en €)", "H. supp. 175% (en h)",
  "H. supp. 175% (en €)", "H. supp. 200% (en h)", "H. supp. 200% (en €)",
  "Majo. jour 25% (en h)", "Majo. jour 25% (en €)", "Majo. jour 50% (en h)",
  "Majo. jour 50% (en €)", "Majo. jour 100% (en h)",
  "Majo. jour 100% (en €)", "Majo. jour 200% (en h)",
  "Majo. jour 200% (en €)", "Majo. nuit 25% (en h)",
  "Majo. nuit 25% (en €)", "Majo. nuit 50% (en h)", "Majo. nuit 50% (en €)",
  "Majo. nuit 100% (en h)", "Majo. nuit 100% (en €)",
  "H. anticipées 100% (en h)", "H. anticipées 100% (en €)",
  "H. anticipées 50% (en h)", "H. anticipées 50% (en €)",
  "Majo. dimanche 100% (en h)", "Majo. dimanche 100% (en €)",
  "Majo. dimanche 50% (en h)", "Majo. dimanche 50% (en €)",
  "Majo. férié 50% (en h)", "Majo. férié 50% (en €)",
  "Majo. férié 100% (en h)", "Majo. férié 100% (en €)",
  "Majo. férié 200% (en h)", "Majo. férié 200% (en €)",
  "Majo. 6ème jour 100% (en h)", "Majo. 6ème jour 100% (en €)",
  "Récup. dimanche (en h)", "Récup. dimanche (en €)",
  "Récup. férié (en h)", "Récup. férié (en €)",
  "Récup. 6ème jour (en h)", "Récup. 6ème jour (en €)",
  "Indem. transport (Prépa) (en h)", "Indem. transport (Prépa) (en €)",
  "Indem. transport (Tournage) (en h)", "Indem. transport (Tournage) (en €)",
  "Indem. transport (PostProduction) (en h)",
  "Indem. transport (PostProduction) (en €)",
  "Indem. voyage (Prépa) (en h)", "Indem. voyage (Prépa) (en €)",
  "Indem. voyage (Tournage) (en h)", "Indem. voyage (Tournage) (en €)",
  "Indem. voyage (PostProduction) (en h)",
  "Indem. voyage (PostProduction) (en €)", "Indem. repas (RP) (en h)",
  "Indem. repas (RP) (en €)", "Indem. repas (Hors RP) (en h)",
  "Indem. repas (Hors RP) (en €)", "Indem. repas (Étranger) (en h)",
  "Indem. repas (Étranger) (en €)", "Indem. repas (Prépa) (en h)",
  "Indem. repas (Prépa) (en €)", "Indem. repas (Postproduction) (en h)",
  "Indem. repas (Postproduction) (en €)", "Indem. casse croûte (RP) (en h)",
  "Indem. casse croûte (RP) (en €)", "Indem. casse croûte (Hors RP) (en h)",
  "Indem. casse croûte (Hors RP) (en €)",
  "Indem. casse croûte (Étranger) (en h)",
  "Indem. casse croûte (Étranger) (en €)",
  "Indem. casse croûte (Prépa) (en h)", "Indem. casse croûte (Prépa) (en €)",
  "Indem. continue (en h)", "Indem. continue (en €)",
  "Retrait équivalence (en h)", "Retrait équivalence (en €)",
  "Retrait equi - H. supp. 125% (en h)", "Retrait equi - H. supp. 125% (en €)",
  "Retrait equi - H. supp. 150% (en h)", "Retrait equi - H. supp. 150% (en €)",
  "Retrait equi - H. supp. 175% (en h)", "Retrait equi - H. supp. 175% (en €)",
  "Retrait equi - H. supp. 200% (en h)", "Retrait equi - H. supp. 200% (en €)",
  "Retrait plafond majo. (en h)", "Retrait plafond majo. (en €)",
  "Prime except. (en h)", "Prime except. (en €)", "Salaire brut (en h)",
  "Salaire brut (en €)", "Coût employeur (en h)", "Coût employeur (en €)",
  "Salaire net imposable (en h)", "Salaire net imposable (en €)",
  "Salaire net (en h)", "Salaire net (en €)",
];

// Largeur de colonne calculée à partir du libellé d'en-tête, pas fixée en
// dur : chaque colonne est assez large pour son mot le plus long, le reste
// du libellé passant à la ligne (cf. wrapText sur la ligne d'en-tête) plutôt
// que d'être masqué.
function widthForLabel(label) {
  const longestWord = Math.max(...label.split(" ").map((w) => w.length));
  return Math.max(11, Math.min(22, longestWord + 4));
}
const HEADER_ROW_HEIGHT = 45;

// Normalise un libellé de colonne pour comparaison : minuscule, accents
// retirés, espaces réduits, et "(h)"/"(€)" unifiés avec leurs variantes
// "(en h)"/"(en €)" utilisées dans le format cible.
function normalizeLabel(s) {
  if (!s) return "";
  let x = String(s).trim().toLowerCase();
  x = x.replace(/\(en\s*€\)/g, "(unit_eur)").replace(/\(€\)/g, "(unit_eur)");
  x = x.replace(/\(en\s*h\)/g, "(unit_h)").replace(/\(h\)/g, "(unit_h)");
  x = stripAccents(x);
  x = x.replace(/\s+/g, " ").trim();
  return x;
}

const TARGET_NORM_TO_LABEL = new Map(STANDARD_HEADERS.map((l) => [normalizeLabel(l), l]));

// Pour les colonnes "extra" (sans équivalent standard), on garde le libellé
// d'origine mais on harmonise le suffixe "(h)"/"(€)" en "(en h)"/"(en €)"
// comme dans le reste du format cible, pour que le format des nombres et la
// lisibilité restent cohérents.
function toDisplayLabel(label) {
  return label.replace(/\(h\)\s*$/i, "(en h)").replace(/\(€\)\s*$/i, "(en €)");
}

// Résout, pour LE FICHIER DÉPOSÉ (ses propres lettres de colonnes, quel que
// soit leur ordre), quelle colonne cible standard chaque colonne source
// alimente, et quelles colonnes n'ont aucun équivalent standard (à ajouter
// en fin de tableau pour ne perdre aucune donnée).
function resolveColumnMapping(headers) {
  const letterToTargetLabel = {};
  const extraEntries = []; // [srcLetter, originalLabel] dans l'ordre du fichier source
  const letters = Object.keys(headers).sort(
    (a, b) => colIndexFromLetter(a) - colIndexFromLetter(b)
  );

  for (const letter of letters) {
    const rawLabel = headers[letter];
    if (!rawLabel) continue;
    const label = String(rawLabel).trim();
    const norm = normalizeLabel(label);
    if (!norm) continue;

    const renamedLabel = IDENTITY_RENAME[norm] || label;
    const renamedNorm = normalizeLabel(renamedLabel);
    if (TARGET_NORM_TO_LABEL.has(renamedNorm)) {
      letterToTargetLabel[letter] = TARGET_NORM_TO_LABEL.get(renamedNorm);
    } else {
      extraEntries.push([letter, toDisplayLabel(label)]);
    }
  }

  return { letterToTargetLabel, extraEntries };
}

// --------------------------------------------------------------------------
// 2. Classement des métiers par département de tournage
// --------------------------------------------------------------------------

// Table extraite directement des référentiels LoV `Job` / `Department` de la
// base de données itsi-production (382 métiers -> 23 départements) : les
// mêmes noms de département que ceux utilisés côté API/back-office.
const METIER_TO_DEPT = {
  // Réalisation - Mise en scène / ADS
  "1er assistant realisateur": "Réalisation - Mise en scène / ADS",
  "1er assistant realisateur (specialise)": "Réalisation - Mise en scène / ADS",
  "2eme assistant realisateur": "Réalisation - Mise en scène / ADS",
  "2eme assistant realisateur (specialise)": "Réalisation - Mise en scène / ADS",
  "assistant au charge de la figuration cinema": "Réalisation - Mise en scène / ADS",
  "assistant realisateur": "Réalisation - Mise en scène / ADS",
  "assistant realisateur adjoint": "Réalisation - Mise en scène / ADS",
  "assistant scripte adjoint": "Réalisation - Mise en scène / ADS",
  "assistant scripte cinema": "Réalisation - Mise en scène / ADS",
  "auxiliaire a la realisation cinema": "Réalisation - Mise en scène / ADS",
  "charge de la figuration cinema": "Réalisation - Mise en scène / ADS",
  "conseiller technique a la realisation cinema": "Réalisation - Mise en scène / ADS",
  "premier assistant a la distribution des roles cinema": "Réalisation - Mise en scène / ADS",
  "premier assistant realisateur cinema": "Réalisation - Mise en scène / ADS",
  "realisateur": "Réalisation - Mise en scène / ADS",
  "realisateur cinema": "Réalisation - Mise en scène / ADS",
  "realisateur de films publicitaires": "Réalisation - Mise en scène / ADS",
  "repetiteur cinema": "Réalisation - Mise en scène / ADS",
  "scripte": "Réalisation - Mise en scène / ADS",
  "scripte (specialise)": "Réalisation - Mise en scène / ADS",
  "scripte cinema": "Réalisation - Mise en scène / ADS",
  "second assistant realisateur cinema": "Réalisation - Mise en scène / ADS",
  "technicien realisateur 2eme equipe cinema": "Réalisation - Mise en scène / ADS",

  // Régie / PA (Locations-Unit-LO-Transport)
  "assistant regisseur adjoint": "Régie / PA (Locations-Unit-LO-Transport)",
  "auxiliaire a la regie cinema": "Régie / PA (Locations-Unit-LO-Transport)",
  "chauffeur": "Régie / PA (Locations-Unit-LO-Transport)",
  "regisseur": "Régie / PA (Locations-Unit-LO-Transport)",
  "regisseur (specialise)": "Régie / PA (Locations-Unit-LO-Transport)",
  "regisseur adjoint": "Régie / PA (Locations-Unit-LO-Transport)",
  "regisseur adjoint (specialise)": "Régie / PA (Locations-Unit-LO-Transport)",
  "regisseur adjoint cinema": "Régie / PA (Locations-Unit-LO-Transport)",
  "regisseur general": "Régie / PA (Locations-Unit-LO-Transport)",
  "regisseur general (specialise)": "Régie / PA (Locations-Unit-LO-Transport)",
  "regisseur general cinema": "Régie / PA (Locations-Unit-LO-Transport)",
  "responsable reperages": "Régie / PA (Locations-Unit-LO-Transport)",
  "responsable reperages (specialise)": "Régie / PA (Locations-Unit-LO-Transport)",

  // Administration / Production
  "administrateur adjoint comptable cinema": "Administration / Production",
  "administrateur de production": "Administration / Production",
  "administrateur de production (specialise)": "Administration / Production",
  "administrateur de production cinema": "Administration / Production",
  "assistant comptable de production cinema": "Administration / Production",
  "assistant de post production": "Administration / Production",
  "assistant de production": "Administration / Production",
  "assistant de production (specialise)": "Administration / Production",
  "assistant de production adjoint": "Administration / Production",
  "charge de postproduction": "Administration / Production",
  "charge de production": "Administration / Production",
  "comptable de production": "Administration / Production",
  "comptable de production (specialise)": "Administration / Production",
  "coordinateur de post-production cinema": "Administration / Production",
  "coordinateur de production web": "Administration / Production",
  "directeur de postproduction": "Administration / Production",
  "directeur de production": "Administration / Production",
  "directeur de production (specialise)": "Administration / Production",
  "directeur de production cinema": "Administration / Production",
  "producteur artistique": "Administration / Production",
  "producteur executif": "Administration / Production",
  "regisseur de plateau": "Administration / Production",
  "secretaire de production cinema": "Administration / Production",

  // Image - prise de vues / Camera - Video & DIT
  "1er assistant opv": "Image - prise de vues / Camera - Video & DIT",
  "1er assistant opv (specialise)": "Image - prise de vues / Camera - Video & DIT",
  "2eme assistant opv": "Image - prise de vues / Camera - Video & DIT",
  "2eme assistant opv (specialise)": "Image - prise de vues / Camera - Video & DIT",
  "assistant opv adjoint": "Image - prise de vues / Camera - Video & DIT",
  "cadreur": "Image - prise de vues / Camera - Video & DIT",
  "cadreur (specialise)": "Image - prise de vues / Camera - Video & DIT",
  "cadreur cinema": "Image - prise de vues / Camera - Video & DIT",
  "cadreur specialise cinema": "Image - prise de vues / Camera - Video & DIT",
  "chef operateur prise de vues": "Image - prise de vues / Camera - Video & DIT",
  "data manager": "Image - prise de vues / Camera - Video & DIT",
  "data manager (audiovisuel)": "Image - prise de vues / Camera - Video & DIT",
  "deuxieme assistant operateur cinema": "Image - prise de vues / Camera - Video & DIT",
  "directeur de la photographie cinema": "Image - prise de vues / Camera - Video & DIT",
  "directeur photo": "Image - prise de vues / Camera - Video & DIT",
  "directeur photo (specialise)": "Image - prise de vues / Camera - Video & DIT",
  "dit": "Image - prise de vues / Camera - Video & DIT",
  "dit (audiovisuel)": "Image - prise de vues / Camera - Video & DIT",
  "operateur de prise de vue (opv)": "Image - prise de vues / Camera - Video & DIT",
  "operateur de prise de vue (opv) (specialise)": "Image - prise de vues / Camera - Video & DIT",
  "operateur magneto ralenti": "Image - prise de vues / Camera - Video & DIT",
  "operateur magnetoscope": "Image - prise de vues / Camera - Video & DIT",
  "operateur qtake": "Image - prise de vues / Camera - Video & DIT",
  "operateur qtake (ccnpa)": "Image - prise de vues / Camera - Video & DIT",
  "operateur special (steadicamer)": "Image - prise de vues / Camera - Video & DIT",
  "operateur special (steadicamer) (specialise)": "Image - prise de vues / Camera - Video & DIT",
  "photographe de plateau cinema": "Image - prise de vues / Camera - Video & DIT",
  "pointeur": "Image - prise de vues / Camera - Video & DIT",
  "pointeur (specialise)": "Image - prise de vues / Camera - Video & DIT",
  "premier assistant operateur cinema": "Image - prise de vues / Camera - Video & DIT",
  "pupitreur lumiere": "Image - prise de vues / Camera - Video & DIT",
  "tech dappareils telecommandes (pdv) cine": "Image - prise de vues / Camera - Video & DIT",
  "technicien retour image cinema": "Image - prise de vues / Camera - Video & DIT",
  "technicien video": "Image - prise de vues / Camera - Video & DIT",
  "technicien video web": "Image - prise de vues / Camera - Video & DIT",
  "to_delete_dit - data - ca": "Image - prise de vues / Camera - Video & DIT",
  "z-old_jobs": "Image - prise de vues / Camera - Video & DIT",

  // Électricité / Electrical
  "assistant lumiere": "Électricité / Electrical",
  "assistant lumiere (specialise)": "Électricité / Electrical",
  "auxiliaire a la regie cinema (affecte electricite)": "Électricité / Electrical",
  "blocker": "Électricité / Electrical",
  "chef electricien": "Électricité / Electrical",
  "chef electricien prise de vues cinema": "Électricité / Electrical",
  "conducteur de groupe": "Électricité / Electrical",
  "conducteur de groupe cinema": "Électricité / Electrical",
  "eclairagiste": "Électricité / Electrical",
  "electricien": "Électricité / Electrical",
  "electricien prise de vues cinema": "Électricité / Electrical",
  "sous-chef electricien prise de vues cinema": "Électricité / Electrical",

  // Machinerie / Grip
  "auxiliaire a la regie cinema (affecte machinerie)": "Machinerie / Grip",
  "chef machiniste": "Machinerie / Grip",
  "chef machiniste prise de vues cinema": "Machinerie / Grip",
  "machiniste": "Machinerie / Grip",
  "machiniste de prise de vues cinema": "Machinerie / Grip",
  "rigger": "Machinerie / Grip",
  "sous-chef machiniste de prise de vues cinema": "Machinerie / Grip",

  // Costumes / Costume
  "1er assistant costume cinema": "Costumes / Costume",
  "auxiliaire a la regie cinema (affecte costumes)": "Costumes / Costume",
  "chef costumier": "Costumes / Costume",
  "chef costumier (specialise)": "Costumes / Costume",
  "chef costumier cinema": "Costumes / Costume",
  "chef d'atelier costumes cinema": "Costumes / Costume",
  "costumier": "Costumes / Costume",
  "costumier (specialise)": "Costumes / Costume",
  "costumier cinema": "Costumes / Costume",
  "couturier costumes cinema": "Costumes / Costume",
  "createur de costume": "Costumes / Costume",
  "createur de costume (specialise)": "Costumes / Costume",
  "createur de costume cinema": "Costumes / Costume",
  "habilleur": "Costumes / Costume",
  "habilleur (specialise)": "Costumes / Costume",
  "habilleur cinema": "Costumes / Costume",
  "regisseur adjoint cinema (affecte costumes)": "Costumes / Costume",
  "styliste": "Costumes / Costume",
  "teinturier patineur costumes cinema": "Costumes / Costume",

  // Maquillage / Make up
  "chef maquilleur": "Maquillage / Make up",
  "chef maquilleur (specialise)": "Maquillage / Make up",
  "chef maquilleur cinema": "Maquillage / Make up",
  "maquilleur": "Maquillage / Make up",
  "maquilleur (specialise)": "Maquillage / Make up",
  "maquilleur cinema": "Maquillage / Make up",
  "maquilleur et coiffeur effets speciaux": "Maquillage / Make up",

  // Coiffure / Hair
  "chef coiffeur cinema": "Coiffure / Hair",
  "coiffeur": "Coiffure / Hair",
  "coiffeur (specialise)": "Coiffure / Hair",
  "coiffeur cinema": "Coiffure / Hair",
  "coiffeur perruquier": "Coiffure / Hair",
  "coiffeur perruquier (specialise)": "Coiffure / Hair",

  // Son / Sound
  "1er assistant operateur du son cinema": "Son / Sound",
  "1er assistant son": "Son / Sound",
  "1er assistant son (specialise)": "Son / Sound",
  "2eme assistant operateur du son cinema": "Son / Sound",
  "assistant son": "Son / Sound",
  "assistant son adjoint": "Son / Sound",
  "auxiliaire a la regie cinema (affecte son)": "Son / Sound",
  "backliner": "Son / Sound",
  "chef operateur de son cinema": "Son / Sound",
  "chef ops (specialise) / chef operateur prise de son": "Son / Sound",
  "chef ops / chef operateur prise de son": "Son / Sound",
  "ingenieur du son": "Son / Sound",
  "ingenieur du son (specialise)": "Son / Sound",
  "operateur prise de son (ops)": "Son / Sound",
  "operateur synthetiseur": "Son / Sound",
  "perchiste": "Son / Sound",
  "perchiste (specialise)": "Son / Sound",
  "technicien instrument": "Son / Sound",

  // Décoration / Art Department
  "1er assistant decorateur": "Décoration / Art Department",
  "1er assistant decorateur (specialise)": "Décoration / Art Department",
  "2eme assistant decorateur": "Décoration / Art Department",
  "2eme assistant decorateur (specialise)": "Décoration / Art Department",
  "accessoiriste de decor cinema": "Décoration / Art Department",
  "assistant decorateur adjoint": "Décoration / Art Department",
  "chef decorateur": "Décoration / Art Department",
  "chef decorateur (specialise)": "Décoration / Art Department",
  "chef decorateur cinema": "Décoration / Art Department",
  "chef tapissier de decor cinema": "Décoration / Art Department",
  "decorateur": "Décoration / Art Department",
  "decorateur (specialise)": "Décoration / Art Department",
  "deuxieme assistant decorateur cinema": "Décoration / Art Department",
  "electricien deco": "Décoration / Art Department",
  "ensemblier - decorateur": "Décoration / Art Department",
  "ensemblier - decorateur (specialise)": "Décoration / Art Department",
  "ensemblier cinema": "Décoration / Art Department",
  "ensemblier decorateur cinema": "Décoration / Art Department",
  "illustrateur de decor cinema": "Décoration / Art Department",
  "infographiste de decor cinema": "Décoration / Art Department",
  "premier assistant decorateur cinema": "Décoration / Art Department",
  "regisseur dexterieurs": "Décoration / Art Department",
  "regisseur dexterieurs (specialise)": "Décoration / Art Department",
  "rippeur": "Décoration / Art Department",
  "tapissier de decor": "Décoration / Art Department",
  "tapissier de decor cinema": "Décoration / Art Department",
  "to_delete_menuisier de decor": "Décoration / Art Department",
  "troisieme assistant decorateur cinema": "Décoration / Art Department",

  // Post-Production Image / Editing-Color Grading
  "1er assistant monteur cinema": "Post-Production Image / Editing-Color Grading",
  "assistant monteur": "Post-Production Image / Editing-Color Grading",
  "assistant monteur (specialise)": "Post-Production Image / Editing-Color Grading",
  "assistant monteur adjoint": "Post-Production Image / Editing-Color Grading",
  "chef monteur": "Post-Production Image / Editing-Color Grading",
  "chef monteur (specialise)": "Post-Production Image / Editing-Color Grading",
  "chef monteur cinema": "Post-Production Image / Editing-Color Grading",
  "conformateur": "Post-Production Image / Editing-Color Grading",
  "deuxieme assistant monteur cinema": "Post-Production Image / Editing-Color Grading",
  "etalonneur": "Post-Production Image / Editing-Color Grading",
  "infographiste": "Post-Production Image / Editing-Color Grading",
  "monteur": "Post-Production Image / Editing-Color Grading",

  // Construction de décors / Construction department
  "accessoiriste": "Construction de décors / Construction department",
  "accessoiriste (specialise)": "Construction de décors / Construction department",
  "accessoiriste de plateau cinema": "Construction de décors / Construction department",
  "chef constructeur": "Construction de décors / Construction department",
  "chef constructeur cinema": "Construction de décors / Construction department",
  "chef d'equipe de decor": "Construction de décors / Construction department",
  "chef electricien de construction cinema": "Construction de décors / Construction department",
  "chef machiniste de construction cinema": "Construction de décors / Construction department",
  "chef menuisier de decor cinema": "Construction de décors / Construction department",
  "chef peintre de decor cinema": "Construction de décors / Construction department",
  "chef sculpteur de decor cinema": "Construction de décors / Construction department",
  "chef serrurier de decor cinema": "Construction de décors / Construction department",
  "chef staffeur de decor cinema": "Construction de décors / Construction department",
  "constructeur de decor": "Construction de décors / Construction department",
  "dessinateur en decor": "Construction de décors / Construction department",
  "dessinateur en decor (specialise)": "Construction de décors / Construction department",
  "electricien de construction cinema": "Construction de décors / Construction department",
  "machiniste de construction cinema": "Construction de décors / Construction department",
  "machiniste decorateur": "Construction de décors / Construction department",
  "macon de decor": "Construction de décors / Construction department",
  "macon de decor cinema": "Construction de décors / Construction department",
  "maquettiste de decor cinema": "Construction de décors / Construction department",
  "mecanicien de decor": "Construction de décors / Construction department",
  "menuisier de decor": "Construction de décors / Construction department",
  "menuisier de decor cinema": "Construction de décors / Construction department",
  "menuisier traceur": "Construction de décors / Construction department",
  "menuisier traceur de decor cinema": "Construction de décors / Construction department",
  "metallier de decor": "Construction de décors / Construction department",
  "peintre d'art de decor cinema": "Construction de décors / Construction department",
  "peintre de decor cinema": "Construction de décors / Construction department",
  "peintre de decors": "Construction de décors / Construction department",
  "peintre en lettres / en faux bois de decor": "Construction de décors / Construction department",
  "peintre en lettres de decor cinema": "Construction de décors / Construction department",
  "peintre faux bois et patine decor cinema": "Construction de décors / Construction department",
  "sculpteur de decor cinema": "Construction de décors / Construction department",
  "serrurier de decor": "Construction de décors / Construction department",
  "serrurier de decor cinema": "Construction de décors / Construction department",
  "sous-chef electricien de construction cinema": "Construction de décors / Construction department",
  "sous-chef machiniste de construction cinema": "Construction de décors / Construction department",
  "sous-chef menuisier de decor cinema": "Construction de décors / Construction department",
  "sous-chef peintre de decor cinema": "Construction de décors / Construction department",
  "sous-chef staffeur de decor cinema": "Construction de décors / Construction department",
  "staffeur de decor": "Construction de décors / Construction department",
  "staffeur de decor cinema": "Construction de décors / Construction department",
  "toupilleur de decor": "Construction de décors / Construction department",
  "toupilleur de decor cinema": "Construction de décors / Construction department",
  "traceur de decor": "Construction de décors / Construction department",

  // Liste artistique - Rôles principaux / MAIN CAST
  "acteur": "Liste artistique - Rôles principaux / MAIN CAST",
  "danseur": "Liste artistique - Rôles principaux / MAIN CAST",
  "petit role": "Liste artistique - Rôles principaux / MAIN CAST",
  "petit role etranger": "Liste artistique - Rôles principaux / MAIN CAST",
  "premier role": "Liste artistique - Rôles principaux / MAIN CAST",
  "premier role etranger": "Liste artistique - Rôles principaux / MAIN CAST",
  "second role": "Liste artistique - Rôles principaux / MAIN CAST",
  "second role etranger": "Liste artistique - Rôles principaux / MAIN CAST",

  // Plateau / Set
  "aide de plateau": "Plateau / Set",
  "animateur": "Plateau / Set",
  "assistant d'emission": "Plateau / Set",
  "assistant technique web": "Plateau / Set",
  "chauffeur de salle": "Plateau / Set",
  "chef de plateau": "Plateau / Set",
  "coordinateur de diffusion web": "Plateau / Set",
  "coordinateur demission": "Plateau / Set",
  "ingenieur de la vision": "Plateau / Set",
  "ingenieur de la vision adjoint": "Plateau / Set",
  "operateur multicam web": "Plateau / Set",
  "operateur regie video": "Plateau / Set",
  "operateur web": "Plateau / Set",

  // Post-Production Son / Sound Post Production
  "assistant bruiteur": "Post-Production Son / Sound Post Production",
  "assistant mixeur cinema": "Post-Production Son / Sound Post Production",
  "assistant monteur son cinema": "Post-Production Son / Sound Post Production",
  "bruiteur": "Post-Production Son / Sound Post Production",
  "chef monteur son cinema": "Post-Production Son / Sound Post Production",
  "illustrateur sonore": "Post-Production Son / Sound Post Production",
  "mixeur": "Post-Production Son / Sound Post Production",
  "mixeur (directs)": "Post-Production Son / Sound Post Production",
  "mixeur cinema": "Post-Production Son / Sound Post Production",
  "to_delete_1er assistant monteur son cinema - cancel": "Post-Production Son / Sound Post Production",

  // Effets Spéciaux / SFX
  "assistant effets physiques cinema": "Effets Spéciaux / SFX",
  "superviseur d'effets physiques cinema": "Effets Spéciaux / SFX",
  "superviseur d'effets speciaux postproduction": "Effets Spéciaux / SFX",
  "superviseur effets speciaux image": "Effets Spéciaux / SFX",
  "technicien truquiste": "Effets Spéciaux / SFX",
  "truquiste": "Effets Spéciaux / SFX",

  // Autres
  "charge de recherche": "Autres",
  "charge de selection": "Autres",
  "charge denquete": "Autres",
  "collaborateur artistique": "Autres",
  "collaborateur de selection": "Autres",
  "concepteur de programme web": "Autres",
  "concepteur web": "Autres",
  "conseiller artistique demission": "Autres",
  "conseiller technique realisation": "Autres",
  "coordinateur decriture (ex script editeur)": "Autres",
  "designer web": "Autres",
  "directeur artistique": "Autres",
  "directeur de collection": "Autres",
  "directeur de jeux": "Autres",
  "directeur de la distribution": "Autres",
  "directeur de la distribution (specialise)": "Autres",
  "directeur de programmation": "Autres",
  "directeur de selection": "Autres",
  "directeur des dialogues": "Autres",
  "documentaliste": "Autres",
  "dresseur": "Autres",
  "editeur artistique web": "Autres",
  "enqueteur": "Autres",
  "gestionnaire de diffusion internet (traffic manager)": "Autres",
  "intervenant": "Autres",
  "operateur de transfert et de traitement numerique": "Autres",
  "permanent cadre": "Autres",
  "permanent non cadre": "Autres",
  "photographe de plateau": "Autres",
  "photographe de plateau (specialise)": "Autres",
  "preparateur de questions": "Autres",
  "programmateur artistique demission": "Autres",
  "prothesiste": "Autres",
  "recherchiste": "Autres",
  "regisseur d'exterieur cinema": "Autres",
  "regulateur de stationnement": "Autres",
  "repetiteur": "Autres",
  "responsable de questions": "Autres",
  "responsable de recherche": "Autres",
  "responsable denquete": "Autres",
  "secretaire de production": "Autres",
  "secretaire de production (specialise)": "Autres",
  "stagiaire": "Autres",
  "technicien de developpement web": "Autres",
  "to_delete_doublon": "Autres",

  // Acteurs de complément / Background actors
  "acteur de complement": "Acteurs de complément / Background actors",
  "acteur de complement mineur": "Acteurs de complément / Background actors",
  "acteur de complement region": "Acteurs de complément / Background actors",
  "cascadeur": "Acteurs de complément / Background actors",
  "doublure": "Acteurs de complément / Background actors",
  "doublure image": "Acteurs de complément / Background actors",
  "doublure image mineure": "Acteurs de complément / Background actors",
  "doublure lumiere": "Acteurs de complément / Background actors",
  "doublure polyvalente": "Acteurs de complément / Background actors",
  "doublure simple texte": "Acteurs de complément / Background actors",
  "figurant": "Acteurs de complément / Background actors",
  "figurant mineur": "Acteurs de complément / Background actors",
  "figurant region": "Acteurs de complément / Background actors",
  "silhouette muette": "Acteurs de complément / Background actors",
  "silhouette muette mineure": "Acteurs de complément / Background actors",
  "silhouette muette region": "Acteurs de complément / Background actors",
  "silhouette parlante": "Acteurs de complément / Background actors",
  "silhouette parlante mineure": "Acteurs de complément / Background actors",
  "silhouette parlante region": "Acteurs de complément / Background actors",

  // Artistes
  "acteur de synchronisation": "Artistes",
  "arrangeur orchestrateur": "Artistes",
  "auteur": "Artistes",
  "cavalier": "Artistes",
  "chanteur": "Artistes",
  "chanteur soliste": "Artistes",
  "chef d'orchestre": "Artistes",
  "chef d'orchestre symphonique": "Artistes",
  "chef des choeurs": "Artistes",
  "choregraphe": "Artistes",
  "choriste": "Artistes",
  "musicien": "Artistes",
  "musicien soliste concertiste": "Artistes",
  "musicien soliste d'orchestre": "Artistes",
  "regisseur d'orchestre": "Artistes",
  "soliste": "Artistes",
  "speaker": "Artistes",

  // Collaborateurs techniques spécialisés
  "animatronicien": "Collaborateurs techniques spécialisés",
  "animatronicien cinema": "Collaborateurs techniques spécialisés",
  "dompteur": "Collaborateurs techniques spécialisés",
  "responsable des enfants": "Collaborateurs techniques spécialisés",
  "responsable des enfants cinema": "Collaborateurs techniques spécialisés",
  "storyboarder": "Collaborateurs techniques spécialisés",
};

function stripAccents(s) {
  return s.normalize("NFKD").replace(/[̀-ͯ]/g, "");
}

// `recognized` distingue un métier réellement absent de la table (à
// signaler à l'utilisateur) d'un métier valide déjà classé dans le
// département "Autres" du référentiel (ex. stagiaire, documentaliste...).
function classifyMetier(metier) {
  if (!metier) return { dept: "Autres", recognized: false };
  const key = stripAccents(String(metier).trim().toLowerCase());
  const dept = METIER_TO_DEPT[key];
  return dept ? { dept, recognized: true } : { dept: "Autres", recognized: false };
}

// Les matricules sont attribués avec le code CNC du département en préfixe :
// trier par matricule revient donc à trier par département, sans avoir à
// maintenir un ordre de département arbitraire ici. Compare numériquement
// si les deux valeurs sont des nombres, sinon en texte.
function compareMatricules(a, b) {
  const na = Number(a);
  const nb = Number(b);
  if (a != null && b != null && !Number.isNaN(na) && !Number.isNaN(nb)) {
    return na - nb;
  }
  return String(a ?? "").localeCompare(String(b ?? ""));
}

// --------------------------------------------------------------------------
// 3. Traduction des formules (mêmes lignes uniquement : taux x heures)
// --------------------------------------------------------------------------

const CELL_RE = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g;

// Traduit une formule vers la nouvelle ligne/colonnes. Une formule qui
// référence une AUTRE ligne (ex. "Ratio MG" = BE2/BE4, où la ligne 4 était
// la ligne "TOTAL" du salarié dans le fichier source) ne peut pas être
// traduite : cette ligne "TOTAL" n'existe plus dans le format cible (seul
// le sous-total par département est conservé). `crossRow` signale ce cas
// à l'appelant, qui fige alors la valeur calculée plutôt que d'écrire une
// formule cassée.
function translateFormula(formula, srcRow, newRow, colmap) {
  let crossRow = false;
  const text = formula.replace(CELL_RE, (match, d1, col, d2, row) => {
    const rowNum = parseInt(row, 10);
    if (rowNum !== srcRow) {
      crossRow = true;
      return match;
    }
    const newCol = colmap[col.toUpperCase()] || col.toUpperCase();
    return `${d1}${newCol}${d2}${newRow}`;
  });
  return { text, crossRow };
}

// --------------------------------------------------------------------------
// 4. Lecture du fichier source
// --------------------------------------------------------------------------

function colLetterFromIndex(idx) {
  let s = "";
  while (idx > 0) {
    const rem = (idx - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    idx = Math.floor((idx - 1) / 26);
  }
  return s;
}

function colIndexFromLetter(letters) {
  let idx = 0;
  for (const ch of letters.toUpperCase()) {
    idx = idx * 26 + (ch.charCodeAt(0) - 64);
  }
  return idx;
}

function excelSerialToDate(serial) {
  // Excel date serial (1900 system) -> JS Date (UTC midday to avoid TZ shift)
  const utcDays = Math.floor(serial - 25569);
  return new Date(utcDays * 86400 * 1000);
}

function parseDate(v) {
  if (v instanceof Date) return v;
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return excelSerialToDate(v);
  const m = String(v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const [, d, mo, y] = m;
    return new Date(Date.UTC(+y, +mo - 1, +d));
  }
  return v; // laisse tel quel si format inattendu
}

function countJours(v) {
  if (v === null || v === undefined || v === "") return null;
  return String(v).split(";").filter((p) => p.trim()).length;
}

// La colonne source "Jour(s) travaillés" stocke une liste de dates séparées
// par ";" (ex. "2026-06-29;2026-06-30"), au format ISO ou JJ/MM/AAAA selon
// les exports. Formatte cette liste en AAAA-MM-JJ séparées par ";" (même
// format que le fichier source) pour la colonne "Jours travaillés (dates)",
// ajoutée juste avant le nombre de jours calculé par `countJours`.
function parseIsoOrFrenchDate(s) {
  const iso = String(s).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    const [, y, mo, d] = iso;
    return new Date(Date.UTC(+y, +mo - 1, +d));
  }
  const fr = String(s).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (fr) {
    const [, d, mo, y] = fr;
    return new Date(Date.UTC(+y, +mo - 1, +d));
  }
  return null;
}

function formatWorkedDates(v) {
  if (v === null || v === undefined || v === "") return null;
  const parts = String(v).split(";").map((p) => p.trim()).filter(Boolean);
  return parts
    .map((p) => {
      const d = parseIsoOrFrenchDate(p);
      if (!d) return p;
      return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
    })
    .join(";");
}

// La lecture du fichier source passe par SheetJS (xlsx) plutôt qu'exceljs :
// les fichiers "Combine" proviennent souvent de copier-coller entre feuilles
// et contiennent des références de "formules partagées" Excel orphelines
// (ref pointant vers des cellules vides) que le lecteur strict d'exceljs
// rejette ("Shared Formula master must exist..."), alors que SheetJS s'en
// accommode sans problème. exceljs reste utilisé côté écriture (styles).
function cellRawValue(cell) {
  if (!cell) return null;
  if (cell.f !== undefined) {
    // cached : valeur de repli si la formule s'avère intraduisible (cf.
    // translateFormula / crossRow)
    return { formula: `=${cell.f}`, cached: typeof cell.v === "number" ? cell.v : null };
  }
  if (cell.v === undefined) return null;
  return cell.v;
}

export async function readSource(arrayBuffer) {
  const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: "array", cellFormula: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const range = XLSX.utils.decode_range(ws["!ref"]);

  const headers = {};
  let codeBulletinCol = range.s.c; // repli sur la 1ère colonne si l'en-tête n'est pas retrouvé par nom
  for (let c = range.s.c; c <= range.e.c; c++) {
    const letter = colLetterFromIndex(c + 1);
    const cell = ws[XLSX.utils.encode_cell({ r: range.s.r, c })];
    if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== "") {
      const label = String(cell.v).trim();
      headers[letter] = label;
      if (normalizeLabel(label) === "code bulletin") codeBulletinCol = c;
    }
  }

  const rows = [];
  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    // r = ligne d'en-tête ; rowNumber reste 1-indexé pour rester cohérent
    // avec les coordonnées Excel utilisées ailleurs.
    const rowNumber = r + 1;
    const aCell = ws[XLSX.utils.encode_cell({ r, c: codeBulletinCol })];
    const codeBulletin = aCell ? aCell.v : undefined;
    if (codeBulletin === undefined || codeBulletin === null || codeBulletin === "") continue;
    if (String(codeBulletin).trim().toUpperCase() === "TOTAL") continue;

    const data = {};
    for (let c = range.s.c; c <= range.e.c; c++) {
      const letter = colLetterFromIndex(c + 1);
      data[letter] = cellRawValue(ws[XLSX.utils.encode_cell({ r, c })]);
    }
    rows.push({ srcRow: rowNumber, data });
  }
  return { headers, rows };
}

// --------------------------------------------------------------------------
// 5. Construction du classeur de sortie
// --------------------------------------------------------------------------

const FILL_DEPT = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCE5CD" } };
const FILL_CONTRACT = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } };
const FONT_LABEL_BOLD = { name: "Arial", bold: true };
const FONT_DEPT = { name: "Arial", size: 8, italic: true };
const FONT_CONTRACT = { name: "Arial", size: 8, italic: true };
const FMT_HOURS = "#,##0.00";
const FMT_EUROS = '#,##0.00"€"';
const FMT_DATE = "dd/mm/yyyy";
const CENTER = { horizontal: "center", vertical: "middle" };

// -- Regroupement visuel des colonnes en 4 zones, avec un dégradé pastel et
// une bordure verticale à chaque changement de zone : infos contrat, puis
// variables de paie (y compris "Jour(s) travaillés" et les colonnes sans
// équivalent standard type Cachet/Déf./Indem.), puis les colonnes "de
// travail" liées à la Garantie Minimale, et enfin — tout à la fin du
// tableau — les totaux bruts (Salaire brut, Coût employeur, Salaire net
// imposable, Salaire net). L'ORDRE des colonnes suit ce même regroupement
// (cf. buildColumnLayout), pas seulement leur couleur.
const IDENTITY_LABELS = new Set([
  "Code bulletin (itsi)", "Statut (itsi)", "Code contrat (itsi)",
  "Code salarié projet", "Matricule salarié production (itsi)", "Nom",
  "Prénom", "Abattement", "Métier", "Date de début", "Date de fin",
  "Taux horaire",
]);
// Libellé de la colonne calculée "Total indemnité (NS)" (somme des colonnes
// "(NS)" euros, cf. isNsLabel) : constante partagée entre la construction du
// plan de colonnes et l'écriture des lignes.
const NS_TOTAL_LABEL = "Total indemnité (NS)";
const GROSS_TOTAL_LABELS = new Set([
  "Coût employeur (en h)", "Coût employeur (en €)",
  "Salaire brut (en h)", "Salaire brut (en €)",
  "Salaire net imposable (en h)", "Salaire net imposable (en €)",
  "Salaire net (en h)", "Salaire net (en €)",
  NS_TOTAL_LABEL,
]);
// Anciennes colonnes de travail "Garantie Minimale" du fichier source
// (Total base, MG, Ratio MG, Supp ap. MG) : reconnues par motif car elles
// n'ont pas de libellé cible standard (ce sont toujours des colonnes
// "extra"). "Total somme" en fait aussi partie à l'origine, mais n'est plus
// reprise dans l'export : "Salaire brut" (calculée) la remplace (cf.
// isTotalSommeLabel, encore utilisée pour l'exclure des colonnes "extra").
const MG_LABEL_PATTERNS = [
  /^total base\b/, /^mg\b/, /^ratio mg\b/, /^supp .*ap\.? mg/,
];

const SECTION_HEADER_FILL = {
  contrat: { type: "pattern", pattern: "solid", fgColor: { argb: "FFD3E5F7" } }, // bleu pastel
  paie: { type: "pattern", pattern: "solid", fgColor: { argb: "FFD6F0DB" } }, // vert pastel
  mg: { type: "pattern", pattern: "solid", fgColor: { argb: "FFF7D3E6" } }, // rose pastel
  totaux: { type: "pattern", pattern: "solid", fgColor: { argb: "FFE6D9F7" } }, // violet pastel
};
const SECTION_DATA_FILL = {
  contrat: { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3F8FD" } },
  paie: { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2FAF4" } },
  mg: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCEEF5" } },
  totaux: { type: "pattern", pattern: "solid", fgColor: { argb: "FFF6F1FC" } },
};
const SECTION_BORDER = { style: "medium", color: { argb: "FFB0B0B0" } };
const HEADER_BOTTOM_BORDER = { style: "thin", color: { argb: "FF999999" } };

function isMgLabel(label) {
  const norm = normalizeLabel(label);
  return MG_LABEL_PATTERNS.some((re) => re.test(norm));
}

function isTotalSommeLabel(label) {
  return /^total somme\b/.test(normalizeLabel(label));
}

// Colonnes "non soumises" (abrégées "(NS)", ex. "Indem. Matériel (NS)", ou
// en toutes lettres, ex. "Déf. non soumis"). Exclues de la somme "Salaire
// brut" et regroupées à part dans "Total indemnité (NS)".
function isNsLabel(label) {
  return /\(ns\)/i.test(label) || /non soumis/i.test(label);
}

function sectionForLabel(label) {
  if (IDENTITY_LABELS.has(label)) return "contrat";
  if (GROSS_TOTAL_LABELS.has(label) || isTotalSommeLabel(label)) return "totaux";
  if (isMgLabel(label)) return "mg";
  return "paie";
}

function labelIsHours(label) {
  return label.endsWith("(en h)");
}
function labelIsEuros(label) {
  return label.endsWith("(en €)");
}

function writeValue(cell, val, label, srcCol, tgtCol, srcRow, targetRow, colmap, frozenStats) {
  if (label === "Date de début" || label === "Date de fin") {
    const d = parseDate(val);
    cell.value = d;
    if (d instanceof Date) cell.numFmt = FMT_DATE;
    return;
  }
  if (label === "Jour(s) travaillés") {
    cell.value = countJours(val);
    return;
  }

  const isFormula = val && typeof val === "object" && typeof val.formula === "string";
  if (isFormula) {
    const { text, crossRow } = translateFormula(val.formula, srcRow, targetRow, colmap);
    if (crossRow) {
      // La formule référence une ligne absente du nouveau format (ex. la
      // ligne "TOTAL" par salarié du fichier source, remplacée ici par le
      // seul sous-total de département) : on fige la dernière valeur
      // calculée plutôt que d'écrire une formule cassée.
      cell.value = val.cached;
      frozenStats.set(label, (frozenStats.get(label) || 0) + 1);
    } else {
      cell.value = { formula: text.slice(1) };
    }
  } else {
    cell.value = val === undefined ? null : val;
  }

  if (labelIsHours(label)) cell.numFmt = FMT_HOURS;
  else if (labelIsEuros(label)) cell.numFmt = FMT_EUROS;
}

export async function buildOutput(headers, sourceRows, options) {
  const { societe, production, objet, idcc } = options;

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("EXPORT BULLETIN", {
    views: [{ showGridLines: true, state: "frozen", ySplit: 4 }],
  });

  // Résolution du mapping colonnes source -> cible d'après les libellés
  // d'en-tête RÉELS du fichier déposé (peu importe leur position). Seules
  // les colonnes standard effectivement présentes dans ce fichier sont
  // retenues (dans l'ordre du format cible) — le tableau de sortie reflète
  // exactement le fichier déposé, ni plus ni moins.
  const { letterToTargetLabel, extraEntries } = resolveColumnMapping(headers);
  const usedTargetLabels = new Set(Object.values(letterToTargetLabel));
  const labelToSrcCol = {};
  for (const [srcCol, label] of Object.entries(letterToTargetLabel)) {
    labelToSrcCol[label] = srcCol;
  }
  const activeStandardLabels = STANDARD_HEADERS.filter((l) => usedTargetLabels.has(l));
  const activePayLabels = activeStandardLabels.filter((l) => !GROSS_TOTAL_LABELS.has(l));
  const activeTotalLabels = activeStandardLabels.filter((l) => GROSS_TOTAL_LABELS.has(l));
  const extrasMg = extraEntries.filter(([, label]) => isMgLabel(label));
  // "Total somme" n'est plus reprise dans l'export : "Salaire brut"
  // (calculée) la remplace. On l'exclut simplement des colonnes "extra".
  const extrasOther = extraEntries.filter(
    ([, label]) => !isMgLabel(label) && !isTotalSommeLabel(label)
  );
  // Colonnes "(NS)" en euros présentes dans le fichier déposé : condition la
  // présence de la colonne calculée "Total indemnité (NS)" en toute fin de
  // tableau (pas de colonne "au cas où" si le fichier n'en a aucune).
  const hasNsEuroCol = extraEntries.some(
    ([, label]) => isNsLabel(label) && labelIsEuros(label)
  );
  const NS_TOTAL_SRC = "__NS_TOTAL__"; // clé "source" synthétique (aucune colonne du fichier déposé)

  // Ordre final des colonnes : infos contrat + variables de paie standard
  // (dans l'ordre du format cible, "Jour(s) travaillés" y compris — précédée
  // d'une colonne dérivée listant les dates travaillées, cf. plus bas), puis
  // les colonnes sans équivalent standard (Cachet, Déf., Indem.…), puis les
  // anciennes colonnes de travail Garantie Minimale, puis les totaux bruts
  // (Salaire brut, Coût employeur, Salaire net imposable, Salaire net), puis
  // "Total indemnité (NS)" (calculée) en toute dernière colonne.
  const activePayEntries = [];
  for (const label of activePayLabels) {
    if (label === "Jour(s) travaillés" && labelToSrcCol[label]) {
      // Colonne dérivée : mêmes données source que "Jour(s) travaillés",
      // affichées en liste de dates plutôt qu'en nombre. `derive` la
      // distingue de la colonne canonique lors de l'écriture des lignes.
      activePayEntries.push({
        label: "Jours travaillés (dates)",
        srcCol: labelToSrcCol[label],
        derive: "dates",
      });
    }
    activePayEntries.push({ label, srcCol: labelToSrcCol[label] });
  }
  const columnPlan = [
    ...activePayEntries,
    ...extrasOther.map(([srcCol, label]) => ({ label, srcCol })),
    ...extrasMg.map(([srcCol, label]) => ({ label, srcCol })),
    ...activeTotalLabels.map((label) => ({ label, srcCol: labelToSrcCol[label] })),
    ...(hasNsEuroCol ? [{ label: NS_TOTAL_LABEL, srcCol: NS_TOTAL_SRC }] : []),
  ];
  const nTotalCols = columnPlan.length;

  // colonne cible (lettre) + libellé pour chaque colonne source reprise,
  // indexé par lettre SOURCE (propre à ce fichier), et zone visuelle de
  // chaque colonne cible (purement dérivée de columnPlan : aucune zone
  // vide ne peut apparaître). Les colonnes "dérivées" (ex. dates
  // travaillées) partagent leur colonne source avec une colonne canonique
  // déjà indexée dans `targetLetterOf` : elles sont donc collectées à part
  // dans `derivedEntries` plutôt que d'écraser cette entrée.
  const targetLetterOf = {};
  const targetLabelOf = {};
  const sectionOfIndex = [];
  const labelOfTargetLetter = {};
  const derivedEntries = [];
  columnPlan.forEach(({ label, srcCol, derive }, i) => {
    const idx = i + 1;
    const letter = colLetterFromIndex(idx);
    labelOfTargetLetter[letter] = label;
    sectionOfIndex[idx] = sectionForLabel(label);
    if (derive) {
      derivedEntries.push({ srcCol, tgtCol: letter, label, derive });
      return;
    }
    targetLetterOf[srcCol] = letter;
    targetLabelOf[srcCol] = label;
  });
  const fullColmap = { ...targetLetterOf };

  // Colonnes à sommer dans les sous-totaux (contrat / département / total
  // général) : toutes les colonnes, à l'exception des infos contrat (Nom,
  // Taux horaire, dates…) qui n'ont pas de sens à additionner.
  function shouldSumForSubtotal(label) {
    return sectionForLabel(label) !== "contrat";
  }
  function numFmtForLabel(label) {
    if (label === NS_TOTAL_LABEL) return FMT_EUROS;
    if (labelIsEuros(label)) return FMT_EUROS;
    if (labelIsHours(label)) return FMT_HOURS;
    return "0"; // Jour(s) travaillés
  }

  const sectionBoundaryCols = [];
  for (let i = 1; i < nTotalCols; i++) {
    if (sectionOfIndex[i] !== sectionOfIndex[i + 1]) sectionBoundaryCols.push(i);
  }
  const sectionOfLetter = {};
  for (let i = 1; i <= nTotalCols; i++) {
    sectionOfLetter[colLetterFromIndex(i)] = sectionOfIndex[i];
  }

  // Colonnes "(en €)" de la zone paie (variables de paie standard + extras
  // hors Garantie Minimale), à l'exclusion des colonnes "(NS)" (cf.
  // nsSumCols) : utilisées pour calculer "Salaire brut" par formule plutôt
  // que de recopier la valeur figée du fichier source.
  const brutSumCols = columnPlan
    .filter(({ label }) => sectionForLabel(label) === "paie" && labelIsEuros(label) && !isNsLabel(label))
    .map(({ srcCol }) => targetLetterOf[srcCol]);

  // Colonnes "(NS)" en euros de la zone paie : sommées séparément dans
  // "Total indemnité (NS)" plutôt que dans "Salaire brut".
  const nsSumCols = columnPlan
    .filter(({ label }) => sectionForLabel(label) === "paie" && labelIsEuros(label) && isNsLabel(label))
    .map(({ srcCol }) => targetLetterOf[srcCol]);

  // lettre source qui alimente un libellé cible donné (pour Métier, dates…)
  const srcLetterForTargetLabel = (targetLabel) =>
    Object.keys(letterToTargetLabel).find((l) => letterToTargetLabel[l] === targetLabel);
  const metierCol = srcLetterForTargetLabel("Métier");
  const debutCol = srcLetterForTargetLabel("Date de début");
  const finCol = srcLetterForTargetLabel("Date de fin");
  const contratCol = srcLetterForTargetLabel("Code contrat (itsi)");
  const matriculeCol = srcLetterForTargetLabel("Matricule salarié production (itsi)");

  // -- bandeau d'en-tête (lignes 1-3) --------------------------------
  ws.getCell("F1").value = "Société"; ws.getCell("F1").font = FONT_LABEL_BOLD;
  ws.getCell("G1").value = "Production"; ws.getCell("G1").font = FONT_LABEL_BOLD;
  ws.getCell("H1").value = "N° objet"; ws.getCell("H1").font = FONT_LABEL_BOLD;
  ws.getCell("J1").value = "IDCC"; ws.getCell("J1").font = FONT_LABEL_BOLD;

  ws.getCell("A2").value = "Période exportée";
  ws.getCell("F2").value = societe;
  ws.getCell("G2").value = production;
  ws.getCell("H2").value = objet;
  ws.getCell("J2").value = idcc;

  const debuts = debutCol
    ? sourceRows.map((r) => parseDate(r.data[debutCol])).filter((d) => d instanceof Date)
    : [];
  const fins = finCol
    ? sourceRows.map((r) => parseDate(r.data[finCol])).filter((d) => d instanceof Date)
    : [];
  if (debuts.length && fins.length) {
    const dmin = new Date(Math.min(...debuts));
    const dmax = new Date(Math.max(...fins));
    const fmt = (d) =>
      `${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}/${d.getUTCFullYear()}`;
    ws.getCell("A3").value = `du ${fmt(dmin)} au ${fmt(dmax)}`;
  } else {
    ws.getCell("A3").value = "du ... au ...";
  }

  // -- ligne d'entête des colonnes (ligne 4) --------------------------
  const headerRow = 4;
  columnPlan.forEach(({ label }, i) => {
    const idx = i + 1;
    const cell = ws.getCell(headerRow, idx);
    cell.value = label;
    cell.font = FONT_LABEL_BOLD;
    cell.alignment = { ...CENTER, wrapText: true };
    cell.fill = SECTION_HEADER_FILL[sectionOfIndex[idx]];
    cell.border = { bottom: HEADER_BOTTOM_BORDER };
    ws.getColumn(idx).width = widthForLabel(label);
  });
  ws.getRow(headerRow).height = HEADER_ROW_HEIGHT;

  // -- regroupement des salariés par département -----------------------
  // Trié par matricule avant regroupement : le préfixe CNC du matricule
  // fait que cet ordre est aussi un ordre de département, sans avoir à
  // maintenir une liste de département arbitraire ici (l'ordre de sortie
  // suit l'ordre de première apparition, comme ProjectJobController::index
  // côté production : tri par référence puis regroupement).
  const byDept = {};
  const unclassified = new Set();
  const sortedForGrouping = matriculeCol
    ? [...sourceRows].sort((a, b) => compareMatricules(a.data[matriculeCol], b.data[matriculeCol]))
    : sourceRows;
  for (const { srcRow, data } of sortedForGrouping) {
    const metier = metierCol ? data[metierCol] : undefined;
    const { dept, recognized } = classifyMetier(metier);
    if (!recognized) unclassified.add(metier);
    if (!byDept[dept]) byDept[dept] = [];
    byDept[dept].push({ srcRow, data });
  }

  const warnings = [];
  const frozenStats = new Map(); // libellé -> nb de formules figées (cf. writeValue)
  let currentRow = headerRow + 2; // ligne 5 vide comme dans le modèle
  const subtotalRefs = {}; // tgtCol -> [ligne sous-total département, ...]

  for (const dept of Object.keys(byDept)) {
    const entries = byDept[dept];
    if (!entries || entries.length === 0) continue;

    const deptHeaderRow = currentRow;
    const deptCell = ws.getCell(deptHeaderRow, 1);
    deptCell.value = dept;
    for (let c = 1; c <= nTotalCols; c++) {
      ws.getCell(deptHeaderRow, c).fill = FILL_DEPT;
    }
    deptCell.font = FONT_DEPT;
    currentRow += 1;

    // Regroupe les lignes consécutives partageant le même code contrat
    // (un salarié peut avoir plusieurs bulletins sur la période pour un
    // même contrat) pour y ajouter un sous-total contrat.
    const contractGroups = [];
    for (const entry of entries) {
      const code = contratCol ? entry.data[contratCol] : undefined;
      const hasCode = code !== undefined && code !== null && code !== "";
      const last = contractGroups[contractGroups.length - 1];
      if (last && hasCode && last.code === code) {
        last.rows.push(entry);
      } else {
        contractGroups.push({ code: hasCode ? code : null, rows: [entry] });
      }
    }

    // tgtCol -> lignes à sommer pour le sous-total département (la ligne de
    // sous-total contrat dès qu'un code contrat existe, sinon la ligne du
    // bulletin lui-même — jamais les deux, pour ne rien compter en double).
    const deptRefRows = {};

    for (const group of contractGroups) {
      const groupFirstRow = currentRow;
      const colsWithAmounts = new Set();

      for (const { srcRow, data } of group.rows) {
        const targetRow = currentRow;
        for (const [srcCol, tgtCol] of Object.entries(targetLetterOf)) {
          const val = data[srcCol];
          const label = targetLabelOf[srcCol];
          const tgtIdx = colIndexFromLetter(tgtCol);
          const outCell = ws.getCell(targetRow, tgtIdx);
          if (label === "Salaire brut (en €)" && brutSumCols.length) {
            // Calculé (somme des colonnes "paie" en €, hors "(NS)"), pas
            // recopié depuis la valeur figée du fichier source : la formule
            // reste vivante si l'utilisateur modifie des heures.
            outCell.value = {
              formula: `SUM(${brutSumCols.map((c) => `${c}${targetRow}`).join(",")})`,
            };
            outCell.numFmt = FMT_EUROS;
          } else if (label === NS_TOTAL_LABEL && nsSumCols.length) {
            // Colonne calculée : somme des colonnes "(NS)" en euros, sans
            // équivalent dans le fichier source (pas de srcCol réel).
            outCell.value = {
              formula: `SUM(${nsSumCols.map((c) => `${c}${targetRow}`).join(",")})`,
            };
            outCell.numFmt = FMT_EUROS;
          } else {
            writeValue(outCell, val, label, srcCol, tgtCol, srcRow, targetRow, fullColmap, frozenStats);
          }
          outCell.fill = SECTION_DATA_FILL[sectionOfLetter[tgtCol]];
          const hasVal =
            (label === "Salaire brut (en €)" && brutSumCols.length) ||
            (label === NS_TOTAL_LABEL && nsSumCols.length) ||
            (val !== null && val !== undefined && val !== "");
          if (shouldSumForSubtotal(label) && hasVal) colsWithAmounts.add(tgtCol);
        }
        for (const { srcCol, tgtCol, derive } of derivedEntries) {
          const outCell = ws.getCell(targetRow, colIndexFromLetter(tgtCol));
          if (derive === "dates") outCell.value = formatWorkedDates(data[srcCol]);
          outCell.fill = SECTION_DATA_FILL[sectionOfLetter[tgtCol]];
        }
        currentRow += 1;
      }

      const groupLastRow = currentRow - 1;
      const sortedGroupCols = [...colsWithAmounts].sort(
        (a, b) => colIndexFromLetter(a) - colIndexFromLetter(b)
      );

      if (group.code) {
        const contractSubtotalRow = currentRow;
        ws.getCell(contractSubtotalRow, 1).value = `SOUS-TOTAL ${group.code}`;
        for (let c = 1; c <= nTotalCols; c++) {
          ws.getCell(contractSubtotalRow, c).fill = FILL_CONTRACT;
        }
        ws.getCell(contractSubtotalRow, 1).font = FONT_CONTRACT;

        for (const tgtCol of sortedGroupCols) {
          const idx = colIndexFromLetter(tgtCol);
          const c = ws.getCell(contractSubtotalRow, idx);
          c.value = { formula: `SUM(${tgtCol}${groupFirstRow}:${tgtCol}${groupLastRow})` };
          c.fill = FILL_CONTRACT;
          c.numFmt = numFmtForLabel(labelOfTargetLetter[tgtCol]);
          (deptRefRows[tgtCol] ||= []).push(contractSubtotalRow);
        }
        currentRow += 1;
      } else {
        for (const tgtCol of sortedGroupCols) {
          (deptRefRows[tgtCol] ||= []).push(groupFirstRow);
        }
      }
    }

    const deptSubtotalRow = currentRow;
    ws.getCell(deptSubtotalRow, 1).value = `SOUS-TOTAL ${dept}`;
    for (let c = 1; c <= nTotalCols; c++) {
      ws.getCell(deptSubtotalRow, c).fill = FILL_DEPT;
    }
    ws.getCell(deptSubtotalRow, 1).font = FONT_DEPT;

    const sortedDeptCols = Object.keys(deptRefRows).sort(
      (a, b) => colIndexFromLetter(a) - colIndexFromLetter(b)
    );
    for (const tgtCol of sortedDeptCols) {
      const idx = colIndexFromLetter(tgtCol);
      const refs = deptRefRows[tgtCol].map((r) => `${tgtCol}${r}`).join(",");
      const c = ws.getCell(deptSubtotalRow, idx);
      c.value = { formula: `SUM(${refs})` };
      c.fill = FILL_DEPT;
      c.numFmt = numFmtForLabel(labelOfTargetLetter[tgtCol]);
      (subtotalRefs[tgtCol] ||= []).push(deptSubtotalRow);
    }

    currentRow = deptSubtotalRow + 2; // ligne vide de séparation
  }

  for (const [label, count] of frozenStats) {
    warnings.push(
      `${label} : ${count} valeur(s) figée(s) au lieu d'une formule vivante ` +
      `(référence à une ligne "TOTAL" par salarié absente du nouveau format).`
    );
  }

  // -- ligne TOTAL GÉNÉRAL ---------------------------------------------
  const totalRow = currentRow;
  const totalLabelCell = ws.getCell(totalRow, 1);
  totalLabelCell.value = "TOTAL GÉNÉRAL";
  totalLabelCell.font = FONT_LABEL_BOLD;
  for (const [tgtCol, rowsArr] of Object.entries(subtotalRefs)) {
    const idx = colIndexFromLetter(tgtCol);
    const refs = rowsArr.map((r) => `${tgtCol}${r}`).join(",");
    const c = ws.getCell(totalRow, idx);
    c.value = { formula: `SUM(${refs})` };
    c.font = FONT_LABEL_BOLD;
    c.numFmt = numFmtForLabel(labelOfTargetLetter[tgtCol]);
  }

  // -- bordures verticales entre zones de colonnes (contrat / paie / totaux)
  for (const boundaryCol of sectionBoundaryCols) {
    for (let r = headerRow; r <= totalRow; r++) {
      const cell = ws.getCell(r, boundaryCol);
      cell.border = { ...cell.border, right: SECTION_BORDER };
    }
  }

  const buffer = await wb.xlsx.writeBuffer();
  return { buffer, warnings, unclassified: [...unclassified].filter(Boolean) };
}

export async function generate(sourceArrayBuffer, options) {
  const { headers, rows } = await readSource(sourceArrayBuffer);
  if (rows.length === 0) {
    throw new Error("Aucune ligne de bulletin trouvée dans le fichier source.");
  }
  return buildOutput(headers, rows, options);
}
