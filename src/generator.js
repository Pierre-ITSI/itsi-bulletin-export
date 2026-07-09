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
  "Prime except. (en h)", "Prime except. (en €)", "Coût employeur (en h)",
  "Coût employeur (en €)", "Salaire brut (en h)", "Salaire brut (en €)",
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

const METIER_TO_DEPT = {
  "directeur de production cinema": "PRODUCTION",
  "administrateur de production cinema": "PRODUCTION",
  "administrateur adjoint comptable cinema": "PRODUCTION",
  "assistant comptable de production cinema": "PRODUCTION",
  "secretaire de production cinema": "PRODUCTION",
  "assistant de production adjoint": "PRODUCTION",
  "responsable des enfants cinema": "PRODUCTION",

  "regisseur general cinema": "REGIE",
  "regisseur adjoint cinema": "REGIE",
  "regisseur d'exterieur cinema": "REGIE",
  "auxiliaire a la regie cinema": "REGIE",

  "premier assistant realisateur cinema": "REALISATION",
  "second assistant realisateur cinema": "REALISATION",
  "auxiliaire a la realisation cinema": "REALISATION",
  "scripte cinema": "REALISATION",
  "assistant scripte cinema": "REALISATION",
  "assistant au charge de la figuration cinema": "REALISATION",
  "charge de la figuration cinema": "REALISATION",
  "choregraphe": "REALISATION",

  "directeur de la photographie cinema": "IMAGE",
  "cadreur cinema": "IMAGE",
  "premier assistant operateur cinema": "IMAGE",
  "deuxieme assistant operateur cinema": "IMAGE",
  "technicien retour image cinema": "IMAGE",

  "chef machiniste prise de vues cinema": "MACHINERIE",
  "sous-chef machiniste de prise de vues cinema": "MACHINERIE",
  "machiniste de prise de vues cinema": "MACHINERIE",

  "chef electricien prise de vues cinema": "ELECTRICITE",
  "sous-chef electricien prise de vues cinema": "ELECTRICITE",
  "electricien prise de vues cinema": "ELECTRICITE",

  "chef operateur de son cinema": "SON",
  "1er assistant operateur du son cinema": "SON",
  "2eme assistant operateur du son cinema": "SON",

  "chef decorateur cinema": "DECORATION",
  "troisieme assistant decorateur cinema": "DECORATION",
  "ensemblier cinema": "DECORATION",
  "accessoiriste de plateau cinema": "DECORATION",
  "chef peintre de decor cinema": "DECORATION",
  "peintre faux bois et patine decor cinema": "DECORATION",
  "menuisier traceur de decor cinema": "DECORATION",
  "machiniste de construction cinema": "DECORATION",
  "infographiste de decor cinema": "DECORATION",

  "chef costumier cinema": "COSTUMES",
  "1er assistant costume cinema": "COSTUMES",
  "costumier cinema": "COSTUMES",
  "habilleur cinema": "COSTUMES",
  "auxiliaire a la regie cinema (affecte costumes)": "COSTUMES",

  "chef maquilleur cinema": "MAQUILLAGE",
  "maquilleur cinema": "MAQUILLAGE",

  "chef coiffeur cinema": "COIFFURE",
  "coiffeur cinema": "COIFFURE",

  "coordinateur de post-production cinema": "MONTAGE",
  "chef monteur cinema": "MONTAGE",
  "1er assistant monteur cinema": "MONTAGE",
};

function stripAccents(s) {
  return s.normalize("NFKD").replace(/[̀-ͯ]/g, "");
}

function classifyMetier(metier) {
  if (!metier) return "AUTRES";
  const key = stripAccents(String(metier).trim().toLowerCase());
  return METIER_TO_DEPT[key] || "AUTRES";
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
// tableau — les totaux bruts (Coût employeur, Salaire brut, Salaire net
// imposable, Salaire net). L'ORDRE des colonnes suit ce même regroupement
// (cf. buildColumnLayout), pas seulement leur couleur.
const IDENTITY_LABELS = new Set([
  "Code bulletin (itsi)", "Statut (itsi)", "Code contrat (itsi)",
  "Code salarié projet", "Matricule salarié production (itsi)", "Nom",
  "Prénom", "Abattement", "Métier", "Date de début", "Date de fin",
  "Taux horaire",
]);
const GROSS_TOTAL_LABELS = new Set([
  "Coût employeur (en h)", "Coût employeur (en €)",
  "Salaire brut (en h)", "Salaire brut (en €)",
  "Salaire net imposable (en h)", "Salaire net imposable (en €)",
  "Salaire net (en h)", "Salaire net (en €)",
]);
// Anciennes colonnes de travail "Garantie Minimale" du fichier source
// (Total base, MG, Ratio MG, Supp ap. MG) : reconnues par motif car elles
// n'ont pas de libellé cible standard (ce sont toujours des colonnes
// "extra"). "Total somme" en fait aussi partie à l'origine, mais rejoint la
// zone violette (cf. isTotalSommeLabel) : elle est placée juste avant les
// totaux bruts plutôt qu'avec le reste des colonnes de travail.
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
  const extrasTotalSomme = extraEntries.filter(([, label]) => isTotalSommeLabel(label));
  const extrasOther = extraEntries.filter(
    ([, label]) => !isMgLabel(label) && !isTotalSommeLabel(label)
  );

  // Ordre final des colonnes : infos contrat + variables de paie standard
  // (dans l'ordre du format cible, "Jour(s) travaillés" y compris), puis
  // les colonnes sans équivalent standard (Cachet, Déf., Indem.…), puis les
  // anciennes colonnes de travail Garantie Minimale, puis "Total somme"
  // juste avant les totaux bruts (Coût employeur, Salaire brut, Salaire net
  // imposable, Salaire net), tout à la fin du tableau.
  const columnPlan = [
    ...activePayLabels.map((label) => ({ label, srcCol: labelToSrcCol[label] })),
    ...extrasOther.map(([srcCol, label]) => ({ label, srcCol })),
    ...extrasMg.map(([srcCol, label]) => ({ label, srcCol })),
    ...extrasTotalSomme.map(([srcCol, label]) => ({ label, srcCol })),
    ...activeTotalLabels.map((label) => ({ label, srcCol: labelToSrcCol[label] })),
  ];
  const nTotalCols = columnPlan.length;

  // colonne cible (lettre) + libellé pour chaque colonne source reprise,
  // indexé par lettre SOURCE (propre à ce fichier), et zone visuelle de
  // chaque colonne cible (purement dérivée de columnPlan : aucune zone
  // vide ne peut apparaître).
  const targetLetterOf = {};
  const targetLabelOf = {};
  const sectionOfIndex = [];
  const labelOfTargetLetter = {};
  columnPlan.forEach(({ label, srcCol }, i) => {
    const idx = i + 1;
    const letter = colLetterFromIndex(idx);
    targetLetterOf[srcCol] = letter;
    targetLabelOf[srcCol] = label;
    sectionOfIndex[idx] = sectionForLabel(label);
    labelOfTargetLetter[letter] = label;
  });
  const fullColmap = { ...targetLetterOf };

  // Colonnes à sommer dans les sous-totaux (contrat / département / total
  // général) : toutes les colonnes, à l'exception des infos contrat (Nom,
  // Taux horaire, dates…) qui n'ont pas de sens à additionner.
  function shouldSumForSubtotal(label) {
    return sectionForLabel(label) !== "contrat";
  }
  function numFmtForLabel(label) {
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
  // hors Garantie Minimale) : utilisées pour calculer "Salaire brut" par
  // formule plutôt que de recopier la valeur figée du fichier source.
  const brutSumCols = columnPlan
    .filter(({ label }) => sectionForLabel(label) === "paie" && labelIsEuros(label))
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
    const dept = classifyMetier(metier);
    if (dept === "AUTRES") unclassified.add(metier);
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
            // Calculé (somme des colonnes "paie" en €), pas recopié
            // depuis la valeur figée du fichier source : la formule reste
            // vivante si l'utilisateur modifie des heures.
            outCell.value = {
              formula: `SUM(${brutSumCols.map((c) => `${c}${targetRow}`).join(",")})`,
            };
            outCell.numFmt = FMT_EUROS;
          } else {
            writeValue(outCell, val, label, srcCol, tgtCol, srcRow, targetRow, fullColmap, frozenStats);
          }
          outCell.fill = SECTION_DATA_FILL[sectionOfLetter[tgtCol]];
          const hasVal =
            (label === "Salaire brut (en €)" && brutSumCols.length) ||
            (val !== null && val !== undefined && val !== "");
          if (shouldSumForSubtotal(label) && hasVal) colsWithAmounts.add(tgtCol);
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
