// Génère un fichier "export bulletin" (format ITSI / type TOURNAGE) à partir
// d'un fichier "Combine" (type Marie-Christine), en conservant les formules
// de calcul (taux horaire x heures, etc.) pour que le fichier reste un outil
// de travail modifiable.
//
// Portage JS (exceljs) du script generate_bulletin_export.py.

import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

// --------------------------------------------------------------------------
// 1. Correspondance des colonnes source -> colonnes cible
// --------------------------------------------------------------------------

// Colonnes "identité" (pas de calcul)
const IDENTITY_MAP = {
  A: "A", // Code bulletin
  B: "B", // Statut
  G: "C", // Code contrat
  E: "D", // Code employé -> Code salarié projet
  H: "E", // Matricule
  C: "F", // Nom
  D: "G", // Prénom
  F: "H", // Abattement
  I: "I", // Métier
  J: "J", // Date de début
  K: "K", // Date de fin
  L: "L", // Taux horaire
  M: "M", // Jour(s) travaillés
};

// Colonnes de calcul (paires heures/euros) qui existent dans les deux formats
const CALC_MAP = {
  N: "N", O: "O", // H. normales
  P: "P", Q: "Q", // H. supp. 125%
  R: "R", S: "S", // H. supp. 150%
  T: "T", U: "U", // H. supp. 175%
  V: "X", W: "Y", // Majo. jour 25%
  X: "Z", Y: "AA", // Majo. jour 50%
  Z: "AB", AA: "AC", // Majo. jour 100%
  AB: "AH", AC: "AI", // Majo. nuit 50%
  AD: "AL", AE: "AM", // H. anticipées 100%
  AF: "BJ", AG: "BK", // Indem. transport (Tournage)
  AH: "BP", AI: "BQ", // Indem. voyage (Tournage)
  AJ: "BV", AK: "BW", // Indem. repas (Hors RP)
  AL: "CF", AM: "CG", // Indem. casse croûte (Hors RP)
  AN: "CL", AO: "CM", // Indem. continue
  AP: "CX", AQ: "CY", // Retrait plafond majo.
  AR: "CZ", AS: "DA", // Prime except.
  BK: "DC", // Coût employeur (€)
  BL: "DE", // Salaire brut (€)
  BM: "DG", // Salaire net imposable (€)
  BN: "DI", // Salaire net (€)
};

// Colonnes source sans équivalent dans le format cible : ajoutées en fin de
// tableau (après DI) pour ne perdre aucune information.
const EXTRA_COLS = [
  ["AT", "Déf. soumis (en h)"],
  ["AU", "Déf. soumis (en €)"],
  ["AV", "Déf. non soumis (en h)"],
  ["AW", "Déf. non soumis (en €)"],
  ["AX", "Rep. continue (en h)"],
  ["AY", "Rep. continue (en €)"],
  ["AZ", "Indem. Matériel (S) (en h)"],
  ["BA", "Indem. Matériel (S) (en €)"],
  ["BB", "Indem. MàL (S) (en h)"],
  ["BC", "Indem. MàL (S) (en €)"],
];

// Colonnes "de travail" du fichier source, sans équivalent, non reprises
// (Total base, MG, Ratio MG, Supp ap. MG x3, Total somme) : elles servaient
// uniquement à répartir la Garantie Minimale entre les bulletins d'un même
// salarié, une notion qui n'existe pas dans le format cible.
// (BD, BE, BF, BG, BH, BI, BJ dans le fichier source)

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

const COL_WIDTHS = { A: 18, B: 7.88, C: 16.38, D: 11.0, E: 9.63, H: 12.38, I: 25.38 };

// --------------------------------------------------------------------------
// 2. Classement des métiers par département de tournage
// --------------------------------------------------------------------------

const METIER_TO_DEPT = {
  "directeur de production cinema": "PRODUCTION",
  "administrateur de production cinema": "PRODUCTION",
  "administrateur adjoint comptable cinema": "PRODUCTION",
  "secretaire de production cinema": "PRODUCTION",
  "assistant de production adjoint": "PRODUCTION",

  "regisseur general cinema": "REGIE",
  "regisseur adjoint cinema": "REGIE",
  "auxiliaire a la regie cinema": "REGIE",

  "premier assistant realisateur cinema": "REALISATION",
  "second assistant realisateur cinema": "REALISATION",
  "auxiliaire a la realisation cinema": "REALISATION",
  "scripte cinema": "REALISATION",
  "assistant scripte cinema": "REALISATION",
  "assistant au charge de la figuration cinema": "REALISATION",

  "directeur de la photographie cinema": "IMAGE",
  "premier assistant operateur cinema": "IMAGE",
  "deuxieme assistant operateur cinema": "IMAGE",
  "technicien retour image cinema": "IMAGE",

  "chef machiniste prise de vues cinema": "MACHINERIE",
  "machiniste de prise de vues cinema": "MACHINERIE",

  "chef electricien prise de vues cinema": "ELECTRICITE",
  "sous-chef electricien prise de vues cinema": "ELECTRICITE",
  "electricien prise de vues cinema": "ELECTRICITE",

  "chef operateur de son cinema": "SON",
  "1er assistant operateur du son cinema": "SON",

  "troisieme assistant decorateur cinema": "DECORATION",
  "accessoiriste de plateau cinema": "DECORATION",

  "chef costumier cinema": "COSTUMES",
  "1er assistant costume cinema": "COSTUMES",
  "costumier cinema": "COSTUMES",
  "habilleur cinema": "COSTUMES",

  "chef maquilleur cinema": "MAQUILLAGE",
  "maquilleur cinema": "MAQUILLAGE",

  "chef coiffeur cinema": "COIFFURE",
  "coiffeur cinema": "COIFFURE",
};

const DEPARTMENT_ORDER = [
  "REGIE", "PRODUCTION", "REALISATION", "IMAGE", "MACHINERIE",
  "ELECTRICITE", "SON", "DECORATION", "COSTUMES", "MAQUILLAGE",
  "COIFFURE", "AUTRES",
];

function stripAccents(s) {
  return s.normalize("NFKD").replace(/[̀-ͯ]/g, "");
}

function classifyMetier(metier) {
  if (!metier) return "AUTRES";
  const key = stripAccents(String(metier).trim().toLowerCase());
  return METIER_TO_DEPT[key] || "AUTRES";
}

// --------------------------------------------------------------------------
// 3. Traduction des formules (mêmes lignes uniquement : taux x heures)
// --------------------------------------------------------------------------

const CELL_RE = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g;

function translateFormula(formula, srcRow, newRow, colmap, warnings, coordForWarning) {
  return formula.replace(CELL_RE, (match, d1, col, d2, row) => {
    const rowNum = parseInt(row, 10);
    if (rowNum !== srcRow) {
      warnings.push(
        `${coordForWarning}: référence à une autre ligne (${match}) non traduite, à vérifier.`
      );
      return match;
    }
    const newCol = colmap[col.toUpperCase()] || col.toUpperCase();
    return `${d1}${newCol}${d2}${newRow}`;
  });
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
  if (cell.f !== undefined) return { formula: `=${cell.f}` };
  if (cell.v === undefined) return null;
  return cell.v;
}

export async function readSource(arrayBuffer) {
  const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: "array", cellFormula: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const range = XLSX.utils.decode_range(ws["!ref"]);

  const rows = [];
  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    // r=0 est la ligne d'en-tête ; rowNumber reste 1-indexé pour rester
    // cohérent avec les coordonnées Excel utilisées ailleurs.
    const rowNumber = r + 1;
    const aCell = ws[XLSX.utils.encode_cell({ r, c: 0 })];
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
  return rows;
}

// --------------------------------------------------------------------------
// 5. Construction du classeur de sortie
// --------------------------------------------------------------------------

const FILL_DEPT = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCE5CD" } };
const FONT_LABEL_BOLD = { name: "Arial", bold: true };
const FONT_DEPT = { name: "Arial", size: 8, italic: true };
const FMT_HOURS = "#,##0.00";
const FMT_EUROS = '#,##0.00"€"';
const FMT_DATE = "dd/mm/yyyy";
const CENTER = { horizontal: "center", vertical: "middle" };

function labelIsHours(label) {
  return label.endsWith("(en h)");
}
function labelIsEuros(label) {
  return label.endsWith("(en €)");
}

function writeValue(cell, val, label, srcCol, tgtCol, srcRow, targetRow, colmap, warnings) {
  if (srcCol === "J" || srcCol === "K") {
    const d = parseDate(val);
    cell.value = d;
    if (d instanceof Date) cell.numFmt = FMT_DATE;
    return;
  }
  if (srcCol === "M") {
    cell.value = countJours(val);
    return;
  }

  const isFormula = val && typeof val === "object" && typeof val.formula === "string";
  if (isFormula) {
    const translated = translateFormula(
      val.formula, srcRow, targetRow, colmap, warnings,
      `${tgtCol}${targetRow} (source ${srcCol}${srcRow})`
    );
    cell.value = { formula: translated.slice(1) };
  } else if (val && typeof val === "object" && val.formula !== undefined) {
    // formule non résolue proprement, on ignore
    cell.value = null;
  } else {
    cell.value = val === undefined ? null : val;
  }

  if (labelIsHours(label)) cell.numFmt = FMT_HOURS;
  else if (labelIsEuros(label)) cell.numFmt = FMT_EUROS;
}

export async function buildOutput(sourceRows, options) {
  const { societe, production, objet, idcc } = options;

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("EXPORT BULLETIN", { views: [{ showGridLines: true }] });

  const nStd = STANDARD_HEADERS.length; // colonnes A..DI
  const nExtra = EXTRA_COLS.length;
  const nTotalCols = nStd + nExtra;

  // colonne cible (lettre) + libellé pour chaque colonne source reprise
  const targetLetterOf = {};
  const targetLabelOf = {};
  for (const [srcCol, tgtCol] of Object.entries({ ...IDENTITY_MAP, ...CALC_MAP })) {
    targetLetterOf[srcCol] = tgtCol;
    targetLabelOf[srcCol] = STANDARD_HEADERS[colIndexFromLetter(tgtCol) - 1];
  }
  EXTRA_COLS.forEach(([srcCol, label], i) => {
    const letter = colLetterFromIndex(nStd + 1 + i);
    targetLetterOf[srcCol] = letter;
    targetLabelOf[srcCol] = label;
  });
  const fullColmap = { ...targetLetterOf };

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

  const debuts = sourceRows.map((r) => parseDate(r.data.J)).filter((d) => d instanceof Date);
  const fins = sourceRows.map((r) => parseDate(r.data.K)).filter((d) => d instanceof Date);
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
  STANDARD_HEADERS.forEach((label, i) => {
    const cell = ws.getCell(headerRow, i + 1);
    cell.value = label;
    cell.font = FONT_LABEL_BOLD;
    cell.alignment = CENTER;
  });
  EXTRA_COLS.forEach(([, label], i) => {
    const cell = ws.getCell(headerRow, nStd + 1 + i);
    cell.value = label;
    cell.font = FONT_LABEL_BOLD;
    cell.alignment = CENTER;
  });

  for (const [col, width] of Object.entries(COL_WIDTHS)) {
    ws.getColumn(col).width = width;
  }

  // -- regroupement des salariés par département -----------------------
  const byDept = {};
  const unclassified = new Set();
  for (const { srcRow, data } of sourceRows) {
    const dept = classifyMetier(data.I);
    if (dept === "AUTRES") unclassified.add(data.I);
    if (!byDept[dept]) byDept[dept] = [];
    byDept[dept].push({ srcRow, data });
  }

  const warnings = [];
  let currentRow = headerRow + 2; // ligne 5 vide comme dans le modèle
  const subtotalRefs = {}; // tgtCol -> [subtotalRow, ...]

  for (const dept of DEPARTMENT_ORDER) {
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

    const firstDataRow = currentRow;
    const colsWithAmounts = new Set();

    for (const { srcRow, data } of entries) {
      const targetRow = currentRow;
      for (const [srcCol, tgtCol] of Object.entries(targetLetterOf)) {
        const val = data[srcCol];
        const label = targetLabelOf[srcCol];
        const tgtIdx = colIndexFromLetter(tgtCol);
        const outCell = ws.getCell(targetRow, tgtIdx);
        writeValue(outCell, val, label, srcCol, tgtCol, srcRow, targetRow, fullColmap, warnings);
        const hasVal = val !== null && val !== undefined && val !== "";
        if (label.endsWith("(en €)") && hasVal) colsWithAmounts.add(tgtCol);
      }
      currentRow += 1;
    }

    const lastDataRow = currentRow - 1;

    const subtotalRow = currentRow;
    ws.getCell(subtotalRow, 1).value = `SOUS-TOTAL ${dept}`;
    for (let c = 1; c <= nTotalCols; c++) {
      ws.getCell(subtotalRow, c).fill = FILL_DEPT;
    }
    ws.getCell(subtotalRow, 1).font = FONT_DEPT;

    const sortedCols = [...colsWithAmounts].sort((a, b) => colIndexFromLetter(a) - colIndexFromLetter(b));
    for (const tgtCol of sortedCols) {
      const idx = colIndexFromLetter(tgtCol);
      const c = ws.getCell(subtotalRow, idx);
      c.value = { formula: `SUM(${tgtCol}${firstDataRow}:${tgtCol}${lastDataRow})` };
      c.fill = FILL_DEPT;
      c.numFmt = FMT_EUROS;
      if (!subtotalRefs[tgtCol]) subtotalRefs[tgtCol] = [];
      subtotalRefs[tgtCol].push(subtotalRow);
    }

    currentRow = subtotalRow + 2; // ligne vide de séparation
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
    c.numFmt = FMT_EUROS;
  }

  const buffer = await wb.xlsx.writeBuffer();
  return { buffer, warnings, unclassified: [...unclassified].filter(Boolean) };
}

export async function generate(sourceArrayBuffer, options) {
  const rows = await readSource(sourceArrayBuffer);
  if (rows.length === 0) {
    throw new Error("Aucune ligne de bulletin trouvée dans le fichier source.");
  }
  return buildOutput(rows, options);
}
