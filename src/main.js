import "./style.css";
import { generate } from "./generator.js";

const app = document.querySelector("#app");

app.innerHTML = `
  <main class="wrap">
    <h1>Générateur export bulletin ITSI</h1>
    <p class="subtitle">
      Déposer un fichier d'export de Bulletins excel extrait d'itsi-production,
      pour obtenir un fichier consolidé, regroupé par département et par
      contrat, avec les formules conservées.
    </p>

    <div id="dropzone" class="dropzone">
      <p class="dz-text">Glissez-déposez le fichier .xlsx ici<br />ou</p>
      <label class="file-btn">
        Choisir un fichier
        <input type="file" id="file-input" accept=".xlsx" hidden />
      </label>
    </div>

    <fieldset class="options">
      <legend>Informations d'en-tête (modifiables)</legend>
      <label>Société <input id="opt-societe" type="text" value="ITSI-APP" /></label>
      <label>Production <input id="opt-production" type="text" value="[Nom du film]" /></label>
      <label>N° objet <input id="opt-objet" type="text" value="[N° objet]" /></label>
      <label>IDCC <input id="opt-idcc" type="number" value="3097" /></label>
    </fieldset>

    <div id="status" class="status" hidden></div>
    <div id="result" class="result" hidden></div>
  </main>
`;

const dropzone = document.querySelector("#dropzone");
const fileInput = document.querySelector("#file-input");
const statusEl = document.querySelector("#status");
const resultEl = document.querySelector("#result");

function setStatus(message, kind = "info") {
  statusEl.hidden = false;
  statusEl.textContent = message;
  statusEl.className = `status status--${kind}`;
}

function clearStatus() {
  statusEl.hidden = true;
  statusEl.textContent = "";
}

["dragenter", "dragover"].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.add("dropzone--active");
  });
});
["dragleave", "drop"].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.remove("dropzone--active");
  });
});
dropzone.addEventListener("drop", (e) => {
  const file = e.dataTransfer.files?.[0];
  if (file) handleFile(file);
});
fileInput.addEventListener("change", (e) => {
  const file = e.target.files?.[0];
  if (file) handleFile(file);
});

async function handleFile(file) {
  resultEl.hidden = true;
  resultEl.innerHTML = "";
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    setStatus("Merci de déposer un fichier .xlsx", "error");
    return;
  }

  setStatus(`Traitement de « ${file.name} »…`, "info");

  const options = {
    societe: document.querySelector("#opt-societe").value || "",
    production: document.querySelector("#opt-production").value || "",
    objet: document.querySelector("#opt-objet").value || "",
    idcc: Number(document.querySelector("#opt-idcc").value) || "",
  };

  try {
    const arrayBuffer = await file.arrayBuffer();
    const { buffer, warnings, unclassified } = await generate(arrayBuffer, options);

    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const outName = file.name.replace(/\.xlsx$/i, "") + "_export_bulletin.xlsx";

    resultEl.hidden = false;
    resultEl.innerHTML = `
      <a class="download-btn" href="${url}" download="${outName}">
        Télécharger ${outName}
      </a>
    `;

    if (unclassified.length || warnings.length) {
      const notes = [];
      if (unclassified.length) {
        notes.push(
          `Métiers non reconnus (classés en "AUTRES") : ${unclassified.join(", ")}`
        );
      }
      if (warnings.length) {
        notes.push(...warnings);
      }
      const notesEl = document.createElement("div");
      notesEl.className = "notes";
      notesEl.innerHTML = notes.map((n) => `<p>⚠️ ${n}</p>`).join("");
      resultEl.appendChild(notesEl);
    }

    setStatus("Fichier généré avec succès.", "success");
  } catch (err) {
    console.error(err);
    setStatus(`Erreur : ${err.message}`, "error");
  }
}
