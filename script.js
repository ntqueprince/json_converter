const fileInput = document.getElementById("fileInput");
const outputJS = document.getElementById("outputJS");
const preview = document.getElementById("preview");
const outputValues = document.getElementById("outputValues");
const errorLog = document.getElementById("errorLog");
const summary = document.getElementById("summary");
const statusEl = document.getElementById("status");
const bar = document.getElementById("bar");

const btnConvert = document.getElementById("btnConvert");
const btnCopy = document.getElementById("btnCopy");
const btnDownload = document.getElementById("btnDownload");
const btnClear = document.getElementById("btnClear");

let lastJS = "";

const FINAL_HEADERS = [
  "InsurerRequirement",
  "Insurer",
  "Requirement",
  "Endorsement type",
  "Documents or any other requirement",
  "TAT",
  "Charges / Deduction",
  "Inspection",
  "Any Exception",
  "Declaration format (if declaration required)"
];

const HEADER_ALIASES = {
  "InsurerRequirement": ["InsurerRequirement", "Insurer Requirement", "Insurer+Requirement", "Key", "Unique Key"],
  "Insurer": ["Insurer", "Company", "Insurance Company", "Insurer Name"],
  "Requirement": ["Requirement", "Correction", "Endorsement", "Endorsement Requirement", "Type"],
  "Endorsement type": ["Endorsement type", "Endorsement Type", "Endt Type", "Endt", "Category"],
  "Documents or any other requirement": ["Documents or any other requirement", "Documents", "Docs", "Document", "Document Required", "Requirements"],
  "TAT": ["TAT", "Tat", "Turnaround", "Turn Around Time"],
  "Charges / Deduction": ["Charges / Deduction", "Charges", "Deduction", "Charges/Deduction"],
  "Inspection": ["Inspection", "Survey", "Vehicle Inspection"],
  "Any Exception": ["Any Exception", "Exception", "Remarks", "Notes", "Comment"],
  "Declaration format (if declaration required)": [
    "Declaration format (if declaration required)",
    "Declaration",
    "Declaration Format",
    "Declaration Required",
    "Declaration Text"
  ]
};

function normalizeHeader(h){
  return (h ?? "").toString().trim().toLowerCase().replace(/\s+/g, " ");
}

function safeValue(v){
  if(v === null || v === undefined) return "";
  return ("" + v).trim();
}

function setProgress(pct, msg){
  bar.style.width = pct + "%";
  statusEl.textContent = msg;
}

function updateSummary(inputLines = 0, outputRows = 0, errors = 0){
  summary.innerHTML = `
    <span>Total input lines: ${inputLines}</span>
    <span>Output rows: ${outputRows}</span>
    <span>Errors: ${errors}</span>
  `;
}

function formatRawInputRows(rows){
  return rows.map((row, idx) => {
    const values = (row || []).map((cell, cellIdx) => `C${cellIdx + 1}: ${safeValue(cell) || "(blank)"}`);
    return `Line ${idx + 1}\n${values.join("\n")}`;
  }).join("\n\n");
}

function formatOutputRows(rows){
  return rows.map((row, idx) => {
    const values = FINAL_HEADERS.map(key => `${key}: ${row[key] || "(blank)"}`);
    return `Output ${idx + 1}\n${values.join("\n")}`;
  }).join("\n\n");
}

function findColumnIndex(headerRow, targetFinalHeader){
  const aliases = HEADER_ALIASES[targetFinalHeader] || [targetFinalHeader];
  const normalizedHeaderRow = headerRow.map(normalizeHeader);

  for(const alias of aliases){
    const idx = normalizedHeaderRow.indexOf(normalizeHeader(alias));
    if(idx !== -1) return idx;
  }
  return -1;
}

async function convertExcel(){
  const errors = [];

  try{
    const file = fileInput.files?.[0];
    if(!file){
      throw new Error("Pehle Excel file upload karo (.xlsx)");
    }

    preview.value = "";
    outputJS.value = "";
    outputValues.value = "";
    errorLog.value = "";
    lastJS = "";
    updateSummary(0, 0, 0);

    setProgress(5, "Reading Excel...");

    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    setProgress(15, "Parsing rows...");

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    if(!rows || rows.length < 2){
      throw new Error("Excel me data nahi mila. Header + rows hona chahiye.");
    }

    preview.value = formatRawInputRows(rows);

    const headerRow = rows[0].map(h => (h ?? "").toString().trim());

    const colMap = {};
    const missing = [];

    for(const finalHeader of FINAL_HEADERS){
      const idx = findColumnIndex(headerRow, finalHeader);
      if(idx === -1){
        missing.push(finalHeader);
      } else {
        colMap[finalHeader] = idx;
      }
    }

    const allowMissingKey = missing.length === 1 && missing[0] === "InsurerRequirement";
    if(missing.length > 0 && !allowMissingKey){
      errors.push("Missing headers:\n- " + missing.join("\n- "));
      throw new Error("Excel headers match nahi ho rahe.");
    }

    if(allowMissingKey){
      errors.push("Warning: InsurerRequirement column missing hai. App Insurer + Requirement se auto bana raha hai.");
    }

    setProgress(30, "Converting...");

    const out = [];
    const totalRows = rows.length - 1;

    for(let i=1; i<rows.length; i++){
      const row = rows[i];
      if(!row || row.join("").trim() === ""){
        errors.push(`Line ${i + 1}: blank row skip ki gayi.`);
        continue;
      }

      const obj = {};
      for(const key of FINAL_HEADERS){
        const idx = colMap[key];
        obj[key] = (idx === undefined) ? "" : safeValue(row[idx]);
      }

      if(!obj["InsurerRequirement"]){
        obj["InsurerRequirement"] = (obj["Insurer"] + obj["Requirement"]).trim();
      }

      if(!obj["Insurer"]){
        errors.push(`Line ${i + 1}: Insurer blank hai.`);
      }

      if(!obj["Requirement"]){
        errors.push(`Line ${i + 1}: Requirement blank hai.`);
      }

      out.push(obj);

      if(i % 200 === 0){
        const pct = Math.min(95, Math.round((i / totalRows) * 100));
        setProgress(pct, `Converting... ${i}/${totalRows}`);
        await new Promise(r => setTimeout(r, 0));
      }
    }

    setProgress(97, "Generating JS output...");

    let jsText = "const endorsementData = [\n";
    for(let i=0; i<out.length; i++){
      const prettyObj = JSON.stringify(out[i], null, 4).replace(/\n/g, "\n    ");
      jsText += "    " + prettyObj + (i === out.length - 1 ? "\n" : ",\n");
    }
    jsText += "];\n";

    lastJS = jsText;
    outputJS.value = jsText;
    outputValues.value = formatOutputRows(out);
    errorLog.value = errors.length ? errors.join("\n\n") : "No errors found.";
    updateSummary(totalRows, out.length, errors.length);

    setProgress(100, `Done! Rows: ${out.length}, Errors: ${errors.length}`);
  } catch(err){
    errors.push(err.message || String(err));
    errorLog.value = errors.join("\n\n");
    updateSummary(0, 0, errors.length);
    setProgress(0, "Error: " + (err.message || String(err)));
  }
}

function copyJS(){
  if(!lastJS.trim()){
    alert("Pehle Convert karo.");
    return;
  }
  navigator.clipboard.writeText(lastJS);
  statusEl.textContent = "✅ Copied!";
}

function downloadJS(){
  if(!lastJS.trim()){
    alert("Pehle Convert karo.");
    return;
  }
  const blob = new Blob([lastJS], { type: "text/javascript" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "endorsementData.js";
  a.click();
  URL.revokeObjectURL(a.href);
  statusEl.textContent = "✅ Download started!";
}

function clearAll(){
  fileInput.value = "";
  outputJS.value = "";
  preview.value = "";
  outputValues.value = "";
  errorLog.value = "";
  lastJS = "";
  updateSummary(0, 0, 0);
  setProgress(0, "Ready.");
}

btnConvert.addEventListener("click", convertExcel);
btnCopy.addEventListener("click", copyJS);
btnDownload.addEventListener("click", downloadJS);
btnClear.addEventListener("click", clearAll);
