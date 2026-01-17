const fileInput = document.getElementById("fileInput");
const outputJS = document.getElementById("outputJS");
const preview = document.getElementById("preview");
const statusEl = document.getElementById("status");
const bar = document.getElementById("bar");

const btnConvert = document.getElementById("btnConvert");
const btnCopy = document.getElementById("btnCopy");
const btnDownload = document.getElementById("btnDownload");
const btnClear = document.getElementById("btnClear");

let lastJS = "";

// FINAL OUTPUT headers (fixed)
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

// AUTO header mapping (Excel me jo bhi naam ho, tool samajh jayega)
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
  return (h ?? "")
    .toString()
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function safeValue(v){
  if(v === null || v === undefined) return "";
  return ("" + v).trim();
}

function setProgress(pct, msg){
  bar.style.width = pct + "%";
  statusEl.textContent = msg;
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
  const file = fileInput.files?.[0];
  if(!file){
    alert("Pehle Excel file upload karo (.xlsx)");
    return;
  }

  setProgress(5, "Reading Excel...");

  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array" });

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  setProgress(15, "Parsing rows...");

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  if(!rows || rows.length < 2){
    alert("Excel me data nahi mila. (Header + rows hona chahiye)");
    return;
  }

  const headerRow = rows[0].map(h => (h ?? "").toString().trim());

  // Build column map
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

  // Agar InsurerRequirement missing hai, hum auto bana denge
  const allowMissingKey = missing.length === 1 && missing[0] === "InsurerRequirement";
  if(missing.length > 0 && !allowMissingKey){
    alert(
      "Excel headers match nahi ho rahe.\nMissing (auto detect nahi hua):\n- " +
      missing.join("\n- ") +
      "\n\nTip: Excel ke headers ko thoda correct karo."
    );
    return;
  }

  setProgress(30, "Converting to endorsementData...");

  const out = [];
  const totalRows = rows.length - 1;

  for(let i=1; i<rows.length; i++){
    const row = rows[i];
    if(!row || row.join("").trim() === "") continue;

    const obj = {};
    for(const key of FINAL_HEADERS){
      const idx = colMap[key];
      obj[key] = (idx === undefined) ? "" : safeValue(row[idx]);
    }

    // Auto build InsurerRequirement
    if(!obj["InsurerRequirement"] || obj["InsurerRequirement"].trim() === ""){
      obj["InsurerRequirement"] = (obj["Insurer"] + obj["Requirement"]).trim();
    }

    out.push(obj);

    if(i % 200 === 0){
      const pct = Math.min(95, Math.round((i / totalRows) * 100));
      setProgress(pct, `Converting... ${i}/${totalRows} rows`);
      await new Promise(r => setTimeout(r, 0));
    }
  }

  setProgress(97, "Generating JS output...");

  // EXACT output style
  let jsText = "const endorsementData = [\n";

  for(let i=0; i<out.length; i++){
    const prettyObj = JSON.stringify(out[i], null, 4)
      .replace(/\n/g, "\n    ");

    jsText += "    " + prettyObj;
    jsText += (i === out.length - 1) ? "\n" : ",\n";
  }

  jsText += "];\n";

  lastJS = jsText;
  outputJS.value = jsText;

  // Preview first 20
  preview.value = out.slice(0, 20).map((r, idx) => {
    return `${idx+1}) ${r["Insurer"]} | ${r["Requirement"]} | ${r["Endorsement type"]}`;
  }).join("\n");

  setProgress(100, `✅ Done! Total rows converted: ${out.length}`);
}

function copyJS(){
  if(!lastJS.trim()){
    alert("Pehle Convert karo.");
    return;
  }
  navigator.clipboard.writeText(lastJS);
  statusEl.textContent = "✅ Copied to clipboard!";
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
  lastJS = "";
  setProgress(0, "Ready.");
}

// Button events
btnConvert.addEventListener("click", convertExcel);
btnCopy.addEventListener("click", copyJS);
btnDownload.addEventListener("click", downloadJS);
btnClear.addEventListener("click", clearAll);
