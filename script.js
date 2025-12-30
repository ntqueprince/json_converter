let headers = [];
let cleanedRows = [];

function normalize(v){
  return String(v ?? "").replace(/\r/g,"");
}

document.getElementById("fileInput").addEventListener("change", function(){
  const file = this.files[0];
  if(!file) return;

  const reader = new FileReader();
  reader.onload = e => {
    const wb = XLSX.read(new Uint8Array(e.target.result), {type:"array"});
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet,{header:1,raw:true});
    renderTable(data);
  };
  reader.readAsArrayBuffer(file);
});

function renderTable(data){
  const table = document.getElementById("previewTable");
  table.innerHTML = "";
  cleanedRows = [];
  headers = data[0] || [];

  document.getElementById("colCount").innerText = headers.length;
  document.getElementById("rowCount").innerText = data.length - 1;

  // Header
  const thead = document.createElement("thead");
  const hr = document.createElement("tr");
  headers.forEach(h=>{
    const th = document.createElement("th");
    th.textContent = h;
    hr.appendChild(th);
  });
  thead.appendChild(hr);
  table.appendChild(thead);

  // Body
  const tbody = document.createElement("tbody");
  data.slice(1).forEach(row=>{
    const tr = document.createElement("tr");
    let fixed = [];

    headers.forEach((_, i)=>{
      const td = document.createElement("td");
      const val = normalize(row[i]);
      td.textContent = val;
      tr.appendChild(td);
      fixed.push(val);
    });

    cleanedRows.push(fixed.join("\t"));
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
}

/* ===== ACTIONS ===== */

function copyClean(){
  navigator.clipboard.writeText(cleanedRows.join("\n"));
  alert("Cleaned data copied");
}

function downloadTXT(){
  downloadFile(
    cleanedRows.join("\n"),
    "cleaned-data.txt",
    "text/plain"
  );
}

function downloadJS(){
  let js = "const endorsementData = [\n";
  cleanedRows.forEach(r=>{
    const c = r.split("\t");
    let obj = {};
    headers.forEach((h,i)=>{
      obj[h] = c[i] || "";
    });
    js += "  " + JSON.stringify(obj, null, 2).replace(/\n/g,"\n  ") + ",\n";
  });
  js += "];";

  downloadFile(js, "endorsementData.js", "application/javascript");
}

function downloadFile(content, filename, type){
  const blob = new Blob([content], {type});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
