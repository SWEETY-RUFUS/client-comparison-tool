let selectedFile = null;
let inputWorkbook = null;
let outputWorkbook = null;

const REQUIRED = ["CL_STORE_NAME", "CL_PA_LINE1"];

const fileInput = document.getElementById("fileInput");
const chooseBtn = document.getElementById("chooseBtn");
const dropZone = document.getElementById("dropZone");
const fileInfo = document.getElementById("fileInfo");
const sheetSection = document.getElementById("sheetSection");
const sheetList = document.getElementById("sheetList");
const warnings = document.getElementById("warnings");
const compareBtn = document.getElementById("compareBtn");
const resultSection = document.getElementById("resultSection");
const resultSummary = document.getElementById("resultSummary");
const downloadBtn = document.getElementById("downloadBtn");

chooseBtn.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", e => {
  if (e.target.files.length) loadFile(e.target.files[0]);
});

["dragenter", "dragover"].forEach(evt => {
  dropZone.addEventListener(evt, e => {
    e.preventDefault();
    dropZone.classList.add("dragover");
  });
});
["dragleave", "drop"].forEach(evt => {
  dropZone.addEventListener(evt, e => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
  });
});
dropZone.addEventListener("drop", e => {
  if (e.dataTransfer.files.length) loadFile(e.dataTransfer.files[0]);
});

async function loadFile(file) {
  // Completely reset previous workbook/result state.
  selectedFile = file;
  inputWorkbook = null;
  outputWorkbook = null;
  resultSection.classList.add("hidden");
  resultSummary.innerHTML = "";
  downloadBtn.disabled = true;
  sheetSection.classList.add("hidden");
  sheetList.innerHTML = "";
  warnings.innerHTML = "";

  try {
    const buffer = await file.arrayBuffer();
    inputWorkbook = XLSX.read(buffer, {
      type: "array",
      cellDates: true,
      cellNF: true,
      cellStyles: true
    });

    fileInfo.classList.remove("hidden");
    fileInfo.innerHTML = `<p><strong>Selected:</strong> ${escapeHtml(file.name)}
      &nbsp; | &nbsp; <strong>Sheets:</strong> ${inputWorkbook.SheetNames.length}</p>`;

    renderSheets();
  } catch (err) {
    alert("Could not open this Excel file. Please check that it is a valid .xlsx/.xls workbook.");
    console.error(err);
  }
}

function renderSheets() {
  const valid = [];
  const invalid = [];

  inputWorkbook.SheetNames.forEach((name, index) => {
    const ws = inputWorkbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    const headers = (rows[0] || []).map(v => String(v).trim());
    const missing = REQUIRED.filter(c => !headers.includes(c));

    if (missing.length) invalid.push({name, index, missing});
    else valid.push({name, index});
  });

  sheetList.innerHTML = inputWorkbook.SheetNames.map((name, i) => {
    const invalidSheet = invalid.find(x => x.index === i);
    return `<div class="sheet-row">
      <span><strong>${i + 1}.</strong> ${escapeHtml(name)}</span>
      <small>${invalidSheet ? "⚠ Missing: " + invalidSheet.missing.join(", ") :
        (i === 0 ? "Reference / oldest sheet" : "Will be compared")}</small>
    </div>`;
  }).join("");

  if (invalid.length) {
    warnings.innerHTML = `<div class="warning"><strong>Warning:</strong>
      ${invalid.map(x => `<div><strong>${escapeHtml(x.name)}</strong> will be skipped because it is missing ${x.missing.join(" and ")}.</div>`).join("")}
      </div>`;
  }

  if (valid.length < 2) {
    compareBtn.disabled = true;
    warnings.innerHTML += `<div class="warning">At least two valid client sheets are needed for comparison.</div>`;
  } else {
    compareBtn.disabled = false;
  }

  sheetSection.classList.remove("hidden");
}

function normalizeText(value) {
  if (value === null || value === undefined) return "";
  return String(value)
    .normalize("NFKC")
    .toUpperCase()
    .replace(/&/g, " AND ")
    .replace(/[^A-Z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeBanner(value) {
  let s = normalizeText(value);
  // Remove common generic business suffixes only at the end.
  s = s.replace(/\b(LLC|INC|INCORPORATED|CORP|CORPORATION|LTD|LIMITED)\b$/g, "");
  return s.replace(/\s+/g, " ").trim();
}

function normalizeAddress(value) {
  let s = normalizeText(value);

  const replacements = [
    [/\bSTREET\b/g, "ST"],
    [/\bAVENUE\b/g, "AVE"],
    [/\bROAD\b/g, "RD"],
    [/\bBOULEVARD\b/g, "BLVD"],
    [/\bDRIVE\b/g, "DR"],
    [/\bLANE\b/g, "LN"],
    [/\bCOURT\b/g, "CT"],
    [/\bCIRCLE\b/g, "CIR"],
    [/\bPARKWAY\b/g, "PKWY"],
    [/\bHIGHWAY\b/g, "HWY"],
    [/\bPLACE\b/g, "PL"],
    [/\bTERRACE\b/g, "TER"],
    [/\bTRAIL\b/g, "TRL"],
    [/\bSUITE\b/g, "STE"],
    [/\bAPARTMENT\b/g, "APT"],
    [/\bNORTH\b/g, "N"],
    [/\bSOUTH\b/g, "S"],
    [/\bEAST\b/g, "E"],
    [/\bWEST\b/g, "W"]
  ];
  replacements.forEach(([pattern, repl]) => s = s.replace(pattern, repl));

  // Normalize "# 101", "STE 101", "SUITE 101" into the same form where possible.
  s = s.replace(/#\s*(\d+[A-Z]?)/g, "STE $1");
  return s.replace(/\s+/g, " ").trim();
}

function tokenize(s) {
  return new Set(s.split(" ").filter(Boolean));
}

function jaccard(a, b) {
  const A = tokenize(a), B = tokenize(b);
  if (!A.size || !B.size) return 0;
  let inter = 0;
  A.forEach(x => { if (B.has(x)) inter++; });
  const union = new Set([...A, ...B]).size;
  return union ? inter / union : 0;
}

function levenshtein(a, b) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  let prev = Array.from({length: b.length + 1}, (_, i) => i);
  for (let i = 0; i < a.length; i++) {
    const cur = [i + 1];
    for (let j = 0; j < b.length; j++) {
      cur[j + 1] = Math.min(
        cur[j] + 1,
        prev[j + 1] + 1,
        prev[j] + (a[i] === b[j] ? 0 : 1)
      );
    }
    prev = cur;
  }
  return prev[b.length];
}

function similarity(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  const maxLen = Math.max(a.length, b.length);
  const editScore = maxLen ? 1 - levenshtein(a, b) / maxLen : 0;
  return Math.max(editScore, jaccard(a, b));
}

function addressSimilarity(a, b) {
  if (a === b) return 1;
  const A = normalizeAddress(a), B = normalizeAddress(b);
  if (!A || !B) return 0;
  if (A === B) return 1;

  // Token similarity handles common additions such as STE 100.
  const tokenScore = jaccard(A, B);

  // If one address is a strong subset of the other, allow normal unit/suite additions.
  const subset = A.includes(B) || B.includes(A);
  if (subset && tokenScore >= 0.75) return 0.96;

  return Math.max(tokenScore, similarity(A, B) * 0.85);
}

function bannersCompatible(a, b) {
  const A = normalizeBanner(a), B = normalizeBanner(b);
  if (!A || !B) return false;
  if (A === B) return true;

  const score = similarity(A, B);

  // Handles names such as:
  // MOSELEY
  // THE MOSELEYPROPER FARE
  // THE MOSELEY/PROPER FARE
  const compactA = A.replace(/\s+/g, "");
  const compactB = B.replace(/\s+/g, "");
  if (compactA.includes(compactB) || compactB.includes(compactA)) return true;

  // Strong similarity for banner variations.
  return score >= 0.72;
}

function classifyMatch(currentBanner, currentAddress, oldBanner, oldAddress) {
  const bExact = normalizeBanner(currentBanner) === normalizeBanner(oldBanner);
  const aExact = normalizeAddress(currentAddress) === normalizeAddress(oldAddress);

  if (bExact && aExact) return "Exact";

  const aScore = addressSimilarity(currentAddress, oldAddress);
  const bCompatible = bannersCompatible(currentBanner, oldBanner);

  // Same/normalized address + banner variation.
  if (aScore >= 0.95 && bCompatible) return "Address + Banner variation";

  // Same/compatible banner + address variation.
  const bScore = similarity(normalizeBanner(currentBanner), normalizeBanner(oldBanner));
  if (bExact && aScore >= 0.75) return "Banner + Address variation";
  if (bScore >= 0.90 && aScore >= 0.75) return "Banner + Address variation";

  return "";
}

function buildRows(ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
}

function headerMap(headers) {
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h).trim();
    if (key && map[key] === undefined) map[key] = i;
  });
  return map;
}

function compareWorkbook() {
  if (!inputWorkbook) return;

  compareBtn.disabled = true;
  compareBtn.textContent = "Comparing...";

  try {
    const sheetData = {};
    const validNames = [];

    for (const name of inputWorkbook.SheetNames) {
      const rows = buildRows(inputWorkbook.Sheets[name]);
      const headers = rows[0] || [];
      const map = headerMap(headers);
      if (REQUIRED.every(c => map[c] !== undefined)) {
        sheetData[name] = {rows, map};
        validNames.push(name);
      }
    }

    if (validNames.length < 2) throw new Error("Fewer than two valid sheets are available.");

    const wb = XLSX.utils.book_new();
    let totalMatches = 0;
    let totalExact = 0;
    let totalVariation = 0;

    for (let i = 0; i < inputWorkbook.SheetNames.length; i++) {
      const name = inputWorkbook.SheetNames[i];
      const originalWs = inputWorkbook.Sheets[name];

      // Preserve invalid/non-client sheets unchanged.
      if (!sheetData[name]) {
        XLSX.utils.book_append_sheet(wb, originalWs, name);
        continue;
      }

      const sourceRows = sheetData[name].rows.map(r => [...r]);
      const headers = sourceRows[0] || [];
      const map = sheetData[name].map;

      let matchedIdx = headers.indexOf("Matched_Sheet");
      let typeIdx = headers.indexOf("Match_Type");

      if (matchedIdx < 0) {
        matchedIdx = headers.length;
        headers[matchedIdx] = "Matched_Sheet";
      }
      if (typeIdx < 0) {
        typeIdx = headers.length;
        headers[typeIdx] = "Match_Type";
      }

      // First sheet has no comparison results.
      if (i === 0) {
        for (let r = 1; r < sourceRows.length; r++) {
          sourceRows[r][matchedIdx] = "";
          sourceRows[r][typeIdx] = "";
        }
      } else {
        for (let r = 1; r < sourceRows.length; r++) {
          const row = sourceRows[r];
          const banner = row[map["CL_STORE_NAME"]];
          const address = row[map["CL_PA_LINE1"]];

          let matchedSheet = "";
          let matchType = "";

          // IMPORTANT: only valid older client sheets are candidates.
          // First/oldest match wins.
          for (let j = 0; j < i; j++) {
            const oldName = inputWorkbook.SheetNames[j];
            if (!sheetData[oldName]) continue;

            const old = sheetData[oldName];
            const oldBannerIdx = old.map["CL_STORE_NAME"];
            const oldAddressIdx = old.map["CL_PA_LINE1"];

            // Build candidate rows on demand. This avoids relying only on exact keys.
            for (let or = 1; or < old.rows.length; or++) {
              const oldRow = old.rows[or];
              const oldBanner = oldRow[oldBannerIdx];
              const oldAddress = oldRow[oldAddressIdx];

              const type = classifyMatch(banner, address, oldBanner, oldAddress);
              if (type) {
                matchedSheet = oldName;
                matchType = type;
                break;
              }
            }
            if (matchedSheet) break;
          }

          row[matchedIdx] = matchedSheet;
          row[typeIdx] = matchType;

          if (matchedSheet) {
            totalMatches++;
            if (matchType === "Exact") totalExact++;
            else totalVariation++;
          }
        }
      }

      const wsOut = XLSX.utils.aoa_to_sheet(sourceRows);
      XLSX.utils.book_append_sheet(wb, wsOut, name);
    }

    outputWorkbook = wb;

    resultSummary.innerHTML = `
      <div class="stat"><strong>${totalMatches}</strong><br>Total matches</div>
      <div class="stat"><strong>${totalExact}</strong><br>Exact matches</div>
      <div class="stat"><strong>${totalVariation}</strong><br>Variation matches</div>
      <p>Comparison completed. The original workbook has not been changed.</p>
    `;
    resultSection.classList.remove("hidden");
    downloadBtn.disabled = false;
  } catch (err) {
    alert("Comparison failed: " + err.message);
    console.error(err);
  } finally {
    compareBtn.disabled = false;
    compareBtn.textContent = "Compare Workbook";
  }
}

compareBtn.addEventListener("click", compareWorkbook);

downloadBtn.addEventListener("click", () => {
  if (!outputWorkbook || !selectedFile) return;
  const base = selectedFile.name.replace(/\.(xlsx|xls)$/i, "");
  XLSX.writeFile(outputWorkbook, `${base}_Comparison_Result.xlsx`);
});

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
