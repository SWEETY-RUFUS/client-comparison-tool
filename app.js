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

  compareBtn.disabled = valid.length < 2;
  if (valid.length < 2) {
    warnings.innerHTML += `<div class="warning">At least two valid sheets are needed for comparison.</div>`;
  }

  sheetSection.classList.remove("hidden");
}

/* -------------------- NORMALIZATION -------------------- */

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
  const A = normalizeAddress(a);
  const B = normalizeAddress(b);

  if (!A || !B) return 0;
  if (A === B) return 1;

  const tokenScore = jaccard(A, B);

  // Allows normal additions such as STE 100 while requiring strong overlap.
  const subset = A.includes(B) || B.includes(A);
  if (subset && tokenScore >= 0.75) return 0.96;

  return Math.max(tokenScore, similarity(A, B) * 0.85);
}

function bannersCompatible(a, b) {
  const A = normalizeBanner(a);
  const B = normalizeBanner(b);

  if (!A || !B) return false;
  if (A === B) return true;

  const score = similarity(A, B);

  const compactA = A.replace(/\s+/g, "");
  const compactB = B.replace(/\s+/g, "");

  if (compactA.includes(compactB) || compactB.includes(compactA)) return true;

  return score >= 0.72;
}

/*
  Important performance fix:
  The previous version compared EVERY row in an older sheet with EVERY
  current row. That becomes very slow and can make the browser show
  "Page Unresponsive".

  This version creates indexes first:
    - exact Banner + Address
    - normalized Address
    - normalized Banner

  Therefore a row normally checks only a small list of candidates instead
  of scanning the whole older worksheet.
*/

function buildIndex(data) {
  const bannerIdx = data.map["CL_STORE_NAME"];
  const addressIdx = data.map["CL_PA_LINE1"];

  const exact = new Map();
  const address = new Map();
  const banner = new Map();

  for (let r = 1; r < data.rows.length; r++) {
    const row = data.rows[r];
    const b = normalizeBanner(row[bannerIdx]);
    const a = normalizeAddress(row[addressIdx]);
    if (!b || !a) continue;

    const record = { row, b, a };

    const exactKey = b + "||" + a;
    if (!exact.has(exactKey)) exact.set(exactKey, []);
    exact.get(exactKey).push(record);

    if (!address.has(a)) address.set(a, []);
    address.get(a).push(record);

    if (!banner.has(b)) banner.set(b, []);
    banner.get(b).push(record);
  }

  return {exact, address, banner};
}

function findMatch(currentBanner, currentAddress, oldIndex) {
  const b = normalizeBanner(currentBanner);
  const a = normalizeAddress(currentAddress);

  if (!b || !a) return "";

  // 1. Exact normalized Banner + Address.
  const exactKey = b + "||" + a;
  if (oldIndex.exact.has(exactKey)) {
    return "Exact";
  }

  // 2. Same/normalized address + banner variation.
  // Usually this is a very small candidate list.
  const addressCandidates = oldIndex.address.get(a) || [];
  for (const candidate of addressCandidates) {
    if (bannersCompatible(b, candidate.b)) {
      return "Address + Banner variation";
    }
  }

  // 3. Same banner + address variation.
  const bannerCandidates = oldIndex.banner.get(b) || [];
  for (const candidate of bannerCandidates) {
    if (addressSimilarity(a, candidate.a) >= 0.75) {
      return "Banner + Address variation";
    }
  }

  // 4. Controlled fallback:
  // If normalized address is not identical, use address token overlap,
  // but only against records sharing at least one meaningful address token.
  // This avoids the old O(rows x rows) scan.
  const addressTokens = [...tokenize(a)].filter(x => x.length >= 3);
  const candidateSet = new Set();

  // Build fallback candidate records from exact banner only first.
  // This is intentionally conservative to avoid false positives.
  for (const candidate of bannerCandidates) {
    candidateSet.add(candidate);
  }

  for (const candidate of candidateSet) {
    const aScore = addressSimilarity(a, candidate.a);
    if (aScore >= 0.90 && similarity(b, candidate.b) >= 0.85) {
      return "Banner + Address variation";
    }
  }

  return "";
}

/* -------------------- COMPARISON -------------------- */

function compareWorkbook() {
  if (!inputWorkbook) return;

  compareBtn.disabled = true;
  compareBtn.textContent = "Comparing...";

  // Give the browser a chance to paint the button/status before processing.
  setTimeout(() => {
    try {
      const sheetData = {};
      const validNames = [];

      for (const name of inputWorkbook.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(
          inputWorkbook.Sheets[name],
          { header: 1, defval: "", raw: false }
        );

        const headers = rows[0] || [];
        const map = headerMap(headers);

        if (REQUIRED.every(c => map[c] !== undefined)) {
          sheetData[name] = {rows, map, index: null};
          validNames.push(name);
        }
      }

      if (validNames.length < 2) {
        throw new Error("Fewer than two valid sheets are available.");
      }

      // Build each valid sheet's indexes once.
      for (const name of validNames) {
        sheetData[name].index = buildIndex(sheetData[name]);
      }

      const wb = XLSX.utils.book_new();

      let totalMatches = 0;
      let totalExact = 0;
      let totalVariation = 0;
      let comparisonRows = 0;

      for (let i = 0; i < inputWorkbook.SheetNames.length; i++) {
        const name = inputWorkbook.SheetNames[i];
        const originalWs = inputWorkbook.Sheets[name];

        if (!sheetData[name]) {
          XLSX.utils.book_append_sheet(wb, originalWs, name);
          continue;
        }

        const data = sheetData[name];
        const sourceRows = data.rows.map(r => [...r]);
        const headers = sourceRows[0] || [];
        const map = data.map;

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

        if (i === 0) {
          for (let r = 1; r < sourceRows.length; r++) {
            sourceRows[r][matchedIdx] = "";
            sourceRows[r][typeIdx] = "";
          }
        } else {
          for (let r = 1; r < sourceRows.length; r++) {
            const row = sourceRows[r];
            const currentBanner = row[map["CL_STORE_NAME"]];
            const currentAddress = row[map["CL_PA_LINE1"]];

            let matchedSheet = "";
            let matchType = "";

            // Check only OLDER sheets and stop at the first match.
            for (let j = 0; j < i; j++) {
              const oldName = inputWorkbook.SheetNames[j];
              if (!sheetData[oldName]) continue;

              const type = findMatch(
                currentBanner,
                currentAddress,
                sheetData[oldName].index
              );

              if (type) {
                matchedSheet = oldName;
                matchType = type;
                break;
              }
            }

            row[matchedIdx] = matchedSheet;
            row[typeIdx] = matchType;

            comparisonRows++;
            if (matchedSheet) {
              totalMatches++;
              if (matchType === "Exact") totalExact++;
              else totalVariation++;
            }
          }
        }

        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sourceRows), name);
      }

      outputWorkbook = wb;

      resultSummary.innerHTML = `
        <div class="stat"><strong>${totalMatches}</strong><br>Total matches</div>
        <div class="stat"><strong>${totalExact}</strong><br>Exact matches</div>
        <div class="stat"><strong>${totalVariation}</strong><br>Variation matches</div>
        <div class="stat"><strong>${comparisonRows}</strong><br>Rows checked</div>
        <p><strong>Completed.</strong> The original workbook was not changed.</p>
        <p>The matching engine uses indexes instead of comparing every row against every other row, so large workbooks should not freeze the browser.</p>
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
  }, 50);
}

compareBtn.addEventListener("click", compareWorkbook);

downloadBtn.addEventListener("click", () => {
  if (!outputWorkbook || !selectedFile) return;

  const base = selectedFile.name.replace(/\.(xlsx|xls)$/i, "");
  XLSX.writeFile(outputWorkbook, `${base}_Comparison_Result.xlsx`);
});

function headerMap(headers) {
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h).trim();
    if (key && map[key] === undefined) map[key] = i;
  });
  return map;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
