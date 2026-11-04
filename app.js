let selectedFile = null;
let inputWorkbook = null;
let outputWorkbook = null;

const REQUIRED = ["CL_STORE_NAME", "CL_PA_LINE1"];
const RESERVED_SHEETS = ["FORCE", "PSU", "LOCAL CODE MATCHING", "MATCHED_ID_EXTRACT", "MATCHED_CL_ID_EXTRACT"];

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
    if (RESERVED_SHEETS.includes(name.trim().toUpperCase())) return;
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
      <small>${RESERVED_SHEETS.includes(name.trim().toUpperCase()) ? "Reserved output sheet; excluded from comparison" : invalidSheet ? "⚠ Missing: " + invalidSheet.missing.join(", ") :
        (valid.filter(x => x.index > i).length === 0 ? "Reference sheet" : "Will compare with sheets to the right")}</small>
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
    return { type: "Exact", record: oldIndex.exact.get(exactKey)[0] };
  }

  // 2. Same/normalized address + banner variation.
  // Usually this is a very small candidate list.
  const addressCandidates = oldIndex.address.get(a) || [];
  for (const candidate of addressCandidates) {
    if (bannersCompatible(b, candidate.b)) {
      return { type: "Address + Banner variation", record: candidate };
    }
  }

  // 3. Same banner + address variation.
  const bannerCandidates = oldIndex.banner.get(b) || [];
  for (const candidate of bannerCandidates) {
    if (addressSimilarity(a, candidate.a) >= 0.75) {
      return { type: "Banner + Address variation", record: candidate };
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
      return { type: "Banner + Address variation", record: candidate };
    }
  }

  return "";
}

/* -------------------- COMPARISON -------------------- */

function compareWorkbook() {
  if (!inputWorkbook) return;

  compareBtn.disabled = true;
  compareBtn.textContent = "Comparing...";

  setTimeout(() => {
    try {
      const sheetData = {};
      const validNames = [];
      const forceSheetName = "Force";
      const extractSheetName = "Local Code Matching";
      const matchedClExtractSheetName = "PSU";
      const utilityNames = new Set(RESERVED_SHEETS);

      // Carry forward the permanent Force mapping register without deleting entries.
      // Older versions stored the two-column mapping table in Local Code Matching;
      // newer versions store it in Force with the additional source-sheet column.
      const mappingSourceNames = inputWorkbook.SheetNames.filter(n => {
        const rows = XLSX.utils.sheet_to_json(inputWorkbook.Sheets[n], { header: 1, defval: "", raw: false });
        const h = headerMap(rows[0] || []);
        return h["Client_Unique_ID"] !== undefined && h["PTY_LOCAL_CODE"] !== undefined;
      });
      let forceRows = [["Client_Unique_ID", "PTY_LOCAL_CODE", "Source_Client_Sheet"]];
      for (const sourceName of mappingSourceNames) {
        const oldRows = XLSX.utils.sheet_to_json(inputWorkbook.Sheets[sourceName], { header: 1, defval: "", raw: false });
        const oldMap = headerMap(oldRows[0] || []);
        const idCol = oldMap["Client_Unique_ID"];
        const codeCol = oldMap["PTY_LOCAL_CODE"];
        const sourceCol = oldMap["Source_Client_Sheet"];
        for (let r = 1; r < oldRows.length; r++) {
          const id = String(oldRows[r][idCol] ?? "").trim();
          // Keep rows with an ID even when the local code is blank.
          if (!id) continue;
          forceRows.push([
            oldRows[r][idCol] ?? "",
            oldRows[r][codeCol] ?? "",
            sourceCol === undefined ? "" : (oldRows[r][sourceCol] ?? "")
          ]);
        }
      }

      // Force and the generated extraction sheet are never client comparison sheets.
      for (const name of inputWorkbook.SheetNames) {
        if (utilityNames.has(name.trim().toUpperCase())) continue;
        const rows = XLSX.utils.sheet_to_json(inputWorkbook.Sheets[name], { header: 1, defval: "", raw: false });
        const headers = rows[0] || [];
        const map = headerMap(headers);
        if (REQUIRED.every(c => map[c] !== undefined)) {
          sheetData[name] = { rows, map, index: null };
          validNames.push(name);
        }
      }

      if (validNames.length < 2) throw new Error("Fewer than two valid client sheets are available. PSU, Local Code Matching, and Force are reserved utility worksheets.");
      for (const name of validNames) sheetData[name].index = buildIndex(sheetData[name]);

      const wb = XLSX.utils.book_new();
      let totalMatches = 0, totalExact = 0, totalVariation = 0, comparisonRows = 0;
      let forceAdded = false;

      function addForceSheet() {
        const forceWs = XLSX.utils.aoa_to_sheet(forceRows);
        forceWs["!cols"] = [{wch: 24}, {wch: 22}, {wch: 28}];
        forceWs["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomLeft", state: "frozen" };
        XLSX.utils.book_append_sheet(wb, forceWs, forceSheetName);
        forceAdded = true;
      }

      for (const originalName of inputWorkbook.SheetNames) {
        if (utilityNames.has(originalName.trim().toUpperCase())) continue;

        if (!sheetData[originalName]) {
          XLSX.utils.book_append_sheet(wb, inputWorkbook.Sheets[originalName], originalName);
          continue;
        }

        const data = sheetData[originalName];
        const sourceRows = data.rows.map(r => [...r]);
        const headers = sourceRows[0] || [];
        const map = data.map;
        let matchedIdx = headers.indexOf("Matched_Sheet");
        let typeIdx = headers.indexOf("Match_Type");
        let matchedClIdIdx = headers.indexOf("Matched_CL_ID");
        let ptyIdx = headers.indexOf("PTY_LOCAL_CODE");

        if (matchedIdx < 0) { matchedIdx = headers.length; headers[matchedIdx] = "Matched_Sheet"; }
        if (typeIdx < 0) { typeIdx = headers.length; headers[typeIdx] = "Match_Type"; }
        if (matchedClIdIdx < 0) { matchedClIdIdx = headers.length; headers[matchedClIdIdx] = "Matched_CL_ID"; }
        if (ptyIdx < 0) { ptyIdx = headers.length; headers[ptyIdx] = "PTY_LOCAL_CODE"; }

        const position = validNames.indexOf(originalName);
        for (let r = 1; r < sourceRows.length; r++) {
          const row = sourceRows[r];
          let matchedSheet = "", matchType = "", matchedClId = "";

          if (position !== validNames.length - 1) {
            for (let j = validNames.length - 1; j > position; j--) {
              const newerName = validNames[j];
              const match = findMatch(row[map["CL_STORE_NAME"]], row[map["CL_PA_LINE1"]], sheetData[newerName].index);
              if (match) {
                matchedSheet = newerName;
                matchType = match.type;
                const clIdColumn = sheetData[newerName].map["CL_ID"];
                matchedClId = clIdColumn === undefined ? "" : (match.record.row[clIdColumn] ?? "");
                break;
              }
            }
            comparisonRows++;
          }

          row[matchedIdx] = matchedSheet;
          row[typeIdx] = matchType;
          row[matchedClIdIdx] = matchedClId;

          // Use the entire shared Force map for every client sheet. If no new
          // mapping is found, retain any PTY_LOCAL_CODE already present in input.
          const excelRow = r + 1;
          const idColLetter = XLSX.utils.encode_col(matchedClIdIdx);
          const oldPtyValue = row[ptyIdx] ?? "";
          const ptyCell = XLSX.utils.encode_cell({r, c: ptyIdx});
          const fallback = oldPtyValue === "" ? '""' : `"${String(oldPtyValue).replace(/"/g, '""')}"`;
          const formula = `IF(${idColLetter}${excelRow}="",${fallback},IFERROR(XLOOKUP(1,('${forceSheetName}'!$A:$A=${idColLetter}${excelRow})*('${forceSheetName}'!$B:$B<>""),'${forceSheetName}'!$B:$B,${fallback}),${fallback}))`;
          row[ptyIdx] = { __formula: formula };

          if (matchedSheet) {
            totalMatches++;
            if (matchType === "Exact") totalExact++; else totalVariation++;
          }
        }

        const outWs = XLSX.utils.aoa_to_sheet(sourceRows.map((row, r) => row.map((value, c) => {
          if (r > 0 && c === ptyIdx && value && typeof value === "object" && value.__formula) return "";
          return value;
        })));
        for (let r = 1; r < sourceRows.length; r++) {
          const value = sourceRows[r][ptyIdx];
          if (value && typeof value === "object" && value.__formula) {
            outWs[XLSX.utils.encode_cell({r, c: ptyIdx})] = { t: "str", f: value.__formula };
          }
        }
        const range = XLSX.utils.decode_range(outWs["!ref"] || "A1");
        range.e.c = Math.max(range.e.c, ptyIdx);
        outWs["!ref"] = XLSX.utils.encode_range(range);
        XLSX.utils.book_append_sheet(wb, outWs, originalName);
      }

      // Utility sheets are appended after all client sheets in the requested order.
      addForceSheet();

      // One reusable extraction sheet. Enter any exact worksheet name in B2.
      // For every row across all client sheets whose Matched_Sheet points to
      // the selected client, return the SOURCE sheet name, that row's actual
      // CL_ID, and its PTY_LOCAL_CODE (not Matched_CL_ID).
      const extractRows = [
        ["Client-wise Matched Record Extract", "", ""],
        ["Enter client worksheet name:", "", ""],
        ["Source_Client_Sheet", "CL_ID", "PTY_LOCAL_CODE"]
      ];
      const extractWs = XLSX.utils.aoa_to_sheet(extractRows);
      extractWs["!cols"] = [{wch: 30}, {wch: 20}, {wch: 24}];
      extractWs["!freeze"] = { xSplit: 0, ySplit: 3, topLeftCell: "A4", activePane: "bottomLeft", state: "frozen" };
      extractWs["B2"] = { t: "s", v: "" };

      const quoteSheet = name => `'${name.replace(/'/g, "''")}'`;
      const stacks = validNames.map(name => {
        const q = quoteSheet(name);
        const headers = [...(sheetData[name].rows[0] || [])];
        let matchedCol = headers.indexOf("Matched_Sheet");
        let clIdCol = headers.indexOf("CL_ID");
        let ptyCol = headers.indexOf("PTY_LOCAL_CODE");
        if (matchedCol < 0) { matchedCol = headers.length; headers[matchedCol] = "Matched_Sheet"; }
        if (ptyCol < 0) { ptyCol = headers.length; headers[ptyCol] = "PTY_LOCAL_CODE"; }
        const mc = XLSX.utils.encode_col(matchedCol);
        const cc = clIdCol < 0 ? null : XLSX.utils.encode_col(clIdCol);
        const pc = XLSX.utils.encode_col(ptyCol);
        const lastRow = Math.max(sheetData[name].rows.length, 2);
        const matchedRange = `${q}!${mc}2:${mc}${lastRow}`;
        const clIdRange = cc ? `${q}!${cc}2:${cc}${lastRow}` : `IF(${matchedRange}<>"","","")`;
        const ptyRange = `${q}!${pc}2:${pc}${lastRow}`;
        const sourceNameRange = `IF(${matchedRange}<>"","${name.replace(/"/g, '""')}","")`;
        return `IFERROR(FILTER(CHOOSE({1,2,3},${sourceNameRange},${clIdRange},${ptyRange}),${matchedRange}=$B$2),{"","",""})`;
      });
      const extractFormula = `IF($B$2="","",LET(allRows,VSTACK(${stacks.join(",")}),FILTER(allRows,CHOOSECOLS(allRows,1)<>"",{"","",""})))`;
      extractWs["A4"] = { t: "s", v: "Enter an exact client worksheet name in B2. Results show the source sheet, its actual CL_ID, and its PTY_LOCAL_CODE." };
      extractWs["A4"].s = { alignment: { wrapText: true } };
      extractWs["A5"] = { t: "n", f: extractFormula };
      extractWs["!ref"] = "A1:C5";
      // The actual-CL_ID extraction is named Local Code Matching.
      XLSX.utils.book_append_sheet(wb, extractWs, extractSheetName);

      // A second, simpler extract: only Matched_CL_ID and PTY_LOCAL_CODE
      // from rows whose Matched_Sheet equals the client name entered in B2.
      const matchedClExtractRows = [
        ["Matched CL_ID and PTY Code Extract", "", ""],
        ["Enter client worksheet name:", "", ""],
        ["Matched_CL_ID", "PTY_LOCAL_CODE", "Source_Client_Sheet"]
      ];
      const matchedClExtractWs = XLSX.utils.aoa_to_sheet(matchedClExtractRows);
      matchedClExtractWs["!cols"] = [{wch: 24}, {wch: 24}, {wch: 30}];
      matchedClExtractWs["!freeze"] = { xSplit: 0, ySplit: 3, topLeftCell: "A4", activePane: "bottomLeft", state: "frozen" };
      matchedClExtractWs["B2"] = { t: "s", v: "" };
      const matchedClStacks = validNames.map(name => {
        const q = quoteSheet(name);
        const headers = [...(sheetData[name].rows[0] || [])];
        const matchedCol = headers.indexOf("Matched_Sheet");
        const matchedClIdCol = headers.indexOf("Matched_CL_ID");
        let ptyCol = headers.indexOf("PTY_LOCAL_CODE");
        if (matchedCol < 0 || matchedClIdCol < 0) return `IF(FALSE,CHOOSE({1,2,3},"","",""),CHOOSE({1,2,3},"","",""))`;
        if (ptyCol < 0) ptyCol = headers.length;
        const mc = XLSX.utils.encode_col(matchedCol);
        const mcc = XLSX.utils.encode_col(matchedClIdCol);
        const pc = XLSX.utils.encode_col(ptyCol);
        const lastRow = Math.max(sheetData[name].rows.length, 2);
        const matchedRange = `${q}!${mc}2:${mc}${lastRow}`;
        const matchedClIdRange = `${q}!${mcc}2:${mcc}${lastRow}`;
        const ptyRange = `${q}!${pc}2:${pc}${lastRow}`;
        const sourceSheetRange = `IF(${matchedRange}<>"","${name.replace(/"/g, '""')}","")`;
        // Filter on Matched_Sheet, not on PTY_LOCAL_CODE, so rows with blank
        // local codes remain visible in PSU along with their Matched_CL_ID.
        return `IFERROR(FILTER(CHOOSE({1,2,3},${matchedClIdRange},${ptyRange},${sourceSheetRange}),${matchedRange}=$B$2),{"","",""})`;
      });
      const matchedClExtractFormula = `IF($B$2="","",LET(allRows,VSTACK(${matchedClStacks.join(",")}),FILTER(allRows,CHOOSECOLS(allRows,3)<>"",{"","",""})))`;
      matchedClExtractWs["A4"] = { t: "s", v: "Enter an exact client worksheet name in B2. Results show Matched_CL_ID, PTY_LOCAL_CODE (including blank codes), and the source client sheet from all client sheets pointing to that client." };
      matchedClExtractWs["A4"].s = { alignment: { wrapText: true } };
      matchedClExtractWs["A5"] = { t: "n", f: matchedClExtractFormula };
      matchedClExtractWs["!ref"] = "A1:C5";
      // PSU is the first utility sheet after all client sheets.
      // Reorder the utility sheets to PSU, Local Code Matching, Force.
      XLSX.utils.book_append_sheet(wb, matchedClExtractWs, matchedClExtractSheetName);
      const clientOutputNames = wb.SheetNames.filter(n => !utilityNames.has(n.trim().toUpperCase()));
      wb.SheetNames = [...clientOutputNames, "PSU", "Local Code Matching", "Force"];

      wb.Workbook = wb.Workbook || {};
      wb.Workbook.CalcPr = { calcMode: "auto", fullCalcOnLoad: true, forceFullCalc: true };
      outputWorkbook = wb;

      resultSummary.innerHTML = `
        <div class="stat"><strong>${totalMatches}</strong><br>Total matches</div>
        <div class="stat"><strong>${totalExact}</strong><br>Exact matches</div>
        <div class="stat"><strong>${totalVariation}</strong><br>Variation matches</div>
        <div class="stat"><strong>${comparisonRows}</strong><br>Rows checked</div>
        <p><strong>Completed.</strong> The original workbook was not changed.</p>
        <p><strong>Local Code Matching:</strong> Enter a client sheet name in B2 to extract the source sheet, actual CL_ID, and PTY_LOCAL_CODE.</p>
        <p><strong>PSU:</strong> Enter a client sheet name in B2 to extract Matched_CL_ID, PTY_LOCAL_CODE (including blank codes), and Source_Client_Sheet from all client sheets whose Matched_Sheet points to that client.</p>
        <p><strong>Force:</strong> Permanent mapping register with Client_Unique_ID, PTY_LOCAL_CODE, and Source_Client_Sheet. Existing entries are carried forward on every run; add new mappings as additional rows. Excel 365 or Excel for the web is required for the PSU and Local Code Matching extracts.</p>
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
