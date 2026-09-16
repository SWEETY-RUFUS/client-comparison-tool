let inputWB=null,resultWB=null,fileNameValue="";
const $=id=>document.getElementById(id);
const fileInput=$("fileInput"),drop=$("dropZone");
fileInput.onchange=e=>e.target.files[0]&&loadFile(e.target.files[0]);
["dragenter","dragover"].forEach(n=>drop.addEventListener(n,e=>{e.preventDefault();drop.classList.add("dragover")}));
["dragleave","drop"].forEach(n=>drop.addEventListener(n,e=>{e.preventDefault();drop.classList.remove("dragover")}));
drop.addEventListener("drop",e=>e.dataTransfer.files[0]&&loadFile(e.dataTransfer.files[0]));
$("compareButton").onclick=compare;
$("downloadButton").onclick=download;
$("resetButton").onclick=reset;

function norm(v){return v==null?"":String(v).trim().toUpperCase().replace(/\s+/g," ")}
function key(a,b){return norm(a)+" || "+norm(b)}
function col(headers,name){return headers.findIndex(x=>norm(x)===norm(name))}
function msg(text,type){$("status").classList.toggle("hidden",type!=="status");$("error").classList.toggle("hidden",type!=="error");(type==="status"?$("status"):$("error")).textContent=text}

async function loadFile(file){
try{
if(!/\.xlsx?$/.test(file.name.toLowerCase()))throw Error("Please choose an Excel .xlsx or .xls file.");
fileNameValue=file.name;
inputWB=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true});
if(inputWB.SheetNames.length<2)throw Error("At least two client sheets are required.");
for(const n of inputWB.SheetNames){
const rows=XLSX.utils.sheet_to_json(inputWB.Sheets[n],{header:1,defval:""});
if(!rows.length)throw Error(`Sheet "${n}" is empty.`);
const h=rows[0].map(x=>String(x??"").trim());
if(col(h,"CL_STORE_NAME")<0||col(h,"CL_PA_LINE1")<0)throw Error(`Sheet "${n}" must contain CL_STORE_NAME and CL_PA_LINE1.`);
}
$("fileName").textContent=file.name;$("sheetCount").textContent=inputWB.SheetNames.length;
$("rowCount").textContent="—";$("matchCount").textContent="—";$("noMatchCount").textContent="—";
renderSheets();$("workbookCard").classList.remove("hidden");$("downloadArea").classList.add("hidden");msg("Workbook loaded. Check the sheet order, then click Compare Sheets.","status");
}catch(e){msg(e.message||String(e),"error")}
}

function renderSheets(){
$("sheetList").innerHTML="";
inputWB.SheetNames.forEach((n,i)=>{
let action=i===0?"No comparison — oldest sheet":`Compare with: ${inputWB.SheetNames.slice(0,i).join(" → ")}`;
let d=document.createElement("div");d.className="sheet-row";
d.innerHTML=`<div class="sheet-number">${i+1}</div><div class="sheet-name">${esc(n)}</div><div class="sheet-action">${esc(action)}</div>`;
$("sheetList").appendChild(d);
});
}

function compare(){
try{
$("compareButton").disabled=true;$("compareButton").textContent="Comparing...";
let out=XLSX.utils.book_new(),rowsChecked=0,matches=0,noMatches=0;
inputWB.SheetNames.forEach((name,i)=>{
const srcRows=XLSX.utils.sheet_to_json(inputWB.Sheets[name],{header:1,defval:""});
if(i===0){XLSX.utils.book_append_sheet(out,XLSX.utils.aoa_to_sheet(srcRows),name);return;}
let rows=srcRows.map(r=>r.slice()),headers=rows[0].map(x=>String(x??"").trim());
let bi=col(headers,"CL_STORE_NAME"),ai=col(headers,"CL_PA_LINE1"),ri=col(headers,"Matched_Sheet");
if(ri<0){ri=rows[0].length;rows[0].push("Matched_Sheet")}else{rows[0][ri]="Matched_Sheet";for(let r=1;r<rows.length;r++)rows[r][ri]="";}
const lookups=[];
for(let p=0;p<i;p++){
const pn=inputWB.SheetNames[p],pr=XLSX.utils.sheet_to_json(inputWB.Sheets[pn],{header:1,defval:""}),ph=pr[0].map(x=>String(x??"").trim()),pbi=col(ph,"CL_STORE_NAME"),pai=col(ph,"CL_PA_LINE1"),set=new Set();
for(let r=1;r<pr.length;r++){let k=key(pr[r][pbi],pr[r][pai]);if(k!==" || ")set.add(k)}
lookups.push({name:pn,set});
}
for(let r=1;r<rows.length;r++){
let k=key(rows[r][bi],rows[r][ai]),match="";
for(const old of lookups){if(k!==" || "&&old.set.has(k)){match=old.name;break}}
rows[r][ri]=match;rowsChecked++;match?matches++:noMatches++;
}
XLSX.utils.book_append_sheet(out,XLSX.utils.aoa_to_sheet(rows),name);
});
resultWB=out;$("rowCount").textContent=rowsChecked.toLocaleString();$("matchCount").textContent=matches.toLocaleString();$("noMatchCount").textContent=noMatches.toLocaleString();
$("summary").textContent=`${matches.toLocaleString()} matches found and ${noMatches.toLocaleString()} rows with no match.`;
$("downloadArea").classList.remove("hidden");msg("Comparison completed successfully.","status");
$("compareButton").disabled=false;$("compareButton").textContent="Compare Sheets";
}catch(e){$("compareButton").disabled=false;$("compareButton").textContent="Compare Sheets";msg(e.message||String(e),"error")}
}

function download(){if(!resultWB)return;XLSX.writeFile(resultWB,fileNameValue.replace(/\.[^.]+$/,"")+"_Comparison_Result.xlsx")}
function reset(){inputWB=null;resultWB=null;fileInput.value="";$("workbookCard").classList.add("hidden");$("downloadArea").classList.add("hidden");$("status").classList.add("hidden");$("error").classList.add("hidden")}
function esc(v){return String(v).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;")}
