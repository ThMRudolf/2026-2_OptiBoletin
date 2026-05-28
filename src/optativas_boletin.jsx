import { useState, useRef } from "react";
import JSZip from "https://esm.sh/jszip@3.10.1";

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const DEPT_ORDER = ["Administración","Bienestar / Formación General","Ciencia Política","Computación","Contabilidad","Derecho","Economía","Estudios Generales","Estudios Internacionales","Ingeniería / Interdepartamental","Ingeniería Industrial y Operaciones","Matemáticas"];
const DEPT_COLORS = {"Administración":"#2563eb","Computación":"#7c3aed","Contabilidad":"#0891b2","Economía":"#059669","Ingeniería Industrial y Operaciones":"#c2410c","Ingeniería / Interdepartamental":"#b45309","Matemáticas":"#be185d","Derecho":"#4338ca","Estudios Generales":"#0f766e","Estudios Internacionales":"#0369a1","Ciencia Política":"#7e22ce","Bienestar / Formación General":"#15803d"};
const DEPT_BG = {"Administración":"#eff6ff","Computación":"#f5f3ff","Contabilidad":"#ecfeff","Economía":"#f0fdf4","Ingeniería Industrial y Operaciones":"#fff7ed","Ingeniería / Interdepartamental":"#fffbeb","Matemáticas":"#fdf2f8","Derecho":"#eef2ff","Estudios Generales":"#f0fdfa","Estudios Internacionales":"#f0f9ff","Ciencia Política":"#faf5ff","Bienestar / Formación General":"#f0fdf4"};
const DEPT_DISPLAY = {"Administración":"ADMINISTRACIÓN","Bienestar / Formación General":"BIENESTAR / FORMACIÓN GENERAL","Ciencia Política":"CIENCIA POLÍTICA","Computación":"COMPUTACIÓN","Contabilidad":"CONTABILIDAD","Derecho":"DERECHO","Economía":"ECONOMÍA","Estudios Generales":"ESTUDIOS GENERALES","Estudios Internacionales":"ESTUDIOS INTERNACIONALES","Ingeniería / Interdepartamental":"INGENIERÍA ELÉCTRICA Y ELECTRÓNICA / INTERDEPARTAMENTAL","Ingeniería Industrial y Operaciones":"INGENIERÍA INDUSTRIAL Y OPERACIONES","Matemáticas":"MATEMÁTICAS"};

// ─── API EXTRACTION ───────────────────────────────────────────────────────────
async function extractCoursesFromDoc(file) {
  const b64 = await new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = e => res(e.target.result.split(",")[1]);
    r.onerror = rej;
    r.readAsDataURL(file);
  });
  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 8000,
      system: `You are a precise data extractor for ITAM, a Mexican university. Extract ALL elective courses from the provided Word document.
Return ONLY a raw JSON array (no markdown fences, no preamble). Each element: {"code":"","title":"","professor":"","prerequisites":"","description":"","department":""}
- code: like "ECO-17201" (normalize dashes)
- title: UPPERCASE as written
- professor: full name(s) as listed
- prerequisites: full text, or "Ninguno" if none
- description: COMPLETE text, do not truncate under any circumstances
- department: infer from prefix — ECO→Economía, ADM→Administración, COM→Computación, CON→Contabilidad, DER→Derecho, EGN→Estudios Generales, EIN→Estudios Internacionales, IIO→Ingeniería Industrial y Operaciones, MAT→Matemáticas, SDI→Ingeniería / Interdepartamental, CEB→Bienestar / Formación General, CSO→Ciencia Política
If no courses found, return [].`,
      messages: [{ role: "user", content: [
        { type: "document", source: { type: "base64", media_type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", data: b64 } },
        { type: "text", text: `Extract all courses from: ${file.name}` }
      ]}]
    })
  });
  if (!resp.ok) { const e = await resp.json().catch(() => ({})); throw new Error(e.error?.message || `API ${resp.status}`); }
  const data = await resp.json();
  const raw = data.content.filter(b => b.type === "text").map(b => b.text).join("");
  return JSON.parse(raw.replace(/```json\s*/g, "").replace(/```\s*/g, "").trim());
}

// ─── DOCX GENERATION ─────────────────────────────────────────────────────────
function esc(s) { return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&apos;"); }
function run(text, bold, size = 20, color = "000000") {
  const b = bold ? "<w:b/><w:bCs/>" : "";
  return `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/>${b}<w:color w:val="${color}"/></w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}
function para(runs, { before=0, after=80, jc="", indent="", shading="" } = {}) {
  return `<w:p><w:pPr>${shading?`<w:shd w:val="clear" w:color="auto" w:fill="${shading}"/>`:""}` +
    `<w:spacing w:before="${before}" w:after="${after}" w:line="276" w:lineRule="auto"/>` +
    `${jc?`<w:jc w:val="${jc}"/>`:""}${indent?`<w:ind w:left="${indent}"/>`:""}` +
    `</w:pPr>${runs}</w:p>`;
}
function courseToXml(c) {
  let out = "";
  out += para(run(`${c.code}   `, true, 20) + run(c.title, true, 20), { before: 200, after: 60 });
  const profLines = (c.professor||"").split("\n");
  out += para(run("PROFESOR: ", true) + run(profLines[0]||""), { after: 40 });
  for (let i = 1; i < profLines.length; i++) out += para(run(profLines[i]), { after: 40, indent: "1260" });
  const preLines = (c.prerequisites||"Ninguno").split("\n");
  out += para(run("PRERREQUISITOS: ", true) + run(preLines[0]), { after: 40 });
  for (let i = 1; i < preLines.length; i++) out += para(run(preLines[i]), { after: i===preLines.length-1?40:20, indent: "1620" });
  out += para(run("DESCRIPCIÓN: ", true) + run(c.description||""), { after: 240, jc: "both" });
  out += `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="2" w:space="1" w:color="CCCCCC"/></w:pBdr><w:spacing w:before="0" w:after="100"/></w:pPr></w:p>`;
  return out;
}

async function generateDocx(selectedCourses, programName) {
  const deptsSorted = DEPT_ORDER.filter(d => selectedCourses.some(c => c.department === d))
    .concat(selectedCourses.map(c => c.department).filter(d => !DEPT_ORDER.includes(d)).filter((v,i,a)=>a.indexOf(v)===i));

  let body = "";
  body += para(run("MATERIAS OPTATIVAS — OTOÑO 2026", true, 28, "1F3864"), { after: 80, jc: "center" });
  body += para(run(programName || "Ingeniería Industrial y de Sistemas (ISI)", false, 22, "555555"), { after: 360, jc: "center" });

  for (const dept of deptsSorted) {
    const courses = selectedCourses.filter(c => c.department === dept);
    if (!courses.length) continue;
    const displayName = DEPT_DISPLAY[dept] || dept.toUpperCase();
    body += `<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="1F3864"/><w:spacing w:before="240" w:after="120"/><w:ind w:left="120"/></w:pPr>${run(displayName, true, 24, "FFFFFF")}</w:p>`;
    courses.forEach(c => { body += courseToXml(c); });
  }

  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable=""><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`;

  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="es-MX"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="0"/></w:pPr></w:style></w:styles>`);
  zip.file("word/document.xml", docXml);
  return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

// ─── STEP INDICATOR ──────────────────────────────────────────────────────────
function StepBar({ step }) {
  const steps = ["Subir archivos", "Seleccionar materias", "Generar documento"];
  return (
    <div style={{ display:"flex", alignItems:"center", gap:0, marginBottom:28 }}>
      {steps.map((s, i) => {
        const n = i + 1;
        const active = step === n;
        const done = step > n;
        return (
          <div key={n} style={{ display:"flex", alignItems:"center", flex: i < steps.length - 1 ? "1" : "0" }}>
            <div style={{ display:"flex", alignItems:"center", gap:8, flexShrink:0 }}>
              <div style={{ width:28, height:28, borderRadius:"50%", background: done?"#15803d": active?"#1F3864":"#e2e8f0", color: (done||active)?"#fff":"#94a3b8", display:"flex", alignItems:"center", justifyContent:"center", fontSize:12, fontFamily:"'DM Mono',monospace", fontWeight:500, flexShrink:0, transition:"all .3s" }}>
                {done ? "✓" : n}
              </div>
              <span style={{ fontSize:13, color: active?"#1F3864": done?"#15803d":"#94a3b8", fontWeight: active?500:400, whiteSpace:"nowrap", transition:"color .3s" }}>{s}</span>
            </div>
            {i < steps.length - 1 && (
              <div style={{ flex:1, height:1, background: done?"#86efac":"#e2e8f0", margin:"0 12px", transition:"background .3s" }} />
            )}
          </div>
        );
      })}
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [step, setStep] = useState(1);
  // Step 1
  const [fileItems, setFileItems] = useState({}); // name → {file,status,courses,error}
  const [allCourses, setAllCourses] = useState([]);
  const [dragOver, setDragOver] = useState(false);
  const [isExtracting, setIsExtracting] = useState(false);
  const fileInputRef = useRef();
  // Step 2
  const [selected, setSelected] = useState(new Set()); // Set of course codes
  const [expandedDepts, setExpandedDepts] = useState(new Set());
  const [searchQuery, setSearchQuery] = useState("");
  const [programName, setProgramName] = useState("Ingeniería Industrial y de Sistemas (ISI)");
  // Step 3
  const [generating, setGenerating] = useState(false);
  const [genDone, setGenDone] = useState(false);

  // ── FILE HELPERS ──
  const addFiles = (flist) => {
    const fresh = {};
    for (const f of flist) {
      if (!f.name.endsWith(".docx")) continue;
      if (!fileItems[f.name]) fresh[f.name] = { file: f, status: "pending", courses: 0, error: null };
    }
    if (Object.keys(fresh).length) setFileItems(prev => ({ ...prev, ...fresh }));
  };

  const removeFile = (name) => setFileItems(prev => { const n = { ...prev }; delete n[name]; return n; });

  const runExtraction = async () => {
    const pending = Object.values(fileItems).filter(x => x.status === "pending");
    if (!pending.length || isExtracting) return;
    setIsExtracting(true);
    const newCourses = [];
    for (const item of pending) {
      setFileItems(prev => ({ ...prev, [item.file.name]: { ...prev[item.file.name], status: "processing" } }));
      try {
        const courses = await extractCoursesFromDoc(item.file);
        newCourses.push(...courses);
        setFileItems(prev => ({ ...prev, [item.file.name]: { ...prev[item.file.name], status: "done", courses: courses.length } }));
      } catch (e) {
        setFileItems(prev => ({ ...prev, [item.file.name]: { ...prev[item.file.name], status: "error", error: e.message } }));
      }
    }
    setAllCourses(prev => {
      const merged = [...prev, ...newCourses];
      const seen = new Set(); return merged.filter(c => { if (seen.has(c.code)) return false; seen.add(c.code); return true; });
    });
    setIsExtracting(false);
  };

  // ── SELECTION HELPERS ──
  const grouped = {};
  for (const c of allCourses) { const d = c.department||"Otro"; if (!grouped[d]) grouped[d] = []; grouped[d].push(c); }
  const deptsSorted = DEPT_ORDER.filter(d => grouped[d]?.length).concat(Object.keys(grouped).filter(d => !DEPT_ORDER.includes(d) && grouped[d]?.length));

  const filteredGrouped = {};
  const q = searchQuery.toLowerCase();
  for (const dept of deptsSorted) {
    const matches = (grouped[dept]||[]).filter(c => !q || c.code?.toLowerCase().includes(q) || c.title?.toLowerCase().includes(q) || c.professor?.toLowerCase().includes(q));
    if (matches.length) filteredGrouped[dept] = matches;
  }
  const filteredDepts = deptsSorted.filter(d => filteredGrouped[d]);

  const toggleCourse = (code) => setSelected(prev => { const n = new Set(prev); n.has(code) ? n.delete(code) : n.add(code); return n; });
  const toggleDept = (dept) => {
    const deptCodes = (filteredGrouped[dept]||[]).map(c => c.code);
    const allSelected = deptCodes.every(code => selected.has(code));
    setSelected(prev => { const n = new Set(prev); deptCodes.forEach(code => allSelected ? n.delete(code) : n.add(code)); return n; });
  };
  const selectAll = () => setSelected(new Set(allCourses.map(c => c.code)));
  const deselectAll = () => setSelected(new Set());
  const toggleExpandDept = (dept) => setExpandedDepts(prev => { const n = new Set(prev); n.has(dept) ? n.delete(dept) : n.add(dept); return n; });

  const selectedCourses = allCourses.filter(c => selected.has(c.code));

  // ── GENERATE ──
  const handleGenerate = async () => {
    setGenerating(true);
    setGenDone(false);
    try {
      const blob = await generateDocx(selectedCourses, programName);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = "Optativas_Otono_2026.docx"; a.click();
      URL.revokeObjectURL(url);
      setGenDone(true);
    } catch(e) { alert("Error: " + e.message); }
    finally { setGenerating(false); }
  };

  const pendingCount = Object.values(fileItems).filter(x => x.status === "pending").length;
  const doneCount = Object.values(fileItems).filter(x => x.status === "done").length;

  // ─ SHARED STYLES ─
  const S = {
    card: { background:"#fff", border:"1px solid #e2e8f0", borderRadius:10, padding:"24px 28px", marginBottom:16 },
    btnPrimary: { background:"#1F3864", color:"#e8ecf4", border:"none", padding:"11px 24px", fontFamily:"'Crimson Pro',serif", fontSize:16, fontWeight:500, cursor:"pointer", borderRadius:6, display:"inline-flex", alignItems:"center", gap:8 },
    btnSuccess: { background:"#15803d", color:"#f0fdf4", border:"none", padding:"11px 24px", fontFamily:"'Crimson Pro',serif", fontSize:16, fontWeight:500, cursor:"pointer", borderRadius:6, display:"inline-flex", alignItems:"center", gap:8 },
    btnSecondary: { background:"none", border:"1px solid #c8c4bc", color:"#555", padding:"9px 16px", fontFamily:"'Crimson Pro',serif", fontSize:15, cursor:"pointer", borderRadius:6 },
    mono: { fontFamily:"'DM Mono',monospace" },
  };

  return (
    <div style={{ minHeight:"100vh", background:"#f5f4f0", fontFamily:"'Crimson Pro',Georgia,serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Crimson+Pro:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;}
        @keyframes spin{to{transform:rotate(360deg)}}
        @keyframes progress{0%{transform:translateX(-100%)}100%{transform:translateX(260%)}}
        @keyframes fadeSlide{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
        .fade{animation:fadeSlide .3s ease both;}
        .drop-z{border:1.5px dashed #b8b3aa;border-radius:8px;padding:28px 20px;text-align:center;cursor:pointer;background:#faf9f7;transition:all .15s;}
        .drop-z:hover,.drop-z.over{border-color:#1F3864;background:#eef2f8;}
        .course-row{display:flex;align-items:flex-start;gap:12px;padding:9px 14px;border-bottom:1px solid #f1f5f9;cursor:pointer;transition:background .1s;}
        .course-row:hover{background:#f8fafc;}
        .course-row:last-child{border-bottom:none;}
        .dept-hdr{display:flex;align-items:center;gap:10px;padding:9px 14px;cursor:pointer;transition:background .1s;border-bottom:1px solid #e2e8f0;}
        .dept-hdr:hover{background:#f8fafc;}
        .chk{width:16px;height:16px;border:1.5px solid #cbd5e1;border-radius:3px;display:flex;align-items:center;justify-content:center;flex-shrink:0;transition:all .15s;cursor:pointer;}
        .chk.on{background:#1F3864;border-color:#1F3864;}
        .chk.partial{background:#e2e8f0;border-color:#94a3b8;}
        .pill{font-family:'DM Mono',monospace;font-size:10px;background:#f1f5f9;color:#475569;padding:2px 6px;border-radius:3px;flex-shrink:0;margin-top:2px;}
        .search-bar{width:100%;font-family:'Crimson Pro',serif;font-size:15px;border:1px solid #e2e8f0;border-radius:6px;padding:9px 14px;background:#f8fafc;color:#1e293b;outline:none;transition:border .15s;}
        .search-bar:focus{border-color:#1F3864;}
        .prog-bar{height:3px;background:#dbeafe;border-radius:2px;overflow:hidden;margin-top:5px;}
        .prog-inner{height:100%;width:55%;background:#1d4ed8;border-radius:2px;animation:progress 1.4s ease-in-out infinite;}
        .btn-p:hover{background:#2c4f8a !important;}
        .btn-s-hover:hover{background:#166534 !important;}
        .tag-chip{display:inline-flex;align-items:center;gap:5px;padding:3px 9px 3px 6px;border-radius:4px;font-size:12px;}
      `}</style>

      {/* ── HEADER ── */}
      <div style={{ background:"#1F3864", padding:"20px 32px 18px", borderBottom:"3px solid #e8b84b" }}>
        <div style={{ maxWidth:900, margin:"0 auto" }}>
          <div style={{ fontFamily:"'DM Mono',monospace", fontSize:10, color:"#8fa4cc", letterSpacing:".12em", marginBottom:5, textTransform:"uppercase" }}>ITAM · División Académica de Ingeniería · Otoño 2026</div>
          <div style={{ color:"#e8ecf4", fontSize:22, fontWeight:500 }}>Generador de Boletín de Optativas</div>
        </div>
      </div>

      <div style={{ maxWidth:900, margin:"0 auto", padding:"28px 20px" }}>
        <StepBar step={step} />

        {/* ══════════════════ STEP 1: UPLOAD ══════════════════ */}
        {step === 1 && (
          <div className="fade">
            <div style={S.card}>
              <div style={{ marginBottom:18 }}>
                <div style={{ fontSize:18, fontWeight:500, color:"#1a1a1a", marginBottom:4 }}>Sube los documentos Word por departamento</div>
                <div style={{ fontSize:13, color:"#64748b" }}>Un .docx por departamento. Claude extrae automáticamente todas las materias.</div>
              </div>

              <div className={`drop-z ${dragOver?"over":""}`}
                onDragOver={e=>{e.preventDefault();setDragOver(true);}}
                onDragLeave={()=>setDragOver(false)}
                onDrop={e=>{e.preventDefault();setDragOver(false);addFiles(e.dataTransfer.files);}}
                onClick={()=>fileInputRef.current?.click()}
                style={{ marginBottom:16 }}
              >
                <div style={{ fontSize:32, marginBottom:8 }}>📂</div>
                <div style={{ fontSize:15, color:"#334155", marginBottom:3 }}>Arrastra los <span style={{ fontFamily:"'DM Mono',monospace", fontSize:12, background:"#ebe9e3", padding:"1px 6px", borderRadius:3 }}>.docx</span> aquí</div>
                <div style={{ fontSize:12, color:"#94a3b8" }}>o haz clic para seleccionar · múltiples archivos a la vez</div>
                <input ref={fileInputRef} type="file" accept=".docx" multiple style={{ display:"none" }} onChange={e=>{addFiles(e.target.files);e.target.value="";}} />
              </div>

              {/* File cards */}
              {Object.values(fileItems).map(item => {
                const cfg = { pending:{bg:"#f8fafc",border:"#e2e8f0",ic:"○",c:"#94a3b8"}, processing:{bg:"#eff6ff",border:"#bfdbfe",ic:"⚙",c:"#1d4ed8"}, done:{bg:"#f0fdf4",border:"#bbf7d0",ic:"✓",c:"#15803d"}, error:{bg:"#fef2f2",border:"#fecaca",ic:"✗",c:"#b91c1c"} }[item.status];
                return (
                  <div key={item.file.name} style={{ background:cfg.bg, border:`1px solid ${cfg.border}`, borderRadius:6, padding:"9px 13px", marginBottom:7 }}>
                    <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                      <span style={{ fontSize:16 }}>📄</span>
                      <div style={{ flex:1, minWidth:0 }}>
                        <div style={{ fontFamily:"'DM Mono',monospace", fontSize:12, color:"#1e293b", fontWeight:500, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{item.file.name}</div>
                        <div style={{ fontSize:11, color:"#94a3b8" }}>{(item.file.size/1024).toFixed(0)} KB</div>
                      </div>
                      {item.status==="processing" && <div style={{ width:13,height:13,border:"2px solid #bfdbfe",borderTopColor:"#1d4ed8",borderRadius:"50%",animation:"spin .7s linear infinite",flexShrink:0 }}/>}
                      <span style={{ fontSize:13, color:cfg.c, fontWeight:600 }}>{cfg.ic}</span>
                      {item.status==="done" && <span style={{ fontFamily:"'DM Mono',monospace",fontSize:10,background:"#1F3864",color:"#d0d9eb",padding:"2px 7px",borderRadius:3 }}>{item.courses} materias</span>}
                      {item.status!=="processing" && <button onClick={()=>removeFile(item.file.name)} style={{ background:"none",border:"none",cursor:"pointer",color:"#94a3b8",fontSize:17,lineHeight:1,padding:"0 2px" }}>×</button>}
                    </div>
                    {item.status==="processing" && <div className="prog-bar"><div className="prog-inner"/></div>}
                    {item.error && <div style={{ fontSize:11,color:"#b91c1c",fontFamily:"'DM Mono',monospace",marginTop:4 }}>⚠ {item.error}</div>}
                  </div>
                );
              })}

              {/* Actions */}
              {Object.keys(fileItems).length > 0 && (
                <div style={{ display:"flex", gap:12, alignItems:"center", marginTop:16, flexWrap:"wrap" }}>
                  <button className="btn-p" style={{ ...S.btnPrimary, opacity:(isExtracting||pendingCount===0)?0.45:1, cursor:(isExtracting||pendingCount===0)?"not-allowed":"pointer" }}
                    onClick={runExtraction} disabled={isExtracting||pendingCount===0}>
                    {isExtracting ? <><div style={{ width:14,height:14,border:"2px solid #8fa4cc",borderTopColor:"#e8ecf4",borderRadius:"50%",animation:"spin .7s linear infinite" }}/> Extrayendo…</> : <><span>⚡</span> Extraer {pendingCount} archivo{pendingCount!==1?"s":""}</>}
                  </button>
                  {allCourses.length > 0 && (
                    <button className="btn-s-hover" style={S.btnSuccess} onClick={()=>setStep(2)}>
                      <span>→</span> Seleccionar materias ({allCourses.length})
                    </button>
                  )}
                  <button style={S.btnSecondary} onClick={()=>{setFileItems({});setAllCourses([]);setSelected(new Set());}}>Limpiar</button>
                  {doneCount > 0 && <span style={{ fontFamily:"'DM Mono',monospace", fontSize:11, color:"#94a3b8", marginLeft:"auto" }}>{doneCount}/{Object.keys(fileItems).length} procesados</span>}
                </div>
              )}
            </div>
          </div>
        )}

        {/* ══════════════════ STEP 2: SELECT ══════════════════ */}
        {step === 2 && (
          <div className="fade">
            {/* Toolbar */}
            <div style={{ display:"flex", gap:12, alignItems:"center", marginBottom:16, flexWrap:"wrap" }}>
              <input className="search-bar" style={{ flex:1, minWidth:200 }} placeholder="Buscar por código, título o profesor…" value={searchQuery} onChange={e=>setSearchQuery(e.target.value)} />
              <button style={{ ...S.btnSecondary, fontSize:13, padding:"8px 14px", whiteSpace:"nowrap" }} onClick={selectAll}>Seleccionar todo</button>
              <button style={{ ...S.btnSecondary, fontSize:13, padding:"8px 14px", whiteSpace:"nowrap" }} onClick={deselectAll}>Deseleccionar</button>
              <div style={{ fontFamily:"'DM Mono',monospace", fontSize:11, color:"#64748b", whiteSpace:"nowrap" }}>
                <span style={{ color:"#1F3864", fontWeight:500 }}>{selected.size}</span> / {allCourses.length} seleccionadas
              </div>
            </div>

            {/* Course tree with checkboxes */}
            <div style={{ background:"#fff", border:"1px solid #e2e8f0", borderRadius:10, overflow:"hidden", marginBottom:16, maxHeight:520, overflowY:"auto" }}>
              {filteredDepts.map((dept, di) => {
                const dCourses = filteredGrouped[dept] || [];
                const selectedInDept = dCourses.filter(c => selected.has(c.code)).length;
                const allIn = selectedInDept === dCourses.length;
                const partialIn = selectedInDept > 0 && !allIn;
                const isOpen = expandedDepts.has(dept);
                const color = DEPT_COLORS[dept] || "#475569";
                const bg = DEPT_BG[dept] || "#f8fafc";
                return (
                  <div key={dept} style={{ borderBottom: di < filteredDepts.length-1 ? "1px solid #e2e8f0":"none" }}>
                    <div className="dept-hdr" style={{ background: isOpen ? bg+"cc" : "#fff" }}>
                      {/* Dept checkbox */}
                      <div className={`chk ${allIn?"on":partialIn?"partial":""}`} onClick={()=>toggleDept(dept)} style={{ marginTop:0 }}>
                        {allIn && <span style={{ color:"#fff", fontSize:10 }}>✓</span>}
                        {partialIn && <span style={{ color:"#64748b", fontSize:10, fontWeight:700 }}>–</span>}
                      </div>
                      {/* Expand toggle */}
                      <div style={{ flex:1, display:"flex", alignItems:"center", gap:10 }} onClick={()=>toggleExpandDept(dept)}>
                        <span style={{ color, fontSize:11, transition:"transform .2s", display:"inline-block", transform:isOpen?"rotate(90deg)":"none" }}>▶</span>
                        <span style={{ fontSize:14, fontWeight:500, color:"#1e293b" }}>{dept}</span>
                        <span style={{ fontFamily:"'DM Mono',monospace", fontSize:10, background:bg, color, border:`1px solid ${color}22`, padding:"1px 7px", borderRadius:3 }}>
                          {selectedInDept}/{dCourses.length}
                        </span>
                      </div>
                    </div>
                    {isOpen && dCourses.map(c => {
                      const isSel = selected.has(c.code);
                      return (
                        <div key={c.code} className="course-row" style={{ background: isSel ? bg+"55" : "transparent" }} onClick={()=>toggleCourse(c.code)}>
                          <div className={`chk ${isSel?"on":""}`} style={{ marginTop:2 }}>
                            {isSel && <span style={{ color:"#fff", fontSize:10 }}>✓</span>}
                          </div>
                          <span className="pill">{c.code}</span>
                          <div style={{ flex:1 }}>
                            <div style={{ fontSize:13, fontWeight: isSel?500:400, color: isSel?"#1e293b":"#374151", lineHeight:1.35 }}>{c.title}</div>
                            <div style={{ fontSize:11, color:"#64748b", marginTop:2, fontFamily:"'DM Mono',monospace" }}>{c.professor}</div>
                            {c.prerequisites && c.prerequisites!=="Ninguno" && (
                              <div style={{ fontSize:11, color:"#94a3b8", marginTop:1 }}>Prereq: {c.prerequisites.length>70?c.prerequisites.slice(0,70)+"…":c.prerequisites}</div>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                );
              })}
              {filteredDepts.length === 0 && (
                <div style={{ padding:"32px", textAlign:"center", color:"#94a3b8", fontSize:14 }}>No se encontraron materias para "{searchQuery}"</div>
              )}
            </div>

            {/* Selected preview chips */}
            {selected.size > 0 && (
              <div style={{ background:"#f8fafc", border:"1px solid #e2e8f0", borderRadius:8, padding:"12px 16px", marginBottom:16 }}>
                <div style={{ fontSize:12, color:"#64748b", fontFamily:"'DM Mono',monospace", marginBottom:8 }}>SELECCIONADAS ({selected.size})</div>
                <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
                  {selectedCourses.slice(0,24).map(c => (
                    <span key={c.code} className="tag-chip" style={{ background: DEPT_BG[c.department]||"#f1f5f9", color: DEPT_COLORS[c.department]||"#475569", border:`1px solid ${(DEPT_COLORS[c.department]||"#cbd5e1")}22` }}>
                      <span style={{ fontFamily:"'DM Mono',monospace", fontSize:10 }}>{c.code}</span>
                      <button onClick={()=>toggleCourse(c.code)} style={{ background:"none",border:"none",cursor:"pointer",color:"inherit",fontSize:13,lineHeight:1,opacity:.6,padding:0 }}>×</button>
                    </span>
                  ))}
                  {selected.size > 24 && <span style={{ fontSize:12, color:"#94a3b8", padding:"3px 0" }}>+{selected.size-24} más</span>}
                </div>
              </div>
            )}

            <div style={{ display:"flex", gap:12, alignItems:"center", flexWrap:"wrap" }}>
              <button style={S.btnSecondary} onClick={()=>setStep(1)}>← Volver</button>
              <button className="btn-s-hover" style={{ ...S.btnSuccess, opacity:selected.size===0?.45:1, cursor:selected.size===0?"not-allowed":"pointer" }}
                onClick={()=>{ if(selected.size>0) setStep(3); }} disabled={selected.size===0}>
                <span>→</span> Continuar con {selected.size} materia{selected.size!==1?"s":""}
              </button>
            </div>
          </div>
        )}

        {/* ══════════════════ STEP 3: GENERATE ══════════════════ */}
        {step === 3 && (
          <div className="fade">
            <div style={S.card}>
              <div style={{ fontSize:18, fontWeight:500, color:"#1a1a1a", marginBottom:16 }}>Configurar y generar documento</div>

              {/* Program name */}
              <div style={{ marginBottom:20 }}>
                <label style={{ fontFamily:"'DM Mono',monospace", fontSize:11, color:"#64748b", letterSpacing:".04em", display:"block", marginBottom:5 }}>NOMBRE DEL PROGRAMA</label>
                <input style={{ fontFamily:"'Crimson Pro',serif", fontSize:15, border:"1px solid #e2e8f0", borderRadius:6, padding:"9px 12px", width:"100%", background:"#f8fafc", color:"#1e293b", outline:"none" }}
                  value={programName} onChange={e=>setProgramName(e.target.value)} />
              </div>

              {/* Summary by dept */}
              <div style={{ marginBottom:20 }}>
                <div style={{ fontFamily:"'DM Mono',monospace", fontSize:11, color:"#64748b", letterSpacing:".04em", marginBottom:10 }}>RESUMEN — {selected.size} MATERIAS SELECCIONADAS</div>
                <div style={{ display:"flex", flexWrap:"wrap", gap:8 }}>
                  {deptsSorted.filter(d => selectedCourses.some(c=>c.department===d)).map(d => {
                    const cnt = selectedCourses.filter(c=>c.department===d).length;
                    const color = DEPT_COLORS[d]||"#475569";
                    const bg = DEPT_BG[d]||"#f8fafc";
                    return (
                      <div key={d} style={{ display:"flex", alignItems:"center", gap:7, background:bg, border:`1px solid ${color}33`, borderRadius:5, padding:"5px 11px 5px 8px" }}>
                        <span style={{ fontFamily:"'DM Mono',monospace", fontSize:11, background:color, color:"#fff", padding:"1px 6px", borderRadius:3, fontWeight:500 }}>{cnt}</span>
                        <span style={{ fontSize:13, color }}>{d}</span>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Document preview snippet */}
              <div style={{ border:"1px solid #e2e8f0", borderRadius:8, overflow:"hidden", marginBottom:24 }}>
                <div style={{ background:"#1F3864", padding:"8px 16px" }}>
                  <div style={{ fontFamily:"'DM Mono',monospace", fontSize:10, color:"#8fa4cc", letterSpacing:".08em" }}>VISTA PREVIA DEL DOCUMENTO</div>
                </div>
                <div style={{ padding:"16px", maxHeight:240, overflowY:"auto", background:"#fafafa", fontFamily:"Arial, sans-serif" }}>
                  <div style={{ textAlign:"center", marginBottom:8 }}>
                    <div style={{ fontSize:14, fontWeight:700, color:"#1F3864" }}>MATERIAS OPTATIVAS — OTOÑO 2026</div>
                    <div style={{ fontSize:11, color:"#555", marginTop:2 }}>{programName}</div>
                  </div>
                  {deptsSorted.filter(d => selectedCourses.some(c=>c.department===d)).map(d => (
                    <div key={d} style={{ marginBottom:6 }}>
                      <div style={{ background:"#1F3864", color:"#fff", padding:"3px 8px", fontSize:10, fontWeight:700, letterSpacing:".04em", marginBottom:3 }}>{DEPT_DISPLAY[d]||d.toUpperCase()}</div>
                      {selectedCourses.filter(c=>c.department===d).map(c => (
                        <div key={c.code} style={{ fontSize:10, marginBottom:4, paddingLeft:4, borderLeft:"2px solid #e2e8f0" }}>
                          <span style={{ fontWeight:700 }}>{c.code}</span> <span style={{ fontWeight:700 }}>{c.title}</span>
                          <div style={{ color:"#555" }}><strong>PROFESOR:</strong> {c.professor}</div>
                        </div>
                      ))}
                    </div>
                  ))}
                </div>
              </div>

              <div style={{ display:"flex", gap:12, alignItems:"center", flexWrap:"wrap" }}>
                <button style={S.btnSecondary} onClick={()=>{setStep(2);setGenDone(false);}}>← Editar selección</button>
                <button className="btn-s-hover" style={{ ...S.btnSuccess, opacity:generating?.45:1, cursor:generating?"not-allowed":"pointer" }}
                  onClick={handleGenerate} disabled={generating}>
                  {generating
                    ? <><div style={{ width:14,height:14,border:"2px solid #86efac",borderTopColor:"#f0fdf4",borderRadius:"50%",animation:"spin .7s linear infinite" }}/> Generando…</>
                    : <><span>↓</span> Descargar Word (.docx)</>
                  }
                </button>
                {genDone && (
                  <div style={{ display:"flex", alignItems:"center", gap:7, background:"#f0fdf4", border:"1px solid #86efac", borderRadius:5, padding:"8px 14px", fontSize:13, color:"#15803d" }}>
                    <span>✓</span> Documento descargado correctamente
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
