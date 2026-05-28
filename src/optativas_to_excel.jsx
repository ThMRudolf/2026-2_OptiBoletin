import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const DEPT_ORDER = [
  "Administración","Bienestar / Formación General","Ciencia Política",
  "Computación","Contabilidad","Derecho","Economía","Estudios Generales",
  "Estudios Internacionales","Ingeniería / Interdepartamental",
  "Ingeniería Industrial y Operaciones","Matemáticas",
];

const DEPT_COLORS = {
  "Administración": "#2563eb",
  "Computación": "#7c3aed",
  "Contabilidad": "#0891b2",
  "Economía": "#059669",
  "Ingeniería Industrial y Operaciones": "#c2410c",
  "Ingeniería / Interdepartamental": "#b45309",
  "Matemáticas": "#be185d",
  "Derecho": "#4338ca",
  "Estudios Generales": "#0f766e",
  "Estudios Internacionales": "#0369a1",
  "Ciencia Política": "#7e22ce",
  "Bienestar / Formación General": "#15803d",
};

const DEPT_BG = {
  "Administración": "#eff6ff",
  "Computación": "#f5f3ff",
  "Contabilidad": "#ecfeff",
  "Economía": "#f0fdf4",
  "Ingeniería Industrial y Operaciones": "#fff7ed",
  "Ingeniería / Interdepartamental": "#fffbeb",
  "Matemáticas": "#fdf2f8",
  "Derecho": "#eef2ff",
  "Estudios Generales": "#f0fdfa",
  "Estudios Internacionales": "#f0f9ff",
  "Ciencia Política": "#faf5ff",
  "Bienestar / Formación General": "#f0fdf4",
};

// ─── MAMMOTH-LIKE TEXT EXTRACTOR (using FileReader + ZIP) ─────────────────────
async function extractDocxText(file) {
  // Use the Anthropic API to do both extraction and structuring in one shot
  // We'll base64 encode the file and send as document
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ─── EXTRACT COURSES VIA ANTHROPIC API ───────────────────────────────────────
async function extractCoursesFromDoc(file) {
  const dataUrl = await extractDocxText(file);
  const base64 = dataUrl.split(",")[1];

  const systemPrompt = `You are a precise data extractor for ITAM, a Mexican university. Extract ALL elective courses from the provided Word document.

For each course return a JSON object with exactly these fields:
- "code": course code like "ECO-17201" or "ADM-11002" (normalize spacing around dashes)
- "title": course title in UPPERCASE as written
- "professor": professor name(s) as listed
- "prerequisites": prerequisites text, or "Ninguno" if none stated
- "description": the full course description paragraph(s) — include all text, do not truncate
- "department": infer from the code prefix using this map: ECO→Economía, ADM→Administración, COM→Computación, CON→Contabilidad, DER→Derecho, EGN→Estudios Generales, EIN→Estudios Internacionales, IIO→Ing. Industrial y Operaciones, MAT→Matemáticas, SDI→Ingeniería/Interdepartamental, CEB→Bienestar/Formación General, CSO→Ciencia Política

Return ONLY a raw JSON array — no markdown fences, no preamble, no explanation. If no courses found, return [].`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 8000,
      system: systemPrompt,
      messages: [{
        role: "user",
        content: [{
          type: "document",
          source: { type: "base64", media_type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", data: base64 }
        }, {
          type: "text",
          text: `Extract all courses from this file: ${file.name}`
        }]
      }]
    })
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err.error?.message || `API error ${response.status}`);
  }

  const data = await response.json();
  const raw = data.content.filter(b => b.type === "text").map(b => b.text).join("");
  const clean = raw.replace(/```json\s*/g, "").replace(/```\s*/g, "").trim();
  return JSON.parse(clean);
}

// ─── EXCEL GENERATION ────────────────────────────────────────────────────────
function generateExcel(allCourses) {
  const sorted = [...allCourses].sort((a, b) => {
    const di = (d) => { const i = DEPT_ORDER.indexOf(d); return i === -1 ? 99 : i; };
    const dd = di(a.department) - di(b.department);
    return dd !== 0 ? dd : (a.code || "").localeCompare(b.code || "");
  });

  const wb = XLSX.utils.book_new();
  const headers = ["Departamento", "Clave", "Título", "Profesor(a)", "Prerrequisitos", "Descripción"];
  const wsData = [headers, ...sorted.map(c => [c.department||"", c.code||"", c.title||"", c.professor||"", c.prerequisites||"", c.description||""])];

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws["!cols"] = [22, 16, 44, 34, 48, 80].map(w => ({ wch: w }));

  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[addr]) continue;
      ws[addr].s = {
        font: { name: "Arial", sz: R === 0 ? 11 : 10, bold: R === 0, color: { rgb: R === 0 ? "FFFFFF" : "000000" } },
        fill: { fgColor: { rgb: R === 0 ? "1F3864" : R % 2 === 0 ? "EBF3FA" : "FFFFFF" }, patternType: "solid" },
        alignment: { wrapText: true, vertical: "top", horizontal: R === 0 ? "center" : "left" },
        border: { top:{style:"thin",color:{rgb:"B8CCE4"}}, bottom:{style:"thin",color:{rgb:"B8CCE4"}}, left:{style:"thin",color:{rgb:"B8CCE4"}}, right:{style:"thin",color:{rgb:"B8CCE4"}} }
      };
    }
  }

  XLSX.utils.book_append_sheet(wb, ws, "Optativas Otoño 2026");
  XLSX.writeFile(wb, "Optativas_Otono_2026.xlsx");
}

// ─── COMPONENTS ───────────────────────────────────────────────────────────────
function DeptBadge({ dept }) {
  const color = DEPT_COLORS[dept] || "#475569";
  const bg = DEPT_BG[dept] || "#f8fafc";
  return (
    <span style={{ fontFamily:"'DM Mono',monospace", fontSize:10, background:bg, color, border:`1px solid ${color}22`, padding:"2px 7px", borderRadius:3, whiteSpace:"nowrap", fontWeight:500 }}>
      {dept}
    </span>
  );
}

function FileItem({ file, status, courses, error, onRemove }) {
  const statusConfig = {
    pending:    { icon: "⏳", color: "#64748b", bg: "#f8fafc",  border: "#e2e8f0" },
    processing: { icon: "⚙",  color: "#1d4ed8", bg: "#eff6ff",  border: "#bfdbfe" },
    done:       { icon: "✓",  color: "#15803d", bg: "#f0fdf4",  border: "#bbf7d0" },
    error:      { icon: "✗",  color: "#b91c1c", bg: "#fef2f2",  border: "#fecaca" },
  };
  const cfg = statusConfig[status] || statusConfig.pending;

  return (
    <div style={{ background: cfg.bg, border: `1px solid ${cfg.border}`, borderRadius: 6, padding: "10px 14px", marginBottom: 8, transition: "all .2s" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <span style={{ fontSize: 18 }}>📄</span>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontFamily:"'DM Mono',monospace", fontSize: 12, color: "#1e293b", fontWeight: 500, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{file.name}</div>
          <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 1 }}>{(file.size / 1024).toFixed(0)} KB</div>
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
          {status === "processing" && (
            <div style={{ width:14, height:14, border:"2px solid #bfdbfe", borderTopColor:"#1d4ed8", borderRadius:"50%", animation:"spin .7s linear infinite", flexShrink:0 }} />
          )}
          <span style={{ fontSize: 14, color: cfg.color, fontWeight: 600 }}>{cfg.icon}</span>
          {status === "done" && courses > 0 && (
            <span style={{ fontFamily:"'DM Mono',monospace", fontSize:10, background:"#1F3864", color:"#d0d9eb", padding:"2px 7px", borderRadius:3 }}>{courses} materias</span>
          )}
          {onRemove && status !== "processing" && (
            <button onClick={onRemove} style={{ background:"none", border:"none", cursor:"pointer", color:"#94a3b8", fontSize:16, lineHeight:1, padding:"0 2px" }} title="Eliminar">×</button>
          )}
        </div>
      </div>
      {status === "processing" && (
        <div style={{ marginTop: 6, height: 3, background: "#dbeafe", borderRadius: 2, overflow:"hidden" }}>
          <div style={{ height:"100%", background:"#1d4ed8", width:"60%", borderRadius:2, animation:"progress 1.5s ease-in-out infinite" }} />
        </div>
      )}
      {error && <div style={{ marginTop:6, fontSize:11, color:"#b91c1c", fontFamily:"'DM Mono',monospace" }}>⚠ {error}</div>}
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [files, setFiles] = useState([]); // [{file, status, courses, error}]
  const [allCourses, setAllCourses] = useState([]);
  const [dragOver, setDragOver] = useState(false);
  const [activeTab, setActiveTab] = useState("upload"); // upload | preview
  const [expandedDept, setExpandedDept] = useState(null);
  const inputRef = useRef();

  const addFiles = useCallback((newFiles) => {
    const docxFiles = [...newFiles].filter(f => f.name.endsWith(".docx"));
    setFiles(prev => {
      const existing = new Set(prev.map(x => x.file.name));
      const fresh = docxFiles.filter(f => !existing.has(f.name)).map(f => ({ file: f, status: "pending", courses: 0, error: null }));
      return [...prev, ...fresh];
    });
  }, []);

  const removeFile = useCallback((name) => {
    setFiles(prev => prev.filter(x => x.file.name !== name));
  }, []);

  const clearAll = () => { setFiles([]); setAllCourses([]); setActiveTab("upload"); };

  const processAll = async () => {
    const pending = files.filter(f => f.status === "pending");
    if (!pending.length) return;

    const newCourses = [];

    for (const item of pending) {
      setFiles(prev => prev.map(x => x.file.name === item.file.name ? { ...x, status: "processing" } : x));
      try {
        const courses = await extractCoursesFromDoc(item.file);
        newCourses.push(...courses);
        setFiles(prev => prev.map(x => x.file.name === item.file.name ? { ...x, status: "done", courses: courses.length } : x));
      } catch (err) {
        setFiles(prev => prev.map(x => x.file.name === item.file.name ? { ...x, status: "error", error: err.message } : x));
      }
    }

    setAllCourses(prev => {
      const merged = [...prev, ...newCourses];
      // Deduplicate by code
      const seen = new Set();
      return merged.filter(c => { if (seen.has(c.code)) return false; seen.add(c.code); return true; });
    });

    setActiveTab("preview");
  };

  // Group courses for preview
  const grouped = {};
  for (const c of allCourses) {
    const d = c.department || "Sin departamento";
    if (!grouped[d]) grouped[d] = [];
    grouped[d].push(c);
  }
  const deptsSorted = DEPT_ORDER.filter(d => grouped[d]?.length).concat(Object.keys(grouped).filter(d => !DEPT_ORDER.includes(d) && grouped[d]?.length));

  const pendingCount = files.filter(f => f.status === "pending").length;
  const processingCount = files.filter(f => f.status === "processing").length;
  const doneCount = files.filter(f => f.status === "done").length;
  const isProcessing = processingCount > 0;

  return (
    <div style={{ minHeight:"100vh", background:"#f5f4f0", fontFamily:"'Crimson Pro', Georgia, serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Crimson+Pro:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
        * { box-sizing:border-box; margin:0; padding:0; }
        @keyframes spin { to { transform:rotate(360deg) } }
        @keyframes progress { 0%{transform:translateX(-100%)} 100%{transform:translateX(250%)} }
        @keyframes fadeIn { from{opacity:0;transform:translateY(6px)} to{opacity:1;transform:none} }
        .fade-in { animation: fadeIn .35s ease both; }
        .tab { background:none; border:none; cursor:pointer; font-family:'Crimson Pro',serif; font-size:15px; padding:8px 18px; border-radius:4px 4px 0 0; border-bottom:2px solid transparent; transition:all .15s; color:#64748b; }
        .tab.active { color:#1F3864; border-bottom-color:#1F3864; font-weight:500; }
        .tab:hover:not(.active) { background:#f1f5f9; color:#334155; }
        .dept-row { cursor:pointer; padding:8px 14px; display:flex; align-items:center; gap:10px; border-bottom:1px solid #e2e8f0; user-select:none; }
        .dept-row:hover { background:#f8fafc; }
        .course-item { padding:8px 14px 8px 42px; border-bottom:1px solid #f1f5f9; display:flex; align-items:flex-start; gap:10px; }
        .course-item:last-child { border-bottom:none; }
        .code-mono { font-family:'DM Mono',monospace; font-size:11px; background:#f1f5f9; color:#475569; padding:2px 7px; border-radius:3px; flex-shrink:0; margin-top:1px; }
        .btn-primary { background:#1F3864; color:#e8ecf4; border:none; padding:11px 26px; font-family:'Crimson Pro',serif; font-size:16px; font-weight:500; cursor:pointer; border-radius:5px; transition:background .12s, transform .1s; display:inline-flex; align-items:center; gap:8px; }
        .btn-primary:hover:not(:disabled) { background:#2c4f8a; }
        .btn-primary:active:not(:disabled) { transform:scale(.98); }
        .btn-primary:disabled { opacity:.4; cursor:not-allowed; }
        .btn-secondary { background:none; border:1px solid #c8c4bc; color:#555; padding:10px 18px; font-family:'Crimson Pro',serif; font-size:15px; cursor:pointer; border-radius:5px; transition:background .1s; }
        .btn-secondary:hover { background:#eae8e2; }
        .btn-success { background:#15803d; color:#f0fdf4; border:none; padding:11px 26px; font-family:'Crimson Pro',serif; font-size:16px; font-weight:500; cursor:pointer; border-radius:5px; transition:background .12s; display:inline-flex; align-items:center; gap:8px; }
        .btn-success:hover { background:#166534; }
        .drop-zone { border:1.5px dashed #b8b3aa; border-radius:8px; padding:32px 20px; text-align:center; cursor:pointer; background:#faf9f7; transition:all .15s; }
        .drop-zone:hover, .drop-zone.over { border-color:#1F3864; background:#eef2f8; }
      `}</style>

      {/* Header */}
      <div style={{ background:"#1F3864", padding:"22px 36px 20px", borderBottom:"3px solid #e8b84b" }}>
        <div style={{ maxWidth:860, margin:"0 auto" }}>
          <div style={{ fontFamily:"'DM Mono',monospace", fontSize:10, color:"#8fa4cc", letterSpacing:".12em", marginBottom:5, textTransform:"uppercase" }}>ITAM · División Académica de Ingeniería</div>
          <div style={{ color:"#e8ecf4", fontSize:24, fontWeight:500 }}>Optativas → Excel</div>
          <div style={{ color:"#7f9abf", fontSize:13, marginTop:3 }}>Sube los documentos Word de optativas por departamento y genera el consolidado</div>
        </div>
      </div>

      <div style={{ maxWidth:860, margin:"0 auto", padding:"28px 20px" }}>

        {/* Tabs */}
        <div style={{ display:"flex", gap:2, borderBottom:"1px solid #e2e8f0", marginBottom:20 }}>
          <button className={`tab ${activeTab==="upload"?"active":""}`} onClick={() => setActiveTab("upload")}>
            📁 Subir documentos {files.length > 0 && `(${files.length})`}
          </button>
          <button className={`tab ${activeTab==="preview"?"active":""}`} onClick={() => setActiveTab("preview")} disabled={!allCourses.length}>
            📊 Vista previa {allCourses.length > 0 && `(${allCourses.length})`}
          </button>
        </div>

        {/* ── UPLOAD TAB ── */}
        {activeTab === "upload" && (
          <div className="fade-in">
            {/* Drop zone */}
            <div
              className={`drop-zone ${dragOver?"over":""}`}
              onDragOver={e => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={e => { e.preventDefault(); setDragOver(false); addFiles(e.dataTransfer.files); }}
              onClick={() => inputRef.current?.click()}
              style={{ marginBottom: 20 }}
            >
              <div style={{ fontSize:36, marginBottom:10 }}>📂</div>
              <div style={{ fontSize:16, color:"#334155", marginBottom:4 }}>Arrastra los archivos <span style={{ fontFamily:"'DM Mono',monospace", fontSize:13, background:"#ebe9e3", padding:"2px 6px", borderRadius:3 }}>.docx</span> aquí</div>
              <div style={{ fontSize:13, color:"#94a3b8" }}>Un archivo por departamento — puedes agregar varios a la vez</div>
              <input ref={inputRef} type="file" accept=".docx" multiple style={{ display:"none" }} onChange={e => { addFiles(e.target.files); e.target.value=""; }} />
            </div>

            {/* File list */}
            {files.length > 0 && (
              <div style={{ marginBottom: 20 }}>
                {files.map(item => (
                  <FileItem
                    key={item.file.name}
                    file={item.file}
                    status={item.status}
                    courses={item.courses}
                    error={item.error}
                    onRemove={item.status !== "processing" ? () => removeFile(item.file.name) : null}
                  />
                ))}
              </div>
            )}

            {/* Actions */}
            {files.length > 0 && (
              <div style={{ display:"flex", gap:12, alignItems:"center", flexWrap:"wrap" }}>
                <button
                  className="btn-primary"
                  onClick={processAll}
                  disabled={isProcessing || pendingCount === 0}
                >
                  {isProcessing
                    ? <><div style={{ width:15, height:15, border:"2px solid #8fa4cc", borderTopColor:"#e8ecf4", borderRadius:"50%", animation:"spin .7s linear infinite" }} /> Extrayendo…</>
                    : <><span>⚡</span> Extraer materias ({pendingCount} archivo{pendingCount!==1?"s":""})</>
                  }
                </button>
                {allCourses.length > 0 && (
                  <button className="btn-success" onClick={() => setActiveTab("preview")}>
                    <span>👁</span> Ver {allCourses.length} materias extraídas
                  </button>
                )}
                <button className="btn-secondary" onClick={clearAll}>Limpiar todo</button>
                {doneCount > 0 && (
                  <span style={{ fontFamily:"'DM Mono',monospace", fontSize:11, color:"#64748b", marginLeft:"auto" }}>
                    {doneCount}/{files.length} archivos procesados
                  </span>
                )}
              </div>
            )}

            {files.length === 0 && (
              <div style={{ background:"#fff", border:"1px solid #e2e8f0", borderRadius:8, padding:"20px 24px" }}>
                <div style={{ fontSize:14, color:"#475569", marginBottom:12, fontWeight:500 }}>¿Qué archivos subir?</div>
                <div style={{ display:"flex", flexWrap:"wrap", gap:8 }}>
                  {["ECO_optativas_202603.docx","ADM_optativas_otoño_2026.docx","COM_optativas_202603.docx","IIO_optativas_202603.docx","CON_optativas_2026-03.docx","DER_optativas_202603.docx"].map(name => (
                    <span key={name} style={{ fontFamily:"'DM Mono',monospace", fontSize:11, background:"#f8fafc", border:"1px solid #e2e8f0", padding:"3px 9px", borderRadius:3, color:"#475569" }}>{name}</span>
                  ))}
                  <span style={{ fontSize:12, color:"#94a3b8", padding:"3px 0" }}>y más…</span>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ── PREVIEW TAB ── */}
        {activeTab === "preview" && allCourses.length > 0 && (
          <div className="fade-in">
            {/* Summary */}
            <div style={{ display:"flex", gap:12, marginBottom:20, flexWrap:"wrap" }}>
              <div style={{ background:"#fff", border:"1px solid #e2e8f0", borderRadius:8, padding:"14px 20px", flex:1, minWidth:120, textAlign:"center" }}>
                <div style={{ fontSize:28, fontWeight:500, color:"#1F3864" }}>{allCourses.length}</div>
                <div style={{ fontSize:12, color:"#64748b", fontFamily:"'DM Mono',monospace" }}>materias</div>
              </div>
              <div style={{ background:"#fff", border:"1px solid #e2e8f0", borderRadius:8, padding:"14px 20px", flex:1, minWidth:120, textAlign:"center" }}>
                <div style={{ fontSize:28, fontWeight:500, color:"#1F3864" }}>{deptsSorted.length}</div>
                <div style={{ fontSize:12, color:"#64748b", fontFamily:"'DM Mono',monospace" }}>departamentos</div>
              </div>
              <div style={{ background:"#fff", border:"1px solid #e2e8f0", borderRadius:8, padding:"14px 20px", flex:1, minWidth:120, textAlign:"center" }}>
                <div style={{ fontSize:28, fontWeight:500, color:"#15803d" }}>{doneCount}</div>
                <div style={{ fontSize:12, color:"#64748b", fontFamily:"'DM Mono',monospace" }}>archivos OK</div>
              </div>
              <div style={{ display:"flex", flexDirection:"column", gap:8, justifyContent:"center" }}>
                <button className="btn-success" onClick={() => generateExcel(allCourses)}>
                  <span>↓</span> Descargar Excel
                </button>
                <button className="btn-secondary" style={{ fontSize:13, padding:"8px 14px" }} onClick={() => setActiveTab("upload")}>
                  + Agregar más
                </button>
              </div>
            </div>

            {/* Course tree */}
            <div style={{ background:"#fff", border:"1px solid #e2e8f0", borderRadius:8, overflow:"hidden" }}>
              {deptsSorted.map((dept, di) => {
                const dCourses = grouped[dept] || [];
                const isOpen = expandedDept === dept;
                const color = DEPT_COLORS[dept] || "#475569";
                const bg = DEPT_BG[dept] || "#f8fafc";
                return (
                  <div key={dept} style={{ borderBottom: di < deptsSorted.length - 1 ? "1px solid #e2e8f0" : "none" }}>
                    <div
                      className="dept-row"
                      onClick={() => setExpandedDept(isOpen ? null : dept)}
                      style={{ background: isOpen ? bg : "#fff" }}
                    >
                      <span style={{ color, fontSize:16, transition:"transform .2s", display:"inline-block", transform: isOpen ? "rotate(90deg)" : "none" }}>▶</span>
                      <div style={{ flex:1, display:"flex", alignItems:"center", gap:10 }}>
                        <span style={{ fontSize:14, fontWeight:500, color:"#1e293b" }}>{dept}</span>
                        <DeptBadge dept={dept} />
                      </div>
                      <span style={{ fontFamily:"'DM Mono',monospace", fontSize:11, color:"#94a3b8" }}>{dCourses.length} materias</span>
                    </div>
                    {isOpen && dCourses.map(c => (
                      <div key={c.code} className="course-item" style={{ background: bg + "88" }}>
                        <span className="code-mono">{c.code}</span>
                        <div style={{ flex:1 }}>
                          <div style={{ fontSize:13, fontWeight:500, color:"#1e293b", lineHeight:1.35 }}>{c.title}</div>
                          <div style={{ fontSize:12, color:"#64748b", marginTop:2, fontFamily:"'DM Mono',monospace" }}>{c.professor}</div>
                          {c.prerequisites && c.prerequisites !== "Ninguno" && (
                            <div style={{ fontSize:11, color:"#94a3b8", marginTop:2 }}>
                              <strong>Prereq:</strong> {c.prerequisites.length > 80 ? c.prerequisites.slice(0,80)+"…" : c.prerequisites}
                            </div>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
