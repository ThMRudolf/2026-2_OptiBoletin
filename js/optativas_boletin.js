/**
 * optativas_boletin.js
 * Lógica React (JSX/Babel) del Generador de Boletín de Optativas — ITAM
 *
 * Dependencias externas (cargadas como UMD en el HTML):
 *   - React 18 + ReactDOM  : componentes y renderizado
 *   - Babel standalone     : transpila JSX en el navegador
 *   - JSZip 3              : empaqueta los archivos del .docx
 *   - SheetJS (xlsx)       : lee el archivo Excel de entrada
 *
 * Flujo de tres pasos:
 *   1. El usuario sube un .xlsx con columnas A-F (dept, código, título, profesor, prereqs, descripción)
 *   2. Puede filtrar y seleccionar qué materias incluir por departamento
 *   3. Elige programa(s) y descarga el boletín como archivo .docx
 */

const { useState, useRef, useCallback } = React;

/**
 * Modo "solo Excel" para la versión independiente de colegas (dist/Optativas_Boletin.html).
 * La activa build/build_standalone.mjs con window.OPTI_EXCEL_ONLY = true.
 * En la versión normal es false y la app se comporta igual que siempre.
 */
const EXCEL_ONLY = window.OPTI_EXCEL_ONLY === true;

// ─── CONSTANTES ───────────────────────────────────────────────────────────────

/** Orden canónico de departamentos en el boletín generado. */
const DEPT_ORDER = OptiParser.DEPT_ORDER;

/** Nombres antiguos de departamento (Excels previos) → nombre actual. */
const DEPT_ALIASES = {
  "Ingeniería IEE": "Ingeniería Eléctrica y Electrónica",
  "Ingeniería / Interdepartamental": "Ingeniería Eléctrica y Electrónica"
};

/** Color de texto/borde para cada departamento (chips y encabezados). */
const DEPT_COLORS = {
  "Administración":"#2563eb","Computación":"#7c3aed","Contabilidad":"#0891b2",
  "Economía":"#059669","Ingeniería Industrial y Operaciones":"#c2410c",
  "Ingeniería Eléctrica y Electrónica":"#b45309","Matemáticas":"#be185d",
  "Derecho":"#4338ca","Estudios Generales":"#0f766e","Estudios Internacionales":"#0369a1",
  "Ciencia Política":"#7e22ce","Bienestar / Formación General":"#15803d",
  "Estadística":"#a21caf"
};

/** Color de fondo suave para cada departamento. */
const DEPT_BG = {
  "Administración":"#eff6ff","Computación":"#f5f3ff","Contabilidad":"#ecfeff",
  "Economía":"#f0fdf4","Ingeniería Industrial y Operaciones":"#fff7ed",
  "Ingeniería Eléctrica y Electrónica":"#fffbeb","Matemáticas":"#fdf2f8",
  "Derecho":"#eef2ff","Estudios Generales":"#f0fdfa","Estudios Internacionales":"#f0f9ff",
  "Ciencia Política":"#faf5ff","Bienestar / Formación General":"#f0fdf4",
  "Estadística":"#fdf4ff"
};

/** Nombre en mayúsculas para el encabezado del boletín Word. */
const DEPT_DISPLAY = {
  "Administración":"ADMINISTRACIÓN",
  "Bienestar / Formación General":"BIENESTAR / FORMACIÓN GENERAL",
  "Ciencia Política":"CIENCIA POLÍTICA",
  "Computación":"COMPUTACIÓN",
  "Contabilidad":"CONTABILIDAD",
  "Derecho":"DERECHO",
  "Economía":"ECONOMÍA",
  "Estadística":"ESTADÍSTICA",
  "Estudios Generales":"ESTUDIOS GENERALES",
  "Estudios Internacionales":"ESTUDIOS INTERNACIONALES",
  "Ingeniería Eléctrica y Electrónica":"INGENIERÍA ELÉCTRICA Y ELECTRÓNICA",
  "Ingeniería Industrial y Operaciones":"INGENIERÍA INDUSTRIAL Y OPERACIONES",
  "Matemáticas":"MATEMÁTICAS"
};

/** Programas académicos disponibles para portada del boletín (máx. 2). */
const PROGRAMS = [
  "Actuaría","Administración de Negocios","Ciencia de Datos","Ciencia Política",
  "Contaduría Analítica y Finanzas Corporativas","Derecho","Dirección Financiera",
  "Dirección de Mercadotecnia","Economía","Matemáticas Aplicadas",
  "Relaciones Internacionales","Ingeniería en Inteligencia Artificial",
  "Ingeniería y Ciencias de la Computación",
  "Ingeniería Industrial y en Sistemas Inteligentes",
  "Ingeniería en Mecatrónica y Robótica Inteligente","Ingeniería en Negocios"
];

/** Semestres del ITAM, en orden del año académico. */
const SEMESTERS = ["Primavera", "Verano", "Otoño"];

/**
 * Periodo propuesto por defecto: el siguiente semestre a partir de hoy
 * (ene–mar → Verano, abr–ago → Otoño, sep–dic → Primavera del año siguiente).
 */
function defaultPeriod(d = new Date()) {
  const m = d.getMonth() + 1, y = d.getFullYear();
  if (m <= 3) return { semester: "Verano", year: y };
  if (m <= 8) return { semester: "Otoño",  year: y };
  return { semester: "Primavera", year: y + 1 };
}

/** Texto seguro para nombres de archivo: sin acentos, espacios → "_". */
const fileSlug = s => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "")
  .replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "");

// ─── GENERACIÓN DE DOCX ───────────────────────────────────────────────────────

/**
 * Escapa caracteres especiales XML para uso seguro en el cuerpo del documento.
 * @param {*} s - Valor a escapar.
 * @returns {string}
 */
function esc(s) {
  return String(s || "")
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&apos;");
}

/**
 * Genera un elemento <w:r> (run de texto) con formato OOXML.
 * @param {string} text  - Contenido del run.
 * @param {boolean} bold - Aplicar negrita.
 * @param {number} size  - Tamaño en semipuntos (default 20 = 10 pt).
 * @param {string} color - Color hex sin # (default "000000").
 * @returns {string} Fragmento XML del run.
 */
function wRun(text, bold, size=20, color="000000") {
  const b = bold ? "<w:b/><w:bCs/>" : "";
  return `<w:r><w:rPr>` +
    `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>` +
    `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>` +
    `${b}<w:color w:val="${color}"/>` +
    `</w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

/**
 * Genera un elemento <w:p> (párrafo) que envuelve los runs dados.
 * @param {string} runs           - Contenido XML de los runs del párrafo.
 * @param {object} opts
 * @param {number} opts.before    - Espacio antes en twips (default 0).
 * @param {number} opts.after     - Espacio después en twips (default 80).
 * @param {string} opts.jc        - Justificación: "center", "both", etc.
 * @param {string} opts.indent    - Sangría izquierda en twips.
 * @param {string} opts.shading   - Color hex de fondo del párrafo.
 * @returns {string} Fragmento XML del párrafo.
 */
function wPara(runs, {before=0, after=80, jc="", indent="", shading=""}={}) {
  return `<w:p><w:pPr>` +
    (shading ? `<w:shd w:val="clear" w:color="auto" w:fill="${shading}"/>` : "") +
    `<w:spacing w:before="${before}" w:after="${after}" w:line="276" w:lineRule="auto"/>` +
    (jc     ? `<w:jc w:val="${jc}"/>` : "") +
    (indent ? `<w:ind w:left="${indent}"/>` : "") +
    `</w:pPr>${runs}</w:p>`;
}

/**
 * Convierte los datos de una materia en bloques de párrafos OOXML.
 * Genera: código+título, profesor, prerrequisitos, descripción y línea divisoria.
 * @param {{code:string, title:string, professor:string, prerequisites:string, description:string}} c
 * @returns {string} Fragmento XML con los párrafos de la materia.
 */
function courseToXml(c) {
  let out = "";
  out += wPara(wRun(`${c.code}   `, true, 20) + wRun(c.title, true, 20), {before:200, after:60});

  const profLines = (c.professor || "").split("\n");
  out += wPara(wRun("PROFESOR: ", true) + wRun(profLines[0] || ""), {after:40});
  for (let i = 1; i < profLines.length; i++)
    out += wPara(wRun(profLines[i]), {after:40, indent:"1260"});

  const preLines = (c.prerequisites || "Ninguno").split("\n");
  out += wPara(wRun("PRERREQUISITOS: ", true) + wRun(preLines[0]), {after:40});
  for (let i = 1; i < preLines.length; i++)
    out += wPara(wRun(preLines[i]), {after: i===preLines.length-1 ? 40 : 20, indent:"1620"});

  out += wPara(wRun("DESCRIPCIÓN: ", true) + wRun(c.description || ""), {after:240, jc:"both"});
  out += `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="2" w:space="1" w:color="CCCCCC"/></w:pBdr><w:spacing w:before="0" w:after="100"/></w:pPr></w:p>`;
  return out;
}

/**
 * Genera y devuelve el blob de un archivo .docx con las materias seleccionadas.
 * El documento se estructura como:
 *   [página de notas generales opcional] → portada → secciones por departamento.
 *
 * @param {Array<object>} selectedCourses - Materias a incluir.
 * @param {string} programName            - Nombre(s) del programa para la portada.
 * @param {string} [generalText=""]       - Texto libre de notas generales.
 * @param {string} [periodLabel=""]       - Periodo, p. ej. "Primavera 2027".
 * @returns {Promise<Blob>} Blob del archivo .docx.
 */
async function generateDocx(selectedCourses, programName, generalText="", periodLabel="") {
  const depts = DEPT_ORDER
    .filter(d => selectedCourses.some(c => c.department === d))
    .concat(
      selectedCourses.map(c => c.department)
        .filter(d => !DEPT_ORDER.includes(d))
        .filter((v, i, a) => a.indexOf(v) === i)
    );

  let body = "";

  if (generalText.trim()) {
    for (const line of generalText.split('\n'))
      body += wPara(wRun(line), {after:80});
    body += `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
  }

  body += wPara(wRun(`MATERIAS OPTATIVAS — ${periodLabel.toUpperCase()}`, true, 28, "1F3864"), {after:80, jc:"center"});
  body += wPara(wRun(programName || "", false, 22, "555555"), {after:360, jc:"center"});

  for (const dept of depts) {
    const courses = selectedCourses.filter(c => c.department === dept);
    if (!courses.length) continue;
    const dn = DEPT_DISPLAY[dept] || dept.toUpperCase();
    body += `<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="1F3864"/>` +
      `<w:spacing w:before="240" w:after="120"/><w:ind w:left="120"/></w:pPr>` +
      `${wRun(dn, true, 24, "FFFFFF")}</w:p>`;
    courses.forEach(c => { body += courseToXml(c); });
  }

  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<w:body>${body}` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>` +
    `</w:sectPr></w:body></w:document>`;

  const zip = new JSZip();
  zip.file("[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    `</Types>`
  );
  zip.file("_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships>`
  );
  zip.file("word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `</Relationships>`
  );
  zip.file("word/styles.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:docDefaults><w:rPrDefault><w:rPr>` +
    `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>` +
    `<w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="es-MX"/>` +
    `</w:rPr></w:rPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">` +
    `<w:name w:val="Normal"/><w:pPr><w:spacing w:after="0"/></w:pPr></w:style>` +
    `</w:styles>`
  );
  zip.file("word/document.xml", docXml);

  return zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
}

// ─── AYUDANTES EXCEL ──────────────────────────────────────────────────────────

/**
 * Infiere el departamento a partir del prefijo del código de materia.
 * Se usa cuando la columna A del Excel está vacía.
 * @param {string} code - Código de materia, e.g. "ECO-1234".
 * @returns {string} Nombre del departamento o el prefijo original si no se reconoce.
 */
function guessDept(code) {
  const p = (code || '').split(/[-–]/)[0].trim().toUpperCase();
  return OptiParser.PREFIX_TO_DEPT[p] || p;
}

/** Asigna un identificador único a cada materia (dos materias pueden compartir código). */
const withUids = courses => courses.map((c, i) => ({...c, uid: `${c.code}#${i}`}));

/** Dispara la descarga de un Blob en el navegador. */
function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ─── COMPONENTE: SELECTOR DE PERIODO ──────────────────────────────────────────

/**
 * Selector de semestre (Primavera / Verano / Otoño) y año.
 * Define el título del boletín y los nombres de los archivos descargados.
 */
function PeriodPicker({semester, year, onSemester, onYear}) {
  const now = new Date().getFullYear();
  const years = [];
  for (let y = now - 1; y <= now + 3; y++) years.push(y);
  if (!years.includes(year)) years.push(year);
  return (
    <div className="period-picker">
      <span className="src-label" style={{marginBottom:0}}>Periodo</span>
      <div className="seg">
        {SEMESTERS.map(s => (
          <button key={s} type="button" className={s === semester ? "on" : ""} onClick={() => onSemester(s)}>{s}</button>
        ))}
      </div>
      <select value={year} onChange={e => onYear(Number(e.target.value))} aria-label="Año">
        {years.sort((a, b) => a - b).map(y => <option key={y} value={y}>{y}</option>)}
      </select>
    </div>
  );
}

// ─── COMPONENTE: BARRA DE PASOS ───────────────────────────────────────────────

/**
 * Barra de progreso de los tres pasos del wizard (Subir / Seleccionar / Generar).
 * @param {{step: number}} props - Paso activo (1, 2 o 3).
 */
function StepBar({step}) {
  const steps = [EXCEL_ONLY ? "Subir Excel" : "Cargar materias", "Seleccionar materias", "Generar documento"];
  return (
    <div style={{display:"flex", alignItems:"center", marginBottom:28}}>
      {steps.map((s, i) => {
        const n = i + 1, active = step === n, done = step > n;
        return (
          <React.Fragment key={n}>
            <div style={{display:"flex", alignItems:"center", gap:8, flexShrink:0}}>
              <div style={{
                width:28, height:28, borderRadius:"50%",
                background: done ? "#15803d" : active ? "#1F3864" : "#e2e8f0",
                color: (done || active) ? "#fff" : "#94a3b8",
                display:"flex", alignItems:"center", justifyContent:"center",
                fontSize:12, fontFamily:"'DM Mono',monospace", fontWeight:500,
                transition:"all .3s"
              }}>
                {done ? "✓" : n}
              </div>
              <span style={{
                fontSize:13,
                color: active ? "#1F3864" : done ? "#15803d" : "#94a3b8",
                fontWeight: active ? 500 : 400,
                whiteSpace:"nowrap", transition:"color .3s"
              }}>{s}</span>
            </div>
            {i < steps.length - 1 &&
              <div className="step-line" style={{background: done ? "#86efac" : "#e2e8f0"}}/>}
          </React.Fragment>
        );
      })}
    </div>
  );
}

// ─── COMPONENTE PRINCIPAL ─────────────────────────────────────────────────────

/**
 * Componente raíz de la aplicación.
 * Gestiona el estado global del wizard y orquesta los tres pasos.
 */
function App() {
  // ── Estado del wizard ──
  const [step, setStep] = useState(1);

  // ── Estado del archivo Excel ──
  const [allCourses, setAllCourses] = useState([]);
  const [xlsxStatus, setXlsxStatus] = useState(null); // null | 'loading' | 'done' | 'error'
  const [xlsxFileName, setXlsxFileName] = useState(null);
  const [xlsxError, setXlsxError] = useState(null);
  const [xlsxDragOver, setXlsxDragOver] = useState(false);
  const xlsxInputRef = useRef();

  // ── Estado de la lectura de documentos Word ──
  const [wordStatus, setWordStatus] = useState(null); // null | 'loading' | 'done' | 'error'
  const [wordSource, setWordSource] = useState(null);
  const [wordError, setWordError]   = useState(null);
  const [wordReport, setWordReport] = useState(null); // {report, duplicates}
  const [wordDragOver, setWordDragOver] = useState(false);
  const folderInputRef = useRef();
  const docsInputRef   = useRef();

  // ── Estado del contenido opcional ──
  const [generalText, setGeneralText] = useState("");

  // ── Estado de selección de materias ──
  const [selected, setSelected] = useState(new Set());
  const [expandedDepts, setExpandedDepts] = useState(new Set());
  const [searchQuery, setSearchQuery] = useState("");

  // ── Estado del paso 3 ──
  const [selectedPrograms, setSelectedPrograms] = useState([]);
  const [generating, setGenerating] = useState(false);
  const [genDone, setGenDone] = useState(false);

  /** Alterna la selección de un programa (máximo 2 simultáneos). */
  const toggleProgram = useCallback((prog) => {
    setSelectedPrograms(prev =>
      prev.includes(prog) ? prev.filter(p => p !== prog)
      : prev.length >= 2  ? prev
      : [...prev, prog]
    );
  }, []);

  const programName = selectedPrograms.join(" / ");

  // ── Periodo (semestre + año): título del boletín y nombres de archivo ──
  const [semester, setSemester] = useState(() => defaultPeriod().semester);
  const [year, setYear]         = useState(() => defaultPeriod().year);
  const periodLabel = `${semester} ${year}`;                     // "Primavera 2027"
  const periodSlug  = `${fileSlug(semester)}_${year}`;           // "Primavera_2027"
  const periodPicker = <PeriodPicker semester={semester} year={year} onSemester={setSemester} onYear={setYear}/>;

  /**
   * Lee el archivo Excel y extrae las materias.
   * Formato esperado: columnas A (dept) B (código) C (título) D (profesor) E (prereqs) F (descripción).
   * La columna A puede estar vacía; el departamento se infiere del prefijo del código.
   */
  const loadXlsx = useCallback((file) => {
    if (!file.name.endsWith('.xlsx')) {
      setXlsxError('Por favor sube un archivo .xlsx válido');
      setXlsxStatus('error');
      setXlsxFileName(file.name);
      return;
    }
    resetData();
    setXlsxFileName(file.name);
    setXlsxStatus('loading');

    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, {type: 'array'});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {header: 1, defval: ''});
        const parsed = [];
        for (const row of rows) {
          let dept = String(row[0] || '').trim();
          dept = DEPT_ALIASES[dept] || dept;
          const code = String(row[1] || '').trim();
          // Sólo acepta filas con código válido (2-4 letras mayúsculas + guión + dígito)
          if (!code || !code.match(/^[A-Z]{2,4}[-–]\d/)) continue;
          parsed.push({
            department: dept || guessDept(code),
            code,
            title:        String(row[2] || '').trim(),
            professor:    String(row[3] || '').trim(),
            prerequisites:String(row[4] || '').trim(),
            description:  String(row[5] || '').trim()
          });
        }
        if (!parsed.length) {
          setXlsxStatus('error');
          setXlsxError('No se encontraron materias. Verifica que el archivo tenga el formato correcto (columnas A-F).');
          return;
        }
        const withIds = withUids(parsed);
        setAllCourses(withIds);
        setSelected(new Set(withIds.map(c => c.uid)));
        setXlsxStatus('done');
      } catch(err) {
        setXlsxStatus('error');
        setXlsxError('Error al leer el archivo: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  /**
   * Lee los .docx de la carpeta (o archivos) elegida(s) con OptiParser y carga
   * las materias encontradas. Equivale a ejecutar py/docx_to_excel.py.
   */
  const loadWordFiles = useCallback(async (fileList) => {
    const files = Array.from(fileList || []).filter(f => /\.docx?$/i.test(f.name) && !f.name.startsWith("~$"));
    resetData();
    if (!files.length) {
      setWordStatus('error');
      setWordError('No se encontraron archivos .doc / .docx en la selección.');
      return;
    }
    setWordStatus('loading');
    setWordSource(files.length === 1 ? files[0].name
      : (files[0].webkitRelativePath ? files[0].webkitRelativePath.split('/')[0] + '/' : '') + ` ${files.length} archivos`);
    try {
      const res = await OptiParser.parseWordFiles(files);
      setWordReport(res);
      if (!res.courses.length) {
        setWordStatus('error');
        setWordError('No se encontraron materias en los documentos.');
        return;
      }
      const withIds = withUids(res.courses);
      setAllCourses(withIds);
      setSelected(new Set(withIds.map(c => c.uid)));
      setWordStatus('done');
    } catch (err) {
      setWordStatus('error');
      setWordError('Error al leer los documentos: ' + err.message);
    }
  }, []);

  /** Borra los datos cargados (Excel o Word). */
  function resetData() {
    setXlsxStatus(null); setXlsxFileName(null); setXlsxError(null);
    setWordStatus(null); setWordSource(null); setWordError(null); setWordReport(null);
    setAllCourses([]); setSelected(new Set());
  }

  /** Descarga las materias cargadas como .xlsx (columnas A-F). */
  const downloadXlsx = () =>
    downloadBlob(OptiParser.coursesToXlsx(allCourses), `Optativas_${periodSlug}.xlsx`);

  // ── Agrupación y filtrado ──

  const grouped = {};
  for (const c of allCourses) {
    const d = c.department || "Otro";
    if (!grouped[d]) grouped[d] = [];
    grouped[d].push(c);
  }
  const deptsSorted = DEPT_ORDER
    .filter(d => grouped[d]?.length)
    .concat(Object.keys(grouped).filter(d => !DEPT_ORDER.includes(d) && grouped[d]?.length));

  const q = searchQuery.toLowerCase();
  const filteredGrouped = {};
  for (const dept of deptsSorted) {
    const m = (grouped[dept] || []).filter(c =>
      !q ||
      c.code?.toLowerCase().includes(q) ||
      c.title?.toLowerCase().includes(q) ||
      c.professor?.toLowerCase().includes(q)
    );
    if (m.length) filteredGrouped[dept] = m;
  }
  const filteredDepts = deptsSorted.filter(d => filteredGrouped[d]);

  // ── Acciones de selección ──

  // La selección se guarda por `uid` (no por código): dos materias pueden compartir código.
  const toggleCourse = uid => setSelected(prev => {
    const n = new Set(prev);
    n.has(uid) ? n.delete(uid) : n.add(uid);
    return n;
  });

  const selectAllInDept = dept => {
    const uids = (filteredGrouped[dept] || []).map(c => c.uid);
    setSelected(prev => { const n = new Set(prev); uids.forEach(u => n.add(u)); return n; });
  };

  const deselectAllInDept = dept => {
    const uids = (filteredGrouped[dept] || []).map(c => c.uid);
    setSelected(prev => { const n = new Set(prev); uids.forEach(u => n.delete(u)); return n; });
  };

  const toggleExpand = dept => setExpandedDepts(prev => {
    const n = new Set(prev);
    n.has(dept) ? n.delete(dept) : n.add(dept);
    return n;
  });

  const selectAll   = () => setSelected(new Set(allCourses.map(c => c.uid)));
  const deselectAll = () => setSelected(new Set());

  const selectedCourses = allCourses.filter(c => selected.has(c.uid));

  // ── Descarga del documento ──

  /** Genera el .docx y dispara la descarga en el navegador. */
  const handleGenerate = async () => {
    setGenerating(true);
    setGenDone(false);
    try {
      const blob = await generateDocx(selectedCourses, programName, generalText, periodLabel);
      const progs = selectedPrograms.map(fileSlug).join("_y_");
      downloadBlob(blob, `Boletin_Optativas_${periodSlug}${progs ? "_" + progs : ""}.docx`);
      setGenDone(true);
    } catch(e) {
      alert("Error al generar: " + e.message);
    } finally {
      setGenerating(false);
    }
  };

  // ── Estilo de tarjeta reutilizado en varios pasos ──
  const card = {
    background:"#fff", border:"1px solid #e2e8f0",
    borderRadius:10, padding:"24px 28px", marginBottom:16
  };

  // ── Render ──
  return (
    <div style={{minHeight:"100vh", background:"#f5f4f0", fontFamily:"'Crimson Pro',Georgia,serif"}}>

      {/* ── Encabezado institucional ── */}
      <div style={{background:"#1F3864", padding:"20px 32px 18px", borderBottom:"3px solid #e8b84b"}}>
        <div style={{maxWidth:920, margin:"0 auto"}}>
          <div style={{fontFamily:"'DM Mono',monospace", fontSize:10, color:"#8fa4cc", letterSpacing:".12em", marginBottom:4, textTransform:"uppercase"}}>
            ITAM · División Académica de Ingeniería · {periodLabel}
          </div>
          <div style={{color:"#e8ecf4", fontSize:22, fontWeight:500}}>Generador de Boletín de Optativas</div>
          <div style={{color:"#7f9abf", fontSize:13, marginTop:2}}>
            {EXCEL_ONLY ? "Carga el Excel con las materias" : "Carga las materias (documentos Word o Excel)"} · selecciona las que incluir · genera el Word
          </div>
        </div>
      </div>

      <div style={{maxWidth:920, margin:"0 auto", padding:"28px 20px"}}>
        <StepBar step={step}/>

        {/* ════════ PASO 1: SUBIR EXCEL ════════ */}
        {step===1 && (
          <div className="fade">
            <div style={card}>
              <div style={{marginBottom:16}}>
                <div style={{fontSize:18, fontWeight:500, color:"#1a1a1a", marginBottom:3}}>
                  {EXCEL_ONLY ? "Sube el archivo Excel de optativas" : "Carga las materias optativas"}
                </div>
                <div style={{fontSize:13, color:"#64748b"}}>
                  {EXCEL_ONLY
                    ? "Usa el Excel de optativas que te fue enviado."
                    : <>Lee los anuncios en Word de los departamentos <b>o</b> sube un Excel ya preparado.</>}
                </div>
              </div>

              <div style={{marginBottom:16}}>{periodPicker}</div>

              <div className={"src-grid" + (EXCEL_ONLY ? " single" : "")}>
              {/* ── Opción A: documentos Word (oculta en modo solo Excel) ── */}
              {!EXCEL_ONLY && <div className="src-panel">
                <div className="src-label">A · Documentos Word (.docx)</div>
                <div
                  className={`drop-z${wordDragOver ? " over" : ""}`}
                  onDragOver={e => { e.preventDefault(); setWordDragOver(true); }}
                  onDragLeave={()  => setWordDragOver(false)}
                  onDrop={e => { e.preventDefault(); setWordDragOver(false); loadWordFiles(e.dataTransfer.files); }}
                  onClick={() => folderInputRef.current?.click()}
                >
                  <div style={{fontSize:32, marginBottom:8}}>📁</div>
                  <div style={{fontSize:15, color:"#334155", marginBottom:3}}>Selecciona la carpeta de entrada</div>
                  <div style={{fontSize:12, color:"#94a3b8"}}>o arrastra aquí los archivos .docx</div>
                  <input ref={folderInputRef} type="file" webkitdirectory="" directory="" multiple style={{display:"none"}}
                    onChange={e => { loadWordFiles(e.target.files); e.target.value=""; }}/>
                  <input ref={docsInputRef} type="file" accept=".docx,.doc" multiple style={{display:"none"}}
                    onChange={e => { loadWordFiles(e.target.files); e.target.value=""; }}/>
                </div>
                <div style={{fontSize:11, color:"#94a3b8", marginTop:-6}}>
                  ¿Prefieres elegir archivos sueltos?{" "}
                  <button className="link-btn" onClick={() => docsInputRef.current?.click()}>Seleccionar archivos</button>
                </div>
              </div>}

              {/* ── Opción B: Excel existente ── */}
              <div className="src-panel">
                {!EXCEL_ONLY && <div className="src-label">B · Excel existente (.xlsx)</div>}
              <div
                className={`drop-z${xlsxDragOver ? " over" : ""}`}
                onDragOver={e => { e.preventDefault(); setXlsxDragOver(true); }}
                onDragLeave={()  => setXlsxDragOver(false)}
                onDrop={e => { e.preventDefault(); setXlsxDragOver(false); if (e.dataTransfer.files[0]) loadXlsx(e.dataTransfer.files[0]); }}
                onClick={() => xlsxInputRef.current?.click()}
              >
                <div style={{fontSize:32, marginBottom:8}}>📊</div>
                <div style={{fontSize:15, color:"#334155", marginBottom:3}}>
                  Arrastra el <code style={{fontFamily:"'DM Mono',monospace", fontSize:12, background:"#ebe9e3", padding:"1px 6px", borderRadius:3}}>.xlsx</code> aquí
                </div>
                <div style={{fontSize:12, color:"#94a3b8"}}>o haz clic para seleccionar</div>
                <input ref={xlsxInputRef} type="file" accept=".xlsx" style={{display:"none"}}
                  onChange={e => { if (e.target.files[0]) loadXlsx(e.target.files[0]); e.target.value=""; }}/>
              </div>
                <div style={{fontSize:11, color:"#94a3b8", marginTop:-6}}>
                  Columnas <span style={{fontFamily:"'DM Mono',monospace", fontSize:10}}>A Departamento · B Código · C Título · D Profesor · E Prerrequisitos · F Descripción</span>
                </div>
              </div>
              </div>

              {/* Estado de lectura de los documentos Word */}
              {wordStatus && (
                <div className={`status-box ${wordStatus}`}>
                  <div style={{display:"flex", alignItems:"center", gap:10}}>
                    <span style={{fontSize:16}}>📁</span>
                    <div style={{flex:1, minWidth:0}}>
                      <div style={{fontFamily:"'DM Mono',monospace", fontSize:12, color:"#1e293b", fontWeight:500, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap"}}>{wordSource || "Documentos Word"}</div>
                      {wordStatus==='done' && <div style={{fontSize:11, color:"#15803d"}}>{allCourses.length} materias · {[...new Set(allCourses.map(c => c.department))].length} departamentos</div>}
                      {wordError && <div style={{fontSize:11, color:"#b91c1c", marginTop:2}}>{wordError}</div>}
                    </div>
                    {wordStatus==='loading' && <div className="spinner" style={{width:13, height:13, border:"2px solid #bfdbfe", borderTopColor:"#1d4ed8"}}/>}
                    {wordStatus!=='loading' &&
                      <button onClick={resetData} style={{background:"none", border:"none", cursor:"pointer", color:"#94a3b8", fontSize:17, lineHeight:1, padding:"0 2px"}}>×</button>}
                  </div>
                  {wordStatus==='loading' && <div className="prog-bar"><div className="prog-inner"/></div>}
                  {wordReport && (
                    <details className="file-report" open={wordStatus==='error'}>
                      <summary>Detalle por archivo ({wordReport.report.length}){wordReport.duplicates.length ? ` · ${wordReport.duplicates.length} duplicado(s) omitido(s)` : ""}</summary>
                      <ul>
                        {wordReport.report.map(r => (
                          <li key={r.file} className={r.status}>
                            <span className="ico">{r.status==='ok' ? "✓" : r.status==='skipped' ? "–" : "✗"}</span>
                            <span className="fname">{r.file}</span>
                            <span className="info">{r.status==='ok' ? `${r.count} materia${r.count!==1?"s":""}` : r.message}</span>
                          </li>
                        ))}
                        {wordReport.duplicates.map((d, i) => (
                          <li key={"dup"+i} className="skipped"><span className="ico">⧉</span><span className="fname">{d}</span><span className="info">duplicado</span></li>
                        ))}
                      </ul>
                    </details>
                  )}
                </div>
              )}

              {/* Estado de carga del archivo */}
              {xlsxStatus && (
                <div style={{
                  background: xlsxStatus==='done' ? "#f0fdf4" : xlsxStatus==='error' ? "#fef2f2" : "#eff6ff",
                  border: `1px solid ${xlsxStatus==='done' ? "#bbf7d0" : xlsxStatus==='error' ? "#fecaca" : "#bfdbfe"}`,
                  borderRadius:6, padding:"9px 13px", marginBottom:7
                }}>
                  <div style={{display:"flex", alignItems:"center", gap:10}}>
                    <span style={{fontSize:16}}>📊</span>
                    <div style={{flex:1, minWidth:0}}>
                      <div style={{fontFamily:"'DM Mono',monospace", fontSize:12, color:"#1e293b", fontWeight:500, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap"}}>{xlsxFileName}</div>
                      {xlsxStatus==='done' && <div style={{fontSize:11, color:"#15803d"}}>{allCourses.length} materias · {[...new Set(allCourses.map(c => c.department))].length} departamentos</div>}
                      {xlsxError         && <div style={{fontSize:11, color:"#b91c1c", marginTop:2}}>{xlsxError}</div>}
                    </div>
                    {xlsxStatus==='loading' && <div className="spinner" style={{width:13, height:13, border:"2px solid #bfdbfe", borderTopColor:"#1d4ed8"}}/>}
                    <span style={{fontSize:13, color: xlsxStatus==='done' ? "#15803d" : xlsxStatus==='error' ? "#b91c1c" : "#1d4ed8", fontWeight:600}}>
                      {xlsxStatus==='done' ? "✓" : xlsxStatus==='error' ? "✗" : "⚙"}
                    </span>
                    {xlsxStatus!=='loading' &&
                      <button
                        onClick={() => { setXlsxStatus(null); setXlsxFileName(null); setXlsxError(null); setAllCourses([]); setSelected(new Set()); }}
                        style={{background:"none", border:"none", cursor:"pointer", color:"#94a3b8", fontSize:17, lineHeight:1, padding:"0 2px"}}>×</button>}
                  </div>
                  {xlsxStatus==='loading' && <div className="prog-bar"><div className="prog-inner"/></div>}
                </div>
              )}

              {allCourses.length > 0 && (
                <div style={{display:"flex", gap:12, alignItems:"center", marginTop:16, flexWrap:"wrap"}}>
                  <button className="btn-green" onClick={() => setStep(2)}>→ Revisar selección ({allCourses.length} materias)</button>
                  {wordStatus==='done' && <button className="btn-sec" onClick={downloadXlsx}>↓ Descargar Excel</button>}
                  <button className="btn-sec" onClick={resetData}>Limpiar</button>
                </div>
              )}

              {!xlsxStatus && !wordStatus && (
                <div style={{background:"#f8fafc", border:"1px solid #e2e8f0", borderRadius:8, padding:"12px 16px", marginTop:4}}>
                  <div style={{fontSize:11, color:"#94a3b8"}}>
                    {EXCEL_ONLY
                      ? "El archivo se lee en tu navegador; no se sube a ningún servidor. La columna Departamento puede quedar en blanco: se infiere del código (ECO, IIO, ADM…)."
                      : <>Los documentos se leen en tu navegador; no se suben a ningún servidor. Se omiten los archivos «Notas al plan de estudios».
                        Los .doc antiguos deben guardarse antes como .docx. En el Excel, la columna Departamento puede quedar en blanco: se infiere del código (ECO, IIO, ADM…).</>}
                  </div>
                </div>
              )}
            </div>

            {/* ── Notas generales (opcional) ── */}
            <div style={{...card, marginBottom:0}}>
              <div style={{display:"flex", alignItems:"center", gap:10, marginBottom:10}}>
                <div style={{flex:1}}>
                  <div style={{fontSize:16, fontWeight:500, color:"#1a1a1a", marginBottom:2}}>Notas generales — información previa</div>
                  <div style={{fontSize:12, color:"#94a3b8"}}>Opcional · Aparecerá al inicio del boletín, antes de las materias.</div>
                </div>
                {generalText.trim() && <span style={{fontFamily:"'DM Mono',monospace", fontSize:10, background:"#15803d", color:"#f0fdf4", padding:"2px 8px", borderRadius:3}}>incluido</span>}
              </div>
              <textarea
                value={generalText}
                onChange={e => setGeneralText(e.target.value)}
                placeholder="Pega o escribe aquí las notas al plan de estudios…"
                style={{width:"100%", minHeight:130, fontFamily:"'Crimson Pro',Georgia,serif", fontSize:14, border:"1px solid #e2e8f0", borderRadius:6, padding:"10px 14px", resize:"vertical", outline:"none", color:"#1e293b", background:"#fafafa", lineHeight:1.6, transition:"border .15s"}}
                onFocus={e => e.target.style.borderColor="#1F3864"}
                onBlur={e  => e.target.style.borderColor="#e2e8f0"}
              />
              {generalText.trim() && (
                <div style={{display:"flex", justifyContent:"flex-end", marginTop:6}}>
                  <button className="btn-sec" style={{fontSize:12, padding:"4px 12px"}} onClick={() => setGeneralText("")}>Limpiar</button>
                </div>
              )}
            </div>
          </div>
        )}

        {/* ════════ PASO 2: SELECCIONAR MATERIAS ════════ */}
        {step===2 && (
          <div className="fade">
            <div style={{display:"flex", gap:10, alignItems:"center", marginBottom:14, flexWrap:"wrap"}}>
              <input className="search-bar" style={{flex:1, minWidth:200}}
                placeholder="Buscar por código, título o profesor…"
                value={searchQuery} onChange={e => setSearchQuery(e.target.value)}/>
              <button className="btn-sec" style={{fontSize:13, padding:"8px 14px", whiteSpace:"nowrap"}} onClick={selectAll}>Todo</button>
              <button className="btn-sec" style={{fontSize:13, padding:"8px 14px", whiteSpace:"nowrap"}} onClick={deselectAll}>Ninguno</button>
              {!EXCEL_ONLY && wordStatus==='done' &&
                <button className="btn-sec" style={{fontSize:13, padding:"8px 14px", whiteSpace:"nowrap"}} onClick={downloadXlsx}
                  title="Descarga todas las materias leídas de los documentos Word">↓ Excel</button>}
              <span style={{fontFamily:"'DM Mono',monospace", fontSize:11, color:"#64748b", whiteSpace:"nowrap"}}>
                <span style={{color:"#1F3864", fontWeight:500}}>{selected.size}</span>/{allCourses.length} seleccionadas
              </span>
            </div>

            {/* Lista de departamentos / materias */}
            <div style={{background:"#fff", border:"1px solid #e2e8f0", borderRadius:10, overflow:"hidden", marginBottom:14, maxHeight:500, overflowY:"auto"}}>
              {filteredDepts.map((dept, di) => {
                const dCourses = filteredGrouped[dept] || [];
                const selIn    = dCourses.filter(c => selected.has(c.uid)).length;
                const allIn    = selIn === dCourses.length;
                const partial  = selIn > 0 && !allIn;
                const isOpen   = expandedDepts.has(dept);
                const color    = DEPT_COLORS[dept] || "#475569";
                const bg       = DEPT_BG[dept]    || "#f8fafc";
                return (
                  <div key={dept} style={{borderBottom: di < filteredDepts.length-1 ? "1px solid #e2e8f0" : "none"}}>
                    <div className="dept-hdr" style={{background: isOpen ? bg+"cc" : "#fff", gap:8}}>
                      <div style={{flex:1, display:"flex", alignItems:"center", gap:10, cursor:"pointer"}} onClick={() => toggleExpand(dept)}>
                        <span style={{color, fontSize:11, transition:"transform .2s", display:"inline-block", transform: isOpen ? "rotate(90deg)" : "none"}}>▶</span>
                        <span style={{fontSize:14, fontWeight:500, color:"#1e293b"}}>{dept}</span>
                        <span style={{fontFamily:"'DM Mono',monospace", fontSize:10, background:bg, color, border:`1px solid ${color}22`, padding:"1px 7px", borderRadius:3}}>{selIn}/{dCourses.length}</span>
                      </div>
                      <button onClick={() => selectAllInDept(dept)} style={{flexShrink:0, padding:"3px 9px", border:`1px solid ${color}`, borderRadius:4, background:"transparent", color, fontSize:11, fontFamily:"'DM Mono',monospace", cursor:"pointer", whiteSpace:"nowrap"}}>
                        + Sel. todo
                      </button>
                      <button onClick={() => deselectAllInDept(dept)} style={{flexShrink:0, padding:"3px 9px", border:"1px solid #e2e8f0", borderRadius:4, background:"transparent", color:"#94a3b8", fontSize:11, fontFamily:"'DM Mono',monospace", cursor:"pointer", whiteSpace:"nowrap"}}>
                        − Quitar todo
                      </button>
                    </div>
                    {isOpen && dCourses.map(c => {
                      const isSel = selected.has(c.uid);
                      return (
                        <div key={c.uid} className="course-row" style={{background: isSel ? bg+"55" : "transparent", paddingLeft:38, alignItems:"center", cursor:"default"}}>
                          <span className="pill" style={{marginTop:0}}>{c.code}</span>
                          <div style={{flex:1}}>
                            <div style={{fontSize:13, fontWeight: isSel?500:400, color: isSel?"#1e293b":"#374151", lineHeight:1.35}}>{c.title}</div>
                            <div style={{fontSize:11, color:"#64748b", marginTop:2, fontFamily:"'DM Mono',monospace"}}>{c.professor}</div>
                            {c.prerequisites && c.prerequisites !== "Ninguno" && (
                              <div style={{fontSize:11, color:"#94a3b8", marginTop:1}}>
                                Prereq: {c.prerequisites.length > 72 ? c.prerequisites.slice(0,72)+"…" : c.prerequisites}
                              </div>
                            )}
                          </div>
                          <button onClick={() => toggleCourse(c.uid)} style={{flexShrink:0, padding:"4px 11px", border: isSel?"1px solid #fca5a5":"1px solid #1F3864", borderRadius:5, background: isSel?"#fef2f2":"#1F3864", color: isSel?"#b91c1c":"#e8ecf4", fontSize:12, fontFamily:"'DM Mono',monospace", cursor:"pointer", whiteSpace:"nowrap", transition:"all .12s"}}>
                            {isSel ? "× Quitar" : "✓ Agregar"}
                          </button>
                        </div>
                      );
                    })}
                  </div>
                );
              })}
              {filteredDepts.length === 0 && (
                <div style={{padding:"32px", textAlign:"center", color:"#94a3b8", fontSize:14}}>No se encontraron materias para "{searchQuery}"</div>
              )}
            </div>

            {/* Chips de materias seleccionadas (máx. 28 visibles) */}
            {selected.size > 0 && (
              <div style={{background:"#f8fafc", border:"1px solid #e2e8f0", borderRadius:8, padding:"12px 16px", marginBottom:14}}>
                <div style={{fontFamily:"'DM Mono',monospace", fontSize:10, color:"#64748b", letterSpacing:".04em", marginBottom:7}}>SELECCIONADAS ({selected.size})</div>
                <div style={{display:"flex", flexWrap:"wrap", gap:6}}>
                  {selectedCourses.slice(0, 28).map(c => (
                    <span key={c.uid} className="tag-chip" style={{background: DEPT_BG[c.department]||"#f1f5f9", color: DEPT_COLORS[c.department]||"#475569", border:`1px solid ${(DEPT_COLORS[c.department]||"#cbd5e1")}22`}}>
                      <span style={{fontFamily:"'DM Mono',monospace", fontSize:10}}>{c.code}</span>
                      <button onClick={() => toggleCourse(c.uid)} style={{background:"none", border:"none", cursor:"pointer", color:"inherit", fontSize:13, lineHeight:1, opacity:.6, padding:0}}>×</button>
                    </span>
                  ))}
                  {selected.size > 28 && <span style={{fontSize:12, color:"#94a3b8", padding:"3px 0"}}>+{selected.size-28} más</span>}
                </div>
              </div>
            )}

            <div style={{display:"flex", gap:12, alignItems:"center", flexWrap:"wrap"}}>
              <button className="btn-sec" onClick={() => setStep(1)}>← Volver</button>
              <button className="btn-green" onClick={() => { if (selected.size > 0) setStep(3); }} disabled={selected.size===0}>
                → Continuar con {selected.size} materia{selected.size!==1?"s":""}
              </button>
            </div>
          </div>
        )}

        {/* ════════ PASO 3: CONFIGURAR Y GENERAR ════════ */}
        {step===3 && (
          <div className="fade">
            <div style={card}>
              <div style={{fontSize:18, fontWeight:500, color:"#1a1a1a", marginBottom:4}}>Configurar y generar documento</div>
              <div style={{fontSize:13, color:"#64748b", marginBottom:18}}>Selecciona el programa o los programas (máximo 2) para el boletín.</div>

              <div style={{marginBottom:18}}>{periodPicker}</div>

              <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10}}>
                <label style={{fontFamily:"'DM Mono',monospace", fontSize:11, color:"#64748b", letterSpacing:".04em"}}>PROGRAMA(S)</label>
                <span style={{fontFamily:"'DM Mono',monospace", fontSize:11, color: selectedPrograms.length===2 ? "#1F3864" : "#94a3b8"}}>
                  {selectedPrograms.length}/2 seleccionados
                </span>
              </div>

              {/* Selector de programas */}
              <div style={{display:"flex", flexWrap:"wrap", gap:7, marginBottom: selectedPrograms.length>0 ? 12 : 20}}>
                {PROGRAMS.map(prog => {
                  const isSel      = selectedPrograms.includes(prog);
                  const isDisabled = !isSel && selectedPrograms.length >= 2;
                  return (
                    <button key={prog} onClick={() => toggleProgram(prog)}
                      style={{padding:"6px 14px", border: isSel?"1.5px solid #1F3864":"1px solid #d1d5db", borderRadius:6, background: isSel?"#1F3864":"#fff", color: isSel?"#e8ecf4": isDisabled?"#c8c4bc":"#334155", fontSize:13, fontFamily:"'Crimson Pro',serif", cursor: isDisabled?"default":"pointer", transition:"all .12s", opacity: isDisabled?.45:1, fontWeight: isSel?500:400}}>
                      {prog}
                    </button>
                  );
                })}
              </div>

              {selectedPrograms.length > 0 && (
                <div style={{display:"flex", alignItems:"center", gap:8, marginBottom:20, padding:"9px 14px", background:"#eff6ff", border:"1px solid #bfdbfe", borderRadius:6}}>
                  <span style={{fontFamily:"'DM Mono',monospace", fontSize:10, color:"#1d4ed8", letterSpacing:".04em", flexShrink:0}}>BOLETÍN PARA:</span>
                  <span style={{fontSize:14, color:"#1e293b", fontWeight:500, flex:1}}>{programName}</span>
                  <button onClick={() => setSelectedPrograms([])} style={{background:"none", border:"none", cursor:"pointer", color:"#94a3b8", fontSize:16, lineHeight:1, padding:0, flexShrink:0}}>×</button>
                </div>
              )}

              {/* Resumen por departamento */}
              <div style={{fontFamily:"'DM Mono',monospace", fontSize:10, color:"#64748b", letterSpacing:".04em", marginBottom:10}}>RESUMEN — {selected.size} MATERIAS</div>
              <div style={{display:"flex", flexWrap:"wrap", gap:8, marginBottom:22}}>
                {deptsSorted.filter(d => selectedCourses.some(c => c.department===d)).map(d => {
                  const cnt   = selectedCourses.filter(c => c.department===d).length;
                  const color = DEPT_COLORS[d] || "#475569";
                  const bg    = DEPT_BG[d]    || "#f8fafc";
                  return (
                    <div key={d} style={{display:"flex", alignItems:"center", gap:7, background:bg, border:`1px solid ${color}33`, borderRadius:5, padding:"5px 11px 5px 8px"}}>
                      <span style={{fontFamily:"'DM Mono',monospace", fontSize:11, background:color, color:"#fff", padding:"1px 6px", borderRadius:3, fontWeight:500}}>{cnt}</span>
                      <span style={{fontSize:13, color}}>{d}</span>
                    </div>
                  );
                })}
              </div>

              {/* Vista previa del documento */}
              <div style={{border:"1px solid #e2e8f0", borderRadius:8, overflow:"hidden", marginBottom:24}}>
                <div style={{background:"#1F3864", padding:"7px 16px"}}>
                  <span style={{fontFamily:"'DM Mono',monospace", fontSize:10, color:"#8fa4cc", letterSpacing:".08em"}}>VISTA PREVIA</span>
                </div>
                <div style={{padding:"16px", maxHeight:220, overflowY:"auto", background:"#fafafa", fontFamily:"Arial,sans-serif"}}>
                  <div style={{textAlign:"center", marginBottom:8}}>
                    <div style={{fontSize:13, fontWeight:700, color:"#1F3864"}}>MATERIAS OPTATIVAS — {periodLabel.toUpperCase()}</div>
                    <div style={{fontSize:11, color:"#555", marginTop:2}}>{programName}</div>
                  </div>
                  {deptsSorted.filter(d => selectedCourses.some(c => c.department===d)).map(d => (
                    <div key={d} style={{marginBottom:6}}>
                      <div style={{background:"#1F3864", color:"#fff", padding:"3px 8px", fontSize:9, fontWeight:700, letterSpacing:".04em", marginBottom:3}}>{DEPT_DISPLAY[d]||d.toUpperCase()}</div>
                      {selectedCourses.filter(c => c.department===d).map(c => (
                        <div key={c.uid} style={{fontSize:10, marginBottom:4, paddingLeft:4, borderLeft:"2px solid #e2e8f0"}}>
                          <span style={{fontWeight:700}}>{c.code}</span> <span style={{fontWeight:700}}>{c.title}</span>
                          <div style={{color:"#555"}}><strong>PROFESOR:</strong> {c.professor}</div>
                        </div>
                      ))}
                    </div>
                  ))}
                </div>
              </div>

              <div style={{display:"flex", gap:12, alignItems:"center", flexWrap:"wrap"}}>
                <button className="btn-sec" onClick={() => { setStep(2); setGenDone(false); }}>← Editar selección</button>
                <button className="btn-green" onClick={handleGenerate} disabled={generating || selectedPrograms.length===0}>
                  {generating
                    ? <><div className="spinner" style={{width:14, height:14, border:"2px solid #86efac", borderTopColor:"#f0fdf4"}}/>Generando…</>
                    : <>↓ Descargar Word (.docx)</>}
                </button>
                {selectedPrograms.length===0 && <span style={{fontSize:12, color:"#94a3b8"}}>Selecciona al menos un programa</span>}
                {genDone && (
                  <div style={{display:"flex", alignItems:"center", gap:7, background:"#f0fdf4", border:"1px solid #86efac", borderRadius:5, padding:"8px 14px", fontSize:13, color:"#15803d"}}>
                    ✓ Documento descargado correctamente
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

// ─── PUNTO DE ENTRADA ─────────────────────────────────────────────────────────
ReactDOM.createRoot(document.getElementById("root")).render(<App/>);
