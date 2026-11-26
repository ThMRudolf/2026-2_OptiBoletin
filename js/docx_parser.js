/**
 * docx_parser.js
 * Lector de anuncios de optativas en Word (.docx) — se ejecuta 100 % en el navegador.
 *
 * Es el port en JavaScript de py/docx_to_excel.py (mismas reglas de parseo).
 * Si cambias las reglas en un archivo, cámbialas también en el otro.
 *
 * Dependencias: JSZip (global), cargado en el HTML.
 *
 * API (global window.OptiParser):
 *   parseWordFiles(files: File[]) → Promise<{courses, report, duplicates}>
 *     courses    : [{department, code, title, professor, prerequisites, description}]
 *     report     : [{file, status: 'ok'|'skipped'|'error', count, message}]
 *     duplicates : ["CODE ← archivo", …]
 *   coursesToXlsx(courses) → Blob   (requiere SheetJS / XLSX global)
 *   DEPT_ORDER, PREFIX_TO_DEPT
 */
(function () {
  "use strict";

  // ─── DEPARTAMENTOS ──────────────────────────────────────────────────────────

  const DEPT_ORDER = [
    "Administración", "Bienestar / Formación General", "Ciencia Política",
    "Computación", "Contabilidad", "Derecho", "Economía", "Estadística",
    "Estudios Generales", "Estudios Internacionales",
    "Ingeniería Eléctrica y Electrónica",
    "Ingeniería Industrial y Operaciones", "Matemáticas",
  ];

  const PREFIX_TO_DEPT = {
    ADM: "Administración",     CEB: "Bienestar / Formación General",
    CSO: "Ciencia Política",   COM: "Computación",
    CON: "Contabilidad",       DER: "Derecho",
    ECO: "Economía",           EST: "Estadística",
    EGN: "Estudios Generales", EIN: "Estudios Internacionales",
    SDI: "Ingeniería Eléctrica y Electrónica",
    IEE: "Ingeniería Eléctrica y Electrónica",
    IIO: "Ingeniería Industrial y Operaciones",
    MAT: "Matemáticas",
  };

  const deptFromCode = code => {
    const p = code.split("-")[0].toUpperCase();
    return PREFIX_TO_DEPT[p] || p;
  };

  // ─── EXPRESIONES REGULARES ─────────────────────────────────────────────────

  const CODE_AT_START = new RegExp(
    "^\\s*(?:(?:CLAVE|MATERIA)\\s*[:.]?\\s*)?" +   // CLAVE: / MATERIA: opcional
    "([A-Z]{2,4})" +                               // prefijo
    "[\\s\\-–—\\u00a0]*" +                         // separador opcional
    "(\\d{4,6})" +                                 // 4-6 dígitos
    "(?:[A-Z](?![A-Za-zÀ-ÿ]))?",                  // sufijo de 1 letra (no 1a letra del título)
    "i"
  );
  const PROF_RE    = /^\s*PROFESOR(?:ES|AS|A)?\s*[:.]?\s*/i;
  const PROF_SPLIT = /PROFESOR(?:ES|AS|A)?\s*[:.]?\s*/i;
  const PREREQ_RE  = /^\s*(?:SIN\s+)?PRERREQUISITOS?\s*[:.]?\s*/i;
  const SIN_PRE_RE = /^\s*SIN\s+PRERREQUISITO/i;
  const DESC_RE    = /^\s*DESCRIPCI[ÓO]N\s*[:.]?\s*/i;
  const NOMBRE_RE  = /^\s*NOMBRE\s+DEL\s+CURSO\s*[:.]?\s*/i;
  const CLAVE_RE   = /^\s*CLAVE\s*[:.]?\s*/i;
  const STOP_RE    = /^\s*(TEMARIO|NOTAS?\s*:|PROGRAMA\s*:)/i;
  const NOTES_RE   = /^\s*NOTAS\s+AL\s+PLAN\s+DE\s+ESTUDIOS/i;
  const BREVE_RE   = /^Breve descripción del contenido\.?\s*/i;

  // ─── AYUDANTES DE CADENAS (equivalentes a str.strip/rstrip de Python) ─────

  const esc = s => s.replace(/[-\\^\]]/g, "\\$&");
  const rstrip = (s, chars) => s.replace(new RegExp("[" + esc(chars) + "]+$"), "");
  const lstrip = (s, chars) => s.replace(new RegExp("^[" + esc(chars) + "]+"), "");
  const strip  = (s, chars) => lstrip(rstrip(s, chars), chars);

  function extractCodeAtStart(text) {
    const m = CODE_AT_START.exec(text);
    if (!m) return [null, null];
    const prefix = m[1].toUpperCase();
    if (!(prefix in PREFIX_TO_DEPT)) return [null, null];
    const remainder = lstrip(text.slice(m[0].length).trim(), "-–—:. ");
    return [`${prefix}-${m[2]}`, remainder];
  }

  // ─── LECTURA DEL .DOCX ─────────────────────────────────────────────────────

  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  /** Texto de un elemento: concatenación de todos sus <w:t>. */
  function paraText(el) {
    let s = "";
    const ts = el.getElementsByTagNameNS(W_NS, "t");
    for (let i = 0; i < ts.length; i++) s += ts[i].textContent || "";
    return s;
  }

  /** Líneas limpias de párrafos y celdas de tabla, en orden de documento. */
  async function docxLines(file) {
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const entry = zip.file("word/document.xml");
    if (!entry) throw new Error("no es un .docx válido (falta word/document.xml)");
    const xml = new DOMParser().parseFromString(await entry.async("string"), "application/xml");
    const body = xml.getElementsByTagNameNS(W_NS, "body")[0];
    if (!body) return [];

    const raw = [];
    for (const child of Array.from(body.childNodes).filter(n => n.nodeType === 1)) {
      if (child.localName === "p") {
        raw.push(paraText(child));
      } else if (child.localName === "tbl") {
        for (const tr of Array.from(child.getElementsByTagNameNS(W_NS, "tr"))) {
          const seen = new Set();
          for (const tc of Array.from(tr.getElementsByTagNameNS(W_NS, "tc"))) {
            const ct = Array.from(tc.getElementsByTagNameNS(W_NS, "p")).map(paraText).join(" ").trim();
            const key = ct.slice(0, 50);
            if (ct && !seen.has(key)) { seen.add(key); raw.push(ct); }
          }
        }
      }
    }

    const out = [];
    for (let line of raw) {
      line = line.replace(/ /g, " ").replace(/\t/g, " ");
      for (const sub of line.split("\n")) {
        const s = sub.replace(/ {2,}/g, " ").trim();
        if (s) out.push(s);
      }
    }
    return out;
  }

  // ─── PARSER (máquina de estados, igual que parse_docx en Python) ──────────

  function parseLines(lines) {
    const courses = [];
    let cur = null, descParts = [], state = null; // "prof" | "prereq" | "desc" | "ignore" | null

    const flush = () => {
      if (cur && cur.code) {
        if (descParts.length)
          cur.description = descParts.join(" ").trim().replace(BREVE_RE, "");
        cur.professor     = rstrip(cur.professor.trim(), ".,;/ ");
        cur.prerequisites = cur.prerequisites.trim();
        courses.push(cur);
      }
      cur = null; descParts = [];
    };

    const newCourse = (code, title = "", prof = "") => ({
      department: deptFromCode(code), code, title, professor: prof,
      prerequisites: "", description: "",
    });

    const labelAhead = (i, n = 3) => {
      for (const l of lines.slice(i + 1, i + 1 + n)) {
        if (PROF_RE.test(l)) return true;
        if (DESC_RE.test(l) || extractCodeAtStart(l)[0]) return false;
      }
      return false;
    };

    lines.forEach((line, i) => {
      // Encabezados de sección (línea en MAYÚSCULAS justo antes de un código)
      const nxt = lines[i + 1] || "";
      if (line.length < 60 && !/[a-zà-ÿ]/.test(line)
          && extractCodeAtStart(nxt)[0] && !extractCodeAtStart(line)[0]
          && ![PROF_RE, PREREQ_RE, DESC_RE].some(r => r.test(line))) return;

      // CLAVE: (formato de ficha)
      if (CLAVE_RE.test(line)) {
        const [code] = extractCodeAtStart(line.replace(CLAVE_RE, "").trim());
        if (code) { flush(); cur = newCourse(code); state = null; }
        return;
      }

      // NOMBRE DEL CURSO:
      if (NOMBRE_RE.test(line) && cur) {
        cur.title = line.replace(NOMBRE_RE, "").trim().toUpperCase();
        state = null;
        return;
      }

      // Línea que empieza con código
      const [code, remainder] = extractCodeAtStart(line);
      if (code && cur) {
        // Dentro de un bloque PROFESOR/PRERREQUISITOS es un prerrequisito,
        // salvo que siga una etiqueta PROFESOR nueva (→ materia nueva).
        if ((state === "prof" || state === "prereq") && !(cur.professor && labelAhead(i))) {
          let p = cur.prerequisites;
          if (p === "Ninguno") p = "";
          cur.prerequisites = p ? p + "\n" + line : line;
          state = "prereq";
          return;
        }
        // Código de materia compartida justo después del título
        if (!remainder && state === null && cur.title && !cur.professor && !cur.prerequisites) {
          cur.title += ` (${code})`;
          return;
        }
      }
      if (code) {
        flush();
        let titleText = remainder, embeddedProf = "";
        const m = PROF_SPLIT.exec(remainder);
        if (m) {
          titleText    = remainder.slice(0, m.index).trim();
          embeddedProf = remainder.slice(m.index + m[0].length).trim();
        }
        cur = newCourse(code, rstrip(titleText.toUpperCase().trim(), ".,"), rstrip(embeddedProf, ".,;"));
        state = embeddedProf ? "prof" : null;
        return;
      }

      if (!cur) return;

      if (PROF_RE.test(line)) {
        cur.professor = rstrip(line.replace(PROF_RE, "").trim(), ".,;");
        state = "prof";
        return;
      }
      if (PREREQ_RE.test(line)) {
        let val = rstrip(line.replace(PREREQ_RE, "").trim(), ".,;");
        if (!val && SIN_PRE_RE.test(line)) val = "Ninguno"; // el documento lo dice explícitamente
        cur.prerequisites = val;
        state = "prereq";
        return;
      }
      if (DESC_RE.test(line)) {
        const val = line.replace(DESC_RE, "").trim();
        descParts = val ? [val] : [];
        state = "desc";
        return;
      }
      if (STOP_RE.test(line)) { state = "ignore"; return; }
      if (state === "ignore") return;

      // Líneas de continuación
      const isProse = line.length >= 90;
      if (state === "prof") {
        if (line.length < 70) {
          const sep = cur.professor ? " / " : "";
          cur.professor = strip(cur.professor + sep + line, " /.,;");
        } else { descParts.push(line); state = "desc"; }
      } else if (state === "prereq") {
        let p = cur.prerequisites;
        if (!p || !isProse) {
          if (p === "Ninguno") p = "";
          cur.prerequisites = p ? p + "\n" + line : line;
        } else { descParts.push(line); state = "desc"; }
      } else if (state === "desc") {
        descParts.push(line);
      } else if (state === null) {
        if (!cur.title) cur.title = rstrip(line.toUpperCase(), ".,");
        else if (isProse) { descParts.push(line); state = "desc"; }
        else if (!cur.professor && !cur.prerequisites) {
          const sep = line.startsWith("(") ? " " : " — ";
          cur.title += sep + rstrip(line.toUpperCase(), ".,");
        }
      }
    });

    flush();
    return courses;
  }

  // ─── API PÚBLICA ───────────────────────────────────────────────────────────

  const normTitle = s => (s || "").replace(/[^\p{L}\p{N}_]/gu, "").toUpperCase();

  const sortCourses = arr => arr.slice().sort((a, b) => {
    const ia = DEPT_ORDER.indexOf(a.department), ib = DEPT_ORDER.indexOf(b.department);
    const da = ia < 0 ? 99 : ia, db = ib < 0 ? 99 : ib;
    return da - db || a.code.localeCompare(b.code);
  });

  /**
   * Lee varios archivos Word y devuelve las materias deduplicadas por código+título.
   * Los .doc (formato binario antiguo) no se pueden leer en el navegador.
   */
  async function parseWordFiles(fileList) {
    const files = Array.from(fileList)
      .filter(f => /\.docx?$/i.test(f.name) && !f.name.startsWith("~$"))
      .sort((a, b) => a.name.localeCompare(b.name));

    const all = [], report = [], duplicates = [], seen = new Set();

    for (const f of files) {
      if (/\.doc$/i.test(f.name)) {
        report.push({ file: f.name, status: "error", count: 0,
          message: "formato .doc no se puede leer en el navegador — guárdalo como .docx en Word" });
        continue;
      }
      try {
        const lines = await docxLines(f);
        if (lines.length && NOTES_RE.test(lines[0])) {
          report.push({ file: f.name, status: "skipped", count: 0, message: "notas al plan de estudios" });
          continue;
        }
        let added = 0;
        for (const c of parseLines(lines)) {
          const key = c.code + "|" + normTitle(c.title);
          if (seen.has(key)) { duplicates.push(`${c.code} ← ${f.name}`); continue; }
          seen.add(key); all.push(c); added++;
        }
        report.push({ file: f.name, status: added ? "ok" : "skipped", count: added,
          message: added ? "" : "no se encontraron materias" });
      } catch (e) {
        report.push({ file: f.name, status: "error", count: 0, message: e.message });
      }
    }
    return { courses: sortCourses(all), report, duplicates };
  }

  /** Construye un .xlsx (columnas A-F) con las materias. */
  function coursesToXlsx(courses) {
    const header = ["Departamento", "Código", "Título", "Profesor", "Prerrequisitos", "Descripción"];
    const rows = courses.map(c => [c.department, c.code, c.title, c.professor, c.prerequisites, c.description]
      .map(v => v || null));
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    ws["!cols"] = [22, 16, 44, 34, 48, 80].map(wch => ({ wch }));
    ws["!autofilter"] = { ref: `A1:F${rows.length + 1}` };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Optativas");
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    return new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  window.OptiParser = { parseWordFiles, coursesToXlsx, parseLines, DEPT_ORDER, PREFIX_TO_DEPT };
})();
