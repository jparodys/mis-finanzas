// Lee y modifica "Presupuesto Parodys.xlsx" directamente en su XML.
// Solo se tocan las celdas de entrada de la hoja Registro (y la marca de
// recálculo del libro); gráficos, fórmulas, formatos y validaciones quedan intactos.

const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

const HOJA_REGISTRO = "Registro";
const HOJA_LISTAS = "Listas";
// Encabezados de la fila de títulos en Registro (sin tildes, en minúscula)
const CAMPOS = {
  fecha: "fecha",
  descripcion: "descripcion",
  movimiento: "movimiento",
  tipo: "tipo de gasto",
  medio: "medio de pago",
  valor: "valor",
  notas: "notas",
};
const OBLIGATORIOS = ["fecha", "descripcion", "movimiento", "tipo", "valor"];

function norm(s) {
  return (s ?? "").toString().normalize("NFD").replace(/[̀-ͯ]/g, "").trim().toLowerCase();
}
function colToNum(c) {
  let n = 0;
  for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}
function numToCol(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function splitRef(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  return [m[1], Number(m[2])];
}

// Fecha "AAAA-MM-DD" -> número de serie de Excel
function fechaASerial(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
}
function serialAFecha(n) {
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);
  return d.toISOString().slice(0, 10);
}

function serializar(doc) {
  const s = new XMLSerializer().serializeToString(doc);
  return s.startsWith("<?xml") ? s : XML_DECL + s;
}

async function abrirLibro(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const leer = async (p) => {
    const f = zip.file(p);
    if (!f) return null;
    return new DOMParser().parseFromString(await f.async("string"), "application/xml");
  };
  const wbDoc = await leer("xl/workbook.xml");
  const rels = await leer("xl/_rels/workbook.xml.rels");
  const destino = {};
  for (const r of rels.getElementsByTagName("Relationship")) destino[r.getAttribute("Id")] = r.getAttribute("Target");
  const hojas = {};
  for (const s of wbDoc.getElementsByTagNameNS(NS, "sheet")) {
    const t = destino[s.getAttributeNS(RNS, "id")];
    hojas[s.getAttribute("name")] = t.startsWith("/") ? t.slice(1) : "xl/" + t;
  }
  const compartidos = [];
  const ss = await leer("xl/sharedStrings.xml");
  if (ss) {
    for (const si of ss.getElementsByTagNameNS(NS, "si")) {
      // Ignora textos fonéticos (rPh), solo concatena los <t> directos de <si> o de <r>
      let txt = "";
      for (const n of si.childNodes) {
        if (n.localName === "t") txt += n.textContent;
        else if (n.localName === "r") for (const t of n.getElementsByTagNameNS(NS, "t")) txt += t.textContent;
      }
      compartidos.push(txt);
    }
  }
  return { zip, wbDoc, hojas, compartidos, leer };
}

function valorCelda(c, compartidos) {
  if (!c) return null;
  const t = c.getAttribute("t");
  const v = c.getElementsByTagNameNS(NS, "v")[0];
  if (t === "inlineStr") {
    return [...c.getElementsByTagNameNS(NS, "t")].map((x) => x.textContent).join("");
  }
  if (!v || v.textContent === "") return null;
  if (t === "s") return compartidos[Number(v.textContent)] ?? "";
  if (t === "str" || t === "e") return v.textContent;
  if (t === "b") return v.textContent === "1";
  return Number(v.textContent);
}

function indexarFilas(doc) {
  const sheetData = doc.getElementsByTagNameNS(NS, "sheetData")[0];
  const filas = new Map();
  for (const row of sheetData.getElementsByTagNameNS(NS, "row")) filas.set(Number(row.getAttribute("r")), row);
  return { sheetData, filas };
}

function celdasDeFila(row) {
  const m = new Map();
  for (const c of row.getElementsByTagNameNS(NS, "c")) m.set(splitRef(c.getAttribute("r"))[0], c);
  return m;
}

async function leerRegistro(libro) {
  const ruta = libro.hojas[HOJA_REGISTRO];
  if (!ruta) throw new Error(`No encontré la hoja "${HOJA_REGISTRO}" en el archivo.`);
  const doc = await libro.leer(ruta);
  const { sheetData, filas } = indexarFilas(doc);

  // Busca la fila de encabezados (la primera que tenga "Fecha" y "Valor")
  let filaTitulos = null;
  const cols = {};
  for (const [r, row] of [...filas].sort((a, b) => a[0] - b[0])) {
    if (r > 40) break;
    const encontrados = {};
    for (const [col, c] of celdasDeFila(row)) {
      const txt = norm(valorCelda(c, libro.compartidos));
      for (const [campo, titulo] of Object.entries(CAMPOS)) if (txt === titulo) encontrados[campo] = col;
    }
    if (encontrados.fecha && encontrados.valor) {
      filaTitulos = r;
      Object.assign(cols, encontrados);
      break;
    }
  }
  if (!filaTitulos) throw new Error("No encontré la fila de encabezados (Fecha / Valor) en Registro.");
  const faltan = OBLIGATORIOS.filter((k) => !cols[k]);
  if (faltan.length) throw new Error("Faltan columnas en Registro: " + faltan.map((k) => CAMPOS[k]).join(", "));

  // Movimientos existentes y última fila usada (según las columnas de entrada)
  const movimientos = [];
  let ultima = filaTitulos;
  for (const [r, row] of filas) {
    if (r <= filaTitulos) continue;
    const cs = celdasDeFila(row);
    const val = (k) => (cols[k] ? valorCelda(cs.get(cols[k]), libro.compartidos) : null);
    const fecha = val("fecha"), desc = val("descripcion"), valor = val("valor");
    if (fecha == null && (desc == null || desc === "") && valor == null) continue;
    ultima = Math.max(ultima, r);
    movimientos.push({
      fila: r,
      fecha: typeof fecha === "number" ? serialAFecha(fecha) : null,
      descripcion: desc ?? "",
      movimiento: val("movimiento") ?? "",
      tipo: val("tipo") ?? "",
      medio: val("medio") ?? "",
      valor: typeof valor === "number" ? valor : Number(valor) || 0,
      notas: val("notas") ?? "",
    });
  }
  movimientos.sort((a, b) => a.fila - b.fila);
  return { ruta, doc, sheetData, filas, cols, filaTitulos, ultima, movimientos };
}

async function leerListas(libro) {
  const ruta = libro.hojas[HOJA_LISTAS];
  if (!ruta) return null;
  const doc = await libro.leer(ruta);
  const { filas } = indexarFilas(doc);
  const fila1 = filas.get(1);
  if (!fila1) return null;
  const colDe = {};
  for (const [col, c] of celdasDeFila(fila1)) colDe[norm(valorCelda(c, libro.compartidos))] = col;
  const columna = (titulo) => {
    const col = colDe[titulo];
    if (!col) return [];
    const out = [];
    for (let r = 2; r < 200; r++) {
      const row = filas.get(r);
      const v = row ? valorCelda(celdasDeFila(row).get(col), libro.compartidos) : null;
      if (v == null || v === "") break;
      out.push(String(v));
    }
    return out;
  };
  const gastos = columna("tipos de gasto");
  const ingresos = columna("tipos de ingreso");
  const medios = columna("medios de pago");
  return gastos.length ? { gastos, ingresos, medios } : null;
}

function obtenerFila(reg, r) {
  let row = reg.filas.get(r);
  if (row) return row;
  row = reg.doc.createElementNS(NS, "row");
  row.setAttribute("r", String(r));
  const siguiente = [...reg.filas.keys()].filter((k) => k > r).sort((a, b) => a - b)[0];
  reg.sheetData.insertBefore(row, siguiente ? reg.filas.get(siguiente) : null);
  reg.filas.set(r, row);
  return row;
}

function obtenerCelda(doc, row, col) {
  const r = row.getAttribute("r");
  let despues = null;
  for (const c of row.getElementsByTagNameNS(NS, "c")) {
    const [cc] = splitRef(c.getAttribute("r"));
    if (cc === col) return c;
    if (colToNum(cc) > colToNum(col)) { despues = c; break; }
  }
  const c = doc.createElementNS(NS, "c");
  c.setAttribute("r", col + r);
  row.insertBefore(c, despues);
  return c;
}

function escribirCelda(doc, row, col, valor) {
  const c = obtenerCelda(doc, row, col);
  while (c.firstChild) c.removeChild(c.firstChild);
  c.removeAttribute("t");
  if (valor == null || valor === "") return;
  if (typeof valor === "number") {
    const v = doc.createElementNS(NS, "v");
    v.textContent = String(valor);
    c.appendChild(v);
  } else {
    c.setAttribute("t", "inlineStr");
    const is = doc.createElementNS(NS, "is");
    const t = doc.createElementNS(NS, "t");
    t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    t.textContent = String(valor);
    is.appendChild(t);
    c.appendChild(is);
  }
}

// entradas: [{fecha:"AAAA-MM-DD", descripcion, movimiento, tipo, medio, valor:Number, notas}]
function agregarMovimientos(libro, reg, entradas) {
  const { doc, cols } = reg;
  for (const e of entradas) {
    const r = ++reg.ultima;
    const row = obtenerFila(reg, r);
    escribirCelda(doc, row, cols.fecha, fechaASerial(e.fecha));
    escribirCelda(doc, row, cols.descripcion, e.descripcion);
    escribirCelda(doc, row, cols.movimiento, e.movimiento);
    escribirCelda(doc, row, cols.tipo, e.tipo);
    if (cols.medio) escribirCelda(doc, row, cols.medio, e.medio || "");
    escribirCelda(doc, row, cols.valor, Number(e.valor));
    if (cols.notas) escribirCelda(doc, row, cols.notas, e.notas || "");
    reg.movimientos.push({ fila: r, ...e });
  }
  // Ajusta el rango declarado de la hoja si crecimos más allá
  const dim = doc.getElementsByTagNameNS(NS, "dimension")[0];
  if (dim) {
    const [ini, fin] = dim.getAttribute("ref").split(":");
    if (fin) {
      const [cf, rf] = splitRef(fin);
      if (reg.ultima > rf) dim.setAttribute("ref", `${ini}:${cf}${reg.ultima}`);
    }
  }
  libro.zip.file(reg.ruta, serializar(doc));

  // Pide a Excel recalcular todo al abrir (Año, Mes, consolidado, dashboard)
  const calc = libro.wbDoc.getElementsByTagNameNS(NS, "calcPr")[0];
  if (calc) {
    calc.setAttribute("fullCalcOnLoad", "1");
    libro.zip.file("xl/workbook.xml", serializar(libro.wbDoc));
  }
}

async function generarArchivo(libro) {
  return libro.zip.generateAsync({
    type: "blob",
    compression: "DEFLATE",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

window.XlsxRegistro = { abrirLibro, leerRegistro, leerListas, agregarMovimientos, generarArchivo };
