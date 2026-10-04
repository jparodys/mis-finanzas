// Mis Finanzas: registra gastos e ingresos en "Presupuesto Personal.xlsx" (OneDrive personal).
const C = window.CONFIG;
const X = window.XlsxRegistro;
const $ = (s) => document.querySelector(s);
const REDIRECT = location.origin + location.pathname.replace(/index\.html$/, "");
const SCOPES = "Files.ReadWrite offline_access";

const guardado = {
  leer(k, def) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : def; } catch { return def; } },
  escribir(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
  borrar(k) { try { localStorage.removeItem(k); } catch {} },
};

// ---------------------------------------------------------------- Inicio de sesión (OAuth PKCE)
function b64url(bytes) {
  return btoa(String.fromCharCode(...new Uint8Array(bytes))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
async function iniciarSesion() {
  const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
  const challenge = b64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
  const state = b64url(crypto.getRandomValues(new Uint8Array(16)));
  sessionStorage.setItem("pkce", JSON.stringify({ verifier, state }));
  const p = new URLSearchParams({
    client_id: C.CLIENT_ID, response_type: "code", redirect_uri: REDIRECT, scope: SCOPES,
    code_challenge: challenge, code_challenge_method: "S256", state, response_mode: "query", prompt: "select_account",
  });
  location.assign(`${C.LOGIN}/authorize?${p}`);
}
async function pedirToken(params) {
  const res = await fetch(`${C.LOGIN}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: C.CLIENT_ID, scope: SCOPES, ...params }),
  });
  const j = await res.json();
  if (!res.ok) throw new Error(j.error_description || j.error || "Error al iniciar sesión");
  guardado.escribir("token", {
    access: j.access_token,
    refresh: j.refresh_token,
    vence: Date.now() + (j.expires_in - 120) * 1000,
  });
}
async function completarInicioSesion() {
  const q = new URLSearchParams(location.search);
  if (!q.has("code") && !q.has("error")) return;
  history.replaceState(null, "", REDIRECT);
  if (q.has("error")) throw new Error(q.get("error_description") || q.get("error"));
  const pkce = JSON.parse(sessionStorage.getItem("pkce") || "{}");
  sessionStorage.removeItem("pkce");
  if (pkce.state !== q.get("state")) throw new Error("La respuesta de inicio de sesión no coincide. Intenta de nuevo.");
  await pedirToken({ grant_type: "authorization_code", code: q.get("code"), redirect_uri: REDIRECT, code_verifier: pkce.verifier });
}
async function token() {
  if (C.TOKEN_PRUEBA) return C.TOKEN_PRUEBA;
  const t = guardado.leer("token", null);
  if (!t) return null;
  if (Date.now() < t.vence) return t.access;
  if (!t.refresh) return null;
  try {
    await pedirToken({ grant_type: "refresh_token", refresh_token: t.refresh });
    return guardado.leer("token", null).access;
  } catch {
    guardado.borrar("token");
    return null;
  }
}
function cerrarSesion() {
  guardado.borrar("token");
  render();
}

// ---------------------------------------------------------------- OneDrive (Microsoft Graph)
async function graph(ruta, opciones = {}) {
  const t = await token();
  if (!t) throw Object.assign(new Error("Debes iniciar sesión."), { sinSesion: true });
  return fetch(C.GRAPH + ruta, { ...opciones, headers: { Authorization: "Bearer " + t, ...(opciones.headers || {}) } });
}
async function obtenerArchivo() {
  const ruta = C.RUTA_ARCHIVO.split("/").map(encodeURIComponent).join("/");
  const res = await graph(`/me/drive/root:/${ruta}`);
  if (res.status === 404) throw new Error(`No encontré "${C.RUTA_ARCHIVO}" en tu OneDrive.`);
  if (res.status === 401) { guardado.borrar("token"); throw Object.assign(new Error("Tu sesión expiró."), { sinSesion: true }); }
  if (!res.ok) throw new Error("OneDrive respondió " + res.status);
  const item = await res.json();
  const dl = await fetch(item["@microsoft.graph.downloadUrl"]);
  if (!dl.ok) throw new Error("No pude descargar el archivo (" + dl.status + ")");
  return { item, buffer: await dl.arrayBuffer() };
}
async function subirArchivo(item, blob) {
  return graph(`/me/drive/items/${item.id}/content`, {
    method: "PUT",
    headers: { "If-Match": item.eTag, "Content-Type": blob.type },
    body: blob,
  });
}

// ---------------------------------------------------------------- Estado
const estado = {
  movimiento: "Gasto",
  tipo: null,
  medio: guardado.leer("ultimoMedio", null),   // se recuerda el último medio usado
  listas: guardado.leer("listas", { gastos: C.GASTOS, ingresos: C.INGRESOS }),
  pendientes: guardado.leer("pendientes", []),
  movs: guardado.leer("movs", []),
  vista: "mes",                       // "mes" o "anio"
  mesVista: mesDe(new Date()),
  anioVista: new Date().getFullYear(),
  abierto: null,
  ocupado: false,
  sesion: false,
};

function mesDe(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}
function moverMes(mes, delta) {
  const [y, m] = mes.split("-").map(Number);
  return mesDe(new Date(y, m - 1 + delta, 1));
}

// Totales de un periodo ("AAAA-MM" para un mes, "AAAA" para un año):
// gastos, ingresos y gasto por tipo con su detalle
function resumir(movs, periodo) {
  const delMes = movs.filter((m) => m.fecha && m.fecha.startsWith(periodo));
  const suma = (tipo) => delMes.filter((m) => m.movimiento === tipo).reduce((a, m) => a + (Number(m.valor) || 0), 0);
  const porTipo = {};
  for (const m of delMes) {
    if (m.movimiento !== "Gasto") continue;
    const t = (porTipo[m.tipo || "Sin tipo"] ||= { total: 0, items: [] });
    t.total += Number(m.valor) || 0;
    t.items.push(m);
  }
  return {
    gastos: suma("Gasto"), ingresos: suma("Ingreso"),
    porTipo: Object.entries(porTipo).sort((a, b) => b[1].total - a[1].total),
  };
}

// Descarga el Excel, agrega los pendientes y lo vuelve a subir
async function sincronizar({ soloLeer = false } = {}) {
  if (estado.ocupado) return;
  estado.ocupado = true;
  aviso(soloLeer ? "Leyendo tu Excel…" : "Guardando en tu Excel…", "info");
  render();
  try {
    for (let intento = 1; intento <= 3; intento++) {
      const { item, buffer } = await obtenerArchivo();
      const libro = await X.abrirLibro(buffer);
      const reg = await X.leerRegistro(libro);
      const listas = await X.leerListas(libro).catch(() => null);
      if (listas) { estado.listas = listas; guardado.escribir("listas", listas); }

      const lote = estado.pendientes.slice();
      if (lote.length && !soloLeer) {
        X.agregarMovimientos(libro, reg, lote);
        const res = await subirArchivo(item, await X.generarArchivo(libro));
        if (res.status === 412 && intento < 3) continue;      // alguien cambió el archivo: reintenta
        if (res.status === 423) throw new Error("El archivo está bloqueado (abierto en otro lado). Ciérralo e intenta de nuevo.");
        if (!res.ok) throw new Error("OneDrive no aceptó el cambio (" + res.status + ")");
        const ids = new Set(lote.map((e) => e.id));
        estado.pendientes = estado.pendientes.filter((e) => !ids.has(e.id));
        guardado.escribir("pendientes", estado.pendientes);
        aviso(`✓ ${lote.length === 1 ? "Movimiento guardado" : lote.length + " movimientos guardados"} en Excel (fila ${reg.ultima})`, "ok");
      } else {
        aviso("Datos actualizados desde tu Excel", "ok");
      }
      estado.movs = reg.movimientos.map(({ fecha, descripcion, movimiento, tipo, valor }) => ({ fecha, descripcion, movimiento, tipo, valor }));
      guardado.escribir("movs", estado.movs);
      break;
    }
  } catch (e) {
    if (e.sinSesion) estado.sesion = false;
    const sinRed = !navigator.onLine || e instanceof TypeError;
    aviso(sinRed ? "Sin conexión: queda pendiente y se guardará cuando vuelva la red." : e.message, "error");
  } finally {
    estado.ocupado = false;
    render();
  }
}

// ---------------------------------------------------------------- Interfaz
const pesos = (n) => "$" + Math.round(n || 0).toLocaleString("es-CO");
const hoyISO = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};
const MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
const fechaCorta = (iso) => (iso ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}` : "");
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

let temporizadorAviso;
function aviso(texto, tipo = "info") {
  const el = $("#aviso");
  el.textContent = texto;
  el.className = "aviso " + tipo;
  clearTimeout(temporizadorAviso);
  if (tipo === "ok") temporizadorAviso = setTimeout(() => (el.className = "aviso oculto"), 4000);
}

function render() {
  const conSesion = estado.sesion || !!C.TOKEN_PRUEBA;
  $("#pantalla-login").hidden = conSesion;
  $("#pantalla-app").hidden = !conSesion;
  if (!conSesion) return;

  // Gasto / Ingreso
  for (const b of document.querySelectorAll("[data-mov]")) b.classList.toggle("activo", b.dataset.mov === estado.movimiento);
  document.body.dataset.mov = estado.movimiento;
  $("#titulo-tipo").textContent = estado.movimiento === "Gasto" ? "Tipo de gasto" : "Tipo de ingreso";
  const opciones = estado.movimiento === "Gasto" ? estado.listas.gastos : (estado.listas.ingresos.length ? estado.listas.ingresos : C.INGRESOS);
  if (!opciones.includes(estado.tipo)) estado.tipo = null;
  $("#tipos").innerHTML = opciones
    .map((t) => `<button type="button" class="chip${t === estado.tipo ? " activo" : ""}" data-tipo="${esc(t)}">${esc(t)}</button>`)
    .join("");

  // Medio de pago (solo para gastos)
  const esGasto = estado.movimiento === "Gasto";
  $("#bloque-medio").hidden = !esGasto;
  const medios = estado.listas.medios && estado.listas.medios.length ? estado.listas.medios : C.MEDIOS;
  if (!medios.includes(estado.medio)) estado.medio = null;
  $("#medios").innerHTML = medios
    .map((t) => `<button type="button" class="chip${t === estado.medio ? " activo" : ""}" data-medio="${esc(t)}">${esc(t)}</button>`)
    .join("");

  $("#guardar").disabled = estado.ocupado;
  $("#guardar").textContent = estado.ocupado ? "Guardando…" : `Guardar ${estado.movimiento.toLowerCase()}`;

  // Pendientes
  const n = estado.pendientes.length;
  $("#pendientes").hidden = n === 0;
  $("#pendientes-texto").textContent = `${n} ${n === 1 ? "movimiento pendiente" : "movimientos pendientes"} por subir`;

  // Resumen del mes o del año elegido (incluye lo pendiente por subir)
  const todos = [...estado.movs, ...estado.pendientes];
  const enAnio = estado.vista === "anio";
  for (const b of document.querySelectorAll("[data-vista]")) b.classList.toggle("activo", b.dataset.vista === estado.vista);
  const hoy = new Date();
  let periodo;
  if (enAnio) {
    periodo = String(estado.anioVista);
    $("#mes-titulo").textContent = `Año ${periodo}`;
    $("#mes-sig").disabled = estado.anioVista >= hoy.getFullYear();
  } else {
    periodo = estado.mesVista;
    const [y, m] = periodo.split("-");
    $("#mes-titulo").textContent = `${MESES[Number(m) - 1]} ${y}`;
    $("#mes-sig").disabled = periodo >= mesDe(hoy);
  }
  const r = resumir(todos, periodo);
  $("#kpi-gastos").textContent = pesos(r.gastos);
  $("#kpi-ingresos").textContent = pesos(r.ingresos);
  const bal = r.ingresos - r.gastos;
  $("#kpi-balance").textContent = pesos(bal);
  $("#kpi-balance").classList.toggle("negativo", bal < 0);

  // Gráfico de barras por mes (solo en la vista anual)
  $("#grafico-anio").hidden = !enAnio;
  if (enAnio) {
    const porMes = Array.from({ length: 12 }, (_, i) => resumir(todos, `${periodo}-${String(i + 1).padStart(2, "0")}`));
    const tope = Math.max(1, ...porMes.flatMap((x) => [x.gastos, x.ingresos]));
    const conGasto = porMes.filter((x) => x.gastos > 0).length;
    $("#barras-mes").innerHTML = porMes.map((x, i) => `
      <button type="button" class="col-mes" data-ir-mes="${i + 1}" aria-label="${MESES[i]}: gastos ${pesos(x.gastos)}, ingresos ${pesos(x.ingresos)}">
        <span class="col-barras"><i class="g" style="height:${(x.gastos / tope) * 100}%"></i><i class="n" style="height:${(x.ingresos / tope) * 100}%"></i></span>
        <small>${MESES[i][0].toUpperCase()}</small>
      </button>`).join("");
    $("#promedio-mes").textContent = conGasto
      ? `Gasto promedio: ${pesos(r.gastos / conGasto)} por mes (${conGasto} ${conGasto === 1 ? "mes" : "meses"} con gastos)`
      : "";
  }

  const max = r.porTipo.length ? r.porTipo[0][1].total : 1;
  const detalleAnual = (items) => {
    const meses = Array(12).fill(0);
    for (const i of items) meses[Number(i.fecha.slice(5, 7)) - 1] += Number(i.valor) || 0;
    return meses.map((v, k) => (v ? `<li><span class="fecha">${MESES[k].slice(0, 3)}</span><span class="desc"></span><b>${pesos(v)}</b></li>` : "")).join("");
  };
  const detalleMes = (items) => items.slice().reverse()
    .map((i) => `<li><span class="fecha">${fechaCorta(i.fecha)}</span><span class="desc">${esc(i.descripcion)}</span><b>${pesos(i.valor)}</b></li>`).join("");
  $("#por-tipo").innerHTML = r.porTipo.length
    ? r.porTipo.map(([t, { total, items }]) => `
      <li class="tipo${estado.abierto === t ? " abierto" : ""}" data-abrir="${esc(t)}">
        <div class="tipo-fila">
          <span class="tipo-nombre">${esc(t)}<small>${items.length} ${items.length === 1 ? "gasto" : "gastos"} · ${Math.round((total / r.gastos) * 100)}%</small></span>
          <span class="barra"><i style="width:${Math.max(4, (total / max) * 100)}%"></i></span>
          <b>${pesos(total)}</b>
        </div>
        <ul class="detalle">${enAnio ? detalleAnual(items) : detalleMes(items)}</ul>
      </li>`).join("")
    : `<li class="vacio">No hay gastos registrados en ${enAnio ? "este año" : "este mes"}</li>`;

  const recientes = todos.slice(-8).reverse();
  $("#recientes").innerHTML = recientes.length
    ? recientes.map((m) => `<li><span class="fecha">${fechaCorta(m.fecha)}</span><span class="desc">${esc(m.descripcion)}<small>${esc(m.tipo)}</small></span><b class="${m.movimiento === "Ingreso" ? "ing" : "gas"}">${m.movimiento === "Ingreso" ? "+" : "−"}${pesos(m.valor)}</b></li>`).join("")
    : `<li class="vacio">Toca ↻ para leer tu Excel</li>`;
}

function valorNumerico() {
  return Number($("#valor").value.replace(/\D/g, "")) || 0;
}

function enviarFormulario(ev) {
  ev.preventDefault();
  const valor = valorNumerico();
  const descripcion = $("#descripcion").value.trim();
  if (!estado.tipo) return aviso("Elige el tipo de " + estado.movimiento.toLowerCase(), "error");
  const esGasto = estado.movimiento === "Gasto";
  if (esGasto && !estado.medio) return aviso("Elige el medio de pago (Efectivo, Débito o Crédito)", "error");
  if (!valor) return aviso("Escribe el valor", "error");
  if (!descripcion) return aviso("Escribe una descripción", "error");
  estado.pendientes.push({
    id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now() + Math.random()),
    fecha: $("#fecha").value || hoyISO(),
    descripcion, movimiento: estado.movimiento, tipo: estado.tipo, valor,
    medio: esGasto ? estado.medio : "",
    notas: $("#notas").value.trim(),
  });
  guardado.escribir("pendientes", estado.pendientes);
  if (esGasto) guardado.escribir("ultimoMedio", estado.medio);
  $("#descripcion").value = ""; $("#valor").value = ""; $("#notas").value = "";
  estado.tipo = null;
  render();
  sincronizar();
}

function conectarEventos() {
  $("#entrar").addEventListener("click", () => {
    if (!C.CLIENT_ID || C.CLIENT_ID.startsWith("PEGA")) return alert("Falta configurar CLIENT_ID en config.js (ver LEEME.md).");
    iniciarSesion();
  });
  $("#salir").addEventListener("click", () => confirm("¿Cerrar sesión?") && cerrarSesion());
  $("#refrescar").addEventListener("click", () => sincronizar({ soloLeer: estado.pendientes.length === 0 }));
  $("#reintentar").addEventListener("click", () => sincronizar());
  for (const b of document.querySelectorAll("[data-mov]")) b.addEventListener("click", () => { estado.movimiento = b.dataset.mov; render(); });
  $("#medios").addEventListener("click", (e) => {
    const b = e.target.closest("[data-medio]");
    if (b) { estado.medio = b.dataset.medio; render(); }
  });
  $("#tipos").addEventListener("click", (e) => {
    const b = e.target.closest("[data-tipo]");
    if (b) { estado.tipo = b.dataset.tipo; render(); }
  });
  $("#valor").addEventListener("input", (e) => {
    const n = valorNumerico();
    e.target.value = n ? n.toLocaleString("es-CO") : "";
  });
  $("#form").addEventListener("submit", enviarFormulario);
  const mover = (delta) => {
    if (estado.vista === "anio") estado.anioVista += delta;
    else estado.mesVista = moverMes(estado.mesVista, delta);
    estado.abierto = null;
    render();
  };
  $("#mes-ant").addEventListener("click", () => mover(-1));
  $("#mes-sig").addEventListener("click", () => mover(1));
  for (const b of document.querySelectorAll("[data-vista]")) {
    b.addEventListener("click", () => {
      estado.vista = b.dataset.vista;
      if (estado.vista === "anio") estado.anioVista = Number(estado.mesVista.slice(0, 4));
      estado.abierto = null;
      render();
    });
  }
  $("#barras-mes").addEventListener("click", (e) => {
    const b = e.target.closest("[data-ir-mes]");
    if (!b) return;
    estado.mesVista = `${estado.anioVista}-${String(b.dataset.irMes).padStart(2, "0")}`;
    estado.vista = "mes";
    estado.abierto = null;
    render();
  });
  $("#por-tipo").addEventListener("click", (e) => {
    const li = e.target.closest("[data-abrir]");
    if (!li) return;
    estado.abierto = estado.abierto === li.dataset.abrir ? null : li.dataset.abrir;
    render();
  });
  window.addEventListener("online", () => estado.pendientes.length && sincronizar());
}

async function arrancar() {
  $("#fecha").value = hoyISO();
  conectarEventos();
  try {
    await completarInicioSesion();
  } catch (e) {
    aviso(e.message, "error");
  }
  estado.sesion = !!(await token());
  render();
  if (estado.sesion || C.TOKEN_PRUEBA) sincronizar({ soloLeer: estado.pendientes.length === 0 });
  if ("serviceWorker" in navigator && !C.TOKEN_PRUEBA) navigator.serviceWorker.register("sw.js").catch(() => {});
}
arrancar();
