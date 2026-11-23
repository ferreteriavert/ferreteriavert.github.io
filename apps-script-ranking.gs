/**
 * RANKING DE LOS JUEGOS DE FERRETERÍA VERT — copia permanente en Google
 * --------------------------------------------------------------------
 * Pega este código en Extensiones → Apps Script de la hoja de cálculo.
 *
 * - Pestañas Quiz, Pachinko y Cometornillos: todas las partidas, una por fila (el historial).
 * - Pestaña Ranking: se ordena sola con los 50 mejores de cada juego.
 * - Los juegos leen el ranking de aquí, así que no se pierde aunque se cambie la página o la pantalla.
 * - Lo que el encargado borra en la pantalla no se elimina: se marca "SÍ" en la columna Oculto.
 *   Para recuperar una partida, borra ese "SÍ".
 */

const ID_HOJA = "1UtQEY0Qh0dK8FBPWbgGM_g-2OoMacfhifGFfJCTLlZs";
const CLAVE = "ferreteria-vert";               // la misma que en config-google.js
const JUEGOS = ["Quiz", "Pachinko", "Cometornillos"];
const CABECERA = ["Fecha de la partida", "Nombre", "Puntos", "Detalle", "Recibido", "Id", "Oculto"];
const EN_RANKING = 50;                         // puestos de la pestaña Ranking

/* ---------- Recibir partidas y borrados ---------- */
function doPost(e) {
  const cerrojo = LockService.getScriptLock();
  cerrojo.waitLock(20000);
  try {
    const datos = JSON.parse(e.postData.contents);
    if (datos.clave !== CLAVE) return respuesta({ ok: false, error: "clave incorrecta" });
    const libro = preparar();
    if (datos.accion === "borrar") return respuesta(borrar(libro, datos.borrar || []));
    return respuesta(guardarFilas(libro, datos.filas || []));
  } catch (err) {
    return respuesta({ ok: false, error: String(err) });
  } finally {
    cerrojo.releaseLock();
  }
}

function guardarFilas(libro, filasNuevas) {
  let nuevas = 0;
  JUEGOS.forEach(juego => {
    const deEste = filasNuevas.filter(f => f.juego === juego);
    if (!deEste.length) return;
    const hoja = libro.getSheetByName(juego);
    const ids = new Set(leerIds(hoja));
    const filas = [];
    deEste.forEach(f => {
      if (ids.has(String(f.id))) return;
      ids.add(String(f.id));
      filas.push([new Date(f.fecha), limpiar(f.nombre), Number(f.puntos) || 0, limpiar(f.detalle), new Date(), String(f.id), ""]);
    });
    if (filas.length) {
      hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, CABECERA.length).setValues(filas);
      nuevas += filas.length;
    }
  });
  return { ok: true, nuevas: nuevas };
}

function borrar(libro, lista) {
  let ocultas = 0;
  JUEGOS.forEach(juego => {
    const deEste = lista.filter(b => b.juego === juego);
    if (!deEste.length) return;
    const hoja = libro.getSheetByName(juego);
    const n = hoja.getLastRow() - 1;
    if (n < 1) return;
    const ids = hoja.getRange(2, 6, n, 1).getValues().map(r => String(r[0]));
    const fechas = hoja.getRange(2, 1, n, 1).getValues().map(r => r[0]);
    const oculto = hoja.getRange(2, 7, n, 1).getValues();
    const porId = new Set(deEste.filter(b => b.id).map(b => String(b.id)));
    const todoHasta = deEste.filter(b => b.todo).reduce((m, b) => Math.max(m, Number(b.desde) || 0), 0);
    ids.forEach((id, i) => {
      const f = fechas[i] instanceof Date ? fechas[i].getTime() : 0;
      if (oculto[i][0] !== "SÍ" && (porId.has(id) || (todoHasta && f <= todoHasta))) { oculto[i][0] = "SÍ"; ocultas++; }
    });
    hoja.getRange(2, 7, n, 1).setValues(oculto);
  });
  return { ok: true, ocultas: ocultas };
}

/* ---------- Enviar el ranking a los juegos ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.accion !== "ranking") {
    return respuesta({ ok: true, mensaje: "La copia del ranking de Ferretería Vert está funcionando" });
  }
  if (p.clave !== CLAVE) return respuesta({ ok: false, error: "clave incorrecta" });
  const n = Math.min(500, Number(p.n) || 100);
  const libro = preparar();
  const ranking = {};
  JUEGOS.forEach(juego => {
    const hoja = libro.getSheetByName(juego);
    const filas = hoja.getLastRow() - 1;
    const datos = filas > 0 ? hoja.getRange(2, 1, filas, CABECERA.length).getValues() : [];
    ranking[juego] = datos
      .filter(r => r[1] !== "" && r[6] !== "SÍ")
      .map(r => {
        const id = String(r[5]);
        const f = Number(id.split("-").pop()) || (r[0] instanceof Date ? r[0].getTime() : 0);
        return { n: String(r[1]).replace(/^'/, ""), p: Number(r[2]) || 0, t: r[3] ? String(r[3]) : "", f: f };
      })
      .sort((a, b) => b.p - a.p || a.f - b.f)
      .slice(0, n);
  });
  return respuesta({ ok: true, ranking: ranking });
}

/* ---------- Preparar la hoja (pestañas y ranking automático) ---------- */
function preparar() {
  const libro = SpreadsheetApp.openById(ID_HOJA);
  JUEGOS.forEach(juego => {
    let hoja = libro.getSheetByName(juego);
    if (!hoja) {
      hoja = libro.insertSheet(juego);
      hoja.getRange(1, 1, 1, CABECERA.length).setValues([CABECERA]).setFontWeight("bold");
      hoja.setFrozenRows(1);
      hoja.getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm");
      hoja.getRange("E:E").setNumberFormat("dd/mm/yyyy hh:mm");
      hoja.hideColumns(6);
    } else if (hoja.getRange(1, 7).getValue() !== "Oculto") {
      hoja.getRange(1, 7).setValue("Oculto").setFontWeight("bold");   // hojas de la primera versión
    }
  });
  if (!libro.getSheetByName("Ranking")) crearRanking(libro);
  return libro;
}

function crearRanking(libro) {
  const hoja = libro.insertSheet("Ranking", 0);
  hoja.getRange("A1").setValue("🏆 RANKING DE LOS JUEGOS — se actualiza solo").setFontWeight("bold").setFontSize(14);
  JUEGOS.forEach((juego, i) => {
    const c = 1 + i * 5;                         // columnas A, F, K
    hoja.getRange(2, c).setValue(juego.toUpperCase()).setFontWeight("bold").setFontColor("#ff0000");
    hoja.getRange(3, c, 1, 4).setValues([["Puesto", "Nombre", "Puntos", "Fecha"]]).setFontWeight("bold").setBackground("#000000").setFontColor("#ffffff");
    const ref = "'" + juego + "'!A2:G";
    hoja.getRange(4, c + 1).setFormula(
      '=IFERROR(QUERY(' + ref + ', "select B, C, A where B is not null and (G is null or G <> \'SÍ\') order by C desc, A asc limit ' + EN_RANKING + '", 0), "")');
    const col = String.fromCharCode(64 + c + 1);
    hoja.getRange(4, c).setFormula('=ARRAYFORMULA(IF(LEN(' + col + '4:' + col + (3 + EN_RANKING) + '), ROW(' + col + '4:' + col + (3 + EN_RANKING) + ') - 3, ""))');
    hoja.getRange(4, c + 3, EN_RANKING, 1).setNumberFormat("dd/mm/yyyy");
    hoja.setColumnWidth(c + 1, 150);
  });
  hoja.setFrozenRows(3);
}

/* Menú en la hoja: por si quieres rehacer la pestaña Ranking */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Juegos").addItem("Rehacer pestaña Ranking", "rehacerRanking").addToUi();
}
function rehacerRanking() {
  const libro = SpreadsheetApp.openById(ID_HOJA);
  const vieja = libro.getSheetByName("Ranking");
  if (vieja) libro.deleteSheet(vieja);
  preparar();
}

function leerIds(hoja) {
  const n = hoja.getLastRow() - 1;
  return n > 0 ? hoja.getRange(2, 6, n, 1).getValues().map(r => String(r[0])) : [];
}
function limpiar(t) {
  t = String(t || "").slice(0, 40);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}
function respuesta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
