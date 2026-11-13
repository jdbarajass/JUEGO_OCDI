/**
 * Apps Script – Juegos OCDI (Mundial de Preguntas + La Ruta de Papá Noel)
 * Un solo Google Sheet con una pestaña de participantes y otra de preguntas POR JUEGO.
 *
 * Parámetro "juego": 'futbol' (por defecto, compatible con futbol.html) o 'navidad'.
 *   GET  ?juego=X                       → lista de participantes
 *   GET  ?action=getPreguntas&juego=X   → banco de preguntas
 *   GET  ?action=deleteAll&juego=X      → borra todos los participantes de ese juego
 *   GET  ?action=deleteRow&row=N&juego=X→ borra una fila de participantes
 *   GET  ?action=ping&juego=X           → {ok, v:2, juego}  (lo usa navidad.html para saber que este script está activo)
 *   POST {juego, nombre, puntaje, ...}  → agrega un participante
 *   POST {juego, action:'savePreguntas', preguntas:[...]} → reemplaza el banco de preguntas
 *
 * Publicar cambios SIN cambiar la URL: Implementar → Administrar implementaciones → ✏️ →
 * Versión: "Nueva versión" → Implementar.  (Una "Nueva implementación" crea otra URL.)
 */

// Si este script NO está vinculado al Sheet (no se creó desde Extensiones → Apps Script), pega aquí el ID del Sheet.
const SHEET_ID = '';

// Nombres de pestañas: se usa la primera que exista; si ninguna existe, se crea la primera de la lista.
const HOJAS = {
  futbol:  { part: ['Participantes Futbol', 'Participantes Fútbol', 'Hoja 1'],
             preg: ['Preguntas Futbol', 'Preguntas Fútbol', 'Preguntas'] },
  navidad: { part: ['Participantes Navidad'],
             preg: ['Preguntas Navidad'] }
};
const CAB_PART     = ['nombre', 'puntaje', 'correctas', 'incorrectas', 'fecha', 'tiempo_segundos'];
const CAB_PART_NAV = CAB_PART.concat(['puntos']);
const CAB_PREG     = ['id', 'texto', 'op0', 'op1', 'op2', 'op3', 'correcta', 'activa', 'fallos'];

function libro_() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function juegoDe_(v) { return String(v || '').toLowerCase() === 'navidad' ? 'navidad' : 'futbol'; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function hoja_(juego, tipo) {
  const ss = libro_();
  const nombres = HOJAS[juego][tipo];
  for (const n of nombres) { const h = ss.getSheetByName(n); if (h) return h; }
  const h = ss.insertSheet(nombres[0]);
  const cab = tipo === 'preg' ? CAB_PREG : (juego === 'navidad' ? CAB_PART_NAV : CAB_PART);
  h.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
  h.setFrozenRows(1);
  return h;
}

// Filas → objetos usando la fila 1 como encabezados. _row = número de fila real (para deleteRow).
function leer_(h) {
  const v = h.getDataRange().getValues();
  if (v.length < 2) return [];
  const cab = v[0].map(c => String(c).trim());
  const out = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i].every(c => c === '' || c === null)) continue;
    const o = { _row: i + 1 };
    cab.forEach((c, k) => { if (c) o[c] = v[i][k]; });
    out.push(o);
  }
  return out;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const juego = juegoDe_(p.juego);
  const action = p.action || '';
  try {
    if (action === 'ping') return json_({ ok: true, v: 2, juego: juego });
    if (action === 'getPreguntas') return json_({ ok: true, juego: juego, data: leer_(hoja_(juego, 'preg')) });
    if (action === 'deleteAll') {
      const h = hoja_(juego, 'part');
      if (h.getLastRow() > 1) h.deleteRows(2, h.getLastRow() - 1);
      return json_({ ok: true, juego: juego });
    }
    if (action === 'deleteRow') {
      const r = parseInt(p.row, 10), h = hoja_(juego, 'part');
      if (r >= 2 && r <= h.getLastRow()) h.deleteRow(r);
      return json_({ ok: true, juego: juego });
    }
    return json_({ ok: true, juego: juego, data: leer_(hoja_(juego, 'part')) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const juego = juegoDe_(b.juego);

    if (b.action === 'savePreguntas') {
      const h = hoja_(juego, 'preg');
      const qs = Array.isArray(b.preguntas) ? b.preguntas : [];
      const filas = [CAB_PREG].concat(qs.map(q => {
        const o = q.options || [];
        return [q.id || '', q.text || '', o[0] || '', o[1] || '', o[2] || '', o[3] || '',
                Number(q.correct) || 0, q.active !== false, Number(q.fails) || 0];
      }));
      h.clearContents();
      h.getRange(1, 1, filas.length, CAB_PREG.length).setValues(filas);
      return json_({ ok: true, juego: juego });
    }

    // Participante: se escribe según los encabezados de la pestaña (agrega columnas que falten, p. ej. "puntos")
    const h = hoja_(juego, 'part');
    let cab = h.getRange(1, 1, 1, Math.max(1, h.getLastColumn())).getValues()[0].map(c => String(c).trim());
    const esperadas = juego === 'navidad' ? CAB_PART_NAV : CAB_PART;
    esperadas.forEach(c => {
      if (cab.indexOf(c) === -1) {
        const col = cab.filter(x => x).length + 1;
        h.getRange(1, col).setValue(c).setFontWeight('bold');
        cab = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(x => String(x).trim());
      }
    });
    h.appendRow(cab.map(c => (c && b[c] !== undefined) ? b[c] : ''));
    return json_({ ok: true, juego: juego });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
