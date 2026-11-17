/**
 * COMERCIAL BIRU · Conexión entre la página web y Google Sheets
 *
 * Instalación (una sola vez):
 *  1. Crea una hoja de cálculo nueva en Google Sheets.
 *  2. Menú Extensiones > Apps Script. Borra lo que haya y pega este archivo completo.
 *  3. Arriba elige la función "instalar" y pulsa Ejecutar. Acepta los permisos.
 *     Se crean las pestañas y el usuario inicial (ver pestaña Usuarios).
 *  4. Botón Implementar > Nueva implementación > tipo "Aplicación web".
 *       Ejecutar como: Yo        Quién tiene acceso: Cualquier usuario
 *  5. Copia la URL que termina en /exec y pégala en la página (Conexión con Google Sheets).
 *
 * Si cambias este código, usa Implementar > Gestionar implementaciones > Editar > Nueva versión.
 */

const TABLAS = {
  Productos: ['codigo', 'nombre', 'categoria', 'unidad', 'precio', 'stockMin', 'stockMax', 'ubicacion', 'activo'],
  Movimientos: ['id', 'fecha', 'local', 'codigo', 'tipo', 'cantidad', 'monto', 'nota', 'usuario'],
  Pedidos: ['id', 'fecha', 'solicita', 'despacha', 'items', 'estado', 'nota', 'usuario', 'actualizado'],
  Llegadas: ['id', 'fecha', 'proveedor', 'local', 'items', 'estado', 'nota', 'usuario'],
  Precios: ['id', 'fecha', 'codigo', 'anterior', 'nuevo', 'usuario', 'balanza1', 'balanza2'],
  Usuarios: ['usuario', 'nombre', 'pin', 'rol', 'local']
};
const CLAVE = { Productos: 'codigo', Usuarios: 'usuario' };

function instalar() {
  Object.keys(TABLAS).forEach(hoja_);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sobra = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (sobra && sobra.getLastRow() === 0) ss.deleteSheet(sobra);
}

function hoja_(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    const cols = TABLAS[nombre];
    sh = ss.insertSheet(nombre);
    // Todo como texto: así los códigos tipo "007" y las fechas no se alteran.
    sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#C8202F').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    if (nombre === 'Usuarios') sh.getRange(2, 1, 1, 5).setValues([['admin', 'Administrador', '1234', 'admin', 'L1']]);
  }
  return sh;
}

function leer_(nombre) {
  const v = hoja_(nombre).getDataRange().getDisplayValues();
  const cab = v[0], filas = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i].join('') === '') continue;
    const o = {};
    cab.forEach(function (c, j) { if (c) o[c] = v[i][j]; });
    filas.push(o);
  }
  return filas;
}

function responder_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return responder_({ ok: true, mensaje: 'Conexión de Comercial Biru activa.' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const usuarios = leer_('Usuarios');
    const u = usuarios.filter(function (x) {
      return x.usuario.toLowerCase() === String(req.usuario || '').toLowerCase() && String(x.pin) === String(req.pin) && x.pin !== '';
    })[0];
    if (!u) return responder_({ ok: false, error: 'Usuario o PIN incorrecto' });

    if (req.ops && req.ops.length) {
      const cerrojo = LockService.getScriptLock();
      cerrojo.waitLock(20000);
      try { aplicar_(req.ops, u); } finally { cerrojo.releaseLock(); }
    }

    const db = {};
    Object.keys(TABLAS).forEach(function (t) { db[t] = leer_(t); });
    db.Usuarios.forEach(function (x) { x.pin = ''; });
    if (u.rol !== 'admin') db.Usuarios = db.Usuarios.filter(function (x) { return x.usuario === u.usuario; });
    const yo = db.Usuarios.filter(function (x) { return x.usuario === u.usuario; })[0] || { usuario: u.usuario, nombre: u.nombre, rol: u.rol, local: u.local, pin: '' };
    return responder_({ ok: true, user: yo, db: db });
  } catch (err) {
    return responder_({ ok: false, error: String(err.message || err) });
  }
}

function permiso_(op, u) {
  if (u.rol === 'admin') return;
  const t = op.t;
  let ok = true;
  if (t === 'Usuarios') ok = false;
  else if (t === 'Productos') ok = op.op === 'upd' && Object.keys(op.f || {}).join() === 'ubicacion';
  else if (t === 'Precios') ok = op.op === 'upd' && Object.keys(op.f || {}).every(function (c) { return c === 'balanza1' || c === 'balanza2'; });
  else if (t === 'Movimientos') ok = op.op === 'add';
  else if (t === 'Pedidos' && op.op === 'upd') ok = ['aprobado', 'rechazado'].indexOf((op.f || {}).estado) < 0;
  if (!ok) throw new Error('No tienes permiso para esta acción. Pídeselo al administrador.');
}

function aplicar_(ops, u) {
  const nuevos = {};
  ops.forEach(function (op) {
    const cols = TABLAS[op.t];
    if (!cols) throw new Error('Tabla desconocida: ' + op.t);
    permiso_(op, u);
    const sh = hoja_(op.t), clave = CLAVE[op.t] || 'id';

    if (op.op === 'add') {
      (nuevos[op.t] = nuevos[op.t] || []).push(cols.map(function (c) { return op.row[c] == null ? '' : String(op.row[c]); }));
      return;
    }
    const ultima = sh.getLastRow();
    if (ultima < 2) return;
    const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    const ids = sh.getRange(2, cab.indexOf(clave) + 1, ultima - 1, 1).getDisplayValues().map(function (r) { return r[0]; });
    const i = ids.indexOf(String(op.id));
    if (i < 0) return;

    if (op.op === 'del') { sh.deleteRow(i + 2); return; }
    Object.keys(op.f || {}).forEach(function (c) {
      const ci = cab.indexOf(c);
      if (ci < 0 || c === clave) return;
      if (op.t === 'Usuarios' && c === 'pin' && !op.f[c]) return;
      sh.getRange(i + 2, ci + 1).setNumberFormat('@').setValue(String(op.f[c]));
    });
  });

  Object.keys(nuevos).forEach(function (t) {
    const sh = hoja_(t), filas = nuevos[t];
    sh.getRange(sh.getLastRow() + 1, 1, filas.length, filas[0].length).setNumberFormat('@').setValues(filas);
  });
}
