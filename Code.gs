/**
 * FORMULARIO DE REGISTRO - CHOHO
 * Google Sheets como base de datos:
 *   - Pestaña "Empresas":  base de datos de empresas autorizadas (NIT -> nombre).
 *   - Pestaña "Registros": administradores y asesores registrados por cada empresa.
 *
 * Flujo: la empresa escribe su NIT (sin dígito de verificación), el sistema trae
 * automáticamente el nombre registrado y, al ingresar, se habilita el registro
 * del administrador y sus asesores.
 *
 * Cómo instalar:
 * 1. Abre la hoja de cálculo de Google > Extensiones > Apps Script.
 * 2. Pega este código en Code.gs y crea un archivo HTML llamado "Index" con el otro archivo.
 * 3. Ejecuta una vez "cargarDatosDePrueba" (crea las pestañas y 5 empresas de prueba).
 * 4. Implementar > Nueva implementación > Aplicación web
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquier usuario
 */
 
// Pestañas
const SHEET_REGISTROS = 'Registros';
const SHEET_EMPRESAS = 'Empresas';
 
// Límite de asesores por administrador
const MAX_ASESORES = 30;
 
// Colores de los encabezados
const HEADER_BG = '#0B2A5B';
const HEADER_FONT = '#FFFFFF';
 
// Columnas de la pestaña Empresas
const COLS_EMPRESAS = ['NIT', 'Nombre de la empresa', 'Ciudad', 'Activo (SI/NO)', 'Fecha de creación'];
 
// Columnas de la pestaña Registros
const COLS_REGISTROS = [
  'NIT',
  'Nombre de la empresa',
  'Nombre del administrador',
  'Cantidad de asesores',
  'Nombre del asesor',
  'Numero de contacto del asesor',
  'Fecha de registro',
  'ID registro'
];
 
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Formulario de Registro | CHOHO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
 
/* ======================= CONFIGURACIÓN DE LA BASE ======================= */
 
/** Crea las pestañas (si no existen) sin agregar datos. */
function configurarHojas() {
  getEmpresasSheet_();
  getRegistrosSheet_();
}
 
/**
 * DATOS DE PRUEBA: crea las pestañas y carga 5 empresas ficticias.
 * Se puede ejecutar varias veces: no duplica NITs que ya existan.
 */
function cargarDatosDePrueba() {
  const sh = getEmpresasSheet_();
  const hoy = new Date();
 
  const pruebas = [
    ['900111222', 'COMERCIALIZADORA LOS ANDES SAS',   'Bogotá',       'SI', hoy],
    ['901234567', 'DISTRIBUCIONES EL PORVENIR LTDA',   'Medellín',     'SI', hoy],
    ['900987654', 'INVERSIONES CARIBE MOTOS SAS',      'Barranquilla', 'SI', hoy],
    ['901555888', 'REPUESTOS DEL PACÍFICO SAS',        'Cali',         'SI', hoy],
    ['900333444', 'MOTOPARTES LA SABANA SAS',          'Bucaramanga',  'SI', hoy]
  ];
 
  const existentes = new Set(leerEmpresas_().map(function (e) { return e.nit; }));
  const nuevas = pruebas.filter(function (p) { return !existentes.has(p[0]); });
 
  if (nuevas.length) {
    const inicio = sh.getLastRow() + 1;
    sh.getRange(inicio, 1, nuevas.length, 1).setNumberFormat('@');
    sh.getRange(inicio, 1, nuevas.length, COLS_EMPRESAS.length).setValues(nuevas)
      .setBorder(true, true, true, true, true, true);
    sh.getRange(inicio, 5, nuevas.length, 1).setNumberFormat('dd/MM/yyyy');
  }
  getRegistrosSheet_();
 
  SpreadsheetApp.getActiveSpreadsheet().toast(
    nuevas.length + ' empresa(s) de prueba cargada(s).', 'CHOHO', 5
  );
}
 
function estiloEncabezado_(range) {
  range.setBackground(HEADER_BG)
    .setFontColor(HEADER_FONT)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}
 
function getEmpresasSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_EMPRESAS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_EMPRESAS);
    estiloEncabezado_(sh.getRange(1, 1, 1, COLS_EMPRESAS.length).setValues([COLS_EMPRESAS]));
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, COLS_EMPRESAS.length, 220);
    sh.setColumnWidth(2, 320);
    sh.getRange('A:A').setNumberFormat('@');
    const regla = SpreadsheetApp.newDataValidation().requireValueInList(['SI', 'NO'], true).build();
    sh.getRange('D2:D').setDataValidation(regla);
  }
  return sh;
}
 
function getRegistrosSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_REGISTROS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_REGISTROS);
 
    sh.getRange('E1:F1').merge()
      .setValue('Según la cantidad se repite este dato')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
 
    estiloEncabezado_(sh.getRange(2, 1, 1, COLS_REGISTROS.length).setValues([COLS_REGISTROS]));
    sh.setFrozenRows(2);
    sh.setColumnWidths(1, COLS_REGISTROS.length, 220);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('F:F').setNumberFormat('@');
  }
  return sh;
}
 
/* ======================= CONSULTA DE EMPRESAS ======================= */
 
/** NIT solo números, sin dígito de verificación ("900.123.456-7" -> "900123456"). */
function normalizarNit_(nit) {
  return String(nit || '').replace(/[\s.]/g, '').split('-')[0].replace(/\D/g, '');
}
 
/** Lee toda la base de empresas. */
function leerEmpresas_() {
  const sh = getEmpresasSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 4).getDisplayValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return {
        nit: normalizarNit_(r[0]),
        empresa: String(r[1]).trim(),
        ciudad: String(r[2]).trim(),
        activa: String(r[3] || 'SI').trim().toUpperCase() !== 'NO'
      };
    });
}
 
/** Busca una empresa activa por NIT. Devuelve {nit, empresa, ciudad} o null. */
function buscarEmpresaPorNit_(nit) {
  const n = normalizarNit_(nit);
  if (!n) return null;
  const e = leerEmpresas_().find(function (x) { return x.nit === n; });
  if (!e || !e.activa) return null;
  return { nit: e.nit, empresa: e.empresa, ciudad: e.ciudad };
}
 
/**
 * Llamado desde el formulario al escribir el NIT.
 * Devuelve {encontrada: true, nit, empresa, ciudad} o {encontrada: false, mensaje}.
 */
function consultarNit(nit) {
  const limpio = String(nit || '').trim();
  if (!/^\d{6,10}$/.test(limpio)) {
    return { encontrada: false, mensaje: 'Ingresa el NIT solo con números, sin dígito de verificación.' };
  }
  const e = buscarEmpresaPorNit_(limpio);
  if (!e) {
    return { encontrada: false, mensaje: 'Este NIT no está registrado o está inactivo.' };
  }
  return { encontrada: true, nit: e.nit, empresa: e.empresa, ciudad: e.ciudad };
}
 
/* ======================= GUARDAR ======================= */
 
/**
 * data = { nit, administrador, cantidad, asesores: [{nombre, telefono}, ...] }
 * El nombre de la empresa se toma siempre de la base de datos, no del navegador.
 */
function guardarRegistro(data) {
  if (!data) throw new Error('No se recibieron datos.');
 
  const empresa = buscarEmpresaPorNit_(data.nit);
  if (!empresa) throw new Error('El NIT no está registrado. Vuelve a ingresar.');
 
  const administrador = String(data.administrador || '').trim();
  const cantidad = parseInt(data.cantidad, 10);
  const asesores = Array.isArray(data.asesores) ? data.asesores : [];
 
  if (!administrador) throw new Error('El nombre del administrador es obligatorio.');
  if (!cantidad || cantidad < 1 || cantidad > MAX_ASESORES) {
    throw new Error('La cantidad de asesores debe estar entre 1 y ' + MAX_ASESORES + '.');
  }
  if (asesores.length !== cantidad) {
    throw new Error('La cantidad de asesores no coincide con los datos enviados.');
  }
 
  const limpios = asesores.map(function (a, i) {
    const nombre = String(a.nombre || '').trim();
    const telefono = String(a.telefono || '').replace(/\D/g, '');
    if (!nombre) throw new Error('Falta el nombre del asesor ' + (i + 1) + '.');
    if (!/^\d{10}$/.test(telefono)) {
      throw new Error('El número del asesor ' + (i + 1) + ' debe tener 10 dígitos.');
    }
    return { nombre: nombre, telefono: telefono };
  });
 
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getRegistrosSheet_();
    const fecha = new Date();
    // ID único del registro: une todas las filas de asesores de un mismo envío
    const id = 'REG-' + Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss')
      + '-' + Math.floor(Math.random() * 1000).toString().padStart(3, '0');
 
    // Empresa, administrador y cantidad solo en la primera fila; una fila por asesor.
    // El ID va en todas las filas para poder filtrar cada registro completo.
    const filas = limpios.map(function (a, i) {
      return [
        i === 0 ? empresa.nit : '',
        i === 0 ? empresa.empresa : '',
        i === 0 ? administrador : '',
        i === 0 ? cantidad : '',
        a.nombre,
        a.telefono,
        i === 0 ? fecha : '',
        id
      ];
    });
 
    const inicio = Math.max(sh.getLastRow(), 2) + 1;
    const rango = sh.getRange(inicio, 1, filas.length, COLS_REGISTROS.length);
    sh.getRange(inicio, 1, filas.length, 1).setNumberFormat('@');
    sh.getRange(inicio, 6, filas.length, 1).setNumberFormat('@');
    rango.setValues(filas);
    rango.setBorder(true, true, true, true, true, true);
    sh.getRange(inicio, 7, 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
 
    return { ok: true, filas: filas.length, id: id };
  } finally {
    lock.releaseLock();
  }
}
 
