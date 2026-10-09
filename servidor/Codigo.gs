/* =====================================================================
   PEDIDOS TIENDA · Quesos El Hidalgo
   Servidor (Apps Script) de la hoja "PEDIDOS TIENDA".
   App: ximo1800.github.io/albaran/

   Circuito:  TIENDA hace el pedido (piezas)
           →  FABRICA lo prepara y apunta piezas servidas y kg
           →  OFICINA/ADMIN emite el albarán (nº automático) e imprime
   Estados del pedido: ENVIADO → PREPARANDO → PREPARADO → ALBARANADO  (o ANULADO)

   Pestañas (se crean solas la primera vez):
     PRODUCTOS       catálogo y precios €/kg (lo que ve la tienda)
     PEDIDOS         un pedido por fila; las líneas van en LINEAS_JSON
     ALBARANES       cabecera de cada albarán
     ALBARAN_LINEAS  una fila por línea de albarán (para facturación / importar)
     USUARIOS        NOMBRE · PIN · ROL (TIENDA, FABRICA, OFICINA, ADMIN)
     LOG             quién cambió qué y cuándo
   ===================================================================== */

var HOJAS = {
  PRODUCTOS:      ['CODIGO','NOMBRE','GRUPO','PRECIO','PZ_BANDEJA','ACTIVO','ORDEN','PESO_MEDIO'],
  PEDIDOS:        ['ID','CREADO','CREADO_POR','ESTADO','OBS','LINEAS_JSON','VERSION','ACTUALIZADO','ACTUALIZADO_POR','ALBARAN','OBS_FABRICA'],
  ALBARANES:      ['NUM','FECHA','CLIENTE','PEDIDOS','KG','TOTAL','OBS','ESTADO','CREADO','CREADO_POR','EXPORTADO','EXPORTADO_POR'],
  ALBARAN_LINEAS: ['NUM','FECHA','CLIENTE','PEDIDO','CODIGO','NOMBRE','PIEZAS','KG','PRECIO','IMPORTE'],
  USUARIOS:       ['NOMBRE','PIN','ROL'],
  LOG:            ['FECHA','USUARIO','ACCION','REF','DETALLE']
};

// Grupos: M mezcla, O oveja, D D.O., C cabra, V avería, X otros
var PRODUCTOS_INICIALES = [
  ['MTG','Mezcla tierno 3 kg','M',8.65,6],
  ['MSG','Mezcla semi 3 kg','M',9.25,6],
  ['MCG','Mezcla curado 3 kg','M',11.05,6],
  ['MCG TOSTADO','Mezcla tostado 3 kg','M',13.20,6],
  ['MAG','Mezcla aceite 3 kg','M',13.20,6],
  ['OTG','Oveja tierno 3 kg','O',13.40,6],
  ['OTM','Oveja tierno 2 kg','O',13.40,1],
  ['OTN','Oveja tierno 1 kg','O',13.40,1],
  ['OSG','Oveja semi 3 kg','O',14.20,6],
  ['OSN','Oveja semi 1 kg','O',14.20,1],
  ['OCG','Oveja curado 3 kg','O',15.10,6],
  ['OCN','Oveja curado 1 kg','O',15.10,1],
  ['OCP BRANDY','Oveja brandy','O',16.35,1],
  ['DCGA','D.O. curado 3 kg','D',18.10,6],
  ['DCP ARTESANO','D.O. artesano','D',18.75,1],
  ['CTN','Cabra tierno 1 kg','C',11.80,1],
  ['CSG','Cabra semi','C',12.50,1],
  ['OCG AVERÍA','Avería oveja 3 kg','V',12.55,1],
  ['OCM AVERÍA','Avería oveja 2 kg','V',12.55,1],
  ['OCP AVERÍA','Avería oveja pequeño','V',12.55,1],
  ['SIN SAL','Sin sal','X',9.40,1]
];

// Cambia los PIN en la pestaña USUARIOS cuando quieras (no hace falta tocar el código)
var USUARIOS_INICIALES = [
  ['JOAQUIN','2205','ADMIN'],
  ['TIENDA','1111','TIENDA'],
  ['FABRICA','2222','FABRICA'],
  ['OFICINA','3333','OFICINA']
];

var DIAS_HISTORIAL = 60;   // pedidos/albaranes cerrados que se mandan a la app

/* ------------------------------------------------------------------ */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.tipo) {
    try { return puente_(p); }
    catch (err) { return ContentService.createTextOutput('ERROR: ' + (err.message || err)).setMimeType(ContentService.MimeType.TEXT); }
  }
  return json_({ok: true, app: 'PEDIDOS TIENDA', hora: new Date().toISOString()});
}

/* ------------- puente al servidor (ordenador del despacho, Windows 7) -------------
   El VBS del ordenador de Joaquín pide los albaranes T*.TXT que aún no ha dejado en ALBTIEND:
     ?tipo=pendientes&clave=…              → «OK» y debajo un nombre por línea
     ?tipo=fichero&nombre=T0000123.TXT&clave=…  → el fichero tal cual (mismo formato de siempre)
     ?tipo=entregado&nombre=T0000123.TXT&clave=… → lo marca entregado (descripción del fichero en Drive)
   Respuestas en texto plano para que el VBS no tenga que leer JSON. La clave NO está en el código:
   se crea una vez ejecutando INSTALAR_PUENTE() desde el editor y queda en las propiedades del proyecto. */
var PUENTE_DIAS = 30;   // solo mira ficheros de los últimos 30 días
function puente_(p) {
  var txt = function (s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT); };
  var props = PropertiesService.getScriptProperties(), clave = props.getProperty('CLAVE_PUENTE');
  if (!clave || p.clave !== clave) return txt('ERROR: clave incorrecta');
  var carpeta = DriveApp.getFolderById(CARPETA_PROGRAMA_ID);
  var buscar = function (nombre) {
    if (!/^T\d{7}\.TXT$/.test(nombre || '')) return null;
    var it = carpeta.getFilesByName(nombre);
    while (it.hasNext()) { var f = it.next(); if (!f.isTrashed()) return f; }
    return null;
  };
  if (p.tipo === 'pendientes') {
    var desde = Math.max(Number(props.getProperty('PUENTE_DESDE') || 0), Date.now() - PUENTE_DIAS * 864e5);
    var nombres = [];
    // Recorre la carpeta sin búsquedas de Drive (cambian de sintaxis según la versión) y se queda con los T*.TXT recientes
    var it = carpeta.getFiles();
    while (it.hasNext()) {
      var f = it.next(), n = f.getName();
      if (!/^T\d{7}\.TXT$/.test(n) || f.isTrashed()) continue;
      if (f.getDateCreated().getTime() <= desde) continue;
      if (String(f.getDescription() || '').indexOf('ENTREGADO') !== 0) nombres.push(n);
    }
    nombres.sort();
    return txt(['OK'].concat(nombres).join('\r\n'));
  }
  if (p.tipo === 'fichero') {
    var f1 = buscar(p.nombre);
    if (!f1) return txt('ERROR: no existe ' + p.nombre);
    return txt(f1.getBlob().getDataAsString());
  }
  if (p.tipo === 'entregado') {
    var f2 = buscar(p.nombre);
    if (!f2) return txt('ERROR: no existe ' + p.nombre);
    f2.setDescription('ENTREGADO ' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm') + ' · ALBTIEND');
    return txt('OK');
  }
  return txt('ERROR: tipo desconocido');
}
// Ejecutar UNA vez desde el editor (botón ▶ con esta función elegida). Crea la clave y la enseña en el registro.
// Solo da como pendientes los albaranes emitidos A PARTIR de este momento (los de antes ya los llevó el puente viejo).
function INSTALAR_PUENTE() {
  var props = PropertiesService.getScriptProperties();
  var clave = props.getProperty('CLAVE_PUENTE');
  if (!clave) { clave = Utilities.getUuid().replace(/-/g, ''); props.setProperty('CLAVE_PUENTE', clave); }
  if (!props.getProperty('PUENTE_DESDE')) props.setProperty('PUENTE_DESDE', String(Date.now()));
  Logger.log('CLAVE para el VBS: ' + clave);
  Logger.log('Pendientes desde: ' + new Date(Number(props.getProperty('PUENTE_DESDE'))));
}
// Si alguna vez hiciera falta volver a mandar un albarán: quitarle la marca y el puente lo vuelve a dejar.
function REENVIAR_ALBARAN(num) {
  var it = DriveApp.getFolderById(CARPETA_PROGRAMA_ID).getFilesByName(nombreAlbaranProg_(num));
  while (it.hasNext()) it.next().setDescription('');
}

function doPost(e) {
  if (e && e.parameter && e.parameter.tipo) {   // puente del despacho (subida de ficheros de la PDA)
    try { return puenteSubir_(e); }
    catch (err) { return ContentService.createTextOutput('ERROR: ' + (err.message || err)).setMimeType(ContentService.MimeType.TEXT); }
  }
  var req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ok: false, error: 'Petición no válida'}); }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (err) { return json_({ok: false, error: 'Servidor ocupado, inténtalo otra vez'}); }

  try {
    asegurarHojas_();
    var u = usuario_(req.pin);
    if (!u) return json_({ok: false, error: 'PIN incorrecto', pin: true});
    var f = ACCIONES[req.accion];
    if (!f) return json_({ok: false, error: 'Acción desconocida: ' + req.accion});
    var r = f(req, u) || {};
    r.ok = true;
    return json_(r);
  } catch (err) {
    return json_({ok: false, error: String(err.message || err)});
  } finally {
    lock.releaseLock();
  }
}

var ACCIONES = {

  login: function (req, u) { return {usuario: u.nombre, rol: u.rol}; },

  datos: function (req, u) {
    var limite = Date.now() - DIAS_HISTORIAL * 864e5;
    var peds = leer_('PEDIDOS').filter(function (p) {
      return (p.ESTADO !== 'ALBARANADO' && p.ESTADO !== 'ANULADO') || fecha_(p.ACTUALIZADO) >= limite;
    }).map(pedidoOut_);
    var resp = {usuario: u.nombre, rol: u.rol, productos: productos_(), pedidos: peds, hora: new Date().toISOString()};
    if (u.rol !== 'FABRICA') {
      var lineas = leer_('ALBARAN_LINEAS');
      resp.albaranes = leer_('ALBARANES').filter(function (a) { return fecha_(a.CREADO) >= limite; })
        .map(function (a) { return albaranOut_(a, lineas); });
    }
    return resp;
  },

  // Crear o modificar un pedido (tienda, oficina, admin)
  guardarPedido: function (req, u) {
    permitir_(u, ['TIENDA','OFICINA','ADMIN']);
    var p = req.pedido || {};
    var lineas = limpiarLineas_(p.lineas);
    if (!lineas.length) throw new Error('El pedido no tiene productos');
    var sh = hoja_('PEDIDOS'), ahora = new Date();

    if (!p.id) {
      var id = siguienteId_();
      sh.appendRow([id, ahora, u.nombre, 'ENVIADO', p.obs || '', JSON.stringify(lineas), 1, ahora, u.nombre, '', '']);
      log_(u, 'NUEVO PEDIDO', id, resumen_(lineas) + (p.obs ? ' · OBS: ' + p.obs : ''));
      return {pedido: pedidoOut_(buscar_('PEDIDOS', 'ID', id).obj)};
    }
    var f = buscar_('PEDIDOS', 'ID', p.id);
    if (!f) throw new Error('Pedido ' + p.id + ' no encontrado');
    comprobarVersion_(f.obj, p.version);
    if (u.rol === 'TIENDA' && f.obj.ESTADO !== 'ENVIADO')
      throw new Error('Este pedido ya se está preparando. Para cambiarlo, llama a la fábrica.');
    if (f.obj.ESTADO === 'ALBARANADO' || f.obj.ESTADO === 'ANULADO')
      throw new Error('El pedido está ' + f.obj.ESTADO.toLowerCase() + ' y no se puede cambiar');
    escribir_(sh, f.fila, {OBS: p.obs || '', LINEAS_JSON: JSON.stringify(lineas),
      VERSION: (+f.obj.VERSION || 1) + 1, ACTUALIZADO: ahora, ACTUALIZADO_POR: u.nombre});
    log_(u, 'MODIFICA PEDIDO', p.id, resumen_(lineas) + (p.obs ? ' · OBS: ' + p.obs : ''));
    return {pedido: pedidoOut_(buscar_('PEDIDOS', 'ID', p.id).obj)};
  },

  // Preparación en fábrica: piezas servidas, kg, observaciones, estado
  prepararPedido: function (req, u) {
    permitir_(u, ['FABRICA','OFICINA','ADMIN']);
    var p = req.pedido || {};
    var f = buscar_('PEDIDOS', 'ID', p.id);
    if (!f) throw new Error('Pedido no encontrado');
    comprobarVersion_(f.obj, p.version);
    if (f.obj.ESTADO === 'ALBARANADO' || f.obj.ESTADO === 'ANULADO')
      throw new Error('El pedido está ' + f.obj.ESTADO.toLowerCase());
    var estado = p.estado === 'PREPARADO' ? 'PREPARADO' : 'PREPARANDO';
    var lineas = limpiarLineas_(p.lineas);
    if (estado === 'PREPARADO') {
      var faltan = lineas.filter(function (l) { return l.pzS > 0 && !(l.kg > 0); });
      if (faltan.length) throw new Error('Falta el peso de: ' + faltan.map(function (l) { return l.c; }).join(', '));
    }
    escribir_(hoja_('PEDIDOS'), f.fila, {ESTADO: estado, OBS_FABRICA: p.obsF || '', LINEAS_JSON: JSON.stringify(lineas),
      VERSION: (+f.obj.VERSION || 1) + 1, ACTUALIZADO: new Date(), ACTUALIZADO_POR: u.nombre});
    log_(u, estado === 'PREPARADO' ? 'PEDIDO PREPARADO' : 'PREPARANDO', p.id, resumen_(lineas, true));
    return {pedido: pedidoOut_(buscar_('PEDIDOS', 'ID', p.id).obj)};
  },

  // Volver un pedido preparado a "preparando" (corregir pesos)
  reabrirPedido: function (req, u) {
    permitir_(u, ['FABRICA','OFICINA','ADMIN']);
    var f = buscar_('PEDIDOS', 'ID', req.id);
    if (!f) throw new Error('Pedido no encontrado');
    if (f.obj.ESTADO !== 'PREPARADO') throw new Error('Solo se puede reabrir un pedido preparado');
    escribir_(hoja_('PEDIDOS'), f.fila, {ESTADO: 'PREPARANDO', VERSION: (+f.obj.VERSION || 1) + 1,
      ACTUALIZADO: new Date(), ACTUALIZADO_POR: u.nombre});
    log_(u, 'REABRE PEDIDO', req.id, '');
    return {pedido: pedidoOut_(buscar_('PEDIDOS', 'ID', req.id).obj)};
  },

  // Historia de lo pedido: el original y cada modificación (sale del LOG)
  historialPedido: function (req, u) {
    var id = String(req.id || ''), out = [];
    var v = hoja_('LOG').getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      var r = v[i], acc = String(r[2]);
      if (String(r[3]) !== id || (acc !== 'NUEVO PEDIDO' && acc !== 'MODIFICA PEDIDO')) continue;
      var det = String(r[4] || ''), obs = null, k = det.indexOf(' · OBS: ');
      if (k >= 0) { obs = det.slice(k + 8); det = det.slice(0, k); }
      var ls = det.split(', ').map(function (t) {
        var m = t.match(/^(-?[\d.]+)\s+(.+)$/); return m ? {c: m[2].trim(), pz: num_(m[1])} : null;
      }).filter(function (x) { return x; });
      out.push({fecha: iso_(r[0]), usuario: String(r[1]), accion: acc, lineas: ls, obs: obs});
    }
    return {historial: out};
  },

  anularPedido: function (req, u) {
    permitir_(u, ['TIENDA','OFICINA','ADMIN']);
    var f = buscar_('PEDIDOS', 'ID', req.id);
    if (!f) throw new Error('Pedido no encontrado');
    if (u.rol === 'TIENDA' && f.obj.ESTADO !== 'ENVIADO') throw new Error('El pedido ya se está preparando. Llama a la fábrica.');
    if (f.obj.ESTADO === 'ALBARANADO') throw new Error('El pedido ya tiene albarán');
    escribir_(hoja_('PEDIDOS'), f.fila, {ESTADO: 'ANULADO', VERSION: (+f.obj.VERSION || 1) + 1,
      ACTUALIZADO: new Date(), ACTUALIZADO_POR: u.nombre});
    log_(u, 'ANULA PEDIDO', req.id, req.motivo || '');
    return {};
  },

  // Emitir albarán a partir de uno o varios pedidos preparados
  emitirAlbaran: function (req, u) {
    permitir_(u, ['OFICINA','ADMIN']);
    var a = req.albaran || {};
    var ids = (a.pedidos || []).map(String);
    if (!ids.length) throw new Error('Elige al menos un pedido');
    var filas = ids.map(function (id) {
      var f = buscar_('PEDIDOS', 'ID', id);
      if (!f) throw new Error('Pedido ' + id + ' no encontrado');
      if (f.obj.ESTADO !== 'PREPARADO') throw new Error('El pedido ' + id + ' no está preparado');
      return f;
    });
    var lineas = (a.lineas || []).map(function (l) {
      return {pedido: String(l.pedido || ''), c: String(l.c || '').trim(), n: String(l.n || ''),
              pz: num_(l.pz), kg: num_(l.kg), p: num_(l.p)};
    }).filter(function (l) { return l.c && (l.kg > 0 || l.pz > 0); });
    if (!lineas.length) throw new Error('El albarán no tiene líneas');

    var shA = hoja_('ALBARANES'), num = siguienteAlbaran_();
    var fecha = a.fecha ? new Date(a.fecha + 'T12:00:00') : new Date();
    var cliente = a.cliente || 'CASA', kg = 0, total = 0, filasL = [];
    lineas.forEach(function (l) {
      var imp = r2_(l.kg * l.p); kg += l.kg; total += imp;
      filasL.push([num, fecha, cliente, l.pedido, l.c, l.n, l.pz, l.kg, l.p, imp]);
    });
    shA.appendRow([num, fecha, cliente, ids.join(','), r3_(kg), r2_(total), a.obs || '', 'EMITIDO', new Date(), u.nombre]);
    var shL = hoja_('ALBARAN_LINEAS');
    shL.getRange(shL.getLastRow() + 1, 1, filasL.length, filasL[0].length).setValues(filasL);
    filas.forEach(function (f) {
      escribir_(hoja_('PEDIDOS'), f.fila, {ESTADO: 'ALBARANADO', ALBARAN: num, VERSION: (+f.obj.VERSION || 1) + 1,
        ACTUALIZADO: new Date(), ACTUALIZADO_POR: u.nombre});
    });
    log_(u, 'EMITE ALBARÁN', num, 'Pedidos ' + ids.join(',') + ' · ' + r2_(total) + ' €');
    var fich = '';
    try {   // fichero para el programa de gestión (si falla, el albarán queda emitido igual)
      fich = exportarAlbaranProg_(num, lineas, fecha);
      var fa = buscar_('ALBARANES', 'NUM', num);
      if (fa) escribir_(shA, fa.fila, {EXPORTADO: new Date(), EXPORTADO_POR: 'AUTO ' + fich});
    } catch (e) { log_(u, 'ERROR EXPORTAR', num, String(e && e.message || e)); }
    return {num: num, total: r2_(total), kg: r3_(kg), fichero: fich};
  },

  // Anular un albarán: sus pedidos vuelven a PREPARADO (solo ADMIN)
  anularAlbaran: function (req, u) {
    permitir_(u, ['ADMIN']);
    var f = buscar_('ALBARANES', 'NUM', req.num);
    if (!f) throw new Error('Albarán no encontrado');
    if (f.obj.ESTADO === 'ANULADO') throw new Error('Ya está anulado');
    escribir_(hoja_('ALBARANES'), f.fila, {ESTADO: 'ANULADO'});
    String(f.obj.PEDIDOS).split(',').forEach(function (id) {
      var p = buscar_('PEDIDOS', 'ID', id.trim());
      if (p) escribir_(hoja_('PEDIDOS'), p.fila, {ESTADO: 'PREPARADO', ALBARAN: '', VERSION: (+p.obj.VERSION || 1) + 1,
        ACTUALIZADO: new Date(), ACTUALIZADO_POR: u.nombre});
    });
    // las líneas se quedan pero marcadas en el NUM con la palabra ANULADO para que no se importen
    var shL = hoja_('ALBARAN_LINEAS'), datos = shL.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) if (String(datos[i][0]) === String(req.num)) shL.getRange(i + 1, 1).setValue(req.num + ' ANULADO');
    try { borrarAlbaranProg_(req.num); } catch (e) {}
    log_(u, 'ANULA ALBARÁN', req.num, req.motivo || '');
    return {};
  },

  // Histórico para la versión de escritorio: pedidos y albaranes entre dos fechas (yyyy-mm-dd)
  historial: function (req, u) {
    permitir_(u, ['OFICINA','ADMIN']);
    var d0 = req.desde ? new Date(req.desde + 'T00:00:00').getTime() : 0;
    var d1 = req.hasta ? new Date(req.hasta + 'T23:59:59').getTime() : Date.now();
    var peds = leer_('PEDIDOS').filter(function (p) { var t = fecha_(p.CREADO); return t >= d0 && t <= d1; }).map(pedidoOut_);
    var lineas = leer_('ALBARAN_LINEAS');
    var albs = leer_('ALBARANES').filter(function (a) { var t = fecha_(a.FECHA); return t >= d0 && t <= d1 + 1; })
      .map(function (a) { return albaranOut_(a, lineas); });
    return {pedidos: peds, albaranes: albs};
  },

  // (Re)envía al programa de gestión el fichero de los albaranes indicados
  enviarProg: function (req, u) {
    permitir_(u, ['OFICINA','ADMIN']);
    var sh = hoja_('ALBARANES'), lineas = leer_('ALBARAN_LINEAS'), ahora = new Date(), hechos = [];
    (req.nums || []).forEach(function (num) {
      var f = buscar_('ALBARANES', 'NUM', num);
      if (!f || f.obj.ESTADO === 'ANULADO') return;
      var ls = lineas.filter(function (l) { return String(l.NUM) === String(num); })
        .map(function (l) { return {c: String(l.CODIGO), pz: num_(l.PIEZAS), kg: num_(l.KG), p: num_(l.PRECIO)}; });
      if (!ls.length) return;
      var fich = exportarAlbaranProg_(num, ls, f.obj.FECHA);
      escribir_(sh, f.fila, {EXPORTADO: ahora, EXPORTADO_POR: 'AUTO ' + fich});
      hechos.push({num: String(num), fichero: fich});
    });
    log_(u, 'ENVÍA AL PROGRAMA', '', hechos.map(function (h) { return h.fichero; }).join(','));
    return {hechos: hechos, fecha: ahora.toISOString()};
  },

  // Marca albaranes como exportados al programa de gestión
  marcarExportado: function (req, u) {
    permitir_(u, ['OFICINA','ADMIN']);
    var sh = hoja_('ALBARANES'), ahora = new Date(), n = 0;
    (req.nums || []).forEach(function (num) {
      var f = buscar_('ALBARANES', 'NUM', num);
      if (f && f.obj.ESTADO !== 'ANULADO') { escribir_(sh, f.fila, {EXPORTADO: ahora, EXPORTADO_POR: u.nombre}); n++; }
    });
    log_(u, 'EXPORTA CSV', '', n + ' albaranes: ' + (req.nums || []).join(','));
    return {marcados: n, fecha: ahora.toISOString()};
  },

  guardarProductos: function (req, u) {
    permitir_(u, ['OFICINA','ADMIN']);
    var ps = (req.productos || []).map(function (p, i) {
      return [String(p.c || '').trim(), String(p.n || ''), String(p.g || 'X'),
              num_(p.p), Math.max(1, Math.round(num_(p.b) || 1)), p.activo === false ? 'NO' : 'SI', i + 1, r3_(num_(p.pm))];
    }).filter(function (r) { return r[0]; });
    var vistos = {};
    ps.forEach(function (r) { if (vistos[r[0]]) throw new Error('Código repetido: ' + r[0]); vistos[r[0]] = 1; });
    var sh = hoja_('PRODUCTOS');
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, HOJAS.PRODUCTOS.length).clearContent();
    if (ps.length) sh.getRange(2, 1, ps.length, HOJAS.PRODUCTOS.length).setValues(ps);
    log_(u, 'CAMBIA PRECIOS', '', ps.length + ' productos');
    return {productos: productos_()};
  }
};

/* ------------------- fichero para el programa de gestión ------------------- */
// Carpeta de Drive «EXPORTACION PROGRAMA» (la misma de la recogida de leche). En el ordenador de
// administración, Google Drive la baja y la tarea «Copia ALBTIEND» copia los T*.TXT a V:\SERVIDORW10\ALBTIEND.
var CARPETA_PROGRAMA_ID = '1qoFDcSbndVMs9u4lIU0_OofFxRvTwJts';
function nombreAlbaranProg_(num) { return 'T' + ('0000000' + String(num).replace(/\D/g, '')).slice(-7) + '.TXT'; }
// Una línea por artículo: FECHA;CODIGO;PIEZAS;KG;PRECIO  (fecha del albarán dd/mm/aaaa)  (separador ; · decimales con punto · sin cabecera · CRLF)
function exportarAlbaranProg_(num, lineas, fecha) {
  var fx = Utilities.formatDate(fecha instanceof Date ? fecha : new Date(fecha || Date.now()), 'Europe/Madrid', 'dd/MM/yyyy');
  var txt = lineas.map(function (l) {
    return [fx, l.c, Math.round(l.pz), r3_(l.kg).toFixed(3), r2_(l.p).toFixed(2)].join(';');
  }).join('\r\n') + '\r\n';
  var nombre = nombreAlbaranProg_(num), carpeta = DriveApp.getFolderById(CARPETA_PROGRAMA_ID);
  var viejos = carpeta.getFilesByName(nombre); while (viejos.hasNext()) viejos.next().setTrashed(true);
  carpeta.createFile(nombre, txt, MimeType.PLAIN_TEXT);
  return nombre;
}
function borrarAlbaranProg_(num) {
  var it = DriveApp.getFolderById(CARPETA_PROGRAMA_ID).getFilesByName(nombreAlbaranProg_(num));
  while (it.hasNext()) it.next().setTrashed(true);
}
// Desde el editor: manda al programa los albaranes emitidos que aún no tienen fichero (columna EXPORTADO vacía)
function EXPORTAR_PENDIENTES() {
  var sh = hoja_('ALBARANES'), lineas = leer_('ALBARAN_LINEAS'), n = 0;
  leer_('ALBARANES').forEach(function (a) {
    if (a.ESTADO === 'ANULADO' || a.EXPORTADO) return;
    var ls = lineas.filter(function (l) { return String(l.NUM) === String(a.NUM); })
      .map(function (l) { return {c: String(l.CODIGO), pz: num_(l.PIEZAS), kg: num_(l.KG), p: num_(l.PRECIO)}; });
    if (!ls.length) return;
    var fich = exportarAlbaranProg_(a.NUM, ls, a.FECHA), f = buscar_('ALBARANES', 'NUM', a.NUM);
    escribir_(sh, f.fila, {EXPORTADO: new Date(), EXPORTADO_POR: 'AUTO ' + fich}); n++;
  });
  Logger.log(n + ' albaranes enviados al programa');
}
// Para probar desde el editor (pide permiso de Drive la primera vez)
function PROBAR_EXPORTAR() { Logger.log(exportarAlbaranProg_('9999999', [{c: 'MAN', pz: 3, kg: 1.05, p: 13.2}], new Date())); borrarAlbaranProg_('9999999'); }

/* ---------------------------- utilidades ---------------------------- */

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hoja_(n) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); }
function num_(v) { var x = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isFinite(x) ? x : 0; }
function r2_(x) { return Math.round((x + 1e-9) * 100) / 100; }
function r3_(x) { return Math.round((x + 1e-9) * 1000) / 1000; }
function fecha_(v) { return v instanceof Date ? v.getTime() : (v ? new Date(v).getTime() : 0); }
function fechaTxt_(v) { return v instanceof Date ? Utilities.formatDate(v, 'Europe/Madrid', 'yyyy-MM-dd') : String(v || ''); }
function iso_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }

function asegurarHojas_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (PropertiesService.getScriptProperties().getProperty('HOJAS_OK') === '4') return;
  Object.keys(HOJAS).forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (!sh) sh = ss.insertSheet(n);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HOJAS[n].length).setValues([HOJAS[n]]).setFontWeight('bold');
      sh.setFrozenRows(1);
      if (n === 'PRODUCTOS') sh.getRange(2, 1, PRODUCTOS_INICIALES.length, 7).setValues(
        PRODUCTOS_INICIALES.map(function (p, i) { return [p[0], p[1], p[2], p[3], p[4], 'SI', i + 1]; }));
      if (n === 'USUARIOS') { sh.getRange('B:B').setNumberFormat('@'); sh.getRange(2, 1, USUARIOS_INICIALES.length, 3).setValues(USUARIOS_INICIALES); }
      if (n === 'PEDIDOS' || n === 'ALBARANES' || n === 'ALBARAN_LINEAS') sh.getRange('A:A').setNumberFormat('@');
      if (n === 'ALBARAN_LINEAS') sh.getRange('D:D').setNumberFormat('@');
    } else {
      // hoja ya existente: añadir columnas nuevas que falten al final
      var cab = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
      HOJAS[n].forEach(function (c) {
        if (cab.indexOf(c) < 0) { cab.push(c); sh.getRange(1, cab.length).setValue(c).setFontWeight('bold'); }
      });
    }
  });
  var h1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (h1 && h1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(h1);
  PropertiesService.getScriptProperties().setProperty('HOJAS_OK', '4');
}

// Ejecuta esta función una vez desde el editor para crear las pestañas
function PREPARAR_HOJA() {
  PropertiesService.getScriptProperties().deleteProperty('HOJAS_OK');
  asegurarHojas_();
}

function albaranOut_(a, lineas) {
  return {num: String(a.NUM), fecha: fechaTxt_(a.FECHA), cliente: a.CLIENTE, pedidos: String(a.PEDIDOS),
          kg: +a.KG, total: +a.TOTAL, obs: a.OBS, estado: a.ESTADO, creadoPor: a.CREADO_POR, creado: iso_(a.CREADO),
          exportado: a.EXPORTADO ? iso_(a.EXPORTADO) : '', exportadoPor: a.EXPORTADO_POR || '',
          lineas: lineas.filter(function (l) { return String(l.NUM) === String(a.NUM); })
            .map(function (l) { return {pedido: String(l.PEDIDO), c: l.CODIGO, n: l.NOMBRE, pz: +l.PIEZAS, kg: +l.KG, p: +l.PRECIO}; })};
}

function leer_(n) {
  var v = hoja_(n).getDataRange().getValues(), cab = v.shift();
  return v.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    var o = {}; cab.forEach(function (c, i) { o[c] = r[i]; }); return o;
  });
}

function buscar_(n, col, valor) {
  var v = hoja_(n).getDataRange().getValues(), cab = v[0], i = cab.indexOf(col);
  for (var k = 1; k < v.length; k++) if (String(v[k][i]) === String(valor)) {
    var o = {}; cab.forEach(function (c, j) { o[c] = v[k][j]; });
    return {fila: k + 1, obj: o};
  }
  return null;
}

function escribir_(sh, fila, cambios) {
  var cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var r = sh.getRange(fila, 1, 1, cab.length), v = r.getValues()[0];
  Object.keys(cambios).forEach(function (k) { var i = cab.indexOf(k); if (i >= 0) v[i] = cambios[k]; });
  r.setValues([v]);
}

function usuario_(pin) {
  if (!pin) return null;
  var us = leer_('USUARIOS');
  for (var i = 0; i < us.length; i++) if (String(us[i].PIN).trim() === String(pin).trim())
    return {nombre: String(us[i].NOMBRE).toUpperCase(), rol: String(us[i].ROL).toUpperCase()};
  return null;
}

function permitir_(u, roles) {
  if (roles.indexOf(u.rol) < 0) throw new Error('Tu usuario (' + u.rol + ') no puede hacer esto');
}

function comprobarVersion_(obj, version) {
  if (version != null && String(version) !== String(obj.VERSION))
    throw new Error('Otra persona ha cambiado este pedido mientras tanto. Se recargará con los datos nuevos.');
}

function productos_() {
  return leer_('PRODUCTOS').map(function (p) {
    return {c: String(p.CODIGO), n: String(p.NOMBRE), g: String(p.GRUPO || 'X'), p: num_(p.PRECIO),
            b: Math.max(1, num_(p.PZ_BANDEJA) || 1), pm: num_(p.PESO_MEDIO), activo: String(p.ACTIVO).toUpperCase() !== 'NO'};
  });
}

function pedidoOut_(p) {
  var l = []; try { l = JSON.parse(p.LINEAS_JSON || '[]'); } catch (e) {}
  return {id: String(p.ID), creado: iso_(p.CREADO), creadoPor: p.CREADO_POR, estado: p.ESTADO, obs: p.OBS,
          lineas: l, version: +p.VERSION, actualizado: iso_(p.ACTUALIZADO), actualizadoPor: p.ACTUALIZADO_POR,
          albaran: String(p.ALBARAN || ''), obsF: String(p.OBS_FABRICA || '')};
}

function limpiarLineas_(ls) {
  return (ls || []).map(function (l) {
    var o = {c: String(l.c || '').trim(), n: String(l.n || ''), pz: num_(l.pz),
             pzS: l.pzS == null || l.pzS === '' ? null : num_(l.pzS), kg: num_(l.kg), p: num_(l.p),
             obs: String(l.obs || ''), extra: !!l.extra, ok: !!l.ok, aprox: !!l.aprox};
    if (!o.aprox) delete o.aprox;
    if (o.pzS === null) delete o.pzS;
    return o;
  }).filter(function (l) { return l.c && (l.pz > 0 || l.pzS > 0 || l.kg > 0); });
}

function siguienteId_() {
  var v = hoja_('PEDIDOS').getRange('A:A').getValues(), max = 0;
  v.forEach(function (r) { var n = parseInt(r[0], 10); if (n > max) max = n; });
  return String(max + 1);
}

function siguienteAlbaran_() {
  var v = hoja_('ALBARANES').getRange('A:A').getValues(), max = 0;
  v.forEach(function (r) { var n = parseInt(r[0], 10); if (n > max) max = n; });
  var inicio = parseInt(PropertiesService.getScriptProperties().getProperty('PRIMER_ALBARAN') || '1', 10);
  return String(Math.max(max + 1, inicio));
}

function resumen_(ls, conKg) {
  return ls.map(function (l) {
    return (l.pzS != null ? l.pzS : l.pz) + ' ' + l.c + (conKg && l.kg ? ' (' + l.kg + ' kg)' : '');
  }).join(', ').slice(0, 1000);
}

function log_(u, accion, ref, detalle) {
  hoja_('LOG').appendRow([new Date(), u.nombre, accion, ref, detalle]);
}

/* ------------- PDA de autoventa: resúmenes «.xps» que deja el programa en V:\SERVIDORW10\AUTOVENT -------------
   El puente del despacho los sube aquí (POST ?tipo=subirPda&nombre=…&clave=…, cuerpo = fichero en base64) y se
   guardan en la carpeta de Drive «AUTOVENT PDA». La app los lista y los lee ella misma (pdaLista / pdaFichero). */
function carpetaPda_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('CARPETA_PDA');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('AUTOVENT PDA');
  props.setProperty('CARPETA_PDA', f.getId());
  return f;
}
function puenteSubir_(e) {
  var txt = function (s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT); };
  var p = e.parameter, clave = PropertiesService.getScriptProperties().getProperty('CLAVE_PUENTE');
  if (!clave || p.clave !== clave) return txt('ERROR: clave incorrecta');
  if (p.tipo !== 'subirPda') return txt('ERROR: tipo desconocido');
  var nombre = String(p.nombre || '').replace(/[\\/:*?"<>|]/g, '_');
  if (!/\.o?xps$/i.test(nombre)) return txt('ERROR: solo ficheros .xps');
  var b64 = String((e.postData && e.postData.contents) || '').replace(/\s/g, '');
  if (!b64) return txt('ERROR: fichero vacío');
  var bytes = Utilities.base64Decode(b64);
  var carpeta = carpetaPda_(), viejos = carpeta.getFilesByName(nombre);
  while (viejos.hasNext()) viejos.next().setTrashed(true);
  carpeta.createFile(Utilities.newBlob(bytes, 'application/vnd.ms-xpsdocument', nombre));
  return txt('OK');
}
ACCIONES.pdaLista = function (req, u) {
  permitir_(u, ['ADMIN']);
  var it = carpetaPda_().getFiles(), l = [];
  while (it.hasNext()) {
    var f = it.next();
    if (f.isTrashed()) continue;
    l.push({id: f.getId(), nombre: f.getName(), fecha: f.getDateCreated().toISOString(), tam: f.getSize()});
  }
  l.sort(function (a, b) { return a.fecha < b.fecha ? 1 : -1; });
  return {ficheros: l.slice(0, 15)};
};
ACCIONES.pdaFichero = function (req, u) {
  permitir_(u, ['ADMIN']);
  var f = DriveApp.getFileById(String(req.id || ''));
  if (f.getParents().next().getId() !== carpetaPda_().getId()) throw new Error('Ese fichero no es de la PDA');
  return {nombre: f.getName(), b64: Utilities.base64Encode(f.getBlob().getBytes())};
};
