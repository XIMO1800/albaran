/* =====================================================================
   CUADRE DE AUTOVENTA — archivo aparte dentro del mismo proyecto de
   Apps Script que Codigo.gs (hoja PEDIDOS TIENDA). Solo lo usa ADMIN.

   INSTALACIÓN (una vez):
   1. En el editor de Apps Script: + (Archivos) → Script → nombre «Autoventa»
      y pega este archivo entero.
   2. En Codigo.gs, dentro de doPost, cambia esta línea:
          var f = ACCIONES[req.accion];
      por esta:
          var f = ACCIONES[req.accion] || ACCIONES_AV[req.accion];
   3. Implementar → Gestionar implementaciones → editar (lápiz) →
      Versión: «Nueva versión» → Implementar. La URL no cambia.

   La pestaña AUTOVENTA se crea sola la primera vez. Cada fila es un
   cuadre (columna DATOS en JSON); la fila con ID «CONFIG» guarda las
   referencias, vendedores y taras.
   ===================================================================== */

var HOJA_AV = 'AUTOVENTA';
var CAB_AV = ['ID', 'VENDEDOR', 'FECHA', 'ESTADO', 'DATOS', 'ACTUALIZADO', 'POR'];

function hojaAv_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA_AV);
  if (!sh) {
    sh = ss.insertSheet(HOJA_AV);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('C:C').setNumberFormat('@');
    sh.getRange(1, 1, 1, CAB_AV.length).setValues([CAB_AV]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function filaAv_(sh, id) {
  var n = sh.getLastRow();
  if (n < 2) return 0;
  var ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return 0;
}

function idAv_(id) {
  id = String(id || '');
  if (!/^[A-Za-z0-9_-]{1,40}$/.test(id)) throw new Error('Cuadre sin identificador válido');
  return id;
}

var ACCIONES_AV = {

  // Todos los cuadres y la configuración
  avDatos: function (req, u) {
    permitir_(u, ['ADMIN']);
    var v = hojaAv_().getDataRange().getValues();
    v.shift();
    var cuadres = [], config = null;
    v.forEach(function (r) {
      if (r[0] === '' || r[0] == null) return;
      var d;
      try { d = JSON.parse(r[4]); } catch (e) { return; }
      if (String(r[0]) === 'CONFIG') config = d;
      else cuadres.push({id: String(r[0]), datos: d});
    });
    return {cuadres: cuadres, config: config};
  },

  // Crea o reemplaza un cuadre (o la configuración con id CONFIG)
  avGuardar: function (req, u) {
    permitir_(u, ['ADMIN']);
    var id = idAv_(req.id), d = req.datos || {}, txt = JSON.stringify(d);
    if (txt.length > 45000) throw new Error('El cuadre es demasiado grande para guardarlo');
    var sh = hojaAv_(), fila = filaAv_(sh, id);
    var row = [id, d.vendedor || '', d.fecha || '', d.estado || '', txt, new Date(), u.nombre];
    if (fila) sh.getRange(fila, 1, 1, row.length).setValues([row]);
    else sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    return {};
  },

  avBorrar: function (req, u) {
    permitir_(u, ['ADMIN']);
    var id = idAv_(req.id), sh = hojaAv_(), fila = filaAv_(sh, id);
    if (fila) sh.deleteRow(fila);
    return {};
  }
};
