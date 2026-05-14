function obtenerMisVentas(documentoUsuario) {
  try {
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var venData = ss.getSheetByName('VENTA').getDataRange().getValues();
    var facData = ss.getSheetByName('FACTURA').getDataRange().getValues();

    // Mapa factura → documento empleado
    var mFac = {};
    for (var i = 1; i < facData.length; i++) {
      if (!facData[i][0]) continue;
      var numFac = String(facData[i][0]).trim();
      var docEmp = String(parseInt(facData[i][4]) || facData[i][4]).trim();
      mFac[numFac] = docEmp;
    }

    var docUsuario = String(documentoUsuario).trim();
    var ventas = [];

    for (var i = 1; i < venData.length; i++) {
      if (!venData[i][0]) continue;
      var numFac = String(venData[i][1]).trim();

      // Solo incluir si la factura pertenece al vendedor loggeado
      if (mFac[numFac] !== docUsuario) continue;

      ventas.push({
        idVenta:       String(venData[i][0]).trim(),
        numeroFactura: numFac,
        imei:          String(venData[i][2]).trim(),
        valorVenta:    parseFloat(venData[i][3]) || 0
      });
    }

    ventas.reverse();
    return { ok: true, ventas: ventas };
  } catch(e) {
    return { ok: false, msg: e.message, ventas: [] };
  }
}
function obtenerTodasLasVentas() {
  try {
    var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    var data = ss.getSheetByName('VENTA').getDataRange().getValues();
    var ventas = [];

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      ventas.push({
        idVenta:       String(data[i][0]).trim(),
        numeroFactura: String(data[i][1]).trim(),
        imei:          String(data[i][2]).trim(),
        valorVenta:    parseFloat(data[i][3]) || 0
      });
    }

    ventas.reverse();
    return { ok: true, ventas: ventas };
  } catch(e) {
    return { ok: false, msg: e.message, ventas: [] };
  }
}

function obtenerMisCompras(documentoUsuario) {
  try {
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var venData = ss.getSheetByName('VENTA').getDataRange().getValues();
    var facData = ss.getSheetByName('FACTURA').getDataRange().getValues();

    // Mapa factura → documento cliente (columna D = índice 3)
    var mFac = {};
    for (var i = 1; i < facData.length; i++) {
      if (!facData[i][0]) continue;
      var numFac  = String(facData[i][0]).trim();
      var docCli  = String(parseInt(facData[i][3]) || facData[i][3]).trim();
      mFac[numFac] = docCli;
    }

    var docUsuario = String(documentoUsuario).trim();
    var compras = [];

    for (var i = 1; i < venData.length; i++) {
      if (!venData[i][0]) continue;
      var numFac = String(venData[i][1]).trim();
      if (mFac[numFac] !== docUsuario) continue;
      compras.push({
        idVenta:       String(venData[i][0]).trim(),
        numeroFactura: numFac,
        imei:          String(venData[i][2]).trim(),
        valorVenta:    parseFloat(venData[i][3]) || 0
      });
    }

    compras.reverse();
    return { ok: true, ventas: compras };
  } catch(e) {
    return { ok: false, msg: e.message, ventas: [] };
  }
}