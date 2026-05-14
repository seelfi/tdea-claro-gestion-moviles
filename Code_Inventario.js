/* ══════════════════════════════════════════════════════════════
   Code_Inventario.gs
══════════════════════════════════════════════════════════════ */

function obtenerInventarioCompleto() {
  try {
    var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    var invData  = ss.getSheetByName('INVENTARIO').getDataRange().getValues();
    var bodData  = ss.getSheetByName('BODEGA').getDataRange().getValues();
    var proData  = ss.getSheetByName('PRODUCTO').getDataRange().getValues();
    var pdvData  = ss.getSheetByName('PUNTO_DE_VENTA').getDataRange().getValues();

    var mProd = {};
    for (var i = 1; i < proData.length; i++) {
      mProd[String(proData[i][0]).trim()] = String(proData[i][1]).trim();
    }

    var mPDV = {};
    for (var i = 1; i < pdvData.length; i++) {
      mPDV[String(pdvData[i][0]).trim()] = String(pdvData[i][1]).trim();
    }

    var mBodega = {};
    for (var i = 1; i < bodData.length; i++) {
      var codInv = String(bodData[i][0]).trim();
      var codPDV = String(bodData[i][1]).trim();
      if (!mBodega[codInv]) mBodega[codInv] = [];
      mBodega[codInv].push({
        codPDV:    codPDV,
        pdvNombre: mPDV[codPDV] || codPDV,
        dispPDV:   parseInt(bodData[i][2]) || 0,
        vendPDV:   parseInt(bodData[i][3]) || 0
      });
    }

    var items = [];
    for (var i = 1; i < invData.length; i++) {
      if (!invData[i][0]) continue;
      var codInv  = String(invData[i][0]).trim();
      var codProd = String(invData[i][1]).trim();
      items.push({
        codInventario:   codInv,
        codProducto:     codProd,
        nombreProducto:  mProd[codProd] || codProd,
        stockDisponible: parseInt(invData[i][2]) || 0,
        stockVendido:    parseInt(invData[i][3]) || 0,
        precioUnitario:  parseFloat(invData[i][4]) || 0,
        sedes:           mBodega[codInv] || []
      });
    }

    var pdvs = Object.keys(mPDV).map(function(k) {
      return { codigo: k, nombre: mPDV[k] };
    });

    return { ok: true, items: items, pdvs: pdvs };

  } catch(e) {
    return { ok: false, msg: e.message, items: [], pdvs: [] };
  }
} // ← cierre correcto de obtenerInventarioCompleto


/* ──────────────────────────────────────────────────────────────
   REGISTRAR EQUIPOS — ahora es función independiente y pública
────────────────────────────────────────────────────────────── */
function registrarEquipos(datos) {
  try {
    var ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    var hojaDetalle = ss.getSheetByName('DETALLE_PRODUCTO');
    var hojaBodega  = ss.getSheetByName('BODEGA');
    var hojaInv     = ss.getSheetByName('INVENTARIO');

    var codProd = String(datos.codProducto).trim();
    var codPDV  = String(datos.codPDV).trim();
    var equipos = datos.equipos;

    if (!codProd || !codPDV || !equipos || !equipos.length) {
      return { ok: false, msg: 'Datos incompletos.' };
    }

    // Verificar duplicados de IMEI en DETALLE_PRODUCTO
    var detalleData     = hojaDetalle.getDataRange().getValues();
    var imeisExistentes = {};
    for (var i = 1; i < detalleData.length; i++) {
      imeisExistentes[String(detalleData[i][0]).trim()] = true;
    }
    for (var i = 0; i < equipos.length; i++) {
      if (imeisExistentes[equipos[i].imei]) {
        return { ok: false, msg: 'IMEI duplicado: ' + equipos[i].imei + '. Regenera los equipos.' };
      }
    }

    // Insertar equipos en DETALLE_PRODUCTO
    // [IMEI, Codigo_Producto, Serial, Color, Tamano, Peso, Garantia, Codigo_PDV]
    var filas = equipos.map(function(eq) {
      return [eq.imei, codProd, eq.serial, eq.color,
              eq.tamano, eq.peso, eq.garantia || '1 Año', codPDV];
    });
    hojaDetalle.getRange(
      hojaDetalle.getLastRow() + 1, 1, filas.length, filas[0].length
    ).setValues(filas);

    var cantidad = equipos.length;

    // Buscar Cod_Inventario para este producto en INVENTARIO
    var invData = hojaInv.getDataRange().getValues();
    var codInv  = null;
    var filaInv = -1;
    for (var i = 1; i < invData.length; i++) {
      if (String(invData[i][1]).trim() === codProd) {
        codInv  = String(invData[i][0]).trim();
        filaInv = i + 1;
        break;
      }
    }
    if (!codInv) {
      return { ok: false, msg: 'Producto no encontrado en INVENTARIO. Verifica el catálogo.' };
    }

    // Actualizar Stock_General_Disponible en INVENTARIO (col 3)
    var stockGenActual = parseInt(invData[filaInv - 1][2]) || 0;
    hojaInv.getRange(filaInv, 3).setValue(stockGenActual + cantidad);

    // Buscar fila en BODEGA para (codInv, codPDV)
    var bodData = hojaBodega.getDataRange().getValues();
    var filaBod = -1;
    for (var i = 1; i < bodData.length; i++) {
      if (String(bodData[i][0]).trim() === codInv &&
          String(bodData[i][1]).trim() === codPDV) {
        filaBod = i + 1;
        break;
      }
    }

    if (filaBod !== -1) {
      var stockBodActual = parseInt(bodData[filaBod - 1][2]) || 0;
      hojaBodega.getRange(filaBod, 3).setValue(stockBodActual + cantidad);
    } else {
      hojaBodega.appendRow([codInv, codPDV, cantidad, 0]);
    }

    try {
      _syncRegistrarEquiposAzure(datos.codProducto, datos.codPDV, equipos);
    } catch(e) {
        Logger.log('sync registrar equipos falló: ' + e.message);
    }

    return {
      ok:  true,
      msg: cantidad + ' equipo' + (cantidad !== 1 ? 's' : '') +
           ' registrado' + (cantidad !== 1 ? 's' : '') + ' correctamente.'
    };

  } catch(e) {
    return { ok: false, msg: 'Error: ' + e.message };
  }
}