/* ══════════════════════════════════════════════════════════════
   Code_PDV.gs — Módulo de Puntos de Venta (solo Rol 1)
══════════════════════════════════════════════════════════════ */

/* ──────────────────────────────────────────────────────────────
   OBTENER TODOS LOS PDVs CON DETALLE
────────────────────────────────────────────────────────────── */
function obtenerPDVsCompleto() {
  try {
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var pdvData = ss.getSheetByName('PUNTO_DE_VENTA').getDataRange().getValues();
    var ubData  = ss.getSheetByName('UBICACION_PUNTOS').getDataRange().getValues();
    var datData = ss.getSheetByName('DATOS_EMPLEADO').getDataRange().getValues();
    var bodData = ss.getSheetByName('BODEGA').getDataRange().getValues();

    // Mapa ubicacion: codPDV → { departamento, ciudad, barrio, direccion }
    var mUbic = {};
    for (var i = 1; i < ubData.length; i++) {
      var cpv = String(ubData[i][1]).trim();
      if (!mUbic[cpv]) {
        mUbic[cpv] = {
          departamento: String(ubData[i][2]).trim(),
          ciudad:       String(ubData[i][3]).trim(),
          barrio:       String(ubData[i][4]).trim(),
          direccion:    String(ubData[i][5]).trim()
        };
      }
    }

    // Mapa empleados por PDV: codPDV → cantidad
    var mEmpleados = {};
    for (var i = 1; i < datData.length; i++) {
      var cpv = String(datData[i][9]).trim();
      if (!cpv) continue;
      mEmpleados[cpv] = (mEmpleados[cpv] || 0) + 1;
    }

    // Stock total disponible por PDV: codPDV → suma de stock
    var mStock = {};
    for (var i = 1; i < bodData.length; i++) {
      var cpv = String(bodData[i][1]).trim();
      if (!cpv) continue;
      mStock[cpv] = (mStock[cpv] || 0) + (parseInt(bodData[i][2]) || 0);
    }

    var pdvs = [];
    for (var i = 1; i < pdvData.length; i++) {
      if (!pdvData[i][0]) continue;
      var cod = String(pdvData[i][0]).trim();
      var ub  = mUbic[cod] || {};
      pdvs.push({
        codigo:       cod,
        nombre:       String(pdvData[i][1]).trim(),
        contacto:     String(pdvData[i][2]).trim(),
        horario:      String(pdvData[i][3]).trim(),
        nit:          String(pdvData[i][4]).trim(),
        departamento: ub.departamento || '—',
        ciudad:       ub.ciudad       || '—',
        barrio:       ub.barrio       || '—',
        direccion:    ub.direccion    || '—',
        empleados:    mEmpleados[cod] || 0,
        stockTotal:   mStock[cod]     || 0
      });
    }

    return { ok: true, pdvs: pdvs };
  } catch(e) {
    return { ok: false, msg: e.message, pdvs: [] };
  }
}

/* ──────────────────────────────────────────────────────────────
   OBTENER DETALLE DE UN PDV (empleados + stock por producto)
────────────────────────────────────────────────────────────── */
function obtenerDetallePDV(codPDV) {
  try {
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var datData = ss.getSheetByName('DATOS_EMPLEADO').getDataRange().getValues();
    var empData = ss.getSheetByName('EMPLEADO').getDataRange().getValues();
    var bodData = ss.getSheetByName('BODEGA').getDataRange().getValues();
    var invData = ss.getSheetByName('INVENTARIO').getDataRange().getValues();
    var proData = ss.getSheetByName('PRODUCTO').getDataRange().getValues();
    var cod     = String(codPDV).trim();

    // Mapa empleado: documento → nombre completo
    var mEmp = {};
    for (var i = 1; i < empData.length; i++) {
      var doc = String(empData[i][0]).trim();
      mEmp[doc] = [empData[i][1], empData[i][3]].filter(Boolean).join(' ').trim();
    }

    // Empleados de este PDV
    var empleados = [];
    for (var i = 1; i < datData.length; i++) {
      if (String(datData[i][9]).trim() !== cod) continue;
      var doc = String(datData[i][8]).trim();
      empleados.push({
        nombre: mEmp[doc] || doc,
        cargo:  String(datData[i][3]).trim(),
        estado: String(datData[i][6]).trim()
      });
    }

    // Mapa producto: codProducto → nombre
    var mProd = {};
    for (var i = 1; i < proData.length; i++) {
      mProd[String(proData[i][0]).trim()] = String(proData[i][1]).trim();
    }

    // Mapa inventario: codInventario → codProducto
    var mInv = {};
    for (var i = 1; i < invData.length; i++) {
      mInv[String(invData[i][0]).trim()] = String(invData[i][1]).trim();
    }

    // Stock de este PDV por producto
    var stock = [];
    for (var i = 1; i < bodData.length; i++) {
      if (String(bodData[i][1]).trim() !== cod) continue;
      var codInv  = String(bodData[i][0]).trim();
      var codProd = mInv[codInv] || '—';
      stock.push({
        codInventario: codInv,
        producto:      mProd[codProd] || codProd,
        disponible:    parseInt(bodData[i][2]) || 0,
        vendido:       parseInt(bodData[i][3]) || 0
      });
    }

    return { ok: true, empleados: empleados, stock: stock };
  } catch(e) {
    return { ok: false, msg: e.message, empleados: [], stock: [] };
  }
}

/* ──────────────────────────────────────────────────────────────
   CREAR PDV
────────────────────────────────────────────────────────────── */
function crearPDV(datos) {
  try {
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var hojaPDV = ss.getSheetByName('PUNTO_DE_VENTA');
    var hojaUb  = ss.getSheetByName('UBICACION_PUNTOS');

    // Validaciones básicas
    if (!datos.nombre || !datos.ciudad || !datos.direccion) {
      return { ok: false, msg: 'Nombre, ciudad y dirección son obligatorios.' };
    }

    // Generar Codigo_PDV autoincremental
    var pdvData  = hojaPDV.getDataRange().getValues();
    var maxCod   = 0;
    for (var i = 1; i < pdvData.length; i++) {
      var n = parseInt(pdvData[i][0]);
      if (!isNaN(n) && n > maxCod) maxCod = n;
    }
    var nuevoCod = maxCod + 1;

    // Generar Id_Ubicacion autoincremental
    var ubData  = hojaUb.getDataRange().getValues();
    var maxUbId = 0;
    for (var i = 1; i < ubData.length; i++) {
      var n = parseInt(ubData[i][0]);
      if (!isNaN(n) && n > maxUbId) maxUbId = n;
    }
    var nuevoUbId = maxUbId + 1;

    // Verificar nombre duplicado
    for (var i = 1; i < pdvData.length; i++) {
      if (String(pdvData[i][1]).trim().toLowerCase() === datos.nombre.trim().toLowerCase()) {
        return { ok: false, msg: 'Ya existe un punto de venta con ese nombre.' };
      }
    }

    // Insertar en PUNTO_DE_VENTA
    // [Codigo_PDV, Nombre, Contacto, Horario, NIT_Empresa]
    hojaPDV.appendRow([
      nuevoCod,
      datos.nombre.trim(),
      datos.contacto  || '',
      datos.horario   || '',
      '900999888'
    ]);

    // Insertar en UBICACION_PUNTOS
    // [Id_Ubicacion, Codigo_PDV, Departamento, Ciudad, Barrio, Direccion1, Dir2, Comp1, Comp2, Otros]
    hojaUb.appendRow([
      nuevoUbId,
      nuevoCod,
      datos.departamento || 'Antioquia',
      datos.ciudad.trim(),
      datos.barrio       || '',
      datos.direccion.trim(),
      '', '', '', ''
    ]);

    // Sync Azure SQL
  try { _syncPDVAzure(datos, nuevoCod); } catch(e) {}

  return { ok: true, msg: 'Punto de venta creado exitosamente.', codigo: nuevoCod };
  } catch(e) {
    return { ok: false, msg: 'Error: ' + e.message };
  }
}

/* ──────────────────────────────────────────────────────────────
   ACTUALIZAR PDV
────────────────────────────────────────────────────────────── */
function actualizarPDV(datos) {
  try {
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var hojaPDV = ss.getSheetByName('PUNTO_DE_VENTA');
    var hojaUb  = ss.getSheetByName('UBICACION_PUNTOS');
    var cod     = String(datos.codigo).trim();

    // Actualizar PUNTO_DE_VENTA
    var pdvData    = hojaPDV.getDataRange().getValues();
    var encontrado = false;
    for (var i = 1; i < pdvData.length; i++) {
      if (String(pdvData[i][0]).trim() !== cod) continue;
      var fila = i + 1;
      if (datos.nombre)   hojaPDV.getRange(fila, 2).setValue(datos.nombre.trim());
      if (datos.contacto) hojaPDV.getRange(fila, 3).setValue(datos.contacto.trim());
      if (datos.horario)  hojaPDV.getRange(fila, 4).setValue(datos.horario.trim());
      encontrado = true;
      break;
    }
    if (!encontrado) return { ok: false, msg: 'Punto de venta no encontrado.' };

    // Actualizar UBICACION_PUNTOS
    var ubData = hojaUb.getDataRange().getValues();
    for (var i = 1; i < ubData.length; i++) {
      if (String(ubData[i][1]).trim() !== cod) continue;
      var filaUb = i + 1;
      if (datos.departamento) hojaUb.getRange(filaUb, 3).setValue(datos.departamento.trim());
      if (datos.ciudad)       hojaUb.getRange(filaUb, 4).setValue(datos.ciudad.trim());
      if (datos.barrio)       hojaUb.getRange(filaUb, 5).setValue(datos.barrio.trim());
      if (datos.direccion)    hojaUb.getRange(filaUb, 6).setValue(datos.direccion.trim());
      break;
    }

    try {
      _syncActualizarPDVAzure(datos);
    } catch(e) {
    Logger.log('sync actualizar PDV falló: ' + e.message);
    }

    return { ok: true, msg: 'Punto de venta actualizado correctamente.' };
  } catch(e) {
    return { ok: false, msg: 'Error: ' + e.message };
  }
}