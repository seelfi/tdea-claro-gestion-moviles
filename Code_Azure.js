// ══════════════════════════════════════════════
// AZURE SQL — Configuración
// ══════════════════════════════════════════════
var AZURE_API = 'https://api-clarotdea-c8egebftbxc7g9hy.brazilsouth-01.azurewebsites.net/api';

// ── Consultar métricas ───────────────────────
function consultarMetricasAzure() {
  try {
    var response = UrlFetchApp.fetch(AZURE_API + '/consultarmetricas', {
      method: 'GET', muteHttpExceptions: true
    });
    var data = JSON.parse(response.getContentText());
    if (data.success) return data;
    return { success: false, msg: data.msg };
  } catch(e) {
    return { success: false, msg: 'Azure no disponible: ' + e.message };
  }
}

// ── Consultar inventario ─────────────────────
function consultarInventarioAzure(codigoPdv) {
  try {
    var url = AZURE_API + '/consultarinventario';
    if (codigoPdv) url += '?pdv=' + codigoPdv;
    var response = UrlFetchApp.fetch(url, { method: 'GET', muteHttpExceptions: true });
    var data = JSON.parse(response.getContentText());
    if (data.success) return data;
    return { success: false, msg: data.msg };
  } catch(e) {
    return { success: false, msg: 'Azure no disponible: ' + e.message };
  }
}

// ── Insertar venta ───────────────────────────
function _syncVentaAzure(numFactura, fecha, total, docCliente, docEmpleado, codigoPdv, items) {
  var payload = {
    numero_factura: numFactura,
    fecha:          fecha,
    total:          total,
    doc_cliente:    docCliente,
    doc_empleado:   docEmpleado || null,
    codigo_pdv:     codigoPdv,
    items:          items
  };
  try {
    var response = UrlFetchApp.fetch(AZURE_API + '/insertarVenta', {
      method:             'POST',
      contentType:        'application/json',
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var resultado = JSON.parse(response.getContentText());
    if (!resultado.success) {
      if (resultado.msg && resultado.msg.indexOf('duplicate') !== -1) {
        Logger.log('Azure: factura ya existe, ignorando — ' + numFactura);
        return;
      }
      _guardarSyncPendiente(numFactura, payload, resultado.msg);
    }
  } catch(e) {
    Logger.log('Azure sync venta falló: ' + e.message);
    _guardarSyncPendiente(numFactura, payload, e.message);
  }
}

// ── Guardar sync pendiente ───────────────────
function _guardarSyncPendiente(numFactura, payload, motivo) {
  try {
    var ss   = SpreadsheetApp.openById(SS_ID);
    var hoja = ss.getSheetByName('SYNC_PENDIENTE');
    if (!hoja) {
      hoja = ss.insertSheet('SYNC_PENDIENTE');
      hoja.appendRow(['Fecha', 'Num_Factura', 'Payload', 'Motivo', 'Estado']);
    }
    hoja.appendRow([new Date(), numFactura, JSON.stringify(payload), motivo, 'PENDIENTE']);
  } catch(e2) {
    Logger.log('No se pudo guardar sync pendiente: ' + e2.message);
  }
}

// ── Reintentar sync pendientes ───────────────
function reintentarSyncPendientes() {
  var ss   = SpreadsheetApp.openById(SS_ID);
  var hoja = ss.getSheetByName('SYNC_PENDIENTE');
  if (!hoja) return;
  var data = hoja.getDataRange().getValues();
  var reintentados = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim() !== 'PENDIENTE') continue;
    try {
      var payload  = JSON.parse(data[i][2]);
      var response = UrlFetchApp.fetch(AZURE_API + '/insertarVenta', {
        method:             'POST',
        contentType:        'application/json',
        payload:            JSON.stringify(payload),
        muteHttpExceptions: true
      });
      var resultado = JSON.parse(response.getContentText());
      if (resultado.success || (resultado.msg && resultado.msg.indexOf('duplicate') !== -1)) {
        hoja.getRange(i + 1, 5).setValue('SINCRONIZADO');
        reintentados++;
      } else {
        hoja.getRange(i + 1, 5).setValue('ERROR: ' + (resultado.msg || ''));
      }
    } catch(e) {
      hoja.getRange(i + 1, 5).setValue('ERROR: ' + e.message);
    }
  }
  Logger.log('Reintento finalizado. Sincronizados: ' + reintentados);
}

// ── Insertar empleado ────────────────────────
function _syncEmpleadoAzure(datos) {
  try {
    var payload = {
      empleado: {
        Documento:        datos.documento,
        Primer_Nombre:    datos.primerNombre,
        Segundo_Nombre:   datos.segundoNombre  || null,
        Primer_Apellido:  datos.primerApellido,
        Segundo_Apellido: datos.segundoApellido || null,
        Tipo_Documento:   datos.tipoDoc
      },
      datosEmpleado: {
        Numero_Contrato:        datos.numContrato    || '',
        Fecha_Ingreso:          datos.fechaIngreso   || new Date().toISOString().split('T')[0],
        Salario_Base:           parseFloat(datos.salario) || 0,
        Cargo:                  datos.cargo          || 'Vendedor',
        Horario_Laboral_Inicio: datos.horIni         || null,
        Horario_Laboral_Fin:    datos.horFin         || null,
        Estado:                 'Activo',
        Correo_Empresa:         datos.correo,
        Documento:              datos.documento,
        Codigo_PDV:             datos.codPDV
      },
      contacto: {
        Correo_Personal: datos.correo,
        Documento:       datos.documento
      }
    };
    UrlFetchApp.fetch(AZURE_API + '/insertarEmpleado', {
      method:             'POST',
      contentType:        'application/json',
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch(e) {
    Logger.log('Azure sync empleado falló: ' + e.message);
  }
}

// ── Actualizar empleado ──────────────────────
function _syncActualizarEmpleadoAzure(datos) {
  try {
    UrlFetchApp.fetch(AZURE_API + '/actualizarEmpleado', {
      method:             'POST',
      contentType:        'application/json',
      payload:            JSON.stringify({
        documento: datos.documento,
        cargo:     datos.cargo  || null,
        estado:    datos.estado || null,
        correo:    datos.correo || null,
        codPDV:    datos.codPDV || null
      }),
      muteHttpExceptions: true
    });
  } catch(e) {
    Logger.log('Azure sync actualizar empleado falló: ' + e.message);
  }
}

// ── Insertar cliente ─────────────────────────
function _syncClienteAzure(documento, datosCliente, correo) {
  try {
    UrlFetchApp.fetch(AZURE_API + '/insertarCliente', {
      method:             'POST',
      contentType:        'application/json',
      payload:            JSON.stringify({
        Documento:        documento,
        Tipo_Documento:   datosCliente.tipoDoc,
        Primer_Nombre:    datosCliente.n1,
        Segundo_Nombre:   datosCliente.n2 || null,
        Primer_Apellido:  datosCliente.a1,
        Segundo_Apellido: datosCliente.a2 || null,
        Telefono:         datosCliente.tel || null,
        Correo:           correo,
        Fecha_Registro:   new Date().toISOString().split('T')[0]
      }),
      muteHttpExceptions: true
    });
  } catch(e) {
    Logger.log('Azure sync cliente falló: ' + e.message);
  }
}

// ── Insertar PDV ─────────────────────────────
function _syncPDVAzure(datos, codigoPDV) {
  try {
    UrlFetchApp.fetch(AZURE_API + '/insertarPDV', {
      method:             'POST',
      contentType:        'application/json',
      payload:            JSON.stringify({
        pdv: {
          Codigo_PDV:  String(codigoPDV),
          Nombre:      datos.nombre,
          Contacto:    datos.contacto || null,
          Horario:     datos.horario  || null,
          NIT_Empresa: '900999888'
        },
        ubicacion: {
          Codigo_PDV:   String(codigoPDV),
          Departamento: datos.departamento || 'Antioquia',
          Ciudad:       datos.ciudad,
          Barrio:       datos.barrio   || null,
          Direccion1:   datos.direccion
        }
      }),
      muteHttpExceptions: true
    });
  } catch(e) {
    Logger.log('Azure sync PDV falló: ' + e.message);
  }
}

// ── Actualizar PDV ───────────────────────────
function _syncActualizarPDVAzure(datos) {
  try {
    UrlFetchApp.fetch(AZURE_API + '/actualizarPDV', {
      method:             'POST',
      contentType:        'application/json',
      payload:            JSON.stringify({
        codigo:       String(datos.codigo),
        nombre:       datos.nombre       || null,
        contacto:     datos.contacto     || null,
        horario:      datos.horario      || null,
        departamento: datos.departamento || null,
        ciudad:       datos.ciudad       || null,
        barrio:       datos.barrio       || null,
        direccion:    datos.direccion    || null
      }),
      muteHttpExceptions: true
    });
  } catch(e) {
    Logger.log('Azure sync actualizar PDV falló: ' + e.message);
  }
}

// ── Registrar equipos ────────────────────────
function _syncRegistrarEquiposAzure(codProducto, codPDV, equipos) {
  try {
    UrlFetchApp.fetch(AZURE_API + '/registrarEquipos', {
      method:             'POST',
      contentType:        'application/json',
      payload:            JSON.stringify({
        codProducto: codProducto,
        codPDV:      codPDV,
        equipos:     equipos
      }),
      muteHttpExceptions: true
    });
  } catch(e) {
    Logger.log('Azure sync registrar equipos falló: ' + e.message);
  }
}

// ── Validar sincronización Sheet vs Azure ────
function validarStockAzureVsSheet() {
  var rAzure = UrlFetchApp.fetch(AZURE_API + '/consultarinventario', {
    method: 'GET', muteHttpExceptions: true
  });
  var azure = JSON.parse(rAzure.getContentText());
  Logger.log('=== AZURE ===');
  if (azure.success) {
    azure.inventario.forEach(function(item) {
      Logger.log(item.Nombre_Producto + ' | Disp: ' + item.stock_total + ' | Vendido: ' + item.vendido_total);
    });
  }
  var ss      = SpreadsheetApp.openById(SS_ID);
  var invData = ss.getSheetByName('INVENTARIO').getDataRange().getValues();
  var proData = ss.getSheetByName('PRODUCTO').getDataRange().getValues();
  var mProd   = {};
  for (var i = 1; i < proData.length; i++) {
    mProd[String(proData[i][0]).trim()] = String(proData[i][1]).trim();
  }
  Logger.log('=== SHEET ===');
  for (var i = 1; i < invData.length; i++) {
    if (!invData[i][0]) continue;
    Logger.log(mProd[String(invData[i][1]).trim()] + ' | Disp: ' + invData[i][2] + ' | Vendido: ' + invData[i][3]);
  }
}