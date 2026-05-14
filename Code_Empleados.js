/* ══════════════════════════════════════════════════════════════
   Code_Empleados.gs — Módulo de Empleados (solo Rol 1)
══════════════════════════════════════════════════════════════ */

/* ──────────────────────────────────────────────────────────────
   OBTENER TODOS LOS EMPLEADOS
────────────────────────────────────────────────────────────── */
function obtenerEmpleados() {
  try {
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var empData = ss.getSheetByName('EMPLEADO').getDataRange().getValues();
    var datData = ss.getSheetByName('DATOS_EMPLEADO').getDataRange().getValues();
    var pdvData = ss.getSheetByName('PUNTO_DE_VENTA').getDataRange().getValues();
    var ubData  = ss.getSheetByName('UBICACION_PUNTOS').getDataRange().getValues();

    // Mapa PDV: codPDV → { nombre, ciudad }
    var mPDV = {};
    for (var i = 1; i < pdvData.length; i++) {
      mPDV[String(pdvData[i][0]).trim()] = String(pdvData[i][1]).trim();
    }

    // Mapa Ubicacion: codPDV → ciudad
    var mUbic = {};
    for (var i = 1; i < ubData.length; i++) {
      var cpv = String(ubData[i][1]).trim();
      if (!mUbic[cpv]) mUbic[cpv] = String(ubData[i][3]).trim(); // Ciudad
    }

    // Mapa datos laborales: documento → objeto
    var mDatos = {};
    for (var i = 1; i < datData.length; i++) {
      var doc = String(datData[i][8]).trim(); // col I = Documento
      mDatos[doc] = {
        numContrato:  String(datData[i][0]).trim(),
        fechaIngreso: String(datData[i][1]).trim(),
        salario:      parseFloat(datData[i][2]) || 0,
        cargo:        String(datData[i][3]).trim(),
        horIni:       String(datData[i][4]).trim(),
        horFin:       String(datData[i][5]).trim(),
        estado:       String(datData[i][6]).trim(),
        correo:       String(datData[i][7]).trim(),
        codPDV:       String(datData[i][9]).trim()
      };
    }

    var empleados = [];
    for (var i = 1; i < empData.length; i++) {
      if (!empData[i][0]) continue;
      var doc    = String(empData[i][0]).trim();
      var datos  = mDatos[doc] || {};
      var codPDV = datos.codPDV || '';
      empleados.push({
        documento:    doc,
        tipoDoc:      String(empData[i][5]).trim(),
        nombre:       [empData[i][1], empData[i][2]].filter(Boolean).join(' ').trim(),
        apellido:     [empData[i][3], empData[i][4]].filter(Boolean).join(' ').trim(),
        cargo:        datos.cargo    || '—',
        correo:       datos.correo   || '—',
        estado:       datos.estado   || '—',
        codPDV:       codPDV,
        pdvNombre:    mPDV[codPDV]   || '—',
        ciudad:       mUbic[codPDV]  || '—',
        numContrato:  datos.numContrato  || '',
        fechaIngreso: datos.fechaIngreso || '',
        salario:      datos.salario      || 0,
        horIni:       datos.horIni       || '',
        horFin:       datos.horFin       || ''
      });
    }

    // Lista de PDVs para el selector
    var pdvs = Object.keys(mPDV).map(function(k) {
      return { codigo: k, nombre: mPDV[k] };
    });

    return { ok: true, empleados: empleados, pdvs: pdvs };
  } catch(e) {
    return { ok: false, msg: e.message, empleados: [], pdvs: [] };
  }
}

/* ──────────────────────────────────────────────────────────────
   CREAR EMPLEADO
────────────────────────────────────────────────────────────── */
function crearEmpleado(datos) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Validar duplicado
    var empData = ss.getSheetByName('EMPLEADO').getDataRange().getValues();
    for (var i = 1; i < empData.length; i++) {
      if (String(empData[i][0]).trim() === String(datos.documento).trim()) {
        return { ok: false, msg: 'Ya existe un empleado con ese documento.' };
      }
    }

    // Validar correo básico
    if (!datos.correo || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(datos.correo)) {
      return { ok: false, msg: 'El correo ingresado no es válido.' };
    }

    // Insertar en EMPLEADO
    ss.getSheetByName('EMPLEADO').appendRow([
      datos.documento,
      datos.primerNombre,
      datos.segundoNombre || '',
      datos.primerApellido,
      datos.segundoApellido || '',
      datos.tipoDoc
    ]);

    // Insertar en DATOS_EMPLEADO
    ss.getSheetByName('DATOS_EMPLEADO').appendRow([
      datos.numContrato    || '',
      datos.fechaIngreso   || new Date(),
      datos.salario        || 0,
      datos.cargo,
      datos.horIni         || '',
      datos.horFin         || '',
      'Activo',
      datos.correo,
      datos.documento,
      datos.codPDV
    ]);

    // Insertar en CONTACTO_EMPLEADO si existe
    try {
      var contHoja = ss.getSheetByName('CONTACTO_EMPLEADO');
      if (contHoja) {
        contHoja.appendRow([
          '', '', datos.correo, datos.documento
        ]);
      }
    } catch(e2) {}
    
    // Sync Azure SQL (no bloquea si falla)
    try { 
      _syncEmpleadoAzure(datos); 
      } catch(e) { 
        Logger.log('crearEmpleado → sync Azure falló: ' + e.message); 
      }

  return { ok: true, msg: 'Empleado registrado exitosamente.' };
    
  } catch(e) {
    return { ok: false, msg: 'Error: ' + e.message };
  }
}

/* ──────────────────────────────────────────────────────────────
   ACTUALIZAR EMPLEADO
────────────────────────────────────────────────────────────── */
function actualizarEmpleado(datos) {
  try {
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var datHoja = ss.getSheetByName('DATOS_EMPLEADO');
    var datData = datHoja.getDataRange().getValues();
    var doc     = String(datos.documento).trim();

    // Validar correo
    if (datos.correo && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(datos.correo)) {
      return { ok: false, msg: 'El correo ingresado no es válido.' };
    }

    var encontrado = false;
    for (var i = 1; i < datData.length; i++) {
      if (String(datData[i][8]).trim() !== doc) continue;
      var fila = i + 1;
      if (datos.cargo)  datHoja.getRange(fila, 4).setValue(datos.cargo);
      if (datos.estado) datHoja.getRange(fila, 7).setValue(datos.estado);
      if (datos.correo) datHoja.getRange(fila, 8).setValue(datos.correo);
      if (datos.codPDV) datHoja.getRange(fila, 10).setValue(datos.codPDV);

      // Actualizar correo en CONTACTO_EMPLEADO también
      if (datos.correo) {
        try {
          var contHoja = ss.getSheetByName('CONTACTO_EMPLEADO');
          if (contHoja) {
            var contData = contHoja.getDataRange().getValues();
            for (var j = 1; j < contData.length; j++) {
              if (String(contData[j][3]).trim() === doc) {
                contHoja.getRange(j + 1, 3).setValue(datos.correo);
                break;
              }
            }
          }
        } catch(e2) {}
      }

      encontrado = true;
      break;
    }

    if (!encontrado) return { ok: false, msg: 'Empleado no encontrado.' };

    // ← Sync Azure
    try {
      _syncActualizarEmpleadoAzure(datos);
    } catch(e) {
      Logger.log('sync actualizar empleado falló: ' + e.message);
    }

    return { ok: true, msg: 'Empleado actualizado correctamente.' };
  } catch(e) {
    return { ok: false, msg: 'Error: ' + e.message };
  }
}

/* ──────────────────────────────────────────────────────────────
   OBTENER PDVs (para selectores)
────────────────────────────────────────────────────────────── */
function obtenerPDVs() {
  try {
    var pdvData = SpreadsheetApp.openById(SPREADSHEET_ID)
                  .getSheetByName('PUNTO_DE_VENTA').getDataRange().getValues();
    var pdvs = [];
    for (var i = 1; i < pdvData.length; i++) {
      if (!pdvData[i][0]) continue;
      pdvs.push({ codigo: String(pdvData[i][0]).trim(), nombre: String(pdvData[i][1]).trim() });
    }
    return { ok: true, pdvs: pdvs };
  } catch(e) {
    return { ok: false, pdvs: [] };
  }
}

function obtenerMiHorario(documentoUsuario) {
  try {
    var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    var data = ss.getSheetByName('DATOS_EMPLEADO').getDataRange().getDisplayValues(); // ← único cambio
    var doc  = String(documentoUsuario).trim();

    for (var i = 1; i < data.length; i++) {
      var rowDoc = String(parseInt(data[i][8]) || data[i][8]).trim();
      if (rowDoc !== doc) continue;

      function formatearHora(val) {
        var str = String(val).trim();
        var match = str.match(/(\d{1,2}):(\d{2})/);
        if (match) return match[1].padStart(2,'0') + ':' + match[2];
        return str;
      }

      return {
        ok:     true,
        horIni: formatearHora(data[i][4]),
        horFin: formatearHora(data[i][5]),
        estado: String(data[i][6]).trim()
      };
    }
    return { ok: false, msg: 'No se encontró información de horario.' };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}