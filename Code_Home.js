/* ══════════════════════════════════════════════════════════════
   Code_Home.gs  —  Backend del Home
   IMPORTANTE: No tiene doGet(), no tiene myFunction(), 
   no depende de constantes de Code.gs
══════════════════════════════════════════════════════════════ */

var SS_ID = "1i-dhDgilp1cE847jUwTLnWPhpq-bRHucHXMBD-Em-TU";

function _hoja(nombre) {
  var hoja = SpreadsheetApp.openById(SS_ID).getSheetByName(nombre);
  if (!hoja) throw new Error('Hoja no encontrada: "' + nombre + '"');
  return hoja;
}

/* ──────────────────────────────────────────────────────────────
   1. NOMBRE DEL USUARIO LOGUEADO
   Cod_Rol 1 o 2 → EMPLEADO  (A=Doc[0], B=Primer_Nombre[1], D=Primer_Apellido[3])
   Cod_Rol 3     → CLIENTE   (A=Doc[0], C=Primer_Nombre[2], E=Primer_Apellido[4])
────────────────────────────────────────────────────────────── */
function obtenerNombreUsuario(documento) {
  try {
    var ss  = SpreadsheetApp.openById(SS_ID);
    var doc = String(documento).trim();

    var filasUsuario = ss.getSheetByName('USUARIO').getDataRange().getValues();
    var codRol = null;
    for (var i = 1; i < filasUsuario.length; i++) {
      if (String(filasUsuario[i][0]).trim() === doc) {
        codRol = String(filasUsuario[i][3]).trim();
        break;
      }
    }
    if (!codRol) return { ok: false, nombre: doc };

    if (codRol === '1' || codRol === '2') {
      var filas = ss.getSheetByName('EMPLEADO').getDataRange().getValues();
      for (var i = 1; i < filas.length; i++) {
        if (String(filas[i][0]).trim() !== doc) continue;
        var n = [String(filas[i][1]).trim(), String(filas[i][3]).trim()]
                .filter(function(x){ return x.length > 0; }).join(' ');
        return { ok: true, nombre: n || doc };
      }
    }

    if (codRol === '3') {
      var filas = ss.getSheetByName('CLIENTE').getDataRange().getValues();
      for (var i = 1; i < filas.length; i++) {
        if (String(filas[i][0]).trim() !== doc) continue;
        var n = [String(filas[i][2]).trim(), String(filas[i][4]).trim()]
                .filter(function(x){ return x.length > 0; }).join(' ');
        return { ok: true, nombre: n || doc };
      }
    }

    return { ok: false, nombre: doc };
  } catch(e) {
    return { ok: false, nombre: String(documento), msg: e.message };
  }
}

/* ──────────────────────────────────────────────────────────────
   2. CAMBIAR CONTRASEÑA (recibe hashes SHA-256 del cliente)
────────────────────────────────────────────────────────────── */
function cambiarPassword(documento, hashActual, hashNueva) {
  try {
    var hoja = _hoja('USUARIO');
    var data = hoja.getDataRange().getValues();
    var doc  = String(documento).trim();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== doc) continue;

      var fila         = i + 1;
      var passGuardada = String(data[i][1]).trim();

      if (passGuardada !== hashActual)  return { ok: false, msg: 'La contraseña actual es incorrecta.' };
      if (hashNueva    === hashActual)  return { ok: false, msg: 'La nueva contraseña debe ser diferente.' };

      hoja.getRange(fila, 2).setValue(hashNueva);
      hoja.getRange(fila, 5).setValue(0);
      hoja.getRange(fila, 8).setValue(false);

      return { ok: true, msg: 'Contraseña actualizada correctamente.' };
    }
    return { ok: false, msg: 'Usuario no encontrado.' };
  } catch(e) {
    return { ok: false, msg: 'Error: ' + e.message };
  }
}

/* ──────────────────────────────────────────────────────────────
   3. DATOS LABORALES
────────────────────────────────────────────────────────────── */
function obtenerDatosLaborales(documento) {
  try {
    var ss   = SpreadsheetApp.openById(SS_ID);
    var data = ss.getSheetByName('DATOS_EMPLEADO').getDataRange().getValues();
    var doc  = String(documento).trim();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][8]).trim() !== doc) continue;

      var codPDV = String(data[i][9]).trim();
      var res = {
        ok:            true,
        cargo:         String(data[i][3]).trim(),
        horarioInicio: String(data[i][4]).trim(),
        horarioFin:    String(data[i][5]).trim(),
        codigoPDV:     codPDV,
        pdvNombre:     '—'
      };

      if (codPDV) {
        try {
          var pdv = ss.getSheetByName('PUNTO_DE_VENTA').getDataRange().getValues();
          for (var j = 1; j < pdv.length; j++) {
            if (String(pdv[j][0]).trim() === codPDV) {
              res.pdvNombre = String(pdv[j][1]).trim();
              break;
            }
          }
        } catch(e) {}
      }
      return res;
    }
    return { ok: false, msg: 'Datos laborales no encontrados.' };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

/* ──────────────────────────────────────────────────────────────
   4. CATÁLOGO DE PRODUCTOS
────────────────────────────────────────────────────────────── */
function obtenerProductos() {
  try {
    var ss          = SpreadsheetApp.openById(SS_ID);
    var dataProductos = ss.getSheetByName('PRODUCTO').getDataRange().getValues();
    var dataMenu      = ss.getSheetByName('MENU').getDataRange().getValues();

    function extraerIdDrive(url) {
      if (!url) return '';
      var match = String(url).match(/\/d\/([a-zA-Z0-9_-]{25,})/);
      return match ? match[1] : '';
    }

    function urlImagen(id) {
      return id ? 'https://drive.google.com/thumbnail?id=' + id + '&sz=w400' : '';
    }

    var imagenes = {
      'xiaomi':    urlImagen(extraerIdDrive(dataMenu[4][1])),
      's23ultra':  urlImagen(extraerIdDrive(dataMenu[5][1])),
      'samsung':   urlImagen(extraerIdDrive(dataMenu[6][1])),
      'iphone':    urlImagen(extraerIdDrive(dataMenu[7][1])),
      'motorola':  urlImagen(extraerIdDrive(dataMenu[8][1]))
    };

    function asignarImagen(nombre) {
      var n = String(nombre).toLowerCase();
      if (n.includes('xiaomi'))                      return imagenes.xiaomi;
      if (n.includes('s23 ultra') || n.includes('s23ultra')) return imagenes.s23ultra;
      if (n.includes('samsung'))                     return imagenes.samsung;
      if (n.includes('iphone') || n.includes('apple')) return imagenes.iphone;
      if (n.includes('motorola'))                    return imagenes.motorola;
      return '';
    }

    var productos = [];
    for (var i = 1; i < dataProductos.length; i++) {
      if (!dataProductos[i][0] && !dataProductos[i][1]) continue;
      var nombre = String(dataProductos[i][1] || '').trim();
      productos.push({
        codigo:    String(dataProductos[i][0] || '').trim(),
        nombre:    nombre,
        modelo:    String(dataProductos[i][2] || '').trim(),
        precio:    parseFloat(dataProductos[i][3]) || 0,
        imagenUrl: asignarImagen(nombre)
      });
    }

    return { ok: true, productos: productos };
  } catch(e) {
    return { ok: false, msg: e.message, productos: [] };
  }
}

/* ──────────────────────────────────────────────────────────────
   5. ROL Y ESTADO
────────────────────────────────────────────────────────────── */
function obtenerRolUsuario(documento) {
  try {
    var data = _hoja('USUARIO').getDataRange().getValues();
    var doc  = String(documento).trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== doc) continue;
      return {
        ok:             true,
        codRol:         String(data[i][3]).trim(),
        estado:         String(data[i][2]).trim(),
        requiereCambio: data[i][7] === true || String(data[i][7]).toUpperCase() === 'TRUE'
      };
    }
    return { ok: false, msg: 'Usuario no encontrado.' };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

/* ──────────────────────────────────────────────────────────────
   6. INVENTARIO PDV
────────────────────────────────────────────────────────────── */
function obtenerInventarioPDV(codigoPDV) {
  try {
    var ss  = SpreadsheetApp.openById(SS_ID);
    var bod = ss.getSheetByName('BODEGA').getDataRange().getValues();
    var inv = ss.getSheetByName('INVENTARIO').getDataRange().getValues();
    var pro = ss.getSheetByName('PRODUCTO').getDataRange().getValues();

    var mInv = {};
    for (var i = 1; i < inv.length; i++)
      mInv[String(inv[i][0]).trim()] = { cod: String(inv[i][1]).trim(), precio: inv[i][4] };

    var mPro = {};
    for (var i = 1; i < pro.length; i++)
      mPro[String(pro[i][0]).trim()] = { nombre: String(pro[i][1]).trim(), modelo: String(pro[i][2]).trim(), precio: pro[i][3] };

    var items = [];
    var cpv   = String(codigoPDV).trim();
    for (var i = 1; i < bod.length; i++) {
      if (String(bod[i][1]).trim() !== cpv) continue;
      var ci = String(bod[i][0]).trim();
      var iv = mInv[ci] || {};
      var pr = mPro[iv.cod] || {};
      items.push({
        codInventario: ci, codProducto: iv.cod || '—',
        nombre: pr.nombre || '—', modelo: pr.modelo || '—',
        stockDisponible: bod[i][2], stockVendido: bod[i][3],
        precioUnitario: iv.precio || pr.precio || 0
      });
    }
    return { ok: true, items: items };
  } catch(e) {
    return { ok: false, msg: e.message, items: [] };
  }
}

/* ──────────────────────────────────────────────────────────────
   7. VENTAS / FACTURAS
────────────────────────────────────────────────────────────── */
function obtenerVentas(documento, codRol) {
  try {
    var ss  = SpreadsheetApp.openById(SS_ID);
    var fac = ss.getSheetByName('FACTURA').getDataRange().getValues();
    var ven = ss.getSheetByName('VENTA').getDataRange().getValues();

    var mV = {};
    for (var i = 1; i < ven.length; i++) {
      var nf = String(ven[i][1]).trim();
      if (!mV[nf]) mV[nf] = [];
      mV[nf].push({ idVenta: ven[i][0], imei: String(ven[i][2]).trim(), valorVenta: ven[i][3] });
    }

    var facturas = [];
    var doc = String(documento).trim();
    var rol = parseInt(codRol);
    for (var i = 1; i < fac.length; i++) {
      var dc = String(fac[i][3]).trim();
      var de = String(fac[i][4]).trim();
      if (rol === 2 && de !== doc) continue;
      if (rol === 3 && dc !== doc) continue;
      var nf = String(fac[i][0]).trim();
      facturas.push({
        numeroFactura: nf, fechaFactura: String(fac[i][1]).trim(),
        total: fac[i][2], docCliente: dc, docEmpleado: de,
        codigoPDV: String(fac[i][5]).trim(), ventas: mV[nf] || []
      });
    }
    return { ok: true, facturas: facturas };
  } catch(e) {
    return { ok: false, msg: e.message, facturas: [] };
  }
}

/* ──────────────────────────────────────────────────────────────
   REGISTRAR VENTA — Función principal
────────────────────────────────────────────────────────────── */
function registrarVenta(datos) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);

    // 1. VALIDAR ROL
    if (datos.codRol !== 1 && datos.codRol !== 2) {
      return { ok: false, msg: 'No tienes permiso para registrar ventas.' };
    }

    // 2. OBTENER PDV DEL EMPLEADO
    var datosEmp = ss.getSheetByName('DATOS_EMPLEADO').getDataRange().getValues();
    var codPDV   = null;
    for (var i = 1; i < datosEmp.length; i++) {
      if (String(datosEmp[i][8]).trim() === String(datos.documentoEmpleado).trim()) {
        codPDV = String(datosEmp[i][9]).trim();
        break;
      }
    }
    if (!codPDV) return { ok: false, msg: 'No se encontró el punto de venta del empleado.' };

    // 3. VALIDAR CLIENTE
    var hojaUsuario  = ss.getSheetByName('USUARIO').getDataRange().getValues();
    var clienteValido = false;
    for (var i = 1; i < hojaUsuario.length; i++) {
      if (String(hojaUsuario[i][0]).trim() === String(datos.documentoCliente).trim()) {
        if (String(hojaUsuario[i][3]).trim() === '3') clienteValido = true;
        break;
      }
    }
    if (!clienteValido) return { ok: false, msg: 'Cliente no encontrado o no tiene rol de cliente.' };

    // 4. VALIDAR IMEI Y OBTENER DATOS DEL DETALLE
    var detalles    = ss.getSheetByName('DETALLE_PRODUCTO').getDataRange().getValues();
    var codProducto = null;
    var codPDVImei  = null;
    var filaDetalle = -1;
    for (var i = 1; i < detalles.length; i++) {
      if (String(detalles[i][0]).trim() === String(datos.imei).trim()) {
        codProducto = String(detalles[i][1]).trim();
        codPDVImei  = String(detalles[i][7]).trim();
        filaDetalle = i + 1;
        break;
      }
    }
    if (!codProducto) return { ok: false, msg: 'IMEI no encontrado en el sistema.' };

    // 5. VALIDAR QUE EL IMEI PERTENECE A LA SEDE DEL EMPLEADO
    if (codPDVImei !== codPDV) {
      return { ok: false, msg: 'Este equipo no pertenece a tu sede. Pertenece al PDV ' + codPDVImei + '.' };
    }

    // 6. VERIFICAR QUE EL IMEI NO ESTÉ YA VENDIDO
    var ventaData = ss.getSheetByName('VENTA').getDataRange().getValues();
    for (var i = 1; i < ventaData.length; i++) {
      if (String(ventaData[i][2]).trim() === String(datos.imei).trim()) {
        return { ok: false, msg: 'Este IMEI ya fue vendido anteriormente.' };
      }
    }

    // 7. OBTENER Cod_Inventario Y PRECIO desde INVENTARIO
    var invData        = ss.getSheetByName('INVENTARIO').getDataRange().getValues();
    var codInv         = null;
    var precioUnitario = 0;
    var filaInv        = -1;
    for (var i = 1; i < invData.length; i++) {
      if (String(invData[i][1]).trim() === String(codProducto).trim()) {
        codInv         = String(invData[i][0]).trim();
        precioUnitario = parseFloat(invData[i][4]) || 0;
        filaInv        = i + 1;
        break;
      }
    }
    if (!codInv) return { ok: false, msg: 'Producto no encontrado en inventario global.' };

    // 8. VALIDAR STOCK EN BODEGA
    var bodData  = ss.getSheetByName('BODEGA').getDataRange().getValues();
    var filaBod  = -1;
    var stockDisp = 0;
    for (var i = 1; i < bodData.length; i++) {
      if (String(bodData[i][0]).trim() === codInv &&
          String(bodData[i][1]).trim() === codPDV) {
        stockDisp = parseInt(bodData[i][2]) || 0;
        filaBod   = i + 1;
        break;
      }
    }
    if (filaBod === -1) return { ok: false, msg: 'Este producto no está disponible en tu sede.' };
    if (stockDisp <= 0) return { ok: false, msg: 'Sin stock disponible en tu sede.' };

    // 9. GENERAR NÚMERO DE FACTURA ÚNICO
    var hojaFac    = ss.getSheetByName('FACTURA');
    var facData    = hojaFac.getDataRange().getValues();
    var numFactura = 'FAC-' + String(Date.now()).slice(-6) + '-' + String(facData.length);

    // 10. REGISTRAR FACTURA
    var fechaHoy = new Date();
    hojaFac.appendRow([
      numFactura,
      fechaHoy,
      precioUnitario,
      datos.documentoCliente,
      datos.documentoEmpleado,
      codPDV
    ]);

    // 11. REGISTRAR VENTA
    var hojaVenta = ss.getSheetByName('VENTA');
    var idVenta   = 'V-' + String(Date.now());
    hojaVenta.appendRow([idVenta, numFactura, datos.imei, precioUnitario]);

    // 12. ACTUALIZAR STOCK BODEGA
    var hojaBod = ss.getSheetByName('BODEGA');
    hojaBod.getRange(filaBod, 3).setValue(stockDisp - 1);
    hojaBod.getRange(filaBod, 4).setValue((parseInt(bodData[filaBod-1][3]) || 0) + 1);

    // 13. ACTUALIZAR INVENTARIO GLOBAL
    var hojaInv     = ss.getSheetByName('INVENTARIO');
    var stockGenDis = parseInt(invData[filaInv-1][2]) || 0;
    var stockGenVen = parseInt(invData[filaInv-1][3]) || 0;
    hojaInv.getRange(filaInv, 3).setValue(stockGenDis - 1);
    hojaInv.getRange(filaInv, 4).setValue(stockGenVen + 1);

    // Sync Azure SQL
    try {
      _syncVentaAzure(numFactura, fechaHoy.toISOString(), precioUnitario,
        datos.documentoCliente, datos.documentoEmpleado, codPDV,
        [{ imei: datos.imei, valor: precioUnitario }]);
    } catch(e) {}

    return {
      ok:         true,
      msg:        'Venta registrada exitosamente.',
      numFactura: numFactura,
      total:      precioUnitario,
      fecha:      fechaHoy.toLocaleDateString('es-CO')
    };

  } catch(e) {
    return { ok: false, msg: 'Error interno: ' + e.message };
  }
}

/* ──────────────────────────────────────────────────────────────
   BUSCAR PRODUCTO POR IMEI
────────────────────────────────────────────────────────────── */
function buscarProductoPorIMEI(imei) {
  try {
    var ss      = SpreadsheetApp.openById(SS_ID);
    var detalle = ss.getSheetByName('DETALLE_PRODUCTO').getDataRange().getValues();
    var imeiStr = String(imei).trim();

    for (var i = 1; i < detalle.length; i++) {
      if (String(detalle[i][0]).trim() !== imeiStr) continue;

      var codProd  = String(detalle[i][1]).trim();
      var codPDV   = String(detalle[i][7]).trim();

      // Verificar si ya fue vendido
      var ventas = ss.getSheetByName('VENTA').getDataRange().getValues();
      for (var k = 1; k < ventas.length; k++) {
        if (String(ventas[k][2]).trim() === imeiStr) {
          return { ok: false, msg: 'Este IMEI ya fue vendido.' };
        }
      }

      // Obtener nombre y precio del producto
      var productos  = ss.getSheetByName('PRODUCTO').getDataRange().getValues();
      var nombreProd = '—';
      var precio     = 0;
      for (var j = 1; j < productos.length; j++) {
        if (String(productos[j][0]).trim() === codProd) {
          nombreProd = String(productos[j][1]).trim();
          precio     = parseFloat(productos[j][3]) || 0;
          break;
        }
      }

      // Obtener nombre del PDV
      var pdvData  = ss.getSheetByName('PUNTO_DE_VENTA').getDataRange().getValues();
      var pdvNombre = '—';
      for (var p = 1; p < pdvData.length; p++) {
        if (String(pdvData[p][0]).trim() === codPDV) {
          pdvNombre = String(pdvData[p][1]).trim();
          break;
        }
      }

      return {
        ok:        true,
        imei:      imeiStr,
        codigo:    codProd,
        nombre:    nombreProd,
        color:     String(detalle[i][3]).trim(),
        storage:   String(detalle[i][4]).trim(),
        precio:    precio,
        codPDV:    codPDV,
        pdvNombre: pdvNombre
      };
    }
    return { ok: false, msg: 'IMEI no encontrado.' };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

/* ──────────────────────────────────────────────────────────────
   BUSCAR CLIENTE POR DOCUMENTO
────────────────────────────────────────────────────────────── */
function buscarCliente(documento) {
  try {
    var ss      = SpreadsheetApp.openById(SS_ID);
    var cliente = ss.getSheetByName('CLIENTE').getDataRange().getValues();
    var doc     = String(documento).trim();

    for (var i = 1; i < cliente.length; i++) {
      if (String(cliente[i][0]).trim() === doc) {
        return {
          ok:       true,
          nombre:   String(cliente[i][2]).trim() + ' ' + String(cliente[i][4]).trim(),
          telefono: String(cliente[i][6]).trim(),
          correo:   String(cliente[i][7]).trim()
        };
      }
    }
    return { ok: false, msg: 'Cliente no encontrado.' };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

/* ──────────────────────────────────────────────────────────────
   CARRITO — Obtener IMEIs disponibles por producto
────────────────────────────────────────────────────────────── */
function obtenerIMEIsDisponibles(codProducto) {
  try {
    var ss      = SpreadsheetApp.openById(SS_ID);
    var detalle = ss.getSheetByName('DETALLE_PRODUCTO').getDataRange().getValues();
    var ventas  = ss.getSheetByName('VENTA').getDataRange().getValues();

    var vendidos = {};
    for (var i = 1; i < ventas.length; i++) {
      vendidos[String(ventas[i][2]).trim()] = true;
    }

    var disponibles = [];
    for (var i = 1; i < detalle.length; i++) {
      var imei    = String(detalle[i][0]).trim();
      var codProd = String(detalle[i][1]).trim();
      if (codProd !== String(codProducto).trim()) continue;
      if (vendidos[imei]) continue;
      disponibles.push({
        imei:   imei,
        color:  String(detalle[i][3]).trim(),
        codPDV: String(detalle[i][7]).trim()
      });
    }
    return { ok: true, disponibles: disponibles };
  } catch(e) {
    return { ok: false, msg: e.message, disponibles: [] };
  }
}

/* ──────────────────────────────────────────────────────────────
   CARRITO — Registrar venta desde cliente (rol 3)
────────────────────────────────────────────────────────────── */
function registrarVentaCliente(datos) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);

    // 1. VALIDAR CLIENTE
    var hojaUsuario = ss.getSheetByName('USUARIO').getDataRange().getValues();
    var clienteValido = false;
    for (var i = 1; i < hojaUsuario.length; i++) {
      if (String(hojaUsuario[i][0]).trim() === String(datos.documentoCliente).trim()) {
        if (String(hojaUsuario[i][3]).trim() === '3') clienteValido = true;
        break;
      }
    }
    if (!clienteValido) return { ok: false, msg: 'Cliente no válido.' };

    // 2. VALIDAR QUE HAY ITEMS
    if (!datos.items || datos.items.length === 0) {
      return { ok: false, msg: 'El carrito está vacío.' };
    }

    var detalles = ss.getSheetByName('DETALLE_PRODUCTO').getDataRange().getValues();
    var invData  = ss.getSheetByName('INVENTARIO').getDataRange().getValues();
    var bodData  = ss.getSheetByName('BODEGA').getDataRange().getValues();
    var venData  = ss.getSheetByName('VENTA').getDataRange().getValues();
    var prodData = ss.getSheetByName('PRODUCTO').getDataRange().getValues();

    var vendidos = {};
    for (var i = 1; i < venData.length; i++) {
      vendidos[String(venData[i][2]).trim()] = true;
    }

    // 3. VALIDAR CADA ITEM Y RECOLECTAR DATOS
    var lineas      = [];
    var totalGlobal = 0;

    for (var k = 0; k < datos.items.length; k++) {
      var item    = datos.items[k];
      var imeiStr = String(item.imei).trim();

      if (vendidos[imeiStr]) {
        return { ok: false, msg: 'El equipo con IMEI ' + imeiStr + ' ya no está disponible.' };
      }

      var codProducto = null;
      var codPDV      = null;
      var filaDetalle = -1;
      for (var i = 1; i < detalles.length; i++) {
        if (String(detalles[i][0]).trim() === imeiStr) {
          codProducto = String(detalles[i][1]).trim();
          codPDV      = String(detalles[i][7]).trim();
          filaDetalle = i;
          break;
        }
      }
      if (!codProducto) return { ok: false, msg: 'IMEI no encontrado: ' + imeiStr };

      var precio = 0;
      for (var i = 1; i < prodData.length; i++) {
        if (String(prodData[i][0]).trim() === codProducto) {
          precio = parseFloat(prodData[i][3]) || 0;
          break;
        }
      }

      var codInv  = null;
      var filaInv = -1;
      for (var i = 1; i < invData.length; i++) {
        if (String(invData[i][1]).trim() === codProducto) {
          codInv  = String(invData[i][0]).trim();
          filaInv = i;
          break;
        }
      }
      if (!codInv) return { ok: false, msg: 'Producto no encontrado en inventario.' };

      var filaBod   = -1;
      var stockDisp = 0;
      for (var i = 1; i < bodData.length; i++) {
        if (String(bodData[i][0]).trim() === codInv &&
            String(bodData[i][1]).trim() === codPDV) {
          stockDisp = parseInt(bodData[i][2]) || 0;
          filaBod   = i;
          break;
        }
      }
      if (filaBod === -1) return { ok: false, msg: 'Producto sin sede asignada.' };
      if (stockDisp <= 0) return { ok: false, msg: 'Sin stock para IMEI ' + imeiStr + '.' };

      totalGlobal += precio;
      lineas.push({
        imei: imeiStr, codProducto: codProducto, codPDV: codPDV,
        precio: precio, codInv: codInv,
        filaInv: filaInv, filaBod: filaBod,
        stockDisp: stockDisp,
        stockVend: parseInt(bodData[filaBod][3]) || 0,
        stockGenDis: parseInt(invData[filaInv][2]) || 0,
        stockGenVen: parseInt(invData[filaInv][3]) || 0
      });
    }

    // 4. GENERAR FACTURA ÚNICA para todo el carrito
    var hojaFac    = ss.getSheetByName('FACTURA');
    var facRows    = hojaFac.getDataRange().getValues();
    var numFactura = 'FAC-' + String(Date.now()).slice(-6) + '-' + String(facRows.length);
    var fechaHoy   = new Date();

    hojaFac.appendRow([
      numFactura,
      fechaHoy,
      totalGlobal,
      datos.documentoCliente,
      'ONLINE',
      lineas[0].codPDV
    ]);

    // 5. REGISTRAR CADA VENTA Y ACTUALIZAR STOCK
    var hojaVenta = ss.getSheetByName('VENTA');
    var hojaInv   = ss.getSheetByName('INVENTARIO');
    var hojaBod   = ss.getSheetByName('BODEGA');

    for (var k = 0; k < lineas.length; k++) {
      var l = lineas[k];
      hojaVenta.appendRow(['V-' + Date.now() + '-' + k, numFactura, l.imei, l.precio]);
      hojaBod.getRange(l.filaBod + 1, 3).setValue(l.stockDisp - 1);
      hojaBod.getRange(l.filaBod + 1, 4).setValue(l.stockVend + 1);
      hojaInv.getRange(l.filaInv + 1, 3).setValue(l.stockGenDis - 1);
      hojaInv.getRange(l.filaInv + 1, 4).setValue(l.stockGenVen + 1);
    }

    // Sync Azure SQL
    try {
      _syncVentaAzure(numFactura, fechaHoy.toISOString(), totalGlobal,
        datos.documentoCliente, 'ONLINE', lineas[0].codPDV,
        lineas.map(function(l) { return { imei: l.imei, valor: l.precio }; }));
    } catch(e) {}

    return {
      ok:         true,
      msg:        'Compra realizada exitosamente.',
      numFactura: numFactura,
      total:      totalGlobal,
      fecha:      fechaHoy.toLocaleDateString('es-CO'),
      items:      lineas.length
    };

  } catch(e) {
    return { ok: false, msg: 'Error interno: ' + e.message };
  }
}

/* ──────────────────────────────────────────────────────────────
   REGISTRAR VENTA — IMEIs disponibles por PDV del empleado
────────────────────────────────────────────────────────────── */
function obtenerIMEIsDisponiblesPorPDV(documentoEmpleado) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);

    // 1. Obtener el PDV del empleado
    var datosEmp = ss.getSheetByName('DATOS_EMPLEADO').getDataRange().getValues();
    var codPDV   = null;
    for (var i = 1; i < datosEmp.length; i++) {
      if (String(datosEmp[i][8]).trim() === String(documentoEmpleado).trim()) {
        codPDV = String(datosEmp[i][9]).trim();
        break;
      }
    }
    if (!codPDV) return { ok: false, msg: 'No se encontró el PDV del empleado.', items: [] };

    // 2. Obtener IMEIs ya vendidos
    var venData  = ss.getSheetByName('VENTA').getDataRange().getValues();
    var vendidos = {};
    for (var i = 1; i < venData.length; i++) {
      vendidos[String(venData[i][2]).trim()] = true;
    }

    // 3. Filtrar DETALLE_PRODUCTO por PDV y no vendidos
    var detalle  = ss.getSheetByName('DETALLE_PRODUCTO').getDataRange().getValues();
    var prodData = ss.getSheetByName('PRODUCTO').getDataRange().getValues();

    // Mapa codigo → nombre + precio
    var mProd = {};
    for (var i = 1; i < prodData.length; i++) {
      mProd[String(prodData[i][0]).trim()] = {
        nombre: String(prodData[i][1]).trim(),
        precio: parseFloat(prodData[i][3]) || 0
      };
    }

    var items = [];
    for (var i = 1; i < detalle.length; i++) {
      var imei   = String(detalle[i][0]).trim();
      var pdvImei = String(detalle[i][7]).trim();

      if (pdvImei !== codPDV)  continue; // no es de este PDV
      if (vendidos[imei])      continue; // ya fue vendido

      var prod = mProd[String(detalle[i][1]).trim()] || {};
      items.push({
        imei:    imei,
        nombre:  prod.nombre || '—',
        color:   String(detalle[i][3]).trim(),
        storage: String(detalle[i][4]).trim(),
        precio:  prod.precio || 0
      });
    }

    // Ordenar por nombre de producto
    items.sort(function(a, b) { return a.nombre.localeCompare(b.nombre); });

    return { ok: true, codPDV: codPDV, items: items };
  } catch(e) {
    return { ok: false, msg: e.message, items: [] };
  }
}