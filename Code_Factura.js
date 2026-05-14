function obtenerFacturas(documentoUsuario, codRol) {  // ← recibe parámetros
  try {
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var facData = ss.getSheetByName('FACTURA').getDataRange().getValues();
    var empData = ss.getSheetByName('EMPLEADO').getDataRange().getValues();
    var pdvData = ss.getSheetByName('PUNTO_DE_VENTA').getDataRange().getValues();

    var mPDV = {};
    for (var i = 1; i < pdvData.length; i++) {
      mPDV[String(pdvData[i][0]).trim()] = String(pdvData[i][1]).trim();
    }

    var mEmp = {};
    for (var i = 1; i < empData.length; i++) {
      var doc = String(parseInt(empData[i][0]) || empData[i][0]).trim();
      var nombre = [
        String(empData[i][1] || '').trim(),
        String(empData[i][2] || '').trim(),
        String(empData[i][3] || '').trim(),
        String(empData[i][4] || '').trim()
      ].filter(Boolean).join(' ');
      mEmp[doc] = nombre || doc;
    }

    var facturas = [];
    for (var i = 1; i < facData.length; i++) {
      if (!facData[i][0]) continue;
      var docEmp = String(parseInt(facData[i][4]) || facData[i][4]).trim();
      var codPDV = String(parseInt(facData[i][5]) || facData[i][5]).trim();

      // ← Si es rol 2, solo incluir sus propias facturas
      if (parseInt(codRol) === 2 && docEmp !== String(documentoUsuario).trim()) continue;

      facturas.push({
        numeroFactura:    String(facData[i][0]).trim(),
        fechaFactura:     String(facData[i][1]).trim(),
        total:            parseFloat(facData[i][2]) || 0,
        documentoCliente: String(facData[i][3]).trim(),
        nombreVendedor:   mEmp[docEmp] || docEmp,
        nombrePDV:        mPDV[codPDV] || codPDV
      });
    }

    facturas.sort(function(a, b) {
      return new Date(b.fechaFactura) - new Date(a.fechaFactura);
    });

    return { ok: true, facturas: facturas };
  } catch(e) {
    return { ok: false, msg: e.message, facturas: [] };
  }
}

function listarEstructuraSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var hojas = ss.getSheets();
  var estructura = {};
  
  hojas.forEach(function(hoja) {
    var nombre = hoja.getName();
    var headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    estructura[nombre] = headers.filter(function(h) { return h !== ''; });
  });
  
  Logger.log(JSON.stringify(estructura, null, 2));
}

function generarSQLDesdeSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sql = '';

  const tablas = [
    { hoja: 'EMPRESA',            cols: ['NIT','Nombre_Empresa','Pagina_Web','PBX'] },
    { hoja: 'PUNTO_DE_VENTA',     cols: ['Codigo_PDV','Nombre','Contacto','Horario','NIT_Empresa'] },
    { hoja: 'EMPLEADO',           cols: ['Documento','Primer_Nombre','Segundo_Nombre','Primer_Apellido','Segundo_Apellido','Tipo_Documento'] },
    { hoja: 'CLIENTE',            cols: ['Documento','Tipo_Documento','Primer_Nombre','Segundo_Nombre','Primer_Apellido','Segundo_Apellido','Telefono','Correo','Fecha_Registro'] },
    { hoja: 'PRODUCTO',           cols: ['Codigo_Producto','Nombre_Producto','Modelo','Precio'] },
    { hoja: 'INVENTARIO',         cols: ['Cod_Inventario','Codigo_Producto','Stock_General_Disponible','Stock_General_Vendido','Precio_Unitario','Precio_General'] },
    { hoja: 'DETALLE_PRODUCTO',   cols: ['IMEI','Codigo_Producto','Serial','Color','Tamano','Peso','Garantia','Codigo_PDV'] },
    { hoja: 'DATOS_EMPLEADO',     cols: ['Numero_Contrato','Fecha_Ingreso','Salario_Base','Cargo','Horario_Laboral_Inicio','Horario_Laboral_Fin','Estado','Correo_Empresa','Documento','Codigo_PDV'] },
    { hoja: 'CONTACTO_EMPLEADO',  cols: ['Telefono','Telefono_Empresa','Correo_Personal','Documento'] },
    { hoja: 'PERSONAL_EMPLEADO',  cols: ['Documento','Genero','Estado_Civil','Hijos','Fecha_Nacimiento','RH'] },
    { hoja: 'DIRECCION_EMPLEADO', cols: ['Direccion1','Direccion2','Complemento1','Complemento2','Otros','Documento'] },
    { hoja: 'ACADEMICO_EMPLEADO', cols: ['Titulo','Fecha_Inicio','Fecha_Fin','Institucion','Estado','Documento'] },
    { hoja: 'LABORAL_EMPLEADO',   cols: ['Nombre_Empresa_Anterior','Fecha_Ingreso','Fecha_Fin','Cargo','Contacto_Jefe_Anterior','Funciones','Documento'] },
    { hoja: 'FAMILIAR_EMPLEADO',  cols: ['Nombre','Parentesco','Fecha_Nacimiento','Documento'] },
    { hoja: 'FACTURA',            cols: ['Numero_Factura','Fecha_Factura','Total','Documento_Cliente','Documento_Empleado','Codigo_PDV'] },
    { hoja: 'VENTA',              cols: ['Numero_Factura','IMEI','Valor_Venta'] },
    { hoja: 'BODEGA',             cols: ['Cod_Inventario','Codigo_PDV','Stock_Disponible','Stock_Vendido'] },
  ];

  tablas.forEach(function(t) {
    const hoja = ss.getSheetByName(t.hoja);
    if (!hoja) return;
    const filas = hoja.getDataRange().getValues();
    if (filas.length <= 1) return;

    sql += '\n-- ' + t.hoja + '\n';
    for (let i = 1; i < filas.length; i++) {
      const fila = filas[i];
      if (!fila[0]) continue; // saltar filas vacías
      const valores = t.cols.map(function(col, idx) {
        const val = fila[idx];
        if (val === '' || val === null || val === undefined) return 'NULL';
        if (val instanceof Date) return "'" + Utilities.formatDate(val, 'America/Bogota', 'yyyy-MM-dd') + "'";
        if (typeof val === 'number') return val;
        return "'" + String(val).replace(/'/g, "''") + "'";
      });
      sql += 'INSERT INTO ' + t.hoja + ' (' + t.cols.join(',') + ') VALUES (' + valores.join(',') + ');\n';
    }
  });

  // Dividir en partes porque Logger tiene límite
  const partes = Math.ceil(sql.length / 8000);
  for (let p = 0; p < partes; p++) {
    Logger.log('-- PARTE ' + (p+1) + ' de ' + partes);
    Logger.log(sql.substring(p * 8000, (p+1) * 8000));
  }

  Logger.log('-- FIN. Total caracteres: ' + sql.length);
}
