/* ==============================================================
   BACK-END: Code.gs
   ============================================================== */

const SPREADSHEET_ID = "1i-dhDgilp1cE847jUwTLnWPhpq-bRHucHXMBD-Em-TU";
const SHEET_USUARIO = "Usuario";

// ── HELPER: verifica si un empleado (rol 1 o 2) está Inactivo en DATOS_EMPLEADO
function _empleadoEstaInactivo(ss, documento, rol) {
  if (String(rol).trim() !== '1' && String(rol).trim() !== '2') return false;
  try {
    const hojaDatos = ss.getSheetByName('DATOS_EMPLEADO');
    if (!hojaDatos) return false;
    const filas = hojaDatos.getDataRange().getValues();
    for (let k = 1; k < filas.length; k++) {
      if (String(filas[k][8]).trim() === String(documento).trim()) { // col I = Documento
        return String(filas[k][6]).trim().toLowerCase() === 'inactivo'; // col G = Estado
      }
    }
  } catch(e) {}
  return false;
}

// 1. MANEJO DE RUTAS Y ARCHIVOS
function doGet(e) {
  var page = e && e.parameter && e.parameter.page;
  if (page === 'home') {
    var sesion = e.parameter.doc ? e.parameter : null;
    var template = HtmlService.createTemplateFromFile('Home');
    template.sesionDocumento = sesion ? String(sesion.doc) : '';
    template.sesionCodRol    = sesion ? String(sesion.rol) : '';
    return template.evaluate()
      .setTitle('Portal TdeA - Home')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Portal TdeA - Login')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function cargarPagina(nombrePagina) {
  return HtmlService.createTemplateFromFile(nombrePagina).evaluate().getContent();
}

function obtenerImagenesMenu() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("MENU");
    const urlScreenshot = sheet.getRange("B3").getValue().toString();
    const urlLogo = sheet.getRange("B4").getValue().toString();
    const forzarImagenDirecta = (url) => {
      if (!url) return "";
      const idMatch = url.match(/[-\w]{25,}/);
      if (idMatch) return "https://lh3.googleusercontent.com/u/0/d/" + idMatch[0];
      return url;
    };
    return {
      screenshot: forzarImagenDirecta(urlScreenshot),
      logo: forzarImagenDirecta(urlLogo)
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

// 2. LOGIN
function validarLogin(usuario, password) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USUARIO);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      let [doc, pass, estado, rol, intentos, token, fechaToken, requiereCambio] = data[i];

      if (doc.toString() !== usuario.toString()) continue;

      // 1. Bloqueado en USUARIO
      if (String(estado).toLowerCase() === 'bloqueado') {
        return { success: false, msg: "Usuario bloqueado. Usa '¿Olvidé mi contraseña?' para recuperar el acceso." };
      }

      // 2. Inactivo en USUARIO
      if (String(estado).toLowerCase() === 'inactivo') {
        return { success: false, msg: "Tu cuenta está inactiva. Contacta al administrador." };
      }

      // 3. Inactivo en DATOS_EMPLEADO (roles 1 y 2) — bloqueo total, sin importar contraseña
      if (_empleadoEstaInactivo(ss, usuario, rol)) {
        return { success: false, msg: "Tu cuenta está inactiva. Contacta al administrador para reactivarla." };
      }

      // 4. Verificar contraseña
      if (pass.toString() !== hashearClave(password)) {
        let nuevosIntentos = (parseInt(intentos) || 0) + 1;
        sheet.getRange(i + 1, 5).setValue(nuevosIntentos);
        if (nuevosIntentos >= 3) {
          sheet.getRange(i + 1, 3).setValue("Bloqueado");
          sheet.getRange(i + 1, 8).setValue("TRUE");
          return { success: false, msg: "Demasiados intentos fallidos. Tu cuenta fue bloqueada. Usa '¿Olvidé mi contraseña?' para recuperar el acceso." };
        }
        const restantes = 3 - nuevosIntentos;
        return { success: false, msg: `Contraseña incorrecta. Te quedan ${restantes} intento(s).` };
      }

      // 5. Contraseña correcta — resetear intentos
      sheet.getRange(i + 1, 5).setValue(0);

      // 6. Requiere cambio obligatorio de contraseña
      if (String(requiereCambio).toUpperCase() === 'TRUE') {
        return { success: false, requiereCambio: true, msg: "Debes cambiar tu contraseña antes de continuar." };
      }

      return { success: true, msg: "Login exitoso", documento: doc.toString(), codRol: rol.toString() };
    }

    return { success: false, msg: "Usuario no encontrado." };
  } catch (e) {
    return { success: false, msg: "Error de conexión: " + e.message };
  }
}

// 3. REGISTRO
function registrarUsuarioYCliente(documento, password, rol, correo, datosCliente) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (rol === "1" || rol === "2") {
      const hojaDatos = ss.getSheetByName("DATOS_EMPLEADO");
      if (!hojaDatos) return { success: false, msg: "Error interno: hoja DATOS_EMPLEADO no encontrada." };
      const datosDatos = hojaDatos.getDataRange().getValues();
      let encontrado = false;
      let cargoValido = false;
      for (let i = 1; i < datosDatos.length; i++) {
        if (String(datosDatos[i][8]).trim() === String(documento).trim()) {
          encontrado = true;
          const cargo = String(datosDatos[i][3]).trim().toLowerCase();
          if (rol === "1" && cargo === "administrador") cargoValido = true;
          if (rol === "2" && cargo === "vendedor")      cargoValido = true;
          break;
        }
      }
      if (!encontrado) return { success: false, msg: "Tu documento no está registrado como empleado." };
      if (!cargoValido) return { success: false, msg: "Tu cargo no corresponde al rol solicitado." };
    }

    // ── Validar correo único ANTES de cualquier inserción ──
    if (rol === "3" && correo) {
      const hojaCliente = ss.getSheetByName('CLIENTE');
      if (hojaCliente) {
        const dataCliente = hojaCliente.getDataRange().getValues();
        for (let i = 1; i < dataCliente.length; i++) {
          if (String(dataCliente[i][7]).trim().toLowerCase() === correo.trim().toLowerCase()) {
            return { success: false, msg: 'Este correo ya está registrado por otro cliente.' };
          }
        }
      }
    }

    // ── Validar documento único ──
    const sheetUser = ss.getSheetByName(SHEET_USUARIO);
    if (!sheetUser) return { success: false, msg: "Error: No existe la hoja de usuarios." };
    const dataUsuarios = sheetUser.getRange("A:A").getValues();
    for (let i = 1; i < dataUsuarios.length; i++) {
      if (dataUsuarios[i][0].toString() === documento.toString()) {
        return { success: false, msg: "El documento ya se encuentra registrado." };
      }
    }

    // ── Todo válido → insertar ──
    sheetUser.appendRow([documento, hashearClave(password), "Activo", rol, 0, "", "", "NO"]);

    if (rol === "3" && datosCliente) {
      const sheetCliente = ss.getSheetByName("CLIENTE");
      if (!sheetCliente) return { success: false, msg: "Error: No existe la hoja 'CLIENTE'." };
      sheetCliente.appendRow([
        documento, datosCliente.tipoDoc, datosCliente.n1, datosCliente.n2,
        datosCliente.a1, datosCliente.a2, datosCliente.tel, correo, new Date()
      ]);
    } 

    // Sync Azure SQL solo para clientes
  if (rol === "3" && datosCliente) {
    try { _syncClienteAzure(documento, datosCliente, correo); } catch(e) {}
    }

return { success: true, msg: "Registro completado con éxito." };
  } catch (e) {
    return { success: false, msg: "Error en el servidor: " + e.message };
  }
}

// 4. RECUPERACIÓN — PASO 1: Generar y enviar token
function generarTokenRecuperacion(documento) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaUsuario = ss.getSheetByName(SHEET_USUARIO);
    const datosUsuario = hojaUsuario.getDataRange().getValues();

    let filaUsuario = -1;
    let codRol = null;
    let estadoUsuario = null;

    for (let i = 1; i < datosUsuario.length; i++) {
      if (String(datosUsuario[i][0]) === String(documento)) {
        filaUsuario  = i + 1;
        codRol       = String(datosUsuario[i][3]);
        estadoUsuario = String(datosUsuario[i][2]).toLowerCase();
        break;
      }
    }

    if (filaUsuario === -1) {
      return { success: false, msg: 'No existe ninguna cuenta con ese documento.' };
    }

    // Bloquear recuperación si el empleado está Inactivo en DATOS_EMPLEADO
    if (_empleadoEstaInactivo(ss, documento, codRol)) {
      return { success: false, msg: 'Tu cuenta está inactiva. Contacta al administrador; no puedes recuperar el acceso por este medio.' };
    }

    let correo = null;
    if (codRol === '1' || codRol === '2') {
      const hojaContacto = ss.getSheetByName('CONTACTO_EMPLEADO');
      const datosContacto = hojaContacto.getDataRange().getValues();
      for (let j = 1; j < datosContacto.length; j++) {
        if (String(datosContacto[j][3]) === String(documento)) {
          correo = datosContacto[j][2];
          break;
        }
      }
    } else if (codRol === '3') {
      const hojaCliente = ss.getSheetByName('CLIENTE');
      const datosCliente = hojaCliente.getDataRange().getValues();
      for (let k = 1; k < datosCliente.length; k++) {
        if (String(datosCliente[k][0]) === String(documento)) {
          correo = datosCliente[k][7];
          break;
        }
      }
    }

    if (!correo) {
      return { success: false, msg: 'No se encontró un correo asociado a este documento.' };
    }

    const token = Math.random().toString(36).substring(2, 8).toUpperCase();
    const expiracion = new Date(new Date().getTime() + 5 * 60 * 1000);
    hojaUsuario.getRange(filaUsuario, 6).setValue(token);
    hojaUsuario.getRange(filaUsuario, 7).setValue(expiracion);

    MailApp.sendEmail({
      to: correo,
      subject: 'Recuperación de contraseña - Claro Colombia',
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:480px;margin:auto;padding:30px;border:1px solid #eee;border-radius:8px;">' +
        '<h2 style="color:#ef3340;">Recuperación de contraseña</h2>' +
        '<p>Tu código de recuperación es:</p>' +
        '<div style="font-size:32px;font-weight:bold;letter-spacing:8px;color:#333;background:#f5f5f5;padding:16px;border-radius:6px;text-align:center;">' + token + '</div>' +
        '<p style="color:#888;font-size:13px;margin-top:16px;">Este código expira en <strong>5 minutos</strong>. Si no solicitaste este cambio, ignora este mensaje.</p>' +
        '<hr style="border:none;border-top:1px solid #eee;margin:20px 0;">' +
        '<p style="color:#aaa;font-size:11px;">Claro Colombia &bull; TdeA</p>' +
        '</div>'
    });

    const correoOculto = correo.replace(/(.{2})(.*)(@.*)/, '$1***$3');
    return { success: true, msg: 'Token enviado a ' + correoOculto };
  } catch (e) {
    return { success: false, msg: 'Error interno: ' + e.message };
  }
}

// 4. RECUPERACIÓN — PASO 2: Validar token y cambiar clave
function validarTokenYCambiarClave(token, nuevaClave) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaUsuario = ss.getSheetByName(SHEET_USUARIO);
    const datosUsuario = hojaUsuario.getDataRange().getValues();

    const ahora = new Date();
    let filaUsuario = -1;
    let rolUsuario  = null;
    let docUsuario  = null;

    for (let i = 1; i < datosUsuario.length; i++) {
      const tokenGuardado = String(datosUsuario[i][5]);
      const fechaExp = datosUsuario[i][6];
      if (tokenGuardado === String(token).toUpperCase()) {
        if (new Date(fechaExp) < ahora) {
          hojaUsuario.getRange(i + 1, 6).setValue('');
          hojaUsuario.getRange(i + 1, 7).setValue('');
          return { success: false, msg: 'El token ha expirado. Solicita uno nuevo.' };
        }
        filaUsuario = i + 1;
        rolUsuario  = String(datosUsuario[i][3]);
        docUsuario  = String(datosUsuario[i][0]);
        break;
      }
    }

    if (filaUsuario === -1) {
      return { success: false, msg: 'Token inválido. Verifica el código ingresado.' };
    }

    // Bloquear cambio de clave si el empleado está Inactivo en DATOS_EMPLEADO
    if (_empleadoEstaInactivo(ss, docUsuario, rolUsuario)) {
      return { success: false, msg: 'Tu cuenta está inactiva. Contacta al administrador; no puedes cambiar la contraseña.' };
    }

    hojaUsuario.getRange(filaUsuario, 2).setValue(hashearClave(nuevaClave));
    hojaUsuario.getRange(filaUsuario, 5).setValue(0);
    hojaUsuario.getRange(filaUsuario, 6).setValue('');
    hojaUsuario.getRange(filaUsuario, 7).setValue('');
    hojaUsuario.getRange(filaUsuario, 8).setValue('FALSE');

    return { success: true, msg: 'Contraseña actualizada correctamente.' };
  } catch (e) {
    return { success: false, msg: 'Error interno: ' + e.message };
  }
}

function hashearClave(clave) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    clave,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function hashearUsuariosExistentes() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoja = ss.getSheetByName(SHEET_USUARIO);
  const datos = hoja.getDataRange().getValues();
  let actualizados = 0;
  for (let i = 1; i < datos.length; i++) {
    const clave = String(datos[i][1]);
    if (clave.length === 64 && /^[a-f0-9]+$/.test(clave)) continue;
    if (!clave) continue;
    hoja.getRange(i + 1, 2).setValue(hashearClave(clave));
    actualizados++;
  }
  Logger.log('Usuarios hasheados: ' + actualizados);
}

function cambiarClaveObligatoria(documento, nuevaClave) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(SHEET_USUARIO);
    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]) === String(documento)) {
        hoja.getRange(i + 1, 2).setValue(hashearClave(nuevaClave));
        hoja.getRange(i + 1, 3).setValue("Activo");
        hoja.getRange(i + 1, 5).setValue(0);
        hoja.getRange(i + 1, 8).setValue("FALSE");
        return { success: true, msg: "Contraseña actualizada." };
      }
    }
    return { success: false, msg: "Usuario no encontrado." };
  } catch (e) {
    return { success: false, msg: "Error: " + e.message };
  }
}

function cargarHome(documento, codRol) {
  var template = HtmlService.createTemplateFromFile('Home');
  template.sesionDocumento = String(documento || '');
  template.sesionCodRol    = String(codRol    || '');
  return template.evaluate().getContent();
}

function getHomeUrl()  { return ScriptApp.getService().getUrl() + '?page=home'; }
function getLoginUrl() { return ScriptApp.getService().getUrl(); }

function resetearPasswordTemporal() {
  var ss    = SpreadsheetApp.openById('1i-dhDgilp1cE847jUwTLnWPhpq-bRHucHXMBD-Em-TU');
  var hoja  = ss.getSheetByName('USUARIO');
  var datos = hoja.getDataRange().getValues();
  var documento  = '1013462095';
  var nuevaClave = 'Vendedor2024';
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === String(documento)) {
      hoja.getRange(i + 1, 2).setValue(hashearClave(nuevaClave));
      hoja.getRange(i + 1, 3).setValue('Activo');
      hoja.getRange(i + 1, 5).setValue(0);
      hoja.getRange(i + 1, 8).setValue('FALSE');
      Logger.log('Listo. Usa: ' + nuevaClave);
      return;
    }
  }
  Logger.log('Usuario no encontrado');
}

function obtenerMiPerfil(documentoUsuario) {
  try {
    var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    var data = ss.getSheetByName('CLIENTE').getDataRange().getValues();
    var doc  = String(documentoUsuario).trim();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== doc) continue;
      return {
        ok:        true,
        documento: String(data[i][0]).trim(),
        tipoDoc:   String(data[i][1]).trim(),
        nombre:    [data[i][2], data[i][3]].filter(Boolean).join(' '),
        apellido:  [data[i][4], data[i][5]].filter(Boolean).join(' '),
        telefono:  String(data[i][6]).trim(),
        correo:    String(data[i][7]).trim()
      };
    }
    return { ok: false, msg: 'No se encontró el perfil.' };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

function actualizarMiPerfil(documentoUsuario, telefono, correo) {
  try {
    var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('CLIENTE');
    var data  = sheet.getDataRange().getValues();
    var doc   = String(documentoUsuario).trim();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== doc) continue;
      sheet.getRange(i + 1, 7).setValue(telefono); // col G = Telefono
      sheet.getRange(i + 1, 8).setValue(correo);   // col H = Correo
      return { ok: true, msg: 'Perfil actualizado correctamente.' };
    }
    return { ok: false, msg: 'Usuario no encontrado.' };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

// ══════════════════════════════════════════════
// SSO REAL: Login con cuenta Google
// ══════════════════════════════════════════════

function validarLoginSSO(correo) {
  try {
    if (!correo || !correo.includes('@')) {
      return { success: false, msg: 'Correo inválido.' };
    }

    const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    const correoLower = correo.trim().toLowerCase();
    const encontrados = []; // puede haber máximo 2: uno empleado, uno cliente

    // 1. Buscar en CONTACTO_EMPLEADO
    const hojaContacto = ss.getSheetByName('CONTACTO_EMPLEADO');
    if (hojaContacto) {
      const rows = hojaContacto.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][2]).trim().toLowerCase() === correoLower) {
          encontrados.push({ documento: String(rows[i][3]).trim(), origen: 'empleado' });
          break; // correo único por hoja, no seguir buscando
        }
      }
    }

    // 2. Buscar en CLIENTE
    const hojaCliente = ss.getSheetByName('CLIENTE');
    if (hojaCliente) {
      const rows = hojaCliente.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][7]).trim().toLowerCase() === correoLower) {
          encontrados.push({ documento: String(rows[i][0]).trim(), origen: 'cliente' });
          break; // correo único por hoja
        }
      }
    }

    // No existe en ninguna hoja
    if (encontrados.length === 0) {
      return { success: false, msg: 'El correo ' + correo + ' no está registrado. Contacta al administrador.' };
    }

    // Existe en ambas hojas → preguntar al usuario
    if (encontrados.length === 2) {
      return { success: false, seleccionRol: true, correo: correo,
               msg: 'Tu correo está asociado a dos perfiles. ¿Con cuál deseas ingresar?' };
    }

    // Solo en una hoja → validar directamente
    return _validarDocumentoEnUsuario(ss, encontrados[0].documento);

  } catch (e) {
    return { success: false, msg: 'Error interno: ' + e.message };
  }
}

function validarLoginSSOConRol(correo, origen) {
  try {
    const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    const correoLower = correo.trim().toLowerCase();
    let documento     = null;

    if (origen === 'empleado') {
      const hojaContacto = ss.getSheetByName('CONTACTO_EMPLEADO');
      const rows = hojaContacto.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][2]).trim().toLowerCase() === correoLower) {
          documento = String(rows[i][3]).trim();
          break;
        }
      }
    } else {
      const hojaCliente = ss.getSheetByName('CLIENTE');
      const rows = hojaCliente.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][7]).trim().toLowerCase() === correoLower) {
          documento = String(rows[i][0]).trim();
          break;
        }
      }
    }

    if (!documento) return { success: false, msg: 'No se encontró el perfil.' };
    return _validarDocumentoEnUsuario(ss, documento);

  } catch (e) {
    return { success: false, msg: 'Error interno: ' + e.message };
  }
}

function _validarDocumentoEnUsuario(ss, documento) {
  const hojaUsuario  = ss.getSheetByName(SHEET_USUARIO);
  const datosUsuario = hojaUsuario.getDataRange().getValues();

  for (let i = 1; i < datosUsuario.length; i++) {
    if (String(datosUsuario[i][0]).trim() !== documento) continue;

    const estado         = String(datosUsuario[i][2]).toLowerCase();
    const rol            = String(datosUsuario[i][3]);
    const requiereCambio = String(datosUsuario[i][7]);

    if (estado === 'bloqueado') return { success: false, msg: 'Cuenta bloqueada. Usa "¿Olvidé mi contraseña?" para recuperar el acceso.' };
    if (estado === 'inactivo')  return { success: false, msg: 'Cuenta inactiva. Contacta al administrador.' };
    if (_empleadoEstaInactivo(ss, documento, rol)) return { success: false, msg: 'Cuenta inactiva. Contacta al administrador.' };
    if (requiereCambio.toUpperCase() === 'TRUE') return { success: false, requiereCambio: true, documento, codRol: rol, msg: 'Debes cambiar tu contraseña antes de continuar.' };

    return { success: true, documento, codRol: rol, msg: 'Acceso concedido' };
  }
  return { success: false, msg: 'El correo está registrado pero no tiene cuenta activa.' };
}

function obtenerUsuarioActual() {
  try {
    var correo = Session.getActiveUser().getEmail();
    if (!correo) return { success: false, msg: 'No se pudo obtener el correo.' };
    return validarLoginSSO(correo);
  } catch(e) {
    return { success: false, msg: e.message };
  }
}


