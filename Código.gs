// =====================================================================
// VARIABLES GLOBALES
// =====================================================================
var FOLDER_ID = '1EUPj6r_VtvQPzUILLa_EARadWQ2YAi1V';
var TEMPLATE_ID = '1NWIA_949kF5TqEQiAQ9EPKloNcKAWVTStguPHpaFjjU'; 
var FOLDER_PLANTILLAS_ID = '16RhCKtWRqvBuBZREd1A-Y296W4Rpuixf';
var USERS_SHEET_ID = '18MRDKeWsbaciGSTrP63bSM-z_l9EvdjKIQZTJ5BjYNo'; 
var BORRADORES_SHEET_ID = '13Dyg_fIxrbsw--0r1k-jLV7TuhPdWOYAr__OgJ06eT8'; 

// ID extraído de la planilla: 00_Memorias_Base de datos
var RUBROS_SHEET_ID = '1e5zU-WiVsRi4WUNdzH9zPNvpA_2td2sas5Ha12C1xhY'; 

// VARIABLES SECCIÓN DOCUMENTOS
var LAMINA_TEMPLATE_ID = '1qaknDVROitOKqXfuyLUkOkLUWC9J5Nl9yGDASZXAtm4';
var FOLDER_CONTENIDOS_ID = '1Gh66FrnRhddUq6vI3iSd1C7G77ctrL_6';

// --- ARRANQUE Y FUSIÓN DE ARCHIVOS ---
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Portal ASSE - Arquitectura')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

function obtenerFechaActualizacion() {
  try {
    var scriptId = ScriptApp.getScriptId();
    var file = DriveApp.getFileById(scriptId);
    var fecha = file.getLastUpdated();
    var meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
    return fecha.getDate() + " de " + meses[fecha.getMonth()] + " de " + fecha.getFullYear();
  } catch (e) {
    return "Actualización reciente";
  }
}

// --- GESTIÓN DE SEGURIDAD Y PERFIL ---
function validarCredenciales(email, password) {
  try {
    var ss = SpreadsheetApp.openById(USERS_SHEET_ID);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var emailBusqueda = email.trim().toLowerCase();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === emailBusqueda && data[i][1].toString() === password) {
        
        // Verificamos permisos desde las columnas F(5), G(6), H(7), I(8), y la NUEVA J(9)
        var esAdminFijo = (emailBusqueda === "ze-sergio@hotmail.com");
        var pAdmin = esAdminFijo || (data[i][5] === true || data[i][5] === "TRUE" || data[i][5] === "VERDADERO");
        
        var pMem = (data[i][6] === "") ? true : (data[i][6] === true || data[i][6] === "TRUE" || data[i][6] === "VERDADERO");
        var pDoc = (data[i][7] === "") ? true : (data[i][7] === true || data[i][7] === "TRUE" || data[i][7] === "VERDADERO");
        var pPla = (data[i][8] === "") ? true : (data[i][8] === true || data[i][8] === "TRUE" || data[i][8] === "VERDADERO");
        
        // Control seguro para la columna 9 (Proyectos) en caso de planillas cortas
        var pProy = true; // Por defecto encendido
        if (data[i].length > 9 && data[i][9] !== "") {
           pProy = (data[i][9] === true || data[i][9] === "TRUE" || data[i][9] === "VERDADERO");
        }

        return { 
          success: true, 
          nombreCompleto: data[i][2] + " " + data[i][3], 
          email: data[i][0],
          permisos: {
            admin: pAdmin,
            memorias: pMem,
            documentos: pDoc,
            plantillas: pPla,
            proyectos: pProy
          }
        };
      }
    }
    return { success: false, error: "Correo o contraseña incorrectos." };
  } catch (e) {
    return { success: false, error: "Error de conexión con la base de datos." };
  }
}

function registrarUsuarioNuevo(email, password, nombre, apellido) {
  try {
    var ss = SpreadsheetApp.openById(USERS_SHEET_ID);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var emailLimpio = email.trim().toLowerCase();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === emailLimpio) {
        return { success: false, error: "Este correo ya está registrado." };
      }
    }
    
    // Al crear usuario: Nombre, Apellido, Fecha, Admin(false), Memorias(true), Documentos(true), Plantillas(true), Proyectos(true)
    sheet.appendRow([emailLimpio, password, nombre.trim(), apellido.trim(), new Date(), false, true, true, true, true]);
    
    // Registramos la acción en la bitácora
    registrarActividad(emailLimpio, nombre + " " + apellido, "se unió al portal de Arquitectura.");
    
    return { success: true };
  } catch (e) { return { success: false, error: "Error al registrar usuario." }; }
}

function obtenerDatosPerfil(email) {
  try {
    var ss = SpreadsheetApp.openById(USERS_SHEET_ID);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var emailBusqueda = email.trim().toLowerCase();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === emailBusqueda) {
        return { nombre: data[i][2], apellido: data[i][3], password: data[i][1] };
      }
    }
    return { error: "Usuario no encontrado." };
  } catch (e) { return { error: e.message }; }
}

function actualizarPerfilUsuario(email, nombre, apellido, password) {
  try {
    var ss = SpreadsheetApp.openById(USERS_SHEET_ID);
    var sheet = ss.getSheets()[0];
    var emailBusqueda = email.trim().toLowerCase();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === emailBusqueda) {
        sheet.getRange(i + 1, 2).setValue(password);
        sheet.getRange(i + 1, 3).setValue(nombre);
        sheet.getRange(i + 1, 4).setValue(apellido);
        
        registrarActividad(email, nombre + " " + apellido, "actualizó sus datos de perfil.");
        return { success: true, nuevoNombre: nombre + " " + apellido };
      }
    }
    return { success: false, error: "Usuario no encontrado." };
  } catch (e) { return { success: false, error: e.message }; }
}

function recuperarClave(email) {
  try {
    var ss = SpreadsheetApp.openById(USERS_SHEET_ID);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var emailBusqueda = email.trim().toLowerCase();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === emailBusqueda) {
        MailApp.sendEmail({
          to: emailBusqueda, subject: "Recuperación de Acceso - Portal ASSE Arquitectura",
          htmlBody: "Hola <b>" + data[i][2] + "</b>,<br><br>Tu contraseña es: <b>" + data[i][1] + "</b><br><br>Saludos."
        });
        return { success: true };
      }
    }
    return { success: false, error: "El correo electrónico no está registrado." };
  } catch (e) { return { success: false, error: "Error al enviar el correo." }; }
}

function limpiarCarpetaAutomaticamente() {
  var carpeta = DriveApp.getFolderById(FOLDER_ID);
  var archivos = carpeta.getFiles();
  var fechaLimite = new Date(new Date().getTime() - (3 * 24 * 60 * 60 * 1000));
  while (archivos.hasNext()) {
    var archivo = archivos.next();
    if (archivo.getDateCreated() < fechaLimite) archivo.setTrashed(true);
  }
}

function obtenerListaUsuarios() {
  try {
    var ss = SpreadsheetApp.openById(USERS_SHEET_ID);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var usuarios = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() !== "") {
        usuarios.push({ email: data[i][0].toString().trim(), nombre: data[i][2].toString().trim() + " " + data[i][3].toString().trim() });
      }
    }
    usuarios.sort(function(a, b) { return a.nombre.localeCompare(b.nombre); });
    return usuarios;
  } catch (e) { return []; }
}

// =====================================================================
// FUNCIONES EXCLUSIVAS DEL ADMINISTRADOR
// =====================================================================

function obtenerListaUsuariosAdmin() {
  try {
    var ss = SpreadsheetApp.openById(USERS_SHEET_ID);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var usuarios = [];
    
    for (var i = 1; i < data.length; i++) {
      var mail = data[i][0].toString().toLowerCase().trim();
      if (mail !== "") {
        var esAdminFijo = (mail === "ze-sergio@hotmail.com");
        var pAdmin = esAdminFijo || (data[i][5] === true || data[i][5] === "TRUE" || data[i][5] === "VERDADERO");
        var pMem = (data[i][6] === "") ? true : (data[i][6] === true || data[i][6] === "TRUE" || data[i][6] === "VERDADERO");
        var pDoc = (data[i][7] === "") ? true : (data[i][7] === true || data[i][7] === "TRUE" || data[i][7] === "VERDADERO");
        var pPla = (data[i][8] === "") ? true : (data[i][8] === true || data[i][8] === "TRUE" || data[i][8] === "VERDADERO");
        
        var pProy = true;
        if (data[i].length > 9 && data[i][9] !== "") {
           pProy = (data[i][9] === true || data[i][9] === "TRUE" || data[i][9] === "VERDADERO");
        }

        usuarios.push({
          email: mail,
          nombre: data[i][2] + " " + data[i][3],
          admin: pAdmin,
          memorias: pMem,
          documentos: pDoc,
          plantillas: pPla,
          proyectos: pProy,
          esSuperAdmin: esAdminFijo // Usado para evitar que te quites los permisos a ti mismo por accidente
        });
      }
    }
    // Ordenamos alfabéticamente
    usuarios.sort(function(a, b) { return a.nombre.localeCompare(b.nombre); });
    return usuarios;
  } catch (e) { return { error: e.toString() }; }
}

function cambiarPermisoUsuario(emailUsuario, modulo, estado) {
  try {
    var ss = SpreadsheetApp.openById(USERS_SHEET_ID);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var emailBusqueda = emailUsuario.toLowerCase();

    // Mapeamos el módulo a su número de columna correspondiente (en setRange empieza en 1)
    var colIndex = -1;
    if (modulo === "admin") colIndex = 6;       // Columna F
    else if (modulo === "memorias") colIndex = 7; // Columna G
    else if (modulo === "documentos") colIndex = 8; // Columna H
    else if (modulo === "plantillas") colIndex = 9; // Columna I
    else if (modulo === "proyectos") colIndex = 10; // Columna J

    if (colIndex === -1) return {success: false, error: "Módulo desconocido."};

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === emailBusqueda) {
         sheet.getRange(i + 1, colIndex).setValue(estado);
         return {success: true};
      }
    }
    return {success: false, error: "Usuario no encontrado."};
  } catch (e) { return {success: false, error: e.toString()}; }
}

// =====================================================================
// MÓDULO: BITÁCORA Y RESUMEN DE ACTIVIDAD (NUEVO)
// =====================================================================

// Llama a esta función para anotar silenciosamente qué hizo cada usuario
function registrarActividad(email, nombre, accion) {
  try {
    var ss = SpreadsheetApp.openById(USERS_SHEET_ID);
    var sheet = ss.getSheetByName('Registro_Actividad');
    
    // Si la hoja no existe, la IA la crea automáticamente la primera vez
    if (!sheet) {
      sheet = ss.insertSheet('Registro_Actividad');
      sheet.appendRow(['Fecha', 'Email', 'Nombre', 'Acción']);
      sheet.getRange("A1:D1").setFontWeight("bold");
    }
    
    var fecha = new Date();
    sheet.appendRow([fecha, email, nombre, accion]);
  } catch(e) {
    // Falla silenciosamente para no interrumpir al usuario si hay un error
    console.error("No se pudo registrar actividad: " + e.message);
  }
}

// Esta es la función que lee el registro y usa a Gemini o lanza el Plan B
function obtenerResumenActividad(nombreUsuarioActivo) {
  try {
    var ss = SpreadsheetApp.openById(USERS_SHEET_ID);
    var sheet = ss.getSheetByName('Registro_Actividad');
    
    // Si la hoja no existe o está vacía
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        tipo: "vacio",
        html: "<span style='color:#888;'>Aún no hay actividad reciente registrada en el portal. ¡Sé el primero en crear algo hoy!</span>"
      };
    }
    
    var data = sheet.getDataRange().getValues();
    var ultimasAcciones = [];
    var maxAcciones = Math.min(data.length - 1, 10); // Agarraremos hasta las últimas 10 acciones
    
    for (var i = data.length - 1; i >= data.length - maxAcciones; i--) {
      var fechaObj = new Date(data[i][0]);
      var diaStr = fechaObj.getDate().toString().padStart(2, '0');
      var mesStr = (fechaObj.getMonth() + 1).toString().padStart(2, '0');
      var horaStr = fechaObj.getHours().toString().padStart(2, '0');
      var minStr = fechaObj.getMinutes().toString().padStart(2, '0');
      
      var strFecha = diaStr + "/" + mesStr + " " + horaStr + ":" + minStr;
      var nom = data[i][2];
      var acc = data[i][3];
      ultimasAcciones.push("• " + strFecha + " - " + nom + " " + acc);
    }
    
    var textoCrudo = ultimasAcciones.join("\n");
    
    // ESTE ES EL PLAN B (LA LISTA CLÁSICA Y SEGURA)
    var htmlRespaldo = "<ul style='margin:0; padding-left: 20px; color:#555; font-size:12px; margin-top: 5px;'>" + 
                       ultimasAcciones.map(function(a){ return "<li style='margin-bottom: 4px;'>"+a+"</li>"; }).join("") + 
                       "</ul>";

    // INTENTAMOS CONECTAR CON GEMINI (PLAN A)
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
       return { tipo: "lista", html: htmlRespaldo }; // Si no hay llave, salta el Plan B automático
    }

    var promptIA = "Eres el presentador amigable del Portal de Arquitectura de ASSE. Escribe UN SOLO PÁRRAFO CORTO (máximo 4 líneas) dirigido a " + nombreUsuarioActivo + " resumiendo las siguientes acciones recientes del equipo. Sé profesional pero cercano. NO uses viñetas, redacta un párrafo continuo.\n\nActividad reciente:\n" + textoCrudo;
    
    var urlIA = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
    var payload = {
      "contents": [{"parts": [{"text": promptIA}]}],
      "generationConfig": {"temperature": 0.4, "maxOutputTokens": 150} 
    };

    var options = {
      'method' : 'post',
      'contentType': 'application/json',
      'payload' : JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    var response = UrlFetchApp.fetch(urlIA, options);
    var json = JSON.parse(response.getContentText());

    // Si Gemini da error o llega al límite, se dispara el PLAN B
    if (json.error || !json.candidates || json.candidates.length === 0) {
       return { tipo: "lista", html: htmlRespaldo };
    }

    var textoIA = json.candidates[0].content.parts[0].text.trim();
    // Limpiamos formato markdown en caso de que la IA responda con negritas (**texto**)
    textoIA = textoIA.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');

    return { tipo: "ia", html: "<span>" + textoIA + "</span>" };

  } catch (e) {
    // Si todo explota, salta este error en vez de romper el portal
    return { tipo: "error", html: "<span style='color:#d32f2f;'>No se pudo cargar la actividad reciente.</span>" };
  }
}

// =====================================================================
// MÓDULO: ASISTENTE VIRTUAL IA
// =====================================================================

function consultarAsistenteIA(preguntaUsuario, emailUsuario) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return "⚠️ Error interno: La llave de la IA no está configurada.\n\nFuente: Interfaz del Portal";

    var ss = SpreadsheetApp.openById(USERS_SHEET_ID);
    var sheetCerebro = ss.getSheetByName('Cerebro_IA');

    // --- MODO ADMINISTRADOR: APRENDIZAJE ---
    var prefijoComando = "APRENDE:";
    if (preguntaUsuario.toUpperCase().startsWith(prefijoComando)) {
        var emailAdmin = "ze-sergio@hotmail.com";
        if (!emailUsuario || emailUsuario.toLowerCase().trim() !== emailAdmin) {
            return "⛔ Permiso denegado. Solo el administrador puede enseñarme nuevas reglas.\n\nFuente: Interfaz del Portal";
        }
        var nuevaRegla = preguntaUsuario.substring(prefijoComando.length).trim();
        sheetCerebro.appendRow([nuevaRegla]);
        return "🧠 ¡Anotado! Regla guardada permanentemente.\n\nFuente: Interfaz del Portal";
    }

    // --- MODO NAVEGACIÓN: DETECTAR URL EN LA PREGUNTA ---
    var infoWebExtra = "";
    var urlRegex = /(https?:\/\/[^\s]+)/g;
    var enlaces = preguntaUsuario.match(urlRegex);

    if (enlaces && enlaces.length > 0) {
      try {
        infoWebExtra = "\n--- CONTENIDO EXTRAÍDO DE LA WEB (" + enlaces[0] + ") ---\n";
        infoWebExtra += extraerTextoDeUrl(enlaces[0]);
        infoWebExtra += "\n--- FIN DEL CONTENIDO WEB ---\n";
      } catch(e) {
        infoWebExtra = "\n(No pude leer el enlace: " + e.message + ")\n";
      }
    }

    // --- CARGA DE CONOCIMIENTOS PREVIOS ---
    var datosCerebro = sheetCerebro.getDataRange().getValues();
    var reglasVivas = "";
    if (datosCerebro.length > 1) {
        reglasVivas = datosCerebro.slice(1).map(r => "- " + r[0]).join("\n");
    }

    // --- LECTURA DE HTML DEL PORTAL ---
    var codigoPortal = HtmlService.createTemplateFromFile('Index').evaluate().getContent();
    var codigoEstructura = codigoPortal.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');

    // --- LECTURA DE LA BASE DE DATOS DE RUBROS ---
    var infoRubros = "";
    try {
      var ssRubros = SpreadsheetApp.openById(RUBROS_SHEET_ID);
      var hojasRubros = ssRubros.getSheets();
      
      for (var h = 0; h < hojasRubros.length; h++) {
        var nombreHoja = hojasRubros[h].getName();
        var dataRubros = hojasRubros[h].getDataRange().getDisplayValues();
        var maxFilas = Math.min(dataRubros.length, 200); 
        
        var bA = "", pA = "", sA = "", tA = "";
        
        for (var j = 1; j < maxFilas; j++) {
          var fila = dataRubros[j];
          if(fila.join("").trim() === "") continue;

          var bloque = fila[1] ? fila[1].toString().trim().toUpperCase() : "";
          var principal = fila[2] ? fila[2].toString().trim() : "";
          var secundario = fila[3] ? fila[3].toString().trim() : "";
          var terciario = fila[4] ? fila[4].toString().trim() : "";
          var texto = fila[5] ? fila[5].toString().trim() : "";

          if (bloque) bA = bloque;
          if (principal) { pA = principal; sA = ""; tA = ""; }
          else if (secundario) { sA = secundario; tA = ""; }
          else if (terciario) { tA = terciario; }

          if (texto) {
             var path = nombreHoja;
             if (bA) path += " - BLOQUE " + bA;
             
             var jerarquia = [];
             if (pA) jerarquia.push(pA);
             if (sA) jerarquia.push(sA);
             if (tA) jerarquia.push(tA);
             
             if (jerarquia.length > 0) path += " - " + jerarquia.join(" / ");
             infoRubros += "[Fuente: " + path + "] -> Texto: " + texto + "\n";
          }
        }
      }
    } catch(err) {
      infoRubros = "(No se pudo cargar la base de datos de rubros en este momento)";
    }

    var contextoOculto = "Eres EXCLUSIVAMENTE el Asistente del Portal de Arquitectura de ASSE. Tu único propósito es ayudar con el uso del portal, normativas de construcción, rubros, y la información provista en las planillas.\n" +
      "REGLA DE ORO 1: NUNCA le digas al usuario que estás leyendo código HTML o una planilla de Excel. Responde con naturalidad.\n" +
      "REGLA DE ORO 2: Responde conciso (1-2 párrafos).\n" +
      "REGLA DE ORO 3 (CERCO PERIMETRAL STRICTO): Si la pregunta del usuario trata sobre temas generales ajenos a la arquitectura, construcción, el funcionamiento de ASSE o tu portal (por ejemplo: deportes, celebridades, recetas, historia general, etc.), TIENES PROHIBIDO RESPONDER. Debes contestar EXACTAMENTE con esta frase: 'Lo siento, como asistente técnico del Portal de Arquitectura, solo estoy autorizado a responder consultas operativas, normativas o técnicas relacionadas con el sistema y los rubros constructivos.'\n" +
      "REGLA DE ORO 4 (FUENTES MÚLTIPLES OBLIGATORIAS): Al final de TODAS tus respuestas, debes indicar las fuentes utilizadas. Si usas información de la Base de datos de Rubros, copia exactamente la ruta que aparece entre corchetes (Ejemplo: 'Sanitaria - BLOQUE A - CONDICIONES PREVIAS / Cañerías'). Otras fuentes válidas son: 'Regla aprendida del Administrador', 'Interfaz del Portal', o 'Navegación Web'.\n" +
      "   - Si usas UNA SOLA fuente, escribe al final: 'Fuente: [nombre de la fuente]'.\n" +
      "   - Si usas DOS O MÁS fuentes, debes obligatoriamente escribir 'Fuentes:' y enumerarlas debajo en forma de lista (1. [Fuente uno], 2. [Fuente dos], etc).\n\n" +
      "REGLAS APRENDIDAS DE TU ADMINISTRADOR:\n" + reglasVivas + "\n\n" +
      infoWebExtra + "\n\n" + 
      "--- BASE DE DATOS DE RUBROS Y MEMORIAS DISPONIBLES ---\n" + infoRubros + "\n--- FIN DE LA BASE DE DATOS ---\n\n" +
      "--- CÓDIGO DE LA INTERFAZ DEL PORTAL ---\n" + codigoEstructura + "\n--- FIN DEL CÓDIGO ---\n\n" +
      "La pregunta del usuario es: " + preguntaUsuario;

    var urlIA = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
    var payload = {
      "contents": [{"parts": [{"text": contextoOculto}]}],
      "generationConfig": {"temperature": 0.1} 
    };

    var options = {
      'method' : 'post',
      'contentType': 'application/json',
      'payload' : JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    var response = UrlFetchApp.fetch(urlIA, options);
    var json = JSON.parse(response.getContentText());

    if (json.error) {
        var msgError = (json.error.message || "").toLowerCase();
        if (msgError.includes("high demand") || msgError.includes("exhausted") || msgError.includes("quota") || msgError.includes("overloaded")) {
            return "⚠️ Por alta demanda en los servidores no podemos contestar en este momento. Espere entre 30 y 60 segundos y vuelva a intentar.\n\nFuente: Interfaz del Portal";
        }
        return "⚠️ Error de conexión: " + json.error.message + "\n\nFuente: Interfaz del Portal";
    }

    if (json.candidates && json.candidates.length > 0 && json.candidates[0].content) {
        return json.candidates[0].content.parts[0].text;
    } else {
        return "⚠️ Sin respuesta de la IA. Intenta de nuevo.\n\nFuente: Interfaz del Portal";
    }
  } catch (e) { return "⚠️ Hubo un error de ejecución: " + e.message + "\n\nFuente: Interfaz del Portal"; }
}

function extraerTextoDeUrl(url) {
  try {
    var html = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText();
    var texto = html
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '') 
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')   
      .replace(/<[^>]+>/g, ' ')                        
      .replace(/\s+/g, ' ')                            
      .substring(0, 10000);                            
    return texto;
  } catch(e) {
    return "Error al leer la web: " + e.message;
  }
}