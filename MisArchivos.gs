// =====================================================================
// MÓDULO: MIS ARCHIVOS E HISTORIAL (Lógica de servidor)
// =====================================================================

function registrarArchivoGenerado(usuarioEmail, nombre, id) {
  try {
    var ss = SpreadsheetApp.openById(BORRADORES_SHEET_ID);
    var sheet = ss.getSheetByName("LOG_ARCHIVOS");
    if (!sheet) {
        sheet = ss.insertSheet("LOG_ARCHIVOS");
        sheet.appendRow(["Email", "Nombre", "ID", "Fecha", "Estado"]);
    }
    
    // Guardamos la fecha ya formateada para lectura rápida después
    var fechaActual = Utilities.formatDate(new Date(), "GMT-3", "dd/MM HH:mm");
    sheet.appendRow([usuarioEmail.toLowerCase(), nombre, id, fechaActual, "Activo"]);
    
    return true;
  } catch (e) { 
    console.log("Error al registrar: " + e.message); 
    return false;
  }
}

// NUEVA FUNCIÓN ULTRA RÁPIDA: Lee la planilla sin preguntarle a Drive
function obtenerMisArchivosPaginados(usuarioEmail, limite, desde) {
  try {
    var ss = SpreadsheetApp.openById(BORRADORES_SHEET_ID);
    var sheet = ss.getSheetByName("LOG_ARCHIVOS");
    if (!sheet) return { archivos: [], tieneMas: false, esInicial: (desde === 0) };
    
    var data = sheet.getDataRange().getValues();
    var misArchivos = [];
    var emailBuscado = usuarioEmail.toString().toLowerCase();
    
    // Recorremos de abajo hacia arriba (más recientes primero)
    for (var i = data.length - 1; i >= 1; i--) { 
      var filaUserEmail = data[i][0] ? data[i][0].toString().toLowerCase() : "";
      
      if (filaUserEmail === emailBuscado) {
        var fileId = data[i][2];
        var rawFecha = data[i][3];
        var fechaFormateada = rawFecha;

        // Por compatibilidad con archivos viejos que guardaste como objeto Date
        if (rawFecha instanceof Date) {
           fechaFormateada = Utilities.formatDate(rawFecha, "GMT-3", "dd/MM HH:mm");
        }

        misArchivos.push({
          id: fileId,
          nombre: data[i][1],
          url: 'https://docs.google.com/document/d/' + fileId + '/edit', // URL autogenerada para ahorrar lecturas
          fecha: fechaFormateada
        });
      }
    }
    
    // Paginación: recortamos solo la porción que la pantalla necesita ver ahora
    var segmento = misArchivos.slice(desde, desde + limite);
    var tieneMas = misArchivos.length > (desde + limite);
    
    return {
      archivos: segmento,
      tieneMas: tieneMas,
      esInicial: (desde === 0)
    };
  } catch (e) { 
    return { error: e.message }; 
  }
}

function borrarArchivoGenerado(fileId) {
  try {
    // 1. Borramos de Drive (intentamos en silencio por si ya lo había borrado el usuario a mano)
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}
    
    // 2. Borramos de la base de datos
    var ss = SpreadsheetApp.openById(BORRADORES_SHEET_ID);
    var sheet = ss.getSheetByName("LOG_ARCHIVOS");
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = data.length - 1; i >= 0; i--) {
        if (data[i][2] && data[i][2].toString() === fileId.toString()) {
          sheet.deleteRow(i + 1); 
          break; 
        }
      }
    }
    return { success: true };
  } catch (e) { 
    return { error: e.message }; 
  }
}

function borrarTodoElHistorial(usuarioEmail) {
  try {
    var ss = SpreadsheetApp.openById(BORRADORES_SHEET_ID);
    var sheet = ss.getSheetByName("LOG_ARCHIVOS");
    if (!sheet) return { success: false, error: "No se encontró la hoja de registro." };

    var data = sheet.getDataRange().getValues();
    var emailBuscado = usuarioEmail.toString().toLowerCase();

    // Borramos desde abajo hacia arriba para que no se corran los índices al eliminar filas
    for (var i = data.length - 1; i >= 1; i--) {
      var filaUserEmail = data[i][0] ? data[i][0].toString().toLowerCase() : "";
      if (filaUserEmail === emailBuscado) {
        var fileId = data[i][2];
        try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
        sheet.deleteRow(i + 1);
      }
    }

    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}