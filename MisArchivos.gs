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

// NUEVA FUNCIÓN ULTRA RÁPIDA: Lee y filtra los de hace 3 días antes de paginar
function obtenerMisArchivosPaginados(usuarioEmail, limite, desde) {
  try {
    var ss = SpreadsheetApp.openById(BORRADORES_SHEET_ID);
    var sheet = ss.getSheetByName("LOG_ARCHIVOS");
    if (!sheet) return { archivos: [], tieneMas: false, esInicial: (desde === 0) };
    
    var data = sheet.getDataRange().getValues();
    var misArchivos = [];
    var emailBuscado = usuarioEmail.toString().toLowerCase();
    
    var hoy = new Date();
    hoy.setHours(0,0,0,0); // Normalizamos a medianoche para calcular días precisos

    // Recorremos de abajo hacia arriba (más recientes primero)
    for (var i = data.length - 1; i >= 1; i--) { 
      var filaUserEmail = data[i][0] ? data[i][0].toString().toLowerCase() : "";
      
      if (filaUserEmail === emailBuscado) {
        var fileId = data[i][2];
        var rawFecha = data[i][3];
        var fechaFormateada = rawFecha;
        var docDate = new Date(); 
        var isValidDate = false;

        // 1. EXTRACTOR DE FECHAS (Soporta formato "27/05 15:23" sin año)
        if (rawFecha instanceof Date) {
           docDate = rawFecha;
           isValidDate = true;
           fechaFormateada = Utilities.formatDate(rawFecha, "GMT-3", "dd/MM HH:mm");
        } else if (typeof rawFecha === 'string') {
           fechaFormateada = rawFecha;
           var match = rawFecha.match(/(\d{1,2})\/(\d{1,2})/); // Extrae estrictamente DD/MM
           if (match) {
             var dia = parseInt(match[1], 10);
             var mes = parseInt(match[2], 10) - 1;
             var anio = hoy.getFullYear();
             
             // Ajuste por si el archivo es de diciembre y hoy es enero
             if (mes === 11 && hoy.getMonth() === 0) anio--; 
             
             docDate.setFullYear(anio, mes, dia);
             isValidDate = true;
           }
        }

        // 2. FILTRO DE 3 DÍAS EN EL SERVIDOR
        var mostrarArchivo = true;
        if (isValidDate) {
           var checkDate = new Date(docDate);
           checkDate.setHours(0,0,0,0);
           var diffDias = (hoy.getTime() - checkDate.getTime()) / (1000 * 60 * 60 * 24);
           
           if (diffDias > 3) {
             mostrarArchivo = false; // El archivo es viejo, lo ignoramos
           }
        }

        // 3. SOLO GUARDAMOS SI PASÓ EL FILTRO
        if (mostrarArchivo) {
            misArchivos.push({
              id: fileId,
              nombre: data[i][1],
              url: 'https://docs.google.com/document/d/' + fileId + '/edit',
              fecha: fechaFormateada
            });
        }
      }
    }
    
    // 4. AHORA SÍ PAGINAMOS (Siempre enviará grupos exactos de 5)
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
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}
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