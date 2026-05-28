// =====================================================================
// BACKEND: GESTIÓN DE PAPELERA DE PROYECTOS (Solo Admins)
// =====================================================================

/**
 * Obtiene todos los proyectos que se encuentran en BD_Proyectos_Eliminados.
 */
function obtenerProyectosEliminados() {
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID); // Utiliza tu ID de configuración
    var sheet = ss.getSheetByName("BD_Proyectos_Eliminados");
    
    if (!sheet) {
      return { success: false, error: "No se encontró la pestaña 'BD_Proyectos_Eliminados'." };
    }

    var data = sheet.getDataRange().getValues();
    var eliminados = [];

    // Iteramos desde la fila 1 (saltando los encabezados)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Verificación básica para no procesar filas vacías
      if (!row[3]) continue; 

      var fechaP = (row[0] instanceof Date) ? row[0].toISOString() : row[0].toString();
      var fechaElim = (row[19] instanceof Date) ? row[19].toISOString().split('T')[0] : row[19].toString();

      // Convertimos el JSON de equipo (Col H - índice 7) si existe
      var equipoData = [];
      try {
        if (row[7]) equipoData = JSON.parse(row[7]);
      } catch (e) {
        equipoData = [];
      }

      eliminados.push({
        fecha: fechaP,                                // A: 0
        departamento: row[1] ? row[1].toString() : "",// B: 1
        centro: row[2] ? row[2].toString() : "",      // C: 2
        nombre: row[3] ? row[3].toString() : "",      // D: 3
        descripcion: row[4] ? row[4].toString() : "", // E: 4
        tipoObra: row[5] ? row[5].toString() : "",    // F: 5
        estado: row[6] ? row[6].toString() : "",      // G: 6
        equipo: equipoData,                           // H: 7
        creador: row[8] ? row[8].toString() : "",     // I: 8
        // Nuevas columnas de la papelera:
        origen: row[18] ? row[18].toString() : "En Curso", // S: 18
        fechaEliminacion: fechaElim,                       // T: 19
        eliminadoPor: row[20] ? row[20].toString() : ""    // U: 20
      });
    }

    return { success: true, datos: eliminados };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Restaura un proyecto desde la Papelera hacia su hoja de Origen ("En Curso" o "Archivado")
 * y también restaura su bitácora.
 */
function restaurarProyectoBackend(nombre, fecha) {
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    var sheetEliminados = ss.getSheetByName("BD_Proyectos_Eliminados");
    
    if (!sheetEliminados) return { success: false, error: "No existe la pestaña BD_Proyectos_Eliminados." };

    var dataEliminados = sheetEliminados.getDataRange().getValues();
    var filaEncontrada = -1;
    var rowData = null;
    var origenDestino = "";

    // 1. Encontrar el proyecto en la Papelera
    for (var i = 1; i < dataEliminados.length; i++) {
      var rowFecha = (dataEliminados[i][0] instanceof Date) ? dataEliminados[i][0].toISOString() : dataEliminados[i][0].toString();
      var rowNombre = dataEliminados[i][3] ? dataEliminados[i][3].toString() : "";
      
      if (rowNombre === nombre && rowFecha === fecha) {
        filaEncontrada = i + 1;
        rowData = dataEliminados[i];
        origenDestino = dataEliminados[i][18] ? dataEliminados[i][18].toString() : "En Curso"; // Columna S
        break;
      }
    }

    if (filaEncontrada === -1) {
      return { success: false, error: "No se encontró el proyecto en la papelera." };
    }

    // 2. Determinar a qué pestaña debe volver
    var nombrePestañaDestino = (origenDestino === "Archivado") ? "BD_Proyectos_Archivados" : "BD_Proyectos";
    var sheetDestino = ss.getSheetByName(nombrePestañaDestino);
    
    if (!sheetDestino) {
      // Fallback: si por algo no existe la de archivados, lo mandamos a En Curso
      sheetDestino = ss.getSheetByName("BD_Proyectos");
    }

    // 3. Quitar las últimas 3 columnas (Origen, Fecha Elim, Usuario) para dejarlo en su estado original
    var filaOriginalRestaurada = rowData.slice(0, rowData.length - 3);

    // 4. Escribir en la pestaña original
    sheetDestino.appendRow(filaOriginalRestaurada);

    // 5. Eliminar de la papelera
    sheetEliminados.deleteRow(filaEncontrada);

    // 6. Restaurar el historial (Bitácora) si existe
    restaurarHistorialEliminado(ss, nombre, fecha);

    // 7. Limpiar Caché para forzar la recarga
    var cache = CacheService.getScriptCache();
    cache.remove("cache_lista_proy_chunks");
    cache.remove("cache_lista_arch_chunks");

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// Función auxiliar para restaurar la bitácora
function restaurarHistorialEliminado(ss, nombre, fecha) {
  var sheetHistorial = ss.getSheetByName("BD_Historial");
  var sheetHistEliminados = ss.getSheetByName("BD_Historial_Eliminados");
  
  if (!sheetHistorial || !sheetHistEliminados) return;
  
  var dataHistElim = sheetHistEliminados.getDataRange().getValues();
  var filasAEliminar = [];

  for (var i = dataHistElim.length - 1; i >= 1; i--) {
    var histNom = dataHistElim[i][0] ? dataHistElim[i][0].toString() : "";
    var histFec = (dataHistElim[i][1] instanceof Date) ? dataHistElim[i][1].toISOString() : dataHistElim[i][1].toString();
    
    if (histNom === nombre && histFec === fecha) {
      // Quitamos la última columna (origen) que le habíamos agregado al borrarlo
      var filaRestaurada = dataHistElim[i].slice(0, dataHistElim[i].length - 1);
      sheetHistorial.appendRow(filaRestaurada);
      filasAEliminar.push(i + 1);
    }
  }

  // Borramos de la papelera de historial
  for (var k = 0; k < filasAEliminar.length; k++) {
    sheetHistEliminados.deleteRow(filasAEliminar[k]);
  }
}