// =====================================================================
// BACKEND: GESTIÓN DE PAPELERA DE PROYECTOS (Solo Admins)
// =====================================================================

function obtenerProyectosEliminados() {
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    var sheet = ss.getSheetByName("BD_Proyectos_Eliminados");
    if (!sheet) {
      return { success: false, error: "No se encontró la pestaña 'BD_Proyectos_Eliminados'." };
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, datos: [] }; // Está vacía

    var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    var idxFecha = headers.indexOf("fecha");
    var idxDepto = headers.indexOf("departamento");
    var idxCentro = headers.indexOf("centro");
    var idxNombre = headers.indexOf("nombre proyecto") !== -1 ? headers.indexOf("nombre proyecto") : headers.indexOf("nombre");
    var idxOrigen = headers.indexOf("origen");
    
    // Identificamos las columnas finales
    var idxFechaElim = headers.indexOf("fecha eliminacion") !== -1 ? headers.indexOf("fecha eliminacion") : (data[0].length - 2); 
    var idxElimPor = headers.indexOf("eliminado por") !== -1 ? headers.indexOf("eliminado por") : (data[0].length - 1);
    
    var eliminados = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (idxNombre === -1 || !row[idxNombre]) continue; // Fila vacía
      
      var fechaP = (idxFecha !== -1 && row[idxFecha] instanceof Date) ? row[idxFecha].toISOString() : (row[idxFecha] || "");
      var fechaElim = (row[idxFechaElim] instanceof Date) ? row[idxFechaElim].toISOString().split('T')[0] : (row[idxFechaElim] || "");
      var fOrigen = idxOrigen !== -1 && row[idxOrigen] ? row[idxOrigen].toString() : "En Curso";
      var fDepto = idxDepto !== -1 && row[idxDepto] ? row[idxDepto].toString() : "";
      var fCentro = idxCentro !== -1 && row[idxCentro] ? row[idxCentro].toString() : "";
      var fElimPor = idxElimPor !== -1 && row[idxElimPor] ? row[idxElimPor].toString() : "";
      
      eliminados.push({
        fecha: fechaP,
        departamento: fDepto,
        centro: fCentro,
        nombre: row[idxNombre].toString(),  
        origen: fOrigen,
        fechaEliminacion: fechaElim,
        eliminadoPor: fElimPor
      });
    }

    return { success: true, datos: eliminados };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// MODIFICACIÓN: Ahora recibe 'usuarioNombre' como tercer parámetro
function restaurarProyectoBackend(nombre, fecha, usuarioNombre) {
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    var sheetEliminados = ss.getSheetByName("BD_Proyectos_Eliminados");
    if (!sheetEliminados) return { success: false, error: "No existe la pestaña BD_Proyectos_Eliminados." };

    var dataEliminados = sheetEliminados.getDataRange().getValues();
    var headers = dataEliminados[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    var idxFecha = headers.indexOf("fecha");
    var idxNombre = headers.indexOf("nombre proyecto") !== -1 ? headers.indexOf("nombre proyecto") : headers.indexOf("nombre");
    var idxOrigen = headers.indexOf("origen");
    var idxId = headers.indexOf("id proyecto");

    var filaEncontrada = -1;
    var rowData = null;
    var origenDestino = "";
    var idProyecto = "";

    // Buscar en papelera asegurando lectura dinámica
    for (var i = 1; i < dataEliminados.length; i++) {
      var rowFecha = (idxFecha !== -1 && dataEliminados[i][idxFecha] instanceof Date) ? dataEliminados[i][idxFecha].toISOString() : (dataEliminados[i][idxFecha] || "").toString();
      var rowNombre = (idxNombre !== -1 && dataEliminados[i][idxNombre]) ? dataEliminados[i][idxNombre].toString() : "";
      
      if (rowNombre === nombre && rowFecha === fecha) {
        filaEncontrada = i + 1;
        rowData = dataEliminados[i];
        origenDestino = idxOrigen !== -1 && dataEliminados[i][idxOrigen] ? dataEliminados[i][idxOrigen].toString() : "En Curso";
        idProyecto = idxId !== -1 && dataEliminados[i][idxId] ? dataEliminados[i][idxId].toString() : "";
        break;
      }
    }

    if (filaEncontrada === -1) {
      return { success: false, error: "No se encontró el proyecto en la papelera." };
    }

    var nombrePestañaDestino = (origenDestino === "Archivado") ? "BD_Proyectos_Archivados" : "BD_Proyectos";
    var sheetDestino = ss.getSheetByName(nombrePestañaDestino);
    if (!sheetDestino) sheetDestino = ss.getSheetByName("BD_Proyectos");

    // Quitar las últimas 3 columnas (Origen, Fecha Elim, Usuario) para dejarlo limpio en su BD original
    var filaOriginalRestaurada = rowData.slice(0, rowData.length - 3);
    sheetDestino.appendRow(filaOriginalRestaurada);
    sheetEliminados.deleteRow(filaEncontrada);

    // Restaurar historial usando el ID exacto
    if (idProyecto) {
       restaurarHistorialEliminado(ss, idProyecto, origenDestino);
       
       // NUEVO: Registrar el hito de restauración con FECHA/HORA exacta y USUARIO real
       var fechaRestauracion = new Date();
       
       // Aprovechamos resolverNombreUsuario de Proyectos.gs para formatear perfecto Nombre y Apellido
       var user = typeof resolverNombreUsuario === "function" ? resolverNombreUsuario(usuarioNombre) : (usuarioNombre || "Sistema");
       
       var nombrePestañaHistorial = (origenDestino === "Archivado") ? "BD_Historial_Archivados" : "BD_Historial";
       var sheetHistorialDestino = ss.getSheetByName(nombrePestañaHistorial);
       
       if (sheetHistorialDestino) {
         sheetHistorialDestino.appendRow([idProyecto, fechaRestauracion, user, "RESTAURADO", "El proyecto fue recuperado desde la papelera.", ""]);
       }
    }

    var cache = CacheService.getScriptCache();
    cache.remove("cache_lista_proy_chunks");
    cache.remove("cache_lista_arch_chunks");

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function restaurarHistorialEliminado(ss, idProyecto, origenDestino) {
  var nombrePestañaDestino = (origenDestino === "Archivado") ? "BD_Historial_Archivados" : "BD_Historial";
  var sheetHistorial = ss.getSheetByName(nombrePestañaDestino);
  var sheetHistEliminados = ss.getSheetByName("BD_Historial_Eliminados");
  
  if (!sheetHistorial || !sheetHistEliminados) return;
  
  var dataHistElim = sheetHistEliminados.getDataRange().getValues();
  var filasAEliminar = [];
  // Recorremos de atrás hacia adelante para no alterar índices al borrar luego
  for (var i = dataHistElim.length - 1; i >= 1; i--) {
    var histId = dataHistElim[i][0] ? dataHistElim[i][0].toString() : "";
    if (histId === idProyecto) {
      // Quitamos la última columna que era "Origen" agregada al eliminar
      var filaRestaurada = dataHistElim[i].slice(0, dataHistElim[i].length - 1);
      sheetHistorial.appendRow(filaRestaurada);
      filasAEliminar.push(i + 1);
    }
  }

  for (var k = 0; k < filasAEliminar.length; k++) {
    sheetHistEliminados.deleteRow(filasAEliminar[k]);
  }
}