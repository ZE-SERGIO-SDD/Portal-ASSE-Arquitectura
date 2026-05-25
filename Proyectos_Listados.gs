// =====================================================================
// MÓDULO: LISTADO DE PROYECTOS (Lógica de servidor)
// =====================================================================

function obtenerListaDeProyectos(forzar) {
  var cache = CacheService.getScriptCache();
  
  // 1. INTENTO DE LECTURA DE CACHÉ FRAGMENTADO (Bypassea la limitación de 100KB)
  if (!forzar) {
    try {
      var chunksStr = cache.get("cache_lista_proy_chunks");
      if (chunksStr) {
        var chunks = parseInt(chunksStr, 10);
        var cachedStr = "";
        var cacheValido = true;
        
        for (var c = 0; c < chunks; c++) {
          var chunk = cache.get("cache_lista_proy_" + c);
          if (!chunk) { 
            cacheValido = false; 
            break; 
          }
          cachedStr += chunk;
        }
        
        if (cacheValido && cachedStr) {
          return JSON.parse(cachedStr);
        }
      }
    } catch(e) {
      // Falla silenciosa de caché, continúa leyendo la base de datos
    }
  }

  // 2. LECTURA DIRECTA DESDE GOOGLE SHEETS
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    var sheet = ss.getSheetByName("BD_Proyectos");
    if (!sheet) return [];
    
    // Optimizamos obteniendo solo el rango exacto de filas con datos reales
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; 
    var data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    
    var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    
    var idxFecha = headers.indexOf("fecha");
    var idxDepto = headers.indexOf("departamento");
    var idxCentro = headers.indexOf("centro");
    var idxNombre = headers.indexOf("nombre proyecto") !== -1 ? headers.indexOf("nombre proyecto") : headers.indexOf("nombre");
    var idxDesc = headers.indexOf("descripción") !== -1 ? headers.indexOf("descripción") : headers.indexOf("descripcion");
    var idxTipo = headers.indexOf("tipo de obra");
    var idxEstado = headers.indexOf("estado");
    var idxEquipo = headers.indexOf("equipo");
    var idxCreador = headers.indexOf("creado por");
    
    // COLUMNAS COMPLEMENTARIAS
    var idxExpediente = headers.indexOf("nro expediente");
    var idxCosto = headers.indexOf("costo obra");
    var idxEmpresa = headers.indexOf("empresa");
    var idxContactoNom = headers.indexOf("contacto nombre");
    var idxContactoTel = headers.indexOf("contacto teléfono") !== -1 ? headers.indexOf("contacto teléfono") : headers.indexOf("contacto telefono");
    var idxUbicArchivos = headers.indexOf("ubicación archivos") !== -1 ? headers.indexOf("ubicación archivos") : headers.indexOf("ubicacion archivos");
    var idxM2 = headers.indexOf("m2");

    var proyectos = [];
    
    // Procesamiento en bucle de alto rendimiento
    for (var i = 1; i < data.length; i++) {
      var fila = data[i];
      var nombre = idxNombre !== -1 ? fila[idxNombre] : fila[3];
      if (!nombre || nombre.toString().trim() === "") continue;
      
      var depto = idxDepto !== -1 ? fila[idxDepto] : (fila[1] || "");
      var centro = idxCentro !== -1 ? fila[idxCentro] : (fila[2] || "");
      var desc = idxDesc !== -1 ? fila[idxDesc] : (fila[4] || "");
      var tipo = idxTipo !== -1 ? fila[idxTipo] : (fila[5] || "");
      var estado = idxEstado !== -1 ? fila[idxEstado] : (fila[6] || "");
      var creador = idxCreador !== -1 ? fila[idxCreador] : (fila[8] || "");
      var fechaCruda = idxFecha !== -1 ? fila[idxFecha] : fila[0];
      
      var fechaSegura = "";
      if (fechaCruda instanceof Date) {
        fechaSegura = fechaCruda.toISOString();
      } else if (fechaCruda) {
        fechaSegura = fechaCruda.toString();
      }
      
      var equipoRaw = idxEquipo !== -1 ? fila[idxEquipo] : (fila[7] || "[]");
      var equipoObj = [];
      if (equipoRaw && equipoRaw.toString().trim() !== "") {
        try { equipoObj = JSON.parse(equipoRaw); } catch(e) {}
      }

      var expediente = idxExpediente !== -1 ? fila[idxExpediente] : "";
      var costo = idxCosto !== -1 ? fila[idxCosto] : "";
      var empresa = idxEmpresa !== -1 ? fila[idxEmpresa] : "";
      var contactoNom = idxContactoNom !== -1 ? fila[idxContactoNom] : "";
      var contactoTel = idxContactoTel !== -1 ? fila[idxContactoTel] : "";
      var ubicArchivos = idxUbicArchivos !== -1 ? fila[idxUbicArchivos] : "";
      var m2 = idxM2 !== -1 ? fila[idxM2] : "";

      proyectos.push({
        fecha: fechaSegura, departamento: depto.toString(), centro: centro.toString(),
        nombre: nombre.toString(), descripcion: desc.toString(), tipoObra: tipo.toString(),
        estado: estado.toString(), equipo: equipoObj, creador: creador.toString(),
        expediente: expediente.toString(), costo: costo.toString(), empresa: empresa.toString(),
        contactoNombre: contactoNom.toString(), contactoTelefono: contactoTel.toString(),
        ubicacionArchivos: ubicArchivos.toString(), m2: m2.toString()
      });
    }
    
    var result = proyectos.reverse();
    
    // 3. ESCRITURA EN CACHÉ FRAGMENTADO (Divide el archivo si supera el límite de la plataforma)
    try {
      var resultStr = JSON.stringify(result);
      var chunkSize = 90000; // Bloques seguros de ~90KB
      var chunks = Math.ceil(resultStr.length / chunkSize);
      
      for (var i = 0; i < chunks; i++) {
        cache.put("cache_lista_proy_" + i, resultStr.slice(i * chunkSize, (i + 1) * chunkSize), 21600);
      }
      cache.put("cache_lista_proy_chunks", chunks.toString(), 21600);
    } catch(e) {
      // Evita bloqueos si la escritura fallara por otra causa
    }
    
    return result;
  } catch (e) {
    return { error: "Error al leer BD: " + e.message };
  }
}