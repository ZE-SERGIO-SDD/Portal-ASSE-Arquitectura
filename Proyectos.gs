// ID de la planilla de configuración de Proyectos
var PROYECTOS_CONFIG_ID = '1NCoOyOz7r14MFF37Cq5NEU8cmd_-XyOewvVe4iE1ZrY';
// ID de la planilla de Departamentos y Centros estructurada en columnas
var DEPARTAMENTOS_CENTROS_ID = '1dq81v5ccrsZhRRJH6VHSXJv8wPpVFRWUHIyX7KNFHgw';
// ID de la planilla externa de origen para los Participantes
var PARTICIPANTES_EXTERNO_ID = '18MRDKeWsbaciGSTrP63bSM-z_l9EvdjKIQZTJ5BjYNo';
// ID de la planilla externa de origen para las Empresas
var EMPRESAS_SHEET_ID = '11-kpuGqKCyOUFvUYEaTCMUec_5OQppTC8VFzzFyA-Io';

function obtenerParametrosProyectos(forzar) {
  var cache = CacheService.getScriptCache();
  var cacheKeyPrefix = "cache_params_proy_chunk_";
  
  if (!forzar) {
    try {
      var chunksStr = cache.get("cache_params_chunks_count");
      if (chunksStr) {
        var chunks = parseInt(chunksStr, 10);
        var cachedStr = "";
        var cacheValido = true;
        for (var c = 0; c < chunks; c++) {
          var chunk = cache.get(cacheKeyPrefix + c);
          if (!chunk) { cacheValido = false; break; }
          cachedStr += chunk;
        }
        if (cacheValido && cachedStr) {
          return JSON.parse(cachedStr);
        }
      }
    } catch(e) {}
  }

  try {
    var ssConfig = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    
    var leerColumna = function(nombrePestaña) {
      var sheet = ssConfig.getSheetByName(nombrePestaña);
      if (!sheet) return [];
      var values = sheet.getDataRange().getValues();
      var lista = [];
      for (var i = 0; i < values.length; i++) {
        if (values[i][0]) lista.push(values[i][0].toString().trim());
      }
      return [...new Set(lista)];
    };

    var listaParticipantes = [];
    try {
      var ssPart = SpreadsheetApp.openById(PARTICIPANTES_EXTERNO_ID);
      var sheetPart = ssPart.getSheetByName("Hoja 1");
      if (sheetPart) {
        var dataPart = sheetPart.getDataRange().getValues();
        for (var i = 1; i < dataPart.length; i++) {
          var nombre = dataPart[i][2];
          var apellido = dataPart[i][3];  
          if (nombre && nombre.toString().trim() !== "") {
            var nombreCompleto = nombre.toString().trim() + " " + (apellido ? apellido.toString().trim() : "");
            listaParticipantes.push(nombreCompleto.trim());
          }
        }
      }
    } catch (e) { console.error("Error partic: " + e.message); }
    listaParticipantes = [...new Set(listaParticipantes)].sort();

    var listaEmpresas = [];
    try {
      var ssEmp = SpreadsheetApp.openById(EMPRESAS_SHEET_ID);
      var sheetEmp = ssEmp.getSheets()[0];
      var dataEmp = sheetEmp.getDataRange().getValues();
      if (dataEmp.length === 0 || (dataEmp.length === 1 && dataEmp[0][0].toString().trim() === "")) {
        sheetEmp.appendRow(["Empresas Registradas (Base Central)"]);
        sheetEmp.appendRow(["Teyma Uruguay S.A."]);
        sheetEmp.appendRow(["Saceem S.A."]);
        sheetEmp.appendRow(["Stiler S.A."]);
        listaEmpresas = ["Teyma Uruguay S.A.", "Saceem S.A.", "Stiler S.A."];
      } else {
        for (var i = 0; i < dataEmp.length; i++) {
          var empVal = dataEmp[i][0] ? dataEmp[i][0].toString().trim() : "";
          if (empVal && !empVal.toLowerCase().includes("empresas registradas")) {
            listaEmpresas.push(empVal);
          }
        }
      }
    } catch (e) {
      console.error("Error empresas: " + e.message);
      listaEmpresas = ["Teyma Uruguay S.A.", "Saceem S.A.", "Stiler S.A."]; 
    }
    listaEmpresas = [...new Set(listaEmpresas)].sort();

    var sheetTipos = ssConfig.getSheetByName('Tipo de Obra');
    var sheetEstados = ssConfig.getSheetByName('Estados');
    var listaTiposObra = [];
    var mapaTiposEstados = {};
    var todosLosEstados = [];

    if (sheetTipos && sheetEstados) {
      var dataTipos = sheetTipos.getDataRange().getValues();
      var dataEstados = sheetEstados.getDataRange().getValues();
      var maxRows = Math.max(dataTipos.length, dataEstados.length);

      for (var i = 0; i < maxRows; i++) {
        var tipo = (dataTipos[i] && dataTipos[i][0]) ? dataTipos[i][0].toString().trim() : "";
        if (tipo) {
          listaTiposObra.push(tipo);
          var estadosRow = [];
          if (dataEstados[i]) {
            for (var c = 0; c < dataEstados[i].length; c++) {
              var val = dataEstados[i][c] ? dataEstados[i][c].toString().trim() : "";
              if (val) {
                if (val.includes(",")) {
                  val.split(",").forEach(function(p) {
                    var clean = p.trim();
                    if (clean) estadosRow.push(clean);
                  });
                } else {
                  estadosRow.push(val);
                }
              }
            }
          }
          mapaTiposEstados[tipo] = [...new Set(estadosRow)];
          todosLosEstados = todosLosEstados.concat(estadosRow);
        }
      }
    }
    listaTiposObra = [...new Set(listaTiposObra)];
    todosLosEstados = [...new Set(todosLosEstados)].sort();

    var listaDeptos = [];
    var mapaCentros = {}; 
    var infoCentros = {};
    
    var procesarFilaCentro = function(row) {
      var depto = (row.length > 6 && row[6]) ? row[6].toString().trim() : "";
      var centro = (row.length > 3 && row[3]) ? row[3].toString().trim() : "";
      
      if (depto !== "") {
        if (!mapaCentros[depto]) {
          mapaCentros[depto] = [];
          listaDeptos.push(depto);
        }
        if (centro !== "" && mapaCentros[depto].indexOf(centro) === -1) {
          mapaCentros[depto].push(centro);
          infoCentros[centro] = {
            calle: (row.length > 19 && row[19]) ? row[19].toString().trim() : "",
            puerta: (row.length > 20 && row[20]) ? row[20].toString().trim() : "",
            dependencia: (row.length > 25 && row[25]) ? row[25].toString().trim() : "",
            ueNum: (row.length > 26 && row[26]) ? row[26].toString().trim() : "",
            ueNom: (row.length > 27 && row[27]) ? row[27].toString().trim() : "",
            nivel: (row.length > 28 && row[28]) ? row[28].toString().trim() : "",
            categoria: (row.length > 29 && row[29]) ? row[29].toString().trim() : ""
          };
        }
      }
    };
    
    try {
      var resSheets = Sheets.Spreadsheets.Values.get(DEPARTAMENTOS_CENTROS_ID, 'A:AD');
      var rows = resSheets.values;
      if (rows && rows.length > 3) {
        for (var i = 3; i < rows.length; i++) { procesarFilaCentro(rows[i]); }
      }
    } catch (err) {
      var sheetDepto = SpreadsheetApp.openById(DEPARTAMENTOS_CENTROS_ID).getSheets()[0];
      var dataDepto = sheetDepto.getRange("A:AD").getValues();
      for (var i = 3; i < dataDepto.length; i++) { procesarFilaCentro(dataDepto[i]); }
    }

    listaDeptos = [...new Set(listaDeptos)].sort();
    for (var d in mapaCentros) { mapaCentros[d].sort(); }
    
    var result = {
      participantes: listaParticipantes,
      roles: leerColumna('Roles'),
      estados: todosLosEstados,
      tiposObra: listaTiposObra,
      mapaTiposEstados: mapaTiposEstados, 
      departamentos: listaDeptos,
      mapaCentros: mapaCentros,
      infoCentros: infoCentros,
      empresas: listaEmpresas 
    };
    
    try {
      var resultStr = JSON.stringify(result);
      var chunkSize = 90000;
      var chunks = Math.ceil(resultStr.length / chunkSize);
      for (var i = 0; i < chunks; i++) {
        cache.put(cacheKeyPrefix + i, resultStr.slice(i * chunkSize, (i + 1) * chunkSize), 21600);
      }
      cache.put("cache_params_chunks_count", chunks.toString(), 21600);
    } catch(e) {}
    
    return result;
    
  } catch (e) {
    return { error: "Error al cargar parámetros: " + e.message };
  }
}

function guardarNuevoProyecto(datosProyecto) {
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    var sheet = ss.getSheetByName("BD_Proyectos") || ss.insertSheet("BD_Proyectos");
    
    var columnasRequeridas = [
      "ID Proyecto", "Fecha", "Departamento", "Centro", "Nombre Proyecto", "Descripción", 
      "Tipo de Obra", "Estado", "Equipo", "Creado por",
      "Nro Expediente", "Costo Obra", "Empresa", "Contacto Nombre", "Contacto Teléfono", "Ubicación Archivos", "Fechas Hitos", "M2"
    ];
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(columnasRequeridas);
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    
    var columnasFaltantes = [
      "ID Proyecto", "Nro Expediente", "Costo Obra", "Empresa", "Contacto Nombre", "Contacto Teléfono", "Ubicación Archivos", "Fechas Hitos", "M2"
    ];
    
    columnasFaltantes.forEach(function(col) {
      var colBaja = col.toLowerCase().trim();
      var colSinTilde = colBaja.replace(/[áéíóú]/g, function(match){
         if(match==='á')return 'a';if(match==='é')return 'e';if(match==='í')return 'i';if(match==='ó')return 'o';if(match==='ú')return 'u';
      });
      if (headers.indexOf(colBaja) === -1 && headers.indexOf(colSinTilde) === -1) {
        sheet.getRange(1, headers.length + 1).setValue(col);
        headers.push(colBaja);
      }
    });
    
    var idUnico = "PRY-" + Date.now().toString(36).toUpperCase();
    
    var mapaDatos = {
      "id proyecto": idUnico,
      "fecha": new Date(),
      "departamento": datosProyecto.departamento,
      "centro": datosProyecto.centro,
      "nombre proyecto": datosProyecto.nombre,
      "descripción": datosProyecto.descripcion,
      "descripcion": datosProyecto.descripcion,
      "tipo de obra": datosProyecto.tipoObra,
      "estado": datosProyecto.estado,
      "equipo": JSON.stringify(datosProyecto.asignados),
      "creado por": datosProyecto.usuario, // Corresponde al usuarioNombre ya enviado desde el front en la creación
      "nro expediente": datosProyecto.expediente || "",
      "costo obra": datosProyecto.costo || "",
      "empresa": datosProyecto.empresa || "",
      "contacto nombre": datosProyecto.contactoNombre || "",
      "contacto teléfono": datosProyecto.contactoTelefono || "",
      "contacto telefono": datosProyecto.contactoTelefono || "",
      "ubicación archivos": datosProyecto.ubicacionArchivos || "",
      "ubicacion archivos": datosProyecto.ubicacionArchivos || "",
      "fechas hitos": JSON.stringify(datosProyecto.hitos || {}),
      "m2": datosProyecto.m2 || ""
    };
    
    var nuevaFila = new Array(headers.length);
    for (var j = 0; j < headers.length; j++) {
      var hName = headers[j];
      nuevaFila[j] = mapaDatos.hasOwnProperty(hName) ? mapaDatos[hName] : "";
    }
    
    sheet.appendRow(nuevaFila);
    
    registrarHitoHistorial(idUnico, new Date(), datosProyecto.usuario, "Creación", "Alta inicial en estado: " + datosProyecto.estado, "Proyecto creado exitosamente.");

    CacheService.getScriptCache().remove("cache_lista_proy");
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function cambiarEstadoProyecto(nombre, fecha, nuevoEstado, usuarioNombre, comentario) {
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    var sheet = ss.getSheetByName("BD_Proyectos");
    
    if (!sheet) return { success: false, error: "No se encontró la hoja de BD_Proyectos" };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    
    var idxNombre = headers.indexOf("nombre proyecto") !== -1 ? headers.indexOf("nombre proyecto") : headers.indexOf("nombre");
    var idxFecha = headers.indexOf("fecha");
    var idxEstado = headers.indexOf("estado");
    var idxId = headers.indexOf("id proyecto");
    
    for (var i = 1; i < data.length; i++) {
      var filaNombre = data[i][idxNombre] ? data[i][idxNombre].toString() : "";
      var filaFechaRaw = data[i][idxFecha];
      var filaFecha = filaFechaRaw instanceof Date ? filaFechaRaw.toISOString() : (filaFechaRaw ? filaFechaRaw.toString() : "");
      
      if (filaNombre === nombre && filaFecha === fecha) {
        var estadoAnterior = data[i][idxEstado] ? data[i][idxEstado].toString() : "Sin estado previo";
        var idProyecto = idxId !== -1 ? (data[i][idxId] ? data[i][idxId].toString() : "") : "";
        
        if (!idProyecto) {
          idProyecto = "PRY-MIG-" + Date.now().toString(36).toUpperCase();
          if (idxId !== -1) {
            sheet.getRange(i + 1, idxId + 1).setValue(idProyecto);
          }
        }
        
        sheet.getRange(i + 1, idxEstado + 1).setValue(nuevoEstado);
        
        var detalleCambio = "Pasó de '" + estadoAnterior + "' a '" + nuevoEstado + "'";
        registrarHitoHistorial(idProyecto, new Date(), usuarioNombre || "Usuario Desconocido", "ESTADO", detalleCambio, comentario || "");

        CacheService.getScriptCache().remove("cache_lista_proy");
        return { success: true };
      }
    }
    return { success: false, error: "No se encontró el proyecto para actualizar." };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function editarDatosProyecto(datosEdicion) {
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    var sheet = ss.getSheetByName("BD_Proyectos");
    
    if (!sheet) return { success: false, error: "No se encontró la hoja BD_Proyectos" };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    
    var columnasFaltantes = [
      "ID Proyecto", "Nro Expediente", "Costo Obra", "Empresa", "Contacto Nombre", "Contacto Teléfono", "Ubicación Archivos", "Fechas Hitos", "M2"
    ];
    
    columnasFaltantes.forEach(function(col) {
      var colBaja = col.toLowerCase().trim();
      var colSinTilde = colBaja.replace(/[áéíóú]/g, function(match){
         if(match==='á')return 'a';if(match==='é')return 'e';if(match==='í')return 'i';if(match==='ó')return 'o';if(match==='ú')return 'u';
      });
      if (headers.indexOf(colBaja) === -1 && headers.indexOf(colSinTilde) === -1) {
        sheet.getRange(1, headers.length + 1).setValue(col);
        headers.push(colBaja);
      }
    });
    
    data = sheet.getDataRange().getValues();

    var idxNombre = headers.indexOf("nombre proyecto") !== -1 ? headers.indexOf("nombre proyecto") : headers.indexOf("nombre");
    var idxFecha = headers.indexOf("fecha");
    var idxId = headers.indexOf("id proyecto");
    
    var filaEncontrada = -1;
    var idProyecto = "";
    
    for (var i = 1; i < data.length; i++) {
      var filaNombre = data[i][idxNombre] ? data[i][idxNombre].toString() : "";
      var filaFechaRaw = data[i][idxFecha];
      var filaFecha = filaFechaRaw instanceof Date ? filaFechaRaw.toISOString() : (filaFechaRaw ? filaFechaRaw.toString() : "");
      
      if (filaNombre === datosEdicion.nombreOriginal && filaFecha === datosEdicion.fechaOriginal) {
        filaEncontrada = i;
        idProyecto = idxId !== -1 && data[i][idxId] ? data[i][idxId].toString() : "";
        break;
      }
    }

    if (filaEncontrada === -1) return { success: false, error: "No se encontró el proyecto original para editar." };

    if (!idProyecto) {
      idProyecto = "PRY-MIG-" + Date.now().toString(36).toUpperCase();
      if (idxId !== -1) sheet.getRange(filaEncontrada + 1, idxId + 1).setValue(idProyecto);
    }

    var columnasAEditar = {
      "descripción": datosEdicion.descripcion,
      "descripcion": datosEdicion.descripcion,
      "nro expediente": datosEdicion.expediente,
      "costo obra": datosEdicion.costo,
      "empresa": datosEdicion.empresa,
      "contacto nombre": datosEdicion.contactoNombre,
      "contacto teléfono": datosEdicion.contactoTelefono,
      "contacto telefono": datosEdicion.contactoTelefono,
      "ubicación archivos": datosEdicion.ubicacionArchivos,
      "ubicacion archivos": datosEdicion.ubicacionArchivos,
      "m2": datosEdicion.m2
    };
    
    var usuario = datosEdicion.usuarioNombre || "Usuario Desconocido";
    var comentariosFront = datosEdicion.comentarios || {};

    // 1. COMPARAR DATOS SIMPLES Y CREAR LOG INDIVIDUAL POR CAMPO
    for (var key in columnasAEditar) {
      var colIndex = headers.indexOf(key);
      if (colIndex !== -1 && columnasAEditar[key] !== undefined) {
        var valorViejo = data[filaEncontrada][colIndex] ? data[filaEncontrada][colIndex].toString().trim() : "";
        var valorNuevo = columnasAEditar[key] ? columnasAEditar[key].toString().trim() : "";
        
        if (valorViejo !== valorNuevo) {
          sheet.getRange(filaEncontrada + 1, colIndex + 1).setValue(valorNuevo);
          
          if (key !== "descripcion" && key !== "contacto telefono" && key !== "ubicacion archivos") {
             var categoria = "";
             var detalle = "";
             
             if (key === "descripción") {
                categoria = "DESCRIPCIÓN";
                detalle = "Se modificó el texto detallado.";
             } else {
                categoria = key.toUpperCase();
                var txtV = valorViejo !== "" ? valorViejo : "vacío";
                var txtN = valorNuevo !== "" ? valorNuevo : "vacío";
                detalle = "Cambió de '" + txtV + "' a '" + txtN + "'";
             }
             
             // Mapeo de la clave para coincidir con el JSON del frontend
             var comentKey = key === "descripción" ? "descripcion" : 
                             key === "contacto teléfono" ? "contactoTelefono" :
                             key === "ubicación archivos" ? "ubicacionArchivos" :
                             key === "nro expediente" ? "expediente" :
                             key === "costo obra" ? "costo" : key;
                             
             var comentario = comentariosFront[comentKey] || "";
             registrarHitoHistorial(idProyecto, new Date(), usuario, categoria, detalle, comentario);
          }
        }
      }
    }

    // 2. COMPARAR EQUIPO PROFUNDO (JSON)
    if (datosEdicion.asignados) {
      var colEquipo = headers.indexOf("equipo");
      if (colEquipo !== -1) {
        var equipoViejoStr = data[filaEncontrada][colEquipo] ? data[filaEncontrada][colEquipo].toString() : "[]";
        var equipoNuevoStr = JSON.stringify(datosEdicion.asignados);
        
        if (equipoViejoStr !== equipoNuevoStr) {
          sheet.getRange(filaEncontrada + 1, colEquipo + 1).setValue(equipoNuevoStr);
          
          var eqViejo = [];
          var eqNuevo = [];
          try { eqViejo = JSON.parse(equipoViejoStr); } catch(e){}
          try { eqNuevo = datosEdicion.asignados; } catch(e){}
          
          var mapV = {};
          eqViejo.forEach(function(p){ mapV[p.participante] = p; });
          var mapN = {};
          eqNuevo.forEach(function(p){ mapN[p.participante] = p; });
          
          var logEquipo = [];
          for (var partV in mapV) {
            if (!mapN[partV]) {
              logEquipo.push(partV + " (que tenía el rol de '" + mapV[partV].rol + "') ya no participa en el proyecto");
            } else {
              var v = mapV[partV];
              var n = mapN[partV];
              var subLog = [];
              if (v.rol !== n.rol) subLog.push("pasó al rol de '" + n.rol + "'");
              if (v.encargado !== n.encargado) {
                 if (n.encargado) subLog.push("ahora es encargado");
                 else subLog.push("dejó de ser encargado");
              }
              if (subLog.length > 0) logEquipo.push(partV + " " + subLog.join(" y "));
            }
          }
          
          for (var partN in mapN) {
            if (!mapV[partN]) {
               var extra = mapN[partN].encargado ? " (como encargado)" : "";
               logEquipo.push("Se incorporó " + partN + " con el rol de '" + mapN[partN].rol + "'" + extra);
            }
          }
          
          if (logEquipo.length > 0) {
             var comentEquipo = comentariosFront["equipo"] || "";
             registrarHitoHistorial(idProyecto, new Date(), usuario, "EQUIPO", logEquipo.join(". "), comentEquipo);
          }
        }
      }
    }

    CacheService.getScriptCache().remove("cache_lista_proy");
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function registrarHitoHistorial(idProyecto, fecha, usuario, categoria, detalle, comentario) {
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    var sheetHistorial = ss.getSheetByName("BD_Historial");
    
    if (!sheetHistorial) {
      sheetHistorial = ss.insertSheet("BD_Historial");
      sheetHistorial.appendRow(["ID Proyecto", "Fecha y Hora", "Usuario Responsable", "Categoría del Hito", "Detalle del Cambio", "Comentario"]);
      sheetHistorial.getRange("A1:F1").setFontWeight("bold").setBackground("#D9E2F3");
    } else {
      // Si la columna de comentario no existe (bases viejas), la crea dinámicamente
      if (sheetHistorial.getLastColumn() < 6) {
         sheetHistorial.getRange(1, 6).setValue("Comentario").setFontWeight("bold").setBackground("#D9E2F3");
      }
    }
    
    var usuarioLimpio = usuario ? usuario.toString().trim() : "Sistema";
    var comentLimpio = comentario ? comentario.toString().trim() : "";
    sheetHistorial.appendRow([idProyecto, fecha, usuarioLimpio, categoria, detalle, comentLimpio]);
  } catch (e) {
    console.error("Error guardando historial: " + e.message);
  }
}

function obtenerHistorialProyecto(nombre, fecha) {
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    var sheetBD = ss.getSheetByName("BD_Proyectos");
    var sheetHistorial = ss.getSheetByName("BD_Historial");
    
    if (!sheetHistorial || !sheetBD) return { success: true, datos: [] };
    var dataBD = sheetBD.getDataRange().getValues();
    var headersBD = dataBD[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    var idxNombre = headersBD.indexOf("nombre proyecto") !== -1 ? headersBD.indexOf("nombre proyecto") : headersBD.indexOf("nombre");
    var idxFecha = headersBD.indexOf("fecha");
    var idxId = headersBD.indexOf("id proyecto");

    var idProyecto = "";
    for (var i = 1; i < dataBD.length; i++) {
      var filaNombre = dataBD[i][idxNombre] ? dataBD[i][idxNombre].toString() : "";
      var filaFechaRaw = dataBD[i][idxFecha];
      var filaFecha = filaFechaRaw instanceof Date ? filaFechaRaw.toISOString() : (filaFechaRaw ? filaFechaRaw.toString() : "");
      
      if (filaNombre === nombre && filaFecha === fecha) {
         idProyecto = idxId !== -1 && dataBD[i][idxId] ? dataBD[i][idxId].toString() : "";
         break;
      }
    }

    if (!idProyecto) return { success: true, datos: [], msg: "Este proyecto es antiguo y aún no posee registros en el historial." };

    var dataHist = sheetHistorial.getDataRange().getValues();
    var historial = [];
    
    for (var j = 1; j < dataHist.length; j++) {
      if (dataHist[j][0] && dataHist[j][0].toString() === idProyecto) {
        var fechaRaw = dataHist[j][1];
        var fechaTxt = "";
        if (fechaRaw instanceof Date) {
          fechaTxt = Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        } else {
          fechaTxt = fechaRaw.toString();
        }

        historial.push({
          fecha: fechaTxt,
          usuario: dataHist[j][2] ? dataHist[j][2].toString() : "Sistema",
          categoria: dataHist[j][3] ? dataHist[j][3].toString() : "-",
          detalle: dataHist[j][4] ? dataHist[j][4].toString() : "-",
          comentario: dataHist[j][5] ? dataHist[j][5].toString() : ""
        });
      }
    }
    
    return { success: true, datos: historial.reverse() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}