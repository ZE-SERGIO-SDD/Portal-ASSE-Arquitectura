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
      "creado por": datosProyecto.usuario, 
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

    CacheService.getScriptCache().remove("cache_lista_proy_chunks");
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function cambiarEstadoProyecto(nombre, fecha, nuevoEstado, usuarioNombre, comentario, fechaCambio) {
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
        var fechaRegistro = new Date();
        if (fechaCambio) {
          var partes = fechaCambio.split("-");
          if(partes.length === 3) fechaRegistro = new Date(partes[0], partes[1] - 1, partes[2], 12, 0, 0);
        }

        registrarHitoHistorial(idProyecto, fechaRegistro, usuarioNombre || "Usuario Desconocido", "ESTADO", detalleCambio, comentario || "");

        CacheService.getScriptCache().remove("cache_lista_proy_chunks");
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

    var fechaRegistro = new Date();
    if (datosEdicion.fechaCambio) {
      var partes = datosEdicion.fechaCambio.split("-");
      if(partes.length === 3) fechaRegistro = new Date(partes[0], partes[1] - 1, partes[2], 12, 0, 0);
    }

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
             
             var txtV = valorViejo !== "" ? valorViejo : "vacío";
             var txtN = valorNuevo !== "" ? valorNuevo : "vacío";

             if (key === "descripción") {
                categoria = "DESCRIPCIÓN";
                detalle = "Cambió de:\n«" + txtV + "»\n\na:\n«" + txtN + "»";
             } else {
                categoria = key.toUpperCase();
                detalle = "Cambió de '" + txtV + "' a '" + txtN + "'";
             }
             
             var comentKey = key === "descripción" ? "descripcion" : 
                             key === "contacto teléfono" ? "contactoTelefono" :
                             key === "ubicación archivos" ? "ubicacionArchivos" :
                             key === "nro expediente" ? "expediente" :
                             key === "costo obra" ? "costo" : key;
                             
             var comentario = comentariosFront[comentKey] || "";
             registrarHitoHistorial(idProyecto, fechaRegistro, usuario, categoria, detalle, comentario);
          }
        }
      }
    }

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
             registrarHitoHistorial(idProyecto, fechaRegistro, usuario, "EQUIPO", logEquipo.join(". "), comentEquipo);
          }
        }
      }
    }

    CacheService.getScriptCache().remove("cache_lista_proy_chunks");
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
    
    // Si no lo encuentra en activos, lo buscamos en los archivados
    if (!sheetBD || !sheetHistorial) {
        sheetBD = ss.getSheetByName("BD_Proyectos_Archivados");
        sheetHistorial = ss.getSheetByName("BD_Historial_Archivados");
        if (!sheetBD || !sheetHistorial) return { success: true, datos: [] };
    }
    
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

    // Doble verificación: Si aún no se encontró el ID, probamos en los archivados por si acaso
    if (!idProyecto) {
      var sheetBDArchivados = ss.getSheetByName("BD_Proyectos_Archivados");
      if(sheetBDArchivados) {
         var dataArch = sheetBDArchivados.getDataRange().getValues();
         for (var k = 1; k < dataArch.length; k++) {
            var fNomb = dataArch[k][idxNombre] ? dataArch[k][idxNombre].toString() : "";
            var fDateR = dataArch[k][idxFecha];
            var fDate = fDateR instanceof Date ? fDateR.toISOString() : (fDateR ? fDateR.toString() : "");
            if(fNomb === nombre && fDate === fecha) {
               idProyecto = idxId !== -1 && dataArch[k][idxId] ? dataArch[k][idxId].toString() : "";
               sheetHistorial = ss.getSheetByName("BD_Historial_Archivados");
               break;
            }
         }
      }
    }

    if (!idProyecto) return { success: true, datos: [], msg: "Este proyecto es antiguo o no posee registros en el historial actual." };

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

function archivarProyectoBackend(nombreOriginal, fechaOriginal, fechaArchivo, comentario, usuarioNombre) {
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    var sheetBD = ss.getSheetByName("BD_Proyectos");
    var sheetHistorial = ss.getSheetByName("BD_Historial");
    
    // Se crean las pestañas de archivados si no existen
    var sheetBDArchivados = ss.getSheetByName("BD_Proyectos_Archivados") || ss.insertSheet("BD_Proyectos_Archivados");
    var sheetHistArchivados = ss.getSheetByName("BD_Historial_Archivados") || ss.insertSheet("BD_Historial_Archivados");

    // Copiar encabezados a las pestañas nuevas si están vacías
    if (sheetBDArchivados.getLastRow() === 0 && sheetBD) {
       sheetBDArchivados.appendRow(sheetBD.getRange(1, 1, 1, sheetBD.getLastColumn()).getValues()[0]);
    }
    if (sheetHistArchivados.getLastRow() === 0 && sheetHistorial) {
       sheetHistArchivados.appendRow(sheetHistorial.getRange(1, 1, 1, sheetHistorial.getLastColumn()).getValues()[0]);
    }

    var dataBD = sheetBD.getDataRange().getValues();
    var headersBD = dataBD[0].map(h => h.toString().toLowerCase().trim());
    var idxNombre = headersBD.indexOf("nombre proyecto") !== -1 ? headersBD.indexOf("nombre proyecto") : headersBD.indexOf("nombre");
    var idxFecha = headersBD.indexOf("fecha");
    var idxId = headersBD.indexOf("id proyecto");

    var filaEncontrada = -1;
    var idProyecto = "";

    // 1. Ubicar el proyecto exacto en la BD principal
    for (var i = 1; i < dataBD.length; i++) {
      var filaNombre = dataBD[i][idxNombre] ? dataBD[i][idxNombre].toString() : "";
      var fDate = dataBD[i][idxFecha];
      var filaFecha = fDate instanceof Date ? fDate.toISOString() : (fDate ? fDate.toString() : "");
      
      if (filaNombre === nombreOriginal && filaFecha === fechaOriginal) {
        filaEncontrada = i + 1; // getRange usa índice base 1
        idProyecto = (idxId !== -1 && dataBD[i][idxId]) ? dataBD[i][idxId].toString() : "";
        break;
      }
    }

    if (filaEncontrada === -1) return { success: false, error: "No se pudo encontrar el proyecto en la base de datos principal para archivarlo." };

    // 2. Mudar el proyecto a BD_Proyectos_Archivados y eliminar de BD_Proyectos
    var filaData = sheetBD.getRange(filaEncontrada, 1, 1, sheetBD.getLastColumn()).getValues()[0];
    sheetBDArchivados.appendRow(filaData);
    sheetBD.deleteRow(filaEncontrada);

    // Formatear la fecha elegida por el usuario
    var fechaRegistro = new Date();
    if (fechaArchivo) {
      var partes = fechaArchivo.split("-");
      if (partes.length === 3) fechaRegistro = new Date(partes[0], partes[1] - 1, partes[2], 12, 0, 0);
    }
    
    // 3. Mudar todo el historial asociado a BD_Historial_Archivados
    if (sheetHistorial && idProyecto) {
       var dataHist = sheetHistorial.getDataRange().getValues();
       // Iteramos desde el final hacia el principio para poder usar deleteRow() sin alterar los índices de las filas de arriba
       for (var j = dataHist.length - 1; j >= 1; j--) {
         if (dataHist[j][0] && dataHist[j][0].toString() === idProyecto) {
            sheetHistArchivados.appendRow(dataHist[j]);
            sheetHistorial.deleteRow(j + 1); // +1 por el encabezado
         }
       }
    }
    
    // 4. Agregar el hito final de "ARCHIVADO" directamente en el historial archivado
    var usuarioLimpio = usuarioNombre ? usuarioNombre.toString().trim() : "Sistema";
    var comentLimpio = comentario ? comentario.toString().trim() : "Sin justificación.";
    sheetHistArchivados.appendRow([idProyecto, fechaRegistro, usuarioLimpio, "ARCHIVADO", "El proyecto y todo su registro fueron transferidos al archivo histórico del portal.", comentLimpio]);
    
    // Limpiamos la caché general para forzar la recarga de los listados
    CacheService.getScriptCache().remove("cache_lista_proy_chunks");
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ---------------------------------------------------------------------
// REACTIVACIÓN / MOVIMIENTO DESDE ARCHIVADOS
// ---------------------------------------------------------------------
function reactivarProyectoBackend(nombreOriginal, fechaOriginal, fechaAccion, comentario, usuarioNombre) {
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    // Hojas de origen
    var sheetBDArchivados = ss.getSheetByName("BD_Proyectos_Archivados");
    var sheetHistArchivados = ss.getSheetByName("BD_Historial_Archivados");
    
    // Hojas de destino
    var sheetBDActivos = ss.getSheetByName("BD_Proyectos") || ss.insertSheet("BD_Proyectos");
    var sheetHistActivos = ss.getSheetByName("BD_Historial") || ss.insertSheet("BD_Historial");
    
    if (!sheetBDArchivados) return { success: false, error: "No existe la base de datos de archivados." };
    
    var dataArch = sheetBDArchivados.getDataRange().getValues();
    var headersArch = dataArch[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    
    var idxNombre = headersArch.indexOf("nombre proyecto") !== -1 ? headersArch.indexOf("nombre proyecto") : headersArch.indexOf("nombre");
    var idxFecha = headersArch.indexOf("fecha");
    var idxId = headersArch.indexOf("id proyecto");

    var filaEncontrada = -1;
    var idProyecto = "";

    // 1. Ubicar el proyecto en Archivados
    for (var i = 1; i < dataArch.length; i++) {
      var filaNombre = dataArch[i][idxNombre] ? dataArch[i][idxNombre].toString() : "";
      var fDate = dataArch[i][idxFecha];
      var filaFecha = fDate instanceof Date ? fDate.toISOString() : (fDate ? fDate.toString() : "");
      
      if (filaNombre === nombreOriginal && filaFecha === fechaOriginal) {
        filaEncontrada = i + 1; // base 1
        idProyecto = (idxId !== -1 && dataArch[i][idxId]) ? dataArch[i][idxId].toString() : "";
        break;
      }
    }

    if (filaEncontrada === -1) return { success: false, error: "No se pudo encontrar el proyecto en la base de datos de archivados." };

    var filaData = sheetBDArchivados.getRange(filaEncontrada, 1, 1, sheetBDArchivados.getLastColumn()).getValues()[0];

    // Formatear fecha
    var fechaRegistro = new Date();
    if (fechaAccion) {
      var partes = fechaAccion.split("-");
      if (partes.length === 3) fechaRegistro = new Date(partes[0], partes[1] - 1, partes[2], 12, 0, 0);
    }
    
    var usuarioLimpio = usuarioNombre ? usuarioNombre.toString().trim() : "Sistema";
    var comentLimpio = comentario ? comentario.toString().trim() : "Reactivado sin justificación.";

    // 2. Mover a Activos
    sheetBDActivos.appendRow(filaData);
    sheetBDArchivados.deleteRow(filaEncontrada);
    
    // 3. Mover historial completo
    if (sheetHistArchivados && idProyecto) {
       var dataHist = sheetHistArchivados.getDataRange().getValues();
       // Iteramos hacia atrás para borrar filas sin afectar los índices
       for (var j = dataHist.length - 1; j >= 1; j--) {
         if (dataHist[j][0] && dataHist[j][0].toString() === idProyecto) {
            sheetHistActivos.appendRow(dataHist[j]);
            sheetHistArchivados.deleteRow(j + 1);
         }
       }
    }
    
    // 4. Agregar hito final en la nueva base de Activos
    sheetHistActivos.appendRow([idProyecto, fechaRegistro, usuarioLimpio, "REACTIVADO", "El proyecto fue reactivado y transferido al listado de Proyectos en Curso.", comentLimpio]);
    
    // Limpiar caché de ambos listados para que el Front-end se refresque instantáneamente
    CacheService.getScriptCache().remove("cache_lista_proy_chunks");
    CacheService.getScriptCache().remove("cache_lista_arch_chunks");
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// =====================================================================
// PAPELERA: ENVIAR PROYECTO A ELIMINADOS (Borrado Lógico)
// =====================================================================
function eliminarProyectoBackend(nombreOriginal, fechaOriginal, usuarioNombre) {
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    var sheetActivos = ss.getSheetByName("BD_Proyectos");
    var filaEncontrada = -1;
    var origen = "";
    var sheetOrigen = null;
    var rowData = null;
    var idProyecto = "";

    // 1. Buscar en BD_Proyectos (Proyectos En Curso)
    if (sheetActivos) {
      var dataActivos = sheetActivos.getDataRange().getValues();
      var headersActivos = dataActivos[0].map(function(h) { return h.toString().toLowerCase().trim(); });
      var idxNombre = headersActivos.indexOf("nombre proyecto") !== -1 ? headersActivos.indexOf("nombre proyecto") : headersActivos.indexOf("nombre");
      var idxFecha = headersActivos.indexOf("fecha");
      var idxId = headersActivos.indexOf("id proyecto");

      for (var i = 1; i < dataActivos.length; i++) {
        var filaNombre = dataActivos[i][idxNombre] ? dataActivos[i][idxNombre].toString() : "";
        var fDate = dataActivos[i][idxFecha];
        var filaFecha = fDate instanceof Date ? fDate.toISOString() : (fDate ? fDate.toString() : "");
        
        if (filaNombre === nombreOriginal && filaFecha === fechaOriginal) {
          filaEncontrada = i + 1;
          origen = "En Curso";
          sheetOrigen = sheetActivos;
          rowData = dataActivos[i];
          idProyecto = (idxId !== -1 && dataActivos[i][idxId]) ? dataActivos[i][idxId].toString() : "";
          break;
        }
      }
    }

    // 2. Si no está en Activos, buscar en BD_Proyectos_Archivados
    if (filaEncontrada === -1) {
      var sheetArchivados = ss.getSheetByName("BD_Proyectos_Archivados");
      if (sheetArchivados) {
        var dataArch = sheetArchivados.getDataRange().getValues();
        var headersArch = dataArch[0].map(function(h) { return h.toString().toLowerCase().trim(); });
        var idxNombreA = headersArch.indexOf("nombre proyecto") !== -1 ? headersArch.indexOf("nombre proyecto") : headersArch.indexOf("nombre");
        var idxFechaA = headersArch.indexOf("fecha");
        var idxIdA = headersArch.indexOf("id proyecto");

        for (var i = 1; i < dataArch.length; i++) {
          var filaNombre = dataArch[i][idxNombreA] ? dataArch[i][idxNombreA].toString() : "";
          var fDate = dataArch[i][idxFechaA];
          var filaFecha = fDate instanceof Date ? fDate.toISOString() : (fDate ? fDate.toString() : "");
          
          if (filaNombre === nombreOriginal && filaFecha === fechaOriginal) {
            filaEncontrada = i + 1;
            origen = "Archivado";
            sheetOrigen = sheetArchivados;
            rowData = dataArch[i];
            idProyecto = (idxIdA !== -1 && dataArch[i][idxIdA]) ? dataArch[i][idxIdA].toString() : "";
            break;
          }
        }
      }
    }

    if (filaEncontrada === -1) {
      return { success: false, error: "No se encontró el proyecto para eliminar." };
    }

    // 3. Escribir en BD_Proyectos_Eliminados
    var sheetEliminados = ss.getSheetByName("BD_Proyectos_Eliminados");
    if (!sheetEliminados) {
      sheetEliminados = ss.insertSheet("BD_Proyectos_Eliminados");
      var headersOrigen = sheetOrigen.getRange(1, 1, 1, sheetOrigen.getLastColumn()).getValues()[0].slice();
      headersOrigen.push("Origen", "Fecha Eliminacion", "Eliminado Por");
      sheetEliminados.appendRow(headersOrigen);
    }

    var fechaEliminacion = new Date();
    var user = usuarioNombre || "Usuario Desconocido";

    var nuevaFila = rowData.slice();
    nuevaFila.push(origen);
    nuevaFila.push(fechaEliminacion);
    nuevaFila.push(user);

    sheetEliminados.appendRow(nuevaFila);
    
    // 4. Eliminar de la hoja original
    sheetOrigen.deleteRow(filaEncontrada);

    // 5. Mudar el Historial (Bitácora) si existe
    if (idProyecto) {
      mudarHistorialEliminado(ss, idProyecto, origen);
    }

    // 6. Limpiar Caché
    var cache = CacheService.getScriptCache();
    cache.remove("cache_lista_proy_chunks"); 
    cache.remove("cache_lista_arch_chunks");

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// Función auxiliar para mudar la bitácora
function mudarHistorialEliminado(ss, idProyecto, origen) {
  var nombrePestañaHistorial = (origen === "Archivado") ? "BD_Historial_Archivados" : "BD_Historial";
  var sheetHistorial = ss.getSheetByName(nombrePestañaHistorial); 
  var sheetHistEliminados = ss.getSheetByName("BD_Historial_Eliminados") || ss.insertSheet("BD_Historial_Eliminados");
  
  if (!sheetHistorial) return;
  
  if (sheetHistEliminados.getLastRow() === 0) {
    var headersHist = sheetHistorial.getRange(1, 1, 1, sheetHistorial.getLastColumn()).getValues()[0].slice();
    headersHist.push("Origen");
    sheetHistEliminados.appendRow(headersHist);
  }
  
  var dataHist = sheetHistorial.getDataRange().getValues();
  var filasAEliminar = [];
  
  // Recorremos de atrás hacia adelante
  for (var i = dataHist.length - 1; i >= 1; i--) { 
    var histId = dataHist[i][0] ? dataHist[i][0].toString() : "";
    
    if (histId === idProyecto) {
      var filaHist = dataHist[i].slice();
      filaHist.push(origen); // Guardamos de dónde vino
      sheetHistEliminados.appendRow(filaHist);
      filasAEliminar.push(i + 1); // Anotamos qué fila borrar
    }
  }

  // Borramos las filas de la hoja original de forma segura
  for (var k = 0; k < filasAEliminar.length; k++) {
    sheetHistorial.deleteRow(filasAEliminar[k]);
  }
}