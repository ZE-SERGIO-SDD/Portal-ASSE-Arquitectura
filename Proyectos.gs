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
  var cacheKey = "cache_params_proy_v5";
  if (!forzar) {
    var cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
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
    } catch (e) { 
      console.error("Error partic: " + e.message);
    }
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
      console.error("Error al mapear/crear empresas externas: " + e.message);
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
    try { cache.put(cacheKey, JSON.stringify(result), 21600); } catch(e) {}
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
      "Nro Expediente", "Costo Obra", "Empresa", "Contacto Nombre", "Contacto Teléfono", "Ubicación Archivos", "Fechas Hitos"
    ];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(columnasRequeridas);
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    
    if (headers.indexOf("id proyecto") === -1) {
      sheet.getRange(1, headers.length + 1).setValue("ID Proyecto");
      headers.push("id proyecto");
    }
    
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
      "fechas hitos": JSON.stringify(datosProyecto.hitos || {})
    };
    
    var nuevaFila = new Array(headers.length);
    for (var j = 0; j < headers.length; j++) {
      var hName = headers[j];
      nuevaFila[j] = mapaDatos.hasOwnProperty(hName) ? mapaDatos[hName] : "";
    }
    
    sheet.appendRow(nuevaFila);
    
    registrarHitoHistorial(idUnico, new Date(), datosProyecto.usuario, "Creación de Proyecto", "Alta inicial en estado: " + datosProyecto.estado);

    CacheService.getScriptCache().remove("cache_lista_proy");
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function cambiarEstadoProyecto(nombre, fecha, nuevoEstado) {
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
        
        var usuarioActivo = "";
        try { usuarioActivo = Session.getActiveUser().getEmail(); } catch(e) {}
        if (!usuarioActivo) usuarioActivo = "Actualización de Estado (App)";

        var detalleCambio = "Pasó de '" + estadoAnterior + "' a '" + nuevoEstado + "'";
        registrarHitoHistorial(idProyecto, new Date(), usuarioActivo, "Cambio de Estado", detalleCambio);

        CacheService.getScriptCache().remove("cache_lista_proy");
        return { success: true };
      }
    }
    return { success: false, error: "No se encontró el proyecto para actualizar." };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// --- NUEVA FUNCIÓN: Gestor de Edición de Datos ---
function editarDatosProyecto(datosEdicion) {
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    var sheet = ss.getSheetByName("BD_Proyectos");
    if (!sheet) return { success: false, error: "No se encontró la hoja BD_Proyectos" };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    
    var idxNombre = headers.indexOf("nombre proyecto") !== -1 ? headers.indexOf("nombre proyecto") : headers.indexOf("nombre");
    var idxFecha = headers.indexOf("fecha");
    var idxId = headers.indexOf("id proyecto");
    
    var filaEncontrada = -1;
    var idProyecto = "";

    // Buscar el proyecto por Nombre y Fecha original (por si no tenemos el ID aún en la vista)
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

    // Si es viejo y no tiene ID, le generamos uno
    if (!idProyecto) {
      idProyecto = "PRY-MIG-" + Date.now().toString(36).toUpperCase();
      if (idxId !== -1) sheet.getRange(filaEncontrada + 1, idxId + 1).setValue(idProyecto);
    }

    var detallesCambios = [];

    // Mapeo de columnas a actualizar
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
      "ubicacion archivos": datosEdicion.ubicacionArchivos
    };

    // Actualizar celdas de texto simples y guardar en historial
    for (var key in columnasAEditar) {
      var colIndex = headers.indexOf(key);
      if (colIndex !== -1 && columnasAEditar[key] !== undefined) {
        var valorViejo = data[filaEncontrada][colIndex] ? data[filaEncontrada][colIndex].toString() : "";
        var valorNuevo = columnasAEditar[key] ? columnasAEditar[key].toString() : "";
        
        if (valorViejo !== valorNuevo) {
          sheet.getRange(filaEncontrada + 1, colIndex + 1).setValue(valorNuevo);
          // Registramos en el detalle solo las columnas principales, no las duplicadas sin tilde
          if (key !== "descripcion" && key !== "contacto telefono" && key !== "ubicacion archivos") {
             detallesCambios.push("Se editó " + key.toUpperCase());
          }
        }
      }
    }

    // Actualizar equipo (participantes) si viene en los datos
    if (datosEdicion.asignados) {
      var colEquipo = headers.indexOf("equipo");
      if (colEquipo !== -1) {
        var equipoViejo = data[filaEncontrada][colEquipo] ? data[filaEncontrada][colEquipo].toString() : "";
        var equipoNuevo = JSON.stringify(datosEdicion.asignados);
        if (equipoViejo !== equipoNuevo) {
          sheet.getRange(filaEncontrada + 1, colEquipo + 1).setValue(equipoNuevo);
          detallesCambios.push("Se modificaron los Participantes/Roles");
        }
      }
    }

    // Si hubo algún cambio real, lo grabamos en la Bitácora
    if (detallesCambios.length > 0) {
      var usuarioActivo = "";
      try { usuarioActivo = Session.getActiveUser().getEmail(); } catch(e) {}
      if (!usuarioActivo) usuarioActivo = datosEdicion.usuario || "Edición de Datos (App)";

      var textoLog = detallesCambios.join(" | ");
      registrarHitoHistorial(idProyecto, new Date(), usuarioActivo, "Edición de Datos", textoLog);
      
      CacheService.getScriptCache().remove("cache_lista_proy");
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function registrarHitoHistorial(idProyecto, fecha, usuario, categoria, detalle) {
  try {
    var ss = SpreadsheetApp.openById(PROYECTOS_CONFIG_ID);
    var sheetHistorial = ss.getSheetByName("BD_Historial");
    
    if (!sheetHistorial) {
      sheetHistorial = ss.insertSheet("BD_Historial");
      sheetHistorial.appendRow(["ID Proyecto", "Fecha y Hora", "Usuario Responsable", "Categoría del Hito", "Detalle del Cambio"]);
      sheetHistorial.getRange("A1:E1").setFontWeight("bold").setBackground("#D9E2F3");
    }
    
    sheetHistorial.appendRow([idProyecto, fecha, usuario, categoria, detalle]);
  } catch (e) {
    console.error("Error guardando historial: " + e.message);
  }
}