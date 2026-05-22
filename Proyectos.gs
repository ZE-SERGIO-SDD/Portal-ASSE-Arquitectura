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
      "Fecha", "Departamento", "Centro", "Nombre Proyecto", "Descripción", 
      "Tipo de Obra", "Estado", "Equipo", "Creado por",
      "Nro Expediente", "Costo Obra", "Empresa", "Contacto Nombre", "Contacto Teléfono", "Ubicación Archivos", "Fechas Hitos"
    ];
    
    if (sheet.getLastRow() === 0) sheet.appendRow(columnasRequeridas);
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    
    var mapaDatos = {
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
    CacheService.getScriptCache().remove("cache_lista_proy");
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}