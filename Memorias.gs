// =====================================================================
// MÓDULO: MEMORIAS DESCRIPTIVAS Y BORRADORES (Lógica de servidor)
// =====================================================================

// --- GESTIÓN DE BORRADORES ---
function guardarBorradorEnDrive(usuarioEmail, disciplina, nombreBorrador, hashes) {
  try {
    var ss = SpreadsheetApp.openById(BORRADORES_SHEET_ID);
    var sheet = ss.getSheets()[0];
    sheet.appendRow([usuarioEmail, disciplina, nombreBorrador, JSON.stringify(hashes), new Date()]);
    
    var cache = CacheService.getScriptCache();
    cache.remove('borradores_v1_' + usuarioEmail.toString().toLowerCase() + '_' + disciplina.toString());
    
    return true;
  } catch (e) { return false; }
}

function obtenerBorradoresDeDrive(usuarioEmail, disciplina, forzarRecarga) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'borradores_v1_' + usuarioEmail.toString().toLowerCase() + '_' + disciplina.toString();
    
    if (!forzarRecarga) {
      var cachedData = cache.get(cacheKey);
      if (cachedData) return JSON.parse(cachedData);
    }

    var ss = SpreadsheetApp.openById(BORRADORES_SHEET_ID);
    var data = ss.getSheets()[0].getDataRange().getValues();
    var misBorradores = {};
    for (var i = 1; i < data.length; i++) { 
      if (data[i][0].toString().toLowerCase() === usuarioEmail.toString().toLowerCase() && data[i][1].toString() === disciplina.toString()) {
        misBorradores[data[i][2]] = {
            hashes: JSON.parse(data[i][3]),
            fecha: data[i][4] ? data[i][4].toString() : "Fecha desconocida" 
        };
      }
    }
    
    cache.put(cacheKey, JSON.stringify(misBorradores), 21600); // 6 horas
    return misBorradores;
  } catch (e) { return {}; }
}

function eliminarBorradorEnDrive(usuarioEmail, disciplina, nombreBorrador) {
  try {
    var ss = SpreadsheetApp.openById(BORRADORES_SHEET_ID);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if (data[i][0].toString().toLowerCase() == usuarioEmail.toString().toLowerCase() && 
          data[i][1].toString() == disciplina.toString() && 
          data[i][2].toString() == nombreBorrador.toString()) {
        sheet.deleteRow(i + 1);
      }
    }
    
    var cache = CacheService.getScriptCache();
    cache.remove('borradores_v1_' + usuarioEmail.toString().toLowerCase() + '_' + disciplina.toString());
    
    return true;
  } catch (e) { return false; }
}

function eliminarTodosBorradoresEnDrive(usuarioEmail, disciplina) {
  try {
    var ss = SpreadsheetApp.openById(BORRADORES_SHEET_ID);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var rowsToDelete = [];
    
    for (var i = data.length - 1; i >= 0; i--) {
      if (data[i][0].toString().toLowerCase() == usuarioEmail.toString().toLowerCase() && 
          data[i][1].toString() == disciplina.toString()) {
        rowsToDelete.push(i + 1);
      }
    }
    
    for (var j = 0; j < rowsToDelete.length; j++) {
      sheet.deleteRow(rowsToDelete[j]);
    }
    
    var cache = CacheService.getScriptCache();
    cache.remove('borradores_v1_' + usuarioEmail.toString().toLowerCase() + '_' + disciplina.toString());
    
    return true;
  } catch (e) { return false; }
}

// --- GENERADOR DE MEMORIAS ---
function obtenerDisciplinas(forzarRecarga) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'listaDisciplinas_v1';
    
    if (!forzarRecarga) {
      var cachedData = cache.get(cacheKey);
      if (cachedData) return JSON.parse(cachedData);
    }

    var ss = SpreadsheetApp.openById(RUBROS_SHEET_ID);
    var disciplinas = ss.getSheets().map(s => s.getName());
    
    cache.put(cacheKey, JSON.stringify(disciplinas), 21600);
    return disciplinas;
  } catch (e) {
    return ["⚠️ Error: Verifica el ID de la base de datos"];
  }
}

function obtenerRubros(nombreHoja, forzarRecarga) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'rubros_v1_' + nombreHoja.toString();
    
    if (!forzarRecarga) {
      var cachedData = cache.get(cacheKey);
      if (cachedData) return JSON.parse(cachedData);
    }

    var ss = SpreadsheetApp.openById(RUBROS_SHEET_ID);
    var data = ss.getSheetByName(nombreHoja).getDataRange().getValues();
    var rubros = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] || data[i][2] || data[i][3] || data[i][4]) {
        rubros.push({ 
          id: i, 
          bloque: data[i][1]?.toString().trim().toUpperCase() || "", 
          principal: data[i][2] || "", 
          secundario: data[i][3] || "", 
          terciario: data[i][4] || "", 
          texto: data[i][5] || "", 
          obligatorio: data[i][6] === true 
        });
      }
    }
    
    try {
      cache.put(cacheKey, JSON.stringify(rubros), 21600);
    } catch (errorCache) {}
    
    return rubros;
    
  } catch (e) {
    return [];
  }
}

function generarMemoriaWeb(datosFormulario, usuarioEmail) {
  var template = DriveApp.getFileById(TEMPLATE_ID);
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var nombreArch = 'Memoria - ' + datosFormulario.disciplina;
  var newFile = template.makeCopy(nombreArch, folder);
  var docUrl = newFile.getUrl(); 
  var docId = newFile.getId();
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  
  body.replaceText('\\(\\*\\*colocar disciplina\\*\\*\\)', datosFormulario.disciplina.toUpperCase());
  
  var tag = body.findText('{{MEMORIA}}');
  if (!tag) return { error: 'No se encontró {{MEMORIA}}' };
  
  var tagElement = tag.getElement().getParent();
  var insertIndex = body.getChildIndex(tagElement);
  var numP = -1, numS = 0, numT = 0, bA = "", indice = [];
  
  datosFormulario.seleccionados.forEach(item => {
    var lB = item.bloque.toUpperCase();
    if (lB !== bA) { bA = lB; numP = -1; numS = 0; numT = 0; }
    
    if (item.principal) {
      numP++; numS = 0; numT = 0;
      var tP = bA + "." + numP + " - " + item.principal;
      indice.push({texto: tP, nivel: 1});
      
      var p = body.insertParagraph(insertIndex++, tP);
      p.setHeading(numP === 0 ? DocumentApp.ParagraphHeading.HEADING1 : DocumentApp.ParagraphHeading.HEADING2)
       .setLineSpacing(1).setSpacingBefore(0).setSpacingAfter(0);
       
    } else if (item.secundario) {
      numS++; numT = 0;
      var tS = bA + "." + (numP === -1 ? 0 : numP) + "." + numS + " - " + item.secundario;
      
      var p = body.insertParagraph(insertIndex++, tS);
      p.setHeading(DocumentApp.ParagraphHeading.HEADING3)
       .setLineSpacing(1).setSpacingBefore(0).setSpacingAfter(0);
       
    } else if (item.terciario) {
      numT++;
      var tT = bA + "." + (numP === -1 ? 0 : numP) + "." + (numS === 0 ? 1 : numS) + "." + numT + " - " + item.terciario;
      var p = body.insertParagraph(insertIndex++, tT);
      p.setHeading(DocumentApp.ParagraphHeading.HEADING4)
       .setLineSpacing(1).setSpacingBefore(0).setSpacingAfter(0);
    }
    
    if (item.texto) {
      var p = body.insertParagraph(insertIndex++, item.texto);
      p.setHeading(DocumentApp.ParagraphHeading.NORMAL)
       .setLineSpacing(1).setSpacingBefore(0).setSpacingAfter(0);
    }
    
    var espacio = body.insertParagraph(insertIndex++, '');
    espacio.setHeading(DocumentApp.ParagraphHeading.NORMAL)
           .setLineSpacing(1).setSpacingBefore(0).setSpacingAfter(0);
  });

  if (indice.length > 0) {
    body.insertPageBreak(insertIndex++);
    var tI = body.insertParagraph(insertIndex++, 'ÍNDICE');
    tI.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    tI.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    tI.setSpacingAfter(10);
    
    indice.forEach(function(i) {
      var p = body.insertParagraph(insertIndex++, i.texto);
      p.setHeading(DocumentApp.ParagraphHeading.NORMAL)
       .setLineSpacing(1).setSpacingBefore(0).setSpacingAfter(0);
    });
  }

  body.replaceText('{{MEMORIA}}', '');
  doc.saveAndClose();
  
  registrarArchivoGenerado(usuarioEmail, nombreArch, docId);
  return { url: docUrl };
}