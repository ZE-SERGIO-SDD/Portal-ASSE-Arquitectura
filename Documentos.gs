// =====================================================================
// MÓDULO: DOCUMENTOS (Lógica de servidor) - VERSIÓN DOS CARPETAS
// =====================================================================

// IDs de las carpetas proporcionadas
var FOLDER_EDITABLES_ID = '1V-2QduJzOVFispP-k2UtGKjhPuXyE5gc';
var FOLDER_ARMADOS_ID = '1tugX6oZL79fguCf8z9XiboP9DSQsWkUQ';

function listarDocumentosEditables(forzarRecarga) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'listaDocsEditables_v2';
    
    if (!forzarRecarga) {
      var cachedData = cache.get(cacheKey);
      if (cachedData) return JSON.parse(cachedData);
    }

    var folder = DriveApp.getFolderById(FOLDER_EDITABLES_ID);
    var files = folder.getFiles();
    var lista = [];
    while (files.hasNext()) {
      var file = files.next();
      if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
        lista.push({ nombre: file.getName(), id: file.getId() });
      }
    }
    lista.sort((a,b) => a.nombre.localeCompare(b.nombre));
    cache.put(cacheKey, JSON.stringify(lista), 21600);
    return lista;
  } catch (e) { 
    return { error: e.message };
  }
}

function listarDocumentosArmados(forzarRecarga) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'listaDocsArmados_v2';
    
    if (!forzarRecarga) {
      var cachedData = cache.get(cacheKey);
      if (cachedData) return JSON.parse(cachedData);
    }

    var folder = DriveApp.getFolderById(FOLDER_ARMADOS_ID);
    var files = folder.getFiles();
    var lista = [];
    while (files.hasNext()) {
      var file = files.next();
      // Listamos los archivos armados (PDFs, Word, etc.) y obtenemos su URL para descarga directa
      lista.push({ nombre: file.getName(), id: file.getId(), url: file.getUrl() });
    }
    lista.sort((a,b) => a.nombre.localeCompare(b.nombre));
    cache.put(cacheKey, JSON.stringify(lista), 21600);
    return lista;
  } catch (e) { 
    return { error: e.message };
  }
}

function generarDocumentoMembretado(idContenido, usuarioEmail, usuarioNombre) {
  try {
    var lamina = DriveApp.getFileById(LAMINA_TEMPLATE_ID);
    var contenidoDoc = DocumentApp.openById(idContenido);
    // Asumo que FOLDER_ID es la carpeta donde se guardan temporalmente los generados (definida globalmente)
    var folder = DriveApp.getFolderById(FOLDER_ID); 
    
    // Formatear el nombre
    var nombreSeguro = usuarioNombre ? usuarioNombre.trim().replace(/\s+/g, '_') : "Usuario";
    var nombreFinal = nombreSeguro + "_" + contenidoDoc.getName();
    
    var copiaDocFile = lamina.makeCopy(nombreFinal, folder);
    var docId = copiaDocFile.getId();
    var copiaDoc = DocumentApp.openById(docId);
    var bodyCopia = copiaDoc.getBody();
    
    var tag = bodyCopia.findText('{{CUERPO}}');
    if (!tag) return { error: 'Marca {{CUERPO}} no encontrada en la lámina.' };
    
    var element = tag.getElement();
    var parent = element.getParent();
    var index = bodyCopia.getChildIndex(parent);
    
    var totalElements = contenidoDoc.getBody().getNumChildren();
    
    for (var j = 0; j < totalElements; j++) {
      var originalElement = contenidoDoc.getBody().getChild(j);
      var type = originalElement.getType();
      var childCopy = originalElement.copy();
      
      if (type == DocumentApp.ElementType.PARAGRAPH) {
        bodyCopia.insertParagraph(index++, childCopy);
      } 
      else if (type == DocumentApp.ElementType.TABLE) {
        bodyCopia.insertTable(index++, childCopy);
      } 
      else if (type == DocumentApp.ElementType.LIST_ITEM) {
        var listItem = childCopy.asListItem();
        var glyphType = originalElement.asListItem().getGlyphType();
        var newListItem = bodyCopia.insertListItem(index++, listItem);
        newListItem.setGlyphType(glyphType);
      }
    }
    
    bodyCopia.replaceText('{{CUERPO}}', '');
    copiaDoc.saveAndClose();
    
    // Si registrarArchivoGenerado existe en otra parte, se ejecuta normal
    if (typeof registrarArchivoGenerado === "function") {
      registrarArchivoGenerado(usuarioEmail, nombreFinal, docId);
    }
    
    return { url: copiaDocFile.getUrl() };
  } catch (e) { 
    return { error: e.message }; 
  }
}