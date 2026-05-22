// =====================================================================
// MÓDULO: DOCUMENTOS (Lógica de servidor) - VERSIÓN CORREGIDA LISTAS
// =====================================================================

function listarDocumentosContenido(forzarRecarga) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'listaDocumentos_v1';
    
    if (!forzarRecarga) {
      var cachedData = cache.get(cacheKey);
      if (cachedData) return JSON.parse(cachedData); 
    }

    var folder = DriveApp.getFolderById(FOLDER_CONTENIDOS_ID);
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

function generarDocumentoMembretado(idContenido, usuarioEmail, usuarioNombre) {
  try {
    var lamina = DriveApp.getFileById(LAMINA_TEMPLATE_ID);
    var contenidoDoc = DocumentApp.openById(idContenido);
    var folder = DriveApp.getFolderById(FOLDER_ID);
    
    // Formatear el nombre: Cambia los espacios por guiones bajos (Ej: Sergio Dogliotti -> Sergio_Dogliotti)
    var nombreSeguro = usuarioNombre ? usuarioNombre.trim().replace(/\s+/g, '_') : "Usuario";
    
    // Armar el nombre final: Nombre_Apellido_NombredelDocumento
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
    
    // Diccionario para rastrear y mantener la continuidad de las listas
    var listaMap = {};

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
        // --- SOLUCIÓN PARA VIÑETAS ---
        var listItem = childCopy.asListItem();
        var glyphType = originalElement.asListItem().getGlyphType();
        
        // Insertamos el ítem y le re-asignamos manualmente su estilo de viñeta
        var newListItem = bodyCopia.insertListItem(index++, listItem);
        newListItem.setGlyphType(glyphType);
      }
    }
    
    bodyCopia.replaceText('{{CUERPO}}', '');
    copiaDoc.saveAndClose();
    
    registrarArchivoGenerado(usuarioEmail, nombreFinal, docId);
    return { url: copiaDocFile.getUrl() };
  } catch (e) { 
    return { error: e.message }; 
  }
}