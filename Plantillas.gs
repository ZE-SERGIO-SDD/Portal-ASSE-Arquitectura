// =====================================================================
// MÓDULO: PLANTILLAS (Lógica de servidor)
// =====================================================================

function obtenerPlantillas(forzarRecarga) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'arbolPlantillas_v1';
    
    // 1. Si NO estamos forzando la recarga, buscamos en la caché primero
    if (!forzarRecarga) {
      var cachedData = cache.get(cacheKey);
      if (cachedData) {
        // La caché guarda texto, así que lo volvemos a convertir en objeto
        return JSON.parse(cachedData); 
      }
    }
    
    // 2. Si no hay caché o el usuario forzó la recarga, vamos a Drive
    var folderRaiz = DriveApp.getFolderById(FOLDER_PLANTILLAS_ID);
    var arbolNuevo = construirArbol(folderRaiz, 0);
    
    // 3. Guardamos el nuevo árbol en caché por 6 horas (21600 segundos)
    // Convertimos el objeto a texto porque la caché solo acepta strings
    cache.put(cacheKey, JSON.stringify(arbolNuevo), 21600);
    
    return arbolNuevo;
    
  } catch (e) { 
    return { error: e.message }; 
  }
}

function construirArbol(folder, profundidad) {
  if (profundidad > 3) return { id: folder.getId(), nombre: folder.getName(), tipo: 'carpeta', elementos: [] };
  var resultado = { id: folder.getId(), nombre: folder.getName(), tipo: 'carpeta', elementos: [] };
  try {
    var subFolders = folder.getFolders();
    while (subFolders.hasNext()) { 
      resultado.elementos.push(construirArbol(subFolders.next(), profundidad + 1)); 
    }
    var files = folder.getFiles();
    while (files.hasNext()) { 
      var f = files.next(); 
      resultado.elementos.push({ id: f.getId(), nombre: f.getName(), tipo: 'archivo', url: f.getUrl() }); 
    }
    resultado.elementos.sort((a,b) => a.tipo === b.tipo ? a.nombre.localeCompare(b.nombre) : (a.tipo === 'carpeta' ? -1 : 1));
  } catch (e) { 
    resultado.elementos.push({ nombre: "⚠️ Error", tipo: 'archivo', url: '#' }); 
  }
  return resultado;
}