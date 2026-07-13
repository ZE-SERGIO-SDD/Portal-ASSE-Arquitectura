/**
 * Obtiene la lista de las últimas 3 actividades del usuario activo desde "LOG_ARCHIVOS".
 * @returns {Array} Arreglo de objetos con los detalles de cada actividad.
 */
function obtenerDatosActividadReciente() {
  try {
    var idSpreadsheet = '13Dyg_fIxrbsw--0r1k-jLV7TuhPdWOYAr__OgJ06eT8';
    var nombreHoja = 'LOG_ARCHIVOS';
    
    // Obtenemos el correo del usuario que está usando la aplicación en este momento
    var emailUsuarioActivo = Session.getActiveUser().getEmail();
    
    // Se conecta a la planilla y busca la pestaña específica
    var ss = SpreadsheetApp.openById(idSpreadsheet);
    var hoja = ss.getSheetByName(nombreHoja);
    
    if (!hoja) {
      return { error: "No se encontró la pestaña '" + nombreHoja + "'." };
    }
    
    // Obtiene toda la tabla de datos
    var datos = hoja.getDataRange().getValues();
    
    // Si la hoja solo tiene los encabezados o está vacía, no hay nada que mostrar
    if (datos.length <= 1) {
      return [];
    }
    
    // Quitamos la primera fila correspondiente a los encabezados
    datos.shift();
    
    // Filtramos los datos para quedarnos SOLO con las filas donde el autor (Columna A) 
    // coincide con el correo del usuario activo.
    var datosDelUsuario = datos.filter(function(fila) {
      var autor = fila[0] ? fila[0].toString().trim() : "";
      return autor === emailUsuarioActivo;
    });
    
    // Si el usuario aún no tiene actividades registradas
    if (datosDelUsuario.length === 0) {
      return [];
    }
    
    // Tomamos las últimas 3 filas del usuario (las más recientes) y las invertimos 
    // para que la última acción aparezca arriba en el panel.
    var ultimasActividades = datosDelUsuario.slice(-3).reverse();
    
    var actividadesFormateadas = ultimasActividades.map(function(fila, index) {
      // Asignación de columnas según la estructura de LOG_ARCHIVOS:
      // A (0) = Usuario
      // B (1) = Nombre del archivo creado
      // D (3) = Fecha
      
      var archivo = fila[1] ? fila[1].toString() : "Archivo sin nombre";
      var fechaCruda = fila[3]; 
      
      var fechaFormateada = formatearFecha(fechaCruda);
      
      // Armamos el título. Ya no es necesario mostrar el correo porque sabemos que es suyo.
      var tituloFinal = archivo;
      
      // Asignación inteligente de iconos según el nombre o extensión del archivo
      var icono = "📝"; 
      var archivoMinuscula = archivo.toLowerCase();
      
      if (archivoMinuscula.indexOf("pdf") > -1) icono = "📄";
      else if (archivoMinuscula.indexOf("dwg") > -1 || archivoMinuscula.indexOf("autocad") > -1) icono = "📐";
      else if (archivoMinuscula.indexOf("rvt") > -1 || archivoMinuscula.indexOf("revit") > -1) icono = "🏗️";
      else if (archivoMinuscula.indexOf("doc") > -1) icono = "📑";
      
      return {
        id: index,
        tipo: "log",
        titulo: tituloFinal,
        fecha: fechaFormateada,
        icono: icono
      };
    });
    
    return actividadesFormateadas;
    
  } catch (error) {
    return { error: "Error de lectura en LOG_ARCHIVOS: " + error.toString() };
  }
}

/**
 * Función auxiliar para darle un formato limpio y legible a la fecha
 */
function formatearFecha(fecha) {
  if (!fecha) return "Fecha desconocida";
  
  if (fecha instanceof Date) {
    var dia = fecha.getDate();
    var mes = fecha.getMonth() + 1;
    var anio = fecha.getFullYear();
    var horas = fecha.getHours();
    var minutos = fecha.getMinutes();
    
    // Agregamos un cero inicial si los números son de un solo dígito
    if (dia < 10) dia = '0' + dia;
    if (mes < 10) mes = '0' + mes;
    if (horas < 10) horas = '0' + horas;
    if (minutos < 10) minutos = '0' + minutos;
    
    return dia + "/" + mes + "/" + anio + " " + horas + ":" + minutos;
  }
  return fecha.toString();
}