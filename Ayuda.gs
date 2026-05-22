// =====================================================================
// MÓDULO: AYUDA (Base de datos central de textos)
// =====================================================================

function obtenerTextosAyuda() {
  return {
    memorias: {
      titulo: "Generador de Memorias",
      icono: "📝",
      texto: "<p><strong>¿Qué hace esta herramienta?</strong><br>Te permite ensamblar rápidamente una Memoria Descriptiva oficial. Extrae los rubros constructivos desde la base de datos de Excel y los estructura automáticamente en un documento de Google Docs membretado.</p><p><strong>¿Cómo utilizarlo?</strong><br>1. Selecciona la disciplina (ej. Arquitectura, Sanitaria).<br>2. Despliega los niveles usando el botón <code>+</code> o el botón 'Expandir Niveles'.<br>3. Marca las casillas de los rubros que necesitas incluir. Si marcas un rubro 'hijo', el sistema marcará automáticamente su rubro 'padre' para mantener la jerarquía.<br>4. Haz clic en 'Previsualizar Memoria' para ver cómo quedará el documento final antes de generarlo.</p><p><strong>Guardado en la Nube (Borradores)</strong><br>Si estás armando una memoria muy larga y no terminaste, usa el botón <strong>💾 GUARDAR</strong>. Le pondrás un nombre y tu selección de casillas quedará guardada en la nube. Otro día, o desde otra computadora, usas <strong>📂 CARGAR</strong>, eliges tu borrador, y las casillas volverán a marcarse solas.</p>"
    },
    documentos: {
      titulo: "Documentos internos",
      icono: "📄",
      texto: "<p><strong>¿Qué hace esta herramienta?</strong><br>Aplica automáticamente la lámina membretada oficial de la oficina a los documentos base (como notas, informes o actas), ahorrándote el trabajo de formato manual.</p><p><strong>¿Cómo utilizarlo?</strong><br>1. Busca en la lista el documento base que necesitas y haz clic en <strong>PREVISUALIZAR ➜</strong>.<br>2. Revisa que sea el contenido correcto en la ventana que se abre.<br>3. Haz clic en <strong>CONFIRMAR Y GENERAR</strong>. El sistema creará una copia nueva uniendo ese texto con el membrete oficial.</p><p><strong>¿Dónde encuentro el resultado?</strong><br>El archivo generado, listo para descargar o imprimir, aparecerá automáticamente en tu sección <strong>Mis Archivos (Historial)</strong>.</p>"
    },
    proyectos: {
      titulo: "Gestión de Proyectos",
      icono: "🏗️",
      texto: "<p><strong>¿Qué hace esta herramienta?</strong><br>Centraliza la información de todas las obras y anteproyectos de ASSE. Permite registrar nuevas obras, asignar responsables, determinar el estado de avance y mantener un historial unificado.</p><p><strong>¿Cómo utilizarlo?</strong><br>1. <strong>Nuevo Proyecto:</strong> Llena el formulario en cascada. Selecciona primero el departamento para poder elegir el centro correspondiente. Asigna roles y guarda.<br>2. <strong>Proyectos en curso:</strong> Visualiza las obras activas. <br>3. <strong>Historial:</strong> Consulta el archivo de obras finalizadas o dadas de baja.</p><p><strong>Datos actualizados</strong><br>Si notas que falta un Arquitecto en la lista o un Estado, utiliza el botón de recargar (🔄) arriba a la derecha. Si el dato sigue faltando, debe agregarse primero a la planilla central de Excel.</p>"
    },
    plantillas: {
      titulo: "Plantillas (Templates)",
      icono: "📂",
      texto: "<p><strong>¿Qué es este repositorio?</strong><br>Es la biblioteca central de la oficina. Aquí encontrarás las plantillas, (templates) de AutoCAD (DWG), y de Revit (RTE).</p><p><strong>¿Cómo utilizarlo?</strong><br>Navega por las carpetas haciendo clic en <strong>ABRIR ➜</strong>. Una vez que encuentres el archivo que necesitas, haz clic en <strong>DESCARGAR</strong> para guardar una copia en tu equipo.</p><p><strong>Importante</strong><br>Estos archivos son de solo lectura para mantener la integridad de los estándares de la oficina. Descárgalos siempre para trabajar sobre ellos de forma local.</p>"
    },
    mis_archivos: {
      titulo: "Mis Archivos (Historial)",
      icono: "🕒",
      texto: "<p><strong>¿Qué es \"Mis Archivos (Historial)\"?</strong><br>Esta es tu bandeja de salida personal. Aquí se agrupan automáticamente todos los documentos oficiales, memorias y archivos que tú generes a través de las distintas herramientas del portal.</p><p><strong>Privacidad y Limpieza</strong><br>Esta lista es estrictamente privada; nadie más puede ver los archivos que has creado. Para no saturar el servidor, los documentos aquí listados son temporales y el sistema los <strong>elimina automáticamente pasados los 3 días</strong>.</p><p><strong>Recomendación</strong><br>Usa esta sección para revisar tu trabajo reciente, descargar los archivos definitivos a tu computadora en formato PDF o Word, o compartir los enlaces finales con tu equipo.</p>"
    }
  };
}