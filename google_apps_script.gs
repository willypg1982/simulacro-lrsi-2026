/**
 * ============================================================================
 * SIMULACRO DE MATEMÁTICAS 2026 - LICEO RURAL SAN ISIDRO
 * Google Apps Script para registro de notas, tiempo y respuestas por estudiante
 * Docente: Profesor William Pineda González
 * ============================================================================
 * 
 * INSTRUCCIONES DE USO:
 * 1. En tu Hoja de Cálculo de Google, ve a: Extensiones > Apps Script.
 * 2. Borra todo el código que haya allí y pega este archivo completo.
 * 3. Haz clic en "Implementar" (botón azul) > "Nueva implementación".
 * 4. Tipo: "Aplicación web".
 * 5. Ejecutar como: "Yo (tu cuenta de correo)".
 * 6. Quién tiene acceso: "Cualquier persona" (Anyone).
 * 7. Haz clic en "Implementar" y autoriza los permisos.
 * 8. Copia la URL de la aplicación web generada (termina en /exec) y asegúrate de que sea la misma en index.html.
 */

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        message: "No se encontró la hoja de cálculo activa." 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = ss.getActiveSheet();
    const data = sheet.getDataRange().getDisplayValues();
    
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      data: data 
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // Si la hoja está en blanco, crear encabezados automáticos con las 14 columnas oficiales
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Marca Temporal",
        "Estudiante",
        "Sección",
        "Puntos",
        "Total Puntos",
        "Porcentaje",
        "Nota Final",
        "Condición",
        "Tiempo Empleado",
        "Tiempo (Segundos)",
        "Correctas",
        "Incorrectas",
        "Preguntas Incorrectas",
        "Respuestas del Estudiante"
      ]);
      const headerRange = sheet.getRange(1, 1, 1, 14);
      headerRange.setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
    
    const data = JSON.parse(e.postData.contents);
    
    sheet.appendRow([
      data.fecha || new Date().toLocaleString("es-CR"),
      data.estudiante || "Estudiante",
      data.seccion || "No indicada",
      data.puntos,
      data.totalPuntos,
      data.porcentaje + "%",
      data.nota,
      data.condicion,
      data.tiempo,
      data.tiempoSegundos,
      data.correctas,
      data.incorrectas,
      data.preguntasIncorrectas || "Ninguna",
      data.respuestasStr || JSON.stringify(data.respuestas || {})
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
