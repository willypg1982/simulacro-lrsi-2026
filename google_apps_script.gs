/**
 * ============================================================================
 * SIMULACRO DE MATEMÁTICAS 2026 - LICEO RURAL SAN ISIDRO
 * Google Apps Script para registro, consulta y eliminación de resultados
 * Docente: Profesor William Pineda González
 * ============================================================================
 * 
 * INSTRUCCIONES DE ACTUALIZACIÓN:
 * 1. En tu Hoja de Cálculo de Google, ve a: Extensiones > Apps Script.
 * 2. Borra TODO el código anterior y pega este archivo completo.
 * 3. Haz clic en "Implementar" (botón azul arriba a la derecha) > "Nueva implementación".
 * 4. Tipo: "Aplicación web".
 * 5. Ejecutar como: "Yo (tu cuenta de correo)".
 * 6. Quién tiene acceso: "Cualquier persona" (Anyone).
 * 7. Haz clic en "Implementar" y autoriza los permisos.
 * 8. Si la URL generada cambia, verifícala en index.html.
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
    const action = e && e.parameter ? e.parameter.action : null;

    // ACCIÓN: ELIMINAR UN REGISTRO VÍA GET
    if (action === "delete") {
      const targetStudent = (e.parameter.estudiante || "").trim().toLowerCase();
      const targetDate = (e.parameter.fecha || "").trim();
      const values = sheet.getDataRange().getDisplayValues();
      let deleted = 0;

      for (let r = values.length - 1; r >= 1; r--) {
        const row = values[r];
        const studentName = String(row[1] || "").trim().toLowerCase();
        let match = (studentName === targetStudent);
        if (match && targetDate && targetDate !== "--") {
          const rowDate = String(row[0] || "").trim();
          match = (rowDate === targetDate);
        }
        if (match) {
          sheet.deleteRow(r + 1);
          deleted++;
          break; // Eliminar la fila encontrada
        }
      }

      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        action: "delete", 
        deletedCount: deleted 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ACCIÓN: VACIAR TODOS LOS REGISTROS VÍA GET (conservando encabezados)
    if (action === "clearAll") {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        action: "clearAll" 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // CONSULTA ESTÁNDAR: OBTENER TODAS LAS FILAS EN VIVO
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

    // ACCIÓN: ELIMINAR VÍA POST
    if (data.action === "delete") {
      const targetStudent = String(data.estudiante || "").trim().toLowerCase();
      const targetDate = String(data.fecha || "").trim();
      const values = sheet.getDataRange().getDisplayValues();
      let deleted = 0;

      for (let r = values.length - 1; r >= 1; r--) {
        const row = values[r];
        const studentName = String(row[1] || "").trim().toLowerCase();
        let match = (studentName === targetStudent);
        if (match && targetDate && targetDate !== "--") {
          const rowDate = String(row[0] || "").trim();
          match = (rowDate === targetDate);
        }
        if (match) {
          sheet.deleteRow(r + 1);
          deleted++;
          break;
        }
      }

      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        action: "delete", 
        deletedCount: deleted 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ACCIÓN: VACIAR HOJA VÍA POST
    if (data.action === "clearAll") {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        action: "clearAll" 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // GUARDAR NUEVO RESULTADO (ACCIÓN POR DEFECTO)
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
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", action: "save" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
