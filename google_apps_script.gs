/**
 * ============================================================================
 * SIMULACRO DE MATEMÁTICAS 2026 - LICEO RURAL SAN ISIDRO
 * Google Apps Script para registro, consulta, descarga y gestión de resultados
 * Docente: Profesor William Pineda González
 * ============================================================================
 * 
 * INSTRUCCIONES DE ACTUALIZACIÓN:
 * 1. En tu Hoja de Cálculo de Google, ve a: Extensiones > Apps Script.
 * 2. Borra TODO el código anterior y pega este archivo completo.
 * 3. Haz clic en "Implementar" (botón azul arriba a la derecha) > "Gestionar implementaciones"
 *    o "Nueva implementación" > Versión nueva > "Implementar".
 * 4. Tipo: "Aplicación web".
 * 5. Ejecutar como: "Yo (tu cuenta de correo)".
 * 6. Quién tiene acceso: "Cualquier persona" (Anyone).
 * 7. Autoriza los permisos si te los solicita.
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

    // ACCIÓN: OBTENER EXAMEN POR CÓDIGO O CLAVE ÚNICA (PARA DESCARGA POSTERIOR)
    if (action === "getExamByCode") {
      const targetCode = String(e.parameter.codigo || "").trim().toLowerCase();
      const values = sheet.getDataRange().getDisplayValues();

      if (!targetCode || values.length <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ 
          status: "not_found", 
          message: "No se proporcionó un código válido o la hoja no tiene registros." 
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // Detectar índice de la columna de código
      let codeColIdx = 14; // Columna 15 por defecto
      if (values.length > 0) {
        const headers = values[0].map(h => String(h).trim().toLowerCase());
        const found = headers.findIndex(h => /c[oó]digo|clave/.test(h));
        if (found !== -1) codeColIdx = found;
      }

      // Buscar de la fila más reciente a la más antigua
      for (let r = values.length - 1; r >= 1; r--) {
        const row = values[r];
        const rowCode = String(row[codeColIdx] || "").trim().toLowerCase();
        if (rowCode && rowCode === targetCode) {
          return ContentService.createTextOutput(JSON.stringify({ 
            status: "success", 
            exam: {
              fecha: row[0] || "",
              estudiante: row[1] || "",
              seccion: row[2] || "",
              puntos: row[3] || "",
              totalPuntos: row[4] || 60,
              porcentaje: row[5] || "",
              nota: row[6] || "",
              condicion: row[7] || "",
              tiempo: row[8] || "",
              buenas: row[10] || row[3] || "",
              incorrectas: row[11] || "",
              preguntasIncorrectas: row[12] || "",
              respuestas: row[13] || "",
              codigo: row[codeColIdx] || ""
            }
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      return ContentService.createTextOutput(JSON.stringify({ 
        status: "not_found", 
        message: "No se encontró ningún examen registrado con la clave ingresada." 
      })).setMimeType(ContentService.MimeType.JSON);
    }

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
    
    // Si la hoja está en blanco, crear encabezados automáticos con las 15 columnas oficiales
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
        "Respuestas del Estudiante",
        "Código Estudiante"
      ]);
      const headerRange = sheet.getRange(1, 1, 1, 15);
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
    
    // GUARDAR NUEVO RESULTADO (ACCIÓN POR DEFECTO CON CÓDIGO)
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
      data.respuestasStr || JSON.stringify(data.respuestas || {}),
      data.codigo || ""
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
