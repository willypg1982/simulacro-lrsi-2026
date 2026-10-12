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

    // ACCIÓN: OBTENER EXAMEN POR CÓDIGO O NOMBRE (PARA DESCARGA POSTERIOR)
    if (action === "getExamByCode") {
      const targetQuery = String(e.parameter.codigo || e.parameter.query || "").trim().toLowerCase();
      const values = sheet.getDataRange().getDisplayValues();

      if (!targetQuery || values.length <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ 
          status: "not_found", 
          message: "No se proporcionó un código o nombre válido o la hoja no tiene registros." 
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // Detectar índice de columnas
      let codeColIdx = 14; // Columna 15 por defecto
      let studentColIdx = 1; // Columna 2 por defecto
      if (values.length > 0) {
        const headers = values[0].map(h => String(h).trim().toLowerCase());
        const foundCode = headers.findIndex(h => /c[oó]digo|clave/.test(h));
        if (foundCode !== -1) codeColIdx = foundCode;
        const foundStudent = headers.findIndex(h => /estudiante|nombre/.test(h));
        if (foundStudent !== -1) studentColIdx = foundStudent;
      }

      const matches = [];

      // Buscar de la fila más reciente a la más antigua
      for (let r = values.length - 1; r >= 1; r--) {
        const row = values[r];
        const rowCode = String(row[codeColIdx] || "").trim().toLowerCase();
        const rowStudent = String(row[studentColIdx] || "").trim().toLowerCase();

        if ((rowCode && rowCode === targetQuery) || (rowStudent && rowStudent.includes(targetQuery))) {
          const examObj = {
            fecha: row[0] || "",
            estudiante: row[studentColIdx] || "",
            seccion: row[2] || "",
            puntos: row[3] || "",
            totalPuntos: row[4] || 60,
            porcentaje: row[5] || "",
            nota: row[6] || "",
            condicion: row[7] || "",
            tiempo: row[8] || "",
            tiempoSegundos: row[9] || "",
            buenas: row[10] || row[3] || "",
            incorrectas: row[11] || "",
            preguntasIncorrectas: row[12] || "",
            respuestas: row[13] || "",
            codigo: row[codeColIdx] || ""
          };

          // Evitar duplicar el mismo estudiante si tiene múltiples filas: preferir la que tenga tiempo real > 0s
          const existingIdx = matches.findIndex(m => m.estudiante.toLowerCase().trim() === rowStudent);
          if (existingIdx === -1) {
            matches.push(examObj);
          } else {
            const existing = matches[existingIdx];
            if ((existing.tiempo === "0 s" || existing.tiempo === "0") && (examObj.tiempo && examObj.tiempo !== "0 s" && examObj.tiempo !== "0")) {
              matches[existingIdx] = examObj;
            }
          }
        }
      }

      if (matches.length > 0) {
        return ContentService.createTextOutput(JSON.stringify({ 
          status: "success", 
          exam: matches[0],
          matches: matches
        })).setMimeType(ContentService.MimeType.JSON);
      }

      return ContentService.createTextOutput(JSON.stringify({ 
        status: "not_found", 
        message: "No se encontró ningún examen registrado con la clave o nombre ingresado." 
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

    // ACCIÓN: DEPURAR DUPLICADOS AUTOMÁTICAMENTE (CONSERVAR SOLO 1 REGISTRO OFICIAL POR ALUMNO)
    if (action === "cleanDuplicates") {
      const msg = limpiarDuplicadosConservarOficial();
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        action: "cleanDuplicates", 
        message: msg 
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
    // Si viene un reenvío accidental con 0 segundos y el estudiante ya tiene su intento oficial con tiempo real, evitar duplicarlo
    if (data.tiempoSegundos === 0 || data.tiempo === "0 s" || data.tiempo === "0") {
      const values = sheet.getDataRange().getDisplayValues();
      const studentName = String(data.estudiante || "").trim().toLowerCase();
      let alreadyExists = false;
      for (let r = values.length - 1; r >= 1; r--) {
        const rowStudent = String(values[r][1] || "").trim().toLowerCase();
        const rowTime = String(values[r][8] || "").trim();
        if (rowStudent === studentName && rowTime !== "0 s" && rowTime !== "0" && rowTime !== "--") {
          alreadyExists = true;
          break;
        }
      }
      if (alreadyExists) {
        return ContentService.createTextOutput(JSON.stringify({ 
          status: "ignored", 
          action: "duplicate_prevented", 
          message: "Se conservó la entrega oficial previa del estudiante con tiempo real." 
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

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

/**
 * ============================================================================
 * FUNCIÓN MANUAL / AUTOMÁTICA DE LIMPIEZA DE DUPLICADOS:
 * ============================================================================
 * Puedes ejecutar esta función directamente desde Apps Script seleccionándola 
 * arriba en el menú desplegable y haciendo clic en "Ejecutar", o activándola 
 * desde el botón de la aplicación web.
 * 
 * ¿Qué hace exactamente?
 * 1. Analiza todos los registros existentes en la hoja de cálculo.
 * 2. Identifica a cada estudiante único.
 * 3. Conserva estrictamente el intento oficial original de cada estudiante (el que tiene tiempo real > 0s).
 * 4. Elimina todas las filas repetidas que fueron reintentos con "0 s" o entregas dobles.
 * 5. Deja la hoja con exactamente 1 registro por estudiante, limpio y organizado.
 */
function limpiarDuplicadosConservarOficial() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return "No se encontró la hoja de cálculo activa.";
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getDisplayValues();

  if (data.length <= 1) return "No hay datos suficientes para depurar.";

  const studentMap = {};
  const rowsToDelete = [];

  for (let r = 1; r < data.length; r++) {
    const student = String(data[r][1] || "").trim().toLowerCase();
    const tiempo = String(data[r][8] || "").trim();
    const rowNumber = r + 1; // 1-index de fila en Google Sheets

    if (!student) continue;

    if (!studentMap[student]) {
      studentMap[student] = {
        rowNumber: rowNumber,
        tiempo: tiempo,
        estudiante: data[r][1]
      };
    } else {
      const existing = studentMap[student];
      if ((tiempo === "0 s" || tiempo === "0") && (existing.tiempo !== "0 s" && existing.tiempo !== "0")) {
        rowsToDelete.push(rowNumber);
      } else if ((existing.tiempo === "0 s" || existing.tiempo === "0") && (tiempo !== "0 s" && tiempo !== "0")) {
        rowsToDelete.push(existing.rowNumber);
        studentMap[student] = { rowNumber, tiempo, estudiante: data[r][1] };
      } else {
        rowsToDelete.push(rowNumber);
      }
    }
  }

  // Eliminar de abajo hacia arriba para mantener estables los números de fila superiores
  rowsToDelete.sort((a, b) => b - a);
  rowsToDelete.forEach(rowNum => {
    sheet.deleteRow(rowNum);
  });

  const msg = `Depuración exitosa: se eliminaron ${rowsToDelete.length} registros duplicados con 0s y se conservaron ${Object.keys(studentMap).length} registros oficiales únicos.`;
  Logger.log(msg);
  return msg;
}

