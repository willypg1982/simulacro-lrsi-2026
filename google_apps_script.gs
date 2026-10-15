/**
 * ============================================================================
 * SIMULACRO DE MATEMÁTICAS 2026 - LICEO RURAL SAN ISIDRO
 * Google Apps Script para registro, consulta, descarga y gestión de resultados
 * Soporte Multi-Pestaña: Pestaña "Oficial" y Pestaña "Práctica"
 * Docente: Profesor William Pineda González
 * ============================================================================
 * 
 * INSTRUCCIONES DE ACTUALIZACIÓN EN GOOGLE SHEETS:
 * 1. En tu Hoja de Cálculo de Google, ve a: Extensiones > Apps Script.
 * 2. Borra TODO el código anterior y pega este archivo completo.
 * 3. Haz clic en "Implementar" (botón azul arriba a la derecha) > "Gestionar implementaciones"
 *    > Editar (ícono lápiz) > Versión: "Nueva versión" > Haz clic en "Implementar".
 * 4. Tipo: "Aplicación web".
 * 5. Ejecutar como: "Yo (tu cuenta de correo)".
 * 6. Quién tiene acceso: "Cualquier persona" (Anyone).
 * 7. Listo. ¡No cambia tu URL de Apps Script!
 */

/**
 * Obtiene o crea la pestaña de cálculo correspondiente ("Oficial" o "Práctica").
 * Si la hoja original se llamaba "Hoja 1", la renombra automáticamente a "Oficial"
 * para conservar intactos todos los resultados previos.
 */
function getTargetSheet(ss, tabName) {
  const isPractica = /pr[aá]ctica/i.test(String(tabName || ""));
  const targetName = isPractica ? "Práctica" : "Oficial";
  
  let sheet = ss.getSheetByName(targetName);
  
  // Si no se encuentra "Oficial", buscar si hay otra hoja como "Hoja 1" o la primera hoja y renombrarla
  if (!isPractica && !sheet) {
    const sheets = ss.getSheets();
    if (sheets.length > 0) {
      if (!/pr[aá]ctica/i.test(sheets[0].getName())) {
        sheet = sheets[0];
        sheet.setName("Oficial");
      }
    }
  }
  
  // Si no existe (por ejemplo "Práctica" en su primera ejecución), crearla automáticamente
  if (!sheet) {
    sheet = ss.insertSheet(targetName);
  }
  
  // Si la hoja no tiene encabezados, crearlos automáticamente
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
    const headerColor = isPractica ? "#0f766e" : "#1e3a8a"; // Verde para Práctica, Azul para Oficial
    headerRange.setBackground(headerColor).setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        message: "No se encontró la hoja de cálculo activa." 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const tabParam = e && e.parameter ? (e.parameter.tab || e.parameter.sheet || "") : "";
    const action = e && e.parameter ? e.parameter.action : null;
    const sheet = getTargetSheet(ss, tabParam);

    // ACCIÓN: OBTENER EXAMEN POR CÓDIGO O NOMBRE (PARA DESCARGA POSTERIOR)
    if (action === "getExamByCode") {
      const targetQuery = String(e.parameter.codigo || e.parameter.query || "").trim().toLowerCase();
      
      // Buscar primero en la pestaña solicitada y, si no se encuentra, en todas las pestañas
      let sheetsToSearch = [sheet];
      const allSheets = ss.getSheets();
      allSheets.forEach(s => {
        if (!sheetsToSearch.some(existing => existing.getName() === s.getName())) {
          sheetsToSearch.push(s);
        }
      });

      const matches = [];

      for (let s = 0; s < sheetsToSearch.length; s++) {
        const currentS = sheetsToSearch[s];
        const values = currentS.getDataRange().getDisplayValues();
        if (values.length <= 1) continue;

        let codeColIdx = 14;
        let studentColIdx = 1;
        if (values.length > 0) {
          const headers = values[0].map(h => String(h).trim().toLowerCase());
          const foundCode = headers.findIndex(h => /c[oó]digo|clave/.test(h));
          if (foundCode !== -1) codeColIdx = foundCode;
          const foundStudent = headers.findIndex(h => /estudiante|nombre/.test(h));
          if (foundStudent !== -1) studentColIdx = foundStudent;
        }

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
              codigo: row[codeColIdx] || "",
              origenPestana: currentS.getName()
            };

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
        if (matches.length > 0 && targetQuery) break;
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
          break;
        }
      }

      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        action: "delete", 
        deletedCount: deleted,
        tab: sheet.getName()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ACCIÓN: VACIAR TODOS LOS REGISTROS VÍA GET
    if (action === "clearAll") {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        action: "clearAll",
        tab: sheet.getName()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ACCIÓN: DEPURAR DUPLICADOS
    if (action === "cleanDuplicates") {
      const msg = limpiarDuplicadosConservarOficial(sheet);
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        action: "cleanDuplicates", 
        message: msg,
        tab: sheet.getName()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // CONSULTA ESTÁNDAR: OBTENER TODAS LAS FILAS DE LA PESTAÑA SOLICITADA
    const data = sheet.getDataRange().getDisplayValues();
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      tab: sheet.getName(),
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
    const data = JSON.parse(e.postData.contents);
    const tabParam = data.tab || data.sheet || (data.isPractica ? "Práctica" : "Oficial");
    const sheet = getTargetSheet(ss, tabParam);
    const isPractica = /pr[aá]ctica/i.test(sheet.getName());

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
        deletedCount: deleted,
        tab: sheet.getName()
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
        action: "clearAll",
        tab: sheet.getName()
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // GUARDAR NUEVO RESULTADO:
    // Si es "Oficial": proteger contra reenvíos accidentales con 0s si ya existe intento con tiempo real
    if (!isPractica) {
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
            message: "Se conservó la entrega oficial previa del estudiante con tiempo real.",
            tab: sheet.getName()
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
    // Si es "Práctica": los estudiantes pueden enviar todas las veces que quieran sin restricción

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
    
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      action: "save",
      tab: sheet.getName()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function limpiarDuplicadosConservarOficial(targetSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return "No se encontró la hoja de cálculo activa.";
  const sheet = targetSheet || ss.getActiveSheet();
  const data = sheet.getDataRange().getDisplayValues();

  if (data.length <= 1) return "No hay datos suficientes para depurar.";

  const studentMap = {};
  const rowsToDelete = [];

  for (let r = 1; r < data.length; r++) {
    const student = String(data[r][1] || "").trim().toLowerCase();
    const tiempo = String(data[r][8] || "").trim();
    const rowNumber = r + 1;

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

  rowsToDelete.sort((a, b) => b - a);
  rowsToDelete.forEach(rowNum => {
    sheet.deleteRow(rowNum);
  });

  const msg = `Depuración en '${sheet.getName()}': se eliminaron ${rowsToDelete.length} duplicados y se conservaron ${Object.keys(studentMap).length} registros únicos.`;
  Logger.log(msg);
  return msg;
}
