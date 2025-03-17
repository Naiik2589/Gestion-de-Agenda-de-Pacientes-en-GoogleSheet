var FORMULAS = {}; // Almacena fórmulas originales

function guardarEstadoInicial(sheet) {
  if (!sheet) return;
  try {
    var sheetName = sheet.getName();
    var range = sheet.getDataRange();
    if (range.isBlank()) return;
    
    FORMULAS[sheetName] = range.getFormulas();
    Logger.log("Estado inicial guardado para la hoja: " + sheetName);
  } catch (err) {
    Logger.log("Error en guardarEstadoInicial: " + err);
  }
}

function onEdit(e) {
  if (!e || !e.range) return;
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;

    // Convertir valores a mayúsculas
    var values = range.getValues().map(row => row.map(cell => typeof cell === 'string' ? cell.toUpperCase() : cell));
    range.setValues(values);
    Logger.log("Valores convertidos a mayúsculas en: " + range.getA1Notation());

    // Pegar solo valores (sin formato ni colores)
    range.setValues(range.getValues());
    Logger.log("Valores pegados sin formato en: " + range.getA1Notation());

    // Aplicar formato de fuente Arial tamaño 12
    range.setFontFamily("Arial");
    range.setFontSize(12);
    Logger.log("Formato de texto estandarizado a Arial 12 en: " + range.getA1Notation());

    // Mantener alineación horizontal y vertical centrada
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
    Logger.log("Alineación centrada aplicada en: " + range.getA1Notation());
    
    // Restaurar fórmulas originales
    var sheetName = sheet.getName();
    if (FORMULAS[sheetName]) {
      var formulas = FORMULAS[sheetName];
      for (var r = 0; r < range.getNumRows(); r++) {
        for (var c = 0; c < range.getNumColumns(); c++) {
          var formula = formulas[range.getRow() - 1 + r][range.getColumn() - 1 + c];
          if (formula) {
            range.getCell(r + 1, c + 1).setFormula(formula);
            Logger.log("Fórmula restaurada en: " + range.getCell(r + 1, c + 1).getA1Notation());
          }
        }
      }
    }
    
    // Descombinar solo si la edición ocurrió en un rango combinado, desde fila 3 y columnas C a J
    if (range.getRow() >= 3 && range.getColumn() >= 3 && range.getColumn() <= 10) {
      var mergedRanges = range.getMergedRanges();
      if (mergedRanges.length > 0) {
        mergedRanges.forEach(mergedRange => {
          mergedRange.breakApart();
          Logger.log("Celdas descombinadas en: " + mergedRange.getA1Notation());
        });
      }
    }
  } catch (err) {
    Logger.log("Error en onEdit: " + err);
  }
}

function inicializar() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(sheet => guardarEstadoInicial(sheet));
  Logger.log("Inicialización completada para todas las hojas.");
}
