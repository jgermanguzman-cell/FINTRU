function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row < 2) return;

  const config = obtenerConfigEtiquetas_(sheet);

  if (!config) return;

  const isNaturalezaEdit = col === config.naturalezaCol;
  const isCategoriaEdit = col === config.categoriaCol;
  const isEtiquetaAdicionalEdit =
    config.etiquetasAdicionalesCol && col === config.etiquetasAdicionalesCol;

  if (!isNaturalezaEdit && !isCategoriaEdit && !isEtiquetaAdicionalEdit) return;

  if (isNaturalezaEdit || isCategoriaEdit) {
    aplicarValidacionEtiquetaRelacional_(sheet, row, config, true);
    return;
  }

  if (isEtiquetaAdicionalEdit) {
    manejarEtiquetasAdicionales_(e);
  }
}

function repararEtiquetasRelacionales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const hojas = ["Presupuesto", "Movimientos"];
  let filasProcesadas = 0;

  hojas.forEach(nombreHoja => {
    const sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) return;

    const config = obtenerConfigEtiquetas_(sheet);
    if (!config) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    for (let row = 2; row <= lastRow; row++) {
      const naturaleza = sheet.getRange(row, config.naturalezaCol).getValue();
      const categoria = sheet.getRange(row, config.categoriaCol).getValue();

      if (!naturaleza && !categoria) continue;

      aplicarValidacionEtiquetaRelacional_(sheet, row, config, false);
      filasProcesadas++;
    }
  });

  ui.alert("Etiquetas relacionales reparadas.\n\nFilas procesadas: " + filasProcesadas);
}

function refrescarEtiquetaRelacionalEnFila_(sheet, row, configManual) {
  const config = configManual || obtenerConfigEtiquetas_(sheet);

  if (!config) return;

  aplicarValidacionEtiquetaRelacional_(sheet, row, config, false);
}

function aplicarValidacionEtiquetaRelacional_(sheet, row, config, limpiarValores) {
  const naturaleza = sheet.getRange(row, config.naturalezaCol).getValue();
  const categoria = sheet.getRange(row, config.categoriaCol).getValue();

  const etiquetaCell = sheet.getRange(row, config.etiquetaCol);
  const etiquetaActual = etiquetaCell.getValue();

  etiquetaCell.clearDataValidations();

  if (limpiarValores) {
    etiquetaCell.clearContent();
  }

  if (config.etiquetasAdicionalesCol) {
    const adicionalesCell = sheet.getRange(row, config.etiquetasAdicionalesCol);
    const adicionalesActual = adicionalesCell.getValue();

    adicionalesCell.clearDataValidations();

    if (limpiarValores) {
      adicionalesCell.clearContent();
    }

    const etiquetas = obtenerEtiquetas_(naturaleza, categoria);

    if (String(naturaleza).trim() === "Egreso" && etiquetas.length > 0) {
      const reglaAdicionales = SpreadsheetApp.newDataValidation()
        .requireValueInList(etiquetas, true)
        .setAllowInvalid(true)
        .build();

      adicionalesCell.setDataValidation(reglaAdicionales);

      if (!limpiarValores) {
        adicionalesCell.setValue(adicionalesActual);
      }
    }
  }

  const etiquetas = obtenerEtiquetas_(naturaleza, categoria);

  if (etiquetas.length === 0) {
    if (!limpiarValores) {
      etiquetaCell.setValue(etiquetaActual);
    }
    return;
  }

  const reglaPrincipal = SpreadsheetApp.newDataValidation()
    .requireValueInList(etiquetas, true)
    .setAllowInvalid(false)
    .build();

  etiquetaCell.setDataValidation(reglaPrincipal);

  if (!limpiarValores) {
    etiquetaCell.setValue(etiquetaActual);
  }
}

function obtenerConfigEtiquetas_(sheet) {
  const sheetName = sheet.getName();

  if (sheetName !== "Presupuesto" && sheetName !== "Movimientos") return null;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = crearMapaHeadersEtiquetas_(headers);

  const naturalezaCol = buscarColEtiqueta_(map, "Naturaleza");
  const categoriaCol =
    buscarColEtiqueta_(map, "Categoria") || buscarColEtiqueta_(map, "Categoría");
  const etiquetaCol = buscarColEtiqueta_(map, "Etiqueta");
  const etiquetasAdicionalesCol = buscarColEtiqueta_(map, "Etiquetas adicionales");

  if (!naturalezaCol || !categoriaCol || !etiquetaCol) return null;

  return {
    naturalezaCol,
    categoriaCol,
    etiquetaCol,
    etiquetasAdicionalesCol: sheetName === "Movimientos" ? etiquetasAdicionalesCol : null
  };
}

function obtenerEtiquetas_(naturaleza, categoria) {
  const naturalezaNormalizada = normalizarTextoFintru_(naturaleza);

  if (naturalezaNormalizada === "ingreso") {
    return obtenerEtiquetasPorNaturaleza_(naturaleza);
  }

  return obtenerEtiquetasPorCategoria_(categoria);
}

function obtenerEtiquetasPorCategoria_(categoria) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const catalogos = ss.getSheetByName("Catálogos");

  if (!catalogos) return [];

  const lastRow = catalogos.getLastRow();
  if (lastRow < 14) return [];

  const categoriaBuscada = normalizarTextoFintru_(categoria);

  const data = catalogos.getRange(14, 1, lastRow - 13, 3).getValues();

  return data
    .filter(row => {
      const categoriaCatalogo = normalizarTextoFintru_(row[0]);
      const estadoCatalogo = normalizarTextoFintru_(row[2]);

      return categoriaCatalogo === categoriaBuscada && estadoCatalogo === "activo";
    })
    .map(row => row[1])
    .filter(value => String(value).trim() !== "");
}

function obtenerEtiquetasPorNaturaleza_(naturaleza) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const catalogos = ss.getSheetByName("Catálogos");

  if (!catalogos) return [];

  const lastRow = catalogos.getLastRow();
  if (lastRow < 14) return [];

  const naturalezaBuscada = normalizarTextoFintru_(naturaleza);

  const data = catalogos.getRange(14, 5, lastRow - 13, 3).getValues();

  return data
    .filter(row => {
      const naturalezaCatalogo = normalizarTextoFintru_(row[0]);
      const estadoCatalogo = normalizarTextoFintru_(row[2]);

      return naturalezaCatalogo === naturalezaBuscada && estadoCatalogo === "activo";
    })
    .map(row => row[1])
    .filter(value => String(value).trim() !== "");
}

function manejarEtiquetasAdicionales_(e) {
  const nuevoValor = e.value;
  const valorAnterior = e.oldValue;

  if (!nuevoValor) return;
  if (!valorAnterior) return;

  const etiquetasActuales = valorAnterior
    .split(",")
    .map(item => item.trim())
    .filter(item => item !== "");

  if (etiquetasActuales.includes(nuevoValor)) {
    e.range.setValue(valorAnterior);
    return;
  }

  etiquetasActuales.push(nuevoValor);
  e.range.setValue(etiquetasActuales.join(", "));
}

function crearMapaHeadersEtiquetas_(headers) {
  const map = {};

  headers.forEach((header, index) => {
    const key = normalizarTextoFintru_(header);
    if (key) {
      map[key] = index + 1;
    }
  });

  return map;
}

function buscarColEtiqueta_(map, nombre) {
  return map[normalizarTextoFintru_(nombre)] || null;
}

function normalizarTextoFintru_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}
function aplicarValidacionEtiquetaMovimiento_(shMovimientos, targetRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCatalogos = ss.getSheetByName("Catálogos");

  if (!shCatalogos) {
    throw new Error("No existe la hoja Catálogos.");
  }

  const categoria = String(shMovimientos.getRange(targetRow, 12).getValue()).trim(); // L
  const naturaleza = String(shMovimientos.getRange(targetRow, 13).getValue()).trim(); // M
  const etiquetaCell = shMovimientos.getRange(targetRow, 14); // N

  etiquetaCell.clearDataValidations();

  if (!categoria || !naturaleza) return;

  const lastRow = shCatalogos.getLastRow();
  if (lastRow < 14) return;

  const data = shCatalogos.getRange(14, 1, lastRow - 13, 3).getValues();

  let etiquetas = [];

  if (naturaleza === "Ingreso") {
    etiquetas = data
      .filter(row =>
        String(row[0]).trim() === "N/A" &&
        String(row[2]).trim() === "Activo" &&
        ["Activo", "Pasivo", "Préstamo recibido"].includes(String(row[1]).trim())
      )
      .map(row => row[1]);
  } else {
    etiquetas = data
      .filter(row =>
        String(row[0]).trim() === categoria &&
        String(row[2]).trim() === "Activo"
      )
      .map(row => row[1]);
  }

  etiquetas = [...new Set(etiquetas)].filter(Boolean);

  if (etiquetas.length === 0) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(etiquetas, true)
    .setAllowInvalid(false)
    .build();

  etiquetaCell.setDataValidation(rule);
}
