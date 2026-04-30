/****************************************************
 * DatosPrueba_Cuentas.gs
 * Pobla la hoja Cuentas con datos de prueba que
 * cubren todas las variaciones posibles de campos.
 *
 * Cómo usar:
 *   1. Abre el editor de Apps Script del archivo
 *   2. Pega este código
 *   3. Ejecuta: poblarDatosPruebaCuentas()
 ****************************************************/

function poblarDatosPruebaCuentas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cuentas");

  if (!sheet) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja 'Cuentas'.");
    return;
  }

  // Limpiar datos anteriores (conserva encabezado fila 1)
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  /*
   * Matriz de datos de prueba.
   * Columnas (según estructura de la hoja Cuentas):
   *   A: ID_Cuenta
   *   B: Nombre
   *   C: Tipo de cuenta
   *   D: Saldo conciliado
   *   E: Fecha conciliación
   *   F: Saldo calculado  ← normalmente calculado; aquí lo dejamos vacío
   *   G: Genera intereses
   *   H: Tasa interés E.A.
   *   I: Periodicidad interés
   *   J: Estado
   *   K: Cuenta patrimonio
   *
   * Variaciones cubiertas:
   *  - Los 7 tipos de cuenta
   *  - Estado Activo e Inactivo
   *  - Genera intereses: Sí y No
   *  - Cuenta patrimonio: Sí y No
   *  - Con y sin saldo conciliado
   *  - Con y sin fecha de conciliación
   *  - Con y sin tasa de interés / periodicidad
   */

  const hoy = new Date();
  const ayer = new Date(hoy); ayer.setDate(hoy.getDate() - 1);
  const haceDias = (n) => { const d = new Date(hoy); d.setDate(hoy.getDate() - n); return d; };

  const datos = [
    // ── TIPO: Efectivo ──────────────────────────────────────────────────────
    // Activo | Sin saldo conciliado | Sin fecha | No intereses | No patrimonio
    [
      "CTA-T01", "Billetera física", "Efectivo",
      "", "", "",
      "No", "", "",
      "Activo", "No"
    ],
    // Activo | Con saldo conciliado | Con fecha | No intereses | No patrimonio
    [
      "CTA-T02", "Caja menor oficina", "Efectivo",
      150000, haceDias(5), "",
      "No", "", "",
      "Activo", "No"
    ],

    // ── TIPO: Cuenta de ahorros ─────────────────────────────────────────────
    // Activo | Con saldo | Con fecha | Sí intereses | Sí patrimonio
    [
      "CTA-T03", "Ahorros Bancolombia principal", "Cuenta de ahorros",
      5200000, haceDias(3), "",
      "Sí", "4.5%", "Mensual",
      "Activo", "Si"
    ],
    // Activo | Con saldo | Con fecha | Sí intereses | No patrimonio
    [
      "CTA-T04", "Ahorros Davivienda operativo", "Cuenta de ahorros",
      980000, haceDias(7), "",
      "Sí", "2.0%", "Mensual",
      "Activo", "No"
    ],
    // Inactivo | Sin saldo | Sin fecha | No intereses | No patrimonio
    [
      "CTA-T05", "Ahorros cerrada prueba", "Cuenta de ahorros",
      "", "", "",
      "No", "", "",
      "Inactivo", "No"
    ],

    // ── TIPO: Cuenta corriente ──────────────────────────────────────────────
    // Activo | Con saldo | Con fecha | No intereses | No patrimonio
    [
      "CTA-T06", "Corriente Bancolombia empresa", "Cuenta corriente",
      3750000, haceDias(2), "",
      "No", "", "",
      "Activo", "No"
    ],
    // Inactivo | Con saldo | Con fecha | No intereses | No patrimonio
    [
      "CTA-T07", "Corriente antigua BBVA", "Cuenta corriente",
      12000, haceDias(180), "",
      "No", "", "",
      "Inactivo", "No"
    ],

    // ── TIPO: Billetera digital ─────────────────────────────────────────────
    // Activo | Sin saldo conciliado | Sin fecha | No intereses | No patrimonio
    [
      "CTA-T08", "Nequi personal", "Billetera digital",
      "", "", "",
      "No", "", "",
      "Activo", "No"
    ],
    // Activo | Con saldo | Con fecha | Sí intereses | Sí patrimonio (ej. Ualá con 4%)
    [
      "CTA-T09", "Ualá rendimientos", "Billetera digital",
      2100000, ayer, "",
      "Sí", "12.0%", "Diaria",
      "Activo", "Si"
    ],
    // Activo | Con saldo | Sin fecha | No intereses | No patrimonio
    [
      "CTA-T10", "MercadoPago saldo", "Billetera digital",
      45000, "", "",
      "No", "", "",
      "Activo", "No"
    ],
    // Inactivo | Sin saldo | Sin fecha | No intereses | No patrimonio
    [
      "CTA-T11", "Daviplata inactiva", "Billetera digital",
      "", "", "",
      "No", "", "",
      "Inactivo", "No"
    ],

    // ── TIPO: Bolsillo ──────────────────────────────────────────────────────
    // Activo | Con saldo | Con fecha | No intereses | Sí patrimonio (bolsillo de ahorro)
    [
      "CTA-T12", "Bolsillo vacaciones Nequi", "Bolsillo",
      800000, haceDias(10), "",
      "No", "", "",
      "Activo", "Si"
    ],
    // Activo | Sin saldo | Sin fecha | No intereses | No patrimonio
    [
      "CTA-T13", "Bolsillo gastos variables", "Bolsillo",
      "", "", "",
      "No", "", "",
      "Activo", "No"
    ],
    // Inactivo | Con saldo | Con fecha | No intereses | No patrimonio
    [
      "CTA-T14", "Bolsillo navidad (cerrado)", "Bolsillo",
      320000, haceDias(90), "",
      "No", "", "",
      "Inactivo", "No"
    ],

    // ── TIPO: Inversión ─────────────────────────────────────────────────────
    // Activo | Con saldo | Con fecha | Sí intereses alta tasa | Sí patrimonio
    [
      "CTA-T15", "CDT Bancolombia 90 días", "Inversión",
      10000000, haceDias(15), "",
      "Sí", "14.2%", "Al vencimiento",
      "Activo", "Si"
    ],
    // Activo | Con saldo | Con fecha | Sí intereses | Sí patrimonio (acciones/ETF)
    [
      "CTA-T16", "Portafolio acciones Trii", "Inversión",
      3500000, hoy, "",
      "Sí", "N/A (variable)", "Variable",
      "Activo", "Si"
    ],
    // Activo | Sin saldo | Sin fecha | Sí intereses | Sí patrimonio (recién abierta)
    [
      "CTA-T17", "Inversión Littio (nueva)", "Inversión",
      "", "", "",
      "Sí", "9.5%", "Mensual",
      "Activo", "Si"
    ],
    // Inactivo | Con saldo residual | Con fecha | No intereses | Sí patrimonio
    [
      "CTA-T18", "Fondo antiguo vencido", "Inversión",
      500, haceDias(365), "",
      "No", "", "",
      "Inactivo", "Si"
    ],

    // ── TIPO: Cuenta bancaria ───────────────────────────────────────────────
    // Activo | Con saldo | Con fecha | No intereses | No patrimonio (genérica)
    [
      "CTA-T19", "Cuenta bancaria genérica Scotiabank", "Cuenta bancaria",
      620000, haceDias(4), "",
      "No", "", "",
      "Activo", "No"
    ],
    // Activo | Con saldo | Con fecha | Sí intereses | Sí patrimonio
    [
      "CTA-T20", "Cuenta remunerada Nubank", "Cuenta bancaria",
      1800000, ayer, "",
      "Sí", "7.0%", "Mensual",
      "Activo", "Si"
    ],
    // Inactivo | Sin saldo | Sin fecha | No intereses | No patrimonio
    [
      "CTA-T21", "Cuenta bancaria inactiva HSBC", "Cuenta bancaria",
      "", "", "",
      "No", "", "",
      "Inactivo", "No"
    ],
  ];

  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  SpreadsheetApp.getUi().alert(
    `✅ Datos de prueba cargados en Cuentas.\n\n` +
    `Total registros: ${datos.length}\n\n` +
    `Variaciones cubiertas:\n` +
    `• 7 tipos de cuenta\n` +
    `• Estado: Activo (15) / Inactivo (6)\n` +
    `• Genera intereses: Sí (8) / No (13)\n` +
    `• Cuenta patrimonio: Sí (9) / No (12)\n` +
    `• Con saldo conciliado: 14 / Sin saldo: 7\n` +
    `• Con fecha conciliación: 14 / Sin fecha: 7`
  );
}


function limpiarDatosPruebaCuentas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cuentas");

  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  SpreadsheetApp.getUi().alert("🗑️ Datos de prueba de Cuentas eliminados.");
}
