/**
 * Ordena las hojas de un libro de cálculo alfabéticamente en sentido ascendente o descendente
 * utilizando el servicio avanzado (API) de hojas de cálculo.
 * @param {boolean} [ascendente]
 */
function ordenarHojasApi(ascendente = true) {

  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  // Precisamos únicamente las letras correspondientes al código del idioma
  const collator = new Intl.Collator(hdc.getSpreadsheetLocale().split('_')[0], { numeric: true, sensitivity: 'base' });

  try {

    // Obtiene vector de propiedades de hojas (nombre, índice)
    const hojasOrdenadas = Sheets.Spreadsheets.get(
      SpreadsheetApp.getActiveSpreadsheet().getId(), { fields: 'sheets.properties(sheetId,title)' }
    ).sheets.sort((hoja1, hoja2) =>
      Math.pow(-1, !ascendente) * collator.compare(hoja1.properties.title, hoja2.properties.title)
    );

    const hojaActual = hdc.getActiveSheet();
    Sheets.Spreadsheets.batchUpdate(
      {
        requests: hojasOrdenadas.map(({ properties: { sheetId } }, pos) => (
          { updateSheetProperties: { properties: { sheetId, index: pos }, fields: 'index' } }
        )),
        includeSpreadsheetInResponse: false
      },
      hdc.getId()
    );
    hdc.setActiveSheet(hojaActual);
    ui.alert(`Se han ordenado alfabéticamente ${hojasOrdenadas.length} hoja(s) en sentido ${ascendente ? 'ascendente' : 'descendente'}.`, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert(`Se ha producido un error inesperado al ordenar las hojas, inténtalo de nuevo.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }

}