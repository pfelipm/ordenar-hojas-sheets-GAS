/**
 * Ordena las hojas de un libro de cálculo alfabéticamente en sentido ascendente o descendente
 * utilizando el servicio integrado de hojas de cálculo.
 * @param {boolean} [ascendente]
 */
function ordenarHojasServicio(ascendente = true) {

  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  // Precisamos únicamente el código del idioma (idioma_región)
  const collator = new Intl.Collator(hdc.getSpreadsheetLocale().split('_')[0], { numeric: true, sensitivity: 'base' });

  try {

    // Vector { hoja, nombre } para evitar múltiples llamadas a getName() en la función del sort().
    const hojasOrdenadas = hdc.getSheets()
      .map(hoja => ({ objeto: hoja, nombre: hoja.getName() }))
      .sort((hoja1, hoja2) => Math.pow(-1, !ascendente) * collator.compare(hoja1.nombre, hoja2.nombre));

    const hojasPorOcultar = [];
    const hojaActual = hdc.getActiveSheet();
    // Identifica las hojas ocultas para dejarlas del mismo modo tras realizar la ordenación
    hojasOrdenadas.forEach((hoja, pos) => {
      if (hoja.objeto.isSheetHidden()) hojasPorOcultar.push(hoja.objeto);
      hdc.setActiveSheet(hoja.objeto);
      hdc.moveActiveSheet(pos + 1);
    });
    // Necesario confirmar cambios antes de procesar hojas a ocultar
    SpreadsheetApp.flush();
    // Sin esta pausa en ocasiones alguna hoja oculta queda visible
    Utilities.sleep(1000);
    if (hojasPorOcultar.length > 0) hojasPorOcultar.forEach(hoja => hoja.hideSheet())
    hdc.setActiveSheet(hojaActual);
    SpreadsheetApp.flush();
    ui.alert(`Se han ordenado alfabéticamente ${hojasOrdenadas.length} hoja(s) en sentido ${ascendente ? 'ascendente' : 'descendente'}.`, ui.ButtonSet.OK);
 
  } catch (e) {
    ui.alert(`Se ha producido un error inesperado al ordenar las hojas, inténtalo de nuevo.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }

}