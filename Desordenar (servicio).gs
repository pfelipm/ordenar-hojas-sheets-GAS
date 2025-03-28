/**
 * Desordena las hojas de un libro de cálculo de forma aleatoria
 * utilizando el servicio integrado de hojas de cálculo.
 * Mantiene el estado de visibilidad de las hojas y restaura la hoja activa.
 * 
 * Generado por Gemini 2.5 Pro (experimental) a partir de la función de ordenación ordenarHojasServicio().
 */
function desordenarHojasServicio() {

  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const hojas = hdc.getSheets();
    const numHojas = hojas.length;

    // Si hay 1 hoja o menos, no hay nada que desordenar
    if (numHojas <= 1) {
      ui.alert(`Solo hay ${numHojas} hoja(s), no es necesario desordenar.`, ui.ButtonSet.OK);
      return;
    }

    // --- Inicio: Algoritmo de Fisher-Yates (Knuth) Shuffle ---
    // Recorre el array desde el final hacia el principio
    for (let i = numHojas - 1; i > 0; i--) {
      // Elige un índice aleatorio j entre 0 e i (inclusive)
      const j = Math.floor(Math.random() * (i + 1));
      // Intercambia el elemento en i con el elemento en j
      [hojas[i], hojas[j]] = [hojas[j], hojas[i]];
    }
    // --- Fin: Algoritmo de Fisher-Yates ---
    // Ahora el array 'hojas' contiene los objetos Sheet en un orden aleatorio.

    const hojasPorOcultar = [];
    const hojaActual = hdc.getActiveSheet(); // Guarda la hoja activa actual

    // Mueve las hojas a su nueva posición aleatoria
    hojas.forEach((hoja, pos) => {
      // Comprueba si la hoja estaba oculta ANTES de moverla
      // (moverla la hará visible temporalmente si estaba oculta)
      if (hoja.isSheetHidden()) {
        hojasPorOcultar.push(hoja); // Añade el objeto hoja a la lista para ocultar después
      }
      hdc.setActiveSheet(hoja); // Activa la hoja para poder moverla
      hdc.moveActiveSheet(pos + 1); // Mueve la hoja a su nueva posición (índice + 1)
    });

    // Necesario confirmar cambios antes de procesar hojas a ocultar
    SpreadsheetApp.flush();
    // Mantenemos la pausa por consistencia con la función original,
    // por si hay problemas de sincronización al ocultar rápidamente.
    Utilities.sleep(1000);

    // Vuelve a ocultar las hojas que estaban ocultas originalmente
    if (hojasPorOcultar.length > 0) {
      hojasPorOcultar.forEach(hoja => {
        // Podría haber un error si la hoja fue eliminada mientras corría el script,
        // aunque es poco probable. Añadimos un check simple.
        try {
          hoja.hideSheet();
        } catch (e) {
          console.warn(`No se pudo ocultar la hoja "${hoja.getName()}". ¿Quizás fue eliminada? Error: ${e.message}`);
        }
      });
    }

    // Restaura la hoja que estaba activa al principio
    hdc.setActiveSheet(hojaActual);
    SpreadsheetApp.flush(); // Asegura que la activación final se aplique

    ui.alert(`Se han desordenado aleatoriamente ${numHojas} hoja(s).`, ui.ButtonSet.OK);

  } catch (e) {
    // Mismo manejo de errores que la función original
    ui.alert(`Se ha producido un error inesperado al desordenar las hojas, inténtalo de nuevo.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
    // Opcional: Registrar el error completo para depuración
    console.error(`Error en desordenarHojas: ${e.message} \n ${e.stack}`);
  }
}