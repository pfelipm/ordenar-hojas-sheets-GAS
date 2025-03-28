/**
 * Código de acompañamiento del artículo:
 * «Velocidad vs. Permisos: Ordenando pestañas de hojas de cálculo con Apps Script»
 * https://pablofelip.online/velocidad-permisos-ordenando-pestanas-apps-script
 * Pablo Felip Monferrer | 2025 
 * 
 * @OnlyCurrentDoc
 */

function onOpen() {

    SpreadsheetApp.getUi().createMenu('Ordenar hojas')
        .addItem('🐢 Ordenar hojas (A → Z) [SERVICIO]', 'ordenarHojasServicioAsc')
        .addItem('🐢 Ordenar hojas (Z → A) [SERVICIO]', 'ordenarHojasServicioDesc')
        .addSeparator()
        .addItem('🔀 Desordenar hojas [SERVICIO]', 'desordenarHojasServicio')
        .addSeparator()
        .addItem('⚡ Ordenar hojas (A → Z) [API]', 'ordenarHojasApiAsc')
        .addItem('⚡ Ordenar hojas (Z → A) [API]', 'ordenarHojasApiDesc')
        .addToUi();
}

// Envoltorios para ordenarHojas()
function ordenarHojasServicioAsc() { ordenarHojasServicio(true); }
function ordenarHojasServicioDesc() { ordenarHojasServicio(false); }
function ordenarHojasApiAsc() { ordenarHojasApi(true); }
function ordenarHojasApiDesc() { ordenarHojasApi(false); }

function foo() {
    const locale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
    console.info('Locale (getSpreadsheetLocale()):' + locale);
    console.info(`Idioma: (getSpreadsheetLocale().split('_')[0]): ` + locale.split('_')[0]);
}