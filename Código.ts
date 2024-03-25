const UI = SpreadsheetApp.getUi();
const ACTIVE_SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const SCRIPT_LIBRARY_NAME = 'PlanilhaMetaInsights';
const FACEBOOK_INSIGHTS_SHEET_NAME:string = 'Facebook Insights';
const INSTAGRAM_INSIGHTS_SHEET_NAME:string = 'Instagram Insights';

//Sheet estilization
const ROW_HEIGHT:number = 42;
const COLUMN_WIDTH_P:number = 84;
const COLUMN_WIDTH_M:number = COLUMN_WIDTH_P * 2;
const COLUMN_WIDTH_G:number = COLUMN_WIDTH_P * 3;

/**
 * Create a new menu option on the top bar of the active spreadsheet named "🤖 Assistente" with all the functions available on this code. Automatically calls the functions on the client side without the need for additional settings, except adding it to the onOpen functions on the client script.
 *
 */
function addMenu() {
  const FIRST_STEPS_SUBMENU = UI.createMenu('Primeiros passos').addItem('Criar planilhas', `${SCRIPT_LIBRARY_NAME}.createFacebookInsightsSheet`);

  UI.createMenu('🤖 Assistente')
    .addSubMenu(FIRST_STEPS_SUBMENU)
    .addSeparator()
    .addItem('Sobre o código', `${SCRIPT_LIBRARY_NAME}.aboutTheScript`)
  .addToUi();
}

/**
 * Create an entire new sheet for the Facebook Insights data. If the sheet already exists, it returns a custom error alert.
 *
 * @returns Either "sheet_already_exists" if the sheet already exists or "sheet_created" if a new sheet is successfully created.
 */
function createFacebookInsightsSheet():'sheet_already_exists' | 'sheet_created' {
  let facebookSheet:GoogleAppsScript.Spreadsheet.Sheet;

  try {
    ACTIVE_SPREADSHEET.toast(`Criando planilha ${FACEBOOK_INSIGHTS_SHEET_NAME}...`)
    facebookSheet = ACTIVE_SPREADSHEET.insertSheet(FACEBOOK_INSIGHTS_SHEET_NAME);
  } catch (error) {
    showCustomErrorAlert('⚠️ Planilha já existente', `A planilha ${FACEBOOK_INSIGHTS_SHEET_NAME} já foi criada. Caso queria criá-la novamente, é necessário excluir a atual e executar essa ação novamente.`);

    return 'sheet_already_exists';
  }

  ACTIVE_SPREADSHEET.setActiveSheet(facebookSheet);

  defineRowsAndColumnsWireframes(facebookSheet, 365, 6);
  defineSheetTextAligment(facebookSheet, 'right');
  commomSheetEstilization(facebookSheet);

  const HEADER_TITLES = [['Nº do Mês', 'Mês', 'Data', 'Alcance', 'Curtidas', 'Seguidores']];

  //TODO: MAKE THE defineRowsAndColumnsWireframes FUNCTION ACCEPT COLUMN WIDTHS AS A PARAMETER.
  const HEADER_COLUMN_WIDTHS = [COLUMN_WIDTH_P, COLUMN_WIDTH_P, COLUMN_WIDTH_P, COLUMN_WIDTH_M, COLUMN_WIDTH_M, COLUMN_WIDTH_M];
  
  facebookSheet.getRange(1,1,1,6).setValues(HEADER_TITLES);

  return 'sheet_created';
}

/**
 * Show a popup with information about the script's author.
 *
 */
function aboutTheScript():void {
  Browser.msgBox('Sobre o código', 'Criado por André Moura Santos, contato@andremourasantos.com.br.', Browser.Buttons.OK);
}

//Auxiliary functions
/**
 * Shows a Browser message saying that nothing was done and offers only a button to click "OK".
 *
 */
function showNothingWasDoneBrowser():void {
  Browser.msgBox('Ação cancelada', 'Nada foi feito.', Browser.Buttons.OK);
}

/**
 * Show a custom error message in the Browser message box style. Offers only a button to click "OK".
 *
 * @param {string} title The title of the message box.
 * @param {string} msg The message to be shown in the message box.
 */
function showCustomErrorAlert(title:string, msg:string):void {
  Browser.msgBox(title, msg, Browser.Buttons.OK);
}

/**
 * Define the length (number of rows) and width (number of columns) of the sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet A sheet object.
 * @param {number} maxRows The number of rows the sheet should have.
 * @param {number} maxColumns The number of columns the sheet should have.
 */
function defineRowsAndColumnsWireframes(sheet:GoogleAppsScript.Spreadsheet.Sheet, maxRows:number, maxColumns:number) {
  sheet.deleteColumns(maxColumns, sheet.getMaxColumns() - maxColumns);
  sheet.deleteRows(maxRows, sheet.getMaxRows() - maxRows);
  sheet.setRowHeightsForced(1, sheet.getMaxRows(), 42);
}

/**
 * Define the text aligment for all the Sheet cells.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet A sheet object.
 * @param {('left' | 'center' | 'right')} type The type of text aligment.
 */
function defineSheetTextAligment(sheet:GoogleAppsScript.Spreadsheet.Sheet, type:'left' | 'center' | 'right') {
  const ALL_SHEET = sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns());

  ALL_SHEET
    .setVerticalAlignment("middle")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setHorizontalAlignment(type)
}

/**
 * Apply the a pre-defined common sheet stylization to the entire sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet A sheet object.
 */
function commomSheetEstilization(sheet:GoogleAppsScript.Spreadsheet.Sheet) {
  const ALL_SHEET = sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns());
  const HEADER_STYLE = SpreadsheetApp.newTextStyle().setBold(true).build();

  sheet.getRange(1,1,1,sheet.getMaxColumns()).setTextStyle(HEADER_STYLE);

  ALL_SHEET
    .setFontFamily('Atkinson Hyperlegible')
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
}