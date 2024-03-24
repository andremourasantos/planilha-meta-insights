const UI = SpreadsheetApp.getUi();
const ACTIVE_SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();

function aboutTheScript() {
  Browser.msgBox('Sobre o código', 'Criado por André Moura Santos, contato@andremourasantos.com.br.', Browser.Buttons.OK);
}