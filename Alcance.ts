const META_REACH_SHEET_NAME:string = 'Alcance';

/**
 * Returns if a sheet named "Alcance" exists in the Spreadsheet.
 *
 * @return {*}  {boolean}
 */
function doesReachSheetExists():boolean {
  const META_REACH_SHEET:GoogleAppsScript.Spreadsheet.Sheet | null = ACTIVE_SPREADSHEET.getSheetByName(META_REACH_SHEET_NAME);

  return META_REACH_SHEET !== null ? true : false;
}

/**
 * Get the Reach data from the Reach CSV file imported into the Spreadsheet.
 * 
 * Automatically calls pasteReachData for Facebook and Instagram data (if found on sheet).
 *
 * @return {*}  {void}
 */
function getReachData():void {
  const META_REACH_SHEET:GoogleAppsScript.Spreadsheet.Sheet | null = ACTIVE_SPREADSHEET.getSheetByName(META_REACH_SHEET_NAME);

  if(!doesReachSheetExists() || META_REACH_SHEET === null){return showCustomErrorAlert('⚠️ Planilha não encontrada', 'Certifique-se de que a planilha "Alcance" foi importada e nomeada corretamente antes de tentar novamente.')};
  
  const START_ROW_FACEBOOK = _getRowIndexForString('Alcance no Facebook');
  let facebookReachData:[Date, string][]  |  null = null;
  const START_ROW_INSTAGRAM = _getRowIndexForString('Alcance do Instagram');
  let instagramReachData:[Date, string][]  |  null = null;

  if(START_ROW_FACEBOOK != -1) {
    ACTIVE_SPREADSHEET.toast('Importando dados de Alcance do Facebook');
    let endRowIndex:number;

    if(START_ROW_INSTAGRAM != -1){
      endRowIndex = _getRowIndexForString('Alcance do Instagram') - 3;
    } else {
      endRowIndex = META_REACH_SHEET.getLastRow();
    }

    facebookReachData = _copyRangeValues(META_REACH_SHEET, 1, (START_ROW_FACEBOOK + 2), 2, endRowIndex - START_ROW_FACEBOOK);
  }

  if(START_ROW_INSTAGRAM != -1) {
    ACTIVE_SPREADSHEET.toast('Importando dados de Alcance do Instagram');
    const endRowIndex = _getLastRowWithValue(META_REACH_SHEET, (START_ROW_INSTAGRAM + 2), 1);

    instagramReachData = _copyRangeValues(META_REACH_SHEET, 1, (START_ROW_INSTAGRAM + 2), 2, endRowIndex);
  }

  if(facebookReachData != null) {
    pasteReachData("FACEBOOK_INSIGHTS_SHEET_NAME", facebookReachData);
  }

  if(instagramReachData != null) {
    pasteReachData('INSTAGRAM_INSIGHTS_SHEET_NAME', instagramReachData);
  }
}

/**
 * Paste the given [Date, string][] data into the respective sheet (Facebook or Instagram).
 * 
 * Is automatically called by getReachData function.
 *
 * @param {('FACEBOOK_INSIGHTS_SHEET_NAME' | 'INSTAGRAM_INSIGHTS_SHEET_NAME')} sheetName The sheet that corresponds to the data given.
 * @param {[Date, string][]} data The data extracted from the Reach CSV file.
 */
function pasteReachData(sheetName:'FACEBOOK_INSIGHTS_SHEET_NAME' | 'INSTAGRAM_INSIGHTS_SHEET_NAME', data:[Date, string][]):void {
  const sheetObjct = ACTIVE_SPREADSHEET.getSheetByName(sheetName);
  const platform = sheetName === 'FACEBOOK_INSIGHTS_SHEET_NAME' ? 'Facebook' : 'Instagram';

  ACTIVE_SPREADSHEET.toast(`teste: ${_findColumnWithValue('Alcance') === -1 ? 'erro' : _findColumnWithValue('Alcance')}`)

  const pasteData = (platform:'Facebook' | 'Instagram') => {
    ACTIVE_SPREADSHEET.toast(`Inserindo dados de Alcance do ${platform}... Aguarde a confirmação de inserção`);
  }


}

//Auxiliary functions
/**
 * Find the row index for a given string.
 *
 * @param {string} string The string to look out for
 * @return {*}  {number} Return the row index (always positive and above 0) if the value was found or -1 if not.
 */
function _getRowIndexForString(string:string):number | -1 {
  const META_REACH_SHEET:GoogleAppsScript.Spreadsheet.Sheet | null = ACTIVE_SPREADSHEET.getSheetByName(META_REACH_SHEET_NAME);

  if(!META_REACH_SHEET){return -1}
  const TEXT_FINDER = META_REACH_SHEET.createTextFinder(string);
  const ROW_INDEX = TEXT_FINDER.findNext()?.getRowIndex();

  if(ROW_INDEX){return ROW_INDEX} else {return -1}
}

/**
 * Copy the values of a given range whose values are compatible with the touple [Date, string][].
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheetObjc The sheet object to look in.
 * @param {number} startColumn The column to start the range.
 * @param {number} startRow The row index to start the range.
 * @param {number} endColumn The column to end the range.
 * @param {number} endRow The row index to end the range.
 * @return {*} Return an array of [Date, string][].
 */
function _copyRangeValues(sheetObjc:GoogleAppsScript.Spreadsheet.Sheet, startColumn:number, startRow:number, endColumn:number, endRow:number):[Date, string][] {
  const DATA = sheetObjc.getRange(startRow, startColumn, endRow, endColumn).getValues();

  return DATA as [Date, string][];
}

/**
 * Get the last row index whose value is different from blank.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheetObjc The sheet object to look in.
 * @param {number} startRow The row to start the search.
 * @param {number} columnToCheck The column to search for the blank row index.
 * @return {number} Returns the row index (always positive and above 0) if found and -1 if no blank values were found.
 */
function _getLastRowWithValue(sheetObjc:GoogleAppsScript.Spreadsheet.Sheet, startRow:number, columnToCheck:number):number | -1 {
  var sheet = sheetObjc;
  var column = sheet.getRange(startRow, columnToCheck, sheetObjc.getMaxRows() - startRow, 1);
  var values = column.getValues();  // get all data in a column
  var lastRow = -1;
  for(var i = 0; i < values.length; i++){
    if(values[i][0] != ""){
      lastRow = i+1;
    }
  }
  return lastRow;  // return -1 if the column is empty
}

function _findColumnWithValue(string:string):number | -1 {
  const META_REACH_SHEET:GoogleAppsScript.Spreadsheet.Sheet | null = ACTIVE_SPREADSHEET.getSheetByName(META_REACH_SHEET_NAME);

  if(!META_REACH_SHEET){return -1};
  const TEXT_FINDER = META_REACH_SHEET.createTextFinder(string);
  const COLUMN_INDEX = TEXT_FINDER.findNext()?.getColumn();

  if(COLUMN_INDEX){return COLUMN_INDEX} else {return -1};
}











// function pasteReachData(platformData) {
//   const FACEBOOK_SHEET = ACTIVESPREADSHEET.getSheetByName('Facebook Insights');
//   const FACEBOOK_DATA = platformData.facebook;
//   const INSTAGRAM_SHEET = ACTIVESPREADSHEET.getSheetByName('Instagram Insights');
//   const INSTAGRAM_DATA = platformData.instagram;

//   const pasteData = (platform, sheetToUse, dataToUse) => {
//     ACTIVESPREADSHEET.toast(Inserindo dados do ${platform}... Aguarde a confirmação de conclusão.);
//     const INTERVAL = sheetToUse.getLastRow();
//     const SEARCH = sheetToUse.getRange('C2:C' + INTERVAL).getValues().map(index => {return new Date(index[0])});

//     for(i=0; i < dataToUse.length; i++){
//       const DATE_TO_CHECK = new Date(dataToUse[i][0]);
//       const INFO_FROM_DATE = dataToUse[i][1];
//       let index = SEARCH.findIndex(date => date.getTime() === DATE_TO_CHECK.getTime());
      
//       if(index !== -1){
//         sheetToUse.getRange('D' + (index + 2)).setValue(INFO_FROM_DATE);
//       } else {
//         const LAST_ROW = getLastRowWithValue('C:C') + 1;
//         sheetToUse.getRange('C' + LAST_ROW).setValue(DATE_TO_CHECK);
//         sheetToUse.getRange('D' + LAST_ROW).setValue(INFO_FROM_DATE);
//       }
//     }

//     ACTIVESPREADSHEET.toast(`Informações do ${platform} inseridas com sucesso.`);
//   }

//   if(FACEBOOK_DATA !== null){
//     pasteData('Facebook', FACEBOOK_SHEET, FACEBOOK_DATA);
//   }

//   if(INSTAGRAM_DATA !== null){
//     pasteData('Instagram', INSTAGRAM_SHEET, INSTAGRAM_DATA);
//   }
// }