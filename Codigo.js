function onOpen(e) {
  addMenu();
}

const ACTIVESPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const UI = SpreadsheetApp.getUi();

function addMenu() {
  SpreadsheetApp.getUi()
    .createMenu(' Automações')
    .addItem('Importar Alcance', 'importReach')
    .addItem('Importar Curtidas', 'importLikes')
    .addItem('Importar Seguidores', 'importFollowers')
    .addSeparator()
    .addItem('Sobre o script', 'aboutTheScript')
    .addToUi();
}

//Main functions
function importReach() {
  let confirmation1 = Browser.msgBox('Importar Alcance', "Faça o upload do arquivo CSV das métricas de Alcance e nomeia a nova guia como \"Alcance\". Apenas após seguir esse passo, inicie o script.", Browser.Buttons.OK_CANCEL).toUpperCase();

  if (confirmation1 === 'OK') {
    if (!doesGetSheetByNameExistis('Alcance')) {
      UI.alert('Não foi encontrada a planilha "Alcance". Tente novamente após criar a planilha.');
    }

    let reachData = {
      facebook: null,
      instagram: null,
    };

    if (findRowWithValue('Alcance', 'A:A', 'Alcance no Facebook')) {
      reachData.facebook = getReachData('Facebook'); // Potential missing parenthesis here (check this function)
    }

    if (findRowWithValue('Alcance', 'A:A', 'Alcance do Instagram')) {
      reachData.instagram = getReachData('Instagram'); // Potential missing parenthesis here (check this function)
    }

    return pasteReachData(reachData);
  } else {
    showNothingWasDoneAlert();
  }
}

/**
 * Retrieves reach data from the active spreadsheet's 'Alcance' sheet based on the specified platform.
 * The data is fetched based on the rows between 'Alcance no Facebook' or 'Alcance do Instagram' and the next empty row, and it is stored in a 2D array containing columns A (Date) and B (String).
 *
 * Each element of the array is a sub-array with two elements: [Date, String].
 *
 * @param {string} platform - The platform for which reach data is retrieved. Must be either 'Facebook' or 'Instagram'.
 * @returns {Array[]} The reach data as a 2D array with columns A (Date) and B (String).
 * 
 * @throws {Error} If 'Alcance' sheet is not found in the active spreadsheet.
 * @throws {Error} If the 'Alcance' and 'Alcance no Facebook' or 'Alcance do Instagram' values are not found in the sheet.
 */
function getReachData(platform) {
  const PLATFORM = platform === 'Facebook' ? 'Alcance no Facebook' : 'Alcance do Instagram'

  ACTIVESPREADSHEET.toast('Importando dados de ' + PLATFORM + '.');
  const ACTIVESHEET = ACTIVESPREADSHEET.getSheetByName('Alcance');

  let startOfData = findRowWithValue('Alcance', 'A:A', PLATFORM) + 2;
  let endOfData;
  if(platform === 'Facebook') {
    endOfData = findRowWithValue('Alcance', 'A:A', '') - 1;
  } else {
    endOfData = ACTIVESHEET.getLastRow();
  }
  
  return dataValues = ACTIVESHEET.getRange("A" + startOfData + ":" + "B" + endOfData).getValues()
}

function importLikes() {
  let confirmation1 = Browser.msgBox('Importar Visitas', "Faça o upload do arquivo CSV das métricas de Visitas e Curtidas e nomeia a nova guia como \"Visitas\". Apenas após seguir esse passo, inicie o script.", Browser.Buttons.OK_CANCEL).toUpperCase();

  //Validations
  if (confirmation1 !== 'OK') {
    return showNothingWasDoneAlert();
  }
  if (!doesGetSheetByNameExistis('Visitas')) {
    UI.alert('Não foi encontrada a planilha "Curtidas". Tente novamente após criar a planilha.');
  }

  let likesData = {
    facebook: null,
    instagram: null,
  };

  if (findRowWithValue('Visitas', 'A:A', 'Visitas ao Facebook')) {
    likesData.facebook = getLikeData('Facebook'); // Potential missing parenthesis here (check this function)
  }

  if (findRowWithValue('Visitas', 'A:A', 'Visitas ao perfil do Instagram')) {
    likesData.instagram = getLikeData('Instagram'); // Potential missing parenthesis here (check this function)
  }

  return pasteLikesData(likesData);
}

function getLikeData(platform) { // Potential missing parenthesis here (update after review)
  const PLATFORM = platform === 'Facebook' ? 'Visitas ao Facebook' : 'Visitas ao perfil do Instagram';

  ACTIVESPREADSHEET.toast('Importando dados de ' + PLATFORM);
  const ACTIVESHEET = ACTIVESPREADSHEET.getSheetByName('Visitas');

  let startOfData = findRowWithValue('Visitas', 'A:A', PLATFORM) + 2;
  let endOfData;

  if (platform === 'Facebook') {
    endOfData = findRowWithValue('Visitas', 'A:A', '') - 1;
  } else {
    endOfData = ACTIVESHEET.getLastRow();
  }

  return dataValues = ACTIVESHEET.getRange("A" + startOfData + ":" + "B" + endOfData).getValues(); // Potential missing parenthesis here (update after review)
}

function getLikeData(platform) {
  const PLATFORM = platform === 'Facebook' ? 'Visitas ao Facebook' : 'Visitas ao perfil do Instagram';

  ACTIVESPREADSHEET.toast('Importando dados de ' + PLATFORM);
  const ACTIVESHEET = ACTIVESPREADSHEET.getSheetByName('Visitas');

  let startOfData = findRowWithValue('Visitas', 'A:A', PLATFORM) + 2;
  let endOfData;

  if(platform === 'Facebook'){
    endOfData = findRowWithValue('Visitas', 'A:A', '') - 1;
  } else {
    endOfData = ACTIVESHEET.getLastRow();
  }

  return dataValues = ACTIVESHEET.getRange("A" + startOfData + ":" + "B" + endOfData).getValues();
}

function importFollowers() {
  let confirmation1 = Browser.msgBox('Importar Novos seguidores', "Faça o upload do arquivo CSV das métricas de Novos seguidores e nomeia a nova guia como \"Seguidores\". Apenas após seguir esse passo, inicie o script.", Browser.Buttons.OK_CANCEL).toUpperCase();

  if(confirmation1 === 'OK') {
    if(!doesGetSheetByNameExistis('Seguidores')){UI.alert('Não foi encontrada a planilha "Seguidores". Tente novamente após criar a planilha.')};

    let followersData = {
      facebook: null,
      instagram: null
    }

    if(findRowWithValue('Seguidores', 'A:A', 'Seguidores')){
      followersData.facebook = getFollowersData('Facebook');
    }

    if(findRowWithValue('Seguidores', 'A:A', 'Seguidos no Instagram')){
      followersData.instagram = getFollowersData('Instagram');
    }

    return pasteFollowersData(followersData)
  } else {
    showNothingWasDoneAlert();
  }
}

/**
 * Retrieves reach data from the active spreadsheet's 'Alcance' sheet based on the specified platform.
 * The data is fetched based on the rows between 'Seguidores' or 'Seguidores no Instagram' and the next empty row, and it is stored in a 2D array containing columns A (Date) and B (String).
 *
 * Each element of the array is a sub-array with two elements: [Date, String].
 *
 * @param {string} platform - The platform for which reach data is retrieved. Must be either 'Facebook' or 'Instagram'.
 * @returns {Array[]} The reach data as a 2D array with columns A (Date) and B (String).
 * 
 * @throws {Error} If 'Alcance' sheet is not found in the active spreadsheet.
 * @throws {Error} If the 'Alcance' and 'Alcance no Facebook' or 'Alcance do Instagram' values are not found in the sheet.
 */
function getFollowersData(platform) {
  const PLATFORM = platform === 'Facebook' ? 'Seguidores' : 'Seguidos no Instagram';

  ACTIVESPREADSHEET.toast('Importando dados de ' + PLATFORM);
  const ACTIVESHEET = ACTIVESPREADSHEET.getSheetByName('Seguidores');

  let startOfData = findRowWithValue('Seguidores', 'A:A', PLATFORM) + 2;
  let endOfData;

  if(platform === 'Facebook'){
    endOfData = findRowWithValue('Seguidores', 'A:A', '') - 1;
  } else {
    endOfData = ACTIVESHEET.getLastRow();
  }

  return dataValues = ACTIVESHEET.getRange("A" + startOfData + ":" + "B" + endOfData).getValues();
}

//Auxiliary functions
/**
 * Displays an alert indicating that nothing was done.
 */
function showNothingWasDoneAlert() {
  UI.alert('Nada foi feito.');
}

/**
 * Checks if a sheet with the specified name exists in the active spreadsheet.
 *
 * @param {string} name - The name of the sheet to check.
 * @returns {boolean} True if the sheet exists, false otherwise.
 */
function doesGetSheetByNameExistis(name) {
  if (ACTIVESPREADSHEET.getSheetByName(name)) {
    return true;
  } else {
    return false;
  }
}

/**
 * Checks if the value at the specified cell in the given sheet is a valid date.
 *
 * @param {string} sheetName - The name of the sheet.
 * @param {number} sheetRow - The row number of the cell.
 * @param {number} sheetColumn - The column number of the cell.
 * @returns {boolean} True if the value is a valid date, false otherwise.
 */
function isValueADate(sheetName, sheetRow, sheetColumn) {
  const str = ACTIVESPREADSHEET.getSheetByName(sheetName).getRange(sheetRow, sheetColumn).getValue().toString();
  const date = new Date(str);
  const isDate = !isNaN(date.getTime()) && !isNaN(Date.parse(str));
  return isDate;
}

/**
 * Search a value in an interval.
 *
 * @param {string} sheetName The name of the sheet in Google Sheets.
 * @param {string} interval The interval (in A1 notation) to look for the value in the sheet, like A:A.
 * @param {string} value The value to look for.
 * @return {number} The value's row number or -1 if not found.
 */
function findRowWithValue(sheetName, interval, value) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var targetValue = value;
  var columnA = sheet.getRange(interval).getValues();

  for (var i = 0; i < columnA.length; i++) {
    if (columnA[i][0] === targetValue) {
      return i + 1;
    }
  }

  Logger.log(targetValue + ' not found in the sheet.');
  return -1;
}

function pasteReachData(platformData) {
  const FACEBOOK_SHEET = ACTIVESPREADSHEET.getSheetByName('Facebook Insights');
  const FACEBOOK_DATA = platformData.facebook;
  const INSTAGRAM_SHEET = ACTIVESPREADSHEET.getSheetByName('Instagram Insights');
  const INSTAGRAM_DATA = platformData.instagram;

  const pasteData = (platform, sheetToUse, dataToUse) => {
    ACTIVESPREADSHEET.toast(`Inserindo dados do ${platform}... Aguarde a confirmação de conclusão.`);
    const INTERVAL = sheetToUse.getLastRow();
    const SEARCH = sheetToUse.getRange('C2:C' + INTERVAL).getValues().map(index => {return new Date(index[0])});

    for(i=0; i < dataToUse.length; i++){
      const DATE_TO_CHECK = new Date(dataToUse[i][0]);
      const INFO_FROM_DATE = dataToUse[i][1];
      let index = SEARCH.findIndex(date => date.getTime() === DATE_TO_CHECK.getTime());
      
      if(index !== -1){
        sheetToUse.getRange('D' + (index + 2)).setValue(INFO_FROM_DATE);
      } else {
        const LAST_ROW = getLastRowWithValue('C:C') + 1;
        sheetToUse.getRange('C' + LAST_ROW).setValue(DATE_TO_CHECK);
        sheetToUse.getRange('D' + LAST_ROW).setValue(INFO_FROM_DATE);
      }
    }

    ACTIVESPREADSHEET.toast(`Informações do ${platform} inseridas com sucesso.`);
  }

  if(FACEBOOK_DATA !== null){
    pasteData('Facebook', FACEBOOK_SHEET, FACEBOOK_DATA);
  }

  if(INSTAGRAM_DATA !== null){
    pasteData('Instagram', INSTAGRAM_SHEET, INSTAGRAM_DATA);
  }
}

function pasteLikesData(platformData) {
  const FACEBOOK_SHEET = ACTIVESPREADSHEET.getSheetByName('Facebook Insights');
  const FACEBOOK_DATA = platformData.facebook;
  const INSTAGRAM_SHEET = ACTIVESPREADSHEET.getSheetByName('Instagram Insights');
  const INSTAGRAM_DATA = platformData.instagram;

  const pasteData = (platform, sheetToUse, dataToUse) => {
    ACTIVESPREADSHEET.toast(`Inserindo dados do ${platform}... Aguarde a confirmação de conclusão.`);
    const INTERVAL = sheetToUse.getLastRow();
    const SEARCH = sheetToUse.getRange('C2:C' + INTERVAL).getValues().map(index => {return new Date(index[0])});

    for(i=0; i < dataToUse.length; i++){
      const DATE_TO_CHECK = new Date(dataToUse[i][0]);
      const INFO_FROM_DATE = dataToUse[i][1];
      let index = SEARCH.findIndex(date => date.getTime() === DATE_TO_CHECK.getTime());
      
      if(index !== -1){
        sheetToUse.getRange('E' + (index + 2)).setValue(INFO_FROM_DATE);
      } else {
        const LAST_ROW = getLastRowWithValue('C:C') + 1;
        sheetToUse.getRange('C' + LAST_ROW).setValue(DATE_TO_CHECK);
        sheetToUse.getRange('E' + LAST_ROW).setValue(INFO_FROM_DATE);
      }
    }

    ACTIVESPREADSHEET.toast(`Informações do ${platform} inseridas com sucesso.`);
  }

  if(FACEBOOK_DATA !== null){
    pasteData('Facebook', FACEBOOK_SHEET, FACEBOOK_DATA);
  }

  if(INSTAGRAM_DATA !== null){
    pasteData('Instagram', INSTAGRAM_SHEET, INSTAGRAM_DATA);
  }
}

function pasteFollowersData(platformData) {
  const FACEBOOK_SHEET = ACTIVESPREADSHEET.getSheetByName('Facebook Insights');
  const FACEBOOK_DATA = platformData.facebook;
  const INSTAGRAM_SHEET = ACTIVESPREADSHEET.getSheetByName('Instagram Insights');
  const INSTAGRAM_DATA = platformData.instagram;

  const pasteData = (platform, sheetToUse, dataToUse) => {
    ACTIVESPREADSHEET.toast(`Inserindo dados do ${platform}... Aguarde a confirmação de conclusão.`);
    const INTERVAL = sheetToUse.getLastRow();
    const SEARCH = sheetToUse.getRange('C2:C' + INTERVAL).getValues().map(index => {return new Date(index[0])});
    
    for(i=0; i < dataToUse.length; i++){
      const DATE_TO_CHECK = new Date(dataToUse[i][0]);
      const INFO_FROM_DATE = dataToUse[i][1];
      let index = SEARCH.findIndex(date => date.getTime() === DATE_TO_CHECK.getTime());
      
      if(index !== -1){
        sheetToUse.getRange('F' + (index + 2)).setValue(INFO_FROM_DATE);
      } else {
        const LAST_ROW = getLastRowWithValue('C:C') + 1;
        sheetToUse.getRange('C' + LAST_ROW).setValue(DATE_TO_CHECK);
        sheetToUse.getRange('F' + LAST_ROW).setValue(INFO_FROM_DATE);
      }
    }
  }

  if(FACEBOOK_DATA !== null){
    pasteData('Facebook', FACEBOOK_SHEET, FACEBOOK_DATA);
  }

  if(INSTAGRAM_DATA !== null){
    pasteData('Instagram', INSTAGRAM_SHEET, INSTAGRAM_DATA);
  }

  ACTIVESPREADSHEET.toast('Informações de inseridas com sucesso.');
}

function getLastRowWithValue(columnToCheck) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var column = sheet.getRange(columnToCheck);
  var values = column.getValues();  // get all data in a column
  var lastRow = -1;
  for(var i = 0; i < values.length; i++){
    if(values[i][0] != ""){
      lastRow = i+1;
    }
  }
  return lastRow;  // return -1 if the column is empty
}

function aboutTheScript() {
  Browser.msgBox('Sobre o script', 'Criado por André Moura Santos, contato@andremourasantos.com.br.', Browser.Buttons.OK)
}