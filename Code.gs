var MC_CODE = '1589'
var COL_NAME = 0 //Which column from sheet has EY Name (first column is 0) 
var COL_ID = 2 //Which column from sheet has EXPA ID (first column is 0)
var COL_OFFSET = 15 //Starts in column 20
var ROW_OFFSET = 3 //Starts in row 4
var STAGE_OFFSET = 4 
var WEEKS_SIZE = 4 //How many weeks in a month
var PRODUCT_SIZE = 12 //Size in number of columns
var PRODUCTS = [  // 1=>GV | 2=>GT | 5=>GE // person | opportunity
  {program: 1,type: 'opportunity',name: 'iGV',offset:0},
  {program: 2,type: 'opportunity',name: 'iGT',offset:(PRODUCT_SIZE*1)},
  {program: 5,type: 'opportunity',name: 'iGE',offset:(PRODUCT_SIZE*2)},
  {program: 1,type: 'person',name: 'oGV',offset:(PRODUCT_SIZE*3)},
  {program: 2,type: 'person',name: 'oGT',offset:(PRODUCT_SIZE*4)},
  {program: 5,type: 'person',name: 'oGE',offset:(PRODUCT_SIZE*5)}
]

function onOpen() {
  var ui = SpreadsheetApp.getUi()
  
  ui.createMenu('EXPA')
    .addItem('Update Week 1 Data', 'updateWeek1')
    .addItem('Update Week 2 Data', 'updateWeek2')
    .addItem('Update Week 3 Data', 'updateWeek3')
    .addItem('Update Week 4 Data', 'updateWeek4')
    .addItem('Update All Weeks Data (Slower)', 'updateAllWeeks')
    .addToUi()
}

function updateWeek1() {
  var day = 1;
  updateWeekData(day)
}

function updateWeek2() {
  var day = 8;
  updateWeekData(day)
}

function updateWeek3() {
  var day = 15;
  updateWeekData(day)
}

function updateWeek4() {
  var day = 22;
  updateWeekData(day)
}

function updateAllWeeks() {
  updateWeek1();
  updateWeek2();
  updateWeek3();
  updateWeek4();
}

function updateWeekData(day) {
  var token = getAndPersistToken();
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = sheet.getSheetName();
  var month = getMonthNumber(sheetName);
  var year = getYearNumber(sheetName);
  
  
  if (month === -1 || isNaN(year)) return;
  
  var start_date = new Date(Date.UTC(year,month,day))
  var week = Math.min(Math.ceil(start_date.getUTCDate()/7),WEEKS_SIZE);
  var end_date = new Date(start_date)
  if(day<21) {
    end_date.setUTCDate(start_date.getUTCDate()+6);
  }
  else {
    //Set the end of the month as end date
    end_date.setUTCMonth(end_date.getUTCMonth()+1);
    end_date.setUTCDate(0);
  }

  if(token != null) {
    var expa_data = PRODUCTS.reduce(function (acc,product){
      var data = getProduct({
        token: token,
        start_date: Utilities.formatDate(start_date,"GMT","yyyy-MM-dd"),
        end_date: Utilities.formatDate(end_date,"GMT","yyyy-MM-dd"),
        program: product.program,
        type: product.type,
      },product);
      sheet.toast('Retrieved '+product.name+' from EXPA.');
      return acc.concat(data);
    },[]);
    
    
    // After retrieving the data from EXPA, populate it all together in the achievement part of the sheet
    
    var achRange = sheet.getDataRange();
    var formulas = achRange.getFormulas();
    var values = achRange.getValues();
    
    console.log(expa_data);
    
    // For each EY in the list, get the related product data and put them in the values array
    for(i=ROW_OFFSET; i<values.length; i++) {
      var ey = values[i];
      var eyName = ey[COL_NAME];
      var expaId = ey[COL_ID];
      var matching = expa_data.filter(function (el){ return expaId == el.key; });
      
      Logger.log(expaId);
      
      matching.forEach(function (data){
        var objKeys = Object.keys(data);
        objKeys.forEach(function (key){
          if(key !== 'key') {
            var col = getColByKeyAndWeek(key, week);
            values[i][col] = data[key];
          }
        })
      });
    }
    
    // Copy formulas as they are into the values array, so we can preserve them when writing back to the spreadsheet
    for(i=ROW_OFFSET; i<formulas.length; i++) {
      for(j=0; j<formulas[i].length; j++) {
        if(formulas[i][j]) {
          values[i][j] = formulas[i][j];
        }
        if(values[i][j] === '') {
          values[i][j] = 0;
        }
      }
    }
    
    sheet.toast('Copying Week '+week+' information into spreadsheet.');
    achRange.setValues(values);
  }
}

function getMonthNumber(sheetName) {
  var months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return months.indexOf(sheetName.substr(0,3));
}

function getYearNumber(sheetName) {
  return 2000 + parseInt(sheetName.substr(3,2));
}

function getColByKeyAndWeek(key, week) {
  //key is of the form: prod_stage (e.g.: iGV_co)
  var prods = PRODUCTS.map(function (prod){ return prod.name; });
  var prod_offsets = PRODUCTS.map(function (prod){ return prod.offset; });
  var stage = ['apd', 're', 'co'];
  
  return COL_OFFSET+prod_offsets[prods.indexOf(key.substr(0,3))]+STAGE_OFFSET*stage.indexOf(key.substr(4))+week-1;
}

function getAndPersistToken() {
  var properties = PropertiesService.getScriptProperties();
  var token = properties.getProperty('token');
  
  if(token != null) {
    if(!isValidToken(token)) {
      PropertiesService.getScriptProperties().deleteProperty('token'); //If token has expired, delete from Properties Service
      token = null;
    }
  }
  
  if(token == null) {
    token = getTokenFromUser();
    
    if(token != null) {
      properties.setProperty('token',token);
    }
  }
  
  return token;
}

function isValidToken(token) {
  try {
    var response = UrlFetchApp.fetch("https://gis-api.aiesec.org/v2/people/my.json?only=facets&access_token="+token); //Gets url
  } catch(err) {
    return false;
  }
  return true;
}

function getProduct(options,product) { //options:JSON (with token,program,type,start_date,end_date)
  var url = 'https://gis-api.aiesec.org/v2/applications/analyze.json?'
    + 'access_token=' + options.token
    + '&basic[home_office_id]=' + MC_CODE
    + '&basic[type]='+options.type // person | opportunity
    + '&start_date='+options.start_date
    + '&end_date='+options.end_date
    + '&programmes[]='+options.program // 1=>GV | 2=>GT | 5=>GE

    Logger.log(url)
    try {
      var data = getURL(url);
      
      //Data format for analytics (not the whole structure though):
        //analytics
        //|__ children
        //    |__ buckets: Array
        //         |-- key
        //         |-- total_approvals
        //         |   |__doc_count: int
        //         |-- total_realized (ídem)
        //         |__ total_completed (ídem)
        
      var lc_data = data.analytics.children.buckets;
      var lcjson;
      var lc_res = [];
      
      var lcvar;
      for(var i=lc_data.length-1 ; i>=0 ; i--) {
        lcvar = lc_data[i];
        
        lcjson = { key: lcvar.key };
        lcjson[product.name+"_apd"] = lcvar.total_approvals.doc_count;
        lcjson[product.name+"_re"] = lcvar.total_realized.doc_count;
        lcjson[product.name+"_co"] = lcvar.total_completed.doc_count;
        
        lc_res.push(lcjson);
      }
      return lc_res;
    } catch(err){
      throw err;
    }
}

//TO-DO:
//response throws error if response code is not 200, fix this to catch it and return it as cb
function getURL(url) { 
  var response = UrlFetchApp.fetch(url); //Gets url
  var responseCode = response.getResponseCode();
  
  if(responseCode == 200) { //If retrieve is successful
    return JSON.parse(response.getContentText()); //Callback with parsed data
  }
  else { // There was an error (most likely 404 or 401, because #EXPA)
    throw Error('EXPA had an error. Response code was '+response.getResponseCode());
  }
}

function getTokenFromUser() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt('Update with data from EXPA','Please enter you access token:',ui.ButtonSet.OK_CANCEL);
  var pressedBtn = response.getSelectedButton();
  
  if(pressedBtn == ui.Button.OK) {
    return response.getResponseText();
  }
  return null;
}
