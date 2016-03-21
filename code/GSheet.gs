//Google Salesforce Connector
//google apps script    code.gs
//Last update:  03/31/2015
//QUNIT:MxL38OxqIK-B73jyDTvCe-OBao7QLBR4j

/////////////////////  ADD CONFIGURATION VARIABLES //////////////////////////////////////////////////
var SF_VERSION = "31.0";
var CLIENT_ID = "3MVG9ZL0ppGP5UrBuJeCR0K0k9Ziiyol6mIuaFx257OgCf2nQ7D09r.NVmuBza.mZBGGFfLUdvBzWrkQNvprX";
var CLIENT_SECRETID = "5440839150561936962";
//var REDIRECT_URI = ScriptApp.getService().getUrl();
var REDIRECT_URI ="https://script.google.com/macros/s/AKfycbwhbI4Dv7E4MOxkJYomdapiFyEe8n7rtHTSpK-27LEnXNkx5jY/exec";

//////////////////////////////////////////////////////////////////////////////////////////////////////
//Added Name field to the ConfigNames for use in Query Dialog .
var CONFIG_NAMES = ["Name","Query","Active","Sheet","Output Mode","Parent","Child","Chunk Size","Chunk Column","Total Column Name","Line Label","Line Formula","Line Label","Line Formula"];
var NUMOFCONFIGS = CONFIG_NAMES.length;
//var QUERY_SHEET= 'https://docs.google.com/spreadsheets/d/1JiUgwgFyys47k-7mXJAmrMij81w1w67vdlM3J9_li_g/edit';
var QUERY_SHEET= 'https://docs.google.com/spreadsheets/d/1jEzDWTuZJZ73yRCBCt9F7ywnorpFVgNL3XMkJ163A_s/edit?usp=sharing';

var GSheet = SpreadsheetApp.getActiveSpreadsheet();

//Storing of Session Tokens at User Level.
var userProperties = PropertiesService.getUserProperties();
var Queries;
var initQueries;
var SF = new SForce(CLIENT_ID, CLIENT_SECRETID, SF_VERSION, REDIRECT_URI);
var RM;
var sfReports;
var editing;

// For Error Handling 
var ErrorHandler = new ErrorLog();

function doGet(e){
  //send a message to window when SF calls back
  SF = new SForce(CLIENT_ID, CLIENT_SECRETID, SF_VERSION, REDIRECT_URI);

  if(typeof(e) != "undefined" && typeof(e.parameters.code) != "undefined" ){
    var params = e.parameters.code;    
   
    userProperties.setProperty("code",params);
    SF.connectSFAuth(params);    
  
    return HtmlService.createHtmlOutput("Salesforce Authentication Complete. You may close this window.");
               
  } else {

    if(CLIENT_ID != null && CLIENT_SECRETID != null){

     //Retrieving from User Level Properties
     if( userProperties.getProperty("sessionId") != null && userProperties.getProperty("instance_url") != null && userProperties.getProperty("code") != null ){
    
      onOpen();
        GSheet.toast("Welcome to the FinancialForce Connector", "GSC");
      }  else {
        GSheet.toast("Please Login to Salesforce!", "GSC", 2000 );
      }
    }
  }

}


function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  
  userProperties.setProperty("currentSheetId", SpreadsheetApp.getActiveSpreadsheet().getSheetId());
   
  _populateQueriesCache();
  
  var amenu = SpreadsheetApp.getUi().createAddonMenu();
  amenu.addItem('FinancialForce Connector', 'onOpen');
  amenu.addToUi();

  try{
    _createQueryListSheet();
    Queries =  _populateSidebarQueue();
  } catch(e){
    Logger.log( "Error " + e.message  );
  }

  showSidebar();

}


function _populateQueriesCache(){
  initQueries= _getQueryListFromSheet();
  CacheService.getDocumentCache().put("len",initQueries.length);
  for( var i = 0; i < initQueries.length; i++){
    CacheService.getDocumentCache().put("query_"+i,initQueries[i]);
  }
  return initQueries;
}

function _getQueriesCache(){
  /* Commented out, setting up the cache during onOpen was proving unreliable, giving not authorised errors
  var tlen = CacheService.getDocumentCache().get("len");
  if(tlen==null || tlen === undefined){
    return _populateQueriesCache();
  }
  */
  _populateQueriesCache();
  var tlen = CacheService.getDocumentCache().get("len");
  var len = parseInt(tlen);

  initQueries = [];
  for( var i = 0; i < len; i++){
    var query = CacheService.getDocumentCache().get("query_"+i)
    initQueries.push(query);
  }
  return initQueries;
}

function showRemoveQuery(i){
  GSheet.getSheetByName("Querylist").deleteRow(i);
  refreshQueue();
}

function showSidebar() {
    
  var html = HtmlService.createTemplateFromFile('sidebar')
  .evaluate();
  html.setTitle("FinancialForce Connector");  
  
  //Setting sandbox Mode
  //If Session is Present,then open the HTML in IFRAME Mode , as Login UI not working in IFRAME Mode.
  if(validateSFsession())
   html.setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  SpreadsheetApp.getUi().showSidebar(html);  
    
}


function showQueueDialog() {
  initQueries = _getQueriesCache();
  editing = false;
  var html = HtmlService.createTemplateFromFile('queue')
  .evaluate();
  html.setWidth(400);
  html.setHeight(480);

  //Setting sandbox Mode
  html.setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  SpreadsheetApp.getUi()
  .showModalDialog(html, 'Add Query to Queue');
}

function showEditQueueDialog(i) {
  initQueries = _getQueriesCache();
  editing = true;
  position = i;
  Queries =  _populateSidebarQueue();
  var edited = i;
  var html = HtmlService.createTemplateFromFile('queue')
  .evaluate();
  html.setWidth(400);
  html.setHeight(480);
  
  //Setting sandbox Mode
  html.setSandboxMode(HtmlService.SandboxMode.IFRAME);

  SpreadsheetApp.getUi()
  .showModalDialog(html, 'Edit Query');
}


function logoutGSC(){
  SF.disconnectSFAuth();
  GSheet.toast("Logging out of FinancialForce Connector!", "GSC");

//clear/delete Report sheets
  var reportsheet = GSheet.getSheetByName("ReportsList");
  if( reportsheet) {
    reportsheet.clearContents();
    GSheet.deleteSheet(reportsheet);
  }
  
  //Calls this to reload the Sidebar.
  showSidebar();
}


function _createQueryListSheet(){
  var newSheet = null;
  var headersRange = null;

  //Create the Query Sheet
  if(GSheet.getSheetByName("QueryList")){

  } else {

    newSheet = GSheet.insertSheet("QueryList");

    //creates the query ueue
    headersRange = newSheet.getRange(1, 1,1,2);
    headersRange.mergeAcross().setValue("Query Queue");
    headersRange.setBackground("#00396B");
    headersRange.setFontSize("10");
    headersRange.setFontColor("white");


    var range_for_names=newSheet.getRange(2,1,2,NUMOFCONFIGS);
    for( var i =0; i < NUMOFCONFIGS; i++){
      range_for_names.getCell(1, i+1).setValue(CONFIG_NAMES[i]).setBackground("#0C8EFF").setFontSize("9").setFontColor("#ffffff");
    }

  }
}//end _createQueryListSheet

function processForm(formdata){
  GSheet.toast("Processing Queue!", "Queue");
  var qdata = [];
  var position = -1;
      
  // Checking if user has either added a new Query or Selected a Query 
  if(formdata.selQuery == "Select a Query"  || (formdata.selQuery == "Add New Query" && formdata.addQuery.trim() == "")){
    GSheet.toast("ERROR:  You must either Select a Query or Add a Query !", "ERROR");
  } else {
    if(formdata.edit !== undefined){
      position = parseInt(formdata.edit);
    }
    for each(var f in formdata){
      if( f != ""){
        qdata.push( f );
      } else {
        qdata.push('null');
      }
    }//end for
    if (  _addToQueue(qdata,position) ){
      GSheet.toast("Done - Query added to Queue!", "Queue");
    }
    refreshQueue();
  }

}

function _addToQueue(qdata,position){
  var activeSheet = GSheet.getSheetByName("QueryList");
  var queue;
  if(position==-1){
    var lastRow = _getLastRowInColumn("QueryList", "A1:A");// last row in Queue
    queue = activeSheet.getRange(lastRow,1);  //range of last row in Queue (Col A)
  }else{
    queue = activeSheet.getRange(position,1);  //range of last row in Queue (Col A)
  }
  var numofConfigs = NUMOFCONFIGS;
  var j=0;
  
  //User has either Added a Query or Selected a Query from the Drop-down List
 // if( qdata[1] != 'Add New Query'){
 //   qdata[2] = qdata[1];}
      
  //find last row in Queue ( Col A)
    
  var j=0;
    
  for( var i = 0; i <numofConfigs; i++){
    
    //To accordingly Iterate through form Data
    if(qdata[i] == 'Add New Query' && i==1)
     j=j+1;
      
    if( qdata[j] != 'null' && typeof(qdata[j]) != 'undefined'){
      queue.offset(0,i).setValue(qdata[j]).setFontWeight("bold").setFontSize(9).setFontColor("#00396b");
    } else {
      queue.offset(0,i).setValue('null').setFontColor("#00396b").setFontSize(9);
    }
    
    if(i==1)
     j=3;
    else
     j++;
  }

  return true;
}



function _getLastRowInColumn(sheetName, colRange){

  try{
    var activeSheet = GSheet.getSheetByName(sheetName);
    var colums = activeSheet.getRange(colRange);
    var lastRowRange = colums.getLastRow();  //range of last row in column Range
    for(var i=1; i < lastRowRange; i++){
      var myrange = colums.getCell(i,1).getValue();
      if( myrange == ""){
        return i;
      }
    }
  } catch(e){
    return;
  }

}

function _getQueryListFromSheet(){
  try{
    var querySheet = SpreadsheetApp.openByUrl(QUERY_SHEET).getSheetByName("Sheet1");
    var lastRow = querySheet.getLastRow();// _getLastRowInColumn("QueryList","N1:N");  //Col D
    var queries = [];

    for( var i =1; i <= lastRow; i++){
      var cellValue = querySheet.getRange(i,1).getValue();
      if( cellValue != "" && cellValue != "Query List"){
        queries.push( cellValue );
      }
    }
  }catch(e){
    Logger.log("Error: "+e);
    var queries = ["Error loading queries" + e];
  }
  return queries;
}

function refreshQueue(){
  GSheet.toast("Refreshing the QUEUE!");
  Queries = _populateSidebarQueue();
  showSidebar();
}

function _populateSidebarQueue(){
  var qries = [];

  try{
    var activeSheet = GSheet.getSheetByName("QueryList");
  } catch(e){
    return null;
  }

  var lastRow = _getLastRowInColumn("QueryList","A1:A");  //Col A

  for( var i =3; i < lastRow; i++){
    //var cellValue = activeSheet.getRange(i,1).getValue();
    //if( cellValue != "***" && cellValue != "Query Queue" && cellValue != "Query Queue" && ){
    
    //Name of the Query
    var queryName = activeSheet.getRange(i,1).getValue();
    var querytxt = activeSheet.getRange(i,1).offset(0,1).getValue();
    var statustxt = activeSheet.getRange(i,1).offset(0,2).getValue();
    var sheetnametxt =activeSheet.getRange(i,1).offset(0,3).getValue();
    var outputmode = activeSheet.getRange(i,1).offset(0,4).getValue();
    var parentef =activeSheet.getRange(i,1).offset(0,5).getValue();
    var childef =activeSheet.getRange(i,1).offset(0,6).getValue();
    var chunksz =activeSheet.getRange(i,1).offset(0,7).getValue();
    var chunkcol =activeSheet.getRange(i,1).offset(0,8).getValue();
    //template
    var templatetotal = activeSheet.getRange(i,1).offset(0,9).getValue();
    var templateLabel1 = activeSheet.getRange(i,1).offset(0,10).getValue();
    var templateFormula1 = activeSheet.getRange(i,1).offset(0,11).getValue();
    var templateLabel2 = activeSheet.getRange(i,1).offset(0,12).getValue();
    var templateFormula2 = activeSheet.getRange(i,1).offset(0,13).getValue();

    if(templateFormula1!="null"){
      templateFormula1="="+templateFormula1;
    }
    if(templateFormula2!="null"){
      templateFormula2="="+templateFormula2;
    }


    //Added Name field 
    var myQueryObj = { num: i, qname:queryName, query: querytxt, status: statustxt, sheetName: sheetnametxt, qMode:outputmode, parentExtField: parentef, childExtField: childef, chunkSize: chunksz, chunkCol: chunkcol,
      templateTotalColumn: templatetotal, templateLabel1:templateLabel1, templateFormula1: templateFormula1 ,templateLabel2:templateLabel2, templateFormula2: templateFormula2 };

    qries.push(myQueryObj);

    //i = i + NUMOFCONFIGS;
    //}
  }
  return qries;
}

function pullSFdata(){
  //method to display the data pulled from Salesforce
  GSheet.toast("In Progress - Pulling Salesforce Data!", "Pull");
  Queries =  _populateSidebarQueue();


  var activeQueries = [];
  for(var i=0; i < Queries.length; i++){
    if( Queries[i].status == 'Yes' ){
      var qObj = new Query(Queries[i].query, Queries[i].sheetName, Queries[i].status, Queries[i].qMode);
      qObj.setParams(Queries[i].query);
      activeQueries.push( qObj );
    }
  }
  if(activeQueries.length > 0){
    //clear or insert result sheet
    //show data
    //insert results sheet if it doesn't exists
    for(var i=0; i < activeQueries.length; i++){

      var resultSheetName = activeQueries[i].resultSheetName;

      if(GSheet.getSheetByName(resultSheetName)){
        GSheet.getSheetByName(resultSheetName).clear();  //clear sheet before new results
        newSheet = GSheet.setActiveSheet(GSheet.getSheetByName(resultSheetName));
        newSheet.activate();
      } else {
        newSheet = GSheet.insertSheet(resultSheetName);
      }
      newSheet.getRange(2,1,newSheet.getMaxRows()-1,newSheet.getMaxColumns()).setFontSize("8");
      Utilities.sleep(3000);  //pause before process next sheet
      _showResultData(activeQueries[i]);

    }//end for

    Utilities.sleep(3000);
    GSheet.toast("Done - Pulling Salesforce Data!");

  } else {
    GSheet.toast("No Active Queries!", "Pull");
  }

} //end pull



//method to get the Data from Salesforce to show on spreadsheet
function _getData(query){
  var data = null;
  SF = new SForce(CLIENT_ID, CLIENT_SECRETID, SF_VERSION, REDIRECT_URI);

  try{
    data = SF.querySF(query);
    if( data.errorCode ){
      GSheet.toast( data.errorCode );
    }

    ////---testing only  ---var data = {"totalSize":26,"done":true,"records":[{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000WIWvMAAX"},"Id":"001i000000WIWvMAAX","Name":"ANEW COMPANY","Description":"A new company/business/establishment"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000b1R9wAAE"},"Id":"001i000000b1R9wAAE","Name":"PeterRobbins, Inc.","Description":"A PAULIE Robbins Company"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000WIZXpAAP"},"Id":"001i000000WIZXpAAP","Name":"Cloud Inc.","Description":"CloudInc is a company."},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000b1RGUAA2"},"Id":"001i000000b1RGUAA2","Name":"Baskin Robbins, Inc.","Description":"A Peter Robbins Company"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000XbDtKAAV"},"Id":"001i000000XbDtKAAV","Name":"CloudForce","Description":"A CloudForce company"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000VMU3QAAX"},"Id":"001i000000VMU3QAAX","Name":"RefAlert Inc","Description":"Organizing Referees and Game Assignments"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000VN53hAAD"},"Id":"001i000000VN53hAAD","Name":"MikaSue Designs","Description":"Your Personal IT Department"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000b1RZZAA2"},"Id":"001i000000b1RZZAA2","Name":"Nerf Inc","Description":"Nerf"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000b1RZaAAM"},"Id":"001i000000b1RZaAAM","Name":"Super Nerf","Description":"Super Nerf Co"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000XaiaGAAR"},"Id":"001i000000XaiaGAAR","Name":"American Water","Description":"Turn Water Off When Brushing Teeth"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000XaiZIAAZ"},"Id":"001i000000XaiZIAAZ","Name":"Brother Enterprises","Description":"A Printing Company"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000XaiOWAAZ"},"Id":"001i000000XaiOWAAZ","Name":"MikaSue Designs Inc.","Description":"Another IT Company"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000Xaj1HAAR"},"Id":"001i000000Xaj1HAAR","Name":"Commandment Co.","Description":"Ten little monkeys"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000UvtWOAAZ"},"Id":"001i000000UvtWOAAZ","Name":"Chicago Public League","Description":"Connecting Officials and Assignors Alike"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000dYPBUAA4"},"Id":"001i000000dYPBUAA4","Name":"CompanyB70","Description":"DescriptionB70"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000VHHPpAAP"},"Id":"001i000000VHHPpAAP","Name":"Gene Point Dynamics","Description":"Genomics company engaged in mapping and sequencing of the human genome and developing gene-based drugs"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000VHHPqAAP"},"Id":"001i000000VHHPqAAP","Name":"United Oil & Gas, United Kingdom","Description":"A United Kingdom company of establishments"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000VHHPrAAP"},"Id":"001i000000VHHPrAAP","Name":"United Oil & Gas, Singapore","Description":"Singapore, singing to the poor"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000VHHPsAAP"},"Id":"001i000000VHHPsAAP","Name":"Edge Case Communications","Description":"Edge Case"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000VHHPtAAP"},"Id":"001i000000VHHPtAAP","Name":"Burlington Textiles Corp of America","Description":"Burlington Company of American Textiles"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000VHHPuAAP"},"Id":"001i000000VHHPuAAP","Name":"Pyramid Construction Inc.","Description":"Egyption Pyramids Construction"},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000VHHPwAAP"},"Id":"001i000000VHHPwAAP","Name":"Grand Hotels & Resorts Ltd","Description":"Chain of hotels and resorts across the US, UK, Eastern Europe, Japan, and SE Asia."},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000VHHPxAAP"},"Id":"001i000000VHHPxAAP","Name":"Express Logistics and Transport","Description":"Commerical logistics and transportation company."},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000VHHPyAAP"},"Id":"001i000000VHHPyAAP","Name":"U OF A","Description":"Leading university in AZ offering undergraduate and graduate programs in arts and humanities, pure sciences, engineering, business, and medicine."},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000VHHPzAAP"},"Id":"001i000000VHHPzAAP","Name":"United Oil & Gas Corp.","Description":"World's third largest oil and gas company."},{"attributes":{"type":"Account","url":"/services/data/v29.0/sobjects/Account/001i000000XaiWiAAJ"},"Id":"001i000000XaiWiAAJ","Name":"MikaSue Learning Center","Description":"A Training Establishment"}]};
    //GSheet.toast("Returning " + data.totalSize + " data rows!", "SALESFORCE");
  } catch(e){
    if( e.errorCode ){
      GSheet.toast( JSON.stringify(e.errorCode) );

      // Call Error Handler 
      ErrorHandler.LogError(JSON.stringify(e.errorCode));
     
      userProperties.deleteAllProperties();
      var ui = SpreadsheetApp.getUi();
      ui.close();
    }
    return false;
  }

  if( typeof(data.message) != 'undefined' ){
    GSheet.toast(JSON.stringify(data.message) + " Did you login?", "SALESFORCE");
  }

  return data;

}



//method to get Data that will be pushed to Salesforce
/* parameter indicates if push chunks or push records button was pressed */
function pushSFData(value){
  GSheet.toast("Pushing data .... ");
  Queries =  _populateSidebarQueue();
  var response;
  var childResponse;
  var parentResponse;
  var resultSheet;
  var parentObj;
  var qcHeaders;
  var activeQueries = [];

  //Determines Active Queries
  for(var i=0; i < Queries.length; i++){
    if( Queries[i].status == 'Yes' ){
      var qObj = new Query(Queries[i].query, Queries[i].sheetName, Queries[i].status, Queries[i].qMode);
      var parent =  ( Queries[i].parentExtField != "null") ? Queries[i].parentExtField : null;
      var child =  ( Queries[i].childExtField != "null") ? Queries[i].childExtField : null;

      if(parent != null){
        qObj.qParentObj = parent.split('.')[0];
        qObj.qExtParentObj = parent.split('.')[1];
      }
      if(child != null){
        qObj.qChildObj = child.split('.')[0];
        qObj.qExtChildObj = child.split('.')[1];
      }

      qObj.chunkSize = Queries[i].chunkSize;
      qObj.chunkCols = Queries[i].chunkCol;
      qObj.setParams(Queries[i].query);

      activeQueries.push(qObj);
    }
  }

  var anyChunkQuery = "-1";
  //Process Active Queries
  if(activeQueries.length > 0){

    for(var i=0; i < activeQueries.length; i++){
      if(activeQueries[i]['qMode'] === 'Chunker' && value === 'true'){
        Logger.log("Chunker Pushing multiple in pushSF " + parent + " | " + child);
        anyChunkQuery = "0";
        //push chunks button was pressed if true
        previewChunks(true);
        break;
      } else if (value === 'true' && activeQueries[i]['qMode'] !== 'Chunker'){
        GSheet.toast("Not a Chunking Query!");

      } else if (activeQueries[i]['qMode'] == 'Multiple'){
        Logger.log("Multiple Pushing multiple in pushSF " + parent + " | " + child);
        pushMultipleData( activeQueries[i], 'split');
      } else {  //Tabular or chunker
        if( activeQueries[i]['qchildColumns'].length > 0 ){
          Logger.log("more Pushing multiple in pushSF " + parent + " | " + child);
          pushMultipleData( activeQueries[i] );
          
          
        } else {
          Logger.log("Tab Pushing multiple in pushSF " + parent + " | " + child);

          pushTabularData( activeQueries[i] );
        }
      }
    }
  }

  //Calling Pull Records 
  pullSFdata();
  
  // If Any Query is Chunk , then call previewChunks as well. 
  if(anyChunkQuery == "0")
   previewChunks(false); 

} //end push



function pushTabularData(tabQuery, mode){
  GSheet.toast("In Progress - Pushing Salesforce Tabular Data!", "SALESFORCE");
  var parentObj = tabQuery.qparentObj;
  var  qcHeaders = tabQuery.qcolumns;
  var resultSheet = null;

  //set sheet to work on
  if(  mode === 'chunkerOnly' )
    resultSheet = GSheet.getSheetByName("PreviewChunks_" + tabQuery.resultSheetName );

  else
    resultSheet = GSheet.getSheetByName(tabQuery.resultSheetName);

  var mystartRange = resultSheet.getRange(2,1);
  var myendRange = resultSheet.getRange(GSheet.getLastRow(), GSheet.getLastColumn());
  var myrange = resultSheet.getRange(mystartRange.getA1Notation() + ":" + myendRange.getA1Notation());
  var myobjs, childrows, parentrows;
  var queryData = tabQuery.rawQuery;

  myobjs = getRowsData(resultSheet, myrange, 1);

  if(resultSheet.getLastRow() < 100){

    if(qcHeaders.indexOf("Id") != -1 ){  //if we don't find an ID column use external field ID
      try{
        response = SF.sendSF(tabQuery, myobjs);
        if( typeof(response.error) != 'undefined' ){
          GSheet.toast("ERROR!" + response.error );
          
          userProperties.deleteAllProperties();
        }
      } catch(e){
        Logger.log("Error " + e );
        // Call Error Handler
        ErrorHandler.LogError(e);
      }
      GSheet.toast( response.message );

    } else {
      GSheet.toast("Pushing Ext Field","Salesforce Push");
      response = SF.sendSFext(tabQuery, "Parent", myobjs);
    }

  } else {
    //use bulk mod
    GSheet.toast("Creating Batch Job!","Salesforce Push");
    response = SF.sendBulkSF(myobjs, parentObj);
  }
}



function pushMultipleData(multipleQuery, mode){
  GSheet.toast("Pushing Multiple Salesforce Data","Salesforce Push");
  var parentObj = multipleQuery.qparentObj;
  var childObj = multipleQuery.qChildObj;
  var  qcHeaders = multipleQuery.qcolumns;
  var resultSheetName = multipleQuery.resultSheetName;
  var childrows, parentrows, parentResponse, childResponse;
  var resultSheet, presultSheet, cresultSheet, chunkerresultSheet;
  var myRanges, childrange, parentrange, childchunkrange, parentchunkrange;

  if(  mode === 'chunkerOnly' ){
    resultSheet = GSheet.getSheetByName("PreviewChunks_" + resultSheetName );
  } else if ( mode === 'split' ){
    presultSheet = GSheet.getSheetByName(resultSheetName);
    cresultSheet = GSheet.getSheetByName(resultSheetName.concat("_Relationships") );

  } else {

    resultSheet = GSheet.getSheetByName(resultSheetName);
  }

  if( mode !== 'split' ){
    //determine parent range and child ranges  if queryObj mode is multiple pass parent then child
    myRanges = getParentChildRange(resultSheet);
    childrange = resultSheet.getRange(myRanges.child.getA1Notation());
    parentrange = resultSheet.getRange(myRanges.parent.getA1Notation());
  }



  if(childObj === null){
    //multiple will always use bulk API
    var tempchildObj = multipleQuery.qchildObj;
    //SALESFORCE RELATIONSHIP RULES  if in plural form, change to singular for push
    //Custom objects endin with __c and __r
    if( tempchildObj.charAt(tempchildObj.length-1) == 'r' && tempchildObj.charAt(tempchildObj.length-2) == '_'){
      tempchildObj = tempchildObj.slice(0,tempchildObj.length-3);
      if( tempchildObj.charAt(tempchildObj.length-1) == 's' && tempchildObj.charAt(tempchildObj.length-2) == 'e' && tempchildObj.charAt(tempchildObj.length-3) == 'i'){
        childObj = tempchildObj.slice(0,tempchildObj.length-3) + 'y';
      }else if( tempchildObj.charAt(tempchildObj.length-1) == 's'){
        childObj = tempchildObj.slice(0,tempchildObj.length-1);
      }
      childObj = childObj + "__c";
    }
    else if( tempchildObj.charAt(tempchildObj.length-1) == 's' && tempchildObj.charAt(tempchildObj.length-2) == 'e' && tempchildObj.charAt(tempchildObj.length-3) == 'i'){
      childObj = tempchildObj.slice(0,tempchildObj.length-3) + 'y';
    }else if( tempchildObj.charAt(tempchildObj.length-1) == 's'){
      childObj = tempchildObj.slice(0,tempchildObj.length-1);
    }else {
      childObj = tempchildObj;
    }
  }

  if( mode === 'split') {
    crangeEnd = cresultSheet.getRange(cresultSheet.getLastRow(), cresultSheet.getLastColumn());
    prangeEnd = presultSheet.getRange(presultSheet.getLastRow(), presultSheet.getLastColumn());

    childrange = cresultSheet.getRange("A2:" + crangeEnd.getA1Notation());
    parentrange = presultSheet.getRange("A2:" + prangeEnd.getA1Notation());

    childrows = getRowsData(cresultSheet, childrange, 1);
    parentrows = getRowsData(presultSheet, parentrange, 1);
  }else {
    childrows = getRowsData(resultSheet, childrange, 1);
    parentrows = getRowsData(resultSheet, parentrange, 1);
  }

  GSheet.toast("Bulk Push - Creating Job", "Salesforce Push");
  parentResponse =SF.sendBulkSF(parentrows, parentObj);

  if( childrows !== null )
    childResponse = SF.sendBulkSF(childrows, childObj);



  GSheet.toast( parentResponse );
  Utilities.sleep(1000);

  if( childResponse != null){
    GSheet.toast( childResponse );
    Utilities.sleep(1000);
  }

  GSheet.toast("Bulk Push Done Successful", "Salesforce Push");

}






//method to show the data on the spreadsheet in the appropriate sheet
//input is a query object
function _showResultData(myQueryObj){
  var dataObjs = [];
  var headers = myQueryObj.qcolumns;
  var printMode;
  var myResultSheet = myQueryObj.resultSheetName;

  myDataObj = _getData(myQueryObj.rawQuery);  //get data from SF



  if( myQueryObj.qMode == 'Multiple' ){
    dataObjs = _getPCSheetRows(myDataObj, myQueryObj);
    printMode = 'split';
    headers = { p: myQueryObj.qparentColumns, c: myQueryObj.qchildColumns };
  } else if( (myQueryObj.qMode == 'Tabular' || myQueryObj.qMode == 'Chunker') &&  myQueryObj.qchildColumns.length == 0){
    dataObjs =  _getSheetRows(myDataObj, headers);
    printMode = 'single';
  } else {
    dataObjs = _getSheetRowsWithChildren(myDataObj, myQueryObj);
    printMode = 'multiple';
  }
  _printSheetRows( myResultSheet, dataObjs, headers, printMode);
}


function _printSheetRows(resultSheet, dataToPrint, hdrs, mode){
  var rsheet, csheet;
  var headersRange, pheadersRange, cheadersRange;

  if( GSheet.getSheetByName(resultSheet) ){
    rsheet = GSheet.getSheetByName(resultSheet).activate();
  } else {
    rsheet = GSheet.insertSheet(resultSheet);
  }
 rsheet.getRange(2,1,rsheet.getMaxRows()-1,rsheet.getMaxColumns()).setFontSize("8");
  if( mode == 'split' ){
    if( GSheet.getSheetByName(resultSheet.concat("_Relationships") ) ){
      GSheet.getSheetByName(resultSheet.concat("_Relationships")).clear();  //clear sheet before new results
      csheet = GSheet.getSheetByName(resultSheet.concat("_Relationships") );
    } else {
      csheet = GSheet.insertSheet(resultSheet.concat("_Relationships"));
    }
    csheet.getRange(2,1,csheet.getMaxRows()-1,csheet.getMaxColumns()).setFontSize("8");
  }


  if(hdrs != null && mode != 'split'){
    
    headersRange = rsheet.getRange(1, 1, 1, hdrs.length);
    headersRange.setValues([hdrs]).setFontSize("9").setBackground("#0C8EFF").setFontColor("#ffffff");
  } else if( mode == 'split'){

    pheadersRange = rsheet.getRange(1, 1, 1, hdrs.p.length);
    pheadersRange.setValues([hdrs.p]).setFontSize("9").setBackground("#0C8EFF").setFontColor("#ffffff");

    cheadersRange = csheet.getRange(1, 1, 1, hdrs.c.length);
    cheadersRange.setValues([hdrs.c]).setFontSize("9").setBackground("#0C8EFF").setFontColor("#ffffff");

  }


  if( mode == 'multiple'){
    for(var x =0; x < dataToPrint.length; x++){
      parentNoChild = false;
      for(var y =0; y < dataToPrint[x].length; y++){
        if(typeof(dataToPrint[x][y]) === 'object')
          rsheet.appendRow(dataToPrint[x][y]);
        else
          parentNoChild = true;
      }
      if(parentNoChild)
        rsheet.appendRow(dataToPrint[x]);
    }
  } else if( mode == 'split'){
    parentRows = dataToPrint.parent;
    childRows = dataToPrint.child;
    _printSheetRows(rsheet.getName(),parentRows, null, "single");
    _printSheetRows(csheet.getName(),childRows, null, "single");

  } else {
    for(var i =0; i < dataToPrint.length; i++){
      rsheet.appendRow(dataToPrint[i]);
    }
  }
 
}//end function

//method to parse the rows returned from Salesforce Query in Tabular mode
function _getSheetRows(queryResultData, hdrs){
  var found = false;
  var temprows = [];
  var headers = hdrs;  //local var

  //for each obj in the returned data
  for( var obj in queryResultData.records ){
    var temparr = [];  //a temporary array to hold values until push to dataObj array
    for( var col in queryResultData.records[obj]){
      //search for the current key in the headers and push value to temparr
      for( var i=0; i < headers.length; i++){
        if(col == headers[i] ){
          temparr.push( queryResultData.records[obj][col] );
          found = true;
          break;
        }
        if(found) continue;
      }
    }
    temprows.push(temparr);
  }

  return temprows;

} //end getSheetRows func



//function to split parent child results
function _getPCSheetRows(queryResultData, qyObj){
  var dataRows = [];
  var prows = [];
  var crows = [];
  var found = false;
  var hdrs = qyObj.qcolumns;
  var chdrs = qyObj.qchildColumns;
  var childObjName = qyObj.qchildObj;
  var parentObj = queryResultData.records;
  var pObjct = queryResultData.totalSize;

  //get parent data
  //get child data

  for( var p=0; p < pObjct; p++){

    //push parent obj values
    var parentTempArray = [];
    for( var col in parentObj[p]){
      //search for the current key in the headers and push value to temparr
      for( var i=0; i < hdrs.length; i++){
        if(col == hdrs[i]){
          parentTempArray.push( parentObj[p][col] );
          found = true;
          break;
        }
        if(found) continue;
      }
    }
    prows.push( parentTempArray );


    var child = parentObj[p][childObjName];
    if( child != null ) {

      //get childRows
      var childRecs = child.records;
      var childRecsSz = childRecs.length;

      if(childRecsSz > 0){
        var tempChildRows = [];

        for(var i=0; i < childRecsSz; i++){
          var tempChildArray = [];
          for( var ccol in childRecs[i]){
            for( var j=0; j < chdrs.length; j++){
              if(ccol == chdrs[j]){
                tempChildArray.push( childRecs[i][ccol] );
                found = true;
                break;
              }
              if(found) continue;
            }
          }
          crows.push( tempChildArray );
        }

      }
    } //if child

  }


  var data =  { "parent": prows, "child":crows };
  return data;

}

function _getSheetRowsWithChildren(queryResultData, qyObj){
  var dataRows = [];

  var tempRows = [];
  var childRows = [];
  var found = false;
  var hdrs = qyObj.qcolumns;
  var chdrs = qyObj.qchildColumns;
  var childObjName = qyObj.qchildObj;
  var parentObj = queryResultData.records;
  var pObjct = queryResultData.totalSize;


  for( var p=0; p < pObjct; p++){
    var rows = [];
    var parentRows = [];
    var child = parentObj[p][childObjName];


    //push parent obj values
    var parentTempArray = [];
    for( var col in parentObj[p]){
      //search for the current key in the headers and push value to temparr
      for( var i=0; i < hdrs.length; i++){
        if(col == hdrs[i]){
          parentTempArray.push( parentObj[p][col] );
          found = true;
          break;
        }
        if(found) continue;
      }
    }

    for(var x in parentTempArray){
      if(parentTempArray[x] == null){
        parentTempArray[x] = " ";
      }
    }


    if(child){

      //get childRows
      var childRecs = child.records;
      var childRecsSz = childRecs.length;

      if(childRecsSz > 0){
        var tempChildRows = [];

        for(var i=0; i < childRecsSz; i++){
          var tempChildArray = [];
          for( var ccol in childRecs[i]){
            for( var j=0; j < chdrs.length; j++){
              if(ccol == chdrs[j]){
                tempChildArray.push( childRecs[i][ccol] );
                found = true;
                break;
              }
              if(found) continue;
            }
          }


          var temprow = null;
          if( qyObj.qMode == 'Multiple' || qyObj.qchildColumns.length > 0 ){
            temprow =parentTempArray.concat(["''"], tempChildArray);
          } else {
            temprow =parentTempArray.concat(tempChildArray);
          }

          rows.push(temprow);

        }
        dataRows.push(rows);
      }

    }  else { // no children  just print parent
      dataRows.push( parentTempArray );
    }
  }

  return dataRows;
}




//method to find the spacer Column in the header row
//returns Range of spacer Column in header row
function findBlankColumnHeader(sheet){
  var lastCol = sheet.getLastColumn();

  var blankCol=sheet.getRange(1, lastCol+1);

  for(var i = 0; i < lastCol; i++){
    var value = sheet.getRange(1, i+1).getValue();

    if( value == "**" ){
      blankCol = sheet.getRange(1, i+1);
      break;
    }
  }
  return blankCol;

}



//method to find the Parent and Child ranges to be able to get those data rows separately to push to SF
function getParentChildRange(sheet){
  //FIND blank header row which determine the break between the parent range and child range
  var newRange;
  var blankCol = findBlankColumnHeader(sheet);

  var pCol = blankCol.offset(0,-1);  //column before blank ie last column in parent range
  var cCol = blankCol.offset(1,1);   //column after blank  ie  first column in child range and row 2

  //parent A2 to this row new range
  var parentEndRange = sheet.getRange(sheet.getLastRow(), pCol.getColumn());
  var parentRowsRange = sheet.getRange("A2:" + parentEndRange.getA1Notation());

  var childStartRange = cCol.getA1Notation();
  var childEndRange = sheet.getRange(sheet.getLastRow(),sheet.getLastColumn());
  var childRowsRange = sheet.getRange(childStartRange + ":" + childEndRange.getA1Notation());

  var myRanges = { "child" : childRowsRange,
    "parent" : parentRowsRange
  };

  return myRanges;

}


/* method called when Preview Chunks menu option is selected
 *  previewChunks populates a Chunker obj from latest chunker data
 *  and shows rows based on latest chunker configs
 */

function previewChunks(value){

  GSheet.toast("Previewing Chunks....");
  Queries =  _populateSidebarQueue();

  var chunkQueries = [];
  //find Chunker Queries
  for(var i=0; i < Queries.length; i++){

    var cSz = Queries[i].chunkSize;
    var cCol = Queries[i].chunkCol;
    var totalColName = Queries[i].templateTotalColumn;
    var lbl1 = Queries[i].templateLabel1;
    var formula1 = Queries[i].templateFormula1;
    var lbl2 = Queries[i].templateLabel2;
    var formula2 = Queries[i].templateFormula2;


    if( cSz != 'null' || cCol != 'null' ){
      if(Queries[i].status == 'Yes' ){
        var qObj = new Query(Queries[i].query, Queries[i].sheetName, Queries[i].status, Queries[i].qMode);
        var parent =  ( Queries[i].parentExtField != "null") ? Queries[i].parentExtField : null;
        var child =  ( Queries[i].childExtField != "null") ? Queries[i].childExtField : null;

        if(parent != null){
          qObj.qParentObj = parent.split('.')[0];
          qObj.qExtParentObj = parent.split('.')[1];
        }
        if(child != null){
          qObj.qChildObj = child.split('.')[0];
          qObj.qExtChildObj = child.split('.')[1];
        }

        qObj.chunkSize = cSz;
        qObj.chunkCols = cCol;

        if(totalColName != 'null' && lbl1 != 'null' && formula1 != 'null'){
          qObj.templateTotalColumn =totalColName;
          qObj.templateLabel1 = lbl1;
          qObj.templateFormula1 = formula1;
          qObj.templateLabel2 = lbl2;
          qObj.templateFormula2 = formula2;
        }

        qObj.setParams(Queries[i].query);
        chunkQueries.push( qObj );
      }
    }
  }// end for for each query

  if( chunkQueries.length < 1 )
    GSheet.toast("Error: No Chunking Queries to Preview", "No Chunker");


  //loop through chunkqueries to display chunks
  for(var j=0; j < chunkQueries.length; j++){

    GSheet.toast("Chunking " + chunkQueries[j]['resultSheetName'], "CHUNKER");

    var chunkerSheet = "PreviewChunks_" + chunkQueries[j]['resultSheetName'];
    var chunkerParent = chunkQueries[j]['qparentObj'];
    var chunkerChild = chunkQueries[j]['qchildObj'];
    var chunkSize =  (chunkQueries[j]['chunkSize'] == 'null') ? 0 : parseInt( chunkQueries[j]['chunkSize'] );
    var chunkCols = chunkQueries[j]['chunkCols'];
    var chunkParents = chunkQueries[j]['qparentColumns'];
    var chunkChildren = chunkQueries[j]['qchildColumns'];
    var templateTotalCol = chunkQueries[j]['templateTotalColumn'];
    var templateLb1 = chunkQueries[j]['templateLabel1'];
    var templateFormula1 = chunkQueries[j]['templateFormula1'];
    var templateLb2 = chunkQueries[j]['templateLabel2'];
    var templateFormula2 = chunkQueries[j]['templateFormula2'];

    var chunkColsArr = [];
    var chunkby;
    var chunkbyValue = [];

    if( chunkCols != 'null' && chunkSize > 0 ){
      chunkby = 'BOTH';
      chunkColsArr = chunkCols.split(",");
      chunkbyValue.push( chunkSize );
      chunkbyValue.push( chunkColsArr );
    } else if ( chunkCols == 'null'  && chunkSize > 0 ){
      chunkby = 'NUM';
      chunkbyValue.push( chunkSize );
    } else if ( chunkCols != 'null' && chunkSize < 1 ){
      chunkby = 'COL';
      chunkColsArr = chunkCols.split(",");
      chunkbyValue.push( chunkColsArr );
    } else {
      chunkby = null;
    }

    // Updated below to include Chunker Output Sheet 
    var myChunkerObj = new Chunker(
        chunkerSheet, chunkerParent, chunkerChild, chunkby, chunkbyValue.toString(), chunkParents, chunkChildren, templateTotalCol, templateLb1, templateFormula1, templateLb2, templateFormula2);
    
    //calculations the number of chunks from chunk Sheet data

    // First Prepare the Chunk Sheet i.e. Keeping the Original Sheet Intact , copy it to another Sheet and use that for Chunking 
    //if chunksheet already exists dont' re-prepare it
   if(! GSheet.getSheetByName("PreviewChunks_" + chunkQueries[j]['resultSheetName']) ){
    
     
      myChunkerObj.prepareChunkSheets(chunkQueries[j]['resultSheetName']);
      Utilities.sleep(1000);
      myChunkerObj.getChunkInformation();

      if(value === true){
        if( myChunkerObj.childHeaders.length > 0 )
          pushMultipleData(chunkQueries[j]);
        else
          pushTabularData(chunkQueries[j]);

        value = false;
      }
      myChunkerObj.displayData();
    }

    if(value === true){
      if( myChunkerObj.childHeaders.length > 0 )
        pushMultipleData(chunkQueries[j], 'chunkerOnly');
      else
        pushTabularData(chunkQueries[j], 'chunkerOnly');
      value = false;
    }


  }//end for  chunkqueries

  if( chunkQueries.length > 0 ){
    if( ! GSheet.getSheetByName(chunkQueries[0]['resultSheetName']) ){
      GSheet.toast("You must Pull data before you can Chunk!","CHUNKER");
      return;
    }


    GSheet.toast("Done Chunking!", "CHUNKER");
  }

}




//LIBRARY FROM GOOGLE APP SCRIPTS EXAMPLES
// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {

  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getLastColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), headers);
}


function getObjects(data, keys) {

  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j] || "";

      object[keys[j]] = cellData ;

      hasData = true;
    }

    if (hasData &&  object.Id != '*') {
      var obj = isObjectEmpty(object);

      if( obj !== null )
        objects.push(obj);
    }
  }
  return objects;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && ( cellData == "" || cellData == "*" );
}

function isDate(date) {
  return (new Date(date) !== "Invalid Date") ? true : false;
}

/* function to determine if the object is an empty object for deletion or just missing some values
 Returns true if object is a candidate for delete */
function isObjectEmpty(obj){
  var cnt = 0;
  for( var o in obj ){
    if( o != "Id"  ){
      var objLen = typeof( obj[o] !== 'undefined' ) ? obj[o].length : 0;
      cnt += cnt + objLen;
    }
  }

  var lengthO = (typeof(obj['Id']) !== 'undefined') ?  obj['Id'].length : 0 ;

  if( lengthO == 0  && cnt == 0 ){ //blank row  INSERT
    return null;
  } else if( lengthO == 0 && cnt > 0 ){
    delete obj['Id'];
    return obj;
  }


  if( cnt == 0 ){
    for( var o in obj){
      if( o != "Id"){
        delete obj[o]; //DELETE  all objs but ID for Delete push
      }
    }
  } else {
    //update
    for( var o in obj ){
      if( o != "Id"  && obj[o] == ""){
        obj[o] = " ";   //need at least 1 char to go up to salesforce   handling blank cells
      } else if( o == 'Id' && obj[o] == "" ){
        delete obj[o];
      }
    }
  }

  return obj;
}
//*******************************




function showReportList(cache){

  if( userProperties.getProperty("instance_url") == null){
    GSheet.toast("You must be logged in to view Reports!", "REPORT");
    throw new Error("Must Be Logged in " );
  } else {
    sfReports = refreshReportsList(cache);
    var numSR =  getScheduledReportsCount();
    return {data: sfReports, count: numSR };
  }
  return null;
}



// call back function to process selected report
function reportRunClick(dpName, whenRun){
  var myReportObj;
  var newSheet = GSheet.getSheetByName("ReportsList");
  newSheet.getRange(2,1,newSheet.getMaxRows()-1,newSheet.getMaxColumns()).setFontSize("8");
  var myobjs = dpName.split("_");
  var rId = myobjs[0];
  var rtype = myobjs[1];
  var rname = myobjs[2];

  //must know which report to create by type
  if ( rtype == "MATRIX") {
    myReportObj = new mtxObj.MatrixReport(rId);
  } else if ( rtype == "SUMMARY"){
    myReportObj = new srtObj.SummaryReport(rId);
  } else {
    myReportObj = new rptObj.ReportObj(rId);
  }


  var reportMsg = null;

  if(whenRun == 'later'){
    GSheet.toast("Scheduling " + rname + " Report", "REPORTS", 000);

    var myASYNReportObj = new rptObj.ReportObj(rId);
    myASYNReportObj.runReportAsyncInstances(rId);

    scheduledReport =  refreshReportsList(false); //refresh  call server to get update
    return scheduledReport;

  } else {

    GSheet.toast("Creating " + rname + " Report", "REPORTS", 4000);
    var msg =  myReportObj.printReport(rname);
    GSheet.toast(msg.message, "REPORTS");
  }

}

function updateScheduleReport(run, id){
  var newsheet = GSheet.getSheetByName("ReportsList");
  newsheet.getRange(2,1,newsheet.getMaxRows()-1,newsheet.getMaxColumns()).setFontSize("8");
  var reportslist = getReportRows();
  for( var r in reportslist ){

    if( reportslist[r].Id == id ){
      var myrow = parseInt(r) + 1;
      newsheet.getRange(myrow, 6).setValue(run);  //RUNSTATE column
    }
  }

}


//refreshes list of reports ASYNC status
function refreshReportsList(cache){
  var newsheet;
  if(! GSheet.getSheetByName("ReportsList")){
    getReportsFromSF();
  }
  newsheet = GSheet.getSheetByName("ReportsList");

  if(newsheet.getRange("A2").getValue() != ""){
    newsheet.getRange(2,1,newsheet.getMaxRows()-1,newsheet.getMaxColumns()).setFontSize("8");
    var reportslist = getReportRows();
    var myReportObj = new rptObj.ReportObj();

    if( !cache ){
      GSheet.toast("Loading Reports","REPORTS");
      //for each report  call sf server get update
      for( var r in reportslist ){
        if( reportslist[r].Name != "Name" && reportslist[r].Id != "Id" && reportslist[r].Type != "Type" ){
          //get async instances, update status in col E
          var reportData = myReportObj.getReportAsyncInstances(reportslist[r].Id);
          var pReportData = null;
          try{
            pReportData = JSON.parse(reportData);
          } catch(e){
            //null response
          }
          if( pReportData.length > 0 ){
            var myrow = parseInt(r) + 1;
            newsheet.getRange(myrow, 4).setValue(pReportData[0].completionDate);
            newsheet.getRange(myrow, 5).setValue(pReportData[0].status);
            newsheet.getRange(myrow, 6).setValue("");
            newsheet.getRange(myrow, 7).setValue(pReportData[0].id);

          }

        }
      }
    } //end cache


    //populate schedulereports array to treturn
    var scheduledReport = [];

    if(!cache){
      reportslist = getReportRows();   //refresh reports list in case of cache
    }

    for( var r in reportslist ){
      //find row that currently selected report is on and mark as async  Col D
      if( reportslist[r].Name != "Name" && reportslist[r].Id != "Id" && reportslist[r].Type != "Type" ){

        var myrow = parseInt(r) + 1;

        //get values from spreadsheet push to schedule Report array
        var rid = newsheet.getRange(myrow,2).getValue();
        var rname = newsheet.getRange(myrow,1).getValue();
        var rtype = newsheet.getRange(myrow, 3).getValue();
        var rstatus = newsheet.getRange(myrow,5).getValue();
        var rlastrun = newsheet.getRange(myrow,4).getValue();
        var instanceId = newsheet.getRange(myrow,7).getValue();
        var runState = newsheet.getRange(myrow,6).getValue();


        SReport = { 'id': rid, 'name': rname, 'type': rtype, 'status':rstatus, 'lastRun': rlastrun, 'instanceid': instanceId, 'runState': runState};


        scheduledReport.push( SReport );
      }
    }//end for

  }
  return scheduledReport;
}


function runScheduledReports(){
  //timestamp current report on ReportList Sheet

  if(GSheet.getSheetByName("ReportsList").getRange("A2").getValue() != ""){
    var reportslist = getReportRows();

    //for each report
    for( var r in reportslist ){
      if( reportslist[r].Name != "Name" && reportslist[r].Id != "Id" && reportslist[r].Type != "Type" ){
        //get async instances, update status in col E
        if( reportslist[r].RunState == "RUN"  && reportslist[r].Instance != ""){
          var myReportObj = new rptObj.ReportObj(reportslist[r].Id);
          myReportObj.printReport();
        }
      }
    }
  }
}


function getScheduledReportsCount(){
  var numReports = 0;


  if( GSheet.getSheetByName("ReportsList") ){

    if(GSheet.getSheetByName("ReportsList").getRange("A2").getValue() != ""){
      var reportslist = getReportRows();

      //for each report
      for( var r in reportslist ){
        if( reportslist[r].Name != "Name" && reportslist[r].Id != "Id" && reportslist[r].Type != "Type" ){
          if( reportslist[r]['Last Scheduled Run'] ){
            var ldate = reportslist[r]['Last Scheduled Run'];

            if( ldate.length > 1 ){
              numReports = numReports + 1;
            }
          }
        }
      } //end for

    }

  }
  return numReports;
}//end function


function getReportsFromSF(){
  GSheet.toast("Getting Reports from SF");
  var newsheet = null;

  if( GSheet.getSheetByName("ReportsList") ){
    newsheet = GSheet.getSheetByName("ReportsList");
  } else {
    newsheet = GSheet.insertSheet("ReportsList");
  }
   
  
  RM = new ReportManager();
  RM.loadReportManager();
  var reportData = RM.getAllItems();

  var myreportRow = [];
  myreportRow[0] = "Name";
  myreportRow[1] = "Id";
  myreportRow[2] = "Type";
  myreportRow[3] = "Last Scheduled Run";
  myreportRow[4] = "Status";
  myreportRow[5] = "RunState";
  myreportRow[6] = "Instance";
  newsheet.appendRow(myreportRow);

  for(var r in reportData){
    myreportRow[0] = reportData[r].reportName;
    myreportRow[1] = reportData[r].reportID;
    myreportRow[2] = reportData[r].reportType;
    myreportRow[3] = " ";
    myreportRow[4] = " ";
    myreportRow[5] = " ";
    myreportRow[6] = " ";
    newsheet.appendRow(myreportRow);
  }
  //newSheet header and data styles 
  newsheet.getRange(1, 1, 1, 7).setFontSize("9").setBackground("#0C8EFF").setFontColor("#ffffff");
  newsheet.getRange(2,1,newsheet.getMaxRows()-1,newsheet.getMaxColumns()).setFontSize("8");

  return reportData.length();
}



function getReportRows(){
  var reportslist = null;

  if(GSheet.getSheetByName("ReportsList").getRange("A2").getValue() != ""){
    var  newSheet = GSheet.setActiveSheet(GSheet.getSheetByName("ReportsList"));
    var myendRange = newSheet.getRange(GSheet.getLastRow(), GSheet.getLastColumn());
     //newSheet.getRange(2,1,newSheet.getMaxRows()-1,newSheet.getMaxColumns()).setFontSize("8");
    var myrange = newSheet.getRange("A1:" + myendRange.getA1Notation());
    reportslist = getRowsData(newSheet, myrange,1);
  }
  return reportslist;
}

function clearReportRows(){
  var newSheet = GSheet.setActiveSheet(GSheet.getSheetByName("ReportsList"));
  newSheet.clear();  //clear sheet before new results

}

/* uncomment if SF org has refresh token permissions */
function validateSFsession(){
  var flagsession = null;
  var  SF = new SForce(CLIENT_ID, CLIENT_SECRETID, SF_VERSION, REDIRECT_URI);



  
  // Getting Session Tokens from User Properties.
  //Have to get the tokens from User Level.
  var tokenresponseInstance_url = userProperties.getProperty("instance_url");
  var tokenresponseAccess_token = userProperties.getProperty("sessionId");


  if( tokenresponseInstance_url !== null && tokenresponseAccess_token !== null){


    flagsession = true;
  } else {
    flagsession = false;
  }

  return flagsession;

}