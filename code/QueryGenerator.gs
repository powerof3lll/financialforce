/**
 * To get all the sObjects
 */
function getObjects() {  
  var data = null;
  SF = new SForce(CLIENT_ID, CLIENT_SECRETID, SF_VERSION, REDIRECT_URI);
  
  try{
    data = SF.getSObjects();
    if( data.errorCode ){
      GSheet.toast( data.errorCode );
    }
  } catch(e){
    if( e.errorCode ){
      GSheet.toast( JSON.stringify(e.errorCode) );
      // Call Error Handler 
      ErrorHandler.LogError(JSON.stringify(e.errorCode));
    }
    return false;
  }
  return data;
}

/**
 * To get all the fields of the sObject
 */
function getSObjectsFields(sObj) {
  var data = null;
  SF = new SForce(CLIENT_ID, CLIENT_SECRETID, SF_VERSION, REDIRECT_URI);
  
  try{
    data = SF.getSObjectsFields(sObj);
    if( data.errorCode ){
      GSheet.toast( data.errorCode );
    }
  } catch(e){
    if( e.errorCode ){
      GSheet.toast( JSON.stringify(e.errorCode) );
      // Call Error Handler 
      ErrorHandler.LogError(JSON.stringify(e.errorCode));
    }
    return false;
  }
  return data;
}

/**
 * To get all the child-relationships of Parent sObject
 */
function getSObjectsChildRelationships(sObj) {
  
  var data = null;
  SF = new SForce(CLIENT_ID, CLIENT_SECRETID, SF_VERSION, REDIRECT_URI);
  
  try{
    data = SF.getSObjectsChildRelationships(sObj);
    if( data.errorCode ){
      GSheet.toast( data.errorCode );
    }
  } catch(e){
    if( e.errorCode ){
      GSheet.toast( JSON.stringify(e.errorCode) );
      // Call Error Handler 
      ErrorHandler.LogError(JSON.stringify(e.errorCode));
    }
    return false;
  }
  return data;
}

/**
 * To get all the saved queries from the sheet
 */
function getSavedQueryListFromSheet(){
  var queries = [];
  try{
    var querySheet = SpreadsheetApp.getActive().getSheetByName("QueryList");
    var lastRow = querySheet.getLastRow();

    for( var i =1; i <= lastRow; i++){
      var cellValue = querySheet.getRange(i,2).getValue();
      if( cellValue != "" && cellValue != "Query"){
        queries.push( cellValue );
      }
    }
  }catch(e){
    if( e.errorCode ){
      GSheet.toast( JSON.stringify(e.errorCode) );
      // Call Error Handler 
      ErrorHandler.LogError(JSON.stringify(e.errorCode));
    }
  }
  return queries;
}

/**
 * Deletes a saved query and refreshes the view
 */
function deleteQueryFromSavedQueryList(index){
  try{
    showRemoveQuery(index);
    return true;
  }catch(e){
    if( e.errorCode ){
      GSheet.toast( JSON.stringify(e.errorCode) );
      // Call Error Handler 
      ErrorHandler.LogError(JSON.stringify(e.errorCode));
    }
    return false;
  }
}

/**
 * Run a query
 */
function querySF(query){
  var data = null;
  SF = new SForce(CLIENT_ID, CLIENT_SECRETID, SF_VERSION, REDIRECT_URI);
  
  try{
    data = SF.querySF(query);
    if( data.errorCode ){
      GSheet.toast( data.errorCode );
    }
  } catch(e){
    if( e.errorCode ){
      GSheet.toast( JSON.stringify(e.errorCode) );
      // Call Error Handler 
      ErrorHandler.LogError(JSON.stringify(e.errorCode));
    }
    return false;
  }
  return data;
}