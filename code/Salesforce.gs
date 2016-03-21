//******************************SALESFORCE OBJECT***************************************************
//create a Salesforce Object that handles all SF functions including Authentication, Push, Pull

//Storing Properties at User Level
var userProperties = PropertiesService.getUserProperties();

var SF_AUTHORIZE_URL = "https://login.salesforce.com/services/oauth2/authorize";
var SF_TOKEN_URL = "https://login.salesforce.com/services/oauth2/token";
var SF_REVOKE_URL = "https://login.salesforce.com/services/oauth2/revoke";
var GSheet = SpreadsheetApp.getActiveSpreadsheet();
var ErrorHandler = new ErrorLog();

function SForce(cid, csid, version, rUrl){

    this.clientID = cid;
    this.clientSecretID = csid;
    this.clientScope = "full";
    this.authorizeURL = null;
    this.tokenURL = null;
    this.sessionId= null;
    this.SFauthURL = SF_AUTHORIZE_URL;
    this.SFtokenURL = SF_TOKEN_URL;
    this.SFinstanceURL = null;
    this.SFversion = version;
    this.redirectUrl = rUrl;
    this.refreshToken = null;



    //Property getters and setters
    Object.defineProperty(this, "sfcid", {
        get: function() { return this.clientID; },
        set: function(v) { this.clientID = v; userProperties.setProperty("clientId", v); }
    });

    Object.defineProperty(this, "sfscope", {
        get: function() { return this.clientScope; },
        set: function(v) { this.clientScope = v; userProperties.setProperty("clientScope", v); }
    });

    Object.defineProperty(this, "sfcsid", {
        get: function() { return this.clientSecretID; },
        set: function(v) { this.clientSecretID = v; userProperties.setProperty("clientSecret", v); }
    });

    Object.defineProperty(this, "sfsessid", {
        get: function() { return this.sessionID; },
        set: function(v) { this.sessionID = v; userProperties.setProperty("sessionId", v);  }
    });
  
   Object.defineProperty(this, "sfrefreshtoken", {
        get: function() { return this.refreshToken; },
        set: function(v) { this.refreshToken = v; userProperties.setProperty("refreshToken", v);  }
    });


    Object.defineProperty(this, "sfinstanceUrl", {
        get: function() { return this.SFinstanceURL; },
        set: function(v) { this.SFinstanceURL = v; userProperties.setProperty("instance_url", v); }
    });

    Object.defineProperty(this, "sfversion", {
        get: function() { return this.SFversion; },
        set: function(v) { this.SFversion = v; userProperties.setProperty("sfversion", v); }

    });

    Object.defineProperty(this, "authurl", {
        get: function() { return (this.authorizeURL == null) ? SF_AUTHORIZE_URL : userProperties.getProperty("authorizeUrl"); },
        set: function(v) { this.authorizeURL = v; userProperties.setProperty("authorizeUrl", v); }
    });

    Object.defineProperty(this, "tokenurl", {
        get: function() { return  (this.tokenURL ==null )? SF_TOKEN_URL : this.tokenURL; },
        set: function(v) { this.tokenURL = v; userProperties.setProperty("tokenUrl", v); }
    });

    Object.defineProperty(this, "sfAuthURL", {
        get: function(){
            this.SFauthURL = this.authurl + "?response_type=code&client_id=" + this.sfcid + "&scope="+this.sfscope+"&redirect_uri=" + REDIRECT_URI;
           userProperties.setProperty("sfAuthUrl", this.SFauthURL);
            return this.SFauthURL;
        }
    });

    Object.defineProperty(this, "sfTokenURL", {
        get: function(){
            this.SFtokenURL = this.tokenurl + "?client_id=" + this.sfcid + "&client_secret=" + this.sfcsid + "&grant_type=authorization_code&redirect_uri=" + REDIRECT_URI;
           userProperties.setProperty("sfTokenUrl", this.SFtokenURL);
            return this.SFtokenURL;
        }
    });


//Methods

   /*
    * Method :  getSFInstanceURL
    * Description : Returns the SF Instance URL
    */
  
   this.getSFInstanceURL = function()
   {
     var inst = this.SFinstanceURL ||  userProperties.getProperty("instance_url");
     return inst;
   }
   
   /*
    * Method :  getAccessToken
    * Description : Returns the SF Access Token
    */
  
   this.getAccessToken = function()
   {
     var accToken = this.sessionID || userProperties.getProperty("sessionId");
     return accToken;
   }
   
   /*
    * Method :  formatDate
    * Description : Convert the date input into valid SFDC Date Format
    * @params : data,can be Date or DateTime
    * @params : type,indicates whether to convert to Date or DateTime format
    */
  
   this.formatDateTime = function(data,type)
   {   
      data = data.trim();
      var res = data;  
      var d;       
     if(data.length > 0)
     {
     if(type == 'date')
     {
       d = new Date(data);       
       var gmtDate = new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate(),d.getHours(),d.getMinutes(),d.getSeconds(),d.getMilliseconds()));
       res = Utilities.formatDate(gmtDate,"GMT","yyyy-MM-dd");       
     }
     
     if(type == 'datetime' && data.indexOf("T")!=-1)
      {
         var str = data;
         var tokens = str.split("T");
         var datePart = tokens[0].split("-");  // Get YYYY-MM-DD
         var second = tokens[1].split("+");
         var timePart = second[0].split(":");
         var timeZone = Session.getScriptTimeZone();         
         // As Javascript start month numbering from 0,decrementing by 1 to get the correct month.         
         d = new Date(Date.UTC(parseInt(datePart[0]),parseInt(datePart[1])-1,parseInt(datePart[2]),parseInt(timePart[0]),parseInt(timePart[1]),parseInt(timePart[2]),0));                    
         var res = Utilities.formatDate(d,"GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
      }
     }     
      return res;
   }
   
  
    //connects to salesforce for Authentication
    this.connectSFAuth = function(code){
              
        var codeURL = this.sfTokenURL + "&code=" + code;

        try{
            var response = UrlFetchApp.fetch(codeURL).getContentText();
        } catch(e){
            Logger.log("ERROR in connectSFAuth" + e.message);
            return false;
        }

        var tokenresponse = JSON.parse(response);
        
       userProperties.setProperty("sf_id", tokenresponse.Id);
        this.SFinstanceURL = tokenresponse.instance_url;
       userProperties.setProperty("instance_url", tokenresponse.instance_url);
        this.sessionID = tokenresponse.access_token;
       userProperties.setProperty("sessionId", tokenresponse.access_token);
        
      
        //Store the refresh token Details at User Level 
       this.refreshToken = tokenresponse.refresh_token;
       userProperties.setProperty("refreshToken",tokenresponse.refresh_token);
      
      
        return true;
      
    };

    //disconnects from salesforce for Authentication
    this.disconnectSFAuth = function(){

        var sid = this.getAccessToken();
        var revokeURL = SF_REVOKE_URL + "?token=" + sid;

        try{
            var response = UrlFetchApp.fetch(revokeURL).getContentText();


        } catch(e){
            Logger.log("ERROR in disconnectSFAuth" + e.message);
            // Call Error Handler
            ErrorHandler.LogError("ERROR in disconnectSFAuth" + e.message);
            //return false;
        }
        var inst = this.getSFInstanceURL();

        this.tokenURL = null;
        this.sessionId= null;
        this.SFinstanceURL = null;
        userProperties.deleteAllProperties();
        
        //Delete all Document Level Properties
        userProperties.deleteAllProperties();
        //var tokenresponse = JSON.parse(response);

        return true;
    };


    this.refreshTokenAuth = function(){
        //The consumer should make POST request to the token endpoint
        /*
         */

        //ErrorHandler.LogError('Inside refreshTokenAuth Method');
        //Logger.log('Inside refreshTokenAuth Method');
        var sid = userProperties.getProperty('refreshToken');
        var refreshURL = this.tokenurl;
        
        var myPayloader = "grant_type=refresh_token&client_id="+this.clientID+"&client_secret="+this.clientSecretID+"&refresh_token="+sid;
        var options = {
            headers : { "Authorization" : "Bearer " + sid,
                "Content-Type" : "application/x-www-form-urlencoded"
            },
            payload : myPayloader
        };

        Logger.log(myPayloader);

        try{
            var response = UrlFetchApp.fetch(refreshURL, options).getContentText();
        } catch(e){
            Logger.log("ERROR in refreshSFAuth" + e.message);
            // Call Error Handler
            ErrorHandler.LogError("ERROR in refreshSFAuth" + e.message);
            return false;
        }

        var refreshtokenresponse = JSON.parse(response);
        this.sessionID = refreshtokenresponse.access_token;
       userProperties.setProperty("sessionId", refreshtokenresponse.access_token);
       
        return refreshtokenresponse;

    };
  
    /* 
     * Utility to check if Access Token is expired or not . 
     * If its invalid , then Refresh Token is used to Refresh the Session
     * This will be used in all places where SFDC Services are invoked.
     */
  
    this.refreshSession = function()
    {      
      var inst = this.getSFInstanceURL();
      var accesstoken = this.getAccessToken();
      var url = inst + "/services/data/v" + this.sfversion;
      var options = {
            headers : { "Authorization" : "Bearer " + accesstoken },
            muteHttpExceptions : true
        };
      var response = this.sfAPIcall(url, options);      

      /* Check the Response Code. As per SFDC , for Invalid Session its 401
       * If its 401, then invoke refreshTokenAuth to get the new Session Token.
       */ 
      if(response.getResponseCode() == 401)
       this.refreshTokenAuth();
    }

    //utility method to process calls to Salesforce and returns raw response
    this.sfAPIcall = function(url, options){
        var response = null; //response from fetch
        try{
            response = UrlFetchApp.fetch(url, options);
        } catch(e){
            Logger.log(e);
            // Call Error Handler
            ErrorHandler.LogError(e);
            return e.errorCode;
        }
        return response;
    };


    //queries Salesforce data with passed in Query
    this.querySF = function(query){
      
        //Get Valid Session Token if possible
        this.refreshSession();
      
        var inst = this.getSFInstanceURL();
        var accesstoken = this.getAccessToken();
        var url = inst + "/services/data/v" + this.sfversion + "/query?q=" + encodeURIComponent(query);
        var options = {
            headers : { "Authorization" : "Bearer " + accesstoken }
            //muteHttpExceptions : true
        };

        var response = this.sfAPIcall(url, options);
        var dataresponse = null;

        if( response.error ){
            Logger.log("Reponse Error : " + response.error);
            ErrorHandler.LogError("Reponse Error : " + response.error);
            return dataresponse = { error: response.error };

        } else {
            try{
                dataresponse = JSON.parse(response);
            } catch(e){
                return {errorCode: "Session Expired!"};
                ErrorHandler.LogError("Session Expired!");
            }
        }
        return dataresponse;
    };


    //Method to get SOBJECTS from Salesforce
    this.getSObjects = function(){

        
        var inst = this.getSFInstanceURL();
        var accesstoken = this.getAccessToken();
        var url = inst + "/services/data/v" + this.SFversion + "/sobjects/";
        var options = {
            headers : { "Authorization" : "Bearer " + accesstoken }
            //muteHttpExceptions : true
        };

        var dataresponse = null;
        var response = this.sfAPIcall(url, options);
        try{
            dataresponse = JSON.parse(response);
        } catch(e){
            var err = { error: e, message: 'Salesforce API call SObjects failed!', title: 'ERROR' };
            return err;
            // Call Error Handler 
            ErrorHandler.LogError('Salesforce API call SObjects failed! ' + e);
        }
        return dataresponse.sobjects;

    };



    //Method to get SOBJECTS FIELDS from Salesforce
    this.getSObjectsFields = function(sobj){

        var inst = this.getSFInstanceURL();
        var accesstoken = this.getAccessToken();
        var url = inst + "/services/data/v" + this.sfversion + "/sobjects/" + sobj + "/describe/";
        var options = {
            headers : { "Authorization" : "Bearer " + accesstoken }
            //muteHttpExceptions : true
        };

        var dataresponse = null;
        var response = this.sfAPIcall(url, options);
        try{
            dataresponse = JSON.parse(response);
        } catch(e){
            var err = { error: e, message: 'Salesforce API call SObject Fields failed!', title: 'ERROR' };
            // Call Error Handler 
            ErrorHandler.LogError('Salesforce API call SObjects failed! ' + e);
            return err;
        }

        return dataresponse.fields;

    };




//Pushes data up to Salesforce  (non-Bulk parent only)
    this.sendSF = function(qobject, myobjs){
     
        //Get Valid Session Token if possible
        this.refreshSession();
      
        var inst = this.getSFInstanceURL();
        var accesstoken = this.getAccessToken();
        var url;
        var myObj = qobject.qparentObj;
        var operation;
        var patchCnt = 0;
        var postCnt = 0;
        var deleteCnt = 0;
        var failureCnt = 0;
        var sheetNameofFailureRecord = qobject.resultSheetName;
        var errorMessage = [];
        var errorRow = [];

        if( inst == null ){
            return  { error: "Instance URL is null" } ;
        }
      
        
        // Here Get the Object Fields and store them .
        // 
        var objFields = this.getSObjectsFields(myObj);

        for(var i =0; i < myobjs.length; i++){
            var myTempID = myobjs[i].Id;
            var tempObj = myobjs[i];


            //Check for Obj lengths.  If only ID exists then this is a deleted obj
            var obLength = 0;
            for( var ob in myobjs[i]){
                if(myobjs[i][ob].length > 0 ){
                    obLength +=  myobjs[i][ob].length;
                }
            }

            if(myobjs[i].Id != null && myobjs[i].Id != "") {
                if(obLength > myobjs[i].Id.length){
                    operation = "PATCH";
                    url = inst + "/services/data/v" + this.sfversion + "/sobjects/" + myObj + "/" + myTempID+"?_HttpMethod="+operation;
                } else {
                    operation = "DELETE";
                    url = inst + "/services/data/v" + this.sfversion + "/sobjects/" + myObj + "/" + myTempID+"?_HttpMethod="+operation;
                }

            }else{
                operation = "POST";
                url = inst + "/services/data/v" + this.sfversion + "/sobjects/" + myObj + "/";
            }


            var mytempob = myobjs[i];
          
      
            if(mytempob.Id){
                delete mytempob.Id;    //can not include ID in payload so delete it
            }
                // Based on Type of Operation , remove the ReadOnly and Non-Updateable Fields 
                
                // For Update Operation --- Remove the fields that cannot be updated
                if(operation == 'PATCH')
                {
                  for(var f in objFields)
                  {  
                    
                    //Convert Date/DateTime fields to its proper SFDC Format
                    if((objFields[f].type == 'date' || objFields[f].type == 'datetime') && mytempob[objFields[f].name])
                    {
                      mytempob[objFields[f].name] = this.formatDateTime(mytempob[objFields[f].name],objFields[f].type);
                    }
                    
                    if(objFields[f].updateable == false && mytempob[objFields[f].name])
                    {
                      delete mytempob[objFields[f].name];
                    }
                    
                    
                    
                  }
                }
              
                // For Insert Operation --- Remove the fields that cannot be created
                if(operation == 'POST')
                {
                  for(var f in objFields)
                  {
                    
                    //Convert Date/DateTime fields to its proper SFDC Format
                    if((objFields[f].type == 'date' || objFields[f].type == 'datetime') && mytempob[objFields[f].name])
                    {
                      mytempob[objFields[f].name] = this.formatDateTime(mytempob[objFields[f].name],objFields[f].type);
                    }
                    
                    if(objFields[f].createable == false && mytempob[objFields[f].name])
                      delete mytempob[objFields[f].name];
                  }
                }
                               
            var myPayloader = JSON.stringify(mytempob);



            //request data from SF
            var options = {
                headers : { "Authorization" : "Bearer " + accesstoken,
                    "Content-Type" : "application/json"
                },
                payload : myPayloader

            };

            var response;
            try{
                response = UrlFetchApp.fetch(url, options);

                //this is where to increment counts
                if( operation == "PATCH" ){
                    patchCnt++;
                } else if ( operation == "DELETE"){
                    deleteCnt++;
                } else if ( operation == "POST" ){
                    postCnt++;
                }
            } catch(e){
                failureCnt++;
                errorRow.push(i+2); //account for starting at 0 and the header row on sheet
                errorMessage.push(e.message);
                // Call Error Handler 
                ErrorHandler.LogError(e.message);
            }

        }

        var infoMsg = "Inserted: " + postCnt + " records " + "/\n Deleted: " + deleteCnt + " records " + " /\n Updated: " + patchCnt + " records";

        if(failureCnt > 0){
            Utilities.sleep(3000);
            infoMsg = infoMsg + "FAILED RECORDS: " + failureCnt + " on Sheet " + sheetNameofFailureRecord + " at row(s) " + errorRow.join();

            //show error rows in read
            var gsSheet = SpreadsheetApp.getActiveSpreadsheet();
            var actSheet = gsSheet.getSheetByName(sheetNameofFailureRecord);
            var savedLastColumn = actSheet.getLastColumn();

            for( var x =0; x < errorRow.length; x++){
                var errRang = actSheet.getRange(
                    actSheet.getRange(errorRow[x], 1).getA1Notation() + ":" +
                    actSheet.getRange(errorRow[x], savedLastColumn).getA1Notation());
                errRang.setBackground("#F6CECE");

                var errMsgRang = actSheet.getRange( errRang.getRow(), savedLastColumn + 1 );
                errMsgRang.setValue( errorMessage[x] ).setFontColor("red").setFontSize(9).setFontWeight("bold").setWrap(false);
            }
        }
        return { response: response, message: infoMsg } ;

    };


    //Pushes data up to Salesforce with externalField IDs
    //inputs a Query Obj that has externalfield ids, the type of obj (parent or child), objs to push
    this.sendSFext = function(qobject, objType, myobjs){
      
        //Get Valid Session Token if possible
        this.refreshSession();
      
        var inst = this.getSFInstanceURL();
        var accesstoken = this.getAccessToken();
        var url;
        var myObj;
        var externalField;
        var url;
        var extId;
        var myPayloader;
        var myResults = [];


        //get Object and ExternalFieldID
        if(objType == "Parent"){
            myObj = qobject.qparentObj;
            externalField = qobject.qExtParentObj;
            if(typeof(externalField) == "undefined"){
            }
        } else {
            //child
            //condition child object to singular
            tempchildObj = qobject.qchildObj;
            //SALESFORCE RELATIONSHIP RULES  if in plural form, change to singular for push
            if( tempchildObj.charAt(tempchildObj.length-1) == 's'){
                myObj = tempchildObj.slice(0,tempchildObj.length-1);
            }else {
                myObj = tempchildObj;
            }
            externalField = qobject.qExtChildObj;
        }

        //Prepare and Send Request
        for(var i=0; i < myobjs.length; i++){
            extId = myobjs[i][externalField].trim().replace(" ", "%20");  //externalField value

            url = inst + "/services/data/v" + this.sfversion + "/sobjects/" + myObj + "/" + externalField + "/" + extId + "?_HttpMethod=PATCH";

            delete myobjs[i][externalField];    //can not include externalField in payload so delete it
            myPayloader = JSON.stringify(myobjs[i]);


            //request data from SF
            var options = {
                headers : { "Authorization" : "Bearer " + accesstoken,
                    "Content-Type" : "application/json",
                    "X" : "PATCH"
                },
                payload : myPayloader
            };

            try{
                var response = UrlFetchApp.fetch(url, options);
                if( response.getContentText() != ""){
                    myResults.push( response.getContent() );
                }
            } catch(e){
                Logger.log( e.message);
                // Call Error Handler 
                ErrorHandler.LogError(e.message);
            }

        }//end for

        var infoMsg = "Processed " + myResults.length + " records using External Field ID";

        return {response: response, infoMsg: infoMsg };
    };












    //////////////////////// Salesforce Bulk API Methods ////////////////////

    //Pushes data up to Salesforce  (Bulk API)
    this.sendBulkSF = function(myobjRows, objType){
      
       //Get Valid Session Token if possible
       this.refreshSession();
      
        //temporary containers for objects until job is created and they are batched appropriately
        var insertObjs = [];
        var updateObjs = [];
        var deleteObjs = [];
      
      // Maps for Removing Duplicate Ids
      var updateIdsMap = new Object();
      var deleteIdsMap = new Object();

      
      // Here Get the Object Fields and store them .
        // 
        var objFields = this.getSObjectsFields(objType);
      
      
        //foreach obj determine operation
        for(var i=0; i < myobjRows.length; i++){
            var obLength = 0;
            for(var ob in myobjRows[i]){
                if(myobjRows[i][ob].toString().length > 0){
                    obLength += myobjRows[i][ob].toString().length;
                }
            }

            if(myobjRows[i].Id != null && myobjRows[i].Id != ""){
                if(obLength > myobjRows[i].Id.length){
                    //then this row has more than just the ID so update

                   // Removing the Duplicate Ids 
                   if(!updateIdsMap[myobjRows[i].Id])
                   { 
                     
                     // Remove the Non-Updateable Fields
                     for(var f in objFields)
                  {  
                    
                    //Convert Date/DateTime fields to its proper SFDC Format
                    if((objFields[f].type == 'date' || objFields[f].type == 'datetime') && myobjRows[i][objFields[f].name])
                    {
                      myobjRows[i][objFields[f].name] = this.formatDateTime(myobjRows[i][objFields[f].name],objFields[f].type);
                    }
                    
                    if(objFields[f].updateable == false && myobjRows[i][objFields[f].name] && objFields[f].name != 'Id')
                    {
                      delete myobjRows[i][objFields[f].name];
                    }
                    
                    
                    
                    
                  }
                     
                     updateObjs.push(myobjRows[i]);
                     updateIdsMap[myobjRows[i].Id] = 1;
                   }
                  
                } else {
                    //if only id exists then this is a delete candidate
                    
                    // Removing the Duplicate Ids
                    
                     
                    if(!deleteIdsMap[myobjRows[i].Id])
                    {
                      deleteObjs.push(myobjRows[i]); 
                      deleteIdsMap[myobjRows[i].Id] = 1;
                    }
                }
            } else {
                // there is no Id on row so insert this row
                //additionally if no other objs  dont' insert  it is a blank row
                var eObj = isObjectEmpty(myobjRows[i]);
                if( eObj != null ){
                  
                    // Remove the Non-Createable Fields
                     for(var f in objFields)
                  {
                    
                    //Convert Date/DateTime fields to its proper SFDC Format
                    if((objFields[f].type == 'date' || objFields[f].type == 'datetime') && myobjRows[i][objFields[f].name])
                    {
                      myobjRows[i][objFields[f].name] = this.formatDateTime(myobjRows[i][objFields[f].name],objFields[f].type);
                    }
                    
                    
                    if(objFields[f].createable == false && myobjRows[i][objFields[f].name])
                      delete myobjRows[i][objFields[f].name];
                  }
                    insertObjs.push(myobjRows[i]);
                }
            }


        } //end for
              

        var jobs = [];  //will hold the job and batch ids returned from processJobBatches func
        var myInserts = {};
        var myUpdates = {};
        var myDeletes = {};
        var eMsg = "Processing " + objType + "....";
      
           
        // Do the Processing here
        // Removal of Readonly and non-Updateable Fields 
      

        //create jobs for operations with rows
        if(insertObjs.length > 0){
            myInserts =  this.processJobBatches(insertObjs, objType, "insert");
            jobs.push(myInserts);
            eMsg += "Inserting " + insertObjs.length + " records ";
        }

        if(updateObjs.length > 0){
            myUpdates =  this.processJobBatches(updateObjs, objType, "update");
            jobs.push(myUpdates);
            eMsg += " Updating " + updateObjs.length + " records ";
        }

        if(deleteObjs.length > 0){
            myDeletes = this.processJobBatches(deleteObjs, objType, "delete");
            jobs.push(myDeletes);
            eMsg += " Deleting " + deleteObjs.length + " records ";
        }

        return eMsg;

    };


    //method to create jobs and process csv data into batches
    //Params:  data  to send in CSV format; object for batch, operation to perform on batch
    this.processJobBatches = function(data, obj, operation){

        var jobID = this.createJob(operation, obj);

        //process data into csv  then pass to add to batch
        var csvData = dataToCSV(data);
        var batchID = this.addBatch(jobID, csvData);

        var job = {"jobID":jobID, "batchID":batchID};


        Utilities.sleep(1000);

        try{
            var res =  this.getBatchResults(jobID, batchID);

            for(var i =0; i < res.length; i++){
                if( res[i][1] === 'false' ){
                    ErrorHandler.LogError(res[i][1] + " " + res[i][3]);
                }
            }

        } catch(e){
            Logger.log( e.message );
            ErrorHandler.LogError(e.message);
        }


        return job;

    };

    //method to create a Salesforce Bulk API job
    //returns the JobID
    this.createJob = function(operation, myobj){

        var inst = this.getSFInstanceURL();
        var accesstoken = this.getAccessToken();
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
        var jobInfo = "<jobInfo xmlns=\"http://www.force.com/2009/06/asyncapi/dataload\">";
        var xmlOperation = "<operation>" + operation + "</operation>";
        var xmlObject = "<object>" + myobj + "</object>";
        var xmlContentType = "<contentType>CSV</contentType></jobInfo>";
        var formattedXML = xml + jobInfo + xmlOperation + xmlObject + xmlContentType;

        //create job
        url = inst + "/services/async/" + this.sfversion + "/job";
        var options = {

            headers : {
                "X-SFDC-Session " : accesstoken,
                "Content-Type " : "application/xml",
                "charset " : "UTF-8"
            },

            payload : formattedXML

        };


        try{
            var response = UrlFetchApp.fetch(url, options);
        } catch(e){
            Logger.log( e.message );
            // Call Error Handler 
            ErrorHandler.LogError(e.message);
        }

        //parse xml response
        var document = XmlService.parse(response);
        var root = document.getRootElement();
        var atom = XmlService.getNamespace("http://www.force.com/2009/06/asyncapi/dataload");
        var jobId = root.getChildText("id", atom);
        var jobState = root.getChildText("state", atom);


        return jobId;

    };



    //method to close a Salesforce Bulk API job
    this.closeJob = function(jobID){

        var inst = this.getSFInstanceURL();
        var accesstoken = this.getAccessToken();
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
        var jobInfo = "<jobInfo xmlns=\"http://www.force.com/2009/06/asyncapi/dataload\">";
        var xmlState = "<state>Closed<state></jobInfo>";
        var formattedXML = xml + jobInfo + xmlState;

        //close a job
        url = inst + "/services/async/" + this.sfversion + "/job/" + jobID;

        var options = {

            headers : {
                "X-SFDC-Session " : accesstoken,
                "Content-Type " : "application/xml",
                "charset " : "UTF-8"
            },

            payload : formattedXML
        };

        try{
            var response = UrlFetchApp.fetch(url, options);
        } catch(e){
            Logger.log( e.message );
            // Call Error Handler 
            ErrorHandler.LogError(e.message);
        }


        //parse xml response
        var document = XmlService.parse(response);
        var root = document.getRootElement();
        var atom = XmlService.getNamespace("http://www.force.com/2009/06/asyncapi/dataload");
        var jobId = root.getChildText("id", atom);
        var jobState = root.getChildText("state", atom);

        return jobState;

    };




    //method to add batch data to a Salesforce Bulk API job
    this.addBatch = function(jobID, batchData){

        //Get Valid Session Token if possible
        this.refreshSession();
      
       var inst = this.getSFInstanceURL();
       var accesstoken = this.getAccessToken();

        //create batch add to job
        url = inst + "/services/async/" + this.sfversion + "/job/" + jobID +"/batch";
        var options = {

            headers : {
                "X-SFDC-Session " : accesstoken,
                "Content-Type " : "text/csv",
                "charset " : "UTF-8"
            },

            payload : batchData
        };

        try{
            var response = UrlFetchApp.fetch(url, options);
        } catch(e){
            Logger.log(e.message);
        }


        //parse xml response
        var document = XmlService.parse(response);
        var root = document.getRootElement();
        var atom = XmlService.getNamespace("http://www.force.com/2009/06/asyncapi/dataload");
        var batchId = root.getChildText("id", atom);
        var batchState = root.getChildText("state", atom);


        return batchId;
    };


    //method to get batch status of Salesforce Bulk API job
    this.getBatchStatus = function(jobID, batchID){
      
        //Get Valid Session Token if possible
        this.refreshSession();
      
      var inst = this.getSFInstanceURL();
      var accesstoken = this.getAccessToken();

        //get batch status
        url = inst + "/services/async/" + this.sfversion + "/job/" + jobID +"/batch/"+batchID;
        var options = {

            headers : {
                "X-SFDC-Session " : accesstoken
            }
        };

        try{
            var response = UrlFetchApp.fetch(url, options);
        } catch(e){
            Logger.log( e.message);
            // Call Error Handler 
            ErrorHandler.LogError(e.message);
        }


        //parse xml response
        var document = XmlService.parse(response);
        var root = document.getRootElement();
        var atom = XmlService.getNamespace("http://www.force.com/2009/06/asyncapi/dataload");
        var batchId = root.getChildText("id", atom);
        var jobId = root.getChildText("jobId", atom);
        var batchState = root.getChildText("state", atom);
        var batchNumProcessed = root.getChildText("numberRecordsProcessed", atom);

        var batchinfo = { "job": jobId, "batch": batchId, "state": batchState, "numProcessed" : batchNumProcessed };
        return batchinfo;

    };


    //method to get batch results of Salesforce Bulk API job
    this.getBatchResults = function(jobID, batchID){
      
        //Get Valid Session Token if possible
        this.refreshSession();
      
      var inst = this.getSFInstanceURL();
      var accesstoken = this.getAccessToken();

        //get batch status
        url = inst + "/services/async/" + this.sfversion + "/job/" + jobID +"/batch/"+batchID+"/result";
        var options = {

            headers : {
                "X-SFDC-Session " : accesstoken
            }
        };

        var response;

        try{
            response = UrlFetchApp.fetch(url, options);
            var res = CSVToArray(response);
            return res;

        } catch(e){
            Logger.log( e.message);
            ErrorHandler.LogError(e.message);
        }

    };





    //private method to convert data to CSV
    var dataToCSV = function(dataToConvert){
        var csvFile;
        var csv = "";
        var tempcsv = "";
        var ct = 0;
        var csvheaders = [];

        if(dataToConvert.length > 0){

            //add header row to csv file
            for(var hr in dataToConvert[0]){
                csvheaders.push(hr);
            }
            tempcsv = csvheaders.join(",");
            tempcsv += "\r\n";
            csv = tempcsv;


            //add data rows to csv file
            for(var i=0; i < dataToConvert.length; i++){
                var csvRow = [];
                tempcsv = "";
                for(var ob in dataToConvert[i]){

                    if (dataToConvert[i][ob].toString().indexOf(",") != -1) {
                        //adding quotes around data
                        dataToConvert[i][ob] = "\"" + dataToConvert[i][ob].trim() + "\"";
                    }
                    csvRow.push(dataToConvert[i][ob]);

                } //end value for

                tempcsv += csvRow.join(",");
                tempcsv += "\r\n";
                csv += tempcsv;
            } //end row for
        }//end if

        csvFile = csv;
        return csvFile;
    };




}//end SForce


/* function to parse Batch Job results from Salesforce */

function CSVToArray(strData, strDelimiter) {
    // Check to see if the delimiter is defined. If not,
    // then default to comma.
    strDelimiter = (strDelimiter || ",");
    // Create a regular expression to parse the CSV values.
    var objPattern = new RegExp((
        // Delimiters.
    "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +
        // Quoted fields.
    "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +
        // Standard fields.
    "([^\"\\" + strDelimiter + "\\r\\n]*))"), "gi");
    // Create an array to hold our data. Give the array
    // a default empty first row.
    var arrData = [[]];
    // Create an array to hold our individual pattern
    // matching groups.
    var arrMatches = null;
    // Keep looping over the regular expression matches
    // until we can no longer find a match.
    while (arrMatches = objPattern.exec(strData)) {
        // Get the delimiter that was found.
        var strMatchedDelimiter = arrMatches[1];
        // Check to see if the given delimiter has a length
        // (is not the start of string) and if it matches
        // field delimiter. If id does not, then we know
        // that this delimiter is a row delimiter.
        if (strMatchedDelimiter.length && (strMatchedDelimiter != strDelimiter)) {
            // Since we have reached a new row of data,
            // add an empty row to our data array.
            arrData.push([]);
        }
        // Now that we have our delimiter out of the way,
        // let's check to see which kind of value we
        // captured (quoted or unquoted).
        if (arrMatches[2]) {
            // We found a quoted value. When we capture
            // this value, unescape any double quotes.
            var strMatchedValue = arrMatches[2].replace(
                new RegExp("\"\"", "g"), "\"");
        } else {
            // We found a non-quoted value.
            var strMatchedValue = arrMatches[3];
        }
        // Now that we have our value string, let's add
        // it to the data array.
        arrData[arrData.length - 1].push(strMatchedValue);
    }
    // Return the parsed data.
    return (arrData);
}