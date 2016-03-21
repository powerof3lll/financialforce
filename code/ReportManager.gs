//******************************** Report Manager *******************************
//a Report Manager holds the user Reports
var userProperties = PropertiesService.getUserProperties();

function ReportManager(){
    var myReports; //array of Report Objs
    this.myReports = [];
    this.myBadReports = [];


    //method to get all Items from the Report Manager
    this.getAllItems = function(){
        return this.myReports;
    };

    //method to get all Bad Items from the Report Manager
    this.getBadItems = function(){
        return this.myBadReports;
    };


    //method to add an Item to the Report Manager
    //input is a report obj
    this.addItem = function(rObj){
        this.myReports.push(rObj);
    };

    //method to get an Item from the Report Manager
    //input is the ID for the Report Obj
    this.getItem = function(rID){
        for(var i = 0; i< this.myReports.length; i++){
            if(this.myReports[i].reportID == rID){
                return this.myReports[i];
            }
        }
    };

    //method to remove an Item from Report Manager
    //input is the report id for the Report Obj
    this.removeItem = function(rID){
        for(var i = 0; i< this.myReports.length; i++){

            if(this.myReports[i].reportID == rID){
                this.myReports.splice(i,1);
            }
        }
    };


    //method to get Bad report data from SF Report object
    this.getReportName = function(myID){
        var mySFobj = new SForce(CLIENT_ID, CLIENT_SECRETID, SF_VERSION, REDIRECT_URI);

        var inst = this.SFinstanceURL || userProperties.getProperty("instance_url");
        var accesstoken = this.sessionID || userProperties.getProperty("sessionId");
        var url = inst +"/services/data/v" + mySFobj.sfversion + "/sobjects/Report/"+myID;
        var options = {
            headers : { "Authorization" : "Bearer " + accesstoken },
            muteHttpExceptions: true   //catch errors quietly

        };

        Utilities.sleep(2000); //make sure all calls have completed to Report obj
        var response = mySFobj.sfAPIcall(url, options);
        var mydata = JSON.parse(response);
        return  mydata.Name;
    };//end getReportData



    //method to load Report Manager with Report Objects from Salesforce Analytics API
    this.loadReportManager = function(){
        var mySFobj = new SForce(CLIENT_ID, CLIENT_SECRETID, SF_VERSION, null);
        var inst = userProperties.getProperty("instance_url");
        var accesstoken =  userProperties.getProperty("sessionId");
        var sfversion = mySFobj.sfversion;

        var url = inst +"/services/data/v" + sfversion + "/sobjects/Report/";
        var options = {
            headers : { "Authorization" : "Bearer " + accesstoken }
        };

        var response = mySFobj.sfAPIcall(url, options);
        var mydata = JSON.parse(response);

        if( mydata != null ){
            //create a report obj from the report ID in the response and push to Report Manager
            for(var rpt in mydata.recentItems){
                var rid = mydata.recentItems[rpt].Id;
                var myReport = new rptObj.ReportObj(rid);
                var supported = myReport.isReportSupported(rid);
                if( supported == "supported" ){
                    this.myReports.push( myReport );
                } else {
                    this.myBadReports.push(rid);
                }

            }
        }//end if
    };

}