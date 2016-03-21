//******************************** Report Object *******************************
//a Report object to hold report details for a particular report as it is being viewed

var userProperties = PropertiesService.getUserProperties();

var rptObj = (function(){

    var myRO = {};

    var rID;
    var rName;
    var rType;
    var rObject;
    var rData;
    var rColumns;
    var rColumnLabels;
    var rGroupDownNames;
    var rGroupAcrossNames;
    var rAggregates;
    var rGroupingsRow;
    var rGroupingsCol;
    var rInstanceURL;
    var rDescribeURL;
    //groupings labels
    var rGroupingRowColumnsLabels;
    var rGroupingColColumnsLabels;


    //rptObj  Constructor  initial variables
    myRO.ReportObj = function(myID){

        this.rColumns = [];
        this.rColumnLabels = [];
        this.rData = [];
        this.rGroupingsRow = [];
        this.rGroupingsCol = [];
        this.rGroupingRowColumnsLabels =[];
        this.rGroupingColColumnsLabels =[];

        if(myID != "" && myID != null){
            //call populate report function with ID to get that report's data
            //when creating a lot of report objects need to wait in between
            Utilities.sleep(3000);
            this.populateReportInformation(myID);
        }
    };

    //encapsulation   gets/sets
    Object.defineProperties(myRO.ReportObj.prototype, "reportID", {
        get: function() { return (this.rID == null) ? userProperties.getProperty("reportID") : this.rID; },
        set: function(v) { this.rID = v;userProperties.setProperty("reportID", v); }
    });
    Object.defineProperties(myRO.ReportObj.prototype, "reportName", {
        get: function() { return (this.rName == null) ? userProperties.getProperty("reportName") : this.rName; },
        set: function(v) { this.rName = v;userProperties.setProperty("reportName", v); }
    });
    Object.defineProperties(myRO.ReportObj.prototype, "reportType", {
        get: function() { return (this.rType == null) ? userProperties.getProperty("reportType") : this.rType; },
        set: function(v) { this.rType = v;userProperties.setProperty("reportType", v); }
    });

    Object.defineProperties(myRO.ReportObj.prototype, "reportObject", {
        get: function() { return (this.rObject == null) ? userProperties.getProperty("reportObject") : this.rObject; },
        set: function(v) { this.rObject = v;userProperties.setProperty("reportObject", v); }
    });

    Object.defineProperties(myRO.ReportObj.prototype, "reportInstanceURL", {
        get: function() { return (this.rInstanceURL == null) ? userProperties.getProperty("reportInstanceURL") : this.rInstanceURL; },
        set: function(v) { this.rInstanceURL = v;userProperties.setProperty("reportInstanceURL", v); }
    });
    Object.defineProperties(myRO.ReportObj.prototype, "reportDescribeURL", {
        get: function() { return (this.rDescribeURL == null) ? userProperties.getProperty("reportDescribeURL") : this.rDescribeURL; },
        set: function(v) { this.rDescribeURL = v;userProperties.setProperty("reportDescribeURL", v); }
    });
    Object.defineProperties(myRO.ReportObj.prototype, "reportData", {
        get: function() { return this.rData; },
        set: function(v) { this.rData = v;userProperties.setProperty("reportData", v); }
    });
    Object.defineProperties(myRO.ReportObj.prototype, "reportColumnLabels", {
        get: function() { return this.rColumnLabels; },
        set: function(v) { this.rColumnLabels = v;  }
    });
    Object.defineProperties(myRO.ReportObj.prototype, "reportGroupingsRow", {
        get: function() { return this.rGroupingsRow; },
        set: function(v) { this.rGroupingsRow = v }
    });
    Object.defineProperties(myRO.ReportObj.prototype, "reportGroupingsCol", {
        get: function() { return this.rGroupingsCol; },
        set: function(v) { this.rGroupingsCol = v }
    });

    Object.defineProperties(myRO.ReportObj.prototype, "reportGroupingRowColumnsLabels", {
        get: function() { return this.rGroupingRowColumnsLabels; },
        set: function(v) { this.rGroupingRowColumnsLabels = v }
    });
    Object.defineProperties(myRO.ReportObj.prototype, "reportGroupingColColumnsLabels", {
        get: function() { return this.rGroupingColColumnsLabels; },
        set: function(v) { this.rGroupingColColumnsLabels = v;  }
    });

    Object.defineProperties(myRO.ReportObj.prototype, "reportAggregates", {
        get: function() { return this.rAggregates; },
        set: function(v) { this.rAggregates = v;userProperties.setProperty("reportAggregates", v); }
    });



    myRO.ReportObj.prototype.populateReportInformation = function(myID){
        var mydata = this.getReportInformation(myID);

        if( mydata != null ){
            this.reportID = mydata['attributes'].reportId;
            this.reportName = mydata['attributes'].reportName;
            this.reportType = mydata.reportMetadata.reportFormat;
            this.reportObject = mydata.reportMetadata.reportType.label;
            this.reportInstanceURL = mydata['attributes'].instancesUrl;
            this.reportDescribeURL = mydata['attributes'].describeUrl;
            this.reportGroupingsRow = mydata.groupingsDown.groupings;
            this.reportGroupingsCol = mydata.groupingsAcross.groupings;
            this.reportAggregates = mydata.reportMetadata.aggregates;


            this.rColumns = mydata.reportMetadata.detailColumns;
            this.reportColumnLabels = [];
            for(var rc in this.rColumns){
                this.reportColumnLabels.push( mydata['reportExtendedMetadata']['detailColumnInfo'][this.rColumns[rc]].label );
            }

            this.rGroupDownNames = mydata.reportMetadata.groupingsDown;
            var gname, glabel;
            this.reportGroupingRowColumnsLabels = [];
            for(var grc in  this.rGroupDownNames){
                gname = this.rGroupDownNames[grc].name;
                glabel = mydata['reportExtendedMetadata']['groupingColumnInfo'][gname].label;
                this.reportGroupingRowColumnsLabels.push(glabel);
            }


            this.rGroupAcrossNames = mydata.reportMetadata.groupingsAcross;
            this.reportGroupingColColumnsLabels = [];
            for(var gcc in this.rGroupAcrossNames){
                gname = this.rGroupAcrossNames[gcc].name;
                glabel = mydata['reportExtendedMetadata']['groupingColumnInfo'][gname].label;
                this.reportGroupingColColumnsLabels.push(glabel);
            }

        }
    }; //end populateReportInformation


    //method to get report information from SF and store in this object
    myRO.ReportObj.prototype.getReportInformation = function(myID){
        var mySFobj = new SForce(CLIENT_ID, CLIENT_SECRETID, SF_VERSION, REDIRECT_URI);
        var inst = mySFobj.SFinstanceURL || userProperties.getProperty("instance_url");
        var accesstoken = mySFobj.sessionID || userProperties.getProperty("sessionId");

        var url = inst + "/services/data/v" + mySFobj.sfversion + "/analytics/reports/"+myID;
        var options = {
            headers : { "Authorization" : "Bearer " + accesstoken },
            muteHttpExceptions: true  //catch errors quietly
        };

        var response =  null;


        var tempResponse = mySFobj.sfAPIcall(url, options);
        tempResponse = JSON.parse(tempResponse);

        var n = JSON.stringify(tempResponse).search("errorCode");
        if ( n > 0 ){
            // then an error was returned save in db
            var badreport = { "Id": myID, "message": tempResponse };
        }  else {
            response = tempResponse
        }

        return response;
    }; //end getReportInformation


    myRO.ReportObj.prototype.isReportSupported = function(myID){
        var supported = null;
        var reportinfo = this.getReportInformation(myID);
        if( reportinfo != null ){
            supported = "supported";
        }
        return supported;
    }



    //method to get report data from SF Analytical API
    myRO.ReportObj.prototype.getReportData = function(){
        var mySFobj = new SForce(CLIENT_ID, CLIENT_SECRETID, SF_VERSION, REDIRECT_URI);
        var inst = mySFobj.SFinstanceURL || userProperties.getProperty("instance_url");
        var accesstoken = mySFobj.sessionID || userProperties.getProperty("sessionId");

        var url = inst + "/services/data/v" + mySFobj.sfversion + "/analytics/reports/"+this.reportID+"?includeDetails=true";
        var options = {
            headers : { "Authorization" : "Bearer " + accesstoken },
            muteHttpExceptions: true   //catch errors quietly

        };

        var response = mySFobj.sfAPIcall(url, options);
        return response;
    };//end getReportData


    //method to parse returned data for view on GSheet
    myRO.ReportObj.prototype.parseReportData = function(){
        var myReportData = JSON.parse(this.getReportData());
        if(this.reportType == "TABULAR"){
            var myreportTypeCode = "T!T";

            var dataRows = myReportData['factMap'][myreportTypeCode].rows;
            var newRow = [];

            for(var i = 0; i < dataRows.length; i++){
                for(var row in dataRows[i]){
                    var newCol = [];
                    for(var cell in dataRows[i][row]){
                        newCol.push(dataRows[i][row][cell].label );
                    }
                    newRow.push(newCol);
                }
            }

            //totals row
            var totalRowHeader = myReportData['reportExtendedMetadata']['aggregateColumnInfo'][this.reportAggregates[0]].label;
            var totalRowValue = myReportData['factMap'][myreportTypeCode].aggregates[0].label;
            newRow.push( [totalRowHeader, totalRowValue] );
        }

        this.rData = newRow;
    }; //end parseReportData



    //method used by report objs to print report heading
    myRO.ReportObj.prototype.printReportHeading = function(){
        var GSheet = SpreadsheetApp.getActiveSpreadsheet();
        var activeReportSheet;

        if( GSheet.getSheetByName(this.reportName)){

            activeReportSheet = GSheet.getSheetByName(this.reportName);
            activeReportSheet.clear();
        } else {
            activeReportSheet = GSheet.insertSheet(this.reportName, GSheet.getNumSheets());
        }

        //show Report Title on Spreadsheet cell A1
        var titleRow = activeReportSheet.getRange("A1:K1").mergeAcross().setValue(this.reportName);
        titleRow.setBackground("#00396B").setFontColor("white").setFontSize(24);
        var reportTypeRow = activeReportSheet.getRange("A2:F2").mergeAcross().setValue("Report Type: " + this.reportType);

        return activeReportSheet;
    }; //end printReportHeading


    myRO.ReportObj.prototype.printReportHeaderColumns = function(activeSheet){
        //show Headers on Spreadsheet beginning A4
        var headersRow = activeSheet.getRange("A4");
        for(var r in this.reportColumnLabels ){
            headersRow.offset(0, r).setValue( this.reportColumnLabels[r] );
            headersRow.offset(0, r).setBackground("#0C8EFF").setFontColor("white").setFontWeight("bold");

        }
    };

    //POST request to run report asynch
    myRO.ReportObj.prototype.runReportAsyncInstances = function(rid){
        var mySFobj = new SForce(CLIENT_ID, CLIENT_SECRETID, SF_VERSION, REDIRECT_URI);
        var inst = mySFobj.SFinstanceURL || userProperties.getProperty("instance_url");
        var accesstoken = mySFobj.sessionID || userProperties.getProperty("sessionId");
        var response = null;
        var reportData = this.getReportInformation(rid);
        var reportMetaData = reportData.reportMetadata;
        var payload = { "reportMetadata" : reportMetaData };

        var myPayLoader = JSON.stringify(payload);


        if( payload != null ){
            //schedule report asynch list
            var url = inst + "/services/data/v" + mySFobj.sfversion + "/analytics/reports/"+rid+"/instances/";


            //request data from SF
            var options = {
                headers : { "Authorization" : "Bearer " + accesstoken,
                    "Content-Type" : "application/json",
                },
                payload : myPayLoader,
            };

            response = mySFobj.sfAPIcall(url, options);
        }
        //return response with instance id in it
        return response;

    }



//GET request for a list of instances for selected report
    myRO.ReportObj.prototype.getReportAsyncInstances = function(rid){

        var mySFobj = new SForce(CLIENT_ID, CLIENT_SECRETID, SF_VERSION, REDIRECT_URI);
        var inst = mySFobj.SFinstanceURL || userProperties.getProperty("instance_url");
        var accesstoken = mySFobj.sessionID || userProperties.getProperty("sessionId");

        //schedule report asynch list
        var url = inst + "/services/data/v" + mySFobj.sfversion + "/analytics/reports/"+rid+"/instances/";
        var options = {
            headers : { "Authorization" : "Bearer " + accesstoken },
            muteHttpExceptions: true   //catch errors quietly

        };

        var response = mySFobj.sfAPIcall(url, options);


        //response will have instance run details:  status, complete date
        return response;

    }


    //method to print report data on GSheet
    myRO.ReportObj.prototype.printReport = function(){

        var activeReportSheet = this.printReportHeading();
        this.printReportHeaderColumns(activeReportSheet);
        var groupings = this.parseReportData(); ///gets data from SF and parse for view
        activeReportSheet.getRange(5,1,activeReportSheet.getMaxRows()-4,activeReportSheet.getMaxColumns()).setFontSize("9");
        //print out rows
        for(var e in this.rData){
            activeReportSheet.appendRow(this.rData[e]);
        }
      
          
        //find groupings rows and format
        var startrng, endrng, groupRang;
        if(typeof(groupings) != "undefined"){
            var data = activeReportSheet.getDataRange().getValues();
        
            for(var n  in data){
                for(var x in groupings){
                    if( data[n][0] == groupings[x].gheader  ){
                        //this is the grouping row
                        var myrow = ++n;
                        startrng =  activeReportSheet.getRange(myrow,1);
                        endrng =activeReportSheet.getRange(myrow, activeReportSheet.getLastColumn());
                        groupRang = activeReportSheet.getRange(startrng.getA1Notation() + ":" + endrng.getA1Notation());
                        groupRang.setBackground("cyan").setFontColor("black");
                        break;
                    }
                }
            }
        }

        //format last row
        var lastRowStart = activeReportSheet.getRange(activeReportSheet.getLastRow(), 1);
        var lastRowEnd = activeReportSheet.getRange(activeReportSheet.getLastRow(),activeReportSheet.getLastColumn());
        var lastrow =  activeReportSheet.getRange(lastRowStart.getA1Notation() + ":" + lastRowEnd.getA1Notation());
        lastrow.setBackground("#0C8EFF").setFontColor("white").setFontWeight("bold");

        activeReportSheet.activate();
        return {message:"Report Data Printed!"};

    }; //end printReport

    return myRO;
}()); //end rptObj






////////////////// SUMMARY Report Obj  ///////////////////////////////////
//SUMMARY Report Obj extends the Report Obj but adds properties for groupings
var srtObj = (function(){

    var mySRO = {};

    //mtrObj  Constructor  initial variables
    mySRO.SummaryReport = function(myID){
        rptObj.ReportObj.call(this, myID);
    };

    //Summary report extends the Report Object
    mySRO.SummaryReport.prototype = new rptObj.ReportObj();
    mySRO.SummaryReport.prototype.constructor = mySRO.SummaryReport;


    //overload parseReportData method to display Summary report with groupings
    mySRO.SummaryReport.prototype.parseReportData = function(){
        var myReportData = JSON.parse(this.getReportData());
        if(this.reportType == "SUMMARY"){
            var mytotalCode = "T!T";
            var newRow = [];
            var groupHeadings = [];


            for(var n=0; n < this.reportGroupingsRow.length; n++){
                var mygroupingCode = n + "!T";

                //push Groupings Row Label and Aggregate   then the rows of data
                var groupingHeader = this.reportGroupingRowColumnsLabels[0];
                var groupingHeaderValue = this.reportGroupingsRow[n].label;
                var groupingHeaderAggregate = myReportData['factMap'][mygroupingCode].aggregates[0].label;

                newRow.push( [groupingHeader + ": " + groupingHeaderValue + " (" +  groupingHeaderAggregate + " records)"]   );
                groupHeadings.push( { "gheader" : [groupingHeader + ": " + groupingHeaderValue + " (" +  groupingHeaderAggregate + " records)"]} );


                var dataRows = myReportData['factMap'][mygroupingCode].rows;
                for(var i = 0; i < dataRows.length; i++){
                    for(var row in dataRows[i]){
                        var newCol = [];
                        for(var cell in dataRows[i][row]){
                            newCol.push(dataRows[i][row][cell].label );
                        }
                        newRow.push(newCol);

                    }
                }
            }//end groupings for


            //totals row
            var totalRowHeader = myReportData['reportExtendedMetadata']['aggregateColumnInfo'][this.reportAggregates[0]].label;
            var totalRowValue = myReportData['factMap']["T!T"].aggregates[0].label;
            newRow.push( [totalRowHeader, totalRowValue] );
        }

        this.rData = newRow;
        return groupHeadings;

    };

    return mySRO;
}());





////////////////// Matrix Report Obj  ///////////////////////////////////
//Matrix Report Obj extends the Report Obj but adds properties for groupings
var mtxObj = (function(){

    var myMRO = {};

    //mtrObj  Constructor  initial variables
    myMRO.MatrixReport = function(myID){
        rptObj.ReportObj.call(this, myID);
    };

    //Matrix report extends the Report Object
    myMRO.MatrixReport.prototype = new rptObj.ReportObj();
    myMRO.MatrixReport.prototype.constructor = myMRO.MatrixReport;

    //overload headers method to display Matrix groupings columns
    myMRO.MatrixReport.prototype.printReportHeaderColumns = function(activeSheet){
        //show Headers on Spreadsheet beginning A4
        var headersRow = activeSheet.getRange("A4");
        var myheaders = [];

        for(var d in this.reportGroupingRowColumnsLabels){
            myheaders.push({"groupheader": this.reportGroupingRowColumnsLabels[d] });
        }

        for(var a in this.reportGroupingColColumnsLabels){
            myheaders.push({"groupheader": this.reportGroupingColColumnsLabels[a] });
        }

        var cols = this.reportColumnLabels.splice(0,1);   //remove first header as not necessary in matrix format
        for(var r in this.reportColumnLabels ){
            myheaders.push( this.reportColumnLabels[r] );
        }

        for(var r in myheaders ){
            if( myheaders[r].groupheader){
                headersRow.offset(0, r).setValue( myheaders[r].groupheader );
                headersRow.offset(0, r).setBackground("navy").setFontColor("white").setFontWeight("bold");

            } else {
                headersRow.offset(0, r).setValue( myheaders[r] );
                headersRow.offset(0, r).setBackground("black").setFontColor("white").setFontWeight("bold");
            }
        }

    };

    //overload parseReportData method to display Matrix report with groupings
    myMRO.MatrixReport.prototype.parseReportData = function(){
        var myReportData = JSON.parse(this.getReportData());

        if(this.reportType == "MATRIX"){
            var myTotalCode;
            var newRow = [];
            var groupHeadings = [];

            for(var n=0; n < this.reportGroupingsRow.length; n++){
                myTotalCode = n + "!T";
                var groupingHeaderValue = this.reportGroupingsRow[n].label;
                var groupingAggregateValue = myReportData['factMap'][myTotalCode]['aggregates'][0].value;
                groupHeadings.push( {"gheader": groupingHeaderValue, "gvalue":groupingAggregateValue});
                newRow.push( [groupingHeaderValue, groupingAggregateValue] );


                var mygroupingCode = n + "!0";
                var dataRows = myReportData['factMap'][mygroupingCode].rows;
                for(var i = 0; i < dataRows.length; i++){
                    for(var row in dataRows[i]){
                        var newCol = [];
                        for(var cell in dataRows[i][row]){
                            if(cell == 1){
                                //skip for col groupings
                                newCol.push("1");
                            }
                            newCol.push(dataRows[i][row][cell].label );

                        }
                        newRow.push(newCol);
                    }
                }
            }//end groupings for


            //totals row
            var totalRowHeader = myReportData['reportExtendedMetadata']['aggregateColumnInfo'][this.reportAggregates[0]].label;
            var totalRowValue = myReportData['factMap']["T!T"].aggregates[0].label;
            newRow.push( [totalRowHeader, totalRowValue] );
        }

        this.rData = newRow;
        return groupHeadings;

    };


    //overload printReport method to display Matrix report with groupings
    //  myMRO.MatrixReport.prototype.printReport = function(){

    //  };

    return myMRO;
}());
