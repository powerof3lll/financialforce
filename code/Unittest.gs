
//Unit Tests for gas_sf project  includes tests for each object in the project - salesforceObj, QueryObj, GoogleSheetObj
//gas_unittests.gs

//main function    change name to TTdoGet when running original project
function XdoGet( e ) {

    //QAS.urlParams( e.parameter );
    QAS.config({ title: "Salesforce | Google App Script Program Test Suite" });
    QAS.load( myGAStests );
    return QAS.getHtml();

};


// Imports the following functions:
// ok, equal, notEqual, deepEqual, notDeepEqual, strictEqual,
// notStrictEqual, throws, module, test, asyncTest, expect
QAS.helpers(this);


function myGAStests() {


///////////////////////////QueryGenerator  ////////////////////////////////////////
    module("QueryGenerator");

    test("Get available objects", 1, function(assert) {
        var objects = getObjects();
        assert.ok(objects != null, "SForce objects exists!");
    });

    test("Get available fields", 1, function(assert) {
        var object = 'Account';
        var fields = getSObjectsFields(object);
        assert.ok(fields != null, "SForce fields exists!");
    });

    test("Get available childRelationships", 1, function(assert) {
        var object = 'Account';
        var childRelationships = getSObjectsChildRelationships(object);
        assert.ok(childRelationships != null, "SForce childRelationships exists!");
    });
  
    test("Get saved queries", 1, function(assert) {
        var queryList = getSavedQueryListFromSheet();
        assert.ok(queryList != null, "Got queries from sheet!");
    });
  
    test("Run query", 1, function(assert) {
        var query = 'SELECT Id FROM Contact';
        var records = querySF(query);
        assert.ok(records != null, "SForce query runs!");
    });


///////////////////////////External Query List  ////////////////////////////////////////
    module( "External query list");

    test("External sheet defined", 1, function(assert){
        assert.ok(typeof(QUERY_SHEET) == "string" && QUERY_SHEET.length > 10  , 'ok');
    });

    test("External sheet usable", 3, function(assert){
        assert.ok( SpreadsheetApp.openByUrl(QUERY_SHEET), "Document accessible");
        assert.ok( SpreadsheetApp.openByUrl(QUERY_SHEET).getSheetByName("Sheet1"), "Sheet present");
        assert.ok( SpreadsheetApp.openByUrl(QUERY_SHEET).getSheetByName("Sheet1").getLastRow()>0, "Sheet not empty");

    });




    //////////////////////////Chunker Object //////////////////////////////////////////////////////

    module( "Chunker");

    var chunker = {
        "sheetName" : "myChunkSheet",
        "parent" : "Invoice",
        "child" : "Invoice_Line_Item",
        "chunkby" : "NUM",
        "chunkbyValue" : "4",
        //  "headers" : "Invoice.Id,Invoice.Date__c,Invoice.Description__c,Invoice_Line_Item.Id,Invoice_Line_Item.Product,Invoice_Line_Item.Quantity,Invoice_Line_Item.Unitprice"
        "pheaders" : ["Invoice.Id","Invoice.Date__c","Invoice.Description__c"],
        "cheaders" : ["Invoice_Line_Item.Id","Invoice_Line_Item.Product","Invoice_Line_Item.Quantity","Invoice_Line_Item.Unitprice"]
    };

    var chk = new Chunker(chunker.sheetName, chunker.parent, chunker.child, chunker.chunkby, chunker.chunkbyValue, chunker.pheaders,chunker.cheaders);

    // If no Active Spreadsheet Present , then create a new one .
    var ssNew = SpreadsheetApp.getActiveSpreadsheet();
    if(ssNew);
    else
        ssNew = SpreadsheetApp.create("QUnitTest1");
    SpreadsheetApp.setActiveSpreadsheet(ssNew);

    test("Chunker Obj Exists", 1, function(assert){
        assert.ok(chk, 'Chunker Object exists');
    });

    test("Chunker Obj Initialized", 7, function(assert){
        assert.equal( chk.chunkSheetName, 'myChunkSheet', "Chunk Sheet Name initialized");
        assert.equal( chk.chunkParentObj, 'Invoice', "Chunker Parent Object initialized");
        assert.equal( chk.chunkChildObj, 'Invoice_Line_Item', "Chunker Child Object initialized");
        assert.equal( chk.chunkByCriteria, 'NUM', "Chunker criteria initialized");
        assert.equal( chk.chunkByValue, 4, "Chunker criteria value initialized");
        assert.equal( chk.parentHeaders.length, 4, "Chunk Parent Headers initialized");
        assert.equal( chk.childHeaders.length, 4, "Chunk Child Headers initialized");
    });


    test("Chunker Creates Sheet with headers", 3, function(assert){

        var result = chk.createChunkSheet();
        assert.equal( result.getActiveSheet().getName(), 'myChunkSheet', 'Chunker Sheet Created');
        assert.equal( result.getActiveSheet().getRange("A1").getValue(), 'Invoice.Id', 'Header Row has Parent values');
        assert.equal( result.getActiveSheet().getRange("E1").getValue(), 'Invoice_Line_Item.Id', 'Header Row has Child values');

        var resSheet = result.getActiveSheet();
        resSheet.appendRow(["02i90000000CcbEAAS","12-03-2014","Test 1","","02i90000000CcbEABS","TestP1","2","10"]);
        resSheet.appendRow(["02i90000000CcbEAAT","12-03-2014","Test 2","","02i90000000CcbEABT","TestP2","3","11"]);
        resSheet.appendRow(["02i90000000CcbEAAU","12-03-2014","Test 3","","02i90000000CcbEABU","TestP3","4","12"]);
        resSheet.appendRow(["02i90000000CcbEAAV","12-03-2014","Test 4","","02i90000000CcbEABV","TestP4","5","13"]);
        resSheet.appendRow(["02i90000000CcbEAAW","12-03-2014","Test 5","","02i90000000CcbEABW","TestP5","6","14"]);
    });

    test("Chunker Get Information", 2, function(assert){
        chk.getChunkInformation();
        assert.equal( chk.totalDataSize,5, 'Total number of data rows chunked by Row');
        assert.ok ( chk.chunkRanges, 'Chunk Ranges are defined');
    });

    test("Chunk by Row Ranges", 1, function(assert){
        chk.getChunkInformation();
        var chunkSize = 4;
        var chunkRange = findChunksByRow(chk.chunkData, chunkSize);
        var numberofChunks = chunkRange.length;

        assert.equal( numberofChunks, 2, 'Number of Chunks');

    });

    test("Chunk by Column Ranges", 1, function(assert){
        chk.getChunkInformation();
        var columnValue = 'Invoice_Line_Item.Product';
        var chunkRange = findChunksByCol(chk.chunkData, columnValue,chk.childHeaders,chk.uniqueParentColumnIdentifier,chk.uniqueChildColumnIdentifier);
        var numberofChunks = chunkRange.length;
        assert.equal( numberofChunks, 5, 'Total number of data rows chunked by Column');

    });



    ////////////////////////////Report Object///////////////////////////////////////
    module("ReportObj");
    test("Report Obj Exists", 1, function(assert){
        var rpt = new rptObj.ReportObj();
        assert.ok(rpt);
    });

    test("Report Properties Set", 5, function(assert){
        var rpt = new rptObj.ReportObj();
        rpt.reportName = "Accounts Report";
        rpt.reportType = "Tabular";
        rpt.reportObject = "Account";
        rpt.reportInstanceURL = "/services/data/v29.0/reports/instances";
        rpt.reportDescribeURL = "/services/data/v29.0/reports/describe";

        assert.equal(rpt.reportName, "Accounts Report", "Report Name set OK!");
        assert.equal(rpt.reportType, "Tabular", "Report Type set OK!");
        assert.equal(rpt.reportObject, "Account", "Report Object set OK!");
        assert.equal(rpt.reportInstanceURL, "/services/data/v29.0/reports/instances", "Instance URL set!");
        assert.equal(rpt.reportDescribeURL, "/services/data/v29.0/reports/describe", "Describe URL set!");
    });

    test("Summary Report Obj Exists", 3, function(assert){
        var srpt = new srtObj.SummaryReport();
        srpt.reportName = "Opportunities Report";
        srpt.reportType = "Summary";
        srpt.reportObject = "Opportunity";

        assert.equal(srpt.reportName, "Opportunities Report", "Summary Report Name set ok!");
        assert.equal(srpt.reportType, "Summary", "Summary Report Type set ok!");
        assert.equal(srpt.reportObject, "Opportunity", "Summary Report Object set ok!");
    });


    test("Matrix Report Obj Exists", 3, function(assert){
        var mxrpt = new mtxObj.MatrixReport();
        mxrpt.reportName = "Opportunities Matrix Report";
        mxrpt.reportType = "Matrix";
        mxrpt.reportObject = "Opportunity";

        assert.equal(mxrpt.reportName, "Opportunities Matrix Report", "Matrix Report Name set ok!");
        assert.equal(mxrpt.reportType, "Matrix", "Matrix Report Type set ok!");
        assert.equal(mxrpt.reportObject, "Opportunity", "Matrix Report Object set ok!");
    });







    //////////////////////////Report Manager //////////////////////////////////////////////////////
    //module Report Manager unit tests
    module("Report Manager");

    test("Report Manager Created Has no Items", 2, function(assert) {
        //var rptTest = new rptManager.RM();
        var rptTest = new ReportManager();
        assert.ok(rptTest,"Report Manager exists!");
        var numItems = rptTest.getAllItems();
        assert.equal(numItems.length, 0, "Report Manager has no items!");

    });

    test("Report Manager - Add an Item", 1, function(assert){
        var rptTest = new ReportManager();
        var myReportObj = new rptObj.ReportObj();
        myReportObj.reportName = "Accounts Report";
        myReportObj.reportID = "1Z1234129879";

        rptTest.addItem(myReportObj);
        var numItems = rptTest.getAllItems();
        assert.equal(numItems.length, 1, "Report manager has 1 item!");
    });

    test("Report Manager - Get an Item", 2, function(assert){
        var rptTest = new ReportManager();
        var myReportObj = new rptObj.ReportObj();
        myReportObj.reportName = "Accounts Report";
        myReportObj.reportID = "1Z1234129879";

        rptTest.addItem(myReportObj);
        var myRPTobj = rptTest.getItem("1Z1234129879");
        assert.equal(myRPTobj.reportName, "Accounts Report", "Report Item returned!" + myRPTobj.reportName);
        assert.equal(myRPTobj.reportID, "1Z1234129879", "Item ID returned is ok!" + myRPTobj.reportID);

    });

    test("Report Manager - Remove an Item", 1, function(assert){
        var rptTest = new ReportManager();
        var myReportObj = new rptObj.ReportObj();
        myReportObj.reportName = "Accounts Report";
        myReportObj.reportID = "1Z1234129879";

        rptTest.addItem(myReportObj);
        rptTest.removeItem("1Z1234129879");
        var numItems = rptTest.getAllItems();
        assert.equal(numItems.length, 0, "Report Item was removed!");
    });



    ////////////////////////////GSheet Object///////////////////////////////////////
    module("GSheet");


    test("Ui App was created", 1, function(assert){
        var app = UiApp.createApplication().setTitle("SF_GAS_Test App");
        assert.equal(UiApp.getActiveApplication(), "UiApplication", "Ui App was created");
    });



    ////////////////////////Salesforce Object///////////////////////////////////////////
    module("SFobj");

    test("SF obj created", 1, function(assert){
        PropertiesService.getUserProperties().deleteAllProperties();
        var mySFObj = new SForce(CLIENT_ID,CLIENT_SECRETID,SF_VERSION,REDIRECT_URI);
        assert.ok(mySFObj, "SForce object exists!");
    });

    test("SF obj ids set", 2, function(assert){
        PropertiesService.getUserProperties().deleteAllProperties();
        var mySFObj = new SForce(CLIENT_ID,CLIENT_SECRETID,SF_VERSION,REDIRECT_URI);
        assert.equal(CLIENT_ID, mySFObj.sfcid, "client IDs setting/getting");
        assert.equal(CLIENT_SECRETID, mySFObj.sfcsid, "client Secret ID setting/getting");

    });




///////////////////////////Query Object ////////////////////////////////////////
    //module QueryObj unit tests
    module("QYobj");

    test("Query obj created", 2, function(assert){
        PropertiesService.getUserProperties().deleteAllProperties();
        var q = "SELECT Id, Name, Description FROM Accounts LIMIT 5";
        var myQueryObj = new Query(q);
        myQueryObj.setParams(q);
        assert.ok(myQueryObj, "Query Obj object exists!");
        assert.equal(myQueryObj.qcolumns.length, 3, "Query created has 3 Columns");

    });


    test("Query Object - Get rawQuery", 1, function(assert){
        var qob = "SELECT Id, Name, Description FROM Accounts LIMIT 5";
        var myQueryObj = new Query(qob);

        //assert.equal(qob, myQueryObj.rawQuery, "Query Obj rawQuery ok!");
        assert.ok(myQueryObj);

    });




}// end myTests func
