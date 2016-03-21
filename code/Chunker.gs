//******************************** Chunker Object *******************************
//a Chunker object to hold data details for a sheet to be chunked

var chunkCtr = 1;

function Chunker(mySheet, parent, child, chunkby, chunkbyval, pheaders, cheaders, templateTotalCol, templateLabel1, templateFormula1, templateLabel2, templateFormula2){
    this.chunkSheetName = mySheet;
    this.chunkParentObj = parent;
    this.chunkChildObj = child;
    this.chunkByCriteria = chunkby;
    this.chunkByValue = chunkbyval;
    this.chunkRanges  = [];
    this.parentHeaders = pheaders;
    this.childHeaders = cheaders;
    pheaders.push("**");
    this.chunkHeaders = pheaders.concat(cheaders);
    this.totalDataSize=0;
    this.chunkData=null;
    this.chunkRanges=null;

    //template variables
    this.templateTotalColumn = templateTotalCol;
    this.templateLb1 = templateLabel1;
    this.templateFormula1= templateFormula1;
    this.templateLb2 = templateLabel2;
    this.templateFormula2= templateFormula2;
  
    // Variables to store what Unique Identifiers to be used to distinguish Column Names of Parent from the Child.
    this.uniqueChildColumnIdentifier = "C.";
    this.uniqueParentColumnIdentifier = "M.";

    this.headersMap = createHashMap(this.chunkHeaders);
  
    //Below will create a HashMap,when there are both Parent and Child , specially to consider cases where both Parent and Child have the same column name.
    this.headersMap2 = createUniqueColumnsNames(pheaders,cheaders,this.uniqueParentColumnIdentifier,this.uniqueChildColumnIdentifier);
  
    //Map for Parent and Child Headers
    this.parentHeadersMap = createHashMap(pheaders);
    this.childHeadersMap = createHashMap(cheaders);
  

    

    //encapsulation   gets/sets
    Object.defineProperties(this, "chunkSheetName", {
        get: function() { return this.chunkSheetName; },
        set: function(v) { this.chunkSheetName = v; }
    });

    //data objects from chunker sheet
    Object.defineProperties(this, "chunkData", {
        get: function() { return this.chunkData; },
        set: function(v) { this.chunkData = v;  }
    });

    //data objects from chunker sheet
    Object.defineProperties(this, "childData", {
        get: function() { return this.childRows; },
        set: function(v) { this.childRows = v;  }
    });

    //data objects from chunker sheet
    Object.defineProperties(this, "templateTotalColumn", {
        get: function() { return this.templateTotalColumn; },
        set: function(v) { this.templateTotalColumn = v;  }
    });

    //data objects from chunker sheet
    Object.defineProperties(this, "templateLabel1", {
        get: function() { return this.templateLb1; },
        set: function(v) { this.templateLb1 = v;  }
    });

    //data objects from chunker sheet
    Object.defineProperties(this, "templateFormula1", {
        get: function() { return this.templateFormula1; },
        set: function(v) { this.templateFormula1 = v;  }
    });
    //data objects from chunker sheet
    Object.defineProperties(this, "templateLabel2", {
        get: function() { return this.templateLb2; },
        set: function(v) { this.templateLb2 = v;  }
    });

    //data objects from chunker sheet
    Object.defineProperties(this, "templateFormula2", {
        get: function() { return this.templateFormula2; },
        set: function(v) { this.templateFormula2 = v;  }
    });

    //data objects from chunker sheet
    Object.defineProperties(this, "parentData", {
        get: function() { return this.parentRows; },
        set: function(v) { this.parentRows = v;  }
    });
    //Total number of data rows on Chunker sheet
    Object.defineProperties(this, "totalChunkDataRows", {
        get: function() { return this.totalDataSize; },
        set: function(v) { this.totalDataSize = v;  }
    });

    //Chunk Ranges of data set
    Object.defineProperties(this, "chunkRanges", {
        get: function() { return this.chunkRanges; },
        set: function(v) { this.chunkRanges = v;  }
    });

    Object.defineProperties(this, "parentHeaders", {
        get: function() { return this.parentHeaders; },
        set: function(v) { this.parentHeaders = v;  }
    });

    Object.defineProperties(this, "childHeaders", {
        get: function() { return this.childHeaders; },
        set: function(v) { this.childHeaders = v;  }
    });


    /*
     * Method to set the value for chunkSheetName
     */

    this.setSheetName = function(sheetname){
        this.chunkSheetName = sheetname;
    };


    /*
     * This method will copy the data to PreviewChunk_<SheetName> where Sheet Name contains the Original Data .
     * This new sheet will then be used for getting the chunk information and accordingly display it.
     * Parameter will be the name of the Sheet that contains the Original Data i.e. SheetName

     * Imp : Invoke this function first , before starting for getting Chunks
     */

    this.prepareChunkSheets = function(originalSheet)
    {       
        var resultSheetName = this.chunkSheetName;
        var GSheet = SpreadsheetApp.getActiveSpreadsheet();
        var newSheet;
        var deleteSheet;
        if(GSheet.getSheetByName(resultSheetName)){                       
           deleteSheet = GSheet.getSheetByName(resultSheetName).setName(resultSheetName + "_ToDelete");
        }       
                
        
     //   GSheet = SpreadsheetApp.getActiveSpreadsheet();
        newSheet = GSheet.getSheetByName(originalSheet);
        // Copying the Sheet
        newSheet = newSheet.copyTo(GSheet).setName(this.chunkSheetName);
      
        if(deleteSheet)
         GSheet.deleteSheet(deleteSheet);
        newSheet.activate();
      
        //Setting the Identifiers for Parent/Child - to TemplateTotalColumn , TemplateLabel1 and TemplateLabel2
        
        //Case 1 : Check if Child Exists or not 
        if(this.childHeaders && this.childHeaders.length > 0)
        {
          // For TemplateTotalColumn
          if(this.templateTotalColumn != null)
          {
           if(this.templateTotalColumn.indexOf(this.uniqueParentColumnIdentifier) == -1 && this.templateTotalColumn.indexOf(this.uniqueChildColumnIdentifier) == -1)
            this.templateTotalColumn = this.uniqueChildColumnIdentifier + this.templateTotalColumn; 
          }
        }
      // Case 2 : Only Parent Exists       
        else
        {
          // For TemplateTotalColumn
          // If Parent Identifier is provided , then remove it . 
          if(this.templateTotalColumn != null)
          {
           if(this.templateTotalColumn.indexOf(this.uniqueParentColumnIdentifier) != -1)
            this.templateTotalColumn = this.templateTotalColumn.substring(2);
          }
        }
       

    };


    /*
     *  method creates the Sheet that will hold the Chunk data when the UI is used
     *  
     */

    this.createChunkSheet = function(){

        var resultSheetName = this.chunkSheetName;
        var newSheet = SpreadsheetApp.getActiveSpreadsheet();
        var GSheet = SpreadsheetApp.getActiveSpreadsheet();

        if(GSheet.getSheetByName(resultSheetName)){
            GSheet.getSheetByName(resultSheetName).clear();  //clear sheet before new results
            newSheet = GSheet.setActiveSheet(GSheet.getSheetByName(resultSheetName));
            newSheet.activate();
        } else {
            newSheet = GSheet.insertSheet(resultSheetName);
        }
       // data font size set to 9
         newSheet.getRange(2,1,newSheet.getMaxRows()-1,newSheet.getMaxColumns()).setFontSize("9");
        //add row headers to sheet
        //print header row
        if(this.chunkHeaders.length > 0){
            headersRange = newSheet.getRange(1, 1, 1, this.parentHeaders.length);
            headersRange.setValues([this.parentHeaders]);
            headersRange = newSheet.getRange(1, this.parentHeaders.length+1, 1, this.childHeaders.length );

            headersRange.setValues([this.childHeaders]);

            headersRange = newSheet.getRange(1, 1, 1, newSheet.getLastColumn());
            headersRange.setBackground("#0C8EFF");
        }


        // Returns the Active SpreadSheet 
        return GSheet;
    }; //end createChunkSheet



    /*
     *  method gathers information about chunk data i.e.  total number of chunks based on criteria
     */
    this.getChunkInformation = function(){
        var resultSheetName = this.chunkSheetName;
        var GSheet = SpreadsheetApp.getActiveSpreadsheet();
        var cSheet = GSheet.setActiveSheet(GSheet.getSheetByName(resultSheetName));


        //find data range on Chunk Sheet
        var mystartRange = cSheet.getRange(2,1);
        var myendRange = cSheet.getRange(cSheet.getLastRow(), cSheet.getLastColumn());
        var myrange = cSheet.getRange(mystartRange.getA1Notation() + ":" + myendRange.getA1Notation());



        if( this.childHeaders.length > 0 ){

            var myPCRanges = getParentChildRange(cSheet);
            var parentrange = cSheet.getRange(myPCRanges.parent.getA1Notation());
            this.parentRows = getRowsData(cSheet, parentrange, 1);

            var childrange = cSheet.getRange(myPCRanges.child.getA1Notation());
            this.childRows = getRowsData(cSheet, childrange, 1);

            var chunkedData = [];
            var temprow = [];
            var tempstring = null;
            var object = {};

            for( var i=0; i < this.parentRows.length; i++)
            {
                var object = {};
              
              //For all Parent Headers , rename them as uniqueParentColumnIdentifier<Column Name> where uniqueParentColumnIdentifier is a configured Prefix
              // Example : If Parent Column Name is Desc__c and uniqueParentColumnIdentifier = 'M." , then rename it to as M.Desc__c 
              // This helps in easy identification of the column names during Chunking .  
              
                for( var h in this.parentHeaders ){
                    if( typeof(this.parentRows[i][this.parentHeaders[h] ]) === 'undefined')
                    { 
                      this.parentRows[i][this.parentHeaders[h] ] = "**";
                      object[this.parentHeaders[h]] = this.parentRows[i][this.parentHeaders[h] ];
                    }
                    else
                      object[this.uniqueParentColumnIdentifier + this.parentHeaders[h]] = this.parentRows[i][this.parentHeaders[h] ];
                }

                //For all Child Headers , rename them as uniqueChildColumnIdentifier<Column Name> where uniqueChildColumnIdentifier is a configured prefix.
              // Example : If Child Column Name is Desc__c and uniqueChildColumnIdentifier = 'C." , then rename it to as C.Desc__c 
              // This helps in easy identification of the column names during Chunking .

                for( var j in this.childHeaders ){

                  
                  //Changing below , to avoid templating logic with regards to Child Records 
                    if( typeof(this.childRows[i]) !== 'undefined'){
                     //   if( this.childHeaders[j] === 'Id' ){
                            object[this.uniqueChildColumnIdentifier + this.childHeaders[j]] = this.childRows[i][this.childHeaders[j] ];
                    //    }  else {
                    //        object[this.childHeaders[j]] = this.childRows[i][this.childHeaders[j] ];
                    //    }
                    }
                }

                chunkedData.push(object);
            }
            this.chunkData = chunkedData;
        } else {
            this.chunkData = getRowsData(cSheet, myrange, 1);
        }

        //set chunk information values
        this.totalDataSize =  this.chunkData.length;

        var chunkSize = 0;
        var chunkCol = [];

        if( this.chunkByCriteria == 'BOTH' ){
            var cbv = this.chunkByValue.split(",");
            this.chunkRanges = findChunksByBoth(this.chunkData, cbv,this.childHeaders,this.uniqueParentColumnIdentifier,this.uniqueChildColumnIdentifier);
        } else if ( this.chunkByCriteria == 'NUM') {
            if( this.chunkByValue > 0 ){
                chunkSize = parseInt( this.chunkByValue );
                chunkCol = null;
                this.chunkRanges = findChunksByRow(this.chunkData, chunkSize);
            }
        } else if ( this.chunkByCriteria == 'COL') {
            chunkCol = this.chunkByValue;
            chunkSize = 0;
            this.chunkRanges = findChunksByCol(this.chunkData, chunkCol,this.childHeaders,this.uniqueParentColumnIdentifier,this.uniqueChildColumnIdentifier);
        }


    }; //end getChunkInformation




    /*
     *  method prepares he data for print
     */
    this.displayData = function(){
        var myPreparedData = [];

        //clear sheet
        var GSheet = SpreadsheetApp.getActiveSpreadsheet();
        var resultSheetName = this.chunkSheetName;
        var cSheet = GSheet.setActiveSheet(GSheet.getSheetByName(resultSheetName));

        // if this is the first set of chunks clear sheet
        var mystartRange = cSheet.getRange(2,1);
        var myendRange = cSheet.getRange(cSheet.getLastRow(), cSheet.getLastColumn());
        var myrange = cSheet.getRange(mystartRange.getA1Notation() + ":" + myendRange.getA1Notation());
        myrange.clear();

        if( this.chunkByCriteria == 'BOTH' ){
            for( var r in this.chunkRanges ){
                for( var x =0; x < this.chunkRanges[r].length; x++){
                    myPreparedData = [];
                    for( var y =0; y < this.chunkRanges[r][x].length; y++){
                        myPreparedData.push( this.chunkData[ this.chunkRanges[r][x][y] ] );
                    }
                    this.displayChunk(myPreparedData);
                }
            }

        } else {

            for( var r in this.chunkRanges ){
                myPreparedData = [];
                for( var x =0; x < this.chunkRanges[r].length; x++){
                    var myDataValue = this.chunkData[ this.chunkRanges[r][x] ] ;
                    myPreparedData.push(myDataValue);
                }
                this.displayChunk(myPreparedData);
            }
        }
    };


    /*  method displays the chunked data to the Google Spreadsheet
     *  params:   startRow - which row to start showing data
     *              endRow - which row to end data and add blank line
     *          dataTo append to Google Spreadsheet
     */
    this.displayChunk = function(chunkToDisplay){
        var GSheet = SpreadsheetApp.getActiveSpreadsheet();

        var resultSheetName = this.chunkSheetName;
        var cSheet = GSheet.setActiveSheet(GSheet.getSheetByName(resultSheetName));
        var displaychunksz = 0;
        var headercol = null;


        // for each data row, match up headers and append that row to spreadsheet      
        for( var dta in chunkToDisplay ){
            if( dta == 0 ){
                displaychunksz = 0;
            }

            var myrow = [];

            if( typeof( chunkToDisplay[dta]) != 'undefined' ){
                var d = JSON.parse(JSON.stringify(chunkToDisplay[dta]));                

                if( this.childHeaders.length > 0 ){
                    this.chunkHeaders = [];

                    //For all Parent Headers , rename them as uniqueParentColumnIdentifier<Column Name> where uniqueParentColumnIdentifier is a configured Prefix
                    // Example : If Parent Column Name is Desc__c and uniqueParentColumnIdentifier = 'M." , then rename it to as M.Desc__c 
                    // This helps in easy identification of the column names during Chunking .  
                  
                    for( var p in this.parentHeaders ){
                      
                        if(this.parentHeaders[p] == "**")
                         this.chunkHeaders.push(this.parentHeaders[p] );
                        else
                         this.chunkHeaders.push( this.uniqueParentColumnIdentifier + this.parentHeaders[p] );
                    }

                      //For all Child Headers , rename them as uniqueChildColumnIdentifier<Column Name> where uniqueChildColumnIdentifier is a configured prefix.
                      // Example : If Child Column Name is Desc__c and uniqueChildColumnIdentifier = 'C." , then rename it to as C.Desc__c 
                      // This helps in easy identification of the column names during Chunking .
                    
                    for( var c in this.childHeaders ){
                   //     if( this.childHeaders[c] === 'Id' )
                   //         this.childHeaders[c] = "c_Id";
                   //              this.childHeaders[c] = "c_" + this.childHeaders[c];       
                        this.chunkHeaders.push( this.uniqueChildColumnIdentifier + this.childHeaders[c] );
                    }

                }
               

                for( var h in this.chunkHeaders ){

                    if( typeof(d[ this.chunkHeaders[h] ]) === 'undefined' ){

                        // only push blanks for missing values  not for missing headers
                        if( this.chunkHeaders[h] !== "**" ){
                            myrow.push(" " );
                            break;
                        }

                    } else {

                        //Changing below for Child Templating Logic Problem
                        if( this.chunkHeaders[h]  === this.templateTotalColumn){
                            headercol = h;
                        }
                        
                        myrow.push( d[ this.chunkHeaders[h] ] );
                    }

                }
                cSheet.appendRow( myrow );
                displaychunksz = displaychunksz + 1;

            }
        }


        //determine blank row
        var lastRowWithValue = cSheet.getLastRow();

        var blankRowRange = cSheet.getRange(cSheet.getRange(lastRowWithValue+1, 1).getA1Notation()+ ":" + cSheet.getRange(lastRowWithValue+1, cSheet.getLastColumn()).getA1Notation());
        
        if( this.templateTotalColumn != null ){            
            var blankRow = cSheet.getRange(lastRowWithValue+1, 1);
            //parse formula
            var startrow = lastRowWithValue+1 - displaychunksz;
            var chunkformula = null;

            var formulaColumn = cSheet.getRange(lastRowWithValue+1, headercol);  //find col for header with same value

            if(this.templateLb1.indexOf("=") != -1)
            {                                
                var LbWithoutColon = this.templateLb1.trim().split(";");
                var i=0;
                for(;i < LbWithoutColon.length;i=i+1)
                {
                    var tokens = LbWithoutColon[i].split("=");
                  
                    var uniqueToken = generateUniqueColumnName(tokens[0].trim(),this.childHeaders,this.uniqueParentColumnIdentifier,this.uniqueChildColumnIdentifier);
                    var k = this.headersMap2[uniqueToken]; 
                                        
                    if(k)
                    {
                        var co = cSheet.getRange(lastRowWithValue+1,parseInt(k)+1);  //find col for header with same value
                        var inp = tokens[1].indexOf("+");
                        if(inp != -1)
                            co.setValue(tokens[1].substring(0,inp).replace(/'/g,"") + chunkCtr);
                        else
                            co.setValue(tokens[1].replace(/'/g,""));
                    }
                }
            }
            else
            {
                formulaColumn.setValue(generateUniqueColumnName(this.templateLb1,this.childHeaders,this.uniqueParentColumnIdentifier,this.uniqueChildColumnIdentifier));
            }
            chunkformula1 = parseFormula(this.templateFormula1, startrow, lastRowWithValue, headercol, cSheet);
            formulaColumn.offset(0,1).setFormula(chunkformula1);
            blankRowRange.setBackground('orange');


            if( this.templateLb2 != 'null' ){

                if(this.templateLb2.indexOf("=") != -1)
                {
                    var LbWithoutColon = this.templateLb2.trim().split(";");
                    var i=0;
                    for(;i < LbWithoutColon.length;i=i+1)
                    {
                        var tokens = LbWithoutColon[i].split("=");
                        
                        var uniqueToken = generateUniqueColumnName(tokens[0].trim(),this.childHeaders,this.uniqueParentColumnIdentifier,this.uniqueChildColumnIdentifier);
                        var k = this.headersMap2[uniqueToken]; 
                        if(k)
                        {
                            var co = cSheet.getRange(lastRowWithValue+2,parseInt(k)+1);  //find col for header with same value

                            var inp = tokens[1].indexOf("+");
                            if(inp != -1)
                                co.setValue(tokens[1].substring(0,inp).replace(/'/g,"") + chunkCtr);
                            else
                                co.setValue(tokens[1].replace(/'/g,""));
                        }
                    }
                }
                else
                    formulaColumn.offset(1,0).setValue(generateUniqueColumnName(this.templateLb2,this.childHeaders,this.uniqueParentColumnIdentifier,this.uniqueChildColumnIdentifier));

                chunkformula2 = parseFormula(this.templateFormula2, startrow, lastRowWithValue, headercol, cSheet);
                formulaColumn.offset(1,1).setFormula(chunkformula2);
                blankRowRange.offset(1,0).setBackground('#0C8EFF');

            }
        }
        else {
            blankRowRange.mergeAcross().setValue("*").setBackground('black');
        }

        chunkCtr+=1;

    }; //end displayChunk

}   //end Chunker()


// Updated for 1-Level of Nesting of Formula 
function parseFormula(formula, rangeStart, rangeEnd, hColumn, csheet){


    // For ABS ( SUM () ) 
    var formulaName = formula.split("(");
    var additionalformula = formula.split(")");

    var formulaMinusEqual = formulaName[0].split("=");    // Will contain ABS 
    var innerFormula = formulaName[1];                    // Will contain SUM 

    var innerAdditionalFormula = additionalformula[1];
    var outerAdditionalFormula = additionalformula[2];

    var chunkRange1 = csheet.getRange(rangeStart, hColumn).offset(0,1);
    var chunkRange2 = csheet.getRange(rangeEnd, hColumn).offset(0,1);

    var chunkRange01 = chunkRange1.getA1Notation();
    var chunkRange02 = chunkRange2.getA1Notation();


    var formula;
    if(formula.indexOf("(") != formula.lastIndexOf("("))
    {
        formula =  innerFormula + "(" + chunkRange01 + ":" + chunkRange02 + ")"   + innerAdditionalFormula;
        formula = formulaMinusEqual[1] + "(" + formula + ")" + outerAdditionalFormula;
    }
    else
    {
        formula = formulaMinusEqual[1] + "(" + chunkRange01 + ":" + chunkRange02 + ")"   + innerAdditionalFormula;
    }


    return formula;
}



/* method called when Push Chunks menu option is selected
 *  pushChunks populates a Chunker obj from latest chunker data, pushes data to SF,
 *  and shows results based on latest chunker configs
 */
function pushChunks(){
}//end push





/*
 *  method to divide data into chunks by Row
 *  Param: chunkData - data to be chunked; chunkSize - num of rows to group
 *  Returns an array of ranges
 */
function findChunksByRow (chunkData, chunkSize){
    var myData = chunkData;
    var myRanges = [];
    var start = 0;


    //while there is still data rows
    while( start < myData.length ){
        var tempRanges = [];

        //break data into chunks of chunkSize
        for( var i = 0; i < chunkSize; i++){
            if( start+i < myData.length ){
                tempRanges.push( start+i);
            } else {
                break;
            }
        }
        start += i; //reset start to next item
        myRanges.push( tempRanges );
    }
    return myRanges;
} //end findChunksByRow



/*
 *  method to divide data into chunks by COL
 *   Param: chunkData - the data to be chunked;  chunkCol - column to match ; childHeaders - List of Child Headers if any
 *          uniqParentId - Id that will be used to prefix chunkCol ; uniqChildId - Id that will be used to prefix chunkCol
 *  Note : chunkCol will now contain the identifier for the column to be used i.e. Whether it belongs to Parent or Child . 
 *         M.<ColName> refers to the Parent . 
 *         C.<ColName> refers to the Child.
 *  
 */
function findChunksByCol(chunkData, chunkCol,childHeaders,uniqParentId,uniqChildId){
    var myData = chunkData;
    var myRanges = [];
    var start = 0 ;
    var rangeIndices = [];
    var tvalues = chunkCol.split(",");

    //Check if Parent/Child Identifiers are provided in the chunkCol or not .
  
    //Case 1 : childHeaders is not null and not empty - Indicates Child is present . 
    //If chunkCol has no identifiers , then we default them to Child i.e. prefix them with uniqChildId(passed as parameter)
    if(childHeaders && childHeaders.length > 0)
    {
       for(var k=0;k<tvalues.length;k++)
       {
         if(tvalues[k].indexOf(uniqChildId)==-1 && tvalues[k].indexOf(uniqParentId)==-1)
          tvalues[k] = uniqChildId + tvalues[k];
       }
    }
  
  //Case 2 : childHeaders is null and empty - Indicates only Parent present .
  //If chunkCol has no identifiers , then we default them to Parent i.e. prefix them with uniqParentId(passed as parameter)
  if(!childHeaders)
  {
    for(var k=0;k<tvalues.length;k++)
       {
         if(tvalues[k].indexOf(uniqParentId)==-1)
          tvalues[k] = uniqParentId + tvalues[k];
       }
  }
          
    var currentVal1 = myData[0][ tvalues[0] ] || null;
    var currentVal2 = myData[0][ tvalues[1] ] || null;
    var currentVal3 = myData[0][ tvalues[2] ] || null;

    if( currentVal1 == null ){
        var e = new Event("No Data Passed in to Chunk By Col!");
        e.triggerEvent();
        return;
    }

    for( var i=0; i < myData.length; i++){
        var sameValue = true;

        //check if same column values
        if( currentVal1 != null && currentVal1 != myData[i][ tvalues[0] ] ){
            //check next level
            sameValue = false;
        }

        if( currentVal2 != null && currentVal2 != myData[i][ tvalues[1] ] ){
            //check another level down
            sameValue = false;
        }

        if( currentVal3 != null && currentVal3 != myData[i][ tvalues[2] ] ){
            sameValue = false;
        }
        if( !sameValue ) {
            //if values dont' match saved previous range reset value
            myRanges.push(rangeIndices);
            rangeIndices = [];
            currentVal1 = myData[i][ tvalues[0] ] || null;
            currentVal2 = myData[i][ tvalues[1] ] || null;
            currentVal3 = myData[i][ tvalues[2] ] || null;
            i--;  //don't increment so that it will store this index
        } else {
            rangeIndices.push(i);
        }// end if save

    } // end for

    //push remaining rangeIndices on to ranges
    if( rangeIndices.length > 0){
        myRanges.push( rangeIndices);
        rangeIndices = [];
    }
    return myRanges;
} //end findChunksByCol



/*
 *  method to divide data into chunks by Row Size and Column
 *   Param: chunkData - the data to be chunked; cbv - Array containing in order( chunkSize- number of rows to group, chunkCol - column to match) ;childHeaders - List of Child Headers if any
 *          uniqParentId - Id that will be used to prefix chunkCol ; uniqChildId - Id that will be used to prefix chunkCol
 *   Returns a pattern of the chunk data  indices do not map to data directly following the pattern returned
 */
function findChunksByBoth(myData, cbv,childHeaders,uniqParentId,uniqChildId) {
    var chunkSize = cbv[0];
    var columnsOnly= [];
    for(var x=1; x < cbv.length; x++){
        columnsOnly.push( cbv[x] ) ;
    }
    var chunkCol = columnsOnly.join(",");

    var myRanges = [];

    var myColRanges = findChunksByCol(myData, chunkCol,childHeaders,uniqParentId,uniqChildId);

    var myRowRanges = [];
    var chunkRanges = [];


    for( var d in myColRanges){

        var tmpRanges = [];

        if( myColRanges[d].length <= chunkSize){
            tmpRanges.push(myColRanges[d]);
        } else {

            //break down ranges further but keep index
            var p =0;
            var tmpArr = [];
            for( var l =0; l < myColRanges[d].length; l++){
                if( p < chunkSize){
                    tmpArr.push(myColRanges[d][l]);
                    p++;
                } else {
                    tmpRanges.push(tmpArr);
                    tmpArr = [];
                    p=0;
                    l--; //allow the current index to get processed
                }
            }
            if( tmpArr.length > 0){
                tmpRanges.push(tmpArr);
            }
        }
        myRanges.push(tmpRanges);

    }

    return myRanges;

}

/*
 *   This method will generate an Associative Array for the Template Labels
 */

function createHashMap(chunkHeaders)
{    
    var map = new Object();
    for( var h in chunkHeaders )
    {
        map[chunkHeaders[h]] = h;
    }
    
    return map;

};

/*
 *   This method will generate an Associative Array for the Template Labels i.e. Child Labels will be relabeled as C.<Label Name>
 *   Example : If Parent and Child both have one Column Name as Desc__c , then Parent one will be referred to as Desc__c while the Child One will be referred to as C.Desc__c. 
 *   This helps during Chunking, where the Column Names are utilized .
 *   Parameters : pH -> refers to the Column Names for the Parent
 *                cH -> refers to the Column Names for the Child.
 *                uniqParentId - Id that will be used to prefix Parent chunkCol ; uniqChildId - Id that will be used to Child prefix chunkCol 
 *   Example : If uniqParentId = "M." and uniqChildId = "C." , then Parent Column 'Desc__c' will now become M.Desc__c and similarly Child Column 'Desc__c' will now become 'C.Desc__c'
 */

function createUniqueColumnsNames(pH,cH,uniqParentId,uniqChildId)
{
    var map = new Object();
    for( var h in pH )
    {
        if(pH[h] != "**")
         map[uniqParentId + pH[h]] = h;
        else
         map[pH[h]] = h; 
    }
 
    for(var c in cH)
    {
      map[uniqChildId + cH[c]] = parseInt(c) + pH.length;
    }
          
    return map;

};


/*
 *   This method will generate a Unique Column Name with Parent/Child Identifiers 
 *   Example : If Parent and Child both have one Column Name as Desc__c , then Parent one will be referred to as Desc__c while the Child One will be referred to as C.Desc__c. 
 *   This helps during Chunking, where the Column Names are utilized .
 *   Parameters : colName -> column Name
 *                childHeaders -> List of Child Headers if any 
 *                uniqParentId - Id that will be used to prefix Parent chunkCol ; uniqChildId - Id that will be used to Child prefix chunkCol 
 *   Example : If uniqParentId = "M." and uniqChildId = "C." , then Parent Column 'Desc__c' will now become M.Desc__c and similarly Child Column 'Desc__c' will now become 'C.Desc__c'
 */

function generateUniqueColumnName(colName,childHeaders,uniqParentId,uniqChildId)
{
  var toReturn = colName;
  // Case 1 : If ChildHeaders is not null and not empty -> Child is Present
  if(childHeaders && childHeaders.length > 0)
  {
    if(colName.indexOf(uniqParentId) == -1 && colName.indexOf(uniqChildId) == -1)
     toReturn = (uniqChildId + colName);
  }
  // Case 2 : If only Parent is present
  else
  {
    if(colName.indexOf(uniqParentId) == -1)
     toReturn =(uniqParentId + colName);
  }
  
  return toReturn;
};
