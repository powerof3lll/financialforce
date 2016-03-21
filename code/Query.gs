//******************************** Query Object *******************************
//a Query object to parse the selected query into relevant parts for pushing and pulling

function Query(query, sheetName, status, outputmode){
    this.qColumns = [];
    this.qParentColumns = [];
    this.qChildColumns = [];
    this.qParentObj=null;
    this.qChildObj=null;
    this.qMode= outputmode;
    this.qStatus = typeof(status) == 'undefined' ? null : status;
    this.resultSheetName = typeof(sheetName) == 'undefined' ? null : sheetName;
    this.rawQuery = typeof(query) == 'undefined' ? null : query;
    this.qExtParentObj=null;
    this.qExtChildObj=null;
    this.qChildFK=null;
    this.chunkSize=null;
    this.chunkCols=null;

    //template variables
    this.templatetotalColumn = null;
    this.templateLb1 = null;
    this.templateFormula1= null;
    this.templateLb2 = null;
    this.templateFormula2= null;


    //encapsulation   gets/sets
    Object.defineProperties(this, "qcolumns", {
        get: function() { return this.qColumns; },
        set: function(v) { this.qColumns = v; }
    });

    Object.defineProperties(this, "qparentColumns", {
        get: function() { return this.qParentColumns; },
        set: function(v) { this.qParentColumns = v; }
    });

    Object.defineProperties(this, "qchildColumns", {
        get: function() { return this.qChildColumns; },
        set: function(v) { this.qChildColumns = v;  }
    });


    Object.defineProperties(this, "qparentObj", {
        get: function() { return this.qParentObj; },
        set: function(v) { this.qParentObj = v; }
    });


    Object.defineProperties(this, "qchildObj", {
        get: function() { return this.qChildObj; },
        set: function(v) { this.qChildObj = v;  }
    });


    Object.defineProperties(this, "qMode", {
        get: function() { return this.qMode; },
        set: function(v) { this.qMode = v; }
    });


    Object.defineProperties(this, "qExtParentObj", {
        get: function() { return this.qExtParentObj; },
        set: function(v) { this.qExtParentObj = v;  }
    });
    Object.defineProperties(this, "qExtChildObj", {
        get: function() { return this.qExtChildObj; },
        set: function(v) { this.qExtChildObj = v; }
    });
    Object.defineProperties(this, "qChildFK", {
        get: function() { return this.qChildFK; },
        set: function(v) { this.qChildFK = v; }
    });

    Object.defineProperties(this, "resultSheetName", {
        get: function() { return this.resultSheetName; },
        set: function(v) { this.resultSheetName = v;}
    });

    Object.defineProperties(this, "qStatus", {
        get: function() { return this.qStatus; },
        set: function(v) { this.qStatus = v; }
    });
    Object.defineProperties(this, "qQuery", {
        get: function() { return this.rawQuery; },
    });

    Object.defineProperties(this, "chunkSize", {
        get: function() { return this.chunkSize; },
        set: function(v) { this.chunkSize = v; }
    });

    Object.defineProperties(this, "chunkCols", {
        get: function() { return this.chunkCols; },
        set: function(v) { this.chunkCols = v; }
    });



    Object.defineProperties(this, "templateTotalColumn", {
        get: function() { return this.templatetotalColumn; },
        set: function(v) { this.templatetotalColumn = v; }
    });

    Object.defineProperties(this, "templateLabel1", {
        get: function() { return this.templateLb1; },
        set: function(v) { this.templateLb1 = v; }
    });
    Object.defineProperties(this, "templateFormula1", {
        get: function() { return this.templateFormula1; },
        set: function(v) { this.templateFormula1 = v; }
    });
    Object.defineProperties(this, "templateLabel2", {
        get: function() { return this.templateLb2; },
        set: function(v) { this.templateLb2 = v; }
    });
    Object.defineProperties(this, "templateFormula2", {
        get: function() { return this.templateFormula2; },
        set: function(v) { this.templateFormula2 = v; }
    });


//method to parse the query object from the query passed in
    this.setParams = function(query){
        var headers = [];
        var pheaders = [];
        var cheaders = [];
        var myQregexp = query.split("(");
        var innerquery;


        if(myQregexp.length > 1 ){
            for(var i = 0; i < myQregexp.length; i++){
                innerquery = myQregexp[i].split(")");
            }

            var parentHeaders = this.getColumnHeaders(myQregexp[0], 'Parent', innerquery[1]);
            var subHeaders = this.getColumnHeaders(innerquery[0],'Child');

            for(var i=0; i < parentHeaders.length; i++){
                if(parentHeaders[i].length > 0){
                    headers.push(parentHeaders[i].trim().replace(",", ""));
                    pheaders.push(parentHeaders[i].trim().replace(",", ""));
                }
            }

            if( (this.qMode == 'Tabular' || this.qMode == 'Chunker') && subHeaders.length > 0 )
                headers.push("**");  //add a space column between Parent and Child headers

            for(var i=0; i < subHeaders.length; i++){
                if(subHeaders[i].length > 0){
                    headers.push(subHeaders[i].trim().replace(",", ""));
                    cheaders.push(subHeaders[i].trim().replace(",", ""));
                }
            }
        } else {
            headers = this.getColumnHeaders(query, 'Parent');
            pheaders = headers;
        }

        this.qcolumns = headers;
        this.qparentColumns = pheaders;
        this.qchildColumns = cheaders;

    };



    //method to return the column headers from the query passed in
    this.getColumnHeaders =  function(query, obj, fromObj){
        var headers = [];
        var myQuery = query.split(" ");

        for(var i = 0; i < myQuery.length; i++){

            if(myQuery[i].toLowerCase() == "from"){
                var j = i + 1;  //the next word after FROM

                if(obj == 'Parent'){
                    this.qparentObj = myQuery[j];
                } else if(obj == 'Child'){
                    this.qchildObj = myQuery[j];
                }
                //if reach the word FROM then its the end of the header list so break out loop
                break;

            } else{


                if(myQuery[i].toLowerCase() != "select"){
                    var nameExt = myQuery[i].trim().split(".");

                    if(nameExt.length > 1){
                        var fieldName = nameExt[1];
                        headers.push(fieldName.trim().replace(",", ""));
                    } else {
                        headers.push(myQuery[i].trim().replace(",", ""));
                    }

                }

            }
        }
        //find from in Parent query
        if(fromObj != null){
            var myFQ = fromObj.trim().split(" ");
            this.qparentObj = myFQ[1];
        }

        return headers;
    }


}
