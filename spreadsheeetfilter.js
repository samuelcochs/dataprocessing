class DataProcessing {
    constructor(ssId) {
      this.ssId = ssId;
      this.dataArr = [];
      this.filteredRowsObj = {};
      this.filteredRowsArr = [];
      this.joinedDataArr;
    };
    makebatchGetByDataFilter(sId) {
        const request =
        {
          'spreadsheetId': this.ssId,
          "dataFilters": [
            {
              "gridRange": {
                "sheetId": sId
              }
            }
          ]
        };
        const data = Sheets.Spreadsheets.Values.batchGetByDataFilter(request, this.ssId);
        return data;
    };

    //filter out certain values based on condition returns arr rows
    getAndFilterDataArr(sId, headerStart, valuesToFilterRow, keepFilterRows, valuesToFilterCol, keepFilterCols,addRowNum=false) {
        //get data from the sheets api
        const data = this.makebatchGetByDataFilter(sId);
        //modifies header and row positions based on headerStart arg
        let header, rows;
        if(headerStart != 1){
        header = data.valueRanges[0].valueRange.values[headerStart-1];
        rows = data.valueRanges[0].valueRange.values.slice(headerStart);
        }else{
        header = data.valueRanges[0].valueRange.values[0];
        rows = data.valueRanges[0].valueRange.values.slice(1);
        };

        let headerIndexLookup = {}
        header.forEach((element) => headerIndexLookup[header.indexOf(element)] = element)

        //converts column names into indexes
        var colIndexes = valuesToFilterCol.map((element) => parseInt(header.indexOf(element)))
        const colIndexesSet = new Set(colIndexes)

        //converts object keys and array into a new object with indexes as the keys and a SET as the array
        const rowIndexes = {}
        let headerError = false;
        Object.keys(valuesToFilterRow).forEach((element) => {
        if(!header.includes(element)){
            headerError = true;
        }
        rowIndexes[header.indexOf(element)] = new Set(valuesToFilterRow[element]);
        })

        //check for a header error
        if(headerError){
        console.warn("Header Error. Please check your filter names and/or header start row.")
        return;
        }

        //modifies the header to remove columns not wanted
        if(keepFilterCols){
        colIndexes = Object.keys(header).filter((element) => !colIndexesSet.has(parseInt(element)))
        colIndexes.forEach((element,index) => {if(index == 0){header.splice(element,1)}else{header.splice(element-index,1)}})
        }else{
        colIndexes.forEach((element,index) => {if(index == 0){header.splice(element,1)}else{header.splice(element-index,1)}})
        }

        //loops through and filters the rows based on there value and/or there column
        var filteredRows = [];
        let indexOfRow = headerStart+1;
        for(const row of rows){
        let keepRow = [];
        //
        if(keepFilterRows == true){
            Logger.log("Please update the keepFilterRows parameter.")
            return;
        }
        if(Object.keys(valuesToFilterRow).length == 0){
            keepRow = true;
        }else{
            Object.keys(rowIndexes).forEach((element) => {
            while(element >= row.length){
                row.push("");
            }
            if(keepFilterRows[headerIndexLookup[element]] == true){
                let rowValueToSearch = row[element];
                if(rowValueToSearch != "" && rowValueToSearch != null){
                rowValueToSearch = rowValueToSearch.replace(/\\/g,"");
                }
                keepRow.push(rowIndexes[element].has(rowValueToSearch));
            }else{
                let rowValueToSearch;
                if(row[element] == null){
                rowValueToSearch = "";
                }else{
                rowValueToSearch = row[element];
                }
                if(rowValueToSearch != "" && rowValueToSearch != null){
                rowValueToSearch = rowValueToSearch.replace(/\\/g,"");
                }
                keepRow.push(!rowIndexes[element].has(rowValueToSearch));
            }
            })
            if(keepRow.includes(false)){
            keepRow = false;
            
            }else{
            keepRow = true;
            
            }
        }

        if(keepRow){
            if(keepFilterCols){
            colIndexes.forEach((element,index) => {if(index == 0){row.splice(element,1)}else{row.splice(element-index,1)}})
            }else{
            colIndexes.forEach((element,index) => {if(index == 0){row.splice(element,1)}else{row.splice(element-index,1)}})
            }
            if(addRowNum){
            const inHeader = header.includes("RowNum");
            if(!inHeader){
                while(header.length > row.length){
                row.push("");
                }
                header.push("RowNum");
                valuesToFilterCol.push("RowNum");
            }else{
                while(header.length-1 > row.length){
                row.push("");
                }
            }
            row.push(indexOfRow);
            }else{
            while(header.length > row.length){
                row.push("");
            }
            }
            filteredRows.push(row);
        }

        indexOfRow += 1;
        }

        //sort columns based on given column orientation
        if(valuesToFilterCol.length != 0 && keepFilterCols == true){
        const desiredHeaderIndexes = valuesToFilterCol.map((element) => header.indexOf(element))
        header = desiredHeaderIndexes.map((element) => header[element])
        filteredRows = filteredRows.map((element) => desiredHeaderIndexes.map((el) => element[el]))
        }

        let finalHeaderIndexes = {}
        header.forEach((element) => finalHeaderIndexes[header.indexOf(element)] = element)

        //concats both the header and the filtered data
        const filteredRowsWithHeader = [header].concat(filteredRows)
        //modified the orginal constructor to update with the new data
        this.filteredRowsArr = filteredRowsWithHeader;
        const returnObj = {
        "header": header,
        "data": filteredRows,
        "combined": filteredRowsWithHeader,
        "indexes": finalHeaderIndexes
        };
        return returnObj;
    }
}