class DataProcessing {
    constructor(ssId) {
        this.ssId = ssId;
        this.dataArr = [];
        this.filteredRowsObj = {};
        this.filteredRowsArr = [];
        this.joinedDataArr;
    };
    
    /**
     * Converts two arrays to one object
     * @param header An array
     * @param values An array
     */
    convertToObj(header, values) {
        return Object.assign(...header.map((k, i) => ({ [k]: values[i] })));
    };

    /**
    Create a batch get request using data filters.
    @param {number} sId - The sheet ID to retrieve data from.
    @return {object} data - The batch get response containing the data.
    */
   makeBatchGetByDataFilter(sId) {
        const request = {
            spreadsheetId: this.ssId,
            dataFilters: [
                {
                    gridRange: {
                        sheetId: sId,
                    },
                },
            ],
        };
        const data = Sheets.Spreadsheets.Values.batchGetByDataFilter(request, this.ssId);
        return data;
    }

    //filter out certain values based on condition returns obj rows
    getAndFilterDataObj(sId, headerStart, objKey, valuesToFilterRow, keepFilterRows, valuesToFilterCol, keepFilterCols) {
        //get data from the sheets api
        const data = this.makeBatchGetByDataFilter(sId);
        //modifies header and row positions based on headerStart arg
        let header, rows;
        if (headerStart != 1) {
            header = data.valueRanges[0].valueRange.values[headerStart - 1];
            rows = data.valueRanges[0].valueRange.values.slice(headerStart);
        } else {
            header = data.valueRanges[0].valueRange.values[0];
            rows = data.valueRanges[0].valueRange.values.slice(1);
        };

        let headerIndexLookup = {}
        header.forEach((element) => headerIndexLookup[header.indexOf(element)] = element)

        if (header.find(element => element == objKey) == undefined) {
            console.warn("objKey does not exist in header.")
            return;
        }
        //converts column names into indexes
        var colIndexes = valuesToFilterCol.map((element) => parseInt(header.indexOf(element)))
        const colIndexesSet = new Set(colIndexes)
        //converts object keys and array into a new object with indexes as the keys and a SET as the array
        // {"Key":[Arr]} => {0:new Set()}
        const rowIndexes = {}

        Object.keys(valuesToFilterRow).forEach((element) => { rowIndexes[header.indexOf(element)] = new Set(valuesToFilterRow[element]) })
        //modifies the header to remove columns not wanted
        if (keepFilterCols) {
            colIndexes = Object.keys(header).filter((element) => !colIndexesSet.has(parseInt(element)))
            colIndexes.forEach((element, index) => { if (index == 0) { header.splice(element, 1) } else { header.splice(element - index, 1) } })
        } else {
            colIndexes.forEach((element, index) => { if (index == 0) { header.splice(element, 1) } else { header.splice(element - index, 1) } })
        }

        //loops through and filters the rows based on there value and/or there column
        var filteredRows = {}
        let indexOfRow = headerStart + 1;
        for (const row of rows) {
            let keepRow = [];
            //
            if (keepFilterRows == true) {
                Logger.log("Please update the keepFilterRows parameter.")
                return;
            }
            if (Object.keys(valuesToFilterRow).length == 0) {
                keepRow = true;
            } else {
                Object.keys(rowIndexes).forEach((element) => {
                    while (element >= row.length) {
                        row.push("");
                    }
                    if (keepFilterRows[headerIndexLookup[element]] == true) {
                        keepRow.push(rowIndexes[element].has(row[element]));
                    } else {
                        keepRow.push(!rowIndexes[element].has(row[element]));
                    }
                })
                if (keepRow.includes(false)) {
                    keepRow = false;
                } else {
                    keepRow = true;
                }
            }
            if (keepRow) {
                if (keepFilterCols) {
                    colIndexes.forEach((element, index) => { if (index == 0) { row.splice(element, 1) } else { row.splice(element - index, 1) } })
                } else {
                    colIndexes.forEach((element, index) => { if (index == 0) { row.splice(element, 1) } else { row.splice(element - index, 1) } })
                }
                var rowObj = this.convertToObj(header, row);
                var tempKey = rowObj[objKey];
                //delete rowObj[objKey];
                if (filteredRows[tempKey] == null) {
                    filteredRows[tempKey] = rowObj;
                } else {
                    let tempKeys = Object.keys(rowObj);
                    tempKeys.forEach((element) => {
                        if (Array.isArray(filteredRows[tempKey][element]) == false) {
                            let tempValue = filteredRows[tempKey][element];
                            filteredRows[tempKey][element] = [];
                            filteredRows[tempKey][element].push(tempValue)
                            filteredRows[tempKey][element].push(rowObj[element]);
                        } else {
                            filteredRows[tempKey][element].push(rowObj[element]);
                        }
                    })
                }
                //add row index
                filteredRows[tempKey]["rowNum"] = indexOfRow;
            }
            indexOfRow += 1;
        }
        this.filteredRowsObj = filteredRows;
        const returnObj = {
            "header": header,
            "data": filteredRows
        };
        return returnObj;
    };
}