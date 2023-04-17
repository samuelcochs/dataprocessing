function exportDataProcessing(ssId) {
    return new DataProcessing(ssId);
}

class DataProcessing {
    constructor(ssId) {
        this.ssId = ssId;
        this.dataArr = [];
        this.filteredRowsObj = {};
        this.filteredRowsArr = [];
        this.joinedDataArr;
    };

    /**
     * Adds the ability to change the orginal Spreadsheet Id
     * @param ssId the Spreadsheet Id
     */
    setSSID(ssId) {
        this.ssId = ssId;
    };

    /**
     * Adds the abilty to change the filteredRowsArr object in the constructor
     * @param data A 2d array
     */
    setFilteredDataArr(data) {
        this.filteredRowsArr = data;
    };

    getFilteredDataArr() {
        return this.filteredRowsArr;
    }

    /**
     * Adds the abilty to change the dataArr object in the constructor
     * @param data A 2d array
     */
    setDataArr(data) {
        this.dataArr = data;
    };

    getDataArr() {
        return this.dataArr;
    };

    /**
     * Adds the abilty to change the joinedDataArr array in the constructor
     * @param data A 2d array
     */
    setJoinedDataArr(data) {
        this.joinedDataArr = data;
    };

    getJoinedDataArr() {
        return this.joinedDataArr;
    }

    /**
     * Adds the ability to refresh formulas on a spreadsheet.
     * @param sId A sheet Id
     * @param startRow The row to start on
     */
    refreshFormulas(sId, startRow = 1) {
        const getRequest =
        {
            'spreadsheetId': this.ssId,
            "dataFilters": [
                {
                    "gridRange": {
                        "sheetId": sId
                    }
                },
            ],
            "valueRenderOption": "FORMULA"
        };
        const values = Sheets.Spreadsheets.Values.batchGetByDataFilter(getRequest, this.ssId);
        const parsedValues = values.valueRanges[0].valueRange.values
        const postRequest =
        {
            "valueInputOption": "USER_ENTERED",
            "data": [
                {
                    "dataFilter": {
                        "gridRange": {
                            "sheetId": sId,
                            "startRowIndex": startRow - 1
                        }
                    },
                    "values": parsedValues,
                    "majorDimension": "ROWS"
                }
            ]
        };
        Sheets.Spreadsheets.Values.batchUpdateByDataFilter(postRequest, this.ssId);
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
        return Sheets.Spreadsheets.Values.batchGetByDataFilter(request, this.ssId);
    }

    /**
     * Converts two arrays to one object
     * @param header An array
     * @param values An array
     */
    convertToObj(header, values) {
        return Object.assign(...header.map((k, i) => ({ [k]: values[i] })));
    };

    /**
     * Returns the sheet id's and title's in a 2d array
     */
    getNamesAndIdsOfSheets() {
        const sheets = Sheets.Spreadsheets.get(this.ssId).sheets;
        var namesAndIds = []
        for (const sheet of sheets) {
            var tempSheetId = sheet.properties.sheetId;
            var tempSheetTitle = sheet.properties.title;
            namesAndIds.push([tempSheetId, tempSheetTitle]);
        };
        return namesAndIds;
    };

    /**
     * Returns an object with the header and the values of the sheet in arrays
     * @param sId A sheet Id
     * @param headerStart The row that the data header is on
     */
    getDataSheetArr(sId, headerStart) {
        const data = this.makeBatchGetByDataFilter(sId);
        let header;
        let values;
        if (headerStart == 1) {
            header = data.valueRanges[0].valueRange.values[0];
            values = data.valueRanges[0].valueRange.values.slice(1);
        } else {
            header = data.valueRanges[0].valueRange.values[headerStart - 1];
            values = data.valueRanges[0].valueRange.values.slice(headerStart);
        };
        const indexes = {};
        header.forEach(columnName => { indexes[columnName] = header.indexOf(columnName) });
        const rowsWithHeader = [header].concat(values);
        const returnObj = {
            "header": header,
            "data": values,
            "combined": rowsWithHeader,
            "indexes": indexes
        };
        this.dataArr = rowsWithHeader;
        return returnObj;
    };

    /**
     * Returns an object with the designated key as the key for each object within the main object
     * @param sId A sheet Id
     * @param objKey The column name to use as the object key
     * @param headerStart The row that the data header is on
     */
    getDataSheetObj(sId, objKey, headerStart) {
        const data = this.makeBatchGetByDataFilter(sId);
        let header;
        let values;
        if (headerStart == 1) {
            header = data.valueRanges[0].valueRange.values[0];
            values = data.valueRanges[0].valueRange.values.slice(1);
        } else {
            header = data.valueRanges[0].valueRange.values[headerStart - 1];
            values = data.valueRanges[0].valueRange.values.slice(headerStart);
        };
        const returnObj = {};
        values.forEach((element, index) => { let tempObj = this.convertToObj(header, element); tempObj["rowNum"] = index + headerStart + 1; var tempKey = tempObj[objKey]; returnObj[tempKey] = tempObj; })
        return returnObj;
    };

    /**
     * Returns an object with all the sheets from the spreadsheet. With the sheet Name as the key and the data as the values.
     */
    getAllDataArr() {
        const sheets = Sheets.Spreadsheets.get(this.ssId).sheets;
        const dataOfSheetsArr = {};
        for (const sheet of sheets) {
            //var tempSheetId = sheet.properties.sheetId;
            var tempSheetTitle = sheet.properties.title;
            dataOfSheetsArr[tempSheetTitle] = this.getDataSheetArr(tempSheetId);
        };
        return dataOfSheetsArr;
    };

    //gets all the sheets and processes each sheet through getDataSheetObj() if a columnKey is given else it proccess through getDataSheetArr()
    /**
     * Returns an object with all the sheets from the spreadsheet. With the sheet Id as the key and the data as the values.
     */
    getAllDataObj(columnKeys, headerStartRows) {
        const sheets = Sheets.Spreadsheets.get(this.ssId).sheets;
        const dataOfSheetsObj = {};
        for (const sheet of sheets) {
            var tempSheetId = sheet.properties.sheetId;
            var tempSheetTitle = sheet.properties.title;
            if (columnKeys[tempSheetTitle] != null) {
                dataOfSheetsObj[tempSheetTitle] = this.getDataSheetObj(tempSheetId, columnKeys[tempSheetTitle], headerStartRows[tempSheetTitle]);
            } else {
                dataOfSheetsObj[tempSheetTitle] = this.getDataSheetArr(tempSheetId);
            };
        };
        return dataOfSheetsObj;
    };

    //returns an object with a count of each unique value
    getColumnUniqueCount(sId, headerStart, objKey, valuesToFilterRow, keepFilterRows, valuesToFilterCol, keepFilterCols, columnName) {

        this.getAndFilterDataObj(sId, headerStart, objKey, valuesToFilterRow, keepFilterRows, valuesToFilterCol, keepFilterCols);

        const data = this.filteredRowsObj;

        const dataKeys = Object.keys(data);

        var countObj = {};

        for (const dataKey of dataKeys) {

            const tempDataPoint = data[dataKey][columnName];

            if (countObj[tempDataPoint] == null) {

                countObj[tempDataPoint] = 1;

            } else {

                countObj[tempDataPoint] += 1;

            };
        };
        return countObj;
    };

    /**
    Retrieve the header row and data rows from the data.
    @param {object} data - The data object containing the value ranges.
    @param {number} headerStart - The row index (0-based) of the header row in the data.
    @return {object} - An object containing the header row and data rows.
    */
    getHeaderAndRows(data, headerStart) {
        let header, rows;
        if (headerStart !== 1) {
            header = data.valueRanges[0].valueRange.values[headerStart - 1];
            rows = data.valueRanges[0].valueRange.values.slice(headerStart);
        } else {
            header = data.valueRanges[0].valueRange.values[0];
            rows = data.valueRanges[0].valueRange.values.slice(1);
        }
        return { header, rows };
    }

    /**
    Create a lookup object for header indexes.
    @param {Array} header - The header row array.
    @return {object} headerIndexLookup - The lookup object with header indexes as keys and header values as values.
    */
    createHeaderIndexLookup(header) {
        let headerIndexLookup = {};
        header.forEach((element) => (headerIndexLookup[header.indexOf(element)] = element));
        return headerIndexLookup;
    }

    /**
    Get column indexes for the given header values.
    @param {Array} header - The header row array.
    @param {Array} valuesToFilterCol - An array of header values to find the indexes for.
    @return {Array} - An array of column indexes.
    */
    getColumnIndexes(header, valuesToFilterCol) {
        return valuesToFilterCol.map((element) => parseInt(header.indexOf(element)));
    }

    /**
    Get row indexes for the given header and row values to filter.
    @param {Array} header - The header row array.
    @param {object} valuesToFilterRow - An object with header values as keys and arrays of row values to filter as values.
    @return {object} - An object containing rowIndexes and headerError (boolean).
    */
    getRowIndexes(header, valuesToFilterRow) {
        let rowIndexes = {};
        let headerError = false;
        Object.keys(valuesToFilterRow).forEach((element) => {
            if (!header.includes(element)) {
                headerError = true;
            }
            rowIndexes[header.indexOf(element)] = new Set(valuesToFilterRow[element]);
        });
        return { rowIndexes, headerError };
    }

    /**
    Remove columns from the header array based on column indexes.
    @param {Array} header - The header row array.
    @param {Array} colIndexes - An array of column indexes to remove or keep.
    @param {boolean} keepFilterCols - A flag indicating whether to keep or remove the columns specified by colIndexes.
    */
    removeColumns(header, colIndexes, keepFilterCols) {
        const colIndexesSet = new Set(colIndexes);
        if (keepFilterCols) {
            colIndexes = Object.keys(header).filter((element) => !colIndexesSet.has(parseInt(element)));
        }
        colIndexes.forEach((element, index) => {
            if (index === 0) {
                header.splice(element, 1);
            } else {
                header.splice(element - index, 1);
            }
        });
    }

    /**
    Determine if a row should be kept based on the row values and rowIndexes.
    @param {Array} row - The row to be evaluated.
    @param {object} rowIndexes - An object containing row indexes and row values to filter.
    @param {object} keepFilterRows - An object with header values as keys and booleans as values indicating whether to keep or remove the row.
    @param {object} headerIndexLookup - The lookup object with header indexes as keys and header values as values.
    @return {boolean} - A boolean value indicating whether the row should be kept.
    */
    shouldKeepRow(row, rowIndexes, keepFilterRows, headerIndexLookup) {
        let keepRow = [];
        if (Object.keys(rowIndexes).length === 0) {
            return true;
        } else {
            Object.keys(rowIndexes).forEach((element) => {
                let rowValueToSearch = row[element] ? row[element].replace(/\\/g, "") : "";
                keepRow.push(keepFilterRows[headerIndexLookup[element]] ? rowIndexes[element].has(rowValueToSearch) : !rowIndexes[element].has(rowValueToSearch));
            });
            return !keepRow.includes(false);
        }
    }

    /**
    Ensure a row has the same length as the target length by adding empty strings.
    @param {Array} row - The row to be modified.
    @param {number} targetLength - The target length for the row.
    */
    ensureRowLength(row, targetLength) {
        while (row.length < targetLength) {
            row.push("");
        }
    }

    /**
    Filter and modify rows based on specified conditions.
    @param {Array} rows - An array of data rows.
    @param {Array} header - The header row array.
    @param {object} rowIndexes - An object containing row indexes and row values to filter.
    @param {object} keepFilterRows - An object with header values as keys and booleans as values indicating whether to keep or remove the row.
    @param {Array} colIndexes - An array of column indexes to remove or keep.
    @param {boolean} keepFilterCols - A flag indicating whether to keep or remove the columns specified by colIndexes.
    @param {boolean} addRowNum - A flag indicating whether to add a row number to the row.
    @param {object} headerIndexLookup - The lookup object with header indexes as keys and header values as values.
    @param {number} headerStart - The row index (0-based) of the header row in the data.
    @return {Array} - An array of filtered and modified rows.
    */
    filterAndModifyRows(rows, header, rowIndexes, keepFilterRows, colIndexes, keepFilterCols, addRowNum, headerIndexLookup, headerStart) {
        let filteredRows = [];
        let indexOfRow = headerStart + 1;
        for (const row of rows) {
            if (this.shouldKeepRow(row, rowIndexes, keepFilterRows, headerIndexLookup)) {
                this.removeColumns(row, colIndexes, keepFilterCols);

                if (addRowNum) {
                    const inHeader = header.includes("RowNum");
                    if (!inHeader) {
                        while (header.length > row.length) {
                            row.push("");
                        }
                        header.push("RowNum");
                        valuesToFilterCol.push("RowNum");
                    } else {
                        while (header.length - 1 > row.length) {
                            row.push("");
                        }
                    }
                    row.push(indexOfRow);
                } else {
                    while (header.length > row.length) {
                        row.push("");
                    }
                }
                filteredRows.push(row);
            }
            indexOfRow += 1;
        }
        return filteredRows;
    }

    /**
    Sort columns based on specified conditions.
    @param {Array} header - The header row array.
    @param {Array} filteredRows - An array of filtered rows.
    @param {Array} valuesToFilterCol - An array of header values to find the indexes for.
    @param {boolean} keepFilterCols - A flag indicating whether to keep or remove the columns specified by valuesToFilterCol.
    @return {object} - An object containing the sorted header and sorted filtered rows.
    */
    sortColumns(header, filteredRows, valuesToFilterCol, keepFilterCols) {
        if (valuesToFilterCol.length !== 0 && keepFilterCols === true) {
            const desiredHeaderIndexes = valuesToFilterCol.map((element) => header.indexOf(element));
            header = desiredHeaderIndexes.map((element) => header[element]);
            filteredRows = filteredRows.map((element) => desiredHeaderIndexes.map((el) => element[el]));
        }
        return { header, filteredRows };
    }

    /**
    Get, filter, and modify data from the Google Spreadsheet based on specified conditions.
    @param {number} sId - The sheet ID to retrieve data from.
    @param {number} headerStart - The row index (0-based) of the header row in the data.
    @param {object} valuesToFilterRow - An object with header values as keys and arrays of row values to filter as values.
    @param {object} keepFilterRows - An object with header values as keys and booleans as values indicating whether to keep or remove the row.
    @param {Array} valuesToFilterCol - An array of header values to find the indexes for.
    @param {boolean} keepFilterCols - A flag indicating whether to keep or remove the columns specified by valuesToFilterCol.
    @param {boolean} [addRowNum=false] - A flag indicating whether to add a row number to the row.
    @return {object} - An object containing the header, data, combined (header and data), and indexes.
    */
    getAndFilterDataArr(sId, headerStart, valuesToFilterRow, keepFilterRows, valuesToFilterCol, keepFilterCols, addRowNum = false) {
        const data = this.makeBatchGetByDataFilter(sId);
        let { header, rows } = this.getHeaderAndRows(data, headerStart);
        let headerIndexLookup = this.createHeaderIndexLookup(header);
        let colIndexes = this.getColumnIndexes(header, valuesToFilterCol);
        let { rowIndexes, headerError } = this.getRowIndexes(header, valuesToFilterRow);

        if (headerError) {
            console.warn("Header Error. Please check your filter names and/or header start row.");
            return;
        }

        this.removeColumns(header, colIndexes, keepFilterCols);
        let filteredRows = this.filterAndModifyRows(rows, header, rowIndexes, keepFilterRows, colIndexes, keepFilterCols, addRowNum, headerIndexLookup, headerStart);
        let { header: sortedHeader, filteredRows: sortedFilteredRows } = this.sortColumns(header, filteredRows, valuesToFilterCol, keepFilterCols);

        let finalHeaderIndexes = this.createHeaderIndexLookup(sortedHeader);
        const filteredRowsWithHeader = [sortedHeader].concat(sortedFilteredRows);
        this.filteredRowsArr = filteredRowsWithHeader;

        const returnObj = {
            header: sortedHeader,
            data: sortedFilteredRows,
            combined: filteredRowsWithHeader,
            indexes: finalHeaderIndexes,
        };

        return returnObj;
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

    //getAndFilterDataObj must be called before this function
    getFilteredColumnUniqueCount(columnName) {
        let data;
        data = this.filteredRowsObj;
        const dataKeys = Object.keys(data);
        var countObj = {};
        for (const dataKey of dataKeys) {
            const tempDataPoint = data[dataKey][columnName];
            if (countObj[tempDataPoint] == null) {
                countObj[tempDataPoint] = 1;
            } else {
                countObj[tempDataPoint] += 1;
            };
        };
        return countObj;
    };

    addSheet(title = "Sheet1", ssId = this.ssId) {
        const request =
        {
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": title,
                            "gridProperties": {
                                "rowCount": 50,
                                "columnCount": 25
                            },
                            "tabColor": {
                                "red": 1.0,
                                "green": 0.3,
                                "blue": 0.4
                            }
                        }
                    }
                }
            ]
        };
        const response = Sheets.Spreadsheets.batchUpdate(request, ssId);
        const newSheetId = response["replies"][0]["addSheet"]["properties"]["sheetId"];
        return newSheetId;
    };

    deleteSheetRange(rowStartIndex, rowQuantDelete, colStartIndex, colQuantDelete, ssId, sId) {
        const colEndIndex = colStartIndex + colQuantDelete;
        const rowEndIndex = rowStartIndex + rowQuantDelete;

        const request =
        {
            "requests": [
                {
                    "deleteDimension": {
                        "range": {
                            "sheetId": sId,
                            "dimension": "COLUMNS",
                            "startIndex": colStartIndex,
                            "endIndex": colEndIndex
                        }
                    }
                },
                {
                    "deleteDimension": {
                        "range": {
                            "sheetId": sId,
                            "dimension": "ROWS",
                            "startIndex": rowStartIndex,
                            "endIndex": rowEndIndex
                        }
                    }
                },
            ]
        };
        Sheets.Spreadsheets.batchUpdate(request, ssId);
    };

    updateSheetRange(ssId, sId, rowStartIndex, rowQuantAdd, colStartIndex, colQuantAdd) {
        const colEndIndex = colStartIndex + colQuantAdd;
        const rowEndIndex = rowStartIndex + rowQuantAdd;

        const request =
        {
            "requests": [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": sId,
                            "dimension": "COLUMNS",
                            "startIndex": colStartIndex,
                            "endIndex": colEndIndex
                        },
                        "inheritFromBefore": true
                    }
                },
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": sId,
                            "dimension": "ROWS",
                            "startIndex": rowStartIndex,
                            "endIndex": rowEndIndex
                        },
                        "inheritFromBefore": true
                    }
                },
            ]
        };
        Sheets.Spreadsheets.batchUpdate(request, ssId);
    };

    getSpreadsheetRange(ssId, sId) {
        const sheets = Sheets.Spreadsheets.get(String(ssId)).sheets;
        let columnIndex;
        let rowIndex;
        for (const sheet of sheets) {
            const sheetProperties = sheet.properties;
            const sheetId = sheetProperties.sheetId;
            if (sheetId == sId) {
                columnIndex = sheetProperties.gridProperties.columnCount;
                rowIndex = sheetProperties.gridProperties.rowCount;
            };
        };
        const returnedObj = {
            "columnIndex": columnIndex,
            "rowIndex": rowIndex
        };
        return returnedObj;
    };

    //typeOfData = filtered or unfiltered
    pushData(ssId, sId, typeOfData, extraCols = 0, startRow = 1, extraRows = 0) {
        let values;
        if (typeOfData == "filtered") {
            values = this.filteredRowsArr;
        } else if (typeOfData == "unfiltered") {
            values = this.dataArr;
        } else if (typeOfData == "joined") {
            values = this.joinedDataArr;
        }

        const currentRange = this.getSpreadsheetRange(ssId, sId);

        let columnDiff = 0;
        let rowDiff = 0;

        let rowStartIndex = currentRange.rowIndex;
        let colStartIndex = currentRange.columnIndex;

        let rowEndIndex = values.length + (startRow - 1) + extraRows;
        let columnEndIndex = values[0].length + extraCols;

        if (colStartIndex <= columnEndIndex) {

            if (rowStartIndex <= rowEndIndex) {

                columnDiff = columnEndIndex - colStartIndex;
                rowDiff = rowEndIndex - rowStartIndex;
                if (rowEndIndex == 1 & extraCols != 0) {
                    rowDiff = 50
                }

                this.updateSheetRange(ssId, sId, rowStartIndex, rowDiff, colStartIndex, columnDiff);

            } else if (rowStartIndex > rowEndIndex) {

                columnDiff = columnEndIndex - colStartIndex;
                rowDiff = rowStartIndex - rowEndIndex;
                //0 ,0 because it will never need to delete columns only add
                //starting at the end of the values because you need to delete everything after that number
                rowStartIndex = rowEndIndex;
                if (rowEndIndex == 1 & extraCols != 0) {
                    rowDiff = 50
                }
                this.deleteSheetRange(rowStartIndex, rowDiff, 0, 0, ssId, sId);

                this.updateSheetRange(ssId, sId, 1, 0, colStartIndex, columnDiff);
            }

        } else if (colStartIndex > columnEndIndex) {

            if (rowStartIndex > rowEndIndex) {

                columnDiff = colStartIndex - columnEndIndex;
                rowDiff = rowStartIndex - rowEndIndex;
                colStartIndex = columnEndIndex;
                rowStartIndex = rowEndIndex;
                if (rowEndIndex == 1 & extraCols != 0) {
                    rowDiff = 50
                }
                this.deleteSheetRange(rowStartIndex, rowDiff, colStartIndex, columnDiff, ssId, sId);

            } else if (rowStartIndex <= rowEndIndex) {

                columnDiff = colStartIndex - columnEndIndex;
                colStartIndex = columnEndIndex;
                this.deleteSheetRange(0, 0, colStartIndex, columnDiff, ssId, sId);
                columnDiff = 0;
                rowDiff = rowEndIndex - rowStartIndex;
                if (rowEndIndex == 1 & extraCols != 0) {
                    rowDiff = 50
                }
                this.updateSheetRange(ssId, sId, rowStartIndex, rowDiff, colStartIndex, columnDiff);

            }
        };

        let endRowIndexRequest;
        if (startRow != 1) {
            endRowIndexRequest = values.length + (startRow - 1);
        } else {
            endRowIndexRequest = values.length;
        }

        const clearRequest =
        {
            "dataFilters": [
                {
                    "gridRange": {
                        "sheetId": sId,
                        "startRowIndex": startRow - 1,
                        "endRowIndex": endRowIndexRequest,
                        "startColumnIndex": 0,
                        "endColumnIndex": values[0].length

                    }
                }
            ]
        }

        Sheets.Spreadsheets.Values.batchClearByDataFilter(clearRequest, ssId)

        const request =
        {

            "valueInputOption": "USER_ENTERED",
            "data": [
                {
                    "dataFilter": {
                        "gridRange": {
                            "sheetId": sId,
                            "startRowIndex": startRow - 1
                        }
                    },
                    "values": values,
                    "majorDimension": "ROWS"
                }
            ]
        };
        Sheets.Spreadsheets.Values.batchUpdateByDataFilter(request, ssId);

    };

    updateRow(ssId, sId, values, rowNum, colNum) {
        const request =
        {

            "valueInputOption": "USER_ENTERED",
            "data": [
                {
                    "dataFilter": {
                        "gridRange": {
                            "sheetId": sId,
                            "startRowIndex": rowNum - 1,
                            "startColumnIndex": colNum - 1
                        }
                    },
                    "values": values,
                    "majorDimension": "ROWS"
                }
            ]
        };
        Sheets.Spreadsheets.Values.batchUpdateByDataFilter(request, ssId);
    };

    updateRows(ssId, sId, updateRowsObj) {
        const request =
        {

            "valueInputOption": "USER_ENTERED",
            "data": [
            ]
        };
        const rowNums = Object.keys(updateRowsObj);
        rowNums.forEach(e => {
            const rowNum = e;
            const colNum = updateRowsObj[e]["colNum"];
            const values = updateRowsObj[e]["values"];

            let tempRequest = {
                "dataFilter": {
                    "gridRange": {
                        "sheetId": sId,
                        "startRowIndex": rowNum - 1,
                        "startColumnIndex": colNum - 1
                    }
                },
                "values": [values],
                "majorDimension": "ROWS"
            };
            request["data"].push(tempRequest);
        });
        Sheets.Spreadsheets.Values.batchUpdateByDataFilter(request, ssId);
    };

    cleanAndAddExtraRow(sId, rowsToAdd) {
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
        const data = Sheets.Spreadsheets.getByDataFilter(request, this.ssId);
        const actualColumnCount = data["sheets"][0]["properties"]["gridProperties"]["columnCount"];
        const actualRowCount = data["sheets"][0]["properties"]["gridProperties"]["rowCount"];

        const data1 = Sheets.Spreadsheets.Values.batchGetByDataFilter(request, this.ssId)
        const usedColumnCount = data1["valueRanges"][0]["valueRange"]["values"][0].length;
        const usedRowCount = data1["valueRanges"][0]["valueRange"]["values"].length + rowsToAdd;

        let columnDiff;
        let rowDiff;
        let rowStartIndex = usedRowCount - 1;
        let colStartIndex = usedColumnCount;
        if (actualColumnCount <= usedColumnCount) {
            if (actualRowCount <= usedRowCount) {
                columnDiff = usedColumnCount - actualColumnCount;
                rowDiff = usedRowCount - actualRowCount;
                this.updateSheetRange(rowStartIndex, rowDiff, colStartIndex, columnDiff, sId, this.ssId);
            } else if (actualRowCount > usedRowCount) {
                columnDiff = usedColumnCount - actualColumnCount;
                rowDiff = actualRowCount - usedRowCount;
                this.deleteSheetRange(rowStartIndex, rowDiff, colStartIndex, columnDiff, sId, this.ssId);
                this.updateSheetRange(rowStartIndex, rowDiff, colStartIndex, columnDiff, sId, this.ssId);
            };
        } else if (actualColumnCount > usedColumnCount) {
            if (actualRowCount > usedRowCount) {
                columnDiff = actualColumnCount - usedColumnCount;
                rowDiff = actualRowCount - usedRowCount;
                this.deleteSheetRange(rowStartIndex, rowDiff, colStartIndex, columnDiff, sId, this.ssId);
            } else if (actualRowCount <= usedRowCount) {
                columnDiff = actualColumnCount - usedColumnCount;
                rowDiff = usedRowCount - actualRowCount;
                this.deleteSheetRange(rowStartIndex, rowDiff, colStartIndex, columnDiff, sId, this.ssId);
                columnDiff = 0;
                this.updateSheetRange(rowStartIndex, rowDiff, colStartIndex, columnDiff, sId, this.ssId);
            }
        };
        const lastRow = usedRowCount;
        return lastRow;
    };

    cleanExtraRowsAndColumns(sId) {
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
        const data = Sheets.Spreadsheets.getByDataFilter(request, this.ssId);
        const actualColumnCount = data["sheets"][0]["properties"]["gridProperties"]["columnCount"];
        const actualRowCount = data["sheets"][0]["properties"]["gridProperties"]["rowCount"];

        const data1 = Sheets.Spreadsheets.Values.batchGetByDataFilter(request, this.ssId)
        const usedColumnCount = data1["valueRanges"][0]["valueRange"]["values"][0].length;
        const usedRowCount = data1["valueRanges"][0]["valueRange"]["values"].length;

        let columnDiff;
        let rowDiff;
        let rowStartIndex = usedRowCount;
        let colStartIndex = usedColumnCount;
        if (actualColumnCount <= usedColumnCount) {
            if (actualRowCount <= usedRowCount) {
                columnDiff = usedColumnCount - actualColumnCount;
                rowDiff = usedRowCount - actualRowCount;
                this.updateSheetRange(rowStartIndex, rowDiff, colStartIndex, columnDiff, sId, this.ssId);
            } else if (actualRowCount > usedRowCount) {
                columnDiff = usedColumnCount - actualColumnCount;
                rowDiff = actualRowCount - usedRowCount;
                this.deleteSheetRange(rowStartIndex, rowDiff, colStartIndex - 1, columnDiff, sId, this.ssId);
                rowDiff = 0;
                this.updateSheetRange(rowStartIndex, rowDiff, colStartIndex, columnDiff, sId, this.ssId);
            };
        } else if (actualColumnCount > usedColumnCount) {
            if (actualRowCount > usedRowCount) {
                columnDiff = actualColumnCount - usedColumnCount;
                rowDiff = actualRowCount - usedRowCount;
                this.deleteSheetRange(rowStartIndex, rowDiff, colStartIndex, columnDiff, sId, this.ssId);
            } else if (actualRowCount <= usedRowCount) {
                columnDiff = actualColumnCount - usedColumnCount;
                rowDiff = usedRowCount - actualRowCount;
                this.deleteSheetRange(rowStartIndex, rowDiff, colStartIndex, columnDiff, sId, this.ssId);
                columnDiff = 0;
                this.updateSheetRange(rowStartIndex, rowDiff, colStartIndex, columnDiff, sId, this.ssId);
            }
        };
    };
    /**
     * Return filtered or unfiltered data from a spreadsheet using the google sheets api.
     * @param  {String} sId  The google sheet id.
     * @param  {Integer} headerStart  The row that the header of your spreadsheet is on.
     * @param  {String} objKey  The column name that is used as the obj key.
     * @param  {Obj} valuesToFilterRow  An object where the key is the column name and the value/values are an array. EX. {"Lease Num":["LCA01011"]}
     * @param  {Obj} keepFilterRows  An object where the key is the column name and the value is a boolean. The boolean is used to determine whether or not to keep the values submitted in the prior argument.
     * @param  {Array} valuesToFilterCol  An array with all of the column names that you would like to keep or drop.
     * @param  {Boolean} keepFilterCols  A boolean indicating whether you would like to drop the columns or keeps them in the previous argument.
     * @param  {Boolean} addRowNum  A boolean indicating whether or not to add the rownum.
     * 
     * @return {Obj} Returns an object with 3 keys header, combined and data. 
     */
    getData(sId, headerStart, objKey = null, valuesToFilterRow = null, keepFilterRows, valuesToFilterCol, keepFilterCols, addRowNum = false) {
        let data;
        if (valuesToFilterRow == null && objKey == null) {
            data = this.getDataSheetArr(sId, headerStart)
        } else if (objKey == null) {
            data = this.getAndFilterDataArr(sId, headerStart, valuesToFilterRow, keepFilterRows, valuesToFilterCol, keepFilterCols, addRowNum)
        } else {
            //getAndFilterDataObj(sId, headerStart, objKey, valuesToFilterRow, keepFilterRows, valuesToFilterCol, keepFilterCols)
            data = this.getAndFilterDataObj(sId, headerStart, objKey, valuesToFilterRow, keepFilterRows, valuesToFilterCol, keepFilterCols)
        }
        return data;
    };

    getUniqueRowsByColumn(colsToFilterUnique, typeOfData) {
        let values;
        if (typeOfData == "filtered") {
            values = this.filteredRowsArr;
        } else if (typeOfData == "unfiltered") {
            values = this.dataArr;
        }
        let header, rows;
        header = values[0];
        rows = values.slice(1);
        //find the indexes of the columns to unique
        const colIndexes = colsToFilterUnique.map((element) => parseInt(header.indexOf(element)))
        const uniqueRowValues = {}

        colIndexes.forEach((element) => { uniqueRowValues[element] = new Set() })

        const filteredRows = rows.filter((element) => { let returned; colIndexes.forEach((el) => { if (uniqueRowValues[el].has(element[el])) { returned = false; } else { uniqueRowValues[el].add(element[el]); returned = true; } }); return returned; })

        if (typeOfData == "filtered") {
            const filteredRowsWithHeader = [header].concat(filteredRows)
            this.filteredRowsArr = filteredRowsWithHeader;
            return filteredRowsWithHeader;
        } else if (typeOfData == "unfiltered") {
            const rowsWithHeader = [header].concat(filteredRows)
            this.dataArr = rowsWithHeader;
            return rowsWithHeader;
        }
    }

    formatData(formatForCols, typeOfData) {
        let values;

        if (typeOfData == "filtered") {
            values = this.filteredRowsArr;
        } else if (typeOfData == "unfiltered") {
            values = this.dataArr;
        }
        let header = values[0];
        let rows = values.slice(1);
        const colsToFormat = Object.keys(formatForCols);
        const indexesOfColsToFormat = colsToFormat.map((element) => header.indexOf(element))
        function formatCol(value, formatRequired) {
            if (value != null && value != undefined && value != "" && value != " ") {
                if (formatRequired == "titlecase") {
                    return value.replace(/\w\S*/g, function (txt) { return txt.charAt(0).toUpperCase() + txt.substring(1).toLowerCase(); });
                } else if (formatRequired == "date") {
                    return new Date(value).toLocaleDateString("en-US");
                } else if (formatRequired == "uppercase") {
                    return value.toUpperCase();
                } else if (formatRequired == "lowercase") {
                    return value.toLowerCase();
                } else if (formatRequired == "zip") {
                    return value.replace('/', '');
                    /*
                    var rx = /[a-zA-Z]/g;
                    let result = value.match(rx); 
                    let length = value.length;
                    if(result == null){
                      if(length > 5){
                        return value.substring(0,5);
                      }else if(length < 5){
                        for(let i = 0; i < Math.max(0, 5 - length); i++){value = "0" + value}
                        return value;
                      }else{
                        return value;
                      }
                    }else{
                      return value;
                    }
                    */
                } else if (formatRequired == "phone") {
                    value = value.replaceAll(/[^\d]/g, '')
                    value = value.slice(0, 3) + '-' + value.slice(3, 6) + '-' + value.slice(6, 10)
                    return value;
                } else if (formatRequired == "duns") {
                    let duns = value.replaceAll(/[^\d]/g, '')
                    if (duns.length != 0) {
                        let diff = Math.max(0, 9 - duns.length);
                        for (let i = 0; i < diff; i++) { duns = "0" + duns }
                    }
                    return duns;
                } else {
                    return value;
                }
            } else {
                return value;
            }
        }
        rows = rows.map((element) => { indexesOfColsToFormat.forEach((el) => { const formatRequired = formatForCols[header[el]]; element.splice(el, 1, formatCol(element[el], formatRequired)) }); return element; })

        //rows = rows.map((element)=>{element.splice(28,29,"Test");return element;})
        if (typeOfData == "filtered") {
            const filteredRowsWithHeader = [header].concat(rows);
            this.filteredRowsArr = filteredRowsWithHeader;
            return filteredRowsWithHeader;
        } else if (typeOfData == "unfiltered") {
            const rowsWithHeader = [header].concat(rows);
            this.dataArr = rowsWithHeader;
            return rowsWithHeader;
        }
    }

    formatRange(sId, formatForCols, headerStart = 1) {
        /**
         * {1:{"type":"DATE"}}
         * 
         * 
         */
        const cols = Object.keys(formatForCols);
        for (let col of cols) {
            const format = formatForCols[col];

            col = Number(col);
            let endCol = col + 1;
            if (col != 0) {
                endCol = col;
                col -= 1
            }
            var resource = {
                "requests": [
                    {
                        "repeatCell": {
                            "cell": {
                                "userEnteredFormat": format
                            },
                            "range": {
                                "sheetId": sId,
                                "startRowIndex": headerStart == 1 ? 1 : headerStart,
                                "startColumnIndex": col,
                                "endColumnIndex": endCol // Specify the end to insert only one column of checkboxe
                            },
                            "fields": "userEnteredFormat"  // or "fields": "userEnteredFormat"
                        }
                    }]
            };
            Sheets.Spreadsheets.batchUpdate(resource, this.ssId);
        }
    };

    calculateFiscalYear(typeOfData, columnToCalculateWith) {
        let values;
        if (typeOfData == "filtered") {
            values = this.filteredRowsArr;
        } else if (typeOfData == "unfiltered") {
            values = this.dataArr;
        }
        const header = values[0].concat("Fiscal Year " + columnToCalculateWith);
        let data = values.slice(1);
        const indexOfColumnToCalculateWith = header.indexOf(columnToCalculateWith);
        data.map((element) => {
            let dateObj = new Date(element[indexOfColumnToCalculateWith]); let [month, year] = [dateObj.getMonth() + 1, dateObj.getFullYear()]; if (month > 9) { year = year + 1; }
            let fiscalYear = String(year); element.push(fiscalYear); return element
        })
        values = [header].concat(data);
        if (typeOfData == "filtered") {
            this.filteredRowsArr = values;
        } else if (typeOfData == "unfiltered") {
            this.dataArr = values;
        }
    };

    addZone(typeOfData, regionColumnName) {
        let values;

        if (typeOfData == "filtered") {
            values = this.filteredRowsArr;
        } else if (typeOfData == "unfiltered") {
            values = this.dataArr;
        }

        let header = values[0];
        let data = values.slice(1);

        let regionColumnIndex = header.indexOf(regionColumnName);

        function getZone(region) {
            let zoneTable = {
                "1": new Set(["1", "2", "3", "5", "11"]),
                "2": new Set(["4", "6", "7"]),
                "3": new Set(["8", "9", "10"])
            }
            for (const zone of Object.keys(zoneTable)) {
                let zoneFound = zoneTable[zone].has(region);
                if (zoneFound) {
                    return zone;
                }
            }
        }

        let newData = data.map((element) => { let region = element[regionColumnIndex]; let zone = getZone(region); return element.concat([zone]); })

        let newValues = [header.concat(["Zone"])].concat(newData);

        if (typeOfData == "filtered") {
            this.filteredRowsArr = newValues;
        } else if (typeOfData == "unfiltered") {
            this.dataArr = newValues;
        }
    };

    addRows(sId, rowsToAdd) {
        const sheetInfo = this.getSpreadsheetRange(this.ssId, sId);
        this.updateSheetRange(this.ssId, sId, sheetInfo["rowIndex"], rowsToAdd, 1, 0);
    };

    maintainNumOfRowsOnSheet(sId, numOfRows) {
        const sheetInfo = this.getSpreadsheetRange(this.ssId, sId);
        const rowIndex = sheetInfo["rowIndex"];
        if (rowIndex < numOfRows) {
            const rowsToAdd = numOfRows - rowIndex;
            this.updateSheetRange(this.ssId, sId, sheetInfo["rowIndex"], rowsToAdd, 1, 0);
        }
    };

    deleteRow(sId, rowNum) {
        const request =
        {
            "requests": [
                {
                    "deleteDimension": {
                        "range": {
                            "sheetId": sId,
                            "dimension": "ROWS",
                            "startIndex": rowNum - 1,
                            "endIndex": rowNum
                        }
                    }
                }
            ]
        };
        Sheets.Spreadsheets.batchUpdate(request, this.ssId);
    };

    deleteRows(sId, rows) {
        rows.sort(function (a, b) { return a - b }).forEach((e, i) => {
            let rowToDelete;
            if (i == 0) {
                rowToDelete = e;
            } else {
                rowToDelete = e - i;
            }
            this.deleteRow(sId, rowToDelete);
        });
    };

    pasteFormula(sId, rowNum, colNum, formula) {
        const request =
        {
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sId,
                            "startRowIndex": rowNum - 1,
                            "endRowIndex": rowNum,
                            "startColumnIndex": colNum - 1,
                            "endColumnIndex": colNum
                        },
                        "cell": {
                            "userEnteredValue": {
                                "formulaValue": formula
                            }
                        },
                        "fields": "userEnteredValue"
                    }
                }
            ]
        };
        Sheets.Spreadsheets.batchUpdate(request, this.ssId);
    };

    getHeaderForSheet() {
    };

    getSumOfColumn(sId, headerStart, valuesToFilterRow, keepFilterRows, valuesToFilterCol, keepFilterCols, typeOfData, columnName) {

        this.getData(sId, headerStart, null, valuesToFilterRow, keepFilterRows, valuesToFilterCol, keepFilterCols)

        let values;
        if (typeOfData == "filtered") {
            values = this.filteredRowsArr;
        } else if (typeOfData == "unfiltered") {
            values = this.dataArr;
        }

        let header = values[0];
        let data = values.slice(1);


        let columnNameIndex = header.indexOf(columnName);

        let sumOfColumn = 0;

        function cleanFormattingNumber(num) {
            return Number(num.replaceAll(",", "").replaceAll("$", ""));

        }

        data.forEach((element) => { sumOfColumn += cleanFormattingNumber(element[columnNameIndex]); })

        return sumOfColumn;
    };

    /**
     * Adds the ability to filter dates out of filtered data before it is pushed to the destination spreadsheet
     * @param typeOfData Filtered or Unfiltered.
     * @param columnToCalculateWith The column that the date is located in.
     * @param startDate The date to start filtering.
     * @param endDate The date to stop filtering.
     * @param inclusive Options of left, right, both.
     */
    filterDates(typeOfData, columnToCalculateWith, startDate, endDate, inclusive) {
        let values;
        if (typeOfData == "filtered") {
            values = this.filteredRowsArr;
        } else if (typeOfData == "unfiltered") {
            values = this.dataArr;
        }
        const header = values[0];
        let data = values.slice(1);
        const indexOfColumnToCalculateWith = header.indexOf(columnToCalculateWith);

        let filteredData = data.filter((element) => {
            let dateObj = new Date(element[indexOfColumnToCalculateWith] + " 00:00:00");
            if (inclusive == "left") {
                return dateObj <= endDate && dateObj > startDate;
            } else if (inclusive == "right") {
                return dateObj < endDate && dateObj >= startDate;
            } else if (inclusive == "both") {
                return dateObj <= endDate && dateObj >= startDate;
            } else {
                Logger.log("Please Format Correctly.")
            }
        })
        values = [header].concat(filteredData);
        if (typeOfData == "filtered") {
            this.filteredRowsArr = values;
        } else if (typeOfData == "unfiltered") {
            this.dataArr = values;
        }
        return values;
    };

    /**
     * 
     */
    getMultipleDataSources(destSpreadsheetSSId, sheetsToGet) {
        const spreadsheets = {
            "Zone 1 Tracker": {
                "ssId": "1jMago0V57teHOkSp1nH2GwZZPFFOKfE6dXZy2MDPi50",
                "Active": 0,
                "Lease File Reviews": 614920586,
                "CPI/Step Rent": 517062235,
                "Taxes": 1064812533,
                "Novation": 1912185094
            },
            "Zone 2 Tracker": {
                "ssId": "1pWIP3-p5lD9WPPhSuhY-SkzbOTDPq-CJ1ECITv-wuT4",
                "Active": 0,
                "Lease File Reviews": 614920586,
                "CPI/Step Rent": 517062235,
                "Taxes": 1064812533,
                "Novation": 1912185094
            },
            "Zone 3 Tracker": {
                "ssId": "16Opf8gm1chhac7DK3KLL6pYskpo-6or4VYCf_qOibxE",
                "Active": 0,
                "Lease File Reviews": 614920586,
                "CPI/Step Rent": 517062235,
                "Taxes": 1064812533,
                "Novation": 1912185094
            },
            "LCA Source": {
                "ssId": "16WfpdW1zRMrb70qx7PWlk1UZRTWyz8mvleTHemcPbqE",
                "Leases": 2097646666,
                "AFR": 1544606965,
                "Withhold-Tax Payments": 1337230005,
                "Lease Blocks": 478042477,
                "BA53 Prompt Pay Interest": 410147085,
                "Zero Occupancy": 1353417856,
                "SAM Contact Info": 196600821,
                "Lease Actions": 1972188303,
                "PRGX Claims": 436930511,
                "SLA": 212934285,
                "RET Cases": 1708324850
            },
            "National Archive": {
                "ssId": "1FXR6qpGKZI-XU4UvHqR06EBmWdOovJbTUgby2JUNJBc",
                "Active": 0,
                "Lease File Reviews": 614920586,
                "CPI/Step Rent": 517062235,
                "Taxes": 1064812533,
                "Novation": 1912185094
            },
            "National Novation Form Responses": {
                "ssId": "1YvOw-w8_tPXD03Wn37TMdQpNyV2eD7XuGG6pIz_i8lg",
                "Form Responses": 449888471
            }
        };

        const sheetToGetNames = Object.keys(sheetsToGet);
        sheetToGetNames.forEach((e) => {
            const spreadSheetName = e;
            const sheetName = sheetsToGet[spreadSheetName];
            const ssId = spreadsheets[spreadSheetName]["ssId"];
            const sId = spreadsheets[spreadSheetName][sheetName];
            this.setSSID(ssId);
            this.getData(sId, 1);
            const newSheetName = spreadSheetName + "-" + sheetName
            const newSheetId = this.addSheet(newSheetName, destSpreadsheetSSId);
            this.pushData(destSpreadsheetSSId, newSheetId, "unfiltered", 0, 1, 0);
        });
    };

    /**
     * Adds the abilty to change the filteredRowsArr object in the constructor
     * @param ssId A spreadsheet Id
     * @param sId A sheet Id
     * @param refreshedData An object where the desired rownum is the key and the updated row is the value
     * @param colNum The column to start on
     */
    refreshData(ssId, sId, refreshedData, colNum = 1) {
        const rowNums = Object.keys(refreshedData);
        rowNums.forEach((e) => {
            this.updateRow(ssId, sId, [refreshedData[e]], e, colNum);
        });
    };

    /**
     * Must be called using an arrow function, in order to pass through the parameters
     * Ex.
     * main.tryCatchWithRetries(() => myFunction("Test"),2)
     * @param func The function to try
     * @param retries The number of times to retry the function
     * @param delay The number of seconds to exponentially delay
     * @param maxDelay The max number of seconds
     */
    tryCatchWithRetries(func, retries = 2, delay = 4, maxDelay = 60) {
        try {
            return func();
        } catch (error) {
            if (retries == 0) {
                Logger.log(error);
                Logger.log("Max retries hit.")
                return error;
            } else {
                Logger.log(error);
                Logger.log("Trying Again.");
                //1.5^(delay/retries)
                const delayTime = Math.round(Math.pow(1.5, (delay / retries)));
                if (delayTime > maxDelay) {
                    delayTime = maxDelay;
                }
                Utilities.sleep(delayTime * 1000);
                return this.tryCatchWithRetries(func, retries - 1, delay, maxDelay);
            }
        }
    };

    /**
     * Convert a date to a date with the format of mm-dd-yyyy
     * @param typeOfData The type of data to be used
     * @param columnsToConvert An array of column names
     */
    convertDates(typeOfData, columnsToConvert) {
        let values;

        if (typeOfData == "filtered") {
            values = this.filteredRowsArr;
        } else if (typeOfData == "unfiltered") {
            values = this.dataArr;
        }

        const header = values[0];
        let data = values.slice(1);

        const indexesOfColumnsToConvert = [];
        columnsToConvert.forEach((e) => {
            const index = header.indexOf(e);
            indexesOfColumnsToConvert.push(index);
        });

        const newDate = (date) => {
            if (date != "") {
                date = new Date(date);
                return [date.getMonth() + 1, date.getDate(), date.getFullYear()]
                    .map(e => e < 10 ? `0${e}` : `${e}`).join('-');
            } else {
                return "";
            }
        };

        data = data.map((e) => {
            indexesOfColumnsToConvert.forEach(el => e[el] = newDate(e[el]))
            return e;
        });

        values = [header].concat(data);
        if (typeOfData == "filtered") {
            this.filteredRowsArr = values;
        } else if (typeOfData == "unfiltered") {
            this.dataArr = values;
        }
        return values;

    };

    /**
     * Add the functionality to merge two sheets together based on a common key.
     * @param arrData Data as an arr, normally a sheet of data
     * @param objData Data as an obj, normally a sheet of data that has been converted into an obj of obj's.
     * @param columnNameToJoin The common key/column between the two data sources/sheets.
     */
    mergeData(arrData, objData, columnNameToJoin) {
        const arrHeader = arrData["header"];
        let arrValues = arrData["data"];
        const columnNameToJoinIndex = arrHeader.indexOf(columnNameToJoin);
        const objHeader = [...objData["header"]];
        objHeader.shift();
        let objValues = objData["data"];

        arrValues = arrValues.map((e) => {
            const joinValue = e[columnNameToJoinIndex];
            const joinObj = objValues[joinValue];
            if (joinObj != null) {
                const joinArr = [];
                objHeader.forEach(el => joinArr.push(joinObj[el]));
                return e.concat(joinArr);
            } else {
                console.warn(`${joinValue} was not found`);
                const joinArr = [];
                objHeader.forEach(el => joinArr.push(""));
                return e.concat(joinArr);
                //return e;
            }
        });
        const finalHeader = arrHeader.concat(objHeader);
        const finalValues = [finalHeader].concat(arrValues);
        this.joinedDataArr = finalValues;
        return finalValues;
    };

    /**
     * Returns the last row of the current sheet plus one.
     * @param sId The sheet id of the sheet to get the last row on.
     */
    getNextRow(sId) {

        const nextRow = this.getData(sId, 1, null).data.length + 2;
        let sheet2Name = "";
        this.getNamesAndIdsOfSheets().forEach((e => { if (e[0] == sId) { sheet2Name = e[1] } }));
        const lastRow = SpreadsheetApp.openById(this.ssId).getSheetByName(sheet2Name).getMaxRows();
        if (nextRow > lastRow) {
            this.addRows(sId, 10);
        }
        return nextRow;
    };

    /**
     * Adds 
     * @param sId The sheet id
     * @param dataValidationForCols An object with the desired options.
     * EX.
     * { ColNum : {
          "startRow":1,
          "rule": {
            "condition": {
              "type": "ONE_OF_LIST",
              "values": [
                {
                  "userEnteredValue": "Not Started"
                }
              ]
            },
            "strict": true,
            "showCustomUi": true
          }
        }
      }
     */
    addDataValidation(sId, dataValidationForCols) {
        Object.keys(dataValidationForCols).forEach(colNum => {
            const rule = dataValidationForCols[colNum]["rule"];
            const startRow = dataValidationForCols[colNum]["startRow"];
            const request =
            {
                "requests": [
                    {
                        "setDataValidation": {
                            "range": {
                                "sheetId": sId,
                                "startColumnIndex": colNum - 1,
                                "endColumnIndex": colNum,
                                "startRowIndex": startRow
                            },
                            "rule": rule
                        }
                    }
                ]
            }
            Sheets.Spreadsheets.batchUpdate(request, this.ssId);
        });
    };
};












