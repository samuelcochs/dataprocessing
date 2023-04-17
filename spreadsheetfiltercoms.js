/**
DataProcessing class for filtering, sorting, and modifying data from a Google Spreadsheet.
*/

class DataProcessing {

    /**
    Create a DataProcessing instance.
    @param {string} ssId - The ID of the Google Spreadsheet to process.
    */
    constructor(ssId) {
        this.ssId = ssId;
        this.dataArr = [];
        this.filteredRowsObj = {};
        this.filteredRowsArr = [];
        this.joinedDataArr;
    }

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
}