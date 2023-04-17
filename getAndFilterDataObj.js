class DataProcessing {
    constructor(ssId) {
        this.ssId = ssId;
        this.dataArr = [];
        this.filteredRowsObj = {};
        this.filteredRowsArr = [];
        this.joinedDataArr;
    };

    // Other methods ...

    /**
     * Get, filter, and modify data from the Google Spreadsheet based on specified conditions.
     * @param {number} sId - The sheet ID to retrieve data from.
     * @param {number} headerStart - The row index (0-based) of the header row in the data.
     * @param {string} objKey - The column name that is used as the obj key.
     * @param {object} valuesToFilterRow - An object with header values as keys and arrays of row values to filter as values.
     * @param {object} keepFilterRows - An object with header values as keys and booleans as values indicating whether to keep or remove the row.
     * @param {Array} valuesToFilterCol - An array of header values to find the indexes for.
     * @param {boolean} keepFilterCols - A flag indicating whether to keep or remove the columns specified by valuesToFilterCol.
     * @return {object} - An object containing the header and data.
     */
    getAndFilterDataObj(sId, headerStart, objKey, valuesToFilterRow, keepFilterRows, valuesToFilterCol, keepFilterCols) {
        const data = this.makeBatchGetByDataFilter(sId);
        let { header, rows } = this.getHeaderAndRows(data, headerStart);

        let headerIndexLookup = this.createHeaderIndexLookup(header);
        if (!header.includes(objKey)) {
            console.warn("objKey does not exist in header.");
            return;
        }

        let colIndexes = this.getColumnIndexes(header, valuesToFilterCol);
        let { rowIndexes, headerError } = this.getRowIndexes(header, valuesToFilterRow);
        
        if (headerError) {
            console.warn("Header Error. Please check your filter names and/or header start row.");
            return;
        }

        this.removeColumns(header, colIndexes, keepFilterCols);

        let filteredRowsObj = {};
        rows.forEach((row, indexOfRow) => {
            if (this.shouldKeepRow(row, rowIndexes, keepFilterRows, headerIndexLookup)) {
                this.removeColumns(row, colIndexes, keepFilterCols);
                const rowObj = this.convertToObj(header, row);
                const tempKey = rowObj[objKey];

                if (filteredRowsObj[tempKey] === undefined) {
                    filteredRowsObj[tempKey] = rowObj;
                } else {
                    Object.keys(rowObj).forEach((element) => {
                        if (!Array.isArray(filteredRowsObj[tempKey][element])) {
                            const tempValue = filteredRowsObj[tempKey][element];
                            filteredRowsObj[tempKey][element] = [tempValue, rowObj[element]];
                        } else {
                            filteredRowsObj[tempKey][element].push(rowObj[element]);
                        }
                    });
                }

                filteredRowsObj[tempKey].rowNum = headerStart + indexOfRow + 1;
            }
        });

        this.filteredRowsObj = filteredRowsObj;

        return {
            header: header,
            data: filteredRowsObj
        };
    };
}