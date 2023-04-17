class DataProcessing {
    constructor(ssId) {
      this.ssId = ssId;
      this.dataArr = [];
      this.filteredRowsObj = {};
      this.filteredRowsArr = [];
      this.joinedDataArr;
    }
  
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
  
    createHeaderIndexLookup(header) {
      let headerIndexLookup = {};
      header.forEach((element) => (headerIndexLookup[header.indexOf(element)] = element));
      return headerIndexLookup;
    }
  
    getColumnIndexes(header, valuesToFilterCol) {
      return valuesToFilterCol.map((element) => parseInt(header.indexOf(element)));
    }
  
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

    ensureRowLength(row, targetLength) {
        while (row.length < targetLength) {
          row.push("");
        }
      }
  
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
  
    sortColumns(header, filteredRows, valuesToFilterCol, keepFilterCols) {
      if (valuesToFilterCol.length !== 0 && keepFilterCols === true) {
        const desiredHeaderIndexes = valuesToFilterCol.map((element) => header.indexOf(element));
        header = desiredHeaderIndexes.map((element) => header[element]);
        filteredRows = filteredRows.map((element) => desiredHeaderIndexes.map((el) => element[el]));
      }
      return { header, filteredRows };
    }
  
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