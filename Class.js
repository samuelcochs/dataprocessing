/*\
  Author: Sam Ochs
  Contact: samuel.ochs@gsa.gov
  Created to centralize my data processing 
  Version 1.0.0
  //
  Change Log:
    Version 6 - Added sorting based on given column orientation
    Version 8 - Added ability to start pushData at a row other than 1
    Version 9 - Fixed deletion of extra rows
    Version 10 - Added help method
    Version 12 - FormatData added
    Version 14 - Added formatRange
    Version 16 - Fixed formatData
    Version 18 - Deleted deprecated methods
    0.1.6 , Version 19 - Added formatData(): phone and duns option
    Version 30 - V0.1.6 Beta, removed substr
    Version 33 - added calculateFiscalYear
    Version 34 - added clear of contents to pushdata
    Version 35 - added new zipcode formatting
    Version 36 - added addZone()
    Version 38 - added extraRows to pushData()
    Version 40 - modified filter data to handle, backslashes
    Version 41 - fixed blank row fitlering, sometimes empty cells on a sheet can be null
    Version 42 - added filterDates
    Version 43 - updates to filterDates
    Version 44 - updated getData to fix bug using getData as an object
    Version 45 - added getDataArr() & getFilteredDataArr()
    Version 46 - added date to formatData()
    Version 47 - set .toLocaleDateString() to .toLocaleDateString("en-US")
    Version 48 - updated setDataArr and getDataArr to correct camel case
    Version 49 - added addSheet()
    Version 50 - added refreshFormulas()
    Version 51 - added getMultipleDataSources()
    Version 52 - added refreshData()
    Version 54 - functionality to add rowNum to getData as arr
    Version 56 - added tryCatchWithRetries()
    Version 57 - added exponential delay to tryCatchWithRetries()
    Version 58 - added convertDates()
    Version 59 - added mergeData()
    Version 60 - added error handler if lookup value is not found for mergeData()
    Version 61 - added combined to getDataSheetArr to returned obj
    Version 62 - Updated getAndFilterDataObj to leave key in values
    Version 63 - Added RowNum to header when add row num is selected
    Version 65 - Added getNextRow() returns the last row plus one
    Version 66 - Fixed bug with getting rowNum when using getData(), rowNum was getting dropped when filtering on both rows and columns 
    Version 67 - updateSheetRange updated parameter order, getNextRow changed to add rows when no blank rows are left on the sheet, and updated zipcode format
    Version 68 - Modified getAndFilterDataObj() to allow for filtering of blank columns, the api drops blanks columns when there is no more data on the right
    Version 69 - Modified getNextRow() to use getMaxRows() instead of getLastRow()
    Version 70 - added addDataValidation()
    Version 71 - modified getAndFilterDataArr() to break if there was an error with the header filtering/header row start
    Version 72 - added addRows() allows, updated getAndFilterDataArr() to work with blanks cells and row num
    Version 73 - added headerStart to formatRange(), non breaking
    Version 74 - changed getAndFilterDataArr() and getAndFilterDataObj() to fill blank rows with "" instead of null to match what google returns
                  there was only a differnce depending on whether there was a last column of data
    Version 75 - Updated updateRow() to use "valueInputOption": "USER_ENTERED", instead of "valueInputOption": "RAW"
    Version 76 - Udpated refreshData() to allow changes to the column number. 
    Version 77 - Added method updateRows(), to allow for multiple non-synchronous rows to be updated at once.
    Version 78-80 - Multiple additions
    Version 81 - Added error thrower to updateRows()
    Version 82 - Fixed multiple errors
  //
  If you have any questions or ways to improve the code
  feel free to email me.
  Addition Log:

  Issue Log:
*/


function exportDataProcessing(ssId) {
  return new DataProcessing(ssId);
}

class DataProcessing {
  /** Start constructor
*/
  constructor(ssId) {
    this.ssId = ssId;
    this.dataArr = [];
    this.filteredDataObj = {};
    this.filteredDataArr = [];
  };
  /** 
*/

  /**----------------------------------------------------------------------- Starts Get and Set Methods ----------------------------------------------------------------------- 
*/
  /**
   * Adds the ability to change the orginal Spreadsheet Id.
   * @param {String} ssId - The Spreadsheet Id.
   */
  setSSId(ssId) {
    this.ssId = ssId;
  };

  /**
   * Adds the abilty to change the filteredDataArr Object in the constructor.
   * @param {Array} data - A 2d Array.
   */
  setFilteredDataArr(data) {
    this.filteredDataArr = data;
  };

  /**
   * Add the ability to get the filteredDataArr stored in the constructor.
   * @return {Array} - The Array from the constructor.
   */
  getFilteredDataArr() {
    return this.filteredDataArr;
  }

  /**
   * Adds the abilty to change the dataArr Object in the constructor.
   * @param {Array} data - A 2d Array.
   */
  setDataArr(data) {
    this.dataArr = data;
  };

  /**
   * Adds the ability to get the dataArr stored in the constructor.
   * @return {Array} - The Array from the constructor.
   */
  getDataArr() {
    return this.dataArr;
  };
  /**----------------------------------------------------------------------- Ends Get and Set Methods ----------------------------------------------------------------------- 
*/

  /** ---------------------------------------------------------------------- Starts Main Data Filtering Methods -----------------------------------------------------------------    
*/

  /**
   * Create a batch get request using data filters.
   * @param {Integer/String} ssId - The spread sheet ID to retrieve data from.
   * @param {Integer/String} sId - The sheet ID to retrieve data from.
   * @return {Array} data - The batch get response containing the data.
  */
  makeBatchGetByDataFilter(ssId, sId) {
    const request = {
      spreadsheetId: ssId,
      dataFilters: [
        {
          gridRange: {
            sheetId: sId,
          },
        },
      ],
    };
    return Sheets.Spreadsheets.Values.batchGetByDataFilter(request, ssId);
  };

  /**
   * Converts two Arrays to one Object.
   * @param {Array} header - An Array.
   * @param {Array} values - An Array.
   * @return {Object} - The combined Object.
   */
  convertToObj(header, values) {
    return Object.assign(...header.map((k, i) => ({ [k]: values[i] })));
  };

  /**
   * Sort row based on the logic in the dataFilters Object
   * @param {Object} dataFilters - The main dataFilters Object - See the getData method for more info
   */
  sortRows(dataFilters) {
    if (dataFilters.sortingConfig != null) {
      dataFilters.sortingHeader = [...dataFilters.header];
      this.removeColumns(dataFilters.sortingHeader, dataFilters.colIndexes, dataFilters.includeColValues);
      const sortingConfigObj = dataFilters.sortingConfig;
      let sortingConfigError = false;
      sortingConfigObj.forEach(e => { if (!dataFilters.header.includes(e.columnName)) { sortingConfigError = true; } });
      if (sortingConfigError) {
        throw new Error("sortingConfig column not found.")
      }
      dataFilters.rows.sort((a, b) => {
        for (const column of sortingConfigObj) {
          const { columnName, ascending } = column;
          const columnIndex = dataFilters.header.indexOf(columnName);
          const sortOrder = ascending ? 1 : -1;
          if (a[columnIndex] < b[columnIndex]) {
            return -1 * sortOrder;
          }
          if (a[columnIndex] > b[columnIndex]) {
            return sortOrder;
          }
        }
        return 0;
      });
    }
  };

  /**
   * Gets the unfiltered data from the spreadsheet based on the supplied data filters through the Sheets API
   * Then splits the data Into headers and rows
   * @param {Object} dataFilters - The main dataFilters Object - See the getData method for more info
   */
  getHeaderAndRows(dataFilters) {
    const data = this.makeBatchGetByDataFilter(dataFilters.ssId, dataFilters.sId);
    if (!dataFilters.headerStart) {
      dataFilters.headerStart = 0
    }
    if (dataFilters.headerStart == 0) {
      dataFilters.header = data.valueRanges[0].valueRange.values[0];
      dataFilters.rows = data.valueRanges[0].valueRange.values.slice(1);
    } else {
      dataFilters.header = data.valueRanges[0].valueRange.values[dataFilters.headerStart];
      dataFilters.rows = data.valueRanges[0].valueRange.values.slice(dataFilters.headerStart + 1);
    };
  };

  /**
   * Creates a lookup for a header
   * @param {Object} dataFilters - The main dataFilters Object - See the getData method for more info
   * @return {Object} headerIndexLookup - Returns an Object where the key is the name of the column and the value is the index
   */
  createHeaderIndexLookup(dataFilters) {
    let headerIndexLookup = {};
    dataFilters.header.forEach((element) => (headerIndexLookup[element] = dataFilters.header.indexOf(element)));
    return headerIndexLookup;
  };

  /**
   * Turns the supplied column filter Into indexes 
   * @param {Object} dataFilters - The main dataFilters Object - See the getData method for more info
   */
  getColumnIndexes(dataFilters) {
    dataFilters.colIndexes = dataFilters.colFilters.map((element) => parseInt(dataFilters.header.indexOf(element)));
  };

  /**
   * Turns the supplied row filter Into indexes 
   * @param {Object} dataFilters - The main dataFilters Object - See the getData method for more info
   */
  getRowIndexes(dataFilters) {
    let rowIndexes = {};
    Object.keys(dataFilters.rowFilters).forEach((element) => {
      if (!dataFilters.header.includes(element)) {
        throw new Error("Header Error. Please check your filter names and/or header start row.")
      }
      rowIndexes[dataFilters.header.indexOf(element)] = new Set(dataFilters.rowFilters[element]);
    });
    dataFilters.rowIndexes = rowIndexes;
  };

  /**
   * Remove columns from the supplied Array inplace
   * @param {Array} array - The Array to have columns removed from
   * @param {Array} colIndexes - An Array with the column indexes to keep or not based on the includeColValues value 
   * @param {bool} includeColValues - A bool on wether or not to keep the supplied column indexes
   */
  removeColumns(arr, colIndexes, includeColValues) {
    const colIndexesSet = new Set(colIndexes);
    if (includeColValues) {
      colIndexes = Object.keys(arr).filter((element) => !colIndexesSet.has(parseInt(element)));
    }
    colIndexes.forEach((element, index) => {
      if (index === 0) {
        arr.splice(element, 1);
      } else {
        arr.splice(element - index, 1);
      }
    });
  };

  /**
   * Determines wether or not to keep the row based on the dataFilters obj
   * @param {Array} row - The row of data to check wether to keep or not
   * @param {Object} dataFilters - The main dataFilters Object - See the getData method for more info
   */
  shouldKeepRow(row, dataFilters) {
    let keepRow = [];
    if (Object.keys(dataFilters.rowIndexes).length === 0) {
      return true;
    } else {
      Object.keys(dataFilters.rowIndexes).forEach(element => {
        //element is a string, converting to a number
        element = Number(element);
        let rowValueToSearch = row[element] ? row[element].replace(/\\/g, "") : "";
        keepRow.push(dataFilters.includeRowValues[dataFilters.header[element]] ? dataFilters.rowIndexes[element].has(rowValueToSearch) : !dataFilters.rowIndexes[element].has(rowValueToSearch));
      });
      return !keepRow.includes(false);
    }
  };

  /**
   * Ensures that the row is the same length as the header, when data is pulled in via the google sheets api an empty cell is dropped if there is no data to the right of the cell
   * @param {Int} targetLength - The length to ensure the row is 
   * @param {Object} dataFilters - The main dataFilters Object - See the getData method for more info
   */
  ensureRowLength(row, targetLength, dataFilters) {
    if (dataFilters.addRowNum) {
      while (row.length < targetLength - 1) {
        row.push("");
      }
    } else {
      while (row.length < targetLength) {
        row.push("");
      }
    }
  };

  /**
   * Handles the main filtering and modifying of the rows
   * @param {Object} dataFilters - The main dataFilters Object - See the getData method for more info
   */
  filterAndModifyRows(dataFilters) {
    if (dataFilters.addRowNum) {
      dataFilters.header.push("rowNum");
    }
    for (let i = dataFilters.rows.length - 1; i >= 0; i--) {
      const row = dataFilters.rows[i];
      const dataRowIndex = i;
      const spreadsheetRowIndex = (dataFilters.headerStart + 1) + i;
      this.processRow(row, dataRowIndex, spreadsheetRowIndex, dataFilters);
    }
    this.sortRows(dataFilters)
  };

  /**
   * Processes each individual row and determines what need to happen to the row
   * @param {Array} row - The row to process
   * @param {Int} dataRowIndex - The index of the row in the data
   * @param {Int} spreadsheetRowIndex - The index of the row on the spreadSheet
   * @param {Object} dataFilters - The main dataFilters Object - See the getData method for more info
   */
  processRow(row, dataRowIndex, spreadsheetRowIndex, dataFilters) {
    if (dataFilters.rowFilters == null) {
      this.ensureRowLength(row, dataFilters.header.length, dataFilters);
      this.removeColumns(row, dataFilters.colIndexes, dataFilters.includeColValues);
      if (dataFilters.addRowNum) {
        this.addRowNumber(dataFilters.rows, dataRowIndex, spreadsheetRowIndex)
      }
    } else if (this.shouldKeepRow(row, dataFilters)) {
      this.ensureRowLength(row, dataFilters.header.length, dataFilters);
      this.modifyRow(row, dataRowIndex, spreadsheetRowIndex, dataFilters);
    } else {
      this.removeRow(dataFilters.rows, dataRowIndex);
    }
  };

  /** 
   * Modifies the row to either remove columns or add in a row num
   * @param {Array} row - The row to modify
   * @param {Int} dataRowIndex - The index of the row in the data
   * @param {Int} spreadsheetRowIndex - The index of the row on the spreadSheet
   * @param {Object} dataFilters - The main dataFilters Object - See the getData method for more info
  */
  modifyRow(row, dataRowIndex, spreadsheetRowIndex, dataFilters) {
    this.removeColumns(row, dataFilters.colIndexes, dataFilters.includeColValues);
    if (dataFilters.addRowNum) {
      this.addRowNumber(dataFilters.rows, dataRowIndex, spreadsheetRowIndex)
    }
  };

  /**
   * Add a row num to the row
   * @param {Array} rows - The rows from the spreadsheet
   * @param {Int} dataRowIndex - The index of the row in the data
   * @param {Int} spreadsheetRowIndex - The index of the row on the spreadSheet
   */
  addRowNumber(rows, dataRowIndex, spreadsheetRowIndex) {
    rows[dataRowIndex].push(spreadsheetRowIndex);
  };

  /**
   * Removes the row from the original rows inplace
   * @param {Array} rows - The rows from the spreadsheet
   * @param {Int} dataRowIndex - The index of the row in the data
   */
  removeRow(rows, dataRowIndex) {
    rows.splice(dataRowIndex, 1);
  };

  /**
   * Dtermines wether or not to sort the columns and if so sorts the columns based upon the passed in order
   * @param {Object} dataFilters - The main dataFilters Object - See the getData method for more info
   */
  sortColumns(dataFilters) {
    if (dataFilters.colFilters.length !== 0 && dataFilters.includeColValues === true) {
      const desiredHeaderIndexes = dataFilters.colFilters.map((element) => dataFilters.header.indexOf(element));
      dataFilters.header = desiredHeaderIndexes.map((element) => dataFilters.header[element]);
      dataFilters.rows = dataFilters.rows.map((element) => desiredHeaderIndexes.map((el) => element[el]));
    }
  };

  /**
   * Gets and filters the spreadsheet data and returns it as an Object with Arrays
   * @param {Object} dataFilters - The main dataFilters Object - See the getData method for more info
   */
  getAndFilterDataArr(dataFilters) {
    //Check if ssId is supplied if not set to one from constructor
    if (!dataFilters.ssId) { dataFilters.ssId = this.ssId }
    //
    this.getHeaderAndRows(dataFilters);
    let headerIndexLookup = this.createHeaderIndexLookup(dataFilters);
    Object.assign(dataFilters, { headerIndexLookup })
    if (dataFilters.rowFilters != null) {
      this.getRowIndexes(dataFilters);
    }
    if (dataFilters.colFilters == null) {
      Object.assign(dataFilters, { colIndexes: [] }, { colFilters: [] }, { includeColValues: false })
    } else {
      this.getColumnIndexes(dataFilters);
    }
    this.filterAndModifyRows(dataFilters);
    this.removeColumns(dataFilters.header, dataFilters.colIndexes, dataFilters.includeColValues);
    this.sortColumns(dataFilters);
    let finalHeaderIndexes = this.createHeaderIndexLookup(dataFilters);
    const filteredRowsWithHeader = [dataFilters.header].concat(dataFilters.rows);
    this.filteredDataArr = filteredRowsWithHeader;
    const returnObj = {
      header: dataFilters.header,
      rows: dataFilters.rows,
      data: filteredRowsWithHeader,
      headerIndexes: finalHeaderIndexes,
      allHeaderIndexes: headerIndexLookup
    };
    return returnObj;
  };

  /**
   * Gets and filters the spreadsheet data and returns it as an Object
   * @param {Object} dataFilters - The main dataFilters Object - See the getData method for more info
   */
  getAndFilterDataObj(dataFilters) {
    const dataObj = this.getAndFilterDataArr(dataFilters);
    if (!dataObj.header.includes(dataFilters.objKey)) {
      throw new Error("Object key error. Object key is not in the header.")
    }
    const header = dataObj.header
    const indexOfObjKey = header.indexOf(dataFilters.objKey);
    header.splice(indexOfObjKey, 1)
    const sortedFilteredRows = dataObj.rows;
    const sortedfilteredDataObj = {};

    sortedFilteredRows.forEach(row => {
      const key = row[indexOfObjKey];
      row.splice(indexOfObjKey, 1);
      sortedfilteredDataObj[key] = this.convertToObj(header, row);
    });

    this.filteredDataObj = sortedfilteredDataObj;
    const returnObj = {
      header: header,
      rows: sortedfilteredDataObj,
      allHeaderIndexes: dataObj.allHeaderIndexes
    };
    return returnObj;
  };

  /**
   * The main filtering method utilizes the two methods above(getAndFilterDataObj, getAndFilterDataArr) to return data
   * The dataFilters parameter is an Object with multiple keys that can either be excluded or included based upon the desired return data type
   * EX.
   * dataFilters = {
   *  ssId: spreadsheetId(Optional)
      sId: sheet Id
      headerStartIndex: Row index of the header
      rowFilters: An Object with the columns as the keys and the values of row values that you want to include or exclude
      includeRowValues: An Object with the column names as the keys and a boolean as the value which determines whether to include or exclude the values provided in rowFilters
      colFilters: An Array of columns names to either include or exclude
      includeColValues: A boolean value which indicates whether to include or exclude the values in colFilters
      sortingConfig: An Array of Object's where the Object has two elements columnName and ascending, the columnName should have a value of the column name and ascending should be a boolean 
      addRowNum: A boolean value of whether to include the rowNum
    }
   */
  getData(dataFilters) {
    if (dataFilters.valuesToFilterRow) {
      throw new Error("valuesToFilterRow is an invalid dataFilters key")
    } else if (dataFilters.valuesToFilterCol) {
      throw new Error("valuesToFilterCol is an invalid dataFilters key")
    }

    if (dataFilters.objKey) {
      return this.getAndFilterDataObj(dataFilters);
    } else if (dataFilters.colFilters || dataFilters.rowFilters) {
      return this.getAndFilterDataArr(dataFilters);
    } else {
      const { header, rows, data, headerIndexes, allHeaderIndexes } = this.getAndFilterDataArr(dataFilters);
      Logger.log("Here22222")
      Logger.log(data)
      this.dataArr = data;
      return {
        header,
        rows,
        data,
        headerIndexes,
        allHeaderIndexes
      }
    }

  };

  /** ---------------------------------------------------------------------------------------- Ends Main Data Filtering Methods ----------------------------------------------------------------- 
*/

  /** ---------------------------------------------------------------------- Starts Push Data Methods -------------------------------------------------------------------------  
*/

  /**
   * Returns the data from the constructor by type
   * @param {string} typeOfData - The type of data to return filtered or unfiltered
   * @return {Array} - An Array containing the desired data
   */
  getTypeOfData(typeOfData) {
    const dataMap = {
      "filtered": this.filteredDataArr,
      "unfiltered": this.dataArr
    }
    if (!dataMap.hasOwnProperty(typeOfData)) {
      throw new Error(`Invalid type of row: ${typeOfData}`)
    }
    return dataMap[typeOfData];
  }

  /**
   * The main function to push new data to a sheet, contains all of the logic to ensure the sheet is the desired size and the data will fit on the destination sheet
   * @param {Object} pushDataObj - An Object containg the desired destination sheet and any extra formatting requirements. The rows key is optional, it allows you to put data Into the pushDataObj instead of 
   * pulling it from the constructor.
   */
  pushData(pushDataObj) {
    let { ssId, sId, startRowIndex, startColIndex, extraRows, extraCols, data: dataParam, typeOfData } = pushDataObj;
    if (!startRowIndex) {
      startRowIndex = 0
    }
    if (!startColIndex) {
      startColIndex = 0
    }
    if (!extraRows) {
      extraRows = 0
    }
    if (!extraCols) {
      extraCols = 0
    }
    const data = dataParam ? dataParam : this.getTypeOfData(typeOfData);
    const currentRange = this.getSpreadsheetRange(ssId, sId);
    const rowDiff = (data.length - 1) - currentRange.rowIndex + startRowIndex + extraRows
    const colDiff = (data[0].length - 1) - currentRange.colIndex + startColIndex + extraCols
    const dataRange = { rowIndex: (data.length - 1), colIndex: (data[0].length - 1) }
    const modifySheetRangeObj = { ssId, sId, currentRange, dataRange, rowDiff, colDiff };
    this.modifySheetRange(modifySheetRangeObj);
    const endRowIndex = startRowIndex != 0 ? data.length + startRowIndex : data.length;
    const clearSheetRangeObj = { ssId, sId, startRowIndex, endRowIndex, startColIndex, endColIndex: data[0].length };
    this.clearSheetRange(clearSheetRangeObj);
    const updateSheetDataRangeObj = { ssId, sId, startRowIndex, startColIndex, data };
    this.updateSheetDataRange(updateSheetDataRangeObj);
  }

  /**
   * Allows you to update row and columns individually via the google sheets api. 
   * @param {Object} updateRowsObj - An Object containing the desired destination sheet and the update rows data and params.
   */
  updateRows(updateRowsObj) {
    const { ssId, sId, updateRowsParams, updateRowsData } = updateRowsObj;
    let valueInputOption = updateRowsParams ? updateRowsParams.valueInputOption : "USER_ENTERED";
    const request =
    {
      valueInputOption,
      "data": [
      ]
    };
    const rowNums = Object.keys(updateRowsData);
    rowNums.forEach(rowNum => {
      const rowValues = updateRowsData[rowNum];
      if (!Array.isArray(rowValues)) {
        throw new Error(`Your colNum and Values are not an array. Ex. {2: { colNum: 0, values: ["LMA04620", "1"] } } vs. {2: [{ colNum: 0, values: ["LMA04620", "1"] }] }`)
      }
      rowValues.forEach(colValue => {
        const colNum = colValue.colNum;
        const values = colValue.values;
        let tempRequest = {
          "dataFilter": {
            "gridRange": {
              "sheetId": sId,
              "startRowIndex": rowNum,
              "startColumnIndex": colNum
            }
          },
          "values": [values],
          "majorDimension": "ROWS"
        };
        request["data"].push(tempRequest);
      })
    });
    Sheets.Spreadsheets.Values.batchUpdateByDataFilter(request, ssId);
  };

  /** ---------------------------------------------------------------------------------------- Ends Push Data Methods ---------------------------------------------------------------  
*/

  /**---------------------------------------------------------------------- Starts Spreadsheet Formula Methods -----------------------------------------------------------------------
*/

  /**
   * Adds the ability to refresh formulas on a spreadsheet. Not always needed but sheets sometimes have issues with formulas when multiple people are in the sheet.
   * @param {String} sId - A sheet Id
   * @param {Int} startRow - The row to start on
   */
  refreshFormulas(sId, startRow = 0) {
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
              "startRowIndex": startRow
            }
          },
          "values": parsedValues,
          "majorDimension": "ROWS"
        }
      ]
    };
    Sheets.Spreadsheets.Values.batchUpdateByDataFilter(postRequest, this.ssId);
  };

  /**----------------------------------------------------------------------- End Spreadsheet Formula Methods -----------------------------------------------------------------------
*/

  /**---------------------------------------------------------------------- Starts Helper Methods -----------------------------------------------------------------------
*/

  /**
   * @return {Array} - The sheet id's and title's in a 2d Array
   */
  getNamesAndIdsOfSheets(ssIdObj) {
    let { ssId } = ssIdObj;
    if (!ssId) { ssId = this.ssId }
    const sheets = Sheets.Spreadsheets.get(ssId).sheets;
    let namesAndIds = []
    for (const sheet of sheets) {
      let tempSheetId = sheet.properties.sheetId;
      let tempSheetTitle = sheet.properties.title;
      namesAndIds.push([tempSheetId, tempSheetTitle]);
    };
    return namesAndIds;
  };

  /**
   * Adds a column with the fiscal year of the desired date column.
   * @param {String} typeOfData - The type of data to use to calculate the fiscal year from.
   * @param {String} columnToCalculateWith - The column name specifies which column to use for the fiscal year calculation.
   */
  calculateFiscalYear(typeOfData, columnToCalculateWith) {
    let values;
    if (typeOfData == "filtered") {
      values = this.filteredDataArr;
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
      this.filteredDataArr = values;
    } else if (typeOfData == "unfiltered") {
      this.dataArr = values;
    }
  };

  /**
   * Adds a column with the corresponding zone.
   * @param {String} typeOfData - The type of data to use to calculate the zone from.
   * @param {String} regionColumnName - The column name that specifies which column to use for the zone calculation.
   */
  addZone(typeOfData, regionColumnName) {
    let values;
    if (typeOfData == "filtered") {
      values = this.filteredDataArr;
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
      this.filteredDataArr = newValues;
    } else if (typeOfData == "unfiltered") {
      this.dataArr = newValues;
    }
  };

  /**
   * Adds the ability to format data in the code before it lands on a spreadsheet.
   * @param {Object} formatForCols - An Object containing the desired format, 
   */
  formatData(formatForCols, typeOfData) {
    const data = this.getTypeOfData(typeOfData);
    let header = data[0];
    let rows = data.slice(1);
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
    if (typeOfData == "filtered") {
      const filteredRowsWithHeader = [header].concat(rows);
      this.filteredDataArr = filteredRowsWithHeader;
    } else if (typeOfData == "unfiltered") {
      const rowsWithHeader = [header].concat(rows);
      this.dataArr = rowsWithHeader;
    }
  }

  /**
   * Adds data validation to a sheet.
   * @param sId The sheet id
   * @param dataValidationForCols An Object with the desired options.
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

  /**
   * Convert a date to a date with the format of mm-dd-yyyy
   * @param typeOfData The type of data to be used
   * @param columnsToConvert An Array of column names
   */
  convertDates(typeOfData, columnsToConvert) {
    let values;

    if (typeOfData == "filtered") {
      values = this.filteredDataArr;
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
      this.filteredDataArr = values;
    } else if (typeOfData == "unfiltered") {
      this.dataArr = values;
    }
    return values;

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
        throw new Error(error)
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
   * Adds the ability to format a range on a sheet.
   * formatObj = {
   *  sId: string,
   *  headerStart: Integer,
   *  formatForCols: Object,
   *    colNum :{
   *      type: the desired user entered format
   *    }
   * }
   * @param {Object} formatObj - An Object with keys of sId, formatForCols and headerStart
   */
  formatRange(formatObj) {
    let { ssId, sId, formatForCols, headerStart } = formatObj;
    if (!ssId) { ssId = this.ssId }
    const cols = Object.keys(formatForCols);
    for (let col of cols) {
      const format = formatForCols[col];
      col = Number(col)
      const request = {
        "requests": [
          {
            "repeatCell": {
              "cell": {
                "userEnteredFormat": format
              },
              "range": {
                "sheetId": sId,
                "startRowIndex": headerStart,
                "startColumnIndex": col,
                "endColumnIndex": col + 1
              },
              "fields": "userEnteredFormat"
            }
          }]
      };
      Sheets.Spreadsheets.batchUpdate(request, ssId);
    }
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
      values = this.filteredDataArr;
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
      this.filteredDataArr = values;
    } else if (typeOfData == "unfiltered") {
      this.dataArr = values;
    }
    return values;
  };

  /**
   * Pastes a formula on to a sheet.
   * @param {String} ssId - The spreadsheet id.
   * @param {String} sId - The sheet id.
   * @param {Integer} rowNum - The row num to start on.
   * @param {Integer} colNum - The col num to start on.
   * @param {String} formula - The formula to put in the cell.
   */

  pasteFormula(ssId, sId, rowNum, colNum, formula) {
    const request =
    {
      "requests": [
        {
          "repeatCell": {
            "range": {
              "sheetId": sId,
              "startRowIndex": rowNum,
              "endRowIndex": rowNum + 1,
              "startColumnIndex": colNum,
              "endColumnIndex": colNum + 1
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
    Sheets.Spreadsheets.batchUpdate(request, ssId);
  };

  sortRowsOnSheet(sortObj) {
    const request =
    {
      "requests": [
        {
          "sortRange": {
            "range": {
              "sheetId": sortObj.sId,
              "startRowIndex": sortObj.dataStart,
              "startColumnIndex": 0,
            },
            "sortSpecs": sortObj.sortSpecs
          }
        }
      ]
    };
    Sheets.Spreadsheets.batchUpdate(request, sortObj.ssId);
  }

  /**
   * Returns the unique data of the column
   * {
   *  columnName: "",
   *  typeOfData: "",
   *  data: [[],[]]
   * }
   */
  getUniqueRowsByCol(uniqueObj) {
    let { colName, typeOfData, includeAllCols = true, data } = uniqueObj;
    if (typeOfData) {
      data = this.getTypeOfData(typeOfData);
    }
    const header = data[0];
    const rows = data.slice(1);
    const indexOfColName = header.indexOf(colName);
    if (includeAllCols) {
      const getColumnUniqueRows = (arr, colIndex) => {
        const seen = {};
        return arr.reduce((acc, row) => {
          if (!seen[row[colIndex]]) {
            seen[row[colIndex]] = true;
            acc.push(row);
          }
          return acc;
        }, []);
      }
      return getColumnUniqueRows(rows, indexOfColName);
    } else {
      return [...new Set(rows.map(row => row[indexOfColName]))];
    }

  }

  /**----------------------------------------------------------------------- Ends Helper Methods -------------------------------------------------------
*/

  /**---------------------------------------------------------------------- Starts Spreadsheet Size Methods --------------------------------------------------------------------------------
*/

  /**
   * Maintains a certain number of rows blank rows on a sheet
   * @param {String} sId - The sheet id.
   * @param {Int} numOfRows - The number of blank rows to maIntain on the sheet.
   */
  maintainNumOfRowsOnSheet(maintainRowsObj) {
    let { ssId, sId, numOfRows } = maintainRowsObj;
    if (!ssId) { ssId = this.ssId }
    const sheetInfo = this.getSpreadsheetRange(ssId, sId);
    const rowIndex = sheetInfo["rowIndex"] + 1;
    if (rowIndex < numOfRows) {
      const rowsToAdd = numOfRows - rowIndex;
      this.updateSheetRange(ssId, [this.buildInsertRequest({ sId, dimension: "ROWS", startIndex: rowIndex, diff: rowsToAdd })]);
    }
  };

  /**
   * Adds the ability to delete a single row
   * @param {Object} deleteRowObj - An object containing ssId, sId, and the row num to delete.
   * 
   */
  deleteRow(deleteRowObj) {
    let { ssId, sId, rowNum } = deleteRowObj;
    if (!ssId) { ssId = this.ssId }
    this.updateSheetRange(ssId, [this.buildDeleteRequest({ sId, dimension: "ROWS", startIndex: rowNum, diff: 1 })])
  };

  /**
   * Adds the ability to delete multiple rows at once. 
   * @param {Object} deleteRowsObj - An object containing ssId, sId and the rows nums(Array) to delete.
   */
  deleteRows(deleteRowsObj) {
    let { ssId, sId, rowNums } = deleteRowsObj;
    if (!ssId) { ssId = this.ssId }
    if (rowNums.length != 0) {
      const rowsToDelete = [];
      rowNums.sort(function (a, b) { return a - b }).forEach((rowNum, index) => {
        let rowNumToDelete;
        if (index == 0) {
          rowNumToDelete = rowNum;
        } else {
          rowNumToDelete = rowNum - index;
        }
        rowsToDelete.push(this.buildDeleteRequest({ sId, dimension: "ROWS", startIndex: rowNumToDelete, diff: 1 }))
      })
      this.updateSheetRange(ssId, rowsToDelete);
    } else {
      Logger.log("No rows were deleted.")
    }
  };

  /**
   * Returns the last row of the current sheet plus one.
   * @param {Object} nextRowObj - An object containing ssId and sId. ssId is optional.
   * @return {Int} - The index of the next row.
   */
  getNextRow(nextRowObj) {
    let { ssId, sId } = nextRowObj;
    if (!ssId) { ssId = this.ssId }
    const nextRow = this.getData({ ssId, sId: sId, headerStart: 0 }).data.length;
    let sheet2Name = "";
    this.getNamesAndIdsOfSheets({ ssId }).forEach((e => { if (e[0] == sId) { sheet2Name = e[1] } }));
    const lastRow = SpreadsheetApp.openById(ssId).getSheetByName(sheet2Name).getMaxRows();
    if (nextRow > lastRow) {
      this.addRows(sId, 10);
    }
    return nextRow;
  };

  /**
   * Adds rows to the inputted sheet
   * @param {Object} nextRowObj - An object containing ssId, sId and the num of rows to add. ssId is optional.
   */
  addRows(addRowsObj) {
    const { ssId, sId, rowsToAdd } = addRowsObj;
    if (!ssId) { ssId = this.ssId }
    const sheetInfo = this.getSpreadsheetRange(ssId, sId);
    this.updateSheetRange(ssId, this.buildInsertRequest({ sId, dimension: "ROWS", startIndex: sheetInfo["rowIndex"] + 1, diff: rowsToAdd }))
  };

  /**
   * Determines how the destination sheet's range needs to be modified
   * @param {Object} modifySheetRangeObj - An Object containing the modifications needed to the destination sheet. ssId is optional.
   */
  modifySheetRange(modifySheetRangeObj) {
    const { ssId, sId, currentRange, dataRange, rowDiff, colDiff } = modifySheetRangeObj;
    if (!ssId) { ssId = this.ssId }
    const rangeObj = [];
    if (rowDiff > 0) {
      rangeObj.push(this.buildInsertRequest({ sId, dimension: "ROWS", startIndex: currentRange.rowIndex, diff: rowDiff }))
    } else if (rowDiff < 0) {
      rangeObj.push(this.buildDeleteRequest({ sId, dimension: "ROWS", startIndex: dataRange.rowIndex, diff: Math.abs(rowDiff) }))
    }
    if (colDiff > 0) {
      rangeObj.push(this.buildInsertRequest({ sId, dimension: "COLUMNS", startIndex: currentRange.colIndex, diff: colDiff }))
    } else if (colDiff < 0) {
      rangeObj.push(this.buildDeleteRequest({ sId, dimension: "COLUMNS", startIndex: dataRange.colIndex, diff: Math.abs(colDiff) }))
    }
    if (rangeObj.length != 0) { this.updateSheetRange(ssId, rangeObj) };
  }

  /**
   * Calls the sheets api to clear the range of a sheet
   * @param {Object} clearSheetRangeObj - An Object containg the range config to be cleared
   */
  clearSheetRange(clearSheetRangeObj) {
    const { ssId, sId, startRowIndex, endRowIndex, startColIndex, endColIndex } = clearSheetRangeObj;
    const clearRequest = {
      "dataFilters": [
        {
          "gridRange": {
            "sheetId": sId,
            "startRowIndex": startRowIndex,
            "endRowIndex": endRowIndex,
            "startColumnIndex": startColIndex,
            "endColumnIndex": endColIndex
          }
        }
      ]
    };
    Sheets.Spreadsheets.Values.batchClearByDataFilter(clearRequest, ssId);
  }

  /**
   * Calls the sheets api to update a sheets range with new data
   * @param {Object} updateSheetDataRangeObj - An Object containing the updated data for the range
   */
  updateSheetDataRange(updateSheetDataRangeObj) {
    const { ssId, sId, startRowIndex, startColIndex, data } = updateSheetDataRangeObj;
    const updateRequest = {
      "valueInputOption": "USER_ENTERED",
      "data": [
        {
          "dataFilter": {
            "gridRange": {
              "sheetId": sId,
              "startRowIndex": startRowIndex,
              "startColumnIndex": startColIndex
            }
          },
          "values": data,
          "majorDimension": "ROWS"
        }
      ]
    };

    Sheets.Spreadsheets.Values.batchUpdateByDataFilter(updateRequest, ssId);
  }

  /**
   * Does a batchupdate call to the sheets api 
   * @param {string} ssId - A string containing the spreadsheet id
   * @param {Object} rangeObj - A formatted request containg the config to update the sheet range
   */
  updateSheetRange(ssId, rangeObj) {
    const request =
    {
      "requests": rangeObj
    };
    Sheets.Spreadsheets.batchUpdate(request, ssId);
  };

  /**
   * Builds the insert range request
   * @param {Object} requestParams - An Object with the config to update the sheet range
   */
  buildInsertRequest(requestParms) {
    return {
      "insertDimension": {
        "range": {
          "sheetId": requestParms.sId,
          "dimension": requestParms.dimension,
          "startIndex": requestParms.startIndex,
          "endIndex": requestParms.startIndex + requestParms.diff
        },
        "inheritFromBefore": requestParms.startIndex == 0 ? false : true
      }
    }
  }

  /**
   * Builds the delete range request
   * @param {Object} requestParams - An Object with the config to delete a sheets range
   */
  buildDeleteRequest(requestParms) {
    return {
      "deleteDimension": {
        "range": {
          "sheetId": requestParms.sId,
          "dimension": requestParms.dimension,
          "startIndex": requestParms.startIndex,
          "endIndex": requestParms.startIndex + requestParms.diff
        }
      }
    }
  }

  /**
   * Gets the grid range column and row indexes.
   * @param {String} ssId - The spreadsheet id.
   * @param {String} sId - The sheet id.
   * @return {Object} - An Object containing the colIndex and rowIndex.
   */
  getSpreadsheetRange(ssId, sId) {
    const sheets = Sheets.Spreadsheets.get(ssId).sheets;
    let columnIndex;
    let rowIndex;
    for (const sheet of sheets) {
      const sheetProperties = sheet.properties;
      const sheetId = sheetProperties.sheetId;
      if (sheetId == sId) {
        columnIndex = sheetProperties.gridProperties.columnCount - 1;
        rowIndex = sheetProperties.gridProperties.rowCount - 1;
      };
    };
    const returnedObj = {
      "colIndex": columnIndex,
      "rowIndex": rowIndex
    };
    return returnedObj;
  };

  /**
   * Gets the value range from the sheet.
   * @param {String} ssId - The spreadsheet id.
   * @param {String} sId - The sheet id.
   * @return {Object} - An Object containing the colIndex and rowIndex.
   */
  getSpreadsheetDataRange(ssId, sId) {
    const request =
    {
      'spreadsheetId': ssId,
      "dataFilters": [
        {
          "gridRange": {
            "sheetId": sId
          }
        }
      ]
    };
    const valueRangeData = Sheets.Spreadsheets.Values.batchGetByDataFilter(request, ssId);
    const valueColIndex = valueRangeData["valueRanges"][0]["valueRange"]["values"][0].length - 1;
    const valueRowIndex = valueRangeData["valueRanges"][0]["valueRange"]["values"].length - 1;

    const returnedObj = {
      "colIndex": valueColIndex,
      "rowIndex": valueRowIndex
    };
    return returnedObj;
  }

  /**
   * Remove all extra cols and rows that do not have data in them.
   * @param {String} ssId - The spreadsheet id.
   * @param {Sting} sId - The sheet id.
   */
  cleanExtraRowsAndColumns(ssId, sId) {
    const currentRange = this.getSpreadsheetRange(ssId, sId);
    const dataRange = this.getSpreadsheetDataRange(ssId, sId)
    const rowDiff = dataRange.rowIndex - currentRange.rowIndex;
    const colDiff = dataRange.colIndex - currentRange.colIndex;
    dataRange.colIndex += 1;
    dataRange.rowIndex += 1;
    const modifySheetRangeObj = { ssId, sId, currentRange, dataRange, rowDiff, colDiff };
    this.modifySheetRange(modifySheetRangeObj);
  };

  /**----------------------------------------------------------------------- Ends Spreadsheet Size Methods -----------------------------------------------------------------------
*/



};


