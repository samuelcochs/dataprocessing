# Data Processing Class Documentation

This class provides a comprehensive set of methods to interact and manipulate data within Google Sheets using the Google Sheets API. From fetching data to formatting and validation, this helper class acts as an essential utility for developers aiming to extend Google Sheets functionalities programmatically.

## Setup and Initialization

Before using the methods, make sure you have set up and authenticated the Google Sheets API. Initialize the class as follows:
```
const dataProc = Data_Processing.exportDataProcessing(Spreadsheet Id);
```

## Helper Methods
Spreadsheet Size Methods: Describe these methods here.
1. Method Examples:
  - Using the getData() method:
    - The getData() method takes in one argument, the argument is an obj. Listed below are all of the options you can use. I have seperated them into required and optional below.
      - Required:
        - sId: sheet Id
      - Optional
        - ssId: spreadsheetId, you only have to input this if you want to change the spreadsheet id from the one you used to intialize the class
        - headerStart: Row index of the header
        - rowFilters: An Object with the columns as the keys and the values of row values that you want to include or exclude
        - includeRowValues: An Object with the column names as the keys and a boolean as the value which determines whether to include or exclude the values provided in rowFilters
        - colFilters: An Array of columns names to either include or exclude
        - includeColValues: A boolean value which indicates whether to include or exclude the values in colFilters
        - sortingConfig: An Array of Object's where the Object has two elements columnName and ascending, the columnName should have a value of the column name and ascending should be a boolean
        - addRowNum: A boolean value of whether to include the rowNum
        - limitRows: The number of rows you want to return
      - There are two main ways to use the getData method both return an obj but one returns the data in the obj as an obj and the other returns it as a 2d array.
        - Obj Option:
        ```javascript
        const {header, rows, allHeaderIndexes} = dataProc.getData({
          ssId: Spreadsheet Id,
          sId: Sheet Id,
          objKey: Column header name to use as the key
        })
        ```
        - Array Option:
        ```javascript
        const { header, rows, data, headerIndexes, allHeaderIndexes } = dataProc.getData({
          ssId: Spreadsheet Id,
          sId: Sheet Id,
        })
        ```
        - Adding in header start index:
        ```javascript
        const { header, rows, data, headerIndexes, allHeaderIndexes } = dataProc.getData({
          ssId: Spreadsheet Id,
          sId: Sheet Id,
          headerStart: int
        })
        ```
        - Adding in row filtering:
        ```javascript
        const { header, rows, data, headerIndexes, allHeaderIndexes } = dataProc.getData({
          ssId: Spreadsheet Id,
          sId: Sheet Id,
          rowFilters: {Column Name: [Values to filter],
          includeRowValues: {Column Name: boolean}
        })
        ```
        - Adding in column filtering:
        ```javascript
        const { header, rows, data, headerIndexes, allHeaderIndexes } = dataProc.getData({
          ssId: Spreadsheet Id,
          sId: Sheet Id,
          colFilters: [Column Names],
          includeColValues: boolean
        })
        ```
        - Adding in sorting:
        ```javascript
        const { header, rows, data, headerIndexes, allHeaderIndexes } = dataProc.getData({
          ssId: Spreadsheet Id,
          sId: Sheet Id,
          sortingConfig: [{columnName: Column Name, ascending: true}]
        })
        ```
        - Adding in row num:
        ```javascript
        const { header, rows, data, headerIndexes, allHeaderIndexes } = dataProc.getData({
          ssId: Spreadsheet Id,
          sId: Sheet Id,
          addRowNum: boolean
        })
        ```
        - Adding in limit rows:
        ```javascript
        const { header, rows, data, headerIndexes, allHeaderIndexes } = dataProc.getData({
          ssId: Spreadsheet Id,
          sId: Sheet Id,
          limitRows: int
        })
        ```
        - Adding in retrying:
        ```javascript
        const { header, rows, data, headerIndexes, allHeaderIndexes } = dataProc.getData({
          ssId: Spreadsheet Id,
          sId: Sheet Id,
          retryConfig: {
            retries: 2,
            delay: 4,
            maxDelay: 60
          }
        })
        ```
        - Final example with real data:
        ```javascript
        const { header, rows, data, headerIndexes, allHeaderIndexes } = dataProc.getData({
          ssId: "1bAgj9v08bJH85Cz8NwwjNWlBeTAKDzq0Ka0WL6oNP_0",
          sId: 0,
          headerStart: 0,
          rowFilters: { "Zone": ["1"] },
          includeRowValues: { "Zone": true },
          colFilters: ["Lease Num", "Zone", "State"],
          includeColValues: true,
          sortingConfig: [{ columnName: "State", ascending: true }],
          addRowNum: true,
          limitRows: 2
        })
  
        This is the result you should expect from the data variable:
        [
          [Lease Num, Zone, State],
          [LCT00346, 1, CA],
          [LCT00286, 1, IN]
        ]
        ```
  - Using the updateRows() method: 
    - This method requires as its only argument. The object should be structured as follows:
    ```javascript
    const updateRowsObj = {
      ssId: Spreadsheet Id,
      sId: Sheet Id,
      updateRowsData: { rowNum: [{ colNum: Column Index, values: An array of values to update the row with }] }
    }
    ```
    - A real example of this is:
    ```javascript
    const updateRowsObj = {
      ssId: "1bAgj9v08bJH85Cz8NwwjNWlBeTAKDzq0Ka0WL6oNP_0",
      sId: 817299795,
      updateRowsData: {
        2:
          [
            { colNum: 0, values: ["LMA04620"] },
            { colNum: 4, values: ["test"]}
          ],
        3:
          [
            { colNum: 0, values: ["LTX03123","Test"] }
          ]
      }
    }
    ```
  - Using the pushData() method: 
    - This method requires an object as its only argument. The object should be strucutred as follows:
    ```javascript
    const pushDataObj = {
      ssId: Spreadsheet Id,
      sId: Sheet Id,
      startRowIndex: Index of row to start on,
      startColIndex: Index of col to start on,
      extraRows: The amount of extra rows you would like after the data,
      extraCols: The amount of extra cols you would like on the right side of the data,
      data: A 2d array of data that you would like to put on the spreadsheet,
      typeOfData: If you would like to use the data in the constructor and leave the data key blank then input filtered or unfiltered here
    }
    ```
    - Basic Example(Using all required inputs)
    ```javascript
    const pushDataObj = {
      ssId: Spreadsheet Id,
      sId: Sheet Id,
      data: [[1,2],[3,4]]
    }
    or
    const pushDataObj = {
      ssId: Spreadsheet Id,
      sId: Sheet Id,
      typeOfData: 'filtered'
    }
    ```
    - More complex example using all options
      - This will start on row 2 and column 
    ```javascript
    cosnt pushDataObj = {
      ssId: Spreadsheet Id,
      sId: Sheet Id,
      data: [[1,2],[3,4]],
      startRowIndex: 1,
      startColIndex: 1,
      extraRows: 10,
      extraCols: 10
    }
    ```


    
  - Using the getNextRow() method:
    - Returns the next avaliable row with no data on the desired spreadsheet. If there are no avaliable rows it will automatically add 5 rows to ensure there is one avaliable.
    ```javascript
    const nextRowObj = {
      ssId: Spreadsheet Id,
      sId: Sheet Id,
    }
    dataProc.getNextRow(nextRowObj)
    ```



    
  - Using the tryCatchWithRetries() method:
    - This method will automatically retry to run a function if it errors out and increase the delay exponentially.
   



  - Using the getUniqueRowsByCol() method:
    - This method will return all of the unique values in a column. 
    ```javascript
    dataProc.getUniqueRowsByColumn({colName: "Lease Num", typeOfData: "unfiltered" })
    ```
  -


Template
  - Using the methodName() method:
    - Description:
    - Arguments: 
    - Argument Descriptions:
    - Practical Example



























