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
        - headerStartIndex: Row index of the header
        - rowFilters: An Object with the columns as the keys and the values of row values that you want to include or exclude
        - includeRowValues: An Object with the column names as the keys and a boolean as the value which determines whether to include or exclude the values provided in rowFilters
        - colFilters: An Array of columns names to either include or exclude
        - includeColValues: A boolean value which indicates whether to include or exclude the values in colFilters
        - sortingConfig: An Array of Object's where the Object has two elements columnName and ascending, the columnName should have a value of the column name and ascending should be a boolean
        - addRowNum: A boolean value of whether to include the rowNum
      - There are two main ways to use the getData method both return an obj but one returns the data in the obj as an obj and the other returns it as a 2d array.
      ```
      Obj Option:
      const {header, rows, allHeaderIndexes} = dataProc.getData({
        ssId: Spreadsheet Id,
        sId: Sheet Id,
        objKey: Column header name to use as the key
      })

      Array Option:
      const { header, rows, data, headerIndexes, allHeaderIndexes } = dataProc.getData({
        ssId: Spreadsheet Id,
        sId: Sheet Id,
      })

      Adding in row filtering:
      const { header, rows, data, headerIndexes, allHeaderIndexes } = dataProc.getData({
        ssId: Spreadsheet Id,
        sId: Sheet Id,
        rowFilters: {Column Name: [Values to filter],
        includeRowValues: {Column Name: boolean}
      })

      Adding in column filtering:
      const { header, rows, data, headerIndexes, allHeaderIndexes } = dataProc.getData({
        ssId: Spreadsheet Id,
        sId: Sheet Id,
        colFilters: [Column Names],
        includeColValues: boolean
      })
      ```
  - Using the updateRows() method: 
    - Explain how to use it.
  - Using the getNextRow() method: 
    - Explain how to use it.
  - Using the updateRows() method: 
    - Explain how to use it.

## Conclusion
The Google Sheets Helper Class offers a multitude of methods that streamline various operations on Google Sheets. With this class, developers can efficiently interact with sheets, format data, and more. Ensure you have the necessary permissions and have set up the Google Sheets API correctly before diving in.


Remember to fill in placeholders (like "Sample code for initialization goes here") with the actual content, and you can add more explanations or sample code for the methods as needed.