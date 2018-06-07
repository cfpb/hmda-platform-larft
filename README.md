# Excel Data Entry Application

This is an Excel application to help financial institutions who want to use it store their HMDA data and format as required by the [FFIEC's
2017 File Specification](http://www.consumerfinance.gov/data-research/hmda/static/for-filers/2017/2017-HMDA-FIG.pdf).

## Dependencies

A copy of Excel is required to open the application on a Windows computer, it is not compatible with Mac computers.
It has been tested only in the Excel 2010 environment.

## Using the Workbook

This workbook is composed of two worksheet.

### Data

Transmittal sheet data is entered on the 3rd row under the blue header in the fields specified by the blue header.

LAR data is entered starting on the 5th row under the green header in the fields specified by the grean header. Each additional LAR should entered on subsequent lines such that the first LAR is on row 5, the second on row 6, and so on.

The export button will format the LAR and Transmittal sheet data to be pipe delimited and add the appropriate record identifier field of "1" for the Transmittal Sheet or "2" for LARs. It will then allow a user to choose a .txt file to export the data into.

### Export

The sheet holds pipe delimted formatted LAR and transmittal sheet data before it is exported into .txt document. It is automatically generated or re-generated when the "Export" button is hit and should not be interacted with.

## Known Issues

The application has been tested on Microsoft Excel 2010 for Windows. Other versions are currently not supported.