# Tag 0.2.0

## Wednesday, May 25, 2016
* Updated ExcelDataFormatter::addRow() and ExcelDataFormatter::getFieldsForObj() so that DataObjects can suggest a default set of Fields to Export.
* Added the ability for the ExcelDataFormatter to use field labels instead of just displaying the Database Field name.
* Added a change log.
* Updated the ReadMe with information about how to choose the column and customise the headers.
* Generalise the logic on the GridFieldExcelExportButton to minimise code duplication between the XLSX, XLS and CSV action.
* Update GridFieldExcelExportAction and GridFieldExcelExportButton so we can set the UseLabelsAsHeaders flag on the underlying ExcelDataFormatter.
