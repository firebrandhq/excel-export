# Silverstripe Excel Export module
This Silverstripe module makes it easy to export a set of Silverstripe DataObjects to:
* Excel 2007 (XLSX)
* Excel 5 (XLS)
* CSV

This module is built by extending the standard [SilverStripe DataFormatter](http://api.silverstripe.org/3.1/class-DataFormatter.html).

## Requirements

 * [silverstripe/cms](https://github.com/silverstripe/silverstripe-cms) >=3.1
 * [phpoffice/phpexcel](https://github.com/PHPOffice/PHPExcel) >=1.8

## Suggestions
* [silverstripe/restfulserver](https://github.com/silverstripe/silverstripe-restfulserver) >=3.1

## Installation

Install the module through [composer](http://getcomposer.org):

```bash
composer require firebrandhq/silverstripe-excel-export
```

## Exporting your DataObjects
There's 3 ways you can export your data to a spread sheet.

### Programmatically by calling the DataFormatter directly
3 DataFormatters are provided:
* ExcelDataFormatter for XLSX
* OldExcelDataFormatter for XLS
* CsvDataFormatter for CSV

You can manually instantiate them to convert a list of DataObjects or a single DataObject.

```
$formatter = new ExcelDataFormatter();

// Will return an Excel Spreadsheet as a string for a single user
$filedata = $formatter->convertDataObject($user);

// Will return an Excel Spreadsheet as a string for a list of user
$filedata = $formatter->convertDataObjectSet(Member::get());
```

`convertDataObjectSet()` and `convertDataObject()` will automatically set the _Content-Type_ HTTP header to an appropriate Mime Type.

You can also retrieve the underlying _PHPExcel_ object and export your DataObject set to whatever format supported by _PHPExcel_.

```
// Get your Data
$formatter = new ExcelDataFormatter();
$excel = $formatter->getPhpExcelObject(SiteTree::get());

// Set up a writer
$writer = PHPExcel_IOFactory::createWriter($excel, 'HTML');

// Save the file somewhere on the server
$writer->save('/tmp/sitetree_list.html');

// Output the results back to the browser
$writer->save('php://output');

// Output the file to a variable
ob_start();
$writer->save('php://output');
$fileData = ob_get_clean();
```

### Add the GridFieldExcelExportButton to a GridField
The `GridFieldExcelExportButton` allows your CMS users to easily export the data from a GridField to a spreadsheet.

```
$rowEntryConfig = GridFieldConfig_RecordEditor::create();
$rowEntryConfig->addComponent(new GridFieldExcelExportButton());
$rowEntryDataGridField = new GridField(
    "ContentRow",
    "Content Row Entry",
    $this->ContentRow(),
    $rowEntryConfig
);
$fields->addFieldToTab('Root.Main', $rowEntryDataGridField);
```

The above code snippet will display a split button allowing the user to export the GridField list to the format of their choice.

Unlike the SilverStripe [GridFieldExportButton](http://api.silverstripe.org/3.1/class-GridFieldExportButton.html), the `GridFieldExcelExportButton` will export all the fields of the provided DataObjects ... not just the summary fields.

You can also use the `GridFieldExcelExportAction` component. This button is added to each row and allows you to export individual records one at a time. Out of the box, `GridFieldExcelExportAction` will export to _xlsx_, but you can get it to export to _xls_ or _csv_ (e.g.: `new GridFieldExcelExportAction('csv')`).

`GridFieldExcelExportAction` and `GridFieldExcelExportButton` can be used in conjunction if you want to give both options to your users.

### Call via the SilverStripe RestfulServer Module
The [SilverStripe RestfulServer Module](https://github.com/silverstripe/silverstripe-restfulserver) allows you to turn any SilverStripe website into a RESTFul Server.

If you use the _SilverStripe RestfulServer Module_ in conjunction with the _Silverstripe Excel Export module_, you'll be able to dynamically export any DataObject set just by entering the right URL in your browser.

#### Access control
Obviously, you don't want everyone to be able to download any data off your website. The SilverStripe RestfulServer Module will only return results for DataObject with the `$api_access` property set.

```
private static $api_access = true;
```

Additionally, access to individual DataObjects is controlled by the `canView` function.

[Configuration the SilverStripe RestfulServer Module ](https://github.com/silverstripe/silverstripe-restfulserver#configuration)

#### Getting to the data
Exporting your data is just as easy as entering a URL.
* Get a list of all Pages in Excel 2007: *http://localhost/api/v1/Page.xlsx*
* Get a list of all Pages in Excel 5: *http://localhost/api/v1/Page.xls*
* Get a list of all Pages in CSV: *http://localhost/api/v1/Page.csv*
* Limit the list to 10 results: *http://localhost/api/v1/Page.csv?limit=10*
* Return a single record: *http://localhost/api/v1/Page/37.xlsx*
* Drill down into relationships: *http://localhost/api/v1/Tag/127/Articles.xlsx*

[SilverStripe RestfulServer Module Supported operations](https://github.com/silverstripe/silverstripe-restfulserver#supported-operations)

## Customising the output

There's 2 ways you can control the output:
* Choose which fields to output ;
* Choose to use field label instead of fields names in the headers.

### Choose which fields to output
Because the `ExcelDataFormatter` extends [DataFormatter](http://api.silverstripe.org/3.3/class-DataFormatter.html), you can use methods like `setCustomFields()`, `setCustomAddFields()` or `setRemoveFields()` to control what fields will be present in the spread sheet.

```
$formatter = new ExcelDataFormatter();

// This formatter instead of returning every field of a DataObject, will only return 3 fields.
$formatter->setCustomFields(['ID', 'Title', 'LastEdited']);

// If youe DataObject has dynamic properties, you can reference them using setCustomAddFields().
$formatter->setCustomAddFields(['ChildrenCount']);
```

#### Defining a default column set
You can customise the default column set that will be return for a specific DataObject class by defining a `getExcelExportFields()` method on your DataOject class.

This `getExcelExportFields()` method should return an array of fields following the same format used by `DataObject::inheritedDatabaseFields()`:
```
return [
    'ID' => 'Int',
    'Name' => 'Varchar',
    'Address' => 'Text'
];
```

You may also reference relationships in this array or dynamic properties:
```
return [
    'Owner.Name' => 'Varchar',
    'Category.Title' => 'Varchar',
    'ChildrenCount' => 'Int',
];
```

This will also allow you to control the order the fields appear in the Spread Sheet. Note that ID will always be the first field and cannot be removed.

This behavior can be overriden for specific instances of `ExcelDataFormatter` by calling the `setCustomFields()` method.

## Use field labels or field names as column headers
Out of the box, the actual field names will be used as column header. (e.g.: `FirstName` rather than `First Name`).

You can customise this behavior and use the Field Labels as define on your DataObject class instead. When generating the header row, `ExcelDataFormatter` will call the `fieldLabel()` method on your Data Object to decide what string to use in each header.

### Change the default for all `ExcelDataFormatter`
In you YML config, you can use the following syntax to change the default headers.
```
ExcelDataFormatter:
  UseLabelsAsHeaders: true
```

### Override the default for a specific instance
You may change the default behavior for a specific instance.
```
$formatter->setUseLabelsAsHeaders(true);
```
