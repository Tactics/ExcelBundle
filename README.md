ExcelBundle
===========
## Getting started
###### Install the bundle using composer
Add following lines to composer.json 
```
{
   "type": "git",
   "url": "git@github.com:Tactics/ExcelBundle"
}

"tactics/excel-bundle": "1.0.*" 
```     

## Basic usage
Getting the excelbuilder service is as easy as:
```
$builder = $this->get('tactics.excel_writer');
```

## Features
##### Exporting collections
###### Exporting array of values
Write an array of values to an excel file
```
//Make array with some kind of values 
$collection = array("First value to write", "Second value to write", ...);
//Offers the excel with values as a file download with filename = awesomeness.xlsx
$builder->getDownloadFromCollection($collection, array(
'filename' => 'awesomeness'                     
));
```
###### Exporting entity collection
Write a collection of objects to an excel file 
```
//get a collection of objects
$collection = $this->getDoctrine()->getManager()->getRepository('SomeBundle:SomeEntity')->findAll();
//Offers the excel with values as a file download with filename = awesomeness.xlsx
$builder->getDownloadFromCollection($collection, array(
'filename' => 'awesomeness',       
));
```
##### Getting excel file from excel template.
In stead of writing to an empty excel file you can create a template excel file and write to this file.
```
//$builder = valid ReportExcelWriter.
$excel = $builder->getExcelFromFile('my_templates_file_name', kernel_service);
$builder->writeToSheet($excel->getActiveSheet();
```
##### Extending the report builder for your customizing needs.
You can extend the default ReportBuilder to add/overwrite (extra) functionality.
```
class myFunkyBuilderClass extends Tactics\Bundle\ExcelBundle\Writer\ReportExcelWriter { ...
```
Existing functionality 
<ol>
<li>get/setActiveCell() (Active cell is coordinate of the cell where content will be written if write method is called)</li>
<li>get/setFirstCell() (First cell is coordinate, pointer will be set to column value when nextRow method is called)</li>
<li>nextCell($amount = 1) (sets active cell to current active cell + $amount * cells to the right)</li>
<li>nextRow($amount = 1)</li>
<li>write(\PHPExcel_Worksheet $sheet, $value) (writes value to a ExcelSheet)</li>
<li>setStyleFromCell(\PHPExcel_Worksheet $sheet, $styledCell, $numberOfColumns = 1, $numberOfRows = 1, $merge = false) (copies style from a cell to a new range)</li>
<li>offerAsXlsxDownload(Response $response, PHPExcel $excel, $filename)</li>
<li>getExcelFromFileName($fileName, $kernel)</li>
<li>nextSheet($sheetTitle, $startcell = 'A1')</li>
<li>getDownloadFromCollection($collection, array $options = array('filename' => 'defaultFileName'))</li>
<li>writeCollection(\PHPExcel_Worksheet $sheet, $collection, $options = array())</li>
</ol>        
##### Customize used helpers for optimal customization.
The ReportExcelWriter depends on helpers to distribute functionality.
Extending/overwriting these helpers allow you to change some behaviors as needed/wanted.

p.e.
class myWriteHelper extends Tactics\Bundle\ExcelBundle\Helpers\WriteHelper {
```
/**
* {@inheritdoc}
*/
       public function write(\PHPExcel_Worksheet $sheet, $value) {
              ...
       }
}
```

//getting the reportwriter with myHelperWriter
see todo why this isn't documented yet

AvailableHelpers
<ol>
<li>CollectionHelper</li>
<li>DownloadHelper</li>
<li>FileReaderHelper</li>
<li>NavigatorHelper</li>
<li>ObjectTransFormerHelper</li>
<li>StylingHelper</li>
<li>WriteHelper</li>
</ol>
## Todo's (aka coming soon ;) )
<ol>
<li>define helper classes in configuration so we don't have to invoke the Reportbuilder constructor when overwriting a helper class.</li>
<li>...</li>
</ol>
