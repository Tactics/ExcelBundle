ExcelBundle
===========
## Getting started
###### Install the bundle using composer
Add following lines to composer.json 
       {
       "type": "git",
       "url": "git@github.com:Tactics/ExcelBundle"
       }
       
       "tactics/excel-bundle": "1.0.*" 
     

## Basic usage
Getting the excelbuilder service is as easy as:
          $builder = $this->get('tactics.excel_writer');
## Features
##### Exporting collections
###### Exporting array of values
Write an array of values to an excel file
          //Make array with some kind of values 
          $collection = array("First value to write", "Second value to write", ...);
          //Offers the excel with values as a file download with filename = awesomeness.xlsx
          $builder->getDownloadFromCollection($collection, array(
              'filename' => 'awesomeness'                     
          ));
###### Exporting entity collection
Write a collection of objects to an excel file 
          //get a collection of objects
          $collection = $this->getDoctrine()->getManager()->getRepository('SomeBundle:SomeEntity')->findAll();
          //Offers the excel with values as a file download with filename = awesomeness.xlsx
          $builder->getDownloadFromCollection($collection, array(
              'filename' => 'awesomeness',       
           ));
##### Getting excel file from excel template.
In stead of writing to an empty excel file you can create a template excel file and write to this file.
          //$builder = valid ReportExcelWriter.
          $excel = $builder->getExcelFromFile('my_templates_file_name', kernel_service);
          $builder->writeToSheet($excel->getActiveSheet();
##### Extending the report builder for your customizing needs.
You can extend the default ReportBuilder to add/overwrite (extra) functionality.
          class myFunkyBuilderClass extends Tactics\Bundle\ExcelBundle\Writer\ReportExcelWriter { ...
Existing functionality 
1. get/setActiveCell() (Active cell is coordinate of the cell where content will be written if write method is called)
2. get/setFirstCell() (First cell is coordinate, pointer will be set to column value when nextRow method is called)
3. nextCell($amount = 1) (sets active cell to current active cell + $amount * cells to the right)
4. nextRow($amount = 1)
5. write(\PHPExcel_Worksheet $sheet, $value) (writes value to a ExcelSheet)
6. setStyleFromCell(\PHPExcel_Worksheet $sheet, $styledCell, $numberOfColumns = 1, $numberOfRows = 1, $merge = false) (copies style from a cell to a new range)
7. offerAsXlsxDownload(Response $response, PHPExcel $excel, $filename)
8. getExcelFromFileName($fileName, $kernel)
9. nextSheet($sheetTitle, $startcell = 'A1')
10. getDownloadFromCollection($collection, array $options = array('filename' => 'defaultFileName'))
11. writeCollection(\PHPExcel_Worksheet $sheet, $collection, $options = array())
        
##### Customize used helpers for optimal customization.
The ReportExcelWriter depends on helpers to distribute functionality.
Extending/overwriting these helpers allow you to change some behaviors as needed/wanted.

p.e.
          class myWriteHelper extends Tactics\Bundle\ExcelBundle\Helpers\WriteHelper {
              /**
              * {@inheritdoc}
              */
              public function write(\PHPExcel_Worksheet $sheet, $value) {
                  ...
              }
          }
        
          //getting the reportwriter with myHelperWriter
          see todo why this isn't documented yet

AvailableHelpers
1. CollectionHelper
2. DownloadHelper
3. FileReaderHelper
4. NavigatorHelper
5. ObjectTransFormerHelper
6. StylingHelper
7. WriteHelper

## Todo's (aka coming soon ;) )
1. define helper classes in configuration so we don't have to invoke the Reportbuilder constructor when overwriting a helper class.
2. ...
