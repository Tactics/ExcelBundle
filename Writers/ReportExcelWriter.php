<?php

namespace Tactics\Bundle\ExcelBundle\Writers;
use PHPExcel_IOFactory;
use Symfony\Component\HttpFoundation\Response;
use PHPExcel;
use Symfony\Component\Validator\Constraints\DateTime;
use Tactics\Bundle\ExcelBundle\Helpers\CollectionHelper;
use Tactics\Bundle\ExcelBundle\Helpers\DownloadHelper;
use Tactics\Bundle\ExcelBundle\Helpers\FileReaderHelper;
use Tactics\Bundle\ExcelBundle\Helpers\NavigatorHelper;
use Tactics\Bundle\ExcelBundle\Helpers\StylingHelper;
use Tactics\Bundle\ExcelBundle\Helpers\WriteHelper;

class ReportExcelWriter
{
    /**
     * @var string
     */
    protected $active_cell;

    protected $collectionHelper;
    protected $reader;
    protected $downloader;
    protected $navigator;
    protected $styler;
    protected $writer;

    public function __construct(CollectionHelper $collectionHelper, FileReaderHelper $fileReaderHelper, DownloadHelper $downloadHelper, NavigatorHelper $navigatorHelper, StylingHelper $stylingHelper, WriteHelper $writeHelper) {
        $this->collectionHelper = $collectionHelper;
        $this->reader = $fileReaderHelper;
        $this->downloader = $downloadHelper;
        $this->navigator = $navigatorHelper;
        $this->styler = $stylingHelper;
        $this->writer = $writeHelper;

        $this->writer->setExcelWriter($this);
        $this->navigator->setExcelWriter($this);
        $this->styler->setExcelWriter($this);
        $this->collectionHelper->setExcelWriter($this);
    }

    /**
     * @var string
     */
    protected $first_cell;

    const EMPTY_CELL_VALUE = '-';
    const EMPTY_FILTER_CELL_VALUE = 'No filter';

    public function getActiveCell()
    {
        return $this->active_cell;
    }

    public function setActiveCell($activeCell)
    {
        $this->active_cell = $activeCell;

        return $this;
    }

    public function getFirstCell()
    {
        return $this->first_cell;
    }

    public function setFirstCell($firstCell)
    {
        $this->first_cell = $firstCell;

        return $this;
    }

    /**
     * Creates a PHPExcel object from a filename
     *
     * @param $fileName
     * @return PHPExcel
     */
    public function createExcelFromFile($fileName)
    {
        return $this->reader->createExcelFromFile($fileName);
    }

    /**
     * Sets active cell to initial column $times rows lower
     *
     * @param int $times
     */
    public function nextRow($times = 1)
    {
        $this->navigator->nextRow($times);
    }

    /**
     * Sets active cell $times columns to the right
     *
     * @param int $times
     */
    public function nextCell($times = 1)
    {
        $this->navigator->nextCell($times);
    }

    /**
     * Writes the value to the cell and puts the active cell one record further.
     *
     * @param \PHPExcel_Worksheet $sheet
     * @param $value
     */
    public function write(\PHPExcel_Worksheet $sheet, $value)
    {
        $this->writer->write($sheet, $value);
    }

    /**
     * Copies the style of a cell to a range defined by a number of columns/rows relative to active cell.
     *
     * @param \PHPExcel_Worksheet $sheet
     * @param $styledCell
     * @param int $numberOfColumns
     * @param int $numberOfRows
     * @param boolean $merge whether the cells in the range should be merged or not
     */
    protected function setStyleFromCell(\PHPExcel_Worksheet $sheet, $styledCell, $numberOfColumns = 1, $numberOfRows = 1, $merge = false)
    {
        $this->styler->setStyleFromCell($sheet, $styledCell, $numberOfColumns, $numberOfRows, $merge);
    }

    /**
     * Makes a new Response and sets the headers so the excel is offered as a Excel 2007 download
     *
     * @param Response $response
     * @param PHPExcel $excel
     * @param $filename
     */
    public function offerAsXlsxDownload(Response $response, PHPExcel $excel, $filename)
    {
        $this->downloader->offerAsXlsxDownload($response, $excel, $filename);
    }

    /**
     * Creates a PHPExcel object from filename (file becomes te template we write to)
     *
     * @param $fileName
     * @param $kernel
     * @return PHPExcel
     */
    public function getExcelFromFileName($fileName, $kernel)
    {
        //@todo inject kernel service into filereader helper?
        return $this->reader->getExcelFromFileName($fileName, $kernel);
    }


    /**
     * Writes the filtered values to sheet of failure report
     *
     * @param \PHPExcel_Worksheet $sheet
     * @param $begincellGeneral
     * @param $begincellPeriod
     */
    protected function writeFilterValues(\PHPExcel_Worksheet $sheet, $begincellGeneral, $begincellPeriod, $filterValues)
    {
        //Set pointer to begin of filter
        $this->setFirstCell($begincellGeneral);
        $this->setActiveCell($begincellGeneral);

        //Write general filter info
        $this->write($sheet, $filterValues['type'] ? : self::EMPTY_FILTER_CELL_VALUE);
        $this->nextRow();
        $this->write($sheet, $filterValues['department'] ? : self::EMPTY_FILTER_CELL_VALUE);
        $this->nextRow();
        $this->write($sheet, $filterValues['equipment_type'] ? : self::EMPTY_FILTER_CELL_VALUE);
        $this->nextRow();
        $this->write($sheet, $filterValues['area'] ? : self::EMPTY_FILTER_CELL_VALUE);
        $this->nextRow();

        //set pointer to begin of period information of filter
        $this->setActiveCell($begincellPeriod);
        $this->setFirstCell($begincellPeriod);

        // write the period filter cells
        if($filterValues['date_failure_from']) {
            $this->writeFormattedDateToCell($sheet, $this->getActiveCell(), $filterValues['date_failure_from']);
        }
        else {
            $this->write($sheet, self::EMPTY_FILTER_CELL_VALUE);
        }
        $this->nextRow();
        if($filterValues['date_failure_to']) {
            $this->writeFormattedDateToCell($sheet, $this->getActiveCell(), $filterValues['date_failure_to']);
        }
        else {
            $this->write($sheet, self::EMPTY_FILTER_CELL_VALUE);
        }
    }

    public function nextSheet($sheetTitle, $startcell = 'A1')
    {
        $this->navigator->nextSheet($sheetTitle, $startcell);
    }

    //@todo make this happen
    protected function previousSheet($startcell = 'A1')
    {
        $this->navigator->nextSheet($startcell);
    }

    /**
     * Writes an image to the sheet.
     *
     * @param $path full pathname to the image
     * @param \PHPExcel_Worksheet $sheet
     * @param array $options
     */
    protected function writeImageToSheet($path, \PHPExcel_Worksheet $sheet, array $options = array())
    {
        $this->writer->writeImageToSheet($path, $sheet, $options);
    }

    /**
     * Returns a response as download with the collection exported to an excel 2007 file .
     *
     * @param $collection
     * @param array $options
     */
    public function getDownloadFromCollection($collection, array $options = array('filename' => 'defaultFileName'))
    {
        $excel = new PHPExcel();
        $excel->setActiveSheetIndex();
        $sheet = $excel->getActiveSheet();

        $this->collectionHelper->writeCollection($sheet, $collection, $options);
        $this->downloader->offerAsXlsxDownload(new Response(), $excel, $options['filename']);
    }

    public function getDownloadFromCrossRefArray($collection, array $options = array('filename' => 'defaultFileName'))
    {
        $excel = new PHPExcel();
        $excel->setActiveSheetIndex();
        $sheet = $excel->getActiveSheet();

        $this->collectionHelper->writeRefTableFromCollection($sheet, $collection, $options);
        $this->downloader->offerAsXlsxDownload(new Response(), $excel, $options['filename']);
    }

    /**
     * Writes a collection to the sheet.
     *
     * @param \PHPExcel_Worksheet $sheet
     * @param $collection
     * @param array $options
     */
    public function writeCollection(\PHPExcel_Worksheet $sheet, $collection, $options = array())
    {
        $this->collectionHelper->writeCollection($sheet, $collection, $options);
    }
}
