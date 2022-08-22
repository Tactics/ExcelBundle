<?php
namespace Tactics\Bundle\ExcelBundle\Helpers;

use Tactics\Bundle\ExcelBundle\Writers\BaseReportExcelWriter;

class WriteHelper {
    protected $excelWriter;

    /**
     * Formats and writes a DateTime instance to a cell.
     *
     * @param \PHPExcel_Worksheet $sheet The sheet to write to.
     * @param string              $cell  The cell to write to.
     * @param \DateTime           $date  The date to format and write.
     */
    protected function writeFormattedDateToCell(\PHPExcel_Worksheet $sheet, $cell, \DateTime $date = null, $includeTime = false)
    {
        if ($date) {
            $sheet->setCellValue($cell, $includeTime ? $date->format('d/m/Y H:i') : $date->format('d/m/Y'));
        }
    }

    /**
     * Writes the value to the cell and puts the active cell one record further.
     *
     * @param \PHPExcel_Worksheet $sheet
     * @param $value
     */
    public function write(\PHPExcel_Worksheet $sheet, $value)
    {
        if($this->isWritableField($value)) {
            if ($value instanceof \DateTime) {
                $this->writeFormattedDateToCell($sheet, $this->excelWriter->getActiveCell(), $value);
            }
            elseif($value === null) {
            }
            else {
                $sheet->setCellValueExplicit($this->excelWriter->getActiveCell(),$value);
            }
        }
        $this->excelWriter->nextCell();
    }

    /**
     * Writes an image to the sheet.
     *
     * @param $path full pathname to the image
     * @param \PHPExcel_Worksheet $sheet
     * @param array $options
     */
    public function writeImageToSheet($path, \PHPExcel_Worksheet $sheet, array $options = array())
    {
        if(!file_exists($path)) {
            return;
        }

        $drawing = new \PHPExcel_Worksheet_Drawing();

        if (isset($options['name'])){
            $drawing->setName($options['name']);
        }
        if (isset($options['description'])) {
            $drawing->setDescription($options['description']);
        }
        if (isset($options['height'])) {
            $drawing->setResizeProportional(false);
            $drawing->setHeight($options['height']);
        }
        if (isset($options['width'])){
            $drawing->setResizeProportional(false);
            $drawing->setWidth($options['width']);
        }

        //@todo offsets + review other possibilities

        $drawing->setPath($path);
        $drawing->setWorksheet($sheet);
        $drawing->setCoordinates(isset($options['coordinate']) ? $options['coordinate'] : $this->getActiveCell());
    }

    /**
     * £Checks if a certain value is writable to the sheet, writebale objects :
     * - null
     * - numeric values
     * - strings
     * - objects with the __toString method
     * - DateTime objectsµ
     *
     * @param $field
     * @return bool
     */
    protected function isWritableField($field)
    {
        return (is_numeric($field) || $field === null || is_string($field) || method_exists($field, '__toString') || $field instanceof \DateTime);
    }

    public function setExcelWriter(BaseReportExcelWriter $excelWriter)
    {
        $this->excelWriter = $excelWriter;
    }
}