<?php
/**
 * Created by JetBrains PhpStorm.
 * User: Jeroen
 * Date: 26/04/13
 * Time: 14:47
 * To change this template use File | Settings | File Templates.
 */

namespace Tactics\Bundle\ExcelBundle\Helpers;

use Tactics\Bundle\ExcelBundle\Writers\ReportExcelWriter;

class StylingHelper {

    protected $excelWriter;

    /**
     * Copies the style of a cell to a range defined by a number of columns/rows relative to active cell.
     *
     * @param \PHPExcel_Worksheet $sheet
     * @param $styledCell
     * @param int $numberOfColumns
     * @param int $numberOfRows
     * @param boolean $merge whether the cells in the range should be merged or not
     */
    public function setStyleFromCell(\PHPExcel_Worksheet $sheet, $styledCell, $numberOfColumns = 1, $numberOfRows = 1, $merge = false)
    {
        $column = substr($this->excelWriter->getActiveCell(),0,1);
        $columnCounter = 1;
        while($columnCounter < $numberOfColumns){
            ++$column;
            $columnCounter++;
        }

        $row = substr($this->excelWriter->getActiveCell(), 1);
        $rowCounter = 1;
        while($rowCounter < $numberOfRows) {
            ++$row;
            $rowCounter++;
        }

        $endCell = $column.$row;
        $range = $this->excelWriter->getActiveCell().':'.$endCell;

        if ($merge) {
            $sheet->mergeCells($range);
        }
        $sheet->duplicateStyle($sheet->getStyle($styledCell),$range);
    }

    public function setExcelWriter(ReportExcelWriter $excelWriter)
    {
        $this->excelWriter = $excelWriter;
    }
}