<?php
/**
 * Created by JetBrains PhpStorm.
 * User: Jeroen
 * Date: 26/04/13
 * Time: 14:45
 * To change this template use File | Settings | File Templates.
 */

namespace Tactics\Bundle\ExcelBundle\Helpers;
use Tactics\Bundle\ExcelBundle\Writers\ReportExcelWriter;

class NavigatorHelper {
    protected $excelWriter;

    /**
     * Sets active cell to initial column $times rows lower
     *
     * @param int $times
     */
    public function nextRow($times = 1)
    {
        do {
            $rowIndex = substr($this->excelWriter->getActiveCell(),1);
            $replacement = ++$rowIndex;
            //up the row
            $this->excelWriter->setActiveCell(str_replace(substr($this->excelWriter->getActiveCell(),1), $replacement, $this->excelWriter->getActiveCell()));
            //reset column
            $firstColumnIndex = substr($this->excelWriter->getFirstCell(),0,1);
            $this->excelWriter->setActiveCell(str_replace(substr($this->excelWriter->getActiveCell(),0,1),$firstColumnIndex, $this->excelWriter->getActiveCell()));
            $times--;
        }while($times >= 1);
    }

    /**
     * Sets active cell $times columns to the right
     *
     * @param int $times
     */
    public function nextCell($times = 1)
    {
        do{
            $columnIndex = substr($this->excelWriter->getActiveCell(),0,1);
            $replacement = ++$columnIndex;
            //up the column
            $this->excelWriter->setActiveCell(str_replace(substr($this->excelWriter->getActiveCell(),0,1), $replacement, $this->excelWriter->getActiveCell()));
            $times--;
        }while($times >= 1);
    }

    //@todo make this
    public function nextSheet($sheetTitle, $startcell = 'A1')
    {

    }

    //@todo make this happen
    public function previousSheet($startcell = 'A1')
    {

    }

    public function setExcelWriter(ReportExcelWriter $excelWriter)
    {
        $this->excelWriter = $excelWriter;
    }
}