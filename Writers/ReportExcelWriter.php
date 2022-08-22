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

class ReportExcelWriter extends BaseReportExcelWriter
{
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
      $this->write($sheet, $this->getFilterValue($filterValues, 'type'));
      $this->nextRow();
      $this->write($sheet, $this->getFilterValue($filterValues, 'department'));
      $this->nextRow();
      $this->write($sheet, $this->getFilterValue($filterValues, 'equipment_type'));
      $this->nextRow();
      $this->write($sheet, $this->getFilterValue($filterValues, 'area'));
      $this->nextRow();

      //set pointer to begin of period information of filter
      $this->setActiveCell($begincellPeriod);
      $this->setFirstCell($begincellPeriod);

      // write the period filter cells
      $this->write($sheet, $this->getFilterValue($filterValues, 'date_failure_from'));
      $this->nextRow();
      $this->write($sheet, $this->getFilterValue($filterValues, 'date_failure_to'));
    }

    private function getFilterValue($filterValues, $index) {
      if(!isset($filterValues[$index])) {
        return self::EMPTY_FILTER_CELL_VALUE;
      }

      return $filterValues[$index] ? : self::EMPTY_FILTER_CELL_VALUE;
    }

}
