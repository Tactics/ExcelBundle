<?php
namespace Tactics\Bundle\ExcelBundle\Helpers;
use PHPExcel_IOFactory;

class FileReaderHelper {
    /**
     * Creates a PHPExcel object from a filename
     *
     * @param $fileName
     * @return PHPExcel
     */
    public function createExcelFromFile($fileName)
    {
        $excelReader = PHPExcel_IOFactory::createReader('Excel2007');
        $excel = $excelReader->load($fileName);

        return $excel;
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
        $pathname = $kernel->getRootDir(). DIRECTORY_SEPARATOR .'..'. DIRECTORY_SEPARATOR . 'web'.DIRECTORY_SEPARATOR. $fileName;
        // replace '/' and '\' by directory seperator
        $pathname = str_replace('/',DIRECTORY_SEPARATOR,$pathname);
        $pathname = str_replace('\\', DIRECTORY_SEPARATOR, $pathname);

        return  $this->createExcelFromFile($pathname);
    }
}