<?php
/**
 * Created by JetBrains PhpStorm.
 * User: Jeroen
 * Date: 26/04/13
 * Time: 14:46
 * To change this template use File | Settings | File Templates.
 */

namespace Tactics\Bundle\ExcelBundle\Helpers;

use Symfony\Component\HttpFoundation\Response;
use PHPExcel;
class DownloadHelper {
    /**
     * Makes a new Response and sets the headers so the excel is offered as a Excel 2007 download
     *
     * @param Response $response
     * @param PHPExcel $excel
     * @param $filename
     */
    public function offerAsXlsxDownload(Response $response, PHPExcel $excel, $filename)
    {
        $response->headers->set('Content-Type', 'application/vnd.openxmlformats.xlsx');
        $response->headers->set('Content-Disposition', 'attachment; filename="'.$filename.'.xlsx"');
        $response->headers->set('Cache-Control', 'max-age=0');
        $response->sendHeaders();

        $objWriter = \PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
        $objWriter->save('php://output');
    }
}