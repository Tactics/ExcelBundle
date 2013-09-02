<?php

namespace Tactics\Bundle\ExcelBundle\Helpers;

use Tactics\Bundle\ExcelBundle\Helpers\ObjectTransformerHelper;
use Tactics\Bundle\ExcelBundle\Helpers\WriteHelper;
use Tactics\Bundle\ExcelBundle\Writers\ReportExcelWriter;

class CollectionHelper {

    protected $objectTransformer;

    protected $writer;

    protected $excelWriter;

    public function __construct(ObjectTransformerHelper $transformer, WriteHelper $writeHelper)
    {
        $this->objectTransformer = $transformer;
        $this->writer = $writeHelper;
    }

    /**
     * Checks if the given collection is an array or can be converted to an array and returns the class of the first element.
     *
     * @param $collection
     * @return mixed
     */
    public function getCollectionObjectClass($collection)
    {
        $collectionArray = $collection;
        //check if collection is array, if not, try to make array out of it.
        if (!is_array($collection) && ! method_exists($collection, 'toArray')) {
            if(! method_exists($collection, 'toArray')) {
                return null;
            }
            else {
                $collectionArray = $collection->toArray();
            }
        }

        //Check if first element in the array is an object.
        if(reset($collectionArray) && is_object(reset($collectionArray))) {
            return get_class(reset($collectionArray));
        }

        return null;
    }

    /**
     * Returns whether we can write the collection (true is all collection items are the same class)
     *
     * @param $collection
     * @return bool
     */
    public function isWritableCollection($collection)
    {
        //@todo what to do with objects from the same superclass?

        $collectionClass = $this->getCollectionObjectClass($collection);
        if ($collectionClass) {
            foreach($collection as $item) {
                if (! is_object($item) && ! get_class($item) == $collectionClass) {
                    return false;
                }
            }
        }

        return true;
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
        if (! $this->isWritableCollection($collection)) {
            return;
        }

        if (isset($options['coordinate'])) {
            $this->excelWriter->setFirstCell($options['coordinate']);
            $this->excelWriter->setActiveCell($options['coordinate']);
        }
        else {
            $this->excelWriter->setFirstCell('A1');
            $this->excelWriter->setActiveCell('A1');
        }

        $headersWritten = false;
        foreach($collection as $collectionItem)
        {
            //Check if the headers need to be written
            if(!$headersWritten && isset($options['include_headers']) && $options['include_headers']) {
                $headers = isset($options['properties']) ? $options['properties'] : null;
                $this->writeCollectionHeaders($sheet, $collectionItem, $headers);
                $headersWritten = true;
            }

            //If its an array, just write the array values
            if (is_array($collectionItem)) {
                foreach($collectionItem as $fieldToWrite) {
                    $this->excelWriter->write($sheet, $fieldToWrite);
                }
            }

            //if the collection contains objects, write the (requested) properties of objects to the sheet
            elseif(is_object($collectionItem)) {
                $propertiesToExport = isset($options['properties']) ? $options['properties'] : null;

                foreach($this->objectTransformer->getWritablePropertiesOfObject($collectionItem, $propertiesToExport) as $propertyName => $propertyValue) {
                    $this->excelWriter->write($sheet, $propertyValue);
                }
            }

            $this->excelWriter->nextRow();
        }
    }


    /**
     * Writes the headers for a collection
     *
     * @param \PHPExcel_Worksheet $sheet
     * @param $object
     * @param null $headers
     */
    private function writeCollectionHeaders(\PHPExcel_Worksheet $sheet, $object, $headers = null)
    {
        if (!$headers) {
            $headers = array_keys($this->objectTransformer->getWritablePropertiesOfObject($object, $headers));
        }

        foreach ($headers as $header) {
            $this->excelWriter->write($sheet, str_replace('_', ' ', ucfirst($header)));
        }

        $this->excelWriter->nextRow();
    }

    public function setExcelWriter(ReportExcelWriter $writer)
    {
        $this->excelWriter = $writer;
    }

    /**
     * Writes a ref table from an array, the array must be build in the following way
     * - column headers must be the indexes of the collection
     * - every of this indexes references an array with the row data
     * - indexes of row data are used as row titles
     *
     * @param \PHPExcel_Worksheet $sheet
     * @param $collection
     * @param $options
     */
    public function writeRefTableFromCollection(\PHPExcel_Worksheet $sheet, $collection, $options)
    {
        $this->excelWriter->setFirstCell('A1');
        $this->excelWriter->setActiveCell('A1');

        $this->writeRefTableColumnHeaders($sheet, $collection);
        $this->writeRefTableRowHeaders($sheet, $collection);

        $this->excelWriter->setFirstCell('B2');
        $this->excelWriter->setActiveCell('B2');
        $this->writeRefValues($sheet, $collection);
    }

    /**
     * Writes the column headers of a ref table based on a collection
     *
     * @param \PHPExcel_Worksheet $sheet
     * @param $collection
     */
    private function writeRefTableColumnHeaders(\PHPExcel_Worksheet $sheet, $collection)
    {
        $this->excelWriter->nextCell();

        foreach($collection as $title => $subArray) {
            $this->excelWriter->write($sheet, $title);
        }
    }

    /**
     * Writes the row headers of a ref table based on a collection
     *
     * @param \PHPExcel_Worksheet $sheet
     * @param $collection
     */
    private function writeRefTableRowHeaders(\PHPExcel_Worksheet $sheet, $collection)
    {
        $this->excelWriter->setActiveCell('A1');
        $this->excelWriter->nextRow();

        foreach($collection as $dataArray) {
            foreach($dataArray as $title => $data) {
                $this->excelWriter->write($sheet, $title);
                $this->excelWriter->nextRow();
            }

            break;
        }
    }

    /**
     * Writes the values in a reference table
     *
     * @param \PHPExcel_Worksheet $
     * @param $collection
     */
    private function writeRefValues(\PHPExcel_Worksheet $sheet, $collection)
    {
        foreach($collection as $subCollection) {
            foreach($subCollection as $value) {
                $this->excelWriter->write($sheet, $value);
                $this->excelWriter->nextRow();
            }

            //Position the cursor to begin writing a new column of data
            $this->excelWriter->setActiveCell($this->excelWriter->getFirstCell());
            $this->excelWriter->nextCell();
            $this->excelWriter->setFirstCell($this->excelWriter->getActiveCell());
        }
    }
}