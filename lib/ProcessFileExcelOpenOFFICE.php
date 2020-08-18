<?php

// được tạo bởi lê văn thắng có gì hỏi nó :)
require_once ('./lib/spout/src/Spout/Autoloader/autoload.php');

use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;

class ProcessFileExcelOpenOFFICE {

    /**
     * Đọc data từ file excel.
     * hàm này excel 300.000 là 40s trả về mãng 2 chiều value
     * Ex:
     * array(
     *  0=>array(0,1,3,6)
     *  1=>array(2,1,4,6)
     * )
     * @param type $fileName | path file excel
     * @return type $arrayData | Array data 2 dimension
     */
    public function ReadDataToArr($fileName) {
        $reader = ReaderEntityFactory::createXLSXReader();
        $reader->open($fileName);
        $cells = array();
        foreach ($reader->getSheetIterator() as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                $cells[] = $row->toArray();
            }
        }
        $reader->close();
        return $cells;
    }

     /**
     * Đọc data từ file excel.
     * hàm này excel 300.000 là 40s trả về mãng 1 chiều value
     * Ex:
     * array(
     *  0=>array(0,1,3,6)
     *  1=>array(2,1,4,6)
     * )
     * @param type $fileName | path file excel
     * @return type $arrayData | Array data 2 dimension
     */
    public function ReadDataToArrOne($fileName) {
        $reader = ReaderEntityFactory::createXLSXReader();
        $reader->open($fileName);
        $cells = array();
        foreach ($reader->getSheetIterator() as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                $cells[] = $row->toArray()[0];
            }
        }
        $reader->close();
        return $cells;
    }

    /**
     * 
     * @param dường dẩn file + 1 màng 1 chiều có 1 sdt
     * @return type xuất excel
     */
    public function ArrDataToFile($pathFileExcel, $dataTmp) {
        $writer = WriterEntityFactory::createXLSXWriter();
        $writer->openToFile($pathFileExcel); // write data to a file or to a PHP streamco
        foreach ($dataTmp as $key => $value) {
            if ($value != '') {
                $rowFromValues = WriterEntityFactory::createRowFromArray(array('key' => $value));
                $writer->addRow($rowFromValues);
            }
        }

        $writer->close();
    }

    /**
     * 
     * @param dường dẩn file + 1 màng 2 chiều 
     * @return type xuất excel
     */
    public function DimensionArrDataToFile($pathFileExcel, $dataTmp) {
        $writer = WriterEntityFactory::createXLSXWriter();
        $writer->openToFile($pathFileExcel); // write data to a file or to a PHP streamco
        foreach ($dataTmp as $key => $value) {
            unset($value['_id']);
            if ($value != '') {
                $rowFromValues = WriterEntityFactory::createRowFromArray($value);
                $writer->addRow($rowFromValues);
            }
        }

        $writer->close();
    }

}
