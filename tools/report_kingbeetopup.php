<?php

ini_set('memory_limit', -1);
ini_set('max_execution_time', 0);
ini_set("display_errors", 1);
ini_set("display_startup_errors", 1);
error_reporting(E_ALL);

require_once './lib/MyMongoDriver.php';
require_once './lib/ProcessFileExcelOpenOFFICE.php';

use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Writer\AbstractMultiSheetsWriter;

class report
{
    private $_svExcelOFi;
    private $_pathFileData;
    private $_mongo_db;

    public function __construct()
    {
        $this->_svExcelOFi = new ProcessFileExcelOpenOFFICE();
        $this->_pathFileData = '/var/www/html/'; // Direction folder export
        $configDatabase = array(
            'mongo_host' => "localhost",
            'mongo_port' => 27018,
            'mongo_db' => "",
            'mongo_user' => "",
            'mongo_pass' => ""
        );
        $this->_mongo_db = new MyMongoDriver($configDatabase, null, null);
    }

    public function export()
    {
        try {
            $ArrSelect   = array(); // field
            $arrData = $this->_mongo_db->where(array())->select($ArrSelect)->get(); // ten bang

            $FileName = 'Baocao' . time() . '.xlsx'; // file name
            $nameFile = $this->_pathFileData . $FileName;
            $writer = WriterEntityFactory::createXLSXWriter();
            $writer->setShouldUseInlineStrings(true);
            $writer->setShouldCreateNewSheetsAutomatically(true);
            $writer->openToFile($nameFile); // write 

            $values = []; // header
            $rowFromValues = WriterEntityFactory::createRowFromArray($values);
            $writer->addRow($rowFromValues);
            $multipleRows = array();
            foreach ($arrData as $value) {

                $cells = [
                    WriterEntityFactory::createCell(isset($value['']) ? $value[''] : 0) // data
                ];
                $multipleRows[] = WriterEntityFactory::createRow($cells);
            }
            $writer->addRows($multipleRows); // write arr to file
            $writer->close();
        } catch (Exception $e) {
            echo $e->getMessage();
        }
    }
}

$a = new report();
$a->export();
