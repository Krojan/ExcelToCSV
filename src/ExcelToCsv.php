<?php

namespace App;

class ExcelToCsv{
    private $file;
    private $exportPath = "public/export/";
    public function __construct($filepath){
        $this->getFiles($filepath);
    }
    public function convert(){
        $filename = pathinfo( $this->file )['filename'];
        $inputFileType = \PhpOffice\PhpSpreadsheet\IOFactory::identify($this->file);
        $objReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($inputFileType)->load($this->file);
        $objWriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($objReader, 'Csv');
        $filename = $this->manipulateFilename($filename);

        $objReaderActiveSheet = $objReader->getActiveSheet();
        $lastRow = $objReaderActiveSheet->getHighestDataRow();
        $objReaderActiveSheet->removeRow($lastRow);
        $objWriter->save($this->exportPath. $filename .".csv");
        $exportFilename =  $filename .".csv";
        echo $exportFilename . " created successfully at " .$this->exportPath. PHP_EOL ;
    }

    public function getFiles($filepath){
        foreach( glob( "$filepath/*.xlsx" ) as $filename ){
            $this->file = $filename;
            $this->convert( $this->file );
        }
    }

    public function manipulateFilename( $filename ){
        return preg_replace("/\s+/", "-", strtolower($filename));    
    }
}
