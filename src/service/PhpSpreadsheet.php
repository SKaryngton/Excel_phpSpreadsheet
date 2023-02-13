<?php

namespace App\service;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Psr\Log\LoggerInterface;

class PhpSpreadsheet
{


    public function __construct(private LoggerInterface $logger)
    {
    }

    public function generateAndSaveOnServerSimpleExcelSheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        //Set Cell Value
        $sheet->setCellValue("A1","Hello World!");

        // (D) Save in the   Public Directory

        $writer = new Xlsx($spreadsheet);
        try {
            $writer->save("uploads/hello.xlsx");
        } catch (Exception $e) {
             $this->logger->error(message: $e->getMessage());
        }


    }

    public function generateAndDownloadSimpleExcelSheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        //Set Cell Value
        $sheet->setCellValue("A1","Hello World!");

        // (D) SEND DOWNLOAD HEADERS
     //ob_clean();
    //  ob_start();
        $writer = new Xlsx($spreadsheet);
        header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header("Content-Disposition: attachment;filename=\"hello.xlsx\"");
        header("Cache-Control: max-age=0");
        header("Expires: Fri, 11 Nov 2011 11:11:11 GMT");
        header("Last-Modified: ". gmdate("D, d M Y H:i:s") ." GMT");
        header("Cache-Control: cache, must-revalidate");
        header("Pragma: public");
        $writer->save("php://output");
        exit();
   //  ob_end_flush();
    }

    public function  readASheetCell(){

        $value='';

        //Read File
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        try {
            $spreadsheet = $reader->load("uploads/hello.xlsx");


            //Read Cells

            $sheet=$spreadsheet->getSheet(0);
            $cell = $sheet->getCell("A1");
            $value=$cell->getValue();
        } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception|\PhpOffice\PhpSpreadsheet\Exception $e) {
            $this->logger->error(message: $e->getMessage());
        }

        return $value;

    }


    public function addWorksheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        //Set Cell Value
        $sheet->setCellValue("A1","Hello World!");

        // Set Title
        $sheet->setTitle("First Sheet");

        // ADD WorkSheet
        $spreadsheet->createSheet();
        $sheet= $spreadsheet->getSheet(1);
        $sheet->setTitle("Second Sheet");
        $sheet->setCellValue("A1","New Sheet");

        // (D) Save in the   Public Directory

        $writer = new Xlsx($spreadsheet);
        try {
            $writer->save("uploads/hello.xlsx");
        } catch (Exception $e) {
            $this->logger->error(message: $e->getMessage());
        }

    }

}
