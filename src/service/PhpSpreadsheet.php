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

        // ADD new WorkSheet
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


    public function copyWorksheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        //Set Cell Value
        $sheet->setCellValue("A1","Hello World!");

        // Set Title
        $sheet->setTitle("First Sheet");

        // COPY WorkSheet and add worksheet
        $copy=clone $spreadsheet->getSheetByName("First Sheet");
        $copy->setTitle("Copy Sheet");
        $spreadsheet->addSheet($copy);

        // (D) Save in the   Public Directory

        $writer = new Xlsx($spreadsheet);
        try {
            $writer->save("uploads/hello.xlsx");
        } catch (Exception $e) {
            $this->logger->error(message: $e->getMessage());
        }

    }


    public function deleteWorksheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        //Set Cell Value
        $sheet->setCellValue("A1","Hello World!");

        // Set Title
        $sheet->setTitle("First Sheet");

        // COPY WorkSheet and add worksheet
        $copy=clone $spreadsheet->getSheetByName("First Sheet");
        $copy->setTitle("Copy Sheet");
        $spreadsheet->addSheet($copy);

        //Delete CopySheet
        $spreadsheet->removeSheetByIndex(1);


        // (D) Save in the   Public Directory

        $writer = new Xlsx($spreadsheet);
        try {
            $writer->save("uploads/hello.xlsx");
        } catch (Exception $e) {
            $this->logger->error(message: $e->getMessage());
        }

    }

    public function countWorksheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        //Set Cell Value
        $sheet->setCellValue("A1","Hello World!");

        // Set Title
        $sheet->setTitle("First Sheet");

        // COPY WorkSheet and add worksheet
        $copy=clone $spreadsheet->getSheetByName("First Sheet");
        $copy->setTitle("Copy Sheet");
        $spreadsheet->addSheet($copy);

        //GET TOTAL NUMBER OF WORKSHEETS
        $total=$spreadsheet->getSheetCount();

        return $total;



    }

    public function SetCellValueWorksheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        //Set Cell Value
        $sheet->setCellValue("A1","Hello World!");

        // Set Title
        $sheet->setTitle("First Sheet");



        //GET Single Cell then Set Value
        $cell=$sheet->getCell("A1");
        $cell->setValue("Hello");

        //GET Single Cell then get Value
        $cell=$sheet->getCell("A1");
        $val=$cell->getValue("Hello");



        // (D) Save in the   Public Directory

        $writer = new Xlsx($spreadsheet);
        try {
            $writer->save("uploads/hello.xlsx");
        } catch (Exception $e) {
            $this->logger->error(message: $e->getMessage());
        }

    }


    public function getHighestRowAndColumnUSeLoopWorksheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        //Set Cell Value
        $sheet->setCellValue("A1","Hello World!");

        // Set Title
        $sheet->setTitle("First Sheet");



        // (C4) GET HIGHEST ROW + COL
        $highestRow = $sheet->getHighestRow();
        $highestCol = $sheet->getHighestColumn();

        // TIP - You can use $highestRow $highestCol to loop through the cells.
        // for ($i=0; i<$highest; i++) { ... }

    }



    public function selectRangeOfCellsWorksheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();



        // Set Title
        $sheet->setTitle("First Sheet");



        //Range Of cells
       // $data=$sheet->rangeToArray("A1:A3");

       //SET DATA FROM ARRAY INTO CELLS
        $data = [100, 53, 86];
        $data = array_chunk($data, 1);
        $sheet->fromArray($data, null, "B1");

        // (D) Save in the   Public Directory

        $writer = new Xlsx($spreadsheet);
        try {
            $writer->save("uploads/hello.xlsx");
        } catch (Exception $e) {
            $this->logger->error(message: $e->getMessage());
        }

    }

    public function formulasWorksheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();



        // Set Title
        $sheet->setTitle("First Sheet");


        //SET DATA FROM ARRAY INTO CELLS
        $data = [100, 53, 86];
        $data = array_chunk($data, 1);
        $sheet->fromArray($data, null, "B1");

        // (E) FORMULAS ACCEPTED - JUST AS IN EXCEL
        $sheet->setCellValue("B4", "=SUM(B1:B3)");

        // (D) Save in the   Public Directory

        $writer = new Xlsx($spreadsheet);
        try {
            $writer->save("uploads/hello.xlsx");
        } catch (Exception $e) {
            $this->logger->error(message: $e->getMessage());
        }

    }


    public function mergeAndUnmergeWorksheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();



        // Set Title
        $sheet->setTitle("First Sheet");


        $sheet->setCellValue("A1", "Hello, this is a very very long string.");
        $sheet->setCellValue("A2", "World!");
        $sheet->setCellValue("A3", "Foo");
        $sheet->setCellValue("A4", "Bar");


        // (C) MERGE & UNMERGE CELLS
        $sheet->mergeCells("A1:D1");
        $sheet->mergeCells("A2:B2");
        //$sheet->unmergeCells("A2:B2");

        // (D) Save in the   Public Directory

        $writer = new Xlsx($spreadsheet);
        try {
            $writer->save("uploads/hello.xlsx");
        } catch (Exception $e) {
            $this->logger->error(message: $e->getMessage());
        }

    }

    public function insertRowAndColumnWorksheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();



        // Set Title
        $sheet->setTitle("First Sheet");


        $sheet->setCellValue("A1", "Hello, this is a very very long string.");
        $sheet->setCellValue("A2", "World!");
        $sheet->setCellValue("A3", "Foo");
        $sheet->setCellValue("A4", "Bar");


        // (C) MERGE & UNMERGE CELLS
        $sheet->mergeCells("A1:D1");
        $sheet->mergeCells("A2:B2");
        //$sheet->unmergeCells("A2:B2");

        // (D) Save in the   Public Directory

        // (D) INSERT ROW & COL
        $sheet->insertNewColumnBefore("A", 1); // 1 new column before column A
        $sheet->insertNewRowBefore(3, 1); // 1 new row before row 3

        $writer = new Xlsx($spreadsheet);
        try {
            $writer->save("uploads/hello.xlsx");
        } catch (Exception $e) {
            $this->logger->error(message: $e->getMessage());
        }

    }

    public function visbilityWorksheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();



        // Set Title
        $sheet->setTitle("First Sheet");


        $sheet->setCellValue("A1", "Hello, this is a very very long string.");
        $sheet->setCellValue("A2", "World!");
        $sheet->setCellValue("A3", "Foo");
        $sheet->setCellValue("A4", "Bar");


        // (C) MERGE & UNMERGE CELLS
        $sheet->mergeCells("A1:D1");
        $sheet->mergeCells("A2:B2");
        //$sheet->unmergeCells("A2:B2");

        // (D) Save in the   Public Directory

        // (D) INSERT ROW & COL
        $sheet->insertNewColumnBefore("A", 1); // 1 new column before column A
        $sheet->insertNewRowBefore(3, 1); // 1 new row before row 3

        // (E) VISIBILITY
        $sheet->getColumnDimension("A")->setVisible(false);
        // $sheet->getColumnDimension("A")->setVisible(true);
        $sheet->getRowDimension(4)->setVisible(false);
       // $sheet->getRowDimension(4)->setVisible(true);

        $writer = new Xlsx($spreadsheet);
        try {
            $writer->save("uploads/hello.xlsx");
        } catch (Exception $e) {
            $this->logger->error(message: $e->getMessage());
        }

    }

    public function withAndHeightCellsWorksheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();



        // Set Title
        $sheet->setTitle("First Sheet");


        $sheet->setCellValue("A1", "Hello, this is a very very long string.");
        $sheet->setCellValue("A2", "World!");
        $sheet->setCellValue("A3", "Foo");
        $sheet->setCellValue("A4", "Bar");


        // (C) MERGE & UNMERGE CELLS
        $sheet->mergeCells("A1:D1");
        $sheet->mergeCells("A2:B2");
        //$sheet->unmergeCells("A2:B2");

        // (D) Save in the   Public Directory

        // (D) INSERT ROW & COL
        $sheet->insertNewColumnBefore("A", 1); // 1 new column before column A
        $sheet->insertNewRowBefore(3, 1); // 1 new row before row 3

        // (E) VISIBILITY
        $sheet->getColumnDimension("A")->setVisible(false);
        // $sheet->getColumnDimension("A")->setVisible(true);
        $sheet->getRowDimension(4)->setVisible(false);
        // $sheet->getRowDimension(4)->setVisible(true);

        // (F) WIDTH & HEIGHT
        $sheet->getRowDimension("4")->setRowHeight(100);
        $sheet->getColumnDimension("C")->setWidth(100);


        $writer = new Xlsx($spreadsheet);
        try {
            $writer->save("uploads/hello.xlsx");
        } catch (Exception $e) {
            $this->logger->error(message: $e->getMessage());
        }

    }

    public function styleFontWorksheet(){

        //Create A new Spreadsheet
        $spreadsheet= new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();



        // Set Title
        $sheet->setTitle("First Sheet");


        $sheet->setCellValue("B1", "Hello");
        $sheet->setCellValue("B2", "World!");
        $sheet->setCellValue("B3", "Foo");
        $sheet->setCellValue("B4", "Bar");
        $sheet->getRowDimension("3")->setRowHeight(50);

// (C) SET STYLE
        $styleSet = [
            // (C1) FONT
            "font" => [
                "bold" => true,
                "italic" => true,
                "underline" => true,
                "strikethrough" => true,
                "color" => ["argb" => "FFFF0000"],
                "name" => "Cooper Hewitt", //schriftart
                "size" => 22
            ],

            // (C2) ALIGNMENT
            "alignment" => [
                "horizontal" => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
                // \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT
                // \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER
                "vertical" => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM
                // \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP
                // \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER
            ],
            // (C3) BORDER
            "borders" => [
                "top" => [
                    "borderStyle" => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                    "color" => ["argb" => "FFFF0000"]
                ],
                "bottom" => [
                    "borderStyle" => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                    "color" => ["argb" => "FF00FF00"]
                ],
                "left" => [
                    "borderStyle" => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
                    "color" => ["argb" => "FF0000FF"]
                ],
                "right" => [
                    "borderStyle" => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    "color" => ["argb" => "FF0000FF"]
                ]
                /* ALTERNATIVELY, THIS WILL SET ALL
                "outline" => [
                  "borderStyle" => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                  "color" => ["argb" => "FFFF0000"]
                ]*/
            ],
            // (C4) FILL
            "fill" => [
                // SOLID FILL
                "fillType" => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                "color" => ["argb" => "FF110000"]

                /* GRADIENT FILL
                "fillType" => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR,
                "rotation" => 90,
                "startColor" => [
                  "argb" => "FF000000",
                ],
                "endColor" => [
                  "argb" => "FFFFFFFF",
                ]*/
            ]
            ];

        $style = $sheet->getStyle("B3");
        // $style = $sheet->getStyle("B1:B4");
        $style->applyFromArray($styleSet);

        $writer = new Xlsx($spreadsheet);
        try {
            $writer->save("uploads/hello.xlsx");
        } catch (Exception $e) {
            $this->logger->error(message: $e->getMessage());
        }

    }
}
