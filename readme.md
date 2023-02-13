<style> 
green { color: #299660} 
yel { color: #9ea647} 
blue { color: #099fc0} 
red {color: #ce4141} 
</style>

# <green> PhpSpreadsheet

- <yel>install PhpSpreadsheet
     - <blue>composer require phpoffice/phpspreadsheet
- <yel>create a new Spreadsheet
    - create and save On server
      ```
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
      ```
    - create and download
     ```
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
  
     ```
  - read xlxs file
   ```
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
   ```
  - add new Worksheet
    ```
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
    ```
  - copy and add Worsheet
   ```
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
   ```
  - Delete Worksheet
    ```
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
    ```
  - Count Worksheets
   ```
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
   ```
  - Get single cell then set value
   ```
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
   
   ```
  - Highest Row and Column For loop
    ```
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
    ```
  - SET DATA FROM ARRAY INTO CELLS
    ```
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
    ```
  - Formulas 
   ```
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
   ```
  - MERGE & UNMERGE CELLS
   ```
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
   ```
  - Insert Row & Col
   ```
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
  
   ```
  - Visibility
   ```
  
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
  
   ```
  - Width & Height
   ```
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
   
   ```
  - 
