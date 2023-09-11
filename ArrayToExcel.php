<?php

require_once('vendor/autoload.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
 use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ArrayToExcel
{
    public function __construct($array)
    {
        $this->row_num = 1;
        $this->col_num = 1;

        $this->spreadsheet = new Spreadsheet();
        $this->spreadsheet->getDefaultStyle()->getNumberFormat()->setFormatCode('#');

        $sheet = $this->spreadsheet->getActiveSheet();
        $finishedsheet = $this->build($array, $sheet); // χτίζουμε με τον τρόπο του SuiteCRM
        return $finishedsheet;
    }

    private function build($array, $sheet)
    {
        foreach ($array as $rownum => $data) {
            $rownum++;
            if (is_array($data)) {
                foreach ($data as $cellnum => $celldata) {
                    $cellnum++;
                    $this->putDataInCell($sheet, $rownum, $cellnum, $celldata);
                }
            }
        }
        return $sheet; // επιστρέφουμε απευθείας το sheet, αν το κάνω αλλιώς θέλει να φτιάξω getter
    }

    private function walk($array, $sheet)
    {
        foreach ($array as $key => $data) {
            if (!is_array($data)) {
                if ($this->row_num == 1) { // στην πρώτη γραμμή βάζουμε τους τίτλους και στη δεύτερη τα πρώτα δεδομένα
                    $this->putDataInCell($sheet, 1, $this->col_num, $key);
                    $this->putDataInCell($sheet, 2, $this->col_num, $data);
                } else { // από την 3η και μετά κανονικά
                    $sheet->setCellValue([$this->col_num, $this->row_num], $data);
                }
                $this->col_num++;
            } else {
                walk($data, $sheet);
                if ($this->row_num == 1) {
                    $this->row_num = 3;
                } else {
                    $this->row_num++;
                }
                $this->col_num = 1;
            }
        }
        return $sheet;
    }

    public function save($excel_file_path)
    {
        $myFile = fopen($excel_file_path, "w");
        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->spreadsheet, "Xlsx");
        try {
            $writer->save($myFile);
        } catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
            return "problem in saving";
        }
        fclose($myFile);
    }

    private function putDataInCell($sheet, $rownum, $cellnum, $celldata){

        //// prepare cell
        $cell = $sheet->getCellByColumnAndRow($cellnum, $rownum);

        /// prepare data
        $celldata = trim($celldata, '"');
        $celldata = trim($celldata, "'");

        if (stripos($celldata, "€") !== false) {
            $celldata = trim($celldata, "€");
            $celldata = $celldata;
        }

            // if the value is a negative number, check it
        if (substr($celldata, 1, 1) == "-") {
            $celldata = trim($celldata, "-");
            $celldata = trim($celldata, "'");
            $celldata = (float)$celldata;
        }

        // if it has decimals, check if is formatted correctly
        $celldata = trim($celldata, "'");
        $commaMatch = preg_match("/\d\,\d\d/", $celldata);
        if ($commaMatch) {
            $celldata = str_replace('.', '', $celldata);
            $celldata = str_replace(",", ".", $celldata);
        }

        // if a number format it
        if (is_numeric($celldata)) {
            $number = floatval($celldata);
            $cell->setDataType('Number');
            $cell->setValue($number);
            $cordC = $cell->getCoordinate();
            $sheet->getStyle($cordC)->getNumberFormat()->setFormatCode('#,##0.00');
        } else {
            $cell->setValue($celldata);
        }

        // if it is a link
        if (stripos($celldata, "://") !== false) {
            $sheet->getCellByColumnAndRow($cellnum, $rownum)->getHyperlink()->setUrl($celldata);
        } elseif (stripos($celldata, "@") !== false) { // or if it is a mail // εδώ χρειάζεται καλύτερο validation
            $sheet->getCellByColumnAndRow($cellnum, $rownum)->getHyperlink()->setUrl("mailto:" . $celldata);
        }

    }

}
