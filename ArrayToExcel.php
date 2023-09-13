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
        $finishedsheet = $this->build($array, $sheet);
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
        return $sheet;
    }

    private function walk($array, $sheet) // recursive
    {
        foreach ($array as $key => $data) {
            if (!is_array($data)) {
                if ($this->row_num == 1) { 
                    $this->putDataInCell($sheet, 1, $this->col_num, $key);
                    $this->putDataInCell($sheet, 2, $this->col_num, $data);
                } else { 
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

        // search for letters and for "+" sign to identify phone numbers, as PhpSpreadsheet searches only for leading zeroes
        if(preg_match("/[a-z]|\+/i", $celldata)){
            $cell->setValueExplicit($celldata, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING );

        // check if it is a link or email
        if (stripos($celldata, "://") !== false) {
            $sheet->getCellByColumnAndRow($cellnum, $rownum)->getHyperlink()->setUrl($celldata);
        } elseif (stripos($celldata, "@") !== false) { // here needs additional validation
            $sheet->getCellByColumnAndRow($cellnum, $rownum)->getHyperlink()->setUrl("mailto:" . $celldata);
        }

        } else { // it is not text or link, we assume it is a number

            // remove currency signs
            if (stripos($celldata, "€") !== false || stripos($celldata, "$") !== false) { // checks only for Euro or Dollar signs, add your currency
                $celldata = trim($celldata, "€$");
                settype($celldata, 'float');
            }

            // change decimal separation sign to default (dot)
            $commaMatch = preg_match("/\d\,\d\d$/", $celldata); // it checks for two digits after a comma on end of the string (greek way), check for your own regional settings if needed
            if ($commaMatch) {
                $celldata = str_replace('.', '', $celldata); // removes dots separating thousands
                $celldata = str_replace(",", ".", $celldata); // replaces comma with default dot
            }

            $cell->setDataType('Number');
            $cell->setValue($number);
            $cordC = $cell->getCoordinate();
            $sheet->getStyle($cordC)->getNumberFormat()->setFormatCode('#,##0.00');
        }
    }
}
