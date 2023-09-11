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

    private function walk($array, $sheet)
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

        if(preg_match("/[a-z]|\+/i", $celldata)){
            $cell->setValueExplicit($celldata, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING );

        } else {
            if (stripos($celldata, "€") !== false || stripos($celldata, "$") !== false) {
                $celldata = trim($celldata, "€$");
                settype($celldata, 'float');
            }

            $commaMatch = preg_match("/\d\,\d\d/", $celldata);
            if ($commaMatch) {
                $celldata = str_replace('.', '', $celldata);
                $celldata = str_replace(",", ".", $celldata);
            }

            if (is_numeric($celldata)) {
                $number = floatval($celldata);
                $cell->setDataType('Number');
                $cell->setValue($number);
                $cordC = $cell->getCoordinate();
                $sheet->getStyle($cordC)->getNumberFormat()->setFormatCode('#,##0.00');
            }
        }

        // if it is a link
        if (stripos($celldata, "://") !== false) {
            $sheet->getCellByColumnAndRow($cellnum, $rownum)->getHyperlink()->setUrl($celldata);
        } elseif (stripos($celldata, "@") !== false) { // here needs additional validation
            $sheet->getCellByColumnAndRow($cellnum, $rownum)->getHyperlink()->setUrl("mailto:" . $celldata);
        }
    }

}
