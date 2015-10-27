<?php

/**
 * @copyright Copyright Victor Demin, 2015
 * @license https://github.com/ruskid/yii2-excel-exporter/LICENSE
 * @link https://github.com/ruskid/yii2-excel-exporter#README
 */

namespace ruskid\csvexporter;

use Exception;

/**
 * Little export helper for yii2
 * @author Victor Demin <demin@trabeja.com>
 */
class CSVExporter {

    /**
     * All row style
     * @var array 
     */
    public $style = [];

    /**
     * Headers row style
     * @var array 
     */
    public $headerStyle = ['font' => ['bold' => true]];

    /**
     * Excel filename
     * @var string 
     */
    public $filename = 'results';

    /**
     * @var \PHPExcel 
     */
    private $_phpexcel;

    /**
     * @return \PHPExcel
     */
    public function getPHPExcel() {
        return $this->_phpexcel;
    }

    /**
     * add PHPExcel dependency
     */
    public function __construct() {
        $this->_phpexcel = new \PHPExcel;
    }

    /**
     * Will set single cell value
     * @param string $cellIndex
     * @param string $value
     * @param integer $sheetIndex
     */
    public function setCellValue($cellIndex, $value, $sheetIndex = 0) {
        $phpexcel = $this->getPHPExcel();
        $sheet = $phpexcel->getSheet($sheetIndex);
        $sheet->setCellValue($cellIndex, $value);
    }

    /**
     * Will set cell style from array
     * @param string $cellIndex
     * @param array $style
     */
    public function setCellStyleFromArray($cellIndex, $style, $sheetIndex = 0) {
        $phpexcel = $this->getPHPExcel();
        $sheet = $phpexcel->getSheet($sheetIndex);
        $sheet->getStyle($cellIndex)->applyFromArray($style);
    }

    /**
     * Will set multiple models to excel.
     * @param string $cellIndexStart
     * @param yii\db\ActiveRecord[]|array $models
     * @param array $config
     * @param integer $sheetIndex
     * @return string Next Free Coordinate
     */
    public function setCellData($cellIndexStart, $models, $config, $sheetIndex = 0) {
        $letters = preg_replace("/[^A-Z]+/", "", $cellIndexStart); //start letter
        $numbers = preg_replace("/[^0-9]+/", "", $cellIndexStart); //start number
        //Set row headers
        $tempLetter = $letters; //reset letter
        foreach ($config as $attributeConfig) {
            $header = isset($attributeConfig['header']) ? $attributeConfig['header'] : "";
            $this->setCellValue($tempLetter . $numbers, $header);
            $this->setCellStyleFromArray($tempLetter . $numbers, $this->headerStyle);
            $tempLetter++;
        }

        //Set row values
        $tempLetter = $letters; //reset letter
        $numbers++; //set rows on new line after headers
        foreach ($models as $model) {
            foreach ($config as $attributeConfig) {
                $value = call_user_func($attributeConfig['value'], $model);
                $this->setCellValue($tempLetter . $numbers, $value);
                if (!empty($this->style)) {
                    $this->setCellStyleFromArray($tempLetter . $numbers, $this->style);
                }
                $tempLetter++;
            }
            $numbers++;
            $tempLetter = $letters;
        }
        return $tempLetter . $numbers;
    }

    /**
     * Will load file into the browser. TODO strategies and formats
     * @return file
     */
    public function export() {
        return $this->sendFileToBrowser();
    }

    /**
     * TODO strategies and formats
     * Will send file to the client's browser
     */
    private function sendFileToBrowser() {
        // ** Redirect output to a client’s web browser (Excel2007)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $this->filename . '.xlsx"');
        header("Pragma: "); //IE8 quick fix.
        header("Cache-Control: "); //IE8 quick fix.

        $objWriter = \PHPExcel_IOFactory::createWriter($this->getPHPExcel(), 'Excel2007');
        $objWriter->save('php://output');
        \Yii::$app->end();
    }

}
