<?php

/**
 * @copyright Copyright Victor Demin, 2015
 * @license https://github.com/ruskid/yii2-excel-exporter/LICENSE
 * @link https://github.com/ruskid/yii2-excel-exporter#README
 */

namespace ruskid\csvexporter;

use Exception;

/**
 * Little Yii2 Helper that exports Active Records to CSV
 * @author Victor Demin <demin@trabeja.com>
 */
class CSVExporter extends \PHPExcel {

    /**
     * Active Records for export
     * @var \yii\db\ActiveRecord 
     */
    public $models;

    /**
     * @var array
     */
    public $values;

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
     * Start point for export X|Y
     * @var string 
     */
    public $startPoint = 'A1';
    private $_letters = [];
    private $_currentRowNumber;

    /**
     * PHPExcel Settings
     */
    public function __construct() {
        parent::__construct();
        $this->setActiveSheetIndex(0); //Use only one sheet
    }

    /**
     * Will validate that required parameters are present
     * @throws Exception
     */
    private function validateInput() {
        if (!$this->models) {
            throw new Exception("'models' parameter is required in " . __CLASS__);
        }
        if (!$this->values) {
            throw new Exception("'values' parameter is required in " . __CLASS__);
        }
    }

    /**
     * Will generate letters by calculating start point X axe and number of values
     * @return array
     */
    public function getLetters() {
        if (empty($this->_letters)) {
             //use letter from start point
            $x = preg_replace("/[^A-Z]+/", "", $this->startPoint);
            $count = 0;
            $max = count($this->values);
            for ($x; $x <= 'Z' && $count < $max; $x++) {
                array_push($this->_letters, $x);
                $count++;
            }
        }
        return $this->_letters;
    }

    /**
     * Will get current row number
     * @return integer
     */
    public function getCurrentRowNumber() {
        if ($this->_currentRowNumber === null) {
            //If null take from start point
            $startNumber = preg_replace("/[^0-9]+/", "", $this->startPoint);
            $this->setCurrentRowNumber($startNumber);
        }
        return $this->_currentRowNumber;
    }

    /**
     * Set new current row number
     * @param integer $new
     */
    public function setCurrentRowNumber($new) {
        $this->_currentRowNumber = $new;
    }

    /**
     * Will set headers of Excel file
     */
    private function setHeaders() {
        $col = 0;
        $letters = $this->getLetters();
        $row = $this->getCurrentRowNumber();
        foreach ($this->values as $config) {
            $label = isset($config['header']) ? $config['header'] : "";
            $this->getActiveSheet()
                    ->setCellValue($letters[$col] . $row, $label)
                    ->getStyle($letters[$col] . $row)
                    ->applyFromArray($this->headerStyle);
            $col++;
        }
        $this->setCurrentRowNumber( ++$row);
    }

    /**
     * Will set Active Records values to excel
     */
    private function setRows() {
        $letters = $this->getLetters();
        $row = $this->getCurrentRowNumber();

        if (!is_array($this->models)) {
            $this->models = [$this->models];
        }

        foreach ($this->models as $model) {
            $col = 0;
            foreach ($this->values as $config) {
                $value = call_user_func($config['value'], $model);
                $cell = $this->getActiveSheet();
                $cell->setCellValue($letters[$col] . $row, $value);
                if (!empty($this->style)) {//styling takes time
                    $cell->getStyle($letters[$col] . $row)->applyFromArray($this->style);
                }
                $col++;
            }
            $row++;
        }
        $this->setCurrentRowNumber($row);
    }

    /**
     * Will export the results
     * @return mixed
     */
    public function export() {
        $this->validateInput();
        $this->setHeaders();
        $this->setRows();
        return $this->sendFileToBrowser();
    }

    /**
     * Will send file to the client's browser
     */
    private function sendFileToBrowser() {
        // ** Redirect output to a client’s web browser (Excel2007)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $this->filename . '.xlsx"');
        header("Pragma: "); //IE8 quick fix.
        header("Cache-Control: "); //IE8 quick fix.

        $objWriter = \PHPExcel_IOFactory::createWriter($this, 'Excel2007');
        $objWriter->save('php://output');
        \Yii::$app->end();
    }

}
