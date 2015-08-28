<?php

/**
 * @copyright Copyright Victor Demin, 2015
 * @license https://github.com/ruskid/yii2-excel-exporter/LICENSE
 * @link https://github.com/ruskid/yii2-excel-exporter#README
 */

namespace ruskid\csvexporter;

/**
 * Little Yii2 Helper that exports Active Records to CSV
 * @author Victor Demin <demin@trabeja.com>
 */
class CSVExporter extends \PHPExcel {

    /**
     * PHPExcel Settings here
     */
    public function __construct() {
        parent::__construct();
        $this->setActiveSheetIndex(0);//Use only one sheet
    }
    
    
    public function export(){
        echo 1;
    }
    
}
