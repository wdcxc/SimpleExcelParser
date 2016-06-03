<?php

require_once "Classes/PHPExcel.php";
require_once "Classes/PHPExcel/IOFactory.php";
require_once "Classes/PHPExcel/Reader/IReadFilter.php";

class chunkReadFilter implements PHPExcel_Reader_IReadFilter{
    private $filterCols;
    private $_startRow = 0;
    private $_endRow = 0;

    public function setFilterCols($filterCols)
    {
        $this->filterCols = $filterCols;
    }

    public function getFilterCols()
    {
        return $this->filterCols;
    }

    public function readCell($column, $row, $worksheetName = '')
    {
        if(($row==1)||($row>=$this->_startRow&&$row<=$this->_endRow)){
            if(empty($this->filterCols)||in_array($column,$this->filterCols))
                return true;
        }
        return false;
    }

    public function setRows($startRow,$chunkSize){
        $this->_startRow=$startRow;
        $this->_endRow=$startRow+$chunkSize;
    }
}


