<?php

require_once "ChunkReadFilter.php";

class SimpleExcelParser{

    private $parseFile;
    private $parseFileType;
    private $chunkSize;
    private $parseReader;
    private $fileInfo;
    private $chunkReadFilter;

    /**
     * @param string  $parseFile 
     * @param string  $parseFileType 
     * @param integer $chunkSize the size of data to be load each time 
     */
    public function __construct($parseFile,$parseFileType=null,$chunkSize=2048){
        if(empty($parseFile)){
            echo "invalid parsefile";
            exit;
        }
        $this->parseFile=$parseFile;

        if(empty($parseFileType)){   
            $this->parseFileType=PHPExcel_IOFactory::identify($parseFile);
        } else {
            $this->parseFileType=$parseType;
        }

        $this->chunkSize=$chunkSize;

        $this->parseReader=PHPExcel_IOFactory::createReader($this->parseFileType);
        $this->fileInfo=$this->parseReader->listWorkSheetInfo($this->parseFile)[0];
        $this->chunkReadFilter=new chunkReadFilter();
    }

    /**
     * @param  array $reqInfo example:array("totalRows","totalColumns")
     * @return array
     */
    public function getFileInfo($reqInfo=null){
        if(!empty($reqInfo)){
            if(is_string($reqInfo))
                return $this->fileInfo[$reqInfo];
            if(is_array($reqInfo)){
                $ret=array();
                foreach($reqInfo as $val){
                   $ret[$val]=$this->fileInfo[$val];
                }
                return $ret;
            }
        }
        return $this->fileInfo;
    }

    /**
     * @param array $filterCols example:array("A","B","L","W")        
     */
    public function getData($filterCols=null){
        
        $this->chunkReadFilter->setFilterCols($filterCols);
        $this->parseReader->setReadFilter($this->chunkReadFilter); 
        $ret=array();

        $totalRows=$this->fileInfo['totalRows'];
        for($row=0;$row<=$totalRows;$row+=$this->chunkSize){
            $this->chunkReadFilter->setRows($row,$this->chunkSize);
            $objPhpExcel=$this->parseReader->load($this->parseFile);
            $sheetData=$objPhpExcel->getActiveSheet()->toArray();

            // if the result is to large, it's better to process it partially;
            $ret=array_merge($ret,$sheetData);
            // var_dump($sheetData);
        }

        return $ret;
    }

    /**
     * @param array $filterCols example:array("A","B","L","W")      
     * @param string $mode default:"wb"      
     */
    public function extractDataIntoFile($filename,$filterCols=null,$mode=null){
        if(empty($filename)){
            echo "invalid filename";
            exit;
        }

        $fp=fopen($filename,empty($mode)?"wb":$mode);
        if(!$fp){
            echo "can not open file";
            exit;
        }

        $this->chunkReadFilter->setFilterCols($filterCols);
        $this->parseReader->setReadFilter($this->chunkReadFilter); 

        $totalRows=$this->fileInfo['totalRows'];
        for($row=0;$row<=$totalRows;$row+=$this->chunkSize){

            $this->chunkReadFilter->setRows($row,$this->chunkSize);
            $objPhpExcel=$this->parseReader->load($this->parseFile);
            $sheetData=$objPhpExcel->getActiveSheet()->toArray();

            // do your own procedure below
            foreach($sheetData as $oneRow) {
                foreach($oneRow as $key=>$val){
                    $data=$val."\t";
                }
                $data.="\n";
                fwrite($fp,$data,strlen($data));
            }
        }

        fclose($fp);
        unset($fp);
    }

}