<?php
require_once "SimpleExcelParser.php";

$sep = new SimpleExcelParser("data.xls");
// $sep->extractDataIntoFile("extdata.txt");
// $sep->getData();
// echo "finish";
$fileInfo=$sep->getFileInfo();
var_dump($fileInfo);