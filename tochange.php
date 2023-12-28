<?php
ini_set('display_errors', '1');
ini_set('display_startup_errors', '1');
error_reporting(E_ALL);
require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Spreadsheet\Style;
use PhpOffice\PhpSpreadsheet\Chart\Chart;
use PhpOffice\PhpSpreadsheet\Chart\DataSeries;
use PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;
use PhpOffice\PhpSpreadsheet\Chart\Legend;
use PhpOffice\PhpSpreadsheet\Chart\PlotArea;
use PhpOffice\PhpSpreadsheet\Chart\Title;
use PhpOffice\PhpSpreadsheet\IOFactory as ExcelIOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
function listFolderFiles($dir): array{
    $ffs = scandir($dir);

    unset($ffs[array_search('.', $ffs, true)]);
    unset($ffs[array_search('..', $ffs, true)]);

    return $ffs;
}

foreach(listFolderFiles('./tablestochange') as $filename){
    $filePath = 'tablestochange/' . $filename;

    $tSpreadsheet = $reader->load($filePath);
    $sheets = $tSpreadsheet->getAllSheets();
    if(count($sheets) < 2) continue;

    $linesFirst = [];
    $linesSecond = [];
    foreach($sheets[0]->getRowIterator() as $firstRow){
        $line = [];
        foreach($firstRow->getCellIterator() as $cell) {
            $val = $cell->getValue();
            $line[] = $val;
        }

        for($i = count($line); $i < 2; $i++) $line[$i] = null;
        $linesFirst[] = $line;
    }

    foreach($sheets[1]->getRowIterator() as $firstRow){
        $line = [];
        foreach($firstRow->getCellIterator() as $cell){
            $val = $cell->getValue();
            $line[] = $val;
        }

        for($i = count($line); $i < 2; $i++) $line[$i] = null;
        $linesSecond[] = $line;
    }

    $newLinesFirst = [];
    foreach($linesFirst as $lineFirst){
        $hasEqual = false;
        foreach($linesSecond as $lineSecond){
            if($lineFirst[1] != null && $lineSecond[0] != null && mb_trim($lineSecond[0]) != '' && translit_file($lineFirst[1]) == translit_file($lineSecond[0])){
                $newLinesFirst[] = [$lineSecond[0], $lineSecond[1]];
                $hasEqual = true;
                break;
            }
        }

        if(!$hasEqual) $newLinesFirst[] = ['', ''];
    }

    $sheets[0]->fromArray($newLinesFirst, NULL, 'C1');

    $writer = ExcelIOFactory::createWriter($tSpreadsheet, 'Xlsx');
    $writer->save($filePath);
}



?>