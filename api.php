<?php
ini_set('display_errors', '1');
ini_set('display_startup_errors', '1');
error_reporting(E_ALL);

function listFolderFiles($dir): array{
    $ffs = scandir($dir);

    unset($ffs[array_search('.', $ffs, true)]);
    unset($ffs[array_search('..', $ffs, true)]);

    return $ffs;
}

function mb_trim($str) {
	return preg_replace("/^\s+|\s+$/u", "", $str); 
}

function translit_file($filename){
	$converter = array(
		'а' => 'a',    'б' => 'b',    'в' => 'v',    'г' => 'g',    'д' => 'd',
		'е' => 'e',    'ё' => 'e',    'ж' => 'zh',   'з' => 'z',    'и' => 'i',
		'й' => 'y',    'к' => 'k',    'л' => 'l',    'м' => 'm',    'н' => 'n',
		'о' => 'o',    'п' => 'p',    'р' => 'r',    'с' => 's',    'т' => 't',
		'у' => 'u',    'ф' => 'f',    'х' => 'h',    'ц' => 'c',    'ч' => 'ch',
		'ш' => 'sh',   'щ' => 'sch',  'ь' => '',     'ы' => 'y',    'ъ' => '',
		'э' => 'e',    'ю' => 'yu',   'я' => 'ya',
 
		'А' => 'A',    'Б' => 'B',    'В' => 'V',    'Г' => 'G',    'Д' => 'D',
		'Е' => 'E',    'Ё' => 'E',    'Ж' => 'Zh',   'З' => 'Z',    'И' => 'I',
		'Й' => 'Y',    'К' => 'K',    'Л' => 'L',    'М' => 'M',    'Н' => 'N',
		'О' => 'O',    'П' => 'P',    'Р' => 'R',    'С' => 'S',    'Т' => 'T',
		'У' => 'U',    'Ф' => 'F',    'Х' => 'H',    'Ц' => 'C',    'Ч' => 'Ch',
		'Ш' => 'Sh',   'Щ' => 'Sch',  'Ь' => '',     'Ы' => 'Y',    'Ъ' => '',
		'Э' => 'E',    'Ю' => 'Yu',   'Я' => 'Ya',
	);

    $newFilename = strtr($filename, $converter);

    $newFilename = mb_ereg_replace('[^\w]', ' ', $newFilename);
    $newFilename = mb_ereg_replace('\s+', ' ', $newFilename);
	
	return $newFilename;
}

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

$names = explode(PHP_EOL, file_get_contents('names.txt'));
for($i = 0; $i < count($names); $i++) $names[$i] = mb_trim($names[$i]);

$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");
$spreadsheet = $reader->load("export.xlsx");

$exTables = ['p1.XLSX', 'p2.XLSX', 'p3.XLSX'];
$exSheets = [];
foreach($exTables as $spreadsheetName) {
    $pSpreadsheet = $reader->load($spreadsheetName);
    $exSheets[] = $pSpreadsheet;
}

foreach($names as $name){
	$sheet = $spreadsheet->getSheetByName('Object');
	$lines = [];
	foreach($sheet->getRowIterator() as $row){
		$cellIterator = $row->getCellIterator();
		foreach($cellIterator as $cell){
			$val = $cell->getValue();
			if(mb_stripos($val, $name) !== false) {
                $lastVal = array_reverse(explode('\\', $val))[0] ?? '0';
                $lines[] = [$val, $lastVal];
            }
		}
	}

    $ffname = translit_file($name);
	$nameSp = new Spreadsheet();
	$nameSp->removeSheetByIndex(0);
	$wsh = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($nameSp, $ffname);
	$nameSp->addSheet($wsh, 0);
	$wsh->getColumnDimension('A')->setWidth(30);

	$wsh->fromArray($lines, NULL, 'A1');

    $pNewSheet = null;
    $rowOffset = 0;
    foreach($exSheets as $exSheet){
        $pSheets = $exSheet->getAllSheets();

        foreach($pSheets as $pSheet){
            $pTranslited = translit_file($pSheet->getTitle());
            if($pTranslited == $ffname){
                $pNewSheet ??= new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($nameSp, 'Equals');

                $lines = [];
                foreach($pSheet->getRowIterator() as $row){
                    if($row->getRowIndex() == 0 || $row->getRowIndex() == 1) continue;
                    $cellIterator = $row->getCellIterator();
                    $line = [];
                    foreach($cellIterator as $cell){
                        $val = $cell->getValue();
                        $line[] = $val;
                    }

                    $lines[] = array_reverse($line);
                }
                $pNewSheet->fromArray($lines, NULL, 'A' . ($rowOffset + 1));
                $rowOffset += count($lines);
            }
        }
    }
    if($pNewSheet != null) $nameSp->addSheet($pNewSheet, 1);

    echo $ffname;
    if($pNewSheet != null) echo ' +';
    echo '<br>';

	$writer = ExcelIOFactory::createWriter($nameSp, 'Xlsx');
	$writer->save(__DIR__ . '/tables/' . $ffname . '.xlsx');
}

foreach(listFolderFiles('./tables') as $filename){
    $filePath = 'tables/' . $filename;

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