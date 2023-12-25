<?php
ini_set('display_errors', '1');
ini_set('display_startup_errors', '1');
error_reporting(E_ALL);

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
	
	$new = '';
	
	$file = pathinfo(trim($filename));
	if (!empty($file['dirname']) && @$file['dirname'] != '.') {
		$new .= rtrim($file['dirname'], '/') . '/';
	}
 
	if (!empty($file['filename'])) {
		$file['filename'] = str_replace(array(' ', ','), '-', $file['filename']);
		$file['filename'] = strtr($file['filename'], $converter);
		$file['filename'] = mb_ereg_replace('[-]+', '-', $file['filename']);
		$file['filename'] = trim($file['filename'], '-');					
		$new .= $file['filename'];
	}
 
	if (!empty($file['extension'])) {
		$new .= '.' . $file['extension'];
	}
	
	return $new;
}

require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style;
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
$sum = 0;
foreach($names as $key => $name){
	$sheet = $spreadsheet->getSheetByName('Object');
	$lines = [];
    $linesno = [];
	foreach($sheet->getRowIterator() as $row){
		$cellIterator = $row->getCellIterator();
		foreach($cellIterator as $cell){
			$val = $cell->getValue();
            $valno = $cell->getValue();
			if(mb_stripos($val, $name) !== false) {
                $lines[] = [$val];
            } else if (mb_stripos($val, $name) == false) {
                $linesno[]=[$valno];
            };
		}
	}
	
	$fname = mb_ereg_replace('[^\w]', ' ', $name);
	$fname = mb_ereg_replace('\s+', ' ', $fname);
	
	$nameSp = new Spreadsheet();
	$nameSp->removeSheetByIndex(0);
	$wsh = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($nameSp, $fname);
	$nameSp->addSheet($wsh, 0);
	$wsh->getColumnDimension('A')->setWidth(30);
	
	$wsh->fromArray($lines, NULL, 'A1');
	
	$ffname = translit_file($fname);
	echo $ffname . '<br>';
    echo count($lines);
    $countn = count($linesno);
    $sum+=$countn;

    $writer = ExcelIOFactory::createWriter($nameSp, 'Xlsx');
	$writer->save(__DIR__ . '/tables/' . $ffname . '.xlsx');
};

// Get the files in the folder
$files = new RecursiveDirectoryIterator('C:\OSPanel\domains\phpspreadsheet\tables');
echo $files;
$iterator = new RecursiveIteratorIterator($files);
$fileNames = new RecursiveIteratorIterator($iterator);
$fileNames->setIterateOnlyExisting(false);

foreach ($fileNames as $fileName) {
    // Check if the file is an Excel file
    if ($fileName->getExtension() === 'xlsx') {
        $spreadsheet = IOFactory::load($fileName->getPath());

        $sheet = $spreadsheet->addSheet('Sheet1');
        $sheet->setTitle('Sheet1');
        $sheet->setDefaultColumnWidth(10);

        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($fileName->getPath());
    }
}



?>