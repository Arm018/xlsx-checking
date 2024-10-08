<?php
ini_set('memory_limit', '2G');

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require 'vendor/autoload.php';


function compareISRCs($file1, $file2, $excludingFile)
{
    $spreadsheet1 = IOFactory::load($file1);
    $spreadsheet2 = IOFactory::load($file2);
    $spreadsheet3 = IOFactory::load($excludingFile);

    $isrcsFile1 = getISRCs($spreadsheet1);
    $isrcsFile2 = getISRCs($spreadsheet2);
    $isrcsFile3 = getISRCs($spreadsheet3);

    $isrcsFile1 = array_diff($isrcsFile1, $isrcsFile3);
    $isrcsFile2 = array_diff($isrcsFile2, $isrcsFile3);

    $inFile1 = array_diff($isrcsFile1, $isrcsFile2);
    $inFile2 = array_diff($isrcsFile2, $isrcsFile1);

    outputComparison($inFile1, $inFile2);

}

function getISRCs($spreadsheet) {
    $sheet = $spreadsheet->getActiveSheet();
    $isrcs = [];

    foreach ($sheet->getRowIterator() as $row) {
        $rowIndex = $row->getRowIndex();
        $isrc = $sheet->getCell('AN' . $rowIndex)->getValue();
        if ($isrc) {
            $isrcs[] = $isrc;
        }
    }
    return $isrcs;
}

function outputComparison($inFile1, $inFile2) {
    $spreadSheet= new Spreadsheet();
    $sheet = $spreadSheet->getActiveSheet();

    $sheet->setCellValue('A1', 'In File 1 but not in File 2');
    $sheet->setCellValue('B1', 'In File 2 but not in File 1');

    $row = 2;
    $maxRows = max(count($inFile1), count($inFile2));

    for ($i = 0; $i < $maxRows; $i++) {
        if (isset($inFile1[$i])) {
            $sheet->setCellValue('A' . $row, $inFile1[$i]);
        }
        if (isset($inFile2[$i])) {
            $sheet->setCellValue('B' . $row, $inFile2[$i]);
        }
        $row++;
    }

    $writer = new Xlsx($spreadSheet);
    $outputFile = 'comparison_result.xlsx';
    $writer->save($outputFile);

    echo "Comparison has been saved to $outputFile\n";

}

compareISRCs(__DIR__ . '/1.xlsx', __DIR__ . '/2.xlsx', __DIR__ . '/3.xlsx');

