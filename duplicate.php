<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function findDuplicates($filePath): void
{
    $spreadsheet = IOFactory::load($filePath);
    $sheet = $spreadsheet->getActiveSheet();

    $data = [];
    $duplicates = [];

    foreach ($sheet->getRowIterator() as $row) {
        $rowIndex = $row->getRowIndex();
        $isrc = $sheet->getCell('AN' . $rowIndex)->getValue();
        $title = $sheet->getCell('F' . $rowIndex)->getValue();
        $artist = $sheet->getCell('E' . $rowIndex)->getValue();

        if ($isrc) {
            if (isset($data[$isrc])) {
                foreach ($data[$isrc] as $existingSong) {
                    if ($existingSong['title'] !== $title || $existingSong['artist'] !== $artist) {
                        $duplicates[] = [
                            'ISRC' => $isrc,
                            'Title' => $title,
                            'Artist' => $artist,
                            'Duplicate of' => $existingSong['title'] . ' - ' . $existingSong['artist']
                        ];
                    }
                }
            }
            $data[$isrc][] = ['title' => $title, 'artist' => $artist];
        }
    }
    outputDuplicates($duplicates);

}

function outputDuplicates($duplicates): void
{

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $sheet->setCellValue('A1', 'ISRC');
    $sheet->setCellValue('B1', 'Title');
    $sheet->setCellValue('C1', 'Artist');
    $sheet->setCellValue('D1', 'Duplicate of');

    $row = 2;
    foreach ($duplicates as $duplicate) {
        $sheet->setCellValue('A' . $row, $duplicate['ISRC']);
        $sheet->setCellValue('B' . $row, $duplicate['Title']);
        $sheet->setCellValue('C' . $row, $duplicate['Artist']);
        $sheet->setCellValue('D' . $row, $duplicate['Duplicate of']);
        $row++;
    }

    $writer = new Xlsx($spreadsheet);
    $outputFile = 'duplicates.xlsx';
    $writer->save($outputFile);

    echo "Duplicates have been saved to $outputFile\n";
}

findDuplicates(__DIR__ . '/1.xlsx');

