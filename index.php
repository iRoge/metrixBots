<?php
namespace Stripmag\Document;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require_once 'vendor/autoload.php';
set_time_limit(1000000);

$collectedData = $_GET['collectedData'];

$spreadsheet = IOFactory::load('file.xlsx');
$heightOfDateRow = 7;
$startRow = 6;
$i = 0;
$worksheet = $spreadsheet->getSheet($i);
while ($worksheet) {
    $currentRow = $startRow;
    while ($worksheet->getCell('B' . $currentRow)->getValue()) {
        $currentRow += $heightOfDateRow;
    }
    if ($currentRow > $startRow) {
        copyRows($worksheet, $currentRow - $heightOfDateRow, $currentRow, 7, 215);
    }
    $thisSitesData = $collectedData[$i];

    // Дата
    $worksheet->setCellValue('B' . $currentRow, date("Y-m-d"));
    // Global Rank
    $worksheet->setCellValue('C' . $currentRow, $thisSitesData['globalRank']);
    // Country Rank
    $worksheet->setCellValue('D' . $currentRow, $thisSitesData['countryRank']);
    // Category Rank
    $worksheet->setCellValue('E' . $currentRow, $thisSitesData['categoryRank']);
    // Category
    $worksheet->setCellValue('F' . $currentRow, $thisSitesData['category']);
    // Total visits
    $worksheet->setCellValue('G' . $currentRow, $thisSitesData['totalVisits']);
    // Avg. Visit Duration
    $worksheet->setCellValue('H' . $currentRow, $thisSitesData['avgVisitsDuration']);
    // Pages per Visit
    $worksheet->setCellValue('I' . $currentRow, $thisSitesData['pagesPerVisit']);
    // Bounce Rate
    $worksheet->setCellValue('J' . $currentRow, $thisSitesData['bounceRate']);

    $i++;
    try {
        $worksheet = $spreadsheet->getSheet($i);
        break;
    } catch (\Exception $e) {
        var_dump($e->getMessage());
        $currentRow = $startRow;
        break;
    }
}

(new Xlsx($spreadsheet))->save('file7.xlsx');
var_dump('done');

function copyRows($sheet,$srcRow,$dstRow,$height,$width) {
    for ($row = 0; $row < $height; $row++) {
        for ($col = 0; $col < $width; $col++) {
            $cell = $sheet->getCellByColumnAndRow($col, $srcRow + $row);
            $style = $sheet->getStyleByColumnAndRow($col, $srcRow + $row);
            $dstCell = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . (string)($dstRow + $row);
            $sheet->setCellValue($dstCell, $cell->getValue());
            $sheet->duplicateStyle($style, $dstCell);
        }

        $h = $sheet->getRowDimension($srcRow + $row)->getRowHeight();
        $sheet->getRowDimension($dstRow + $row)->setRowHeight($h);
    }

    foreach ($sheet->getMergeCells() as $mergeCell) {
        $mc = explode(":", $mergeCell);
        $col_s = preg_replace("/[0-9]*/", "", $mc[0]);
        $col_e = preg_replace("/[0-9]*/", "", $mc[1]);
        $row_s = ((int)preg_replace("/[A-Z]*/", "", $mc[0])) - $srcRow;
        $row_e = ((int)preg_replace("/[A-Z]*/", "", $mc[1])) - $srcRow;

        if (0 <= $row_s && $row_s < $height) {
            $merge = $col_s . (string)($dstRow + $row_s) . ":" . $col_e . (string)($dstRow + $row_e);
            $sheet->mergeCells($merge);
        }
    }
}

