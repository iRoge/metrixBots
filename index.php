<?php
namespace Stripmag\Document;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require_once 'vendor/autoload.php';
set_time_limit(1000000);

$collectedData = $_POST['collectedData'];

$spreadsheet = IOFactory::load('newFile.xlsx');

$heightOfDateRow = 7;
$startRow = 6;
foreach ($collectedData as $sheetName => $thisSitesData) {
    try {
        $worksheet = $spreadsheet->getSheetByName($sheetName);
    } catch (\Exception $e) {
        var_dump($e->getMessage());
        $currentRow = $startRow;
        break;
    }
    echo 'Заполняем ' . $sheetName . PHP_EOL;
    echo 'Открыта страница таблицы с ссылкой ' . $worksheet->getCell('C2')->getValue() . PHP_EOL;

    $currentRow = $startRow;
    while ($worksheet->getCell('B' . $currentRow)->getValue()) {
        $currentRow += $heightOfDateRow;
    }
    if ($currentRow > $startRow) {
        copyRows($worksheet, $currentRow - $heightOfDateRow, $currentRow, 7, 215);
    }

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

    // Countries Block
    for ($column = 'K'; $column != 'DW'; $column++) {
        $countryCell = $worksheet->getCell($column . 5);
        if (isset($thisSitesData['countriesInfo'][strtolower($countryCell->getValue())])) {
            $currentColumn = $column;
            // Country cell percent
            $worksheet->setCellValue($currentColumn . $currentRow, $thisSitesData['countriesInfo'][strtolower($countryCell->getValue())]['percent']);
            // Country cell difference
            $worksheet->setCellValue((++$currentColumn) . $currentRow, $thisSitesData['countriesInfo'][strtolower($countryCell->getValue())]['difference']);
            if (!$thisSitesData['countriesInfo'][strtolower($countryCell->getValue())]['direction']) {
                $worksheet->getStyle($currentColumn . $currentRow)->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
            } else {
                $worksheet->getStyle($currentColumn . $currentRow)->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_GREEN);
            }
        }
    }

    // Direct
    $worksheet->setCellValue('DW' . $currentRow, $thisSitesData['directPercent']);
    // Refferals
    $worksheet->setCellValue('DX' . $currentRow, $thisSitesData['referralsPercent']);
    // Search
    $worksheet->setCellValue('ER' . $currentRow, $thisSitesData['searchPercent']);
    // Social
    $worksheet->setCellValue('FL' . $currentRow, $thisSitesData['socialPercent']);
    // Mail
    $worksheet->setCellValue('FQ' . $currentRow, $thisSitesData['mailPercent']);
    // Display
    $worksheet->setCellValue('FR' . $currentRow, $thisSitesData['displayPercent']);

    // Top Referring Sites Block
    if (!empty($thisSitesData['topReferringSitesInfo'])) {
        for ($column = 'DX', $n = 0; $column != 'EH'; $column++, $column++, $n++) {
            if (!isset($thisSitesData['topReferringSitesInfo'][$n]) || $n >= 5) {
                break;
            }
            $currentColumn = $column;
            $worksheet->setCellValue($column . ($currentRow + 4), $thisSitesData['topReferringSitesInfo'][$n]['siteName']);
            $worksheet->setCellValue(++$currentColumn . ($currentRow + 4), $thisSitesData['topReferringSitesInfo'][$n]['difference']);
            if (!$thisSitesData['topReferringSitesInfo'][$n]['direction']) {
                $worksheet->getStyle($currentColumn . ($currentRow + 4))->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
            } else {
                $worksheet->getStyle($currentColumn . ($currentRow + 4))->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_GREEN);
            }
            $worksheet->setCellValue($column . ($currentRow + 5), $thisSitesData['topReferringSitesInfo'][$n]['percent']);
        }
    }
    // Top Destination Sites Block
    if (!empty($thisSitesData['topDestinationSitesInfo'])) {
        for ($column = 'EH', $n = 0; $column != 'ER' || isset($thisSitesData['topDestinationSitesInfo'][$n]) || $n != 5; $column++, $column++, $n++) {
            if (!isset($thisSitesData['topDestinationSitesInfo'][$n]) || $n >= 5) {
                break;
            }
            $currentColumn = $column;
            $worksheet->setCellValue($column . ($currentRow + 4), $thisSitesData['topDestinationSitesInfo'][$n]['siteName']);
            $worksheet->setCellValue(++$currentColumn . ($currentRow + 4), $thisSitesData['topDestinationSitesInfo'][$n]['difference']);
            if (!$thisSitesData['topDestinationSitesInfo'][$n]['direction']) {
                $worksheet->getStyle($currentColumn . ($currentRow + 4))->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
            } else {
                $worksheet->getStyle($currentColumn . ($currentRow + 4))->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_GREEN);
            }
            $worksheet->setCellValue($column . ($currentRow + 5), $thisSitesData['topDestinationSitesInfo'][$n]['percent']);
        }
    }

    // Organic Search Percent
    $worksheet->setCellValue('ER' . ($currentRow + 2), $thisSitesData['organicSearchPercent']);
    // Organic Search Block
    if (!empty($thisSitesData['organicSearchInfo'])) {
        for ($column = 'ER', $n = 0; $column != 'FB'; $column++, $column++, $n++) {
            if (!isset($thisSitesData['organicSearchInfo'][$n]) || $n >= 5) {
                break;
            }
            $currentColumn = $column;
            $worksheet->setCellValue($column . ($currentRow + 5), $thisSitesData['organicSearchInfo'][$n]['searchText']);
            $worksheet->setCellValue(++$currentColumn . ($currentRow + 5), $thisSitesData['organicSearchInfo'][$n]['difference']);
            if (!$thisSitesData['organicSearchInfo'][$n]['direction']) {
                $worksheet->getStyle($currentColumn . ($currentRow + 5))->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
            } else {
                $worksheet->getStyle($currentColumn . ($currentRow + 5))->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_GREEN);
            }
            $worksheet->setCellValue($column . ($currentRow + 6), $thisSitesData['organicSearchInfo'][$n]['percent']);
        }
    }

    // Paid Search Block
    if (!empty($thisSitesData['paidSearchInfo'])) {
        for ($column = 'FB', $n = 0; $column != 'FL'; $column++, $column++, $n++) {
            if (!isset($thisSitesData['paidSearchInfo'][$n]) || $n >= 5) {
                break;
            }
            $currentColumn = $column;
            $worksheet->setCellValue($column . ($currentRow + 5), $thisSitesData['paidSearchInfo'][$n]['searchText']);
            $worksheet->setCellValue(++$currentColumn . ($currentRow + 5), $thisSitesData['paidSearchInfo'][$n]['difference']);
            if (!$thisSitesData['paidSearchInfo'][$n]['direction']) {
                $worksheet->getStyle($currentColumn . ($currentRow + 5))->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
            } else {
                $worksheet->getStyle($currentColumn . ($currentRow + 5))->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_GREEN);
            }
            $worksheet->setCellValue($column . ($currentRow + 6), $thisSitesData['paidSearchInfo'][$n]['percent']);
        }
    }
    // Paid Search Percent
    $worksheet->setCellValue('FB' . ($currentRow + 2), $thisSitesData['paidSearchPercent']);

    // Social Block
    for ($column = 'FL'; $column != 'FQ'; $column++) {
        $socialCell = $worksheet->getCell($column . ($currentRow + 2));
        if (isset($thisSitesData['socialInfo'][strtolower($socialCell->getValue())])) {
            // Country cell percent
            $worksheet->setCellValue($column . ($currentRow + 3), $thisSitesData['socialInfo'][strtolower($socialCell->getValue())]['percent']);
        }
    }

    // Audience Interests Block
    if (!empty($thisSitesData['audienceInterestsInfo'])) {
        for ($column = 'FS', $n = 0; $column != 'FX'; $column++, $n++) {
            if (!isset($thisSitesData['audienceInterestsInfo'][$n]) || $n >= 5) {
                break;
            }
            $worksheet->setCellValue($column . ($currentRow + 1), $thisSitesData['audienceInterestsInfo'][$n]);
        }
    }

    // Also visited websites Block
    if (!empty($thisSitesData['alsoVisitedWebsitesInfo'])) {
        for ($column = 'FX', $n = 0; $column != 'GC'; $column++, $n++) {
            if (!isset($thisSitesData['alsoVisitedWebsitesInfo'][$n]) || $n >= 5) {
                break;
            }
            $worksheet->setCellValue($column . ($currentRow + 1), $thisSitesData['alsoVisitedWebsitesInfo'][$n]);
        }
    }

    // Similarity Block
    if (!empty($thisSitesData['similarSitesInfo'])) {
        for ($column = 'GC', $n = 0; $column != 'GM'; $column++, $n++) {
            if (!isset($thisSitesData['similarSitesInfo'][$n]) || $n >= 10) {
                break;
            }
            $worksheet->setCellValue($column . ($currentRow + 1), $thisSitesData['similarSitesInfo'][$n]);
        }
    }

    // Rank Block
    if (!empty($thisSitesData['rankSitesInfo'])) {
        for ($column = 'GM', $n = 0; $column != 'GW'; $column++, $n++) {
            if (!isset($thisSitesData['rankSitesInfo'][$n]) || $n >= 10) {
                break;
            }
            $worksheet->setCellValue($column . ($currentRow + 1), $thisSitesData['rankSitesInfo'][$n]);
        }
    }

    // Android Apps Block
    if (!empty($thisSitesData['androidAppsInfo'])) {
        for ($column = 'GW', $n = 0; $column != 'HB'; $column++, $n++) {
            if (!isset($thisSitesData['androidAppsInfo'][$n]) || $n >= 5) {
                break;
            }
            $worksheet->setCellValue($column . ($currentRow + 1), $thisSitesData['androidAppsInfo'][$n]);
        }
    }

    // Apple Apps Block
    if (!empty($thisSitesData['appleAppsInfo'])) {
        for ($column = 'HB', $n = 0; $column != 'HG'; $column++, $n++) {
            if (!isset($thisSitesData['appleAppsInfo'][$n]) || $n >= 5) {
                break;
            }
            $worksheet->setCellValue($column . ($currentRow + 1), $thisSitesData['appleAppsInfo'][$n]);
        }
    }

}

(new Xlsx($spreadsheet))->save('newFile.xlsx');
echo 'XLSX CREATED!!!';

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

