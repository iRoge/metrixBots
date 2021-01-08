<?php
namespace Stripmag\Document;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require_once 'vendor/autoload.php';
set_time_limit(1000000);

$collectedData = $_POST['collectedData'];

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

    // Countries Block
    for ($column = 'K'; $column != 'DW'; $column++) {
        $countryCell = $worksheet->getCell($column . 5);
        if (isset($thisSitesData['countriesInfo'][strtolower($countryCell->getValue())])) {
            $currentColumn = $column;
            // Country cell percent
            $worksheet->setCellValue($currentColumn . $currentRow, $thisSitesData['countriesInfo'][strtolower($countryCell->getValue())]['percent']);
            // Country cell difference
            $worksheet->setCellValue((++$currentColumn) . $currentRow, $thisSitesData['countriesInfo'][strtolower($countryCell->getValue())]['difference']);
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
        for ($column = 'DX', $i = 0; $column != 'EH'; $column++, $column++, $i++) {
            if (!isset($thisSitesData['topReferringSitesInfo'][$i]) || $i >= 5) {
                break;
            }
            $currentColumn = $column;
            $worksheet->setCellValue($column . ($currentRow + 4), $thisSitesData['topReferringSitesInfo'][$i]['siteName']);
            $worksheet->setCellValue(++$currentColumn . ($currentRow + 4), $thisSitesData['topReferringSitesInfo'][$i]['difference']);
            $worksheet->setCellValue($column . ($currentRow + 5), $thisSitesData['topReferringSitesInfo'][$i]['percent']);

        }
    }
    // Top Destination Sites Block
    if (!empty($thisSitesData['topDestinationSitesInfo'])) {
        for ($column = 'EH', $i = 0; $column != 'ER' || isset($thisSitesData['topDestinationSitesInfo'][$i]) || $i != 5; $column++, $column++, $i++) {
            if (!isset($thisSitesData['topDestinationSitesInfo'][$i]) || $i >= 5) {
                break;
            }
            $currentColumn = $column;
            $worksheet->setCellValue($column . ($currentRow + 4), $thisSitesData['topDestinationSitesInfo'][$i]['siteName']);
            $worksheet->setCellValue(++$currentColumn . ($currentRow + 4), $thisSitesData['topDestinationSitesInfo'][$i]['difference']);
            $worksheet->setCellValue($column . ($currentRow + 5), $thisSitesData['topDestinationSitesInfo'][$i]['percent']);
        }
    }

    // Organic Search Percent
    $worksheet->setCellValue('ER' . ($currentRow + 2), $thisSitesData['organicSearchPercent']);
    // Organic Search Block
    if (!empty($thisSitesData['organicSearchInfo'])) {
        for ($column = 'ER', $i = 0; $column != 'FB'; $column++, $column++, $i++) {
            if (!isset($thisSitesData['organicSearchInfo'][$i]) || $i >= 5) {
                break;
            }
            $currentColumn = $column;
            $worksheet->setCellValue($column . ($currentRow + 5), $thisSitesData['organicSearchInfo'][$i]['searchText']);
            $worksheet->setCellValue(++$currentColumn . ($currentRow + 5), $thisSitesData['organicSearchInfo'][$i]['difference']);
            $worksheet->setCellValue($column . ($currentRow + 6), $thisSitesData['organicSearchInfo'][$i]['percent']);
        }
    }

    // Paid Search Block
    if (!empty($thisSitesData['paidSearchInfo'])) {
        for ($column = 'FB', $i = 0; $column != 'FL'; $column++, $column++, $i++) {
            if (!isset($thisSitesData['paidSearchInfo'][$i]) || $i >= 5) {
                break;
            }
            $currentColumn = $column;
            $worksheet->setCellValue($column . ($currentRow + 5), $thisSitesData['paidSearchInfo'][$i]['searchText']);
            $worksheet->setCellValue(++$currentColumn . ($currentRow + 5), $thisSitesData['paidSearchInfo'][$i]['difference']);
            $worksheet->setCellValue($column . ($currentRow + 6), $thisSitesData['paidSearchInfo'][$i]['percent']);
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
        for ($column = 'FS', $i = 0; $column != 'FX'; $column++, $i++) {
            if (!isset($thisSitesData['audienceInterestsInfo'][$i]) || $i >= 5) {
                break;
            }
            $worksheet->setCellValue($column . ($currentRow + 1), $thisSitesData['audienceInterestsInfo'][$i]);
        }
    }

    // Also visited websites Block
    if (!empty($thisSitesData['alsoVisitedWebsitesInfo'])) {
        for ($column = 'FX', $i = 0; $column != 'GC'; $column++, $i++) {
            if (!isset($thisSitesData['alsoVisitedWebsitesInfo'][$i]) || $i >= 5) {
                break;
            }
            $worksheet->setCellValue($column . ($currentRow + 1), $thisSitesData['alsoVisitedWebsitesInfo'][$i]);
        }
    }

    // Similarity Block
    if (!empty($thisSitesData['similarSitesInfo'])) {
        for ($column = 'GC', $i = 0; $column != 'GM'; $column++, $i++) {
            if (!isset($thisSitesData['similarSitesInfo'][$i]) || $i >= 10) {
                break;
            }
            $worksheet->setCellValue($column . ($currentRow + 1), $thisSitesData['similarSitesInfo'][$i]);
        }
    }

    // Rank Block
    if (!empty($thisSitesData['rankSitesInfo'])) {
        for ($column = 'GM', $i = 0; $column != 'GW'; $column++, $i++) {
            if (!isset($thisSitesData['rankSitesInfo'][$i]) || $i >= 10) {
                break;
            }
            $worksheet->setCellValue($column . ($currentRow + 1), $thisSitesData['rankSitesInfo'][$i]);
        }
    }

    // Android Apps Block
    if (!empty($thisSitesData['androidAppsInfo'])) {
        for ($column = 'GW', $i = 0; $column != 'HB'; $column++, $i++) {
            if (!isset($thisSitesData['androidAppsInfo'][$i]) || $i >= 5) {
                break;
            }
            $worksheet->setCellValue($column . ($currentRow + 1), $thisSitesData['androidAppsInfo'][$i]);
        }
    }

    // Apple Apps Block
    if (!empty($thisSitesData['appleAppsInfo'])) {
        for ($column = 'HB', $i = 0; $column != 'HG'; $column++, $i++) {
            if (!isset($thisSitesData['appleAppsInfo'][$i]) || $i >= 5) {
                break;
            }
            $worksheet->setCellValue($column . ($currentRow + 1), $thisSitesData['appleAppsInfo'][$i]);
        }
    }

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

