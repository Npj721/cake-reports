<?php

declare(strict_types=1);

require __DIR__ . '/vendor/autoload.php';
require __DIR__ . '/config/bootstrap.php';

use App\Service\Pdf\GotenbergService;
use App\Service\Pdf\LibreOfficeService;
use App\Service\Pdf\PdfConverterGateway;
use Cake\Core\Configure;

$tmpDir = __DIR__ . DIRECTORY_SEPARATOR . 'tmp';
if (!is_dir($tmpDir)) {
    mkdir($tmpDir, 0777, true);
}

// On cree un petit XLSX de test avec PhpSpreadsheet
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Test de conversion PDF');
$sheet->setCellValue('B1', 'Ligne 2');

$testXlsx = $tmpDir . '/test_source.xlsx';
$writer = new Xlsx($spreadsheet);
$writer->save($testXlsx);

echo "Fichier XLSX de test cree : " . realpath($testXlsx) . "\n\n";

// On nettoie les PDF precedents pour eviter les faux positifs
foreach (glob($tmpDir . '/test_source.pdf') as $f) {
    unlink($f);
}

// --- SCENARIO 1 : Gotenberg seul ---
echo "=== SCENARIO 1 : Gotenberg seul ===\n";
$gotenberg = new GotenbergService();
echo "Disponible : " . ($gotenberg->isAvailable() ? 'OUI' : 'NON') . "\n";
if ($gotenberg->isAvailable()) {
    try {
        $pdf = $gotenberg->convert($testXlsx, $tmpDir);
        echo "SUCCES : " . realpath($pdf) . "\n";
    } catch (\Throwable $e) {
        echo "ECHEC : " . $e->getMessage() . "\n";
    }
} else {
    echo "Skipped : Gotenberg non joignable.\n";
}

// On nettoie pour le prochain test
foreach (glob($tmpDir . '/test_source.pdf') as $f) {
    unlink($f);
}

// --- SCENARIO 2 : LibreOffice seul ---
echo "\n=== SCENARIO 2 : LibreOffice seul ===\n";
$libre = new LibreOfficeService();
echo "Disponible : " . ($libre->isAvailable() ? 'OUI' : 'NON') . "\n";
if ($libre->isAvailable()) {
    try {
        $pdf = $libre->convert($testXlsx, $tmpDir);
        echo "SUCCES : " . realpath($pdf) . "\n";
    } catch (\Throwable $e) {
        echo "ECHEC : " . $e->getMessage() . "\n";
    }
} else {
    echo "Skipped : LibreOffice non detecte.\n";
}

// On nettoie pour le prochain test
foreach (glob($tmpDir . '/test_source.pdf') as $f) {
    unlink($f);
}

// --- SCENARIO 3 : Passerelle complete (fallback) ---
echo "\n=== SCENARIO 3 : Passerelle (gotenberg -> libreoffice) ===\n";
echo "Strategies configurees : " . implode(', ', Configure::read('Pdf.strategies')) . "\n";
try {
    $gateway = PdfConverterGateway::fromConfig();
    $pdf = $gateway->convert($testXlsx, $tmpDir);
    echo "SUCCES FINAL : " . realpath($pdf) . "\n";
} catch (\Throwable $e) {
    echo "ECHEC FINAL (toutes strategies epuisees) :\n" . $e->getMessage() . "\n";
}

echo "\nTermine.\n";