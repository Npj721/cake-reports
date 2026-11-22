
<?php
require __DIR__ . '/vendor/autoload.php';
require __DIR__ . '/config/paths.php';

use App\Service\ExcelReportGenerator;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$generator = new ExcelReportGenerator();

$spreadsheet = $generator->loadTemplate('foo.xlsx');

$generator->replaceVariables($spreadsheet, [
    'name' => 'John',
]);

$writer = new Xlsx($spreadsheet);
$writer->save(__DIR__ . '/tmp/foo-generated.xlsx');

echo 'Rapport généré avec succès !';
