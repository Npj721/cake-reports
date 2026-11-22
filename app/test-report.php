
<?php
require __DIR__ . '/vendor/autoload.php';
require __DIR__ . '/config/paths.php';

use App\Service\ExcelReportGenerator;

$generator = new ExcelReportGenerator();

$output = $generator->generateXlsx(
    'foo.xlsx',
    [
        'name' => 'John',
    ],
    __DIR__ . '/tmp/foo-generated.xlsx'
);

echo 'Rapport généré : ' . $output . PHP_EOL;
