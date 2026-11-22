<?php

declare(strict_types=1);

require __DIR__ . '/vendor/autoload.php';
require __DIR__ . '/config/paths.php'; 

use App\Service\ExcelReportGenerator;
use App\Service\PdfConverterService;

// 1. Preparation de l'environnement
$tmpDir = __DIR__ . '/tmp';
if (!is_dir($tmpDir)) {
    mkdir($tmpDir, 0777, true);
}

$templateName = 'rapport_test.xlsx';
$baseFileName = 'rapport_test_genere_' . date('Ymd_His');
$xlsxPath = $tmpDir . DIRECTORY_SEPARATOR . $baseFileName . '.xlsx';

echo "DEBUT DU TEST DE GENERATION EXCEL...\n";
echo "Template attendu : " . (ROOT . DS . 'templates_excel' . DS . $templateName) . "\n\n";

try {
    // 2. Initialisation du generateur Excel
    $generator = new ExcelReportGenerator(ROOT . DS . 'templates_excel');

    // 3. Definition des donnees (simulant ce que le controleur enverrait)
    $jsonData = [
        'Sheet1' => [ // Doit correspondre EXACTEMENT au nom de l'onglet dans votre Excel
            'variables' => [
                'company_name' => 'TechCorp Solutions Inc.',
                'report_date'  => new DateTime(), // Test d'un renderer sur un objet
                'total_amount' => 192500,
                'unmatched_var' => 'Ceci sera supprime', // Test de remove_unmatched_variables
            ],
            'arrays' => [
                'employees' => [
                    'data' => [
                        ['id' => 101, 'name' => 'Alice Dupont', 'department' => 'IT', 'salary' => 65000, 'status' => 'Actif'],
                        ['id' => 102, 'name' => 'Bob Martin', 'department' => 'RH', 'salary' => 55000, 'status' => 'En conge'],
                        ['id' => 103, 'name' => 'Charlie Durand', 'department' => 'Finance', 'salary' => 72000, 'status' => 'Actif'],
                        ['id' => 104, 'name' => 'Diana Prince', 'department' => 'IT', 'salary' => 80000, 'status' => 'Actif'],
                        ['id' => 105, 'name' => 'Evan Wright', 'department' => 'Marketing', 'salary' => 60000, 'status' => 'Termine'],
                    ],
                    // Renderers specifiques aux colonnes du tableau
                    'renderers' => [
                        'salary' => fn($val) => number_format((float)$val, 0, ',', ' ') . ' EUR',
                        'status' => fn($val) => strtoupper($val),
                    ],
                    // Validations (listes deroulantes)
                    'validations' => [
                        'status' => ['Actif', 'En conge', 'Termine', 'En formation']
                    ]
                ]
            ]
        ]
    ];

    // 4. Definition des options de formatage
    $options = [
        'remove_unmatched_variables' => true,
        'arrays_no_header' => false, // On veut garder la ligne d'en-tete
        'pageBreakInterval' => 3, // Saut de page toutes les 3 lignes de donnees
        'autoSizeColumns' => true,
        'columnWidths' => [
            'A' => 10, // Force la largeur de la colonne ID
        ],
        'style' => [
            'employees' => [
                'color1' => 'FFFFFFFF', // Blanc
                'color2' => 'FFE2E8F0', // Gris bleute clair (effet zebre)
                'border' => 'all',      // Bordure fine partout
                'alignment' => [
                    'salary' => 'right',
                    'id' => 'center'
                ],
                'fontBold' => ['name'], // Mettre la colonne 'name' en gras
                'borders' => [
                    // Exemple : bordure specifique renforcee a droite de la colonne salary
                    'salary' => ['right' => true] 
                ]
            ]
        ],
        // Renderers globaux (s'appliquent aux variables simples)
        'renderers' => [
            'report_date' => fn($val) => $val instanceof DateTime ? $val->format('d/m/Y a H:i') : $val,
            'total_amount' => fn($val) => number_format((float)$val, 2, ',', ' ') . ' EUR',
        ]
    ];

    // 5. Execution du processus Excel
    echo "Chargement du template...\n";
    $spreadsheet = $generator->loadTemplate($templateName);

    foreach ($jsonData as $sheetName => $sheetData) {
        $sheet = $spreadsheet->getSheetByName($sheetName);
        if ($sheet) {
            echo "Traitement de la feuille : '$sheetName'...\n";
            $generator->processSheet(
                $sheet,
                $sheetData['variables'] ?? [],
                $sheetData['arrays'] ?? [],
                $options
            );
        } else {
            echo "ATTENTION : Feuille '$sheetName' introuvable dans le template.\n";
        }
    }

    echo "Sauvegarde du fichier Excel...\n";
    $generator->saveXlsx($spreadsheet, $xlsxPath);
    echo "SUCCES : Rapport Excel genere avec succes.\n";
    echo "   -> " . realpath($xlsxPath) . "\n\n";

    // 6. Execution du processus de conversion PDF
    echo "DEBUT DE LA CONVERSION PDF...\n";
    
    // Note: Dans un script CLI standalone, assurez-vous que le .env est charge 
    // ou que l'auto-detection de LibreOffice fonctionnera sur votre machine.
    $converter = new PdfConverterService();
    
    echo "Recherche de LibreOffice et conversion en cours...\n";
    $pdfPath = $converter->convert($xlsxPath, $tmpDir);
    
    echo "SUCCES : Conversion PDF realisee avec succes.\n";
    echo "   -> " . realpath($pdfPath) . "\n";

} catch (\Throwable $e) {
    echo "\nERREUR FATALE : " . $e->getMessage() . "\n";
    echo "Fichier : " . $e->getFile() . " (Ligne " . $e->getLine() . ")\n";
    
    // Astuce de debug pour LibreOffice
    if (str_contains($e->getMessage(), 'LibreOffice')) {
        echo "\nACTION REQUISE : Verifiez que LibreOffice est installe et que le chemin est correct\n";
        echo "dans votre fichier .env (variable LIBREOFFICE_BINARY_PATH) ou qu'il est dans le PATH systeme.\n";
    }
}