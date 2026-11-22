<?php

declare(strict_types=1);

namespace App\Service;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use RuntimeException;

class ExcelReportGenerator
{
    private string $templateDirectory;

    public function __construct(?string $templateDirectory = null)
    {
        $this->templateDirectory = $templateDirectory
            ?? ROOT . DS . 'templates_excel';
    }

    /**
     * Charge un template Excel.
     */
    public function loadTemplate(string $templateName): Spreadsheet
    {
        // On accepte uniquement un nom de fichier, pas un chemin.
        if (
            $templateName === '' ||
            basename($templateName) !== $templateName ||
            !preg_match('/^[a-zA-Z0-9_-]+\.xlsx$/', $templateName)
        ) {
            throw new RuntimeException('Nom de template invalide.');
        }

        $templatePath = $this->templateDirectory . DS . $templateName;

        if (!is_file($templatePath)) {
            throw new RuntimeException(
                sprintf('Template introuvable : %s', $templateName)
            );
        }

        $reader = IOFactory::createReader('Xlsx');
        $reader->setIncludeCharts(true);

        return $reader->load($templatePath);
    }


    /**
     * Remplace les variables {{variable}} dans toutes les cellules du classeur.
     */
    public function replaceVariables(
        Spreadsheet $spreadsheet,
        array $variables
    ): void {
        foreach ($spreadsheet->getAllSheets() as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                foreach ($row->getCellIterator() as $cell) {
                    $value = $cell->getValue();

                    if (!is_string($value) || !str_contains($value, '{{')) {
                        continue;
                    }

                    $newValue = preg_replace_callback(
                        '/\{\{([a-zA-Z0-9_.-]+)\}\}/',
                        function (array $matches) use ($variables): string {
                            $key = $matches[1];

                            if (!array_key_exists($key, $variables)) {
                                return $matches[0];
                            }

                            $value = $variables[$key];

                            if ($value === null) {
                                return '';
                            }

                            if (is_bool($value)) {
                                return $value ? '1' : '0';
                            }

                            if (is_scalar($value)) {
                                return (string)$value;
                            }

                            throw new RuntimeException(
                                sprintf('La variable "%s" doit être scalaire.', $key)
                            );
                        },
                        $value
                    );

                    if ($newValue !== $value) {
                        $cell->setValue($newValue);
                    }
                }
            }
        }
    }


    /**
     * Génère un fichier Excel à partir d'un template.
     *
     * @param string $templateName Nom du template XLSX.
     * @param array $variables Variables à remplacer.
     * @param string $outputPath Chemin complet du fichier généré.
     *
     * @return string Chemin du fichier généré.
     */
    public function generateXlsx(
        string $templateName,
        array $variables,
        string $outputPath
    ): string {
        $spreadsheet = $this->loadTemplate($templateName);

        $this->replaceVariables($spreadsheet, $variables);

        $outputDirectory = dirname($outputPath);

        if (!is_dir($outputDirectory)) {
            throw new RuntimeException(
                sprintf('Le dossier de destination n’existe pas : %s', $outputDirectory)
            );
        }

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->setIncludeCharts(true);
        $writer->save($outputPath);

        $spreadsheet->disconnectWorksheets();

        return $outputPath;
    }
}
