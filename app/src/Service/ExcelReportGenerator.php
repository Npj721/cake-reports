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
}
