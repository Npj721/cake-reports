<?php

declare(strict_types=1);

namespace App\Controller;

use App\Service\ExcelReportGenerator;
use App\Service\PdfConverterService;
use Cake\Core\Configure;
use RuntimeException;

class ReportsController extends AppDBController
{
    private const TEMPLATE_PATH = ROOT . DS . 'templates_excel' . DS;

    public function generate($templateName, $jsonData, $options = [])
    {
        $this->autoRender = false;
        $exportType = $options['type'] ?? 'PDF';

        $templateFile = self::TEMPLATE_PATH . $templateName;
        if (!file_exists($templateFile)) {
            return $this->response->withStatus(404)->withStringBody("Le Template $templateName n'a pas ete trouve.");
        }

        try {
            $generator = new ExcelReportGenerator(self::TEMPLATE_PATH);
            $spreadsheet = $generator->loadTemplate($templateName);

            foreach ($jsonData as $sheetName => $sheetData) {
                $sheet = $spreadsheet->getSheetByName($sheetName);
                if ($sheet) {
                    $sheetOptions = $options;
                    if (isset($sheetData['renderers'])) {
                        $sheetOptions['renderers'] = $sheetData['renderers'];
                    }
                    $generator->processSheet($sheet, $sheetData['variables'] ?? [], $sheetData['arrays'] ?? [], $sheetOptions);
                }
            }

            if ($exportType === 'PDF') {
                $generator->applyPdfFormatting($spreadsheet);
            }

            $uniqueId = uniqid();
            $baseFilename = sprintf('%s_%s-%s', date('Ymd'), str_replace('.xlsx', '', $templateName), $uniqueId);
            $xlsxPath = TMP . $baseFilename . '.xlsx';

            $generator->saveXlsx($spreadsheet, $xlsxPath);

            if ($exportType === 'EXCEL') {
                return $this->response
                    ->withType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                    ->withDownload($baseFilename . '.xlsx')
                    ->withFile($xlsxPath);
            }

            if ($exportType === 'PDF') {
                // Lecture de la config CakePHP et injection dans le service
                $configuredPath = Configure::read('LibreOffice.binaryPath');
                $converter = new PdfConverterService($configuredPath);
                
                $pdfPath = $converter->convert($xlsxPath, TMP);

                return $this->response
                    ->withType('application/pdf')
                    ->withDownload($baseFilename . '.pdf')
                    ->withFile($pdfPath);
            }

            return $this->response->withStatus(400)->withStringBody("Type d'export non supporte.");

        } catch (RuntimeException $e) {
            // Log::error('Erreur export : ' . $e->getMessage());
            return $this->response->withStatus(500)->withStringBody("Erreur lors de la generation : " . $e->getMessage());
        }
    }
}