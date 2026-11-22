<?php

declare(strict_types=1);

namespace App\Controller;

use App\Service\ExcelReportGenerator;
use App\Service\PdfConverterService;
use Cake\Core\Configure;
use RuntimeException;

class ReportsController extends AppController
{
    private const TEMPLATE_PATH = ROOT . DS . 'templates_excel' . DS;

    public function generate()
    {
        $this->autoRender = false;

        $data = $this->request->getData();

        $templateName = $data['templateName'] ?? null;
        $jsonData = $data['jsonData'] ?? [];
        $options = $data['options'] ?? [];

        if (empty($templateName)) {
            return $this->response
                ->withStatus(400)
                ->withType('application/json')
                ->withStringBody(json_encode(['error' => 'Le paramètre "templateName" est obligatoire.'], JSON_UNESCAPED_UNICODE));
        }

        $exportType = strtoupper($options['type'] ?? 'PDF');
        $templateFile = self::TEMPLATE_PATH . $templateName;
        if (!file_exists($templateFile)) {
            return $this->response
                ->withStatus(404)
                ->withType('application/json')
                ->withStringBody(json_encode(['error' => "Le template '$templateName' n'a pas été trouvé."], JSON_UNESCAPED_UNICODE));
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
                    $generator->processSheet(
                        $sheet, 
                        $sheetData['variables'] ?? [], 
                        $sheetData['arrays'] ?? [], 
                        $sheetOptions
                    );
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
                $configuredPath = Configure::read('LibreOffice.binaryPath');
                $converter = new PdfConverterService($configuredPath);
                $pdfPath = $converter->convert($xlsxPath, TMP);

                return $this->response
                    ->withType('application/pdf')
                    ->withDownload($baseFilename . '.pdf')
                    ->withFile($pdfPath);
            }

            return $this->response
                ->withStatus(400)
                ->withType('application/json')
                ->withStringBody(json_encode(['error' => "Type d'export non supporté. Utilisez 'PDF' ou 'EXCEL'."], JSON_UNESCAPED_UNICODE));

        } catch (RuntimeException $e) {
            return $this->response
                ->withStatus(500)
                ->withType('application/json')
                ->withStringBody(json_encode(['error' => 'Erreur lors de la génération : ' . $e->getMessage()], JSON_UNESCAPED_UNICODE));
        }
    }
}