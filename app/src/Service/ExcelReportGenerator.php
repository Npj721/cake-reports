<?php

declare(strict_types=1);

namespace App\Service;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use RuntimeException;

class ExcelReportGenerator
{
    private string $templateDirectory;

    public function __construct(?string $templateDirectory = null)
    {
        $this->templateDirectory = $templateDirectory ?? ROOT . DS . 'templates_excel';
    }

    /**
     * Charge un template Excel en toute sécurité.
     */
    public function loadTemplate(string $templateName): Spreadsheet
    {
        if ($templateName === '' || basename($templateName) !== $templateName || !preg_match('/^[a-zA-Z0-9_-]+\.xlsx$/', $templateName)) {
            throw new RuntimeException('Nom de template invalide.');
        }

        $templatePath = $this->templateDirectory . DS . $templateName;

        if (!is_file($templatePath)) {
            throw new RuntimeException(sprintf('Template introuvable : %s', $templateName));
        }

        $reader = IOFactory::createReader('Xlsx');
        $reader->setIncludeCharts(true);

        return $reader->load($templatePath);
    }

    /**
     * Point d'entrée principal pour traiter une feuille de calcul.
     * Agit comme un chef d'orchestre vers les méthodes spécialisées.
     */
    public function processSheet(
        Worksheet $sheet,
        array $variables,
        array $arrays,
        array $options = []
    ): void {
        // 1. Analyse du template et préparation des données à insérer
        $rowsToInsert = $this->parseAndMarkCellsForInsertion($sheet, $variables, $arrays, $options);

        if (empty($rowsToInsert)) {
            return; // Rien à insérer, on peut sortir tôt
        }

        // 2. Insertion physique des lignes et peuplement des données
        $this->insertRowsAndPopulateData($sheet, $rowsToInsert);

        // 3. Application des règles métier par tableau inséré
        foreach ($rowsToInsert as $row => $info) {
            $this->applyTableStyles($sheet, $row, $info, $options);
            $this->applyPageBreaks($sheet, $row, $info, $options);
            $this->applyDataValidations($sheet, $row, $info, $options);
        }

        // 4. Ajustement final de la mise en page (largeurs de colonnes)
        $this->adjustColumnWidths($sheet, $rowsToInsert, $options);
    }

    // =========================================================================
    // MÉTHODES PRIVÉES SPÉCIALISÉES
    // =========================================================================

    /**
     * Parcourt les cellules, remplace les variables simples et détecte les marqueurs de tableau.
     * Retourne un tableau structuré des insertions à effectuer.
     */
    private function parseAndMarkCellsForInsertion(
        Worksheet $sheet,
        array $variables,
        array $arrays,
        array $options
    ): array {
        $removeUnmatched = $options['remove_unmatched_variables'] ?? false;
        $arraysNoHeader = $options['arrays_no_header'] ?? false;
        $styles = $options['style'] ?? [];
        $renderers = $options['renderers'] ?? [];

        $highestRow = $sheet->getHighestRow();
        $highestColumn = $sheet->getHighestColumn();
        $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn);

        $rowsToInsert = [];

        for ($row = 1; $row <= $highestRow; $row++) {
            for ($col = 1; $col <= $highestColumnIndex; $col++) {
                $cell = $sheet->getCell([$col, $row]);
                $cellValue = $cell->getValue();

                if (is_string($cellValue)) {
                    // A. Remplacement des variables simples {{variable}}
                    if (str_contains($cellValue, '{{')) {
                        $cellValue = preg_replace_callback(
                            '/\{\{([a-zA-Z0-9_.-]+)\}\}/',
                            function (array $matches) use ($variables, $renderers, $removeUnmatched): string {
                                $key = $matches[1];

                                if (!array_key_exists($key, $variables)) {
                                    return $removeUnmatched ? '' : $matches[0];
                                }

                                $result = $variables[$key];

                                if (isset($renderers[$key]) && is_callable($renderers[$key])) {
                                    $result = $renderers[$key]($result, $key);
                                }

                                if ($result === null) return '';
                                if (is_bool($result)) return $result ? '1' : '0';
                                if (is_scalar($result)) return (string) $result;

                                throw new RuntimeException("La variable '{$key}' doit produire une valeur scalaire.");
                            },
                            $cellValue
                        );
                    }

                    // B. Détection des tableaux [[arrayKey]]
                    if (preg_match('/\[\[([a-zA-Z0-9_.-]+)\]\]/', $cellValue, $matches)) {
                        $arrayKey = $matches[1];

                        if (isset($arrays[$arrayKey]) && is_array($arrays[$arrayKey]) && !empty($arrays[$arrayKey]['data'])) {
                            $tableData = $arrays[$arrayKey]['data'];

                            // Appliquer les renderers spécifiques au tableau
                            if (isset($arrays[$arrayKey]['renderers']) && is_array($arrays[$arrayKey]['renderers'])) {
                                foreach ($tableData as &$rowData) {
                                    foreach ($rowData as $colName => &$value) {
                                        if (isset($arrays[$arrayKey]['renderers'][$colName]) && is_callable($arrays[$arrayKey]['renderers'][$colName])) {
                                            $value = $arrays[$arrayKey]['renderers'][$colName]($value, $colName);
                                        }
                                    }
                                }
                                unset($rowData, $value);
                            }

                            // Ajouter les en-têtes si demandé
                            if (!$arraysNoHeader) {
                                $headers = array_keys(reset($tableData));
                                array_unshift($tableData, $headers);
                            }

                            $rowsToInsert[$row] = [
                                'numRows' => count($tableData),
                                'colStart' => $col,
                                'data' => $tableData,
                                'style' => $styles[$arrayKey] ?? null,
                                'validations' => $arrays[$arrayKey]['validations'] ?? null,
                            ];

                            // Nettoie la cellule qui contenait le marqueur [[arrayKey]]
                            $cellValue = str_replace($matches[0], '', $cellValue);
                        } elseif ($removeUnmatched) {
                            $cellValue = str_replace($matches[0], '', $cellValue);
                        }
                    }

                    $cell->setValue($cellValue);
                }
            }
        }

        return $rowsToInsert;
    }

    /**
     * Insère les lignes physiques dans la feuille et y injecte les données.
     */
    private function insertRowsAndPopulateData(Worksheet $sheet, array $rowsToInsert): void
    {
        // 1. Insérer les lignes nécessaires (en partant du bas pour ne pas décaler les indices des suivantes)
        foreach (array_reverse($rowsToInsert, true) as $row => $info) {
            if ($info['numRows'] > 1) {
                $sheet->insertNewRowBefore($row + 1, $info['numRows'] - 1);
            }
        }

        // 2. Remplir les données avec fromArray
        foreach ($rowsToInsert as $row => $info) {
            $startCol = $info['colStart'];
            $tableData = $info['data'];
            $startColumnLetter = Coordinate::stringFromColumnIndex($startCol);

            $sheet->fromArray($tableData, null, $startColumnLetter . $row);
        }
    }

    /**
     * Applique tous les styles (en-tête, zébrage, bordures, alignements, polices) à un tableau inséré.
     */
    private function applyTableStyles(Worksheet $sheet, int $startRow, array $info, array $options): void
    {
        $style = $info['style'] ?? null;
        if (!$style) {
            return;
        }

        $tableData = $info['data'];
        $startCol = $info['colStart'];
        $startColumnLetter = Coordinate::stringFromColumnIndex($startCol);
        $endColumnLetter = Coordinate::stringFromColumnIndex($startCol + count(array_keys($tableData[0])) - 1);

        $hasStripedColors = isset($style['color1'], $style['color2']);
        $hasBorder = isset($style['border']);
        $headerStyle = $style['header'] ?? null;

        for ($i = 0; $i < count($tableData); $i++) {
            $currentRow = $startRow + $i;
            $range = $startColumnLetter . $currentRow . ':' . $endColumnLetter . $currentRow;
            $rowStyle = $sheet->getStyle($range);

            // 1. Style spécifique de l'en-tête (première ligne du bloc inséré)
            if ($i === 0 && $headerStyle) {
                if (isset($headerStyle['backgroundColor'])) {
                    $rowStyle->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                        ->getStartColor()->setARGB($headerStyle['backgroundColor']);
                }
                if (!empty($headerStyle['fontBold'])) {
                    $rowStyle->getFont()->setBold(true);
                }
                if (isset($headerStyle['fontColor'])) {
                    $rowStyle->getFont()->getColor()->setARGB($headerStyle['fontColor']);
                }
                if (isset($headerStyle['fontSize'])) {
                    $rowStyle->getFont()->setSize($headerStyle['fontSize']);
                }
                if (isset($headerStyle['alignment'])) {
                    $rowStyle->getAlignment()->setHorizontal(
                        match (strtolower($headerStyle['alignment'])) {
                            'center' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                            'right'  => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
                            default  => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                        }
                    );
                }
            } 
            // 2. Style des lignes de données (zébrage)
            elseif ($hasStripedColors) {
                // Si un header existe, la première donnée est à $i=1. On ajuste l'index pour que le zébrage commence proprement.
                $dataIndex = $headerStyle ? $i - 1 : $i;
                $fillColor = ($dataIndex % 2 === 0) ? $style['color1'] : $style['color2'];
                
                $rowStyle->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                    ->getStartColor()->setARGB($fillColor);
            }

            // 3. Bordures globales du tableau
            if ($hasBorder) {
                $rowStyle->getBorders()->getAllBorders()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            }

            // 4. Alignements spécifiques par colonne
            if (isset($style['alignment']) && is_array($style['alignment'])) {
                foreach ($style['alignment'] as $colName => $alignment) {
                    $colIndex = $this->getColumnIndexByName($tableData[0], $colName);
                    if ($colIndex !== null) {
                        $cellLetter = Coordinate::stringFromColumnIndex($startCol + $colIndex);
                        $horizontal = match (strtolower($alignment)) {
                            'left' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                            'center' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                            'right' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
                            default => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_GENERAL,
                        };
                        $sheet->getStyle($cellLetter . $currentRow)->getAlignment()->setHorizontal($horizontal);
                    }
                }
            }

            // 5. Polices en gras par colonne
            if (isset($style['fontBold']) && is_array($style['fontBold'])) {
                foreach ($style['fontBold'] as $colName) {
                    $colIndex = $this->getColumnIndexByName($tableData[0], $colName);
                    if ($colIndex !== null) {
                        $cellLetter = Coordinate::stringFromColumnIndex($startCol + $colIndex);
                        $sheet->getStyle($cellLetter . $currentRow)->getFont()->setBold(true);
                    }
                }
            }

            // 6. Bordures spécifiques par colonne
            if (isset($style['borders']) && is_array($style['borders'])) {
                foreach ($style['borders'] as $colName => $borderTypes) {
                    $colIndex = $this->getColumnIndexByName($tableData[0], $colName);
                    if ($colIndex !== null && is_array($borderTypes)) {
                        $cellLetter = Coordinate::stringFromColumnIndex($startCol + $colIndex);
                        $cellBorders = $sheet->getStyle($cellLetter . $currentRow)->getBorders();

                        foreach ($borderTypes as $borderType => $active) {
                            if ($active) {
                                $borderStyle = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
                                match (strtolower($borderType)) {
                                    'left' => $cellBorders->getLeft()->setBorderStyle($borderStyle),
                                    'right' => $cellBorders->getRight()->setBorderStyle($borderStyle),
                                    'top' => $cellBorders->getTop()->setBorderStyle($borderStyle),
                                    'bottom' => $cellBorders->getBottom()->setBorderStyle($borderStyle),
                                    'all' => $cellBorders->getAllBorders()->setBorderStyle($borderStyle),
                                };
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * Insère des sauts de page à intervalle régulier dans les données du tableau.
     */
    private function applyPageBreaks(Worksheet $sheet, int $startRow, array $info, array $options): void
    {
        $pageBreakInterval = $options['pageBreakInterval'] ?? false;

        if ($pageBreakInterval !== false && is_int($pageBreakInterval) && $pageBreakInterval > 0) {
            $tableData = $info['data'];
            for ($i = 0; $i < count($tableData); $i++) {
                if ($i > 0 && $i % $pageBreakInterval === 0) {
                    $sheet->setBreak("A" . ($startRow + $i), Worksheet::BREAK_ROW);
                }
            }
        }
    }

    /**
     * Configure les listes déroulantes (validations de données) pour des colonnes spécifiques.
     */
    private function applyDataValidations(Worksheet $sheet, int $startRow, array $info, array $options): void
    {
        $validations = $info['validations'] ?? null;
        if (!$validations || !is_array($validations)) {
            return;
        }

        $tableData = $info['data'];
        $startCol = $info['colStart'];
        $arraysNoHeader = $options['arrays_no_header'] ?? false;
        $headers = $tableData[0];

        foreach ($validations as $colName => $allowedValues) {
            $colIndex = array_search($colName, $headers, true);
            if ($colIndex === false) {
                continue; // Colonne non trouvée dans les données
            }

            $colLetter = Coordinate::stringFromColumnIndex($startCol + $colIndex);
            $startDataRow = $arraysNoHeader ? $startRow : $startRow + 1;
            $endDataRow = $startRow + count($tableData) - 1;

            if ($startDataRow <= $endDataRow) {
                // On crée la validation sur la première cellule de données
                $validation = $sheet->getCell($colLetter . $startDataRow)->getDataValidation();
                $validation->setType(DataValidation::TYPE_LIST);
                $validation->setErrorStyle(DataValidation::STYLE_STOP);
                $validation->setAllowBlank(true);
                $validation->setShowInputMessage(true);
                $validation->setShowErrorMessage(true);
                $validation->setShowDropDown(true);
                $validation->setFormula1('"' . implode(',', $allowedValues) . '"');

                // On clone cette validation sur toutes les cellules de la colonne en dessous
                for ($r = $startDataRow + 1; $r <= $endDataRow; $r++) {
                    $sheet->getCell($colLetter . $r)->setDataValidation(clone $validation);
                }
            }
        }
    }

    /**
     * Ajuste la largeur des colonnes concernées par les tableaux insérés.
     */
    private function adjustColumnWidths(Worksheet $sheet, array $rowsToInsert, array $options): void
    {
        $autoSizeColumns = $options['autoSizeColumns'] ?? true;
        $columnWidths = $options['columnWidths'] ?? [];
        $processedColumns = [];

        // Identifier toutes les colonnes touchées par les insertions
        foreach ($rowsToInsert as $info) {
            $startColIdx = $info['colStart'];
            $numCols = count(array_keys($info['data'][0]));
            for ($i = 0; $i < $numCols; $i++) {
                $processedColumns[Coordinate::stringFromColumnIndex($startColIdx + $i)] = true;
            }
        }

        // Appliquer les règles de largeur
        foreach ($processedColumns as $colLetter => $_) {
            if (isset($columnWidths[$colLetter])) {
                $sheet->getColumnDimension($colLetter)->setWidth($columnWidths[$colLetter]);
            } elseif ($autoSizeColumns) {
                $sheet->getColumnDimension($colLetter)->setAutoSize(true);
            }
        }
    }

    /**
     * Utilitaire : Trouve l'index (0-based) d'une colonne dans un tableau associatif en fonction de son nom de clé.
     */
    private function getColumnIndexByName(array $rowData, string $colName): ?int
    {
        $index = array_search($colName, array_keys($rowData), true);
        return $index !== false ? (int)$index : null;
    }

    // =========================================================================
    // MÉTHODES PUBLIQUES UTILITAIRES
    // =========================================================================

    /**
     * Applique les formatages spécifiques pour une exportation PDF propre.
     */
    public function applyPdfFormatting(Spreadsheet $spreadsheet): void
    {
        foreach ($spreadsheet->getAllSheets() as $sheet) {
            $sheet->getPageSetup()->setFitToWidth(1)->setFitToHeight(0);
            
            foreach ($sheet->getRowDimensions() as $rowDimension) {
                $rowDimension->setRowHeight(max($rowDimension->getRowHeight() ?: 15, 15) * 1.4);
            }
            
            $sheet->getStyle($sheet->calculateWorksheetDimension())
                  ->getAlignment()
                  ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
                  
            foreach ($sheet->getCellCollection()->getCoordinates() as $coordinate) {
                $alignment = $sheet->getStyle($coordinate)->getAlignment();
                $horizontal = $alignment->getHorizontal();
                
                if (in_array($horizontal, [
                    \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_GENERAL,
                    \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT
                ], true)) {
                    $alignment->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT)->setIndent(1);
                } elseif ($horizontal === \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT) {
                    $alignment->setIndent(1);
                }
            }
        }
    }

    /**
     * Sauvegarde l'objet Spreadsheet au format Xlsx.
     */
    public function saveXlsx(Spreadsheet $spreadsheet, string $outputPath): void
    {
        $outputDirectory = dirname($outputPath);
        if (!is_dir($outputDirectory)) {
            throw new RuntimeException(sprintf('Le dossier de destination n’existe pas : %s', $outputDirectory));
        }

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->setIncludeCharts(true);
        $writer->save($outputPath);

        $spreadsheet->disconnectWorksheets();
    }
}