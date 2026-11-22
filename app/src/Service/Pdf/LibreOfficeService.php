<?php

declare(strict_types=1);

namespace App\Service\Pdf;

use RuntimeException;

class LibreOfficeService implements PdfConverterInterface
{
    private ?string $configuredBinaryPath;

    public function __construct(?string $binaryPath = null)
    {
        $this->configuredBinaryPath = $binaryPath
            ?? \Cake\Core\Configure::read('LibreOffice.binaryPath');
    }

    public function getName(): string
    {
        return 'libreoffice';
    }

    public function isAvailable(): bool
    {
        try {
            $this->getValidatedBinaryPath();
            return true;
        } catch (\Throwable $e) {
            return false;
        }
    }

    public function convert(string $xlsxPath, string $outputDir): string
    {
        if (!is_file($xlsxPath)) {
            throw new RuntimeException(sprintf('Le fichier source est introuvable : %s', $xlsxPath));
        }

        if (!is_dir($outputDir)) {
            throw new RuntimeException(sprintf('Le dossier de destination est introuvable : %s', $outputDir));
        }

        $binary = $this->getValidatedBinaryPath();
        $profileDir = $outputDir . DIRECTORY_SEPARATOR . 'lo_profile_' . uniqid();
        mkdir($profileDir, 0777, true);

        try {
            $profileUri = 'file:///' . str_replace('\\', '/', $profileDir);

            $cmd = sprintf(
                '%s "-env:UserInstallation=%s" --headless --invisible --convert-to pdf %s --outdir %s 2>&1',
                escapeshellarg($binary),
                escapeshellarg($profileUri),
                escapeshellarg($xlsxPath),
                escapeshellarg($outputDir)
            );

            $output = [];
            $returnCode = 0;
            exec($cmd, $output, $returnCode);

            $pdfPath = str_replace('.xlsx', '.pdf', $xlsxPath);

            if ($returnCode === 0 && is_file($pdfPath)) {
                return $pdfPath;
            }

            throw new RuntimeException(
                sprintf("Echec de la conversion PDF. Code : %d. Details :\n%s", $returnCode, implode("\n", $output))
            );

        } finally {
            $this->deleteDirectory($profileDir);
        }
    }

    private function getValidatedBinaryPath(): string
    {
        if (!empty($this->configuredBinaryPath)) {
            $isWindows = (PHP_OS_FAMILY === 'Windows');
            $isValid = $isWindows ? is_file($this->configuredBinaryPath) : is_executable($this->configuredBinaryPath);

            if (!$isValid) {
                throw new RuntimeException(
                    sprintf('Chemin LibreOffice invalide : "%s"', $this->configuredBinaryPath)
                );
            }
            return $this->configuredBinaryPath;
        }

        if (PHP_OS_FAMILY === 'Windows') {
            foreach ([
                'C:\Program Files\LibreOffice\program\soffice.exe',
                'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
            ] as $path) {
                if (is_file($path)) {
                    return $path;
                }
            }
        } else {
            $binary = trim((string)(shell_exec('which libreoffice 2>/dev/null') ?: shell_exec('which soffice 2>/dev/null')));
            if (!empty($binary) && is_executable($binary)) {
                return $binary;
            }
        }

        throw new RuntimeException('LibreOffice est introuvable.');
    }

    private function deleteDirectory(string $dir): void
    {
        if (!is_dir($dir)) {
            return;
        }
        $items = new \RecursiveIteratorIterator(
            new \RecursiveDirectoryIterator($dir, \RecursiveDirectoryIterator::SKIP_DOTS),
            \RecursiveIteratorIterator::CHILD_FIRST
        );
        foreach ($items as $item) {
            $item->isDir() ? rmdir($item->getRealPath()) : unlink($item->getRealPath());
        }
        rmdir($dir);
    }
}