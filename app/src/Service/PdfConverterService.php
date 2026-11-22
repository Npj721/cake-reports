<?php

declare(strict_types=1);

namespace App\Service;

use RuntimeException;

class PdfConverterService
{
    private ?string $forcedBinaryPath;

    /**
     * @param string|null $forcedBinaryPath Permet de forcer le chemin de LibreOffice via la config si la détection auto échoue.
     */
    public function __construct(?string $forcedBinaryPath = null)
    {
        $this->forcedBinaryPath = $forcedBinaryPath;
    }

    /**
     * Convertit un fichier XLSX en PDF via LibreOffice en ligne de commande.
     *
     * @param string $xlsxPath Chemin absolu du fichier Excel source.
     * @param string $outputDir Chemin absolu du dossier où déposer le PDF.
     * @return string Chemin absolu du fichier PDF généré.
     * @throws RuntimeException Si LibreOffice est introuvable ou si la conversion échoue.
     */
    public function convert(string $xlsxPath, string $outputDir): string
    {
        if (!is_file($xlsxPath)) {
            throw new RuntimeException(sprintf('Le fichier source est introuvable : %s', $xlsxPath));
        }

        if (!is_dir($outputDir)) {
            throw new RuntimeException(sprintf('Le dossier de destination est introuvable : %s', $outputDir));
        }

        // 1. Trouver la binaire LibreOffice
        $binary = $this->forcedBinaryPath ?? $this->detectBinary();

        // 2. Créer un profil utilisateur temporaire isolé (évite les verrous de fichiers)
        $profileDir = $outputDir . DIRECTORY_SEPARATOR . 'lo_profile_' . uniqid();
        mkdir($profileDir, 0777, true);

        try {
            // 3. Construire la commande de manière SÉCURISÉE
            // Note : LibreOffice attend une URI pour UserInstallation, on remplace les \ par des /
            $profileUri = 'file:///' . str_replace('\\', '/', $profileDir);
            
            $cmd = sprintf(
                '%s "-env:UserInstallation=%s" --headless --invisible --convert-to pdf %s --outdir %s 2>&1',
                escapeshellarg($binary),
                escapeshellarg($profileUri),
                escapeshellarg($xlsxPath),
                escapeshellarg($outputDir)
            );

            // 4. Exécution
            $output = [];
            $returnCode = 0;
            exec($cmd, $output, $returnCode);

            // 5. Vérification du résultat
            $pdfPath = str_replace('.xlsx', '.pdf', $xlsxPath);
            
            if ($returnCode === 0 && is_file($pdfPath)) {
                return $pdfPath;
            }

            throw new RuntimeException(
                sprintf(
                    "Échec de la conversion PDF. Code de retour : %d. Détails : %s",
                    $returnCode,
                    implode("\n", $output)
                )
            );

        } finally {
            // 6. Nettoyage garanti du dossier temporaire
            $this->deleteDirectory($profileDir);
        }
    }

    /**
     * Détecte automatiquement le chemin de l'exécutable LibreOffice selon l'OS.
     */
    private function detectBinary(): string
    {
        if (PHP_OS_FAMILY === 'Windows') {
            $paths = [
                'C:\Program Files\LibreOffice\program\soffice.exe',
                'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
            ];
            foreach ($paths as $path) {
                if (is_file($path)) {
                    return $path;
                }
            }
            throw new RuntimeException(
                'LibreOffice est introuvable sous Windows. Veuillez l\'installer ou spécifier le chemin dans la configuration.'
            );
        }

        // Linux / macOS
        $binary = trim((string) shell_exec('which libreoffice 2>/dev/null') ?: shell_exec('which soffice 2>/dev/null'));
        
        if (empty($binary) || !is_executable($binary)) {
            throw new RuntimeException(
                'LibreOffice n\'est pas installé ou n\'est pas dans le PATH du système. Installez-le (ex: apt-get install libreoffice).'
            );
        }

        return $binary;
    }

    /**
     * Supprime récursivement un dossier.
     */
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
            if ($item->isDir()) {
                rmdir($item->getRealPath());
            } else {
                unlink($item->getRealPath());
            }
        }
        
        rmdir($dir);
    }
}