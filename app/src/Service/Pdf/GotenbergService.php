<?php

declare(strict_types=1);

namespace App\Service\Pdf;

use Cake\Http\Client;
use RuntimeException;

class GotenbergService implements PdfConverterInterface
{
    private string $baseUrl;
    private bool $enabled;

    public function __construct(?string $baseUrl = null, ?bool $enabled = null)
    {
        $this->baseUrl = rtrim($baseUrl ?? \Cake\Core\Configure::read('Gotenberg.url') ?? 'http://localhost:3000', '/');
        $this->enabled = $enabled ?? (bool)(\Cake\Core\Configure::read('Gotenberg.enabled') ?? false);
    }

    public function getName(): string
    {
        return 'gotenberg';
    }

    public function isAvailable(): bool
    {
        if (!$this->enabled) {
            return false;
        }
        try {
            $client = new Client(['timeout' => 2]);
            $response = $client->get($this->baseUrl . '/health');
            return in_array($response->getStatusCode(), [200, 204], true);
        } catch (\Throwable $e) {
            return false;
        }
    }

    public function convert(string $inputPath, string $outputDir): string
    {
        if (!is_file($inputPath)) {
            throw new RuntimeException(sprintf('Fichier source introuvable : %s', $inputPath));
        }

        $client = new Client(['timeout' => 30]);
        
        // LA MÉTHODE CAKEPHP 5 : 
        // On passe simplement un tableau avec la ressource du fichier.
        // CakePHP détecte automatiquement le 'fopen' et encode la requête en multipart/form-data.
        // La clé 'files' correspond exactement à ce qu'attend l'API de Gotenberg.
        $data = [
            'files' => fopen($inputPath, 'r')
        ];

        try {
            $response = $client->post(
                $this->baseUrl . '/forms/libreoffice/convert', 
                $data
            );
        } catch (\Throwable $e) {
            throw new RuntimeException(
                sprintf('Impossible de joindre Gotenberg (%s) : %s', $this->baseUrl, $e->getMessage()),
                0,
                $e
            );
        }

        if ($response->getStatusCode() !== 200) {
            throw new RuntimeException(sprintf(
                'Gotenberg a renvoyé une erreur HTTP %d. Body : %s',
                $response->getStatusCode(),
                substr((string)$response->getBody(), 0, 500)
            ));
        }

        $pdfContent = (string)$response->getBody();

        if (strlen($pdfContent) < 100) {
            throw new RuntimeException('Gotenberg a renvoyé un PDF vide ou corrompu.');
        }

        $baseName = pathinfo($inputPath, PATHINFO_FILENAME);
        $pdfPath = $outputDir . DIRECTORY_SEPARATOR . $baseName . '.pdf';
        
        file_put_contents($pdfPath, $pdfContent);

        if (!is_file($pdfPath)) {
            throw new RuntimeException('Écriture du PDF Gotenberg impossible dans : ' . $outputDir);
        }

        return $pdfPath;
    }
}