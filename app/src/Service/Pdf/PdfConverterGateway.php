<?php

declare(strict_types=1);

namespace App\Service\Pdf;

use Cake\Log\Log;
use RuntimeException;

/**
 * Passerelle de conversion PDF.
 *
 * Tente chaque strategie dans l'ordre configure.
 * Si une strategie echoue ou est indisponible, on passe a la suivante.
 * Si toutes echouent, on leve une exception regroupee avec le detail.
 */
class PdfConverterGateway
{
    /** @var PdfConverterInterface[] */
    private array $strategies;

    /**
     * @param PdfConverterInterface[] $strategies Ordonnees par priorite.
     */
    public function __construct(array $strategies)
    {
        // On ne garde que les strategies valides
        $this->strategies = array_values(array_filter(
            $strategies,
            fn($s) => $s instanceof PdfConverterInterface
        ));

        if (empty($this->strategies)) {
            throw new RuntimeException('Aucune strategie de conversion PDF n\'a ete fournie a la passerelle.');
        }
    }

    /**
     * Fabrique la passerelle a partir de la config CakePHP.
     * Strategie par defaut : gotenberg -> libreoffice.
     */
    public static function fromConfig(): self
    {
        $order = \Cake\Core\Configure::read('Pdf.strategies') ?? ['gotenberg', 'libreoffice'];

        $map = [
            'gotenberg'   => new GotenbergService(),
            'libreoffice' => new LibreOfficeService(),
        ];

        $strategies = [];
        foreach ($order as $name) {
            if (isset($map[$name])) {
                $strategies[] = $map[$name];
            }
        }

        return new self($strategies);
    }

    public function convert(string $xlsxPath, string $outputDir): string
    {
        $errors = [];

        foreach ($this->strategies as $strategy) {
            $name = $strategy->getName();

            // Test rapide de disponibilite (evite de lancer un process lourd si le service est mort)
            if (!$strategy->isAvailable()) {
                $msg = sprintf('[%s] Service indisponible, on passe a la strategie suivante.', $name);
                Log::info($msg);
                $errors[$name] = 'indisponible';
                continue;
            }

            try {
                $pdfPath = $strategy->convert($xlsxPath, $outputDir);
                Log::info(sprintf('[%s] Conversion PDF reussie : %s', $name, $pdfPath));
                return $pdfPath;
            } catch (\Throwable $e) {
                $msg = sprintf('[%s] Echec de conversion : %s', $name, $e->getMessage());
                Log::warning($msg);
                $errors[$name] = $e->getMessage();
                // On continue vers la strategie suivante
            }
        }

        // Toutes les strategies ont echoue
        throw new RuntimeException(
            "Aucune strategie de conversion PDF n'a fonctionne. Details :\n" .
            implode("\n", array_map(
                fn($k, $v) => "  - $k : $v",
                array_keys($errors),
                array_values($errors)
            ))
        );
    }
}