<?php

declare(strict_types=1);

namespace App\Service\Pdf;

/**
 * Contrat commun a tous les convertisseurs XLSX -> PDF.
 */
interface PdfConverterInterface
{
    /**
     * Verifie rapidement si le service est operationnel
     * (sans lancer une vraie conversion).
     */
    public function isAvailable(): bool;

    /**
     * Convertit un XLSX en PDF.
     *
     * @param string $xlsxPath  Chemin absolu du fichier source.
     * @param string $outputDir Dossier ou ecrire le PDF.
     * @return string Chemin absolu du PDF genere.
     * @throws \RuntimeException Si la conversion echoue.
     */
    public function convert(string $xlsxPath, string $outputDir): string;

    /**
     * Nom de la strategie (pour le log).
     */
    public function getName(): string;
}