<?php


# Inclure la bibliothèque phpSpreadsheet
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Cell;



function convertCsvFilesToExcel($csvFolderPath)
{

    # Génération du nom de notre dossier
    $dateHeure = date('Y-m-d_H-i-s'); # Format: YYYY-MM-DD_HH-MM-SS
    $excelFolderPath = basename($csvFolderPath) . "_to_excel" . "_$dateHeure";

    # Vérifions si le dossier CSV existe
    if (!is_dir($csvFolderPath)) {
        die("Le dossier CSV '$csvFolderPath' n'existe pas.");
    }

    # Nom du dossier comportant nos convertion
    // $excelFolderPath = 

    # Créer le dossier Excel s'il n'existe pas
    if (!is_dir($excelFolderPath)) {
        mkdir($excelFolderPath, 0777, true);
    }

    # Récupérer la liste des fichiers CSV dans le dossier
    $csvFiles = glob($csvFolderPath . "/*.csv");

    if (empty($csvFiles)) {
        die("Aucun fichier CSV trouvé dans le dossier '$csvFolderPath' .");
    }

    foreach ($csvFiles as $csvFile) {
        $excelFileName = pathinfo($csvFile, PATHINFO_FILENAME);

        # Convertir chaque fichier CSV en fichier Excel
        convertCsvToExcel($csvFile, $excelFolderPath, $excelFileName);
    }

    echo "La conversion des fichiers CSV en fichiers Excel est terminée.";
}


function convertCsvToExcel($csvFilePath, $excelFolderPath, $excelFileName)
{

    # Créer un nouvel objet Spreadsheet
    $spreadsheet = new Spreadsheet();

    try {
        # Créer un lecteur CSV personnalisé
        $csvReader = PhpOffice\PhpSpreadsheet\IOFactory::createReader('Csv');
        $csvReader->setDelimiter(';'); # Définir le délimiteur personnalisé
        $csvReader->setInputEncoding('UTF-8'); # Définir l'encodage du fichier CSV
        $csvReader->setReadDataOnly(true); // Ne pas essayer d'inférer automatiquement le type de données

        # Charger les données du fichier CSV dans un tableau
        $lines = file($csvFilePath, FILE_IGNORE_NEW_LINES);

        # Parcourir chaque ligne du fichier CSV
        foreach ($lines as &$line) {
            # Diviser la ligne en valeurs individuelles en utilisant le délimiteur ;
            $values = explode(';', $line);

            # Parcourir chaque valeur dans la ligne
            foreach ($values as &$value) {
                # Ajouter d'un espace au début de chaque valeur
                $value = " $value";
            }

            # Réassembler les valeurs modifiées en une ligne
            $line = implode(';', $values);
        }

        # Réassembler toutes les lignes en une seule chaîne
        $csvData = implode("\n", $lines);

        # Créer un fichier temporaire avec les données modifiées
        $tempCsvFilePath = tempnam(sys_get_temp_dir(), 'csv');
        file_put_contents($tempCsvFilePath, $csvData);

        # Charger les données du fichier CSV dans la feuille de calcul
        $spreadsheet = $csvReader->load($tempCsvFilePath);

        # Supprimer le fichier temporaire
        unlink($tempCsvFilePath);

    } catch (\Exception $e) {
        die("Erreur lors de la lecture du fichier CSV : " . $e->getMessage());
    }


    # Créer un objet Writer pour XLSX (Excel)
    $writer = new Xlsx($spreadsheet);

    # Chemin du dossier où vous souhaitez enregistrer le fichier Excel
    # Nom du fichier Excel à générer avec la date et l'heure actuelle

    $excelFilePath = $excelFolderPath . '/' . "$excelFileName" . ".xlsx";

    try {
        # Enregistrement du fichier excel
        $writer->save($excelFilePath);
        echo "Le fichier excel '$excelFilePath' a été généré avec succès à partir du fichier CSV '$csvFilePath' ! <br>";
    } catch (\Exception $e) {
        die("Erreur lors de l'écriture du fichier Excel : " . $e->getMessage());
    }
}

# Utilisation de la fonction pour convertir tous les fichiers csv d'un dossier en Excel
$csvFolderPath = "dossier_csv"; # Chemin vers dossier csv
$excelFolderPath = "dossier_excel"; # Chemin vers le dossier excel
convertCsvFilesToExcel($csvFolderPath);
?>