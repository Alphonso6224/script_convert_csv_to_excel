<?php


# Inclure la bibliothèque phpSpreadsheet
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\IOFactory;

function convertCsvToExcel($csvFilePath) {

    # Vérifier si le fichier CSV existe
    if (!file_exists($csvFilePath)) {
        die("Le fichier Csv '$csvFilePath' n'existe pas.");
    }
    
    # Récupérer le nom du fichier san l'extension
    $excelFileName = pathinfo($csvFilePath, PATHINFO_FILENAME);
    
    # Créer un nouvel objet Spreadsheet
    $spreadsheet = new Spreadsheet();
    
    try {
        # Créer un lecteur CSV personnalisé
        $csvReader = PhpOffice\PhpSpreadsheet\IOFactory::createReader('Csv');
        $csvReader->setDelimiter(';'); # Définir le délimiteur personnalisé
        $csvReader->setInputEncoding('UTF-8'); # Définir l'encodage du fichier CSV

        # Charger les données du fichier CSV dans la feuille de calcul
        $spreadsheet = $csvReader->load($csvFilePath);
    } catch (\Exception $e) {
        die("Erreur lors de la lecture du fichier CSV : " . $e->getMessage());
    }
    
    
    # Créer un objet Writer pour XLSX (Excel)
    $writer = new Xlsx($spreadsheet);
    
    # Chemin du dossier où vous souhaitez enregistrer le fichier Excel
    $folderPath = 'src/fichier_excel/';

    # Nom du fichier Excel à générer avec la date et l'heure actuelle
    $dateHeure = date('Y-m-d_H-i-s'); # Format: YYYY-MM-DD_HH-MM-SS
    $excelFilePath = $folderPath . "$excelFileName" . "_$dateHeure.xlsx";

    try {
        # Enregistrement du fichier excel
        $writer->save($excelFilePath);
        echo "Le fichier excel '$excelFilePath' a été généré avec succès à partir du fichier CSV '$csvFilePath' !";
    } catch(\Exception $e) {
        die("Erreur lors de l'écriture du fichier Excel : " . $e->getMessage());
    }
}

# Utilisation de la fonction pour convertir un fichier csv en Excel
$csvFilePath = 'affectations_chef_plateaux.csv';
convertCsvToExcel($csvFilePath);
?>