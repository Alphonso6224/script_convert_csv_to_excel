L'utilisation du script de conversion de fichiers CSV en fichiers Excel (XLSX) :

---

# Conversion de fichiers CSV en fichiers Excel

Ce script PHP permet de prendre un dossier et de convertir plusieurs fichiers CSV en fichiers Excel (XLSX) dans un dossier spécifié.

## Prérequis

- PHP version 7.0 ou ultérieure
- Extension PHP `php-mbstring` installée (si vous utilisez PhpSpreadsheet)

## Installation

1. Clonez ce dépôt ou téléchargez le script `convertCsvToExcel.php`.
2. Assurez-vous d'avoir installé Composer sur votre système.
3. Exécutez `composer install` pour installer les dépendances nécessaires.

## Utilisation

1. Placez vos fichiers CSV dans un dossier.
2. Ouvrez le fichier `convert_csv_to_Excel.php` et configurez la variable `$csvFolderPath` avec le chemin du dossier contenant les fichiers CSV 
3. Exécutez le script en ligne de commande ou via un serveur web.

Exemple d'utilisation en ligne de commande :

```bash
php convert_csv_to_Excel.php
```

## Configuration

Dans le fichier `convert_csv_to_Excel.php`, vous pouvez modifier les options suivantes :

- **Délimiteur CSV** : Modifiez la variable `$delimiter` pour définir le délimiteur des fichiers CSV.
- **Encodage CSV** : Modifiez la variable `$encoding` pour définir l'encodage des fichiers CSV.

---