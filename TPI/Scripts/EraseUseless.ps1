<# 
Ce script a pour but de récupérer les données du fichier Excel créé par GetACL.ps1
et de remplacer toutes les lignes redondantes par un espace, afin de rendre la case vide. 
#>
# ------ VARIABLES -------
# Récupération année et mois pour vérification des dossiers. 
$year = get-date -f yyyy # Année en chiffre
$monthName = (Get-culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetMonthName((get-date -f MM))) # Nom du mois
$monthNumber = get-date -f MM # Numéro du mois
$month = "$monthNumber" + "_" + "$monthName" # Égal à 01_Janvier, 02_Février, ...
# Récupération des données sur le fichier Excel
$fileContent = import-excel "..\Rapports\$year\$month\results1_$(get-date -f dd-MM-yyyy).xlsx"
# Variables pour comparer les données supprimables
$storedPath = ""
$storedFileSystemRights =""
# Variable pour chaque première ligne de chaque entrée
$firstLine = 0

try {
    # Pour chaque ligne du fichier Excel
    foreach ($item in $fileContent) { 
        # Si le chemin enregistré n'est pas égal à celui de l'objet actuel, l'enregistrer
        # et afficher les informations normalement  
        if ($storedPath -ne $item.Path) {
            $storedPath = $item.Path
            $storedFileSystemRights = $item.FileSystemRights
            $firstLine = "1"
        } else {
        # Sinon, effacer le chemin       
            $item.Path = ""                
        }
        # Si les types de droits comparés ne sont pas égaux, 
        # ou s'il s'agit de la première ligne d'un fichier,
        # afficher les informations normalement
        if ($storedFileSystemRights -ne $item.FileSystemRights -or $firstLine -eq "1") {
            $storedFileSystemRights = $item.FileSystemRights
            $firstLine = "0"
        } elseif ($storedFileSystemRights -eq $item.FileSystemRights) {
        # Sinon, les effacer
            $item.FileSystemRights = ""
        }      
    }
    
    if (Test-Path "..\Rapports\$year\$month\results2_$(get-date -f dd-MM-yyyy).xlsx") {
        remove-item "..\Rapports\$year\$month\results2_$(get-date -f dd-MM-yyyy).xlsx"
    }
    
    $fileContent | Export-Excel "..\Rapports\$year\$month\results2_$(get-date -f dd-MM-yyyy).xlsx" -AutoSize
    Invoke-Expression ".\Cleaning.ps1"
} catch {    
    $error > "..\logs\Errors_$(get-date -f dd-MM-yyyy).txt"
    
}

