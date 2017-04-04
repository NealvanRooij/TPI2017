<#
Ce script récupère les informations sur le second fichier pour n'afficher que les lignes 
qui contiennent des données. 
#>
# ------ VARIABLES -------
# Récupération année et mois pour vérification des dossiers. 
$year = get-date -f yyyy
$monthName = (Get-culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetMonthName((get-date -f MM)))
$monthNumber = get-date -f MM
$month = "$monthNumber" + "_" + "$monthName"
# Récupération des données sur le fichier Excel 
# -DataOnly permet de ne récupérer que les cellules qui contiennent des données
$fileContent = Import-Excel "..\Rapports\$year\$month\results2_$(get-date -f dd-MM-yyyy).xlsx" -DataOnly
# Array pou y ajouter les résultats à ajouter par la suite. 
$array = @()

try {
    # Pour chaque ligne dans les données importéesm sélectionner les données et les sortir.
    foreach ($item in $fileContent) {
        $a = $item | Select-Object Path, AccessControlType, FileSystemRights, IdentityReference
        $array += $item    
    }

    if (Test-Path "..\Rapports\$year\$month\results_$(get-date -f dd-MM-yyyy).xlsx") {
        Remove-Item "..\Rapports\$year\$month\results_$(get-date -f dd-MM-yyyy).xlsx"
    }

    $array | Export-Excel "..\Rapports\$year\$month\results_$(get-date -f dd-MM-yyyy).xlsx" `    -WorkSheetname "Rapport" -AutoSize
    
} catch {
    $error > "..\logs\Errors_$(get-date -f dd-MM-yyyy).txt"
}

$error > "..\logs\Errors_$(get-date -f dd-MM-yyyy).txt"


Remove-Item  "..\Rapports\$year\$month\results1_$(get-date -f dd-MM-yyyy).xlsx"
Remove-Item  "..\Rapports\$year\$month\results2_$(get-date -f dd-MM-yyyy).xlsx"
Invoke-Expression ".\GetGroupsArborescence.ps1"

