<#
Ce script récupère un chemin donné, récupère son contenu et obtiens les droits ACL de ceux-ci. 
Il exporte les résultats vers un fichier Excel
#>
# ------ VARIABLES -------
# Chemin pour tester l'arborescence
$sourcePath = Get-Content "..\Settings\source.txt"
# Récupération année et mois pour vérification des dossiers. 
$year = get-date -f yyyy # Année en chiffre
$monthName = (Get-culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetMonthName((get-date -f MM))) # Nom du mois
$monthNumber = get-date -f MM # Numéro du mois
$month = "$monthNumber" + "_" + "$monthName" # Égal à 01_Janvier, 02_Février, ...
$destinationPath = "..\Rapports\$year\$month\" # Chemin de destination pour enregistrer les rapports
# Nettoyage des erreurs pour la nouvelle session
$error.Clear()

# Suppression du fichier de log existant, s'il y en a un.
if(Test-path "..\logs\Errors_$(get-date -f dd-MM-yyyy).txt") {
    remove-item "..\logs\Errors_$(get-date -f dd-MM-yyyy).txt"
}
# Try pour gérer les erreurs en cas de problème
try {
    # Vérifie si le chemin mène quelque part
    if (Test-Path $sourcePath) {        
        # ---- Création des variables pour les données ---- 
        # Variables pour comparer les droits dossier parent/dossier actuel
        $storedParentACL = "" 
        $storedItemACL = ""  
        # Array pour y stocker les résultats
        $results = @()
        # Récupération de l'arborescence demandée
        $table = Get-ChildItem -Recurse -Path $sourcePath `        | Sort-Object FullName
        # Récupération des données pour afficher le chemin donné sur la première ligne
        $Path = (Get-Acl -Path $sourcePath) `        | Select-Object PSPath
        $Path = (Convert-Path $Path.PSPath)
        $ACL = (Get-Acl -Path $sourcePath).Access `        | Select-Object FileSystemRights, AccessControlType, IdentityReference
        # Ajoute le premier dossier
        $results += ($ACL | Add-Member -Name Path -value $Path -MemberType NoteProperty -PassThru `
        | Select-Object Path, FileSystemRights, AccessControlType, IdentityReference)
    
        # Pour chaque entrée dans l'arborescence
        foreach ($item in $table) {
	        # Enregistre les droits de chaque entrée, et des variables de comparaison pour les droits
	        $ACLs = (Get-Acl $item.FullName).Access 
            $StoredParentACL = (Get-ACL $item.PSParentPath).Access `
            | Select-Object IdentityReference
            $StoredItemACL = (Get-Acl $item.FullName).Access `
            | Select-Object IdentityReference
	
            # Si l'entrée est un fichier
            if ((Get-Item $item.FullName) -is [System.IO.FileInfo]) {
                # Si la comparaison des droits entre le dossier parent et le fichier actuel, afficher juste le chemin
                if ((Compare-Object $StoredParentACL $StoredItemACL) -eq $null) {
                    $results += $ACLs `                    | Add-Member -Name Path -Value $item.FullName -MemberType NoteProperty -PassThru `                    | Select-Object Path
                # Sinon, afficher le chemin, l'AccessControleType (Access/Deny), 
                # le FileSystemRights (FullControl, Read/write, ..) et les utilisateurs/groupes concernés
                } else {
                    $results += $ACLs `                    | Add-Member -Name Path -Value $item.FullName -MemberType NoteProperty -PassThru `
                    | Select-Object Path, AccessControlType, FileSystemRights, IdentityReference
                }
            # Si l'entrée est un dossier, afficher le chemin, l'AccessControleType (Access/Deny), 
            # le FileSystemRights (FullControl, Read/write, ..) et les utilisateurs/groupes concernés
            } else {
                $results += $ACLs `                | Add-Member -Name Path -Value $item.FullName -MemberType NoteProperty -PassThru `                | Select-Object Path, AccessControlType, FileSystemRights, IdentityReference
            }            
        }

        # Vérifie si le dossier de sauvegarde existe, le crée si ce n'est pas le cas
        if (!(Test-Path $destinationPath)) {
            if (!(Test-Path "..\Rapports\$year\")) {
                mkdir "..\Rapports\$year\"
            } elseif (!(Test-Path ".\Rapports\$year\$month")) {
                mkdir "..\Rapports\$year\$month\"
            }
        }

        # Vérifier si le fichier existe déjà, le supprime si c'est le cas
        if (Test-Path "..\Rapports\$year\$month\results1_$(get-date -f dd-MM-yyyy).xlsx") {
            Remove-Item "..\Rapports\$year\$month\results1_$(get-date -f dd-MM-yyyy).xlsx"
        }
        # Vérifie si le fichier logs existe, supprime les deux fichiers logs si c'est le cas 
        if (Test-Path "..\logs\ErrorSource_$(get-date -f dd-MM-yyyy).txt") {
            Remove-Item "..\logs\ErrorSource_$(get-date -f dd-MM-yyyy).txt"
            Remove-Item "..\Rapports\ErrorSource_$(get-date -f dd-MM-yyyy).txt"
        }
        # Exporte le résultat sur le fichier en question
        $results | Export-Excel -Path "..\Rapports\$year\$month\results1_$(get-date -f dd-MM-yyyy).xlsx"
        # Appelle le script suivant 
        Invoke-Expression "..\Scripts\EraseUseless.ps1"
        
        

    } else { 
        # Efface le fichier de logs s'il existe et exporte la phrase comme quoi le chemin n'existe pas vers un fichier
        if (Test-Path "..\logs\ErrorSource_$(get-date -f dd-MM-yyyy).txt") {
            Remove-Item "..\logs\ErrorSource_$(get-date -f dd-MM-yyyy).txt"
            Remove-Item "..\Rapports\ErrorSource_$(get-date -f dd-MM-yyyy).txt"
        }
        "The source's path $sourcePath doesn't exists." > "..\logs\ErrorSource_$(get-date -f dd-MM-yyyy).txt"
        Copy-Item -Path .\logs\ErrorSource.txt -Destination "..\Rapports\ErrorSource_$(get-date -f dd-MM-yyyy).txt"
    }
    
} catch {    
    $error > "..\logs\Errors_$(get-date -f dd-MM-yyyy).txt"
}

