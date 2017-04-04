<#
Ce script récupère les droits de l'arborescence et si un groupe est listé dans les droits, 
les utilisateurs de ce groupe sont alors listés sur une secodne feuille sur le même document.
#>
# ------ VARIABLES -------
# Récupération du chemin de l'arborescence.
$sourcePath = Get-Content "..\Settings\source.txt"
# Récupération des droits sur l'arborescence.
$arborescence = Get-ChildItem $sourcePath -Recurse | Select-Object FullName
# Nettoyage des erreurs pour avoir un rapport ne concenrant que les utlisateurs et groupes. 
$error.Clear()
# Array pour stocker les différents groupes présents dans le résultat
$groups = @()
"hllo"
foreach ($line in $arborescence) {
    $acls = (get-ACL $line.FullName).Access
    foreach ($a in $acls) {       
       try {    
          # Sépare la ligne en deux par un \ et sélectionne le 
          #second bloc de résultat avec [1], soit le nom du groupe 
          #ou de l'utilisateur
          $entry = ($a.identityReference.toString().split("\")[1])   
          # Enregistre l'entrée s'il n'est pas encore présent dans 
          #l'Array $groups        
          if (!($groups -contains $entry)) {
            $groups += $entry            
          }
       } catch {          
          continue
       }
    }
}
# Array pour stocker les différents utilisateurs
$groupsOutput = @()
$groupsOutput = foreach ($group in $groups) {
    try {
        $results = Get-ADGroupMember -Identity $group.ToString() `        | Get-ADUser -Properties displayname, ObjectClass, name, samAccountName
        # Pour chaque entrée de membres de groupes dans $results, 
        # récupérer les informations sur les utilisateurs
        # et stockage de ces inforamtions dans un object pour exporter le tout
        foreach ($result in $results) {
            New-Object PSObject -Property @{
                GroupeName = $group
                Username = $result.Name
                SAN = $result.SamAccountName
                DisplayName = $result.DisplayName                
            }
        }  
    } catch {
        continue
    } 
}
# Si la variable contenant toutes les erreurs est plsu longue que 1, alors il y a une erreure, donc l'afficher.
if ($error.length -gt 0) {
    $error > "..\logs\ErrorsAD_$(get-date -f dd-MM-yyyy).txt"
}

# Exporter le résultat en sélectionnant l'ordre des données, triant par notre de groupe et d'utilisateur
$groupsOutput | Select-Object GroupeName, UserName, SAN, DisplayName `| Sort-Object GroupeName, UserName `| Export-Excel "..\Rapports\$year\$month\results_$(get-date -f dd-MM-yyyy).xlsx" -WorkSheetname Droits -AutoSize

