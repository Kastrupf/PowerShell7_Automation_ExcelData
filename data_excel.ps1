# Le raccourci clavier pour commenter plusieurs lines : shift + alt + A


<# Vérification si le Module "Import-Excel" est déjà instalé, sinon le faire à force, sans demander "YES" #> 
if (!(Get-Module Import-Excel)) {
    Install-Module ImportExcel -Force -Scope CurrentUser
}
else {
    Write-Output "Le module est déjà installé"
}


<# Récuperation de tous les fichiers qui sont dans le répertoire "Data" et attibuition du résultat à la nouvelle variable créé $csvs #>
$csvs = Get-ChildItem -Path "C:\CBTNuggets\cours_pxsh7_manip_data\Data"


<# Pour chaque variable $csv qui se trouuve dans la collection $csvs, importer son nom complet et l'exporter comme fichier du type Excel, en créant une variable qui comporte le même nom du fichier d'origine, mais avec l'extension ".xlsx" #>
ForEach($csv in $csvs){
    Import-Csv -Path $csv.FullName | Export-Excel -Path "C:\CBTNuggets\cours_pxsh7_manip_data\Data\$($csv.BaseName).xlsx"
} 


<# Jonction de deux collections de données - characters et planets dans un seul résultat.
Import des objets du type "character" à partir du fichier "characters.xlsx". 
Selection des proprietés name, eye_color, homeworld et attibuition du résultat à la nouvelle variable créé $chars 
Du même avec "planets" #>
$chars = Import-Excel -Path "C:\CBTNuggets\cours_pxsh7_manip_data\Data\characters.xlsx" | Select-Object name, eye_color, homeworld
$planets = Import-Excel -Path "C:\CBTNuggets\cours_pxsh7_manip_data\Data\planets.xlsx"


<# Pour chaque variable $char qui se trouuve dans la collection $chars et, pour chaque variable $planet qui se trouuve dans la collection $planets, si le "homeworld" de $char est égual au "name" du planet, alors commence pour prendre chaque objet du type character et ajoute-le une nouvelle proprieté qui s'appele "rotation" qui, à son tour, correspond à la proprieté "rotation_period" qui vient de l'objet planet
-Force est là pour ne pas demander une confirmation au terminal. #>
foreach($char in $chars){
    foreach($planet in $planets){
        if($char.homeworld -eq $planet.name){
            $char | Add-Member -MemberType NoteProperty -Name rotation -Value $planet.rotation_period -Force  
        }
    }
}


<# A la fin, récupere le résultat e l'export pour un nouveau fichier "output.xlsx" qui contien 4 proprietés : name, eye_color, homeworld et rotation  #>
$chars | Export-Excel -Path "C:\CBTNuggets\cours_pxsh7_manip_data\Data\output.xlsx"

