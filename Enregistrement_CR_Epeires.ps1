# Script automatisant l'enregistrement des comptes-rendus quotidiens d'EPEIRES
# version : 1.0 ( 16 janvier 2018 )
# par : Aurelien LOHO ( ST / QST )

# Constantes de connexion à EPEIRES
$epeiresUrl = 'http://epeires.cdg-lb.aviation'
$userName = 'sauvegarde'
$password = 'epeires'

# Table de correspondance entre le mois et le nom du répertoire de sauvegarde
# le tableau commence à l'index 0, d'où le premier élément vide
$monthNameLookup = @('', '01-JANVIER', '02-FEVRIER', '03-MARS', '04-AVRIL', '05-MAI', '06-JUIN', '07-JUILLET', '08-AOUT', '09-SEPTEMBRE', '10-OCTOBRE', '11-NOVEMBRE', '12-DECEMBRE')

# Variables nécessaires pour générer le nom de fichier à partir de la date du rapport
$yesterday = (Get-Date).AddDays(-1) # la date du rapport que l'on souhaite sauvegarder (toujours celui de la veille pour ce script)
$reportYear = $yesterday.Year
$reportMonth = $monthNameLookup[$yesterday.Month]
$reportFileName = Get-Date $yesterday -UFormat 'rapport_du_%d_%m_%Y.pdf'

# le répertoire de sauvegarde dépend de la date et est de la forme : YYYY\MM-MMMM (par exemple 2018\01-JANVIER)
$reportFullPath = "L:\SE\01.TEMPS_REEL\8.PV\$reportYear\$reportMonth\$reportFileName"

if(!(Test-Path $reportFullPath)) # envoi les requêtes que si le rapport n'a pas déjà été sauvegardé
{
    # Envoi de la requête pour se connecter à EPEIRES
    $connexionParameters = @{identity=$userName;credential=$password;redirect='application';submit=''}
    $loginUrl = "$epeiresUrl/user/login"
    Invoke-WebRequest $loginUrl -SessionVariable epeiresSession -Method Post -Body $connexionParameters | Out-Null

    # Envoi de la requête pour télécharger le compte-rendu et enregistre le résultat dans le fichier spécifié
    $reportDate = Get-Date $yesterday -format r # EPEIRES attend la date au format RFC1123
    $saveReportUrl = "$epeiresUrl/report/daily?day=$reportDate"
    Invoke-WebRequest $saveReportUrl -WebSession $epeiresSession -Method Get -OutFile $reportFullPath | Out-Null   
}