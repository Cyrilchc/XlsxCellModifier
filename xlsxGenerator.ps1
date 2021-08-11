### Fonctions ###
function Install-Dependances 
{
    param (
        $ModuleName
    )

    if(!(Get-Module -ListAvailable -Name $ModuleName))
    {
        Write-Host "Module manquant. Démarrage de l'installation du module $ModuleName..." -ForegroundColor Yellow
        
        try
        {
            Install-Module $ModuleName
        }
        catch
        {
            Write-Host "Impossible de télécharger les dépendances. Réessayez en mode administrateur si vous ne l'étes pas déjé." -ForegroundColor Cyan
            Write-Host $_ -ForegroundColor Cyan
            exit
        }
    }
}

### Variables é modifier ###

# Combien de fichiers é créer ?
$howManyToCreate = 10

# Répertoire ou générer les fichiers xlsx, Présuppose que ce répertoire existe 
$rootPath = "C:\temp\generatedXlsx"

# Dans cet exemple, on utilise le contenu du disque C pour simuler des données é insérer dans les fichiers xlsx
# Pour ce faire, on se place simplement dans le disque C
Set-Location "C:\"

# Démarrage du script
Clear-Host

# Installe le module nécessaire
$module = "PSExcel"
Install-Dependances -ModuleName $module

# Importe le module
Import-Module $module 

# On boucle 10 fois
for ($i = 0; $i -lt $howManyToCreate; $i++)
{
    # Nom du fichier é créer
    $fileName = "xlsx$i.xlsx"

    # Chemin du fichier à créer
    $filePath = Join-Path -Path $rootPath -ChildPath $fileName

    # Export du contenu du répertoire dans lequel on se trouve dans le fichier
    try
    {
        #Get-ChildItem | Export-Excel $filePath
        Get-ChildItem | Export-XLSX $filePath
        Write-Host "$fileName créé dans le répertoire $rootPath" -ForegroundColor Green
    }
    catch
    {
        Write-Host "Impossible de créer le fichier $fileName dans le répertoire $rootPath. Consultez l'erreur pour en savoir plus." -ForegroundColor Cyan
        Write-Host $_ -ForegroundColor Red
    } 
}

explorer.exe $rootPath