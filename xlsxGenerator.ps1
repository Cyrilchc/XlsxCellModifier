### Fonctions ###
function Install-Dependances 
{
    param (
        $ModuleName
    )

    if(!(Get-Module -ListAvailable -Name $ModuleName))
    {
        Write-Host "Module manquant. D�marrage de l'installation du module $ModuleName..." -ForegroundColor Yellow
        
        try
        {
            Install-Module $ModuleName
        }
        catch
        {
            Write-Host "Impossible de t�l�charger les d�pendances. R�essayez en mode administrateur si vous ne l'�tes pas d�j�." -ForegroundColor Cyan
            Write-Host $_ -ForegroundColor Cyan
            exit
        }
    }
}

### Variables � modifier ###

# Combien de fichiers � cr�er ?
$howManyToCreate = 10

# R�pertoire ou g�n�rer les fichiers xlsx, Pr�suppose que ce r�pertoire existe 
$rootPath = "C:\temp\generatedXlsx"

# Dans cet exemple, on utilise le contenu du disque C pour simuler des donn�es � ins�rer dans les fichiers xlsx
# Pour ce faire, on se place simplement dans le disque C
Set-Location "C:\"

# D�marrage du script
Clear-Host

# Installe le module n�cessaire
$module = "PSExcel"
Install-Dependances -ModuleName $module

# Importe le module
Import-Module $module 

# On boucle 10 fois
for ($i = 0; $i -lt $howManyToCreate; $i++)
{
    # Nom du fichier � cr�er
    $fileName = "xlsx$i.xlsx"

    # Chemin du fichier � cr�er
    $filePath = Join-Path -Path $rootPath -ChildPath $fileName

    # Export du contenu du r�pertoire dans lequel on se trouve dans le fichier
    try
    {
        #Get-ChildItem | Export-Excel $filePath
        Get-ChildItem | Export-XLSX $filePath
        Write-Host "$fileName cr�� dans le r�pertoire $rootPath" -ForegroundColor Green
    }
    catch
    {
        Write-Host "Impossible de cr�er le fichier $fileName dans le r�pertoire $rootPath. Consultez l'erreur pour en savoir plus." -ForegroundColor Cyan
        Write-Host $_ -ForegroundColor Red
    } 
}

explorer.exe $rootPath