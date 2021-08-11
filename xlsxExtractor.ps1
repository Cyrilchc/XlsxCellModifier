### Variables à modifier ### 

# Chemin des fichiers xlsx
$path = "C:\temp\generatedXlsx"

# Coordonnées de la cellule à copier
# Pour rechercher plutôt par texte et non pas par coordonnées :
# https://stackoverflow.com/a/59408529
$abscisse = 3
$ordonnee = 5

# Chemin du fichier xlsx à peupler
$targetPath = "C:\temp\target\target.xlsx"

# Coordonnées des cellules à peupler
$targetAbscisse = 1
$targetOrdonnee = 1

# Feuille du fichier cible à sélectionner
$sheetTarget = 1

# Feuille à sélectionner
$sheet = 1

# Afficher Excel ATTENTION => Il faut un PC puissant
$displayExcel = $true

# Laisser Excel ouvert => Idem consomme beaucoup de mémoire
$keepOpened = $true

### Démarrage du script ###
Clear-Host

Add-Type -AssemblyName Microsoft.Office.Interop.Excel

### Ouverture du fichier cible

# Création de l'objet Excel
$excelTarget = New-Object -ComObject Excel.Application
$excelTarget.Visible = $displayExcel
$excelTarget.ScreenUpdating = $displayExcel

# Ouverture du fichier et sélection de la feuille
$workbookTarget = $excelTarget.Workbooks.Open($targetPath)
$worksheetTarget = $workbookTarget.WorkSheets.item($sheetTarget)
[void]$worksheetTarget.Activate()


# Parcours les fichiers du répertoire
foreach($file in Get-ChildItem -Path $path)
{
    try
    {
        # Création de l'objet Excel
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = !$displayExcel
        $excel.ScreenUpdating = !$displayExcel

        # Ouverture du fichier et sélection de la feuille
        $workbook = $excel.Workbooks.Open($file.FullName)
        $worksheet = $workbook.WorkSheets.item($sheet)
        [void]$worksheet.Activate()

        # Sélection de la cellule à copier
        $targetValue = $worksheet.Cells.Item($ordonnee, $abscisse).Value2

        # Copie de la valeur dans le fichier cible
        $worksheetTarget.Cells.Item($targetOrdonnee, $targetAbscisse) = $targetValue

        # Libération des ressources du fichier courant
        if(!$keepOpened)
        {
            [void]$workbook.Close()
            [void]$excel.Quit()  
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }

        # Incrémentation de l'abscisse du fichier cible pour écrire sur la ligne suivante
        $targetOrdonnee++
    }
    catch
    {
        Write-Host "Erreur sur le fichier $($file.FullName). Consultez l'erreur pour en savoir plus." -ForegroundColor Cyan
        Write-Host $_ -ForegroundColor Red
    }
}

# Fin de script

# Enregistrement du fichier cible
[void]$workbookTarget.Save()