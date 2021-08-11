### Variables à modifier ### 
# Chemin des fichiers xlsx
$path = "C:\temp\generatedXlsx"

# Coordonnées de la cellule à modifier ici c'est A1
$abscisse = 1
$ordonnee = 1

# Feuille à sélectionner
$sheet = 1

# Valeur à donner à la cellule
$valeur = "foobar"

# Afficher Excel ATTENTION => Il faut un PC puissant
$displayExcel = $true

# Laisser Excel ouvert
$keepOpened = $true

### Démarrage du script ###
Clear-Host

Add-Type -AssemblyName Microsoft.Office.Interop.Excel

# Parcours les fichiers
foreach($file in Get-ChildItem -Path $path)
{
    try
    {
        # Création de l'objet Excel
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $displayExcel
        $excel.ScreenUpdating = $displayExcel

        # Ouverture du fichier et sélection de la feuille
        $workbook = $excel.Workbooks.Open($file.FullName)
        $worksheet = $workbook.WorkSheets.item($sheet)
        [void]$worksheet.Activate()

        # Modification de la valeur de la cellule
        $worksheet.Cells.Item($abscisse, $ordonnee) = $valeur

        # Enregistrement du fichier et libération des ressources
        [void]$workbook.Save()
        if(!$keepOpened)
        {
            [void]$workbook.Close()
            [void]$excel.Quit()  
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    }
    catch
    {
        Write-Host "Erreur sur le fichier $($file.FullName). Consultez l'erreur pour en savoir plus." -ForegroundColor Cyan
        Write-Host $_ -ForegroundColor Red
    }
}
