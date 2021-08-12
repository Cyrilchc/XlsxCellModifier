### Variables à modifier ### 

# Chemin des fichiers xlsx
$path = "C:\temp\generatedXlsx"

$coorToCopy = @(
    "A1"
    "H5"
    "R9"
    "L11"
)

# Chemin du fichier xlsx à peupler
$targetPath = "C:\temp\target\target.xlsx"

# à partir de quelle ligne commencer à copier
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
        
        # Il faut changer de colonne à chaque itération pour ne pas écraser le précédent
        $column = 1
        foreach($cell in $coorToCopy)
        {
            $tries = 0 # Cette partie du code ne semble pas très fiable, je lui donne 5 tentatives pour réussir
            while ($tries -lt 5)
            {
                try
                {
                    # Sélection de la cellule à copier
                    $range = $worksheet.Range($cell)

                    # Copie la cellule
                    $range.Copy() | Out-Null

                    # Sélectionne la cellule cible
                    $Range = $worksheetTarget.Cells($targetOrdonnee, $column)

                    # Colle la cellule dans la cible
                    $worksheetTarget.Paste($range)

                    $column++

                    $tries = 5 # L'opération a réussi, on annule les essais
                }
                catch
                {
                    $tries++
                    if($tries -eq 5)
                    {
                        throw $_
                    }
                }
            }
        }

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