[CmdletBinding()]
param(
    [string]$Dossier = $PSScriptRoot,
    [switch]$SansMessage
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# Le lanceur CMD peut transmettre un guillemet parasite lorsqu'un chemin cité
# se termine par un antislash. Nettoyer puis normaliser le dossier avant toute
# utilisation évite l'erreur Test-Path « Caractères non conformes ».
$Dossier = ([string]$Dossier).Trim().Trim('"')
try {
    $Dossier = [IO.Path]::GetFullPath($Dossier)
}
catch {
    throw "Chemin de dossier non valide : $Dossier"
}

$RawBase = "https://raw.githubusercontent.com/Nathan20202/excel-catalogues-auto-update/main"
$ConfigUrl = "$RawBase/config/workbooks.json"
$LogPath = Join-Path $Dossier "Actualisation.log"
$BackupFolder = Join-Path $Dossier "_sauvegardes_actualisation"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    Write-Host $line
}

function Show-Result {
    param([string]$Message, [string]$Title, [string]$Icon = "Information")
    if ($SansMessage) {
        return
    }
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $iconValue = [System.Enum]::Parse([System.Windows.Forms.MessageBoxIcon], $Icon)
        [System.Windows.Forms.MessageBox]::Show(
            $Message,
            $Title,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            $iconValue
        ) | Out-Null
    }
    catch {
        Write-Host $Message
    }
}

function Get-RemoteJson {
    param([string]$Url)
    $separator = if ($Url.Contains("?")) { "&" } else { "?" }
    $cacheBuster = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    Invoke-RestMethod -Uri "$Url${separator}v=$cacheBuster" -Method Get -TimeoutSec 120 -UseBasicParsing
}

function Get-OpenWorkbook {
    param($Excel, [string]$FullPath)
    for ($index = 1; $index -le $Excel.Workbooks.Count; $index++) {
        $candidate = $Excel.Workbooks.Item($index)
        if ([string]::Equals(
                [IO.Path]::GetFullPath([string]$candidate.FullName),
                [IO.Path]::GetFullPath($FullPath),
                [StringComparison]::OrdinalIgnoreCase
            )) {
            return $candidate
        }
    }
    return $null
}

function Get-HeaderMap {
    param($Table)
    $map = @{}
    for ($column = 1; $column -le $Table.ListColumns.Count; $column++) {
        $name = [string]$Table.ListColumns.Item($column).Name
        $map[$name] = $column
    }
    return $map
}

function Get-PropertyValue {
    param($Record, [string]$Name)
    $property = $Record.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }
    return $property.Value
}

function Test-EqualCellValue {
    param($Left, $Right)
    if ($null -eq $Left -and $null -eq $Right) {
        return $true
    }
    return ([string]$Left).Trim() -eq ([string]$Right).Trim()
}

function ConvertTo-ExcelCellValue {
    param($Value)

    if ($null -eq $Value) {
        return $null
    }
    # Invoke-RestMethod de Windows PowerShell 5.1 deserialise certains nombres
    # JSON en System.Decimal. L'automatisation COM d'Excel n'accepte pas toujours
    # VT_DECIMAL pour Value2 ; un Double est le type numerique natif attendu.
    if ($Value -is [decimal]) {
        return [double]$Value
    }
    if ($Value -is [long] -or $Value -is [ulong]) {
        return [double]$Value
    }
    if ($Value -is [DateTime]) {
        return [double]$Value.ToOADate()
    }
    return $Value
}

function Set-WorksheetCellValue {
    param($Worksheet, [string]$Address, $Value)

    $range = $Worksheet.Range($Address)
    # Excel refuse l'ecriture dans une cellule fusionnee qui n'est pas la
    # cellule superieure gauche (HRESULT 0x800A03EC). Toujours viser la cellule
    # principale de la zone fusionnee rend le moteur compatible avec tous les
    # tableaux de bord.
    if ($range.MergeCells -eq $true) {
        $range = $range.MergeArea.Cells.Item(1, 1)
    }
    $range.Value2 = $Value
}

function Set-NewRowDefaults {
    param($Table, [int]$RowIndex, $Dataset, [hashtable]$HeaderMap)
    foreach ($property in $Dataset.defaultPersonalValues.PSObject.Properties) {
        $columnName = [string]$property.Name
        if (-not $HeaderMap.ContainsKey($columnName)) {
            continue
        }
        $cell = $Table.DataBodyRange.Cells.Item($RowIndex, $HeaderMap[$columnName])
        if ([string]::IsNullOrWhiteSpace([string]$cell.Value2)) {
            $cell.Value2 = ConvertTo-ExcelCellValue -Value $property.Value
        }
    }
    if ($RowIndex -gt 1) {
        foreach ($columnName in @($Dataset.formulaColumns)) {
            if (-not $HeaderMap.ContainsKey([string]$columnName)) {
                continue
            }
            $columnIndex = $HeaderMap[[string]$columnName]
            $source = $Table.DataBodyRange.Cells.Item($RowIndex - 1, $columnIndex)
            $target = $Table.DataBodyRange.Cells.Item($RowIndex, $columnIndex)
            if (-not [string]::IsNullOrWhiteSpace([string]$source.FormulaR1C1)) {
                $target.FormulaR1C1 = $source.FormulaR1C1
            }
        }
    }
}

function Update-Dataset {
    param($Workbook, $Dataset, [string]$RawBaseUrl)

    $payload = Get-RemoteJson -Url "$RawBaseUrl/$($Dataset.file)"
    if ([int]$payload.schemaVersion -ne 1) {
        throw "Version de schéma distante non prise en charge pour $($Dataset.file)."
    }
    $worksheet = $Workbook.Worksheets.Item([string]$Dataset.sheet)
    $table = $worksheet.ListObjects.Item([string]$Dataset.table)
    $headers = Get-HeaderMap -Table $table
    $idColumn = [string]$Dataset.idColumn
    if (-not $headers.ContainsKey($idColumn)) {
        throw "Colonne ID '$idColumn' absente de $($Dataset.sheet) / $($Dataset.table)."
    }

    $idIndex = $headers[$idColumn]
    $rowsById = @{}
    for ($row = 1; $row -le $table.ListRows.Count; $row++) {
        $identifier = ([string]$table.DataBodyRange.Cells.Item($row, $idIndex).Value2).Trim()
        if (-not [string]::IsNullOrWhiteSpace($identifier)) {
            $rowsById[$identifier] = $row
        }
    }

    $added = 0
    $updatedCells = 0
    foreach ($record in @($payload.records)) {
        $identifier = ([string](Get-PropertyValue -Record $record -Name $idColumn)).Trim()
        if ([string]::IsNullOrWhiteSpace($identifier)) {
            throw "Un enregistrement de $($Dataset.file) ne possède pas d'ID."
        }

        $isNew = -not $rowsById.ContainsKey($identifier)
        if ($isNew) {
            $newListRow = $table.ListRows.Add()
            $rowIndex = [int]$newListRow.Index
            $rowsById[$identifier] = $rowIndex
            $table.DataBodyRange.Cells.Item($rowIndex, $idIndex).Value2 = $identifier
            Set-NewRowDefaults -Table $table -RowIndex $rowIndex -Dataset $Dataset -HeaderMap $headers
            $added++
        }
        else {
            $rowIndex = [int]$rowsById[$identifier]
        }

        foreach ($property in $record.PSObject.Properties) {
            $columnName = [string]$property.Name
            if ($columnName.StartsWith("_")) {
                continue
            }
            if (@($Dataset.preserveColumns) -contains $columnName) {
                continue
            }
            if (@($Dataset.formulaColumns) -contains $columnName) {
                continue
            }
            if (-not $headers.ContainsKey($columnName)) {
                continue
            }
            $cell = $table.DataBodyRange.Cells.Item($rowIndex, $headers[$columnName])
            $newValue = ConvertTo-ExcelCellValue -Value $property.Value
            if (-not (Test-EqualCellValue -Left $cell.Value2 -Right $newValue)) {
                $cell.Value2 = $newValue
                $updatedCells++
            }
        }
    }

    return [PSCustomObject]@{
        Added = $added
        UpdatedCells = $updatedCells
        RemoteCount = @($payload.records).Count
        UpdatedAt = [string]$payload.updatedAt
    }
}

function Backup-Workbook {
    param($Workbook, [string]$OriginalPath)
    New-Item -ItemType Directory -Path $BackupFolder -Force | Out-Null
    $Workbook.Save()
    $baseName = [IO.Path]::GetFileNameWithoutExtension($OriginalPath)
    $extension = [IO.Path]::GetExtension($OriginalPath)
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $backupPath = Join-Path $BackupFolder "$baseName-$stamp$extension"
    Copy-Item -LiteralPath $OriginalPath -Destination $backupPath -Force

    Get-ChildItem -LiteralPath $BackupFolder -Filter "$baseName-*$extension" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -Skip 12 |
        Remove-Item -Force
    return $backupPath
}

function Backup-FileBeforeRepair {
    param([string]$OriginalPath)

    New-Item -ItemType Directory -Path $BackupFolder -Force | Out-Null
    $baseName = [IO.Path]::GetFileNameWithoutExtension($OriginalPath)
    $extension = [IO.Path]::GetExtension($OriginalPath)
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $backupPath = Join-Path $BackupFolder "$baseName-$stamp-avant-reparation$extension"
    Copy-Item -LiteralPath $OriginalPath -Destination $backupPath -Force
    return $backupPath
}

if (-not (Test-Path -LiteralPath $Dossier -PathType Container)) {
    throw "Dossier introuvable : $Dossier"
}

Write-Log "Début de l'actualisation depuis $Dossier."
$config = Get-RemoteJson -Url $ConfigUrl
if ([int]$config.schemaVersion -ne 1) {
    throw "Version de configuration non prise en charge."
}

$excel = $null
$createdExcel = $false
$processed = 0
$totalAdded = 0
$totalUpdated = 0
$errors = [System.Collections.Generic.List[string]]::new()

try {
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        Write-Log "Instance Excel existante détectée."
    }
    catch {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $createdExcel = $true
        Write-Log "Instance Excel locale démarrée."
    }

    $files = Get-ChildItem -LiteralPath $Dossier -Filter "*.xlsx" -File
    foreach ($file in $files) {
        $workbook = $null
        $openedHere = $false
        $stage = "ouverture"
        try {
            # Les fichiers telecharges depuis un navigateur peuvent conserver
            # la marque Windows "provenant d'Internet". Excel les ouvre alors
            # en affichage protege et l'automatisation COM echoue sur Open.
            # Les .xlsx ne contiennent pas de macros VBA ; on peut donc retirer
            # cette marque pour les classeurs du dossier choisi par l'utilisateur.
            Unblock-File -LiteralPath $file.FullName -ErrorAction SilentlyContinue

            $workbook = Get-OpenWorkbook -Excel $excel -FullPath $file.FullName
            if ($null -eq $workbook) {
                # Ne pas actualiser les liaisons externes pendant l'ouverture et
                # demander explicitement une ouverture en lecture/ecriture.
                try {
                    $workbook = $excel.Workbooks.Open($file.FullName, 0, $false)
                    $openedHere = $true
                }
                catch {
                    # Certains classeurs produits par une bibliotheque Open XML
                    # sont lisibles mais Excel demande une reparation silencieuse.
                    # Conserver l'original avant d'autoriser xlRepairFile (1).
                    $repairBackup = Backup-FileBeforeRepair -OriginalPath $file.FullName
                    Write-Log "$($file.Name) : ouverture standard refusee ; tentative de reparation. Original conserve dans $repairBackup" "AVERTISSEMENT"

                    $missing = [Type]::Missing
                    $workbook = $excel.Workbooks.Open(
                        $file.FullName,
                        0,
                        $false,
                        $missing,
                        $missing,
                        $missing,
                        $true,
                        $missing,
                        $missing,
                        $missing,
                        $false,
                        $missing,
                        $false,
                        $true,
                        1
                    )
                    $openedHere = $true
                    Write-Log "$($file.Name) : classeur ouvert en mode reparation Excel." "AVERTISSEMENT"
                }
            }

            $stage = "identification"
            $updateSheet = $null
            $key = ""
            try {
                $updateSheet = $workbook.Worksheets.Item("00 - Actualisation")
                $key = ([string]$updateSheet.Range("N60").Value2).Trim()
            }
            catch {
                foreach ($candidateSheet in @($workbook.Worksheets)) {
                    $candidateKey = ([string]$candidateSheet.Range("AZ1").Value2).Trim()
                    if ($null -ne $config.workbooks.PSObject.Properties[$candidateKey]) {
                        $updateSheet = $candidateSheet
                        $key = $candidateKey
                        break
                    }
                }
            }
            if ($null -eq $updateSheet -or [string]::IsNullOrWhiteSpace($key)) {
                if ($openedHere) {
                    $workbook.Close($false)
                }
                continue
            }
            $workbookConfig = $config.workbooks.PSObject.Properties[$key].Value
            if ($null -eq $workbookConfig) {
                throw "Clé de classeur inconnue : '$key'."
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$workbookConfig.updateSheet)) {
                $updateSheet = $workbook.Worksheets.Item([string]$workbookConfig.updateSheet)
            }

            if ($workbook.ReadOnly) {
                throw "Le classeur est en lecture seule. Ferme Excel, attends la fin de la synchronisation OneDrive, puis relance."
            }

            $stage = "sauvegarde de securite"
            $backupPath = Backup-Workbook -Workbook $workbook -OriginalPath $file.FullName
            Write-Log "Sauvegarde créée : $backupPath"

            $stage = "mise a jour des donnees"
            $bookAdded = 0
            $bookUpdated = 0
            $latestRemote = ""
            foreach ($dataset in @($workbookConfig.datasets)) {
                $result = Update-Dataset -Workbook $workbook -Dataset $dataset -RawBaseUrl ([string]$config.rawBase)
                $bookAdded += $result.Added
                $bookUpdated += $result.UpdatedCells
                $latestRemote = $result.UpdatedAt
                Write-Log (
                    "{0} / {1}: {2} ajout(s), {3} cellule(s) mise(s) à jour." -f
                    $dataset.sheet, $dataset.table, $result.Added, $result.UpdatedCells
                )
            }

            # Le recalcul complet peut echouer a cause d'un autre classeur deja
            # ouvert dans Excel. Cela ne doit pas annuler les donnees mises a jour.
            try {
                $excel.CalculateFull()
            }
            catch {
                Write-Log "$($file.Name) : recalcul complet ignore ($($_.Exception.Message))." "AVERTISSEMENT"
            }

            $stage = "ecriture du statut"
            $status = "Dernière actualisation : {0}`n{1} nouvelle(s) ligne(s) • {2} cellule(s) publique(s) mise(s) à jour`nSource : {3}" -f (
                Get-Date -Format "dd/MM/yyyy HH:mm"
            ), $bookAdded, $bookUpdated, $latestRemote
            $statusCell = if ([string]::IsNullOrWhiteSpace([string]$workbookConfig.statusCell)) { "F6" } else { [string]$workbookConfig.statusCell }
            $lastUpdateCell = if ([string]::IsNullOrWhiteSpace([string]$workbookConfig.lastUpdateCell)) { "N62" } else { [string]$workbookConfig.lastUpdateCell }
            $resultCell = if ([string]::IsNullOrWhiteSpace([string]$workbookConfig.resultCell)) { "N63" } else { [string]$workbookConfig.resultCell }
            Set-WorksheetCellValue -Worksheet $updateSheet -Address $statusCell -Value $status
            Set-WorksheetCellValue -Worksheet $updateSheet -Address $lastUpdateCell -Value (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            Set-WorksheetCellValue -Worksheet $updateSheet -Address $resultCell -Value "$bookAdded|$bookUpdated"

            $stage = "enregistrement final"
            $workbook.Save()

            if ($openedHere) {
                $workbook.Close($true)
            }
            $processed++
            $totalAdded += $bookAdded
            $totalUpdated += $bookUpdated
            Write-Log "$($file.Name) actualisé avec succès."
        }
        catch {
            $message = "$($file.Name) [$stage] : $($_.Exception.Message)"
            $errors.Add($message)
            Write-Log $message "ERREUR"
            if ($null -ne $workbook -and $openedHere) {
                try { $workbook.Close($false) } catch {}
            }
        }
    }
}
finally {
    if ($createdExcel -and $null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    if ($null -ne $excel) {
        try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {}
    }
}

if ($processed -eq 0 -and $errors.Count -eq 0) {
    $message = "Aucun classeur connecté n'a été trouvé dans :`n$Dossier"
    Write-Log $message "AVERTISSEMENT"
    Show-Result -Message $message -Title "Actualisation Excel" -Icon "Warning"
    exit 2
}

if ($errors.Count -gt 0) {
    $message = "Actualisation terminée avec erreur(s).`n`n" + ($errors -join "`n") +
        "`n`nConsulte Actualisation.log. Les sauvegardes n'ont pas été supprimées."
    Show-Result -Message $message -Title "Actualisation Excel" -Icon "Warning"
    exit 1
}

$message = "Actualisation réussie pour $processed classeur(s).`n`n" +
    "$totalAdded nouvelle(s) ligne(s)`n$totalUpdated cellule(s) publique(s) mise(s) à jour`n`n" +
    "Tes données personnelles ont été conservées."
Write-Log $message
Show-Result -Message $message -Title "Actualisation Excel terminée"
