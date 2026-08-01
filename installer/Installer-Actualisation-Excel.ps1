[CmdletBinding()]
param([string]$DossierCible)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$RawBase = "https://raw.githubusercontent.com/Nathan20202/excel-catalogues-auto-update/main"

function Select-Folder {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Choisis le dossier contenant les cinq classeurs Excel."
    $dialog.ShowNewFolderButton = $true
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "Installation annulée."
    }
    return $dialog.SelectedPath
}

function Download-File {
    param([string]$Url, [string]$Destination)
    Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing -TimeoutSec 120
    Unblock-File -LiteralPath $Destination -ErrorAction SilentlyContinue
}

if ([string]::IsNullOrWhiteSpace($DossierCible)) {
    $DossierCible = Select-Folder
}
if (-not (Test-Path -LiteralPath $DossierCible -PathType Container)) {
    throw "Le dossier choisi n'existe pas."
}

$enginePath = Join-Path $DossierCible "Actualiser_Excel.ps1"
$launcherPath = Join-Path $DossierCible "Actualiser_Excel.cmd"
Download-File -Url "$RawBase/installer/Actualiser_Excel.ps1" -Destination $enginePath
Download-File -Url "$RawBase/installer/Actualiser_Excel.cmd" -Destination $launcherPath

$localGuide = @"
ACTUALISATION DES CLASSEURS EXCEL

1. Double-clique sur Actualiser_Excel.cmd ou sur le raccourci du Bureau.
2. Le moteur sauvegarde chaque classeur avant toute écriture.
3. Les sauvegardes se trouvent dans _sauvegardes_actualisation.
4. Le journal se trouve dans Actualisation.log.

Le moteur télécharge uniquement des données publiques depuis :
https://github.com/Nathan20202/excel-catalogues-auto-update

Il n'envoie aucune donnée vers Internet.
"@
Set-Content -LiteralPath (Join-Path $DossierCible "Lisez-moi_Actualisation.txt") -Value $localGuide -Encoding UTF8

$shell = New-Object -ComObject WScript.Shell
$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "Actualiser mes classeurs Excel.lnk"
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $launcherPath
$shortcut.WorkingDirectory = $DossierCible
$shortcut.Description = "Sauvegarde puis actualise les cinq classeurs Excel."
$shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,238"
$shortcut.Save()

$scheduled = $false
try {
    $taskName = "Nathan - Actualisation catalogues Excel"
    $action = New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\cmd.exe" -Argument "/c `"$launcherPath`""
    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At "12:15"
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings `
        -Description "Synchronisation locale gratuite des catalogues Excel." -Force | Out-Null
    $scheduled = $true
}
catch {
    $scheduled = $false
}

$scheduleMessage = if ($scheduled) {
    "Une synchronisation locale hebdomadaire a aussi été programmée le dimanche."
}
else {
    "La tâche automatique Windows n'a pas pu être créée, mais le bouton et le raccourci fonctionnent."
}

Add-Type -AssemblyName System.Windows.Forms
$answer = [System.Windows.Forms.MessageBox]::Show(
    "Installation terminée.`n`nDossier : $DossierCible`n$scheduleMessage`n`nLancer une première actualisation maintenant ?",
    "Actualisation Excel",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Information
)
if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
    Start-Process -FilePath $launcherPath -WorkingDirectory $DossierCible
}

