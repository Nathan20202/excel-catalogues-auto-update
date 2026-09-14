[CmdletBinding()]
param(
    [string]$Dossier = $PSScriptRoot,
    [ValidateRange(0, 60)]
    [int]$DelaiMinutes = 3,
    [switch]$Desinstaller
)

$ErrorActionPreference = "Stop"
$TaskName = "Actualisation automatique des classeurs Excel"

if ($Desinstaller) {
    $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($null -ne $existingTask) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Lancement automatique supprimé." -ForegroundColor Green
    }
    else {
        Write-Host "Aucune tâche '$TaskName' n'était installée." -ForegroundColor Yellow
    }
    exit 0
}

$Dossier = ([string]$Dossier).Trim().Trim('"')
try {
    $Dossier = [IO.Path]::GetFullPath($Dossier)
    $Dossier = $Dossier.TrimEnd(
        [IO.Path]::DirectorySeparatorChar,
        [IO.Path]::AltDirectorySeparatorChar
    )
}
catch {
    throw "Chemin de dossier non valide : $Dossier"
}

$Actualiseur = Join-Path $Dossier "Actualiser_Excel.ps1"
if (-not (Test-Path -LiteralPath $Actualiseur -PathType Leaf)) {
    throw "Actualiser_Excel.ps1 est introuvable dans : $Dossier"
}

$PowerShellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$Arguments = (
    '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass ' +
    '-WindowStyle Hidden -File "{0}" -Dossier "{1}" -SansMessage' -f
    $Actualiseur, $Dossier
)

$ActionParameters = @{
    Execute = $PowerShellExe
    Argument = $Arguments
    WorkingDirectory = $Dossier
}
$Action = New-ScheduledTaskAction @ActionParameters

$CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$Trigger = New-ScheduledTaskTrigger -AtLogOn -User $CurrentUser
if ($DelaiMinutes -gt 0) {
    $Trigger.Delay = "PT$($DelaiMinutes)M"
}

$SettingsParameters = @{
    StartWhenAvailable = $true
    AllowStartIfOnBatteries = $true
    DontStopIfGoingOnBatteries = $true
    MultipleInstances = "IgnoreNew"
    ExecutionTimeLimit = (New-TimeSpan -Hours 1)
}
$Settings = New-ScheduledTaskSettingsSet @SettingsParameters

$PrincipalParameters = @{
    UserId = $CurrentUser
    LogonType = "Interactive"
    RunLevel = "Limited"
}
$Principal = New-ScheduledTaskPrincipal @PrincipalParameters

$RegisterParameters = @{
    TaskName = $TaskName
    Action = $Action
    Trigger = $Trigger
    Settings = $Settings
    Principal = $Principal
    Description = "Actualise les classeurs Excel après l'ouverture de session Windows."
    Force = $true
}
Register-ScheduledTask @RegisterParameters | Out-Null

Write-Host ""
Write-Host "Installation réussie." -ForegroundColor Green
Write-Host "L'actualisation se lancera automatiquement $DelaiMinutes minute(s) après chaque ouverture de session."
Write-Host "Aucune fenêtre de résultat ne s'affichera ; le détail restera dans Actualisation.log."
