# inject_all_modules.ps1
# Script master d'injection de tous les modules VBA dans Planning_2026_RUNTIME.xlsm
# Usage: powershell -File inject_all_modules.ps1 [-DryRun] [-SkipSave]

param(
    [switch]$DryRun,
    [switch]$SkipSave
)

$ErrorActionPreference = "Stop"

# --- Configuration ---
$xlsmPath = "C:\Users\hercl\OneDrive\ACTIF\10_Horaires\2026\Planning_2026_RUNTIME - 20 2.xlsm"
$repoPath = "C:\Users\hercl\planning_2026_repo"

# Modules a injecter (ordre important: dependances d'abord)
$modules = @(
    @{ Name = "Module_Config_Personnel";    File = "Module_Config_Personnel.bas";       Action = "NEW" },
    @{ Name = "Module_SuiviRH";            File = "Module_SuiviRH_v2.bas";            Action = "REPLACE" },
    @{ Name = "Module_SyntheseMensuelle";  File = "Module_SyntheseMensuelle_v2.bas";   Action = "REPLACE" },
    @{ Name = "Module_MAJ_HeuresAPrester"; File = "Module_MAJ_HeuresAPrester_v2.bas";  Action = "REPLACE" },
    @{ Name = "Module_Conges_Engine";      File = "Module_Conges_Engine.bas";          Action = "NEW" },
    @{ Name = "Module_HeuresSup";          File = "Module_HeuresSup.bas";              Action = "NEW" },
    @{ Name = "Module_Alertes";            File = "Module_Alertes.bas";                Action = "NEW" },
    @{ Name = "Module_ManageLeaves_UI";    File = "Module_ManageLeaves_UI.bas";        Action = "NEW" }
)

Write-Host "============================================" -ForegroundColor Cyan
Write-Host " INJECTION MODULES VBA - Planning 2026" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

if ($DryRun) {
    Write-Host "[DRY RUN] Aucune modification ne sera effectuee" -ForegroundColor Yellow
    Write-Host ""
}

# --- Verification des fichiers source ---
Write-Host "--- Verification des fichiers source ---" -ForegroundColor White
$allFilesOk = $true
foreach ($mod in $modules) {
    $fullPath = Join-Path $repoPath $mod.File
    if (Test-Path $fullPath) {
        $lines = (Get-Content $fullPath).Count
        Write-Host "  [OK] $($mod.File) ($lines lignes)" -ForegroundColor Green
    } else {
        Write-Host "  [MISSING] $($mod.File)" -ForegroundColor Red
        $allFilesOk = $false
    }
}

if (-not $allFilesOk) {
    Write-Host ""
    Write-Host "ERREUR: Fichiers manquants. Abandon." -ForegroundColor Red
    exit 1
}

# --- Verification du fichier Excel ---
Write-Host ""
Write-Host "--- Verification du fichier Excel ---" -ForegroundColor White
if (-not (Test-Path $xlsmPath)) {
    Write-Host "  [MISSING] $xlsmPath" -ForegroundColor Red
    Write-Host "  Le fichier Excel n'est pas accessible (OneDrive sync?)." -ForegroundColor Yellow
    exit 1
}
Write-Host "  [OK] Fichier Excel trouve" -ForegroundColor Green

# --- Backup ---
Write-Host ""
Write-Host "--- Backup ---" -ForegroundColor White
$backupDir = Join-Path $repoPath "backups"
if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupPath = Join-Path $backupDir "Planning_BACKUP_$timestamp.xlsm"

if (-not $DryRun) {
    Copy-Item $xlsmPath $backupPath
    Write-Host "  [OK] Backup cree: $backupPath" -ForegroundColor Green
} else {
    Write-Host "  [DRY RUN] Backup serait cree: $backupPath" -ForegroundColor Yellow
}

if ($DryRun) {
    Write-Host ""
    Write-Host "--- DRY RUN termine ---" -ForegroundColor Yellow
    Write-Host "Modules qui seraient injectes:" -ForegroundColor White
    foreach ($mod in $modules) {
        Write-Host "  $($mod.Action): $($mod.Name) <- $($mod.File)" -ForegroundColor Cyan
    }
    Write-Host ""
    Write-Host "Pour executer: powershell -File inject_all_modules.ps1" -ForegroundColor White
    exit 0
}

# --- Ouverture Excel COM ---
Write-Host ""
Write-Host "--- Ouverture Excel ---" -ForegroundColor White

$excel = $null
$wb = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1  # msoAutomationSecurityLow

    Write-Host "  [OK] Excel COM demarre" -ForegroundColor Green

    $wb = $excel.Workbooks.Open($xlsmPath)
    Write-Host "  [OK] Classeur ouvert" -ForegroundColor Green

    $vbProj = $wb.VBProject
    if ($null -eq $vbProj) {
        throw "Impossible d'acceder au VBProject. Verifiez: Fichier > Options > Centre de gestion de la confidentialite > Parametres > Parametres des macros > Cochez 'Acceder au modele objet du projet VBA'"
    }
    Write-Host "  [OK] VBProject accessible" -ForegroundColor Green

    # --- Injection des modules ---
    Write-Host ""
    Write-Host "--- Injection des modules ---" -ForegroundColor White

    $injectedCount = 0
    foreach ($mod in $modules) {
        $modName = $mod.Name
        $modFile = Join-Path $repoPath $mod.File
        $action = $mod.Action

        Write-Host "  [$action] $modName..." -NoNewline

        # Supprimer l'ancien module si REPLACE
        $existing = $null
        try {
            foreach ($comp in $vbProj.VBComponents) {
                if ($comp.Name -eq $modName) {
                    $existing = $comp
                    break
                }
            }
        } catch {}

        if ($null -ne $existing) {
            $vbProj.VBComponents.Remove($existing)
            Write-Host " (ancien supprime)" -NoNewline -ForegroundColor Yellow
        }

        # Importer le nouveau module
        $imported = $vbProj.VBComponents.Import($modFile)
        if ($null -ne $imported) {
            Write-Host " OK ($($imported.Name))" -ForegroundColor Green
            $injectedCount++
        } else {
            Write-Host " ECHEC" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host "--- Resultat: $injectedCount/$($modules.Count) modules injectes ---" -ForegroundColor Cyan

    # --- Sauvegarde ---
    if (-not $SkipSave) {
        Write-Host ""
        Write-Host "--- Sauvegarde ---" -ForegroundColor White
        $wb.Save()
        Write-Host "  [OK] Classeur sauvegarde" -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "  [SKIP] Sauvegarde ignoree (-SkipSave)" -ForegroundColor Yellow
    }

} catch {
    Write-Host ""
    Write-Host "ERREUR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack: $($_.ScriptStackTrace)" -ForegroundColor DarkRed
} finally {
    # --- Cleanup COM ---
    Write-Host ""
    Write-Host "--- Fermeture Excel ---" -ForegroundColor White

    if ($null -ne $wb) {
        try { $wb.Close($false) } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host "  [OK] COM nettoye" -ForegroundColor Green
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " TERMINE" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Prochaines etapes:" -ForegroundColor White
Write-Host "  1. Ouvrir Planning_2026_RUNTIME.xlsm" -ForegroundColor White
Write-Host "  2. Alt+F8 > InitialiserConfigPersonnel > Executer" -ForegroundColor White
Write-Host "  3. Alt+F8 > MigrerQuotasDepuisSuiviRH > Executer" -ForegroundColor White
Write-Host "  4. Alt+F8 > InitialiserFeuillesConges > Executer" -ForegroundColor White
Write-Host "  5. Alt+F8 > RecalculerTousSoldes > Executer" -ForegroundColor White
Write-Host "  6. Alt+F8 > ShowManageLeavesForm > Tester le dashboard" -ForegroundColor White
