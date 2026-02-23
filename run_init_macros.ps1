# run_init_macros.ps1
# Execute les 5 macros d'initialisation dans l'ordre

$ErrorActionPreference = "Stop"

$xlsmPath = "C:\Users\hercl\OneDrive\ACTIF\10_Horaires\2026\Planning_2026_RUNTIME - 20 2.xlsm"

$macros = @(
    @{ Name = "InitialiserConfigPersonnel";  Desc = "Cree la feuille Config_Personnel" },
    @{ Name = "MigrerQuotasDepuisSuiviRH";  Desc = "Remplit les 25 agents + quotas" },
    @{ Name = "InitialiserFeuillesConges";   Desc = "Cree Soldes_Conges + Historique_Conges" },
    @{ Name = "RecalculerTousSoldes";        Desc = "Scanne 12 mois, calcule les soldes" },
    @{ Name = "ShowManageLeavesForm";        Desc = "Cree le Dashboard Conges" }
)

Write-Host "============================================" -ForegroundColor Cyan
Write-Host " INITIALISATION MODULES - Planning 2026" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$excel = $null
$wb = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1

    Write-Host "[OK] Excel COM demarre" -ForegroundColor Green

    $wb = $excel.Workbooks.Open($xlsmPath)
    Write-Host "[OK] Classeur ouvert" -ForegroundColor Green
    Write-Host ""

    foreach ($macro in $macros) {
        Write-Host "--- $($macro.Name) ---" -ForegroundColor White
        Write-Host "    $($macro.Desc)" -ForegroundColor Gray

        try {
            $excel.Run($macro.Name)
            Write-Host "    [OK] Execute avec succes" -ForegroundColor Green
        } catch {
            $errMsg = $_.Exception.Message
            # Les MsgBox bloquent en COM, on les intercepte
            if ($errMsg -match "dialog" -or $errMsg -match "MsgBox" -or $errMsg -match "1004") {
                Write-Host "    [WARN] Execute mais MsgBox intercepte (normal)" -ForegroundColor Yellow
            } else {
                Write-Host "    [ERREUR] $errMsg" -ForegroundColor Red
            }
        }
        Write-Host ""
    }

    # Sauvegarde
    Write-Host "--- Sauvegarde ---" -ForegroundColor White
    $wb.Save()
    Write-Host "    [OK] Classeur sauvegarde" -ForegroundColor Green

} catch {
    Write-Host "ERREUR GLOBALE: $($_.Exception.Message)" -ForegroundColor Red
} finally {
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
    Write-Host ""
    Write-Host "[OK] COM nettoye" -ForegroundColor Green
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " INITIALISATION TERMINEE" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Tu peux maintenant ouvrir le fichier Excel et verifier :" -ForegroundColor White
Write-Host "  - Feuille 'Config_Personnel' (25 agents + quotas)" -ForegroundColor White
Write-Host "  - Feuille 'Soldes_Conges' (soldes calcules)" -ForegroundColor White
Write-Host "  - Feuille 'Historique_Conges' (journal vide)" -ForegroundColor White
Write-Host "  - Feuille 'Dashboard_Conges' (interface de saisie)" -ForegroundColor White
