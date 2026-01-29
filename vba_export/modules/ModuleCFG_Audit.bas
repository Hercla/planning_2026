Attribute VB_Name = "ModuleCFG_Audit"
' ExportedAt: 2026-01-28 00:00:00 | Workbook: Planning_2026.xlsm
Option Explicit

' Ensure required config keys exist (stub for now).
Public Sub EnsureConfigKeys(Optional ByVal silent As Boolean = True)
    ' TODO: wire to real audit/auto-create logic.
End Sub

' Hook on close to cancel pending actions (stub for now).
Public Sub CFG_OnWorkbookClose_CancelPending()
    ' No-op
End Sub
