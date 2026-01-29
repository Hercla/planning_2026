# Audit technique (VBA export)

Date: 2026-01-29
Scope: vba_export (modules, classes, forms)

## Findings (prioritized)

1) Multiple ThisWorkbook module exports present
- Evidence: ba_export/classes/ThisWorkbook1.cls, ThisWorkbook2.cls, ThisWorkbook3.cls, ThisWorkbook4.cls.
- Risk: Only one ThisWorkbook should exist in the live project. Extra copies in export usually mean stale exports or name drift, and make imports ambiguous.
- Action: In Excel, verify which ThisWorkbook module is actually active, then keep only the active one in source control.

2) Hardcoded paths will break on other machines
- Evidence:
  - ba_export/modules/ImportSelectedVBA.bas (targetFolder = C:\Users\hercl\planning_2026\)
  - ba_export/modules/Module_ImportAll.bas (folderPath = C:\Users\hercl\planning_2026)
  - ba_export/modules/UpdateFeuilConfig_And_Import.bas (CSV_PATH and MODULE_PATH in C:\Users\hercl\planning-vba-automation\...)
  - ba_export/modules/Module_Macros_Genere_copy_model.bas (OneDrive path for another user)
- Risk: Import/export and config operations silently fail for other users or even on this machine if folders change.
- Action: Centralize paths in one config function (ProjectRoot) and replace literals.

3) Two different export/import systems coexist
- Evidence:
  - ba_export/modules/Module_ExportAllVBA.bas exports to _EXPORT_VBA.
  - Custom export/import macros are expected to write to ba_export/.
- Risk: Users can export to the wrong folder or import stale code.
- Action: Pick one system; delete or disable the other entrypoints.

4) Empty procedures in userforms
- Evidence:
  - ba_export/forms/Menu.frm: CommandButton21_Click, CommandButton22_Click, ComboBox2_Change
  - ba_export/forms/UserForm1.frm: ComboBox4_Change, ComboBox1_Change, Frame2_Click
- Risk: Dead event handlers and noise during maintenance.
- Action: Remove if truly unused, or implement logic if still needed.

5) Module_ConfigAudit is empty
- Evidence: ba_export/modules/Module_ConfigAudit.bas contains only attribute line.
- Risk: No functional impact, but adds noise and can confuse imports.
- Action: Remove or implement.

## Option Explicit coverage

Missing Option Explicit in 42 files. Highest priority to fix in modules/forms (not sheet modules).

Missing list:
- classes: Acceuil11.cls, Acceuil2.cls, CImpactCodeSuggestion.cls, CNormeJour.cls, CRegleComblement.cls, CReplacementInfo.cls, Decembre11.cls, Feuil131.cls, Feuil141.cls, Feuil16.cls, Feuil21.cls, Feuil241.cls, Feuil31.cls, Feuil41.cls, Feuil51.cls, Feuil61.cls, Feuil71.cls, Feuil81.cls, Feuil91.cls, Fev11.cls, Juin1.cls, Salarie1.cls, Sheet2121.cls, Sheet41.cls
- forms: ManageLeavesForm.frm, ManagePersonalInfoForm.frm, ManagePositionsForm.frm, ManageWorkTimeForm.frm, Menu.frm, UfPrintPlanM.frm, UserForm2.frm, UserForm3.frm
- modules: ColorShortages.bas, Debug_Columns.bas, Debug_Probe.bas, FillSchedule.bas, GénèreEtcolorieRlt.bas, ModImpression.bas, ModNotes.bas, Module_ConfigAudit.bas, ModuleColor.bas, ModuleModes_ConfigDriven.bas, Moduletokencode.bas, NettoyerPlannings.bas

## Suspected duplicates or legacy variants (manual verify)

- Module_Planning.bas vs ModulePlanning.bas (both large planning logic)
- Module_Remplacements.bas vs Module_Remplacements_Premium.bas (one may be a superset)
- GenerateurCalendrier.bas vs GenerateurCalendrier_V2.bas
- ModuleModes.bas vs ModuleModes_ConfigDriven.bas
- Module_Macros_Genere_copy_model.bas (appears as a copy/model)
- previsionupgradetest.bas (name suggests test or upgrade script)
- Debug_Columns.bas, Debug_Probe.bas, Module_Debug_Config.bas, Module_Debug_Couleurs.bas (debug utilities)

Recommendation: do not delete yet. Tag as LEGACY/DEBUG, confirm entrypoints, then quarantine.

## Next steps (Pass 0 -> Pass 1/2)

1) Manual compile in Excel (Debug > Compile VBAProject) and fix any missing references.
2) Choose the single export/import system and remove the other entrypoints.
3) Add Option Explicit to all non-sheet modules and forms.
4) Decide which "duplicate" modules are active. Keep only one.
5) Export code again and commit as "Compile clean".
