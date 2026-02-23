# PLAN DE REMÉDIATION & ARCHITECTURE MODULE CONGÉS
## Planning_2026_RUNTIME — Document Expert

**Date** : 2026-02-20
**Fichier** : `Planning_2026_RUNTIME - 20 2.xlsm`
**Méthode** : 10 agents spécialisés (audit) + 4 agents (remédiation/architecture)
**Objectif** : Stabiliser le planning existant (Phase 0) puis intégrer un module congés complet

---

# PARTIE A — PATCHS CRITIQUES PHASE 0

## Priorité 1 : Clé API Gemini en clair (SÉCURITÉ)

```
FICHIER: vba_export\_inbox\GeminiModule.bas
LIGNE(S): 16-17
AVANT:
    Dim apiKey As String
    apiKey = "AIzaSyDJaPMS4uJVYFpTU7Ivv2NW7o4xezcjz4k"

APRÈS:
    Dim apiKey As String
    On Error Resume Next
    apiKey = Trim(CStr(ThisWorkbook.Sheets("Feuil_Config").Range("B2").Value))
    On Error GoTo 0
    If Len(apiKey) = 0 Then
        CallGemini = "Erreur: Cle API manquante dans Feuil_Config."
        Exit Function
    End If

RISQUE: MOYEN — Révoquer immédiatement la clé dans Google Cloud Console.
```

---

## Priorité 2 : ModuleRemplacementFraction — ne compile pas

### 2a. Sub/Function mismatch (5 corrections)

```
FICHIER: vba_export\_inbox\ModuleRemplacementFraction.bas
LIGNE 131: End Function → End Sub
LIGNE 164: Exit Function → Exit Sub  (dans TraiterUneFeuilleDeMois_ParRegles)
LIGNE 261: End Function → End Sub
LIGNE 471: Exit Function → Exit Sub
LIGNE 507: End Function → End Sub
RISQUE: AUCUN
```

### 2b. Fériés français → belges + syntaxe Return invalide

```
FICHIER: vba_export\_inbox\ModuleRemplacementFraction.bas
LIGNES: 269-291
AVANT:
Private Function ObtenirCodeJourFerie(dateJour As Date) As String
    ...
    If jour = 8 And mois = 5 Then Return "F 8-5"    ' Victoire 1945 [FRANÇAIS]
    If jour = 14 And mois = 7 Then Return "F 14-7"  ' Fête nationale [FRANÇAIS]
    ...

APRÈS:
Private Function ObtenirCodeJourFerie(dateJour As Date) As String
    Dim jour As Integer, mois As Integer
    jour = Day(dateJour): mois = Month(dateJour)

    ' Jours fériés fixes BELGES
    If jour = 1 And mois = 1 Then ObtenirCodeJourFerie = "F 1-1": Exit Function
    If jour = 1 And mois = 5 Then ObtenirCodeJourFerie = "F 1-5": Exit Function
    If jour = 21 And mois = 7 Then ObtenirCodeJourFerie = "F 21-7": Exit Function  ' BELGE
    If jour = 15 And mois = 8 Then ObtenirCodeJourFerie = "F 15-8": Exit Function
    If jour = 1 And mois = 11 Then ObtenirCodeJourFerie = "F 1-11": Exit Function
    If jour = 11 And mois = 11 Then ObtenirCodeJourFerie = "F 11-11": Exit Function
    If jour = 25 And mois = 12 Then ObtenirCodeJourFerie = "F 25-12": Exit Function
    ObtenirCodeJourFerie = ""
End Function

CHANGEMENTS:
  - 8 mai SUPPRIMÉ (Victoire = français, pas belge)
  - 14 juillet SUPPRIMÉ (fête nationale française)
  - 21 juillet AJOUTÉ (fête nationale belge)
  - Syntaxe "Return" remplacée par "Fonction = val: Exit Function" (VBA valide)
RISQUE: FAIBLE
```

---

## Priorité 3 : Worksheet0_Change sur 10 feuilles (événements muets)

```
FICHIERS: _inbox\Mars.cls, Avril.cls, Mai1.cls, Juin.cls, Juillet.cls,
          Aout.cls, Sept.cls, Oct.cls, Nov.cls, Decembre.cls
LIGNE: 11 (dans chaque fichier)
AVANT: Private Sub Worksheet0_Change(ByVal Target As Range)
APRÈS: Private Sub Worksheet_Change(ByVal Target As Range)
RISQUE: AUCUN — "Worksheet0_Change" n'est pas un event VBA valide.
NOTE: Fev1.cls et Decembre1.cls sont déjà corrects. Janv.cls introuvable dans l'export.
```

---

## Priorité 4 : ThisWorkbook — event handler invalide

```
FICHIER: _inbox\ThisWorkbook.cls
LIGNES: 39-48
AVANT:
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("AG41")) Is Nothing Then
        ...
    End If
End Sub

APRÈS:
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Not Intersect(Target, Sh.Range("AG41")) Is Nothing Then
        If Target.Value = "x" Then
            Call UpdateF3
            Sh.Range("Us 1D 1").Copy
            Sheets.Add After:=Sheets(Sheets.Count)
            ActiveSheet.Range("Us 1D 1").PasteSpecial
        End If
    End If
End Sub

RISQUE: FAIBLE — "Us 1D 1" est un named range suspect, à vérifier.
```

---

## Priorité 5 : SelectAnnee — bloqué 2003-2020

```
FICHIER: _inbox\SelectAnnee.frm
LIGNE 34: If annee < 2003 → If annee < Year(Date) - 5
LIGNE 48: If annee > 2020 → If annee > Year(Date) + 5
LIGNE 85-87: For i = 2003 To 2020 → For i = Year(Date) - 5 To Year(Date) + 5
LIGNE 66, 70: ComboBox1 = "2003" → ComboBox1 = CStr(Year(Date))
RISQUE: AUCUN
```

---

## Priorité 6 : Zones impression Août/Octobre corrompues

```
FICHIER: _inbox\UfPrintPlanM.frm
LIGNE 42 (Août):
  AVANT: Call PrintPlanMois(ZoneImprim:="&AY$19:$BD$53")
  APRÈS: Call PrintPlanMois(ZoneImprim:="$AY$19:$BD$53")

LIGNE 46 (Octobre):
  AVANT: Call PrintPlanMois(ZoneImprim:="^BM$19:$BBR$53")
  APRÈS: Call PrintPlanMois(ZoneImprim:="$BM$19:$BR$53")

RISQUE: AUCUN — Pattern confirmé mathématiquement (colonnes espacées de 6).
```

---

## Résumé Phase 0 — Effort estimé

| Patch | Temps | Risque | Fichiers |
|-------|-------|--------|----------|
| Clé API Gemini | 10 min + révocation | MOYEN | 1 |
| ModuleRemplacementFraction Sub/Function | 15 min | AUCUN | 1 |
| ModuleRemplacementFraction fériés | 15 min | FAIBLE | 1 |
| Worksheet0_Change × 10 feuilles | 30 min | AUCUN | 10 |
| ThisWorkbook event | 10 min | FAIBLE | 1 |
| SelectAnnee dynamique | 10 min | AUCUN | 1 |
| Zones impression | 5 min | AUCUN | 1 |
| **TOTAL PHASE 0** | **~2h** | | **15 fichiers** |

---

# PARTIE B — NETTOYAGE CODE MORT

## Fichiers à SUPPRIMER (confirmés par analyse de dépendances)

| Fichier | LOC | Raison | Appelants |
|---------|-----|--------|-----------|
| Module_Debug_Config.bas | 72 | Debug one-shot | Aucun |
| Module_Migration_Structure.bas | 205 | Migration one-shot exécutée | Aucun |
| clsCalendarDay.cls | 11 | Classe vide | Aucun |
| CNormeJour.cls | 10 | Classe vide, remplacée par Type local | Aucun |
| CImpactCodeSuggestion.cls | 10 | Classe vide | Aucun |
| Menu.frm | 227 | Doublon obsolète de UserForm1 | Aucun |
| FillSchedule.bas | 95 | Remplacé par Module_Planning_Core | Aucun |
| ModuleGlobals.bas | 9 | Variables mortes (plage, col, Li) | Aucun |
| ModuleDeclarations.bas | 7 | Module vide | Aucun |
| ModuleModes.bas | 116 | Legacy, remplacé par MODULEMODES_CONFIGDRIVEN | Aucun |

**6 fichiers candidats n'existent déjà plus** : Module_Reset_CFG, Module_TestConfig, Module_Diagnostic_Config, Module_Cleanup_Formats, Module_CLEANUP_MASTER, ModuleMigration.

**Total** : 10 fichiers à supprimer = **762 LOC** de code mort éliminées.

## Fichier à RENOMMER

```
previsionupgradetest.bas → Module_Prevision_Remplacements.bas
```
Raison : module CRITIQUE en production (appelé par UserForm1 bouton 132), nom trompeur "test".

## Duplication Pâques/Fériés — Consolidation

**Version canonique** : `Module_Planning_Core.BuildFeriesBE` (Public, Dictionary, 10 fériés belges corrects)

| Module à modifier | Fonction dupliquée | Type retour | Action |
|--------------------|-------------------|-------------|--------|
| GenerateurCalendrier.bas | BuildFeriesBE (Private) | Collection | Remplacer par Module_Planning_Core.BuildFeriesBE |
| Module_JoursFeries.bas | BuildFeriesBE (Private) | Variant/Array | Remplacer par Module_Planning_Core.BuildFeriesBE |
| Module_Calculer_Totaux.bas | BuildFeriesFast (Private) | Dictionary | Remplacer par Module_Planning_Core.BuildFeriesBE |
| ModuleRemplacementFraction.bas | ObtenirCodeJourFerie | String | Remplacer par Module_Planning_Core.EstDansFeries |

**Attention** : adaptation nécessaire car les types de retour diffèrent (Collection vs Array vs Dictionary).

---

# PARTIE C — CONSOLIDATION CONFIGURATION (8 → 2 modules)

## État actuel : 8 systèmes concurrents

```
Feuille "Feuil_Config" (A=clé, B=valeur, ligne 2+)
         |
    +----+--------+-------------+-----------+
    |              |             |           |
 M1: ConfigEngine  M3: Helpers  M5: MODES  M8: JoursFeries
 (Dict, row 2)    (Dict, row 1!) (Find)    (For loop)
    |              |
 M2: Config       [CONFLIT:
 (facade)         CfgTextOr × 2]
    |
 +--+----+---+
 |       |   |
View  Colors Calendar   M4: Utilities        M6: Planning_Core
Apply        Generateur (AUTRE feuille!)     (Dict injectable)
                                               |
                                          M7: Calculer_Totaux
                                          (copie locale de M6)
```

**Conflits critiques** :
- `CfgTextOr` et `CfgValueOr` existent dans Module_Config ET Module_Config_Helpers
- 3 noms pour la même clé année : `CFG_Year`, `AnneePlanning`, `ANNEE_PLANNING`
- 2 feuilles sources : `Feuil_Config` et `Configuration_GenerateNewWorkbo`
- Ligne 1 vs ligne 2 comme début de lecture

## Architecture cible : 2 modules

```
Module_ConfigEngine.bas (MOTEUR — enrichi)
  + CFG_StrOr(key, default), CFG_LongOr, CFG_BoolOr, CFG_ValueOr
  + CFG_Exists(key), CFG_ToDict()

Module_Config.bas (FACADE — compatibilité)
  CfgText → CFG_StrOr
  CfgLong → CFG_LongOr
  CfgTextOr → CFG_StrOr
  CfgValueOr → CFG_ValueOr
  + LireParametre (alias pour compat PDF)
```

## Plan de migration en 7 phases

| Phase | Action | Risque | Fichiers touchés |
|-------|--------|--------|-----------------|
| 1 | Enrichir ConfigEngine (CFG_*Or, CFG_ToDict) | AUCUN (additif) | 1 |
| 2 | Vérifier consommateurs bien branchés | AUCUN | 0 (ViewApply, Calendar, Remplacements = déjà OK) |
| 3 | Supprimer CfgValueOr Private dans PlanningColors | FAIBLE | 1 |
| 4 | **Supprimer Module_Config_Helpers entièrement** | MOYEN (conflit noms) | 1 suppr + audit appelants |
| 4 | Migrer MODULEMODES_CONFIGDRIVEN Private → Module_Config | FAIBLE | 1 |
| 4 | Migrer ModulePDFGeneration Private → Module_Config/Utilities | FAIBLE | 1 |
| 4 | Migrer Module_JoursFeries Private → Module_Config | FAIBLE | 1 |
| 5 | Migrer Module_Utilities (feuille alternative) | MOYEN (diff clés) | 1 |
| 6 | Consolider Planning_Core/Calculer_Totaux | MOYEN | 2 |
| 7 | Consolider BuildFeriesBE dupliqué | MOYEN (types retour) | 4 |

**Risque principal** : Phase 4 (suppression Module_Config_Helpers) — vérifier tous les appels non-préfixés à CfgTextOr/CfgValueOr avant suppression.

---

# PARTIE D — ARCHITECTURE MODULE CONGÉS

## D.1 Cartographie de l'existant exploitable

| Composant existant | Exploitable ? | Pour quoi ? |
|--------------------|---------------|-------------|
| Codes congés (CA, EL, ANC, MAL...) dans Config_Codes | OUI | Classification déjà dans le planning |
| Module_Planning_Core.BuildFeriesBE | OUI | 10 fériés belges, algorithme Meeus vérifié |
| Module_Planning_Core.ParseCode | OUI | Parsing codes horaires pour calcul heures |
| Feuille Personnel (matricule, nom, fonction, %) | OUI | Base de données agents |
| Colonnes synthèse existantes (Heures prestées, Jours congé...) | OUI | Headers déjà en place |
| ManageLeavesForm.frm (designer) | PARTIEL | Coquille vide, redesigner les contrôles |
| Module_SuiviRH.GetQuotas (25 agents hardcodés) | MIGRATION | Source initiale pour Config_Personnel |

## D.2 Nouvelles feuilles Excel

### Config_Personnel (remplace les quotas hardcodés)

| Col | Header | Type | Description |
|-----|--------|------|-------------|
| A | Matricule | String | Clé unique |
| B | Nom | String | |
| C | Prénom | String | |
| D | Fonction | String | INF / AS / CEFA |
| E | DateEntree | Date | |
| F | DateSortie | Date | Vide si actif |
| G | ContratBase | String | CDI / CDD |
| H | PctTemps | Double | 50, 75, 80, 90, 100 |
| I | RegimeCTR | String | NEANT / CTR_1_5 / CTR_1_2 |
| J | QuotaCA | Double | Défaut: 24 (20 légal + 4 CCT hospitalier) |
| K | QuotaEL | Double | Selon situation familiale |
| L | QuotaANC | Double | Selon ancienneté CCT |
| M | QuotaCSoc | Double | |
| N | QuotaDP | Double | |
| O | QuotaCRP | Double | |
| P-T | Autres quotas + reports | Double | |

### Soldes_Conges (dashboard temps réel)

| Col | Header | Description |
|-----|--------|-------------|
| A-B | Matricule, NomComplet | Identification |
| C-E | CA_Acquis, CA_Pris, CA_Solde | Congé annuel |
| F-H | EL_Acquis, EL_Pris, EL_Solde | Petit chômage |
| I-K | ANC_Acquis, ANC_Pris, ANC_Solde | Ancienneté |
| L-Z | CSOC, DP, CRP, MAL, MAT, RCT | Autres types |
| AA-AE | CTR_Type, Totaux, DernièreMAJ | Agrégats |

### Historique_Conges (journal d'audit)

| Col | Header | Description |
|-----|--------|-------------|
| A | ID | Auto-incrément |
| B | DateHeure | Timestamp |
| C-D | Matricule, NomComplet | Agent |
| E | TypeConge | CA / EL / MAL... |
| F | Action | PRISE / ANNULATION / AJUSTEMENT |
| G-I | DateDebut, DateFin, NbJours | Période |
| J-K | SoldeAvant, SoldeApres | Traçabilité |
| L-O | Source, MoisPlanning, Utilisateur, Commentaire | Contexte |

## D.3 Modules VBA à créer

### Module_Conges_Engine.bas — Moteur central

```vba
' Fonctions publiques principales :
Public Sub RecalculerTousSoldes()
Public Sub RecalculerSoldesMois(ByVal nomMois As String)
Public Function CompterCongesParAgent(nomAgent, typeConge, moisDeb, moisFin) As Double
Public Function GetSoldeAgent(matricule, typeConge) As Double
Public Function GetQuotaAgent(matricule, typeConge) As Double
Public Function ClassifierCodeConge(code As String) As String
Public Function EstCodeAbsence(code As String) As Boolean
Public Function ScannerPlanningMensuel(nomMois) As Object  ' Dictionary
Public Sub EcrireSoldesConges(dictGlobal As Object)
Public Sub EcrireHistorique(matricule, nom, type, action, dateDeb, dateFin, ...)
Public Sub InitialiserFeuillesConges()
Public Sub MigrerQuotasDepuisSuiviRH()
Public Function ValiderPriseConge(matricule, type, dateDeb, dateFin, ByRef msg) As Boolean
Public Function CalculerNbJoursOuvrables(dateDeb, dateFin, annee) As Long
```

**ClassifierCodeConge** — logique de classification :
```vba
Select Case True
    Case c = "CA": → "CA"
    Case c = "EL": → "EL"
    Case c = "ANC": → "ANC"
    Case c = "C SOC": → "CSOC"
    Case Left(c, 3) = "CRP": → "CRP"
    Case Left(c, 7) = "MAL-GAR": → "MAL"
    Case Left(c, 7) = "MAL-MUT", Left(c, 3) = "MUT": → "MUT"
    Case Left(c, 7) = "MAT-EMP", Left(c, 7) = "MAT-MUT": → "MAT"
    Case Left(c, 1) = "F" And InStr(c, "-") > 0: → "FERIE"
    ...
End Select
```

### Module_HeuresTravaillees.bas — Calcul réel des heures

```vba
' Ce qui MANQUE aujourd'hui et que ce module apporte :
Public Function HeuresTheoriquesMois(annee, moisNum, pctTemps) As Double
    ' joursOuvrables × 7.6 × (pctTemps / 100)
    ' DÉDUIT les jours fériés belges (contrairement à HeuresAPresterDyn)
End Function

Public Function HeuresPresteesMois(nomMois, nomAgent, annee) As Double
    ' Parse chaque code horaire dans le planning
    ' Utilise Module_Planning_Core.ParseCode pour extraire les heures
End Function

Public Function DureeEffectiveCode(code As String) As Double
    ' "7:00 15:15" → 8.25h
    ' "6:45 12:00 12:30 15:00" → 7.75h (avec pause)
    ' "19:45 6:45" → 11h (nuit, passage minuit)
End Function

Public Function JoursOuvrablesMois(annee, moisNum) As Long
    ' Lun-Ven hors fériés belges via BuildFeriesBE
End Function

Public Function ProrataCongesEntree(quotaAnnuel, dateEntree, annee) As Double
    ' Règle belge : quota × (moisRestants / 12)
    ' Mois d'entrée compte si entrée ≤ 15
End Function
```

### ManageLeavesForm — Code-behind complet

**Contrôles UI** :
- `cmbAgent` : sélection agent → charge soldes + historique
- `cmbTypeConge` : CA, EL, ANC, CSOC, DP, CRP, MAL, MAT, RCT, CTR
- `txtDateDebut`, `txtDateFin` : saisie dates
- `lblNbJours` : calcul auto jours ouvrables (preview temps réel)
- `lblSoldeApres` : solde restant après déduction (vert si OK, rouge si négatif)
- `btnPoser` : validation → écriture planning + historique + recalcul soldes
- `btnRecalculer` : force recalcul complet depuis les 12 feuilles
- `lstHistorique` : 5 derniers mouvements de l'agent

**Flux principal** :
```
[Sélection agent] → [Affiche soldes]
[Saisie type + dates] → [Preview: nb jours, solde après]
[Clic Poser] → ValiderPriseConge() → EcrireCongesDansPlanning()
            → EcrireHistorique() → RecalculerTousSoldes()
            → Rafraîchir UI
```

## D.4 Intégration avec l'existant (non-destructive)

### Hook Worksheet_Change → mise à jour automatique des soldes

```vba
' À AJOUTER dans ThisWorkbook.cls (une seule fois, couvre tous les mois)
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Not IsMonthSheet(Sh.Name) Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 6 Or Target.Row > 50 Then Exit Sub
    If Target.Column < 3 Or Target.Column > 33 Then Exit Sub

    If Module_Conges_Engine.EstCodeAbsence(CStr(Target.Value)) Then
        Application.EnableEvents = False
        On Error Resume Next
        Module_Conges_Engine.RecalculerSoldesMois Sh.Name
        On Error GoTo 0
        Application.EnableEvents = True
    End If
End Sub
```

### Module_SuiviRH — migration GetQuotas()

```vba
' Remplacer les 25 quotas hardcodés par lecture de Config_Personnel
Private Function GetQuotas() As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Config_Personnel")
    If ws Is Nothing Then
        Set GetQuotas = GetQuotasLegacy()  ' fallback ancien code
        Exit Function
    End If
    ' Lecture depuis Config_Personnel...
End Function
```

### Module_MAJ_HeuresAPrester — correction calcul

```vba
' Remplacer HeuresAPresterDyn (ne déduit pas les fériés)
arrResultats(i, 1) = Module_HeuresTravaillees.HeuresTheoriquesMois( _
    annee, Module_Planning_Core.MoisNumero(mois), pourcentage * 100)
```

### Module_Calculer_Totaux — hook optionnel

```vba
' Après le calcul du cockpit, déclencher la MAJ soldes congés
On Error Resume Next
If CStr(GCS(cfg, "CONGE_AUTO_UPDATE")) = "TRUE" Then
    Module_Conges_Engine.RecalculerSoldesMois nomOnglet
End If
On Error GoTo 0
```

## D.5 Règles métier belges

| Type | Base | Prorata temps partiel | Particularité |
|------|------|-----------------------|---------------|
| CA | 24j/an (20 légal + 4 CCT hospitalier) | Int(24 × %/100 + 0.5) | Report N-1 autorisé (max 5j configurable) |
| EL | Selon événement (1-3j) | Non | Pas de prorata |
| ANC | 1-7j selon ancienneté CCT | Oui | 5-9ans=1j, 10-14=2j, 15-19=3j, 20-24=5j, 25+=7j |
| MAL | Illimité | N/A | 30j salaire garanti employeur, puis mutuelle |
| RCT | Heures sup accumulées | 1 jour = 7.6h × %/100 | Conversion heures → jours |
| CTR | Réduction contractuelle | N/A | Pas de quota, code posé les jours non prestés |
| MAT | 15 semaines (105j) | N/A | 30j employeur (MAT-EMP), puis mutuelle (MAT-MUT) |

## D.6 Schéma d'architecture

```
  ┌─────────────────────────────────────────────────────────────┐
  │                  PLANNING_2026_RUNTIME.xlsm                 │
  ├─────────────────────────────────────────────────────────────┤
  │                                                             │
  │  FEUILLES EXISTANTES          NOUVELLES FEUILLES            │
  │  ──────────────────           ──────────────────            │
  │  Janv..Dec (×12)              Config_Personnel              │
  │  Feuil_Config                 Soldes_Conges                 │
  │  Config_Codes                 Historique_Conges             │
  │  Personnel                                                  │
  │                                                             │
  │  FLUX PRINCIPAL :                                           │
  │                                                             │
  │  [Code congé posé dans onglet mois]                         │
  │           │                                                 │
  │           ▼                                                 │
  │  Workbook_SheetChange (ThisWorkbook)                        │
  │           │                                                 │
  │           ▼                                                 │
  │  Module_Conges_Engine.EstCodeAbsence()                      │
  │           │ Oui                                             │
  │           ▼                                                 │
  │  RecalculerSoldesMois()                                     │
  │    ├── Scanner onglet mois (compter codes par agent)        │
  │    ├── Lire quotas (Config_Personnel)                       │
  │    ├── Calculer soldes = acquis - pris                      │
  │    ├── Écrire dans Soldes_Conges                            │
  │    └── Logger dans Historique_Conges                         │
  │                                                             │
  │  FORMULAIRE (ManageLeavesForm) :                            │
  │                                                             │
  │  [Sélection agent] → [Affiche soldes]                       │
  │  [Type + dates]    → [Preview jours/solde]                  │
  │  [Poser]           → Validation → Écriture planning         │
  │                    → Historique → Recalcul → Refresh UI     │
  │                                                             │
  │  Module_HeuresTravaillees :                                 │
  │    HeuresThéoriquesMois = joursOuvrables × 7.6 × %          │
  │    (DÉDUIT les fériés, contrairement à l'existant)          │
  │                                                             │
  └─────────────────────────────────────────────────────────────┘
```

## D.7 Plan d'implémentation

| Phase | Durée | Contenu | Dépend de |
|-------|-------|---------|-----------|
| **0. Stabilisation** | 2h | 15 patchs critiques (Partie A) | Rien |
| **0b. Nettoyage** | 1h | Supprimer 10 fichiers morts (Partie B) | Phase 0 |
| **1. Fondations congés** | 2j | Créer 3 feuilles, migrer quotas, clés config | Phase 0 |
| **2. Moteur calcul** | 3j | Module_Conges_Engine complet + tests | Phase 1 |
| **3. Heures** | 2j | Module_HeuresTravaillees + remplacement HeuresAPrester | Phase 1 |
| **4. Intégration** | 2j | Hook SheetChange, modifier SuiviRH, hook Calculer_Totaux | Phases 2+3 |
| **5. Formulaire** | 3j | ManageLeavesForm designer + code-behind | Phase 4 |
| **6. Config consolidation** | 3j | 7 phases de migration config (Partie C) | Phase 0 |
| **7. Validation** | 2j | Tests complets, comparaison avec SuiviRH, cas limites | Tout |

**Durée totale estimée** : ~15 jours de développement (en parallèle: Phases 2+3, Phases 5+6)

---

# ANNEXE : Clés de configuration à ajouter dans Feuil_Config

```
CONGE_HEURES_JOUR_TP         7.6
CONGE_HEURES_SEMAINE_TP      38
CONGE_CA_BASE_JOURS          24
CONGE_MAL_SALAIRE_GARANTI    30
CONGE_FERIES_BELGES          10
CONGE_ANNEE                  2026
CONGE_AUTO_UPDATE            TRUE
CONGE_REPORT_CA_AUTORISE     TRUE
CONGE_REPORT_CA_MAX          5
```

---

*Document généré le 2026-02-20 par analyse de 160+ fichiers VBA via 14 agents spécialisés.*
