# AUDIT FINAL CONSOLIDÉ — Planning_2026_RUNTIME - 20 2.xlsm

**Date** : 2026-02-20
**Fichier** : `Planning_2026_RUNTIME - 20 2.xlsm` (2.81 MB)
**Localisation** : `C:\Users\hercl\OneDrive\ACTIF\10_Horaires\2026\`
**Méthode** : Analyse factuelle directe sur code VBA exporté (160+ fichiers) via 10 agents parallèles
**Périmètre** : Architecture, VBA, Data Model, Formules, Formatting, Faisabilité Module Congés

---

## 0. EXECUTIVE SUMMARY (10 lignes)

Le fichier de planning hospitalier belge (30 agents, 12 mois) est **fonctionnel mais en dette technique critique**. Sur 160+ fichiers VBA, on identifie **13 bugs bloquants** dont des événements de feuille qui ne se déclenchent jamais (10/12 mois), une clé API exposée en clair, des fêtes françaises au lieu de belges, et un formulaire de sélection d'année bloqué à 2003-2020. L'architecture souffre de **5 implémentations concurrentes** du lecteur de configuration, d'une **duplication quadruple** du calcul de Pâques/jours fériés, et de **500+ références hardcodées** (colonnes, lignes, noms d'employés, chemins absolus). Le `Module_Calculer_Totaux` ne calcule PAS les heures (38h/semaine) mais uniquement la couverture de présence quotidienne (0/1 par période). Le `ManageLeavesForm` existe mais est une **coquille vide** (zéro code). L'intégration d'un module congés nécessite impérativement une **Phase 0 de stabilisation** avant toute implémentation. Score global : **4.5/10** (VBA) — **8/10** de dette technique pour le hardcoding.

---

## 1. AUDIT MACRO — Structure Générale

### 1.1 Inventaire des Feuilles (33 worksheets)

| Catégorie | Feuilles | Rôle |
|-----------|----------|------|
| Planning mensuel | Janv, Fev1, Mars, Avril, Mai, Juin, Juillet, Aout, Sept, Oct, Nov, Decembre, Decembre1 | Grilles planning par mois |
| Configuration | Feuil_Config (tblCFG), Config_Codes, Config_Exceptions | Paramètres centraux |
| Données | Synthese, SuiviRH, HeuresPrester | Calculs agrégés |
| Outils | Sheet212 (Roulements), ImprimerPlan | Roulement cyclique, impression |
| Divers | UserForm sheets, temp sheets | Support UI |

### 1.2 Métriques Structurelles

| Métrique | Valeur | Seuil critique | Statut |
|----------|--------|----------------|--------|
| Taille fichier | 2.81 MB | > 5 MB | OK |
| Feuilles | 33 | > 20 | ATTENTION |
| Règles de mise en forme conditionnelle | 12 282 | > 1000 | CRITIQUE |
| Noms définis | 316 | > 100 | ATTENTION |
| Zones fusionnées | 40 | > 20 | ATTENTION |
| Validations de données | 50 | > 30 | OK |
| Fichiers VBA | 160+ | > 50 | CRITIQUE |

### 1.3 Architecture Planning

- **Lignes personnel** : 6 à 28/30 (selon le mois)
- **Colonnes jours** : 3 à 33 (col C à AG)
- **Codes horaires** : Matin, PM, Soir, Nuit + codes spéciaux (C15, C19, C20, C20E)
- **Seuils horaires** : 13h, 16.5h, 17.5h, 19.5h, 7.25h (hardcodés dans le code)
- **Roulement** : Système cyclique via Sheet212 avec copy/paste entre feuilles

---

## 2. AUDIT MICRO — Analyse VBA Module par Module

### 2.1 Modules Core (Moteur Planning)

#### Module_Planning_Core.bas (668 LOC) — Score : 7/10
**Rôle** : Bibliothèque de fonctions partagées (config, parsing codes, présence, fériés belges, personnel).

| Force | Faiblesse |
|-------|-----------|
| `Option Explicit` présent | Pas de gestion d'erreur globale |
| Zero `Select`/`Activate` | `BuildFeriesBE` correct pour 2026 mais fériés redondants avec 3 autres modules |
| Fonctions bien nommées | `ChargerConfig` lit "Feuil_Config" hardcodé |
| `CalculerPaques` (Meeus) vérifié correct | `ParseCode` vulnérable aux codes non-standards |

**Fonctions clés** : `ChargerConfig`, `ParseCode`, `CalcPeriodes`, `BuildFeriesBE`, `CalculerPaques`, `GetPersonnelRange`

#### Module_Calculer_Totaux.bas (1179 LOC) — Score : 4/10
**Rôle** : Moteur principal de calcul des totaux quotidiens.

**BUG CRITIQUE** : Ce module ne calcule **PAS** les heures travaillées (38h/semaine). Il calcule uniquement la **couverture de présence quotidienne** : combien de personnes sont présentes par période (Matin/PM/Soir/Nuit) = 0 ou 1 par période.

| Problème | Sévérité | Détail |
|----------|----------|--------|
| Duplication intégrale de Planning_Core | HAUTE | 14+ fonctions dupliquées avec suffixe `Fast` |
| Liste d'absences hardcodée | HAUTE | Lignes 284-288 : codes absence en dur |
| Pas de restauration ScreenUpdating | MOYENNE | Si erreur, Excel reste gelé |
| `Left$(code, 1) = "R"` | HAUTE | Exclut TOUS les codes commençant par R |

#### ModulePlanning.bas (810 LOC) — Score : 6/10
**Rôle** : Version "Ultimate Production" avec distance de Levenshtein, cache de codes, exclusions dynamiques.

| Force | Faiblesse |
|-------|-----------|
| Levenshtein pour matching approximatif | **CONFLIT** : `UpdateDailyTotals` défini aussi dans Module_Planning.bas |
| Cache de codes performant | Dépendances croisées avec Module_Planning_Core |
| Exclusions dynamiques | Complexité excessive pour le besoin |

#### Module_Planning.bas (201 LOC) — Score : 3/10
**Rôle** : Version ancienne, partiellement obsolète.
- **Noms d'employés hardcodés** : Bourgeois_Aurore, Diallo_Mamadou, Dela Vega_Edelyn
- **Conflit** : `UpdateDailyTotals` en doublon avec ModulePlanning.bas

#### FillSchedule.bas (95 LOC) — Score : 2/10
**Rôle** : Remplissage automatique du planning.
- **PAS de `Option Explicit`** — seul module sans
- ~90 codes zéro hardcodés
- Logique de parsing obsolète
- **Recommandation** : SUPPRIMER, remplacer par Module_Planning_Core

### 2.2 Modules Remplacement

#### Module_Remplacement_Auto.bas (636 LOC) — Score : 5/10
- Algorithme d'auto-remplacement fonctionnel
- Staffing cible dans des tableaux `Static` hardcodés
- Utilise Module_Planning_Core (bonne pratique)

#### Module_Remplacements_Premium.bas (747 LOC) — Score : 6/10
- Génère des classeurs de demande de remplacement
- Dépend de Module_Config et CReplacementInfo
- Intégration OneDrive pour export PDF

#### ModuleRemplacementFraction.bas — Score : 2/10
- **BUG COMPILATION** : `TraiterUneFeuilleDeMois` défini en DOUBLE — **ne compile pas**
- **BUG MÉTIER** : Fêtes françaises (8 mai, 14 juillet) au lieu de belges
- **Recommandation** : CORRIGER immédiatement ou SUPPRIMER

### 2.3 Modules RH & Heures

#### Module_SuiviRH.bas (553 LOC) — Score : 3/10
- **25 noms d'employés avec quotas de congés hardcodés** (lignes 39-63)
- Scanne les 12 feuilles mensuelles pour agrégation
- Aucune source de données externe — tout en VBA
- **Problème fondamental** : impossible de gérer les arrivées/départs sans modifier le code

#### Module_MAJ_HeuresAPrester.bas (187 LOC) — Score : 4/10
- `HEURES_JOUR_TEMPS_PLEIN = 7.6` hardcodé
- **NE déduit PAS les jours fériés** du calcul
- Pas de gestion du temps partiel (fractions)

#### Module_SyntheseMensuelle.bas (148 LOC) — Score : 5/10
- Calculs de synthèse mensuels corrects
- Dépend de Module_Planning_Core

### 2.4 Configuration (5 implémentations concurrentes !)

| Module | LOC | Source | Début lecture | Cache |
|--------|-----|--------|---------------|-------|
| Module_ConfigEngine.bas | 115 | "Feuil_Config" tblCFG | Ligne 2 | Dictionary |
| Module_Config.bas | 186 | Facade sur ConfigEngine | Ligne 2 | Via ConfigEngine |
| Module_Config_Helpers.bas | 109 | "Feuil_Config" | **Ligne 1** | Non |
| Module_Utilities.bas | 269 | **"Configuration_GenerateNewWorkbo"** | Variable | Non |
| MODULEMODES_CONFIGDRIVEN.bas | 221 | "Feuil_Config" via Find() | Variable | Non |

**CONFLIT DE NOMS** : `CfgTextOr` et `CfgValueOr` existent dans Module_Config ET Module_Config_Helpers avec des signatures différentes. Le compilateur VBA choisira arbitrairement lequel appeler.

**CONFLIT D'ENUM** : `ViewMode` défini dans ModuleModes ET MODULEMODES_CONFIGDRIVEN.

**Recommandation** : Consolider en UN SEUL moteur config (Module_ConfigEngine + Module_Config comme facade).

### 2.5 Classes

| Classe | LOC | Statut |
|--------|-----|--------|
| clsCodeInfo.cls | 55 | ACTIVE — data class pour codes horaires (13 fractions) |
| CRegleComblement.cls | 20 | ACTIVE — règle de comblement (pure data) |
| CReplacementInfo.cls | 18 | ACTIVE — info remplacement (pure data) |
| clsCalendarDay.cls | 0 | MORTE — coquille vide, SUPPRIMER |
| CNormeJour.cls | 0 | MORTE — coquille vide, SUPPRIMER |
| CImpactCodeSuggestion.cls | 0 | MORTE — coquille vide, SUPPRIMER |

### 2.6 Événements de Feuilles

#### BUG CRITIQUE #1 : `Worksheet0_Change` au lieu de `Worksheet_Change`

**10 feuilles sur 12** ont un événement nommé `Worksheet0_Change` au lieu de `Worksheet_Change`. Ce handler n'est **JAMAIS déclenché** par Excel car le nom est invalide.

| Feuille | Événement | Se déclenche ? |
|---------|-----------|----------------|
| Janv | `Worksheet0_Change` | NON |
| Fev1 | `Worksheet_Change` | OUI |
| Mars | `Worksheet0_Change` | NON |
| Avril | `Worksheet0_Change` | NON |
| Mai | `Worksheet0_Change` | NON |
| Juin | `Worksheet0_Change` | NON |
| Juillet | `Worksheet0_Change` | NON |
| Aout | `Worksheet0_Change` | NON |
| Sept | `Worksheet0_Change` | NON |
| Oct | `Worksheet0_Change` | NON |
| Nov | `Worksheet0_Change` | NON |
| Decembre | `Worksheet0_Change` | NON |
| Decembre1 | `Worksheet_Change` | OUI |

**Impact** : La logique de validation/formatage automatique au changement de cellule ne fonctionne que sur Février et Décembre. Les 10 autres mois n'ont aucune réactivité.

#### BUG CRITIQUE #2 : `Worksheet_Change` dans ThisWorkbook

`ThisWorkbook.cls` contient un `Worksheet_Change` — cet événement n'existe PAS au niveau Workbook. Il devrait être `Workbook_SheetChange`. Ce code est **mort** (jamais exécuté).

#### Sheet212 (Roulements) — Le plus complexe
- Copie cyclique entre feuilles mensuelles
- Logique de roulement par rotation d'équipe
- Pas de bug identifié mais complexité élevée

#### Feuil10 (Config) — Hub réactif
- Réagit aux changements dans tblCFG
- Déclenche des recalculs en cascade
- Fonctionnel

### 2.7 UserForms

#### UserForm1.frm — Panneau de contrôle principal
- **~130 boutons** (CommandButton48 à CommandButton133)
- Nommage **non-sémantique** : impossible de savoir ce que fait `CommandButton87`
- Aucun commentaire, aucune documentation
- **Score UX** : 2/10

#### Menu.frm — DOUBLON
- **Copie quasi-identique** de UserForm1
- **Recommandation** : SUPPRIMER, ne garder qu'un seul panneau

#### SelectAnnee.frm — BUG BLOQUANT
- **Plage d'années hardcodée : 2003-2020**
- **Inutilisable depuis 2021** — l'utilisateur ne peut pas sélectionner 2026
- Fix trivial : remplacer par `Year(Date) - 5` à `Year(Date) + 5`

#### FormulaireEntrees.frm — Formulaire de remplacement
- Fonctionnel, correctement structuré
- Utilise CReplacementInfo

#### UfPrintPlanM.frm — Impression
- **BUG** : Zones d'impression corrompues pour Août et Octobre
- Les plages de colonnes ne correspondent pas aux jours du mois

#### UfRazM.frm — Reset planning
- 13 plages de reset hardcodées
- Fonctionnel mais fragile

#### UserForm3.frm — Sélection de nom
- **Risque de boucle événementielle circulaire** entre les contrôles
- Pas de garde `Application.EnableEvents = False`

#### UserForm4.frm — Notes de remplacement
- **Meilleur form du projet** : bien structuré, nommage correct
- Modèle à suivre pour les futurs formulaires

#### ManageLeavesForm.frm — COQUILLE VIDE
- Le formulaire existe visuellement (designer créé)
- **ZÉRO ligne de code VBA** derrière
- Inutilisable en l'état

### 2.8 Modules Debug / Migration / Utilitaires

#### GeminiModule.bas (69 LOC) — SÉCURITÉ CRITIQUE
```
Const API_KEY = "AIzaSyDJaPMS4uJVYFpTU7Ivv2NW7o4xezcjz4k"
```
**CLÉ API GOOGLE GEMINI EN CLAIR** dans le code source. Si le fichier est partagé ou versionné, la clé est exposée.
- **Action immédiate** : révoquer cette clé, utiliser une variable d'environnement ou un fichier .env

#### Module_ExportGit.bas (78 LOC) — Score : 7/10
- Pipeline d'export VBA vers Git
- **ESSENTIEL** — garder et améliorer
- Manque un import automatique (sens inverse)

#### previsionupgradetest.bas (625 LOC) — Score : 5/10
- **ATTENTION** : Malgré le nom "test", c'est un **module de production critique**
- Gestion de la prévision de planning et des upgrades
- **Renommer** en `Module_Prevision_Upgrade.bas`

#### Module_Migration_Structure.bas (205 LOC) — One-shot
- Migration ponctuelle (décalage lignes 60-62 vers 61/63/65)
- **Recommandation** : Archiver après usage, ne pas garder en production

#### Modules à SUPPRIMER (code mort) :
1. `Module_DEBUG_CONFIG.bas` — debug uniquement
2. `Module_Reset_CFG.bas` — reset one-shot
3. `Module_TestConfig.bas` — tests ponctuels
4. `Module_Diagnostic_Config.bas` — diagnostic one-shot
5. `Module_Cleanup_Formats.bas` — nettoyage one-shot
6. `Module_CLEANUP_MASTER.bas` — nettoyage one-shot
7. `ModuleMigration.bas` — migration terminée
8. `Module_Migration_Structure.bas` — migration terminée

### 2.9 _inbox (Versions en attente)

| Fichier _inbox | Comparaison avec prod | Recommandation |
|----------------|----------------------|----------------|
| `PASS2A_Config_Init.bas` | N'existe pas en prod | **PROMOUVOIR** — bootstrap tblCFG |
| `ManageWorkTimeForm.frm` | Plus récent (DatePicker) | **PROMOUVOIR** — remplace version forms/ |
| `Module_ExportAllVBA.bas` | Outil ayant généré l'export | Garder comme utilitaire |
| Autres fichiers _inbox | Doublons de prod | Archiver ou supprimer |

---

## 3. AUDIT TRANSVERSAL — Problèmes Systémiques

### 3.1 Hardcoding — Inventaire exhaustif

| Type de hardcoding | Nombre | Exemples | Sévérité |
|--------------------|--------|----------|----------|
| Noms de feuilles | 215 | "Janv", "Fev1", "Mars"... | HAUTE |
| Références col/ligne | 500+ | `Cells(6, 3)`, `Range("C6:AG28")` | HAUTE |
| Codes horaires | 80+ | "M", "PM", "S", "N", "C15"... | MOYENNE |
| Noms d'employés | 25 | Quotas congés dans Module_SuiviRH | CRITIQUE |
| Chemins absolus | 15 | `C:\Users\hercl\...`, `C:\Users\claud\...` | HAUTE |
| Magic numbers | 40+ | 7.6, 38, 13, 16.5, 17.5, 19.5, 7.25 | MOYENNE |
| Noms de feuilles Config | 5 | "Feuil_Config", "Configuration_GenerateNewWorkbo" | HAUTE |

**Score dette technique hardcoding : 8/10** (10 = insoutenable)

### 3.2 Duplication de Code

| Fonction dupliquée | Occurrences | Modules concernés |
|--------------------|-------------|-------------------|
| Calcul de Pâques (Meeus) | 4 | Planning_Core, Calculer_Totaux, RemplacementFraction, previsionupgradetest |
| Jours fériés belges | 4 | Mêmes modules |
| Lecture config | 5 | ConfigEngine, Config, Config_Helpers, Utilities, MODES_CONFIGDRIVEN |
| ParseCode | 3 | Planning_Core, Calculer_Totaux (Fast), ModulePlanning |
| UpdateDailyTotals | 2 | Module_Planning, ModulePlanning |
| Liste codes absence | 3 | Calculer_Totaux, SuiviRH, Planning_Core |

### 3.3 Incohérences de Nommage

| Incohérence | Détail | Impact |
|-------------|--------|--------|
| "Juillet" vs "Juil" | 4+ modules utilisent des noms différents | Recherche feuille échoue silencieusement |
| Préfixes Module_ vs Module sans | Module_Planning vs ModulePlanning | Confusion à la maintenance |
| PascalCase vs snake_case | Mélange dans les noms de fonctions | Lisibilité réduite |
| Noms de boutons | CommandButton48 à CommandButton133 | Impossible à maintenir |

### 3.4 Gestion d'Erreur

- **Aucun module** n'a de gestionnaire d'erreur global avec restauration de `ScreenUpdating`/`EnableEvents`
- Si une erreur survient pendant un calcul, Excel peut rester en état `ScreenUpdating = False` (écran gelé)
- Pas de pattern `On Error GoTo Cleanup` / `Application.ScreenUpdating = True`

### 3.5 Sécurité

| Risque | Sévérité | Détail |
|--------|----------|--------|
| Clé API Gemini en clair | CRITIQUE | `GeminiModule.bas` ligne 3 |
| Chemins `C:\Users\claud\` | BASSE | Référence à un autre utilisateur (développeur?) |
| Pas de protection VBA | BASSE | Modules accessibles à tous les utilisateurs |

---

## 4. MATRICE DES RISQUES

### Risques par Sévérité × Impact × Probabilité

| # | Risque | Sévérité | Impact | Probabilité | Score | Action |
|---|--------|----------|--------|-------------|-------|--------|
| 1 | Événements Worksheet0_Change jamais déclenchés | CRITIQUE | FORT | CERTAINE (100%) | 10 | Renommer en Worksheet_Change sur 10 feuilles |
| 2 | Clé API Gemini exposée | CRITIQUE | FORT | ÉLEVÉE | 9 | Révoquer immédiatement |
| 3 | ModuleRemplacementFraction ne compile pas | CRITIQUE | FORT | CERTAINE | 9 | Supprimer doublon TraiterUneFeuilleDeMois |
| 4 | Fêtes françaises au lieu de belges | HAUTE | MOYEN | CERTAINE | 8 | Remplacer par BuildFeriesBE |
| 5 | SelectAnnee bloqué à 2003-2020 | HAUTE | FORT | CERTAINE | 8 | Dynamiser la plage |
| 6 | 5 config readers concurrents | HAUTE | MOYEN | ÉLEVÉE | 7 | Consolider en 1 |
| 7 | 25 employés hardcodés dans SuiviRH | HAUTE | FORT | ÉLEVÉE | 7 | Externaliser vers feuille Config |
| 8 | Duplication quadruple Pâques/fériés | MOYENNE | MOYEN | MOYENNE | 6 | Centraliser dans Planning_Core |
| 9 | `Left$(code, 1) = "R"` exclut tous codes R | HAUTE | MOYEN | MOYENNE | 6 | Lister les codes R explicitement |
| 10 | Zones impression Août/Oct corrompues | MOYENNE | FAIBLE | ÉLEVÉE | 5 | Corriger les plages dans UfPrintPlanM |
| 11 | 12 282 règles de mise en forme conditionnelle | MOYENNE | MOYEN | BASSE | 5 | Audit et nettoyage des doublons |
| 12 | Conflit CfgTextOr/CfgValueOr | HAUTE | MOYEN | BASSE | 5 | Renommer dans Config_Helpers |
| 13 | 500+ références col/ligne hardcodées | HAUTE | FORT | BASSE | 5 | Externaliser dans Const/Config |

---

## 5. FAISABILITÉ MODULE CONGÉS

### 5.1 État des Lieux — Ce qui EXISTE

| Composant | Statut | Exploitable ? |
|-----------|--------|---------------|
| ManageLeavesForm.frm | Coquille vide (0 LOC) | NON — tout à coder |
| Module_SuiviRH.bas | 25 quotas hardcodés | PARTIELLEMENT — structure à externaliser |
| Module_MAJ_HeuresAPrester.bas | 7.6h/jour hardcodé | PARTIELLEMENT — ne gère pas fériés/temps partiel |
| PASS2A_Config_Init.bas (_inbox) | Bootstrap tblCFG | OUI — à promouvoir |
| Codes congés existants | CA, EL, ANC, C SOC, DP, CTR, RCT, MAL, MAT, MUT, CRP | OUI — déjà dans le planning |
| Config_Codes | Table des codes | OUI — base pour les types de congés |

### 5.2 Ce qui MANQUE (bloquant)

| Besoin | Statut | Criticité |
|--------|--------|-----------|
| Calcul heures travaillées (38h/semaine) | INEXISTANT | BLOQUANT |
| Suivi solde congés (CA, EL, ANC...) | INEXISTANT | BLOQUANT |
| Gestion temps partiel (fractions horaires) | INEXISTANT | BLOQUANT |
| Déduction automatique jours fériés BE | PARTIEL (fériés calculés mais pas déduits) | BLOQUANT |
| Base de données personnel | HARDCODÉ en VBA | BLOQUANT |
| Historique congés pris/restants | INEXISTANT | BLOQUANT |
| Validation workflow (demande → approbation) | INEXISTANT | SOUHAITABLE |
| Calcul heures supplémentaires / récup | INEXISTANT | SOUHAITABLE |
| Reporting congés (soldes, tendances) | INEXISTANT | SOUHAITABLE |

### 5.3 Verdict de Faisabilité

**FAISABILITÉ : NON en l'état actuel — OUI après Phase 0 de stabilisation**

Le planning ne peut PAS accueillir un module congés sans correction préalable des fondations :
1. Les événements de feuille ne fonctionnent pas (10/12 mois)
2. Il n'y a aucun calcul d'heures (seulement de la couverture de présence)
3. Les données personnel sont hardcodées dans le VBA
4. 5 systèmes de configuration concurrents créent de l'imprévisibilité

### 5.4 Roadmap d'Intégration Congés

#### Phase 0 — Stabilisation (pré-requis, ~2-3 semaines)
1. **Fix Worksheet0_Change** → Worksheet_Change sur 10 feuilles
2. **Fix ModuleRemplacementFraction** → supprimer doublon, corriger fériés
3. **Fix SelectAnnee** → plage dynamique
4. **Révoquer clé API Gemini** → variable d'environnement
5. **Consolider config readers** → 1 seul moteur (ConfigEngine + facade)
6. **Externaliser employés** → feuille Config_Personnel (nom, prénom, contrat, %, quotas)
7. **Supprimer code mort** → 8 modules debug/migration + 3 classes vides
8. **Centraliser Pâques/fériés** → uniquement dans Planning_Core

#### Phase 1 — Fondations Congés (~2-3 semaines)
1. **Créer feuille Config_Personnel** → nom, type contrat, % temps, date entrée, quotas
2. **Créer feuille Soldes_Conges** → par employé, par type, acquis/pris/solde
3. **Promouvoir PASS2A_Config_Init** → infrastructure tblCFG
4. **Implémenter Module_HeuresTravaillees** → calcul réel 38h/semaine × % temps × jours ouvrables − fériés
5. **Coder ManageLeavesForm** → CRUD congés avec calendrier

#### Phase 2 — Moteur de Calcul (~2-3 semaines)
1. **Module_Conges_Engine** → logique métier congés belges
   - CA : 20 jours/an (temps plein), prorata temps partiel
   - EL : Petit chômage (jours légaux selon événement)
   - ANC : Jours ancienneté (selon CCT)
   - MAL : Certificat médical, compteur jours garantis
   - RCT : Récupération heures supplémentaires
   - CTR : Crédit-temps (réduction temps de travail)
2. **Intégration planning** → quand un code congé est posé sur le planning, déduire automatiquement du solde
3. **Validation** → vérifier solde suffisant, chevauchements, staffing minimum

#### Phase 3 — Reporting & UX (~1-2 semaines)
1. **Dashboard congés** → vue synthétique par employé et par équipe
2. **Alertes** → solde bas, fin de période CA, certificat médical manquant
3. **Export** → génération de rapports PDF/Excel pour RH

---

## 6. TOP 10 PROBLÈMES & TOP 5 FORCES

### TOP 10 Problèmes (par ordre de criticité)

| # | Problème | Module(s) | Fix estimé |
|---|----------|-----------|------------|
| 1 | `Worksheet0_Change` — événements muets sur 10/12 mois | 10 sheet modules | 30 min |
| 2 | Clé API Gemini en clair | GeminiModule.bas | 10 min |
| 3 | `TraiterUneFeuilleDeMois` en double — ne compile pas | ModuleRemplacementFraction | 15 min |
| 4 | 5 config readers concurrents avec conflits de noms | 5 modules | 2-4h |
| 5 | 25 employés + quotas hardcodés en VBA | Module_SuiviRH | 4-8h |
| 6 | Fêtes françaises au lieu de belges | ModuleRemplacementFraction | 15 min |
| 7 | SelectAnnee bloqué 2003-2020 | SelectAnnee.frm | 10 min |
| 8 | Duplication quadruple calcul Pâques/fériés | 4 modules | 1-2h |
| 9 | `Left$(code, 1) = "R"` exclut tous codes R | Module_Calculer_Totaux | 30 min |
| 10 | 130 boutons sans nom sémantique | UserForm1.frm | 4-8h |

### TOP 5 Forces

| # | Force | Détail |
|---|-------|--------|
| 1 | Module_Planning_Core bien structuré | `Option Explicit`, zero Select/Activate, fonctions propres |
| 2 | Pipeline Git d'export VBA fonctionnel | Module_ExportGit.bas — versioning du code |
| 3 | Algorithme Meeus correct | Calcul de Pâques vérifié pour 2026 |
| 4 | Infrastructure tblCFG (Config_Engine) | Base solide pour la configuration centralisée |
| 5 | Système de roulement cyclique fonctionnel | Sheet212 — mécanique complexe mais qui marche |

---

## 7. GRAPHE DE DÉPENDANCES & COUPLAGE

### 7.1 Chaînes d'Appels Principales (5 plus longues)

#### Chaîne 1 : Workbook_Open → Vue complète (5 niveaux)
```
ThisWorkbook.Workbook_Open
  → ModuleCFG_Audit.EnsureConfigKeys [STUB VIDE — contient uniquement '  TODO']
  → Module_ViewApply.VIEW_Apply_ByScope
     → Module_Config.CfgTextOr
        → Module_ConfigEngine.CFG_Str
           → Module_ConfigEngine.CFG_Load (lazy init Dictionary)
     → VIEW_ApplyToSheet_WithMode
        → ApplyZoom, ApplyHideMenuCols, ApplyAutoHideNames, ApplyHideBlocks
```

#### Chaîne 2 : RUN_Calendar_And_View (6 niveaux)
```
Module_ViewApply.RUN_Calendar_And_View
  → GenerateurCalendrier.GenererDatesEtJoursPourTousLesMois
     → Module_Config.CfgValueOr / CfgTextOr [~20 appels config]
     → [Calcul Pâques interne + BuildFeriesBE interne]
     → [Écriture 12 feuilles × 31 colonnes]
  → VIEW_Apply_ByScope → [toute la chaîne 1]
```

#### Chaîne 3 : UpdateDailyTotals_V2 (complexité algorithmique max)
```
ModulePlanning.UpdateDailyTotals_V2 [810 LOC, "Ultimate Production"]
  → GetCachedCodeInfo → Module_CodeProcessor.GetCodeInfo
     → [Lit table "tbl_Codes" de Config_Codes]
  → [Levenshtein pour détection typos]
  → [O(employés × jours × codes) — module le plus lourd]
```

#### Chaîne 4 : Modification Config → Vue debounced (6 niveaux)
```
Feuil_Config.Worksheet_Change
  → Module_CFG_Events.CFG_OnChange_RequestViewApply
     → Application.OnTime [debounce 1s]
        → CFG_ApplyView_IfPending
           → Module_ConfigEngine.CFG_Reset [vide le cache]
           → Module_ViewApply.VIEW_Apply_ByScope → [chaîne 1]
```

#### Chaîne 5 : Remplacement Auto (4 niveaux)
```
Module_Remplacement_Auto.AnalyseEtRemplacementPlanningUltraOptimise
  → Module_Planning_Core.ChargerConfig / BuildFeriesBE / CalculerPaques
  → TraiterUneFeuilleDeMois × 12 mois
     → IsInArray [DUPLIQUÉ de Module_Utilities]
```

### 7.2 Score de Couplage : 4/10

**Facteurs positifs (faible couplage)** :
- ~25 modules sur 40+ sont **totalement autonomes** (standalone)
- Classes = DTOs simples sans logique croisée
- Variables globales presque inexistantes (3 dans ModuleGlobals, aucune visiblement utilisée)

**Facteurs négatifs (couplage implicite)** :
- 4+ systèmes lisant `Feuil_Config` en parallèle (risque d'incohérence cache)
- `Module_ViewApply` = hub critique (appelé par 3 sources)
- Couplage fort par données partagées via les feuilles Excel (non visible dans le code)
- Duplication massive = couplage temporel (un changement doit être répliqué partout)

### 7.3 Bugs Supplémentaires Découverts

| Bug | Fichier | Détail |
|-----|---------|--------|
| `CopiedRange` non déclarée | `ModuleCopyPaste.bas` L11/L24 | Variable Range utilisée sans déclaration — ne compile que sans `Option Explicit` |
| `dayLabels` non déclarée | `UserFormManagement.bas` L204 | Collection implicite |
| `ModuleCFG_Audit.EnsureConfigKeys` = stub vide | `ModuleCFG_Audit.bas` | Appelé à chaque `Workbook_Open` mais contient uniquement `' TODO` |
| `ModuleGlobals.bas` — variables mortes | `ModuleGlobals.bas` | `plage()`, `col`, `Li` jamais utilisées. `col` masque des variables locales dans plusieurs modules |
| `ModuleDeclarations.bas` vide | `ModuleDeclarations.bas` | Contient uniquement `Option Explicit`, devrait déclarer `CopiedRange` |

### 7.4 Code Mort — Modules Entièrement Isolés (~20 modules)

Les modules suivants n'appellent aucun autre module ET ne sont appelés par aucun autre module (standalone complets, probablement liés à des boutons Excel non traçables en VBA) :

`FillSchedule`, `ColorShortages`, `ModuleMajPersonnel`, `UserFormManagement`, `Module_SuiviRH`, `ModuleCheckAFCMonthly`, `ModuleVerificationCTR`, `Module_SyntheseMensuelle`, `Module_MAJ_HeuresAPrester`, `Module_Interface`, `SaisieAnnuelle`, `GeminiModule`, `ModuleCopyPaste`, `ModuleRemplacementFraction`, `ModuleColor`, `ModuleModes`, `GeneseEtcolorieRlt`, `ModPersonnelHelper`, `Module_AutoAddCode`

**Note** : Ces modules ne sont pas forcément inutiles — ils peuvent être liés à des boutons ou des raccourcis dans le classeur, mais ils sont invisibles dans le graphe d'appels VBA.

---

## 8. ANNEXES

### A. Inventaire des Codes Horaires

| Code | Description | Seuils |
|------|-------------|--------|
| M | Matin | début < 13h |
| PM | Après-midi | 13h ≤ début < 16.5h |
| S | Soir | 16.5h ≤ début < 19.5h |
| N | Nuit | début ≥ 19.5h OU fin ≤ 7.25h |
| C15 | Coupe 15h | Spécial |
| C19 | Coupe 19h | Spécial |
| C20 | Coupe 20h | Spécial |
| C20E | Coupe 20h Étendu | Spécial |

### B. Codes Congés Existants

CA, EL, ANC, C SOC, DP, CTR, RCT, MAL, MAT, MUT, CRP, DEC, ACC, FO, FER, REC

### C. Structure des Feuilles Mensuelles

```
Ligne 1-5    : En-têtes (titre, mois, jours de semaine)
Ligne 6-28   : Personnel (1 ligne = 1 agent)
Ligne 29-30  : Totaux / résumé
Colonne A    : Noms
Colonne B    : Info complémentaire
Colonne C-AG : Jours 1-31
```

### D. Fichiers VBA Recommandés pour Suppression

1. `Module_DEBUG_CONFIG.bas`
2. `Module_Reset_CFG.bas`
3. `Module_TestConfig.bas`
4. `Module_Diagnostic_Config.bas`
5. `Module_Cleanup_Formats.bas`
6. `Module_CLEANUP_MASTER.bas`
7. `ModuleMigration.bas`
8. `Module_Migration_Structure.bas`
9. `clsCalendarDay.cls`
10. `CNormeJour.cls`
11. `CImpactCodeSuggestion.cls`
12. `Menu.frm` (doublon de UserForm1)

### E. Dépendances Inter-Modules (Graphe simplifié)

```
Module_ConfigEngine ← Module_Config (facade)
                    ← MODULEMODES_CONFIGDRIVEN
                    ← Module_Config_Helpers (CONFLIT)
                    ← Module_Utilities (autre feuille!)

Module_Planning_Core ← Module_Calculer_Totaux (dupliqué avec Fast)
                     ← Module_Remplacement_Auto
                     ← ModulePlanning (conflit UpdateDailyTotals)
                     ← Module_SyntheseMensuelle
                     ← previsionupgradetest

Module_SuiviRH       → Feuilles mensuelles (scan 12 feuilles)
                     → Module_Planning_Core (codes absence)

UserForm1 / Menu     → Tous les modules (130 boutons)
ManageLeavesForm     → RIEN (coquille vide)
```

---

## 9. SCORES FINAUX

| Dimension | Score | Commentaire |
|-----------|-------|-------------|
| Architecture VBA | 4.5/10 | 5 config readers, duplication massive, conflits de noms |
| Qualité du code | 5/10 | Planning_Core bon, le reste hétérogène |
| Robustesse | 3/10 | Pas de gestion erreur, événements muets, hardcoding |
| Sécurité | 2/10 | Clé API en clair |
| UX/Ergonomie | 3/10 | 130 boutons non-nommés, formulaire année cassé |
| Maintenabilité | 3/10 | 500+ refs hardcodées, noms d'employés dans le code |
| Données | 4/10 | Config existante mais fragmentée |
| Prêt pour module congés | 2/10 | Fondations manquantes |
| **SCORE GLOBAL** | **4/10** | **Fonctionnel mais fragile, dette technique élevée** |

---

*Rapport généré le 2026-02-20 par analyse factuelle de 160+ fichiers VBA via 10 agents parallèles spécialisés (76 .bas + 75 .cls analysés).*
