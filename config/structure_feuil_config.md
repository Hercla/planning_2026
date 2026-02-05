# Structure proposée pour Feuil_Config

## Feuille Excel à créer : `Feuil_Config`

### Tableau 1 : Règles de Présence (A1:C4)

| Indicateur | Règle | Valeur_Attendue |
|------------|-------|-----------------|
| P_0645 | Commence à 6h45 pile (rapport de nuit) | 1 |
| P_7H8H | Présent entre 7h-8h | 3 |
| P_8H1630 | Couvre 8h à 16h30 (binôme C19) | 1 |

---

### Tableau 2 : Codes Coupés (A7:E11)

| Code | Plage1_Debut | Plage1_Fin | Plage2_Debut | Plage2_Fin |
|------|--------------|------------|--------------|------------|
| C 15 | 8:00 | 12:15 | 16:30 | 20:15 |
| C 20 | 8:00 | 12:00 | 16:00 | 20:00 |
| C 20 E | 8:00 | 11:30 | 15:30 | 20:00 |
| C 19 | 7:00 | 11:30 | 15:30 | 19:00 |

---

### Tableau 3 : Lignes de Destination (A14:C25)

| Ligne | Type_Donnee | Description |
|-------|-------------|-------------|
| 60 | Matin | Total présences matin |
| 61 | AM | Total présences après-midi |
| 62 | Soir | Total présences soir |
| 63 | Nuit | Total présences nuit |
| 64 | P_0645 | Présences à 6h45 |
| 65 | P_7H8H | Présences 7h-8h |
| 66 | P_8H1630 | Présences 8h-16h30 |
| 67 | C15 | Codes C15 |
| 68 | C20 | Codes C20 |
| 69 | C20E | Codes C20E |
| 70 | C19 | Codes C19 |

---

## Avantages

✅ **Centralisation** : Toutes les règles métier au même endroit  
✅ **Maintenance** : Modifier les horaires sans toucher au code VBA  
✅ **Documentation** : Visible directement dans Excel  
✅ **Évolutivité** : Ajouter de nouveaux codes facilement

---

## Prochaine étape

Voulez-vous que je modifie `Module_CalculTotaux_Unifie.bas` pour qu'il **lise ces configurations** depuis `Feuil_Config` au lieu d'avoir les valeurs en dur dans le code ?
