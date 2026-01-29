# Planning 2026

## Rituel de travail (recommande)

1. Ouvrir workbook\\Planning_2026.xlsm.
2. Travailler (formules, mises en forme, VBA, etc.).
3. Cliquer sur le bouton **Exporter VBA**.
4. Enregistrer le classeur (Ctrl+S).
5. Dans le terminal :

`powershell
git add .
git commit -m "Description courte des changements"
`

## Arborescence

- workbook/ : classeur Excel principal (Planning_2026.xlsm).
- ba_export/ : export lisible du code VBA.
  - modules/ : modules standard (.bas)
  - classes/ : classes (.cls)
  - orms/ : formulaires (.frm/.frx)

## Notes

- L'option Excel "Acces approuve au modele d'objet du projet VBA" doit etre activee.
