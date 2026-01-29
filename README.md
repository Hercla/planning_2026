#  Projet Planning 2026

Ce depot contient le code source et le fichier Excel de gestion du planning 2026.

##  Architecture du projet

*   **workbook/** : Contient le fichier binaire Excel (Planning_2026.xlsm). C'est le fichier utilisable par l'utilisateur.
*   **ba_export/** : Contient tout le code VBA extrait (Modules, Classes, UserForms). C'est ce dossier qui permet le suivi de version (Git).
*   **M_VersionControl.bas** : Le module utilitaire integre au classeur qui gere l'exportation.

##  Comment travailler sur ce projet (Workflow)

Pour assurer la synchronisation entre le fichier Excel et le depot Git, merci de suivre cette procedure :

1.  **Ouvrir** le fichier workbook/Planning_2026.xlsm.
2.  **Effectuer les modifications** (Code VBA, formules, design...).
3.  **AVANT DE SAUVEGARDER/QUITTER** :
    *   Aller sur la feuille 'Admin' (ou la ou est le bouton).
    *   Cliquer sur le bouton **"EXPORTER LE CODE"**.
    *   Attendre le message de confirmation.
4.  **Sauvegarder** le fichier Excel (Ctrl + S).
5.  **Commiter** les changements :
    *   Les fichiers dans ba_export/ montreront les lignes de code modifiees.
    *   Le fichier .xlsm sera mis a jour en tant que binaire.

##  Installation / Restauration

Si vous recuperez ce depot pour la premiere fois :
1.  Le fichier Planning_2026.xlsm est pret a l'emploi dans le dossier workbook.
2.  Si besoin de mettre a jour le code depuis les fichiers sources, utilisez la macro ImportAllVBA (usage avance uniquement).

##  Notes techniques
*   Le fichier .gitattributes est configure pour traiter les .xlsm comme des fichiers binaires (evite les conflits de fusion illisibles).
