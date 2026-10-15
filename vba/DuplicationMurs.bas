Attribute VB_Name = "DuplicationMurs"
Option Explicit

' Module « Duplication des murs » : le nombre de murs se saisit en A1.
' Le tableau de descente de charges (colonnes B à M, de la ligne 2 jusqu'au
' récapitulatif G / Q) est recopié vers la droite pour atteindre ce nombre de
' tableaux, avec une colonne vide entre deux tableaux (pas de 13 colonnes) :
'   mur 1 : B à M (tableau d'origine), mur 2 : O à Z, mur 3 : AB à AM, ...
' Le titre « Mur k » est écrit en ligne 2 au-dessus de chaque tableau.
' Niveau d'étage (C1) et catégorie (E1) restent communs à tous les murs.
' Appelé automatiquement quand A1 change (voir CodeFeuille.txt), ou à la main.
' À lancer après RemplirChargePermanente et RemplirChargeExploitation.

Public Const PAS_MUR As Long = 13          ' 12 colonnes (B à M) + 1 colonne vide

Public Sub DupliquerMurs(Optional ws As Worksheet)
    Dim nbMurs As Long, k As Long, derniere As Long, c As Long
    Dim source As Range, colB As String

    If ws Is Nothing Then Set ws = ActiveSheet
    nbMurs = Val(ws.Range("A1").Value)
    If nbMurs < 1 Then nbMurs = 1

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Fin

    ' Anciennes copies effacées (tout ce qui est à droite du tableau d'origine)
    ws.Range(ws.Cells(1, 2 + PAS_MUR - 1), ws.Cells(ws.Rows.Count, ws.Columns.Count)).Clear

    ' Tableau d'origine : de la ligne 2 jusqu'au récapitulatif G / Q
    derniere = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    ws.Range("B2").Value = "Mur 1"
    ws.Range("B2").Font.Bold = True
    Set source = ws.Range("B2:M" & derniere)

    For k = 1 To nbMurs - 1
        source.Copy Destination:=ws.Cells(2, 2 + k * PAS_MUR)
        ws.Cells(2, 2 + k * PAS_MUR).Value = "Mur " & k + 1

        ' La toiture de chaque copie dépend de son propre titre (B5 -> O5, AB5, ...)
        colB = Split(ws.Cells(1, 2 + k * PAS_MUR).Address(True, True), "$")(1)
        ' (Replace agit directement sur les formules ; pas d'argument LookIn ici)
        source.Offset(0, k * PAS_MUR).Replace What:="$B$5", Replacement:="$" & colB & "$5", _
            LookAt:=xlPart, MatchCase:=False

        ' Mêmes largeurs de colonnes que le tableau d'origine
        For c = 2 To 13
            ws.Columns(c + k * PAS_MUR).ColumnWidth = ws.Columns(c).ColumnWidth
        Next c
    Next k

Fin:
    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
