Attribute VB_Name = "DuplicationMurs"
Option Explicit

' Module « Duplication des murs » : le nombre de murs se saisit en G1
' (« Entrez le nombre de murs », créé par la macro du tableau).
' Le tableau de descente de charges (colonnes B à K, de la ligne 2 jusqu'au
' récapitulatif G / Q) est recopié vers la droite pour atteindre ce nombre de
' tableaux, avec une colonne vide entre deux tableaux (pas de 11 colonnes) :
'   mur 1 : B à K (tableau d'origine), mur 2 : M à V, mur 3 : X à AG, ...
' Le titre « Mur k » est écrit en ligne 2 au-dessus de chaque tableau.
' Le pourcentage des ouvertures (H1:I1) est aussi recopié au-dessus de chaque copie
' (S1:T1, AD1:AE1, ...) : chaque mur a son propre pourcentage, utilisé dans la
' charge G de ses murs. Les pourcentages déjà saisis sur les copies sont conservés.
' Niveau d'étage (C1) et catégorie (E1) restent communs à tous les murs.
' Appelé automatiquement quand G1 change (voir CodeFeuille.txt), ou à la main.
' À lancer après RemplirChargePermanente et RemplirChargeExploitation.

Public Const PAS_MUR As Long = 11          ' 10 colonnes (B à K) + 1 colonne vide

Public Sub DupliquerMurs(Optional ws As Worksheet)
    Dim nbMurs As Long, k As Long, derniere As Long, c As Long
    Dim source As Range, colB As String, ouvertures() As Variant

    If ws Is Nothing Then Set ws = ActiveSheet
    nbMurs = Val(ws.Range("G1").Value)
    If nbMurs < 1 Then nbMurs = 1

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Fin

    ' Pourcentages d'ouvertures déjà saisis sur les copies (conservés)
    ReDim ouvertures(1 To nbMurs)
    For k = 1 To nbMurs - 1
        ouvertures(k) = ws.Cells(1, 9 + k * PAS_MUR).Value
    Next k

    ' Anciennes copies effacées (tout ce qui est à droite du tableau d'origine, au-dessus
    ' de la matrice des micropieux, recréée ensuite)
    ws.Range(ws.Cells(1, 2 + PAS_MUR - 1), ws.Cells(LIGNE_MATRICE - 1, ws.Columns.Count)).Clear

    ' Tableau d'origine : de la ligne 2 jusqu'au récapitulatif G / Q
    derniere = ws.Cells(LIGNE_MATRICE - 1, "B").End(xlUp).Row
    ws.Range("B2").Value = "Mur 1"
    ws.Range("B2").Font.Bold = True
    Set source = ws.Range("B2:K" & derniere)

    For k = 1 To nbMurs - 1
        source.Copy Destination:=ws.Cells(2, 2 + k * PAS_MUR)
        ws.Cells(2, 2 + k * PAS_MUR).Value = "Mur " & k + 1
        ' Pourcentage des ouvertures propre à ce mur (repris s'il existait déjà)
        ws.Range("H1:I1").Copy Destination:=ws.Cells(1, 8 + k * PAS_MUR)
        If Not IsEmpty(ouvertures(k)) Then ws.Cells(1, 9 + k * PAS_MUR).Value = ouvertures(k)

        ' La toiture de chaque copie dépend de son propre titre (B5 -> M5, X5, ...)
        colB = Split(ws.Cells(1, 2 + k * PAS_MUR).Address(True, True), "$")(1)
        ' (Replace agit directement sur les formules ; pas d'argument LookIn ici)
        source.Offset(0, k * PAS_MUR).Replace What:="$B$5", Replacement:="$" & colB & "$5", _
            LookAt:=xlPart, MatchCase:=False

        ' Mêmes largeurs de colonnes que le tableau d'origine
        For c = 2 To 11
            ws.Columns(c + k * PAS_MUR).ColumnWidth = ws.Columns(c).ColumnWidth
        Next c
    Next k

    ' Matrice des micropieux : une colonne par mur
    ConstruireMatrice ws

Fin:
    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
