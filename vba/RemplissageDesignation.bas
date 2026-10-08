Attribute VB_Name = "RemplissageDesignation"
Option Explicit

' Listes déroulantes sous « Désignation de l'ouvrage », à partir de B5 :
'   B5 à B8   : Toiture / Toiture terrasse, puis Poutre, Poteau et Mur
'   B9 à B13  : Étage 5, puis Plancher, Poutre, Poteau et Mur
'   ...
'   B29 à B33 : Étage 1, puis Plancher, Poutre, Poteau et Mur
'   B34 à B37 : Soubassement, puis Poutre, Poteau et Mur (toujours affiché)
' Pour un mur : le type se choisit en B, l'épaisseur en F (liste adaptée au type).
' Seuls les étages jusqu'au niveau choisi en C1 sont affichés (voir AfficherEtages).

Private Const PLANCHERS As String = "Plancher dalle pleine,Plancher à poutrelles,Plancher bois traditionnel"
Private Const POUTRES As String = "Poutre BA,Poutre bois,Poutre métallique"
Private Const POTEAUX As String = "Poteau BA,Poteau bois,Poteau métallique"
Private Const MURS As String = "Agglos béton creux,Agglos béton plein,Béton banché," & _
    "Briques creuses,Briques Monomur,Briques pleines,Briques taillées"

Public Sub RemplirDesignationOuvrage()
    Dim listes As Variant, i As Long, e As Long, r As Long

    r = 5
    For e = 6 To 0 Step -1             ' 6 = toiture, étages 5 à 1, 0 = soubassement
        If e = 6 Then
            listes = Array("Toiture,Toiture terrasse", POUTRES, POTEAUX, MURS)
        ElseIf e = 0 Then
            listes = Array("Soubassement", POUTRES, POTEAUX, MURS)
        Else
            listes = Array("Étage " & e, PLANCHERS, POUTRES, POTEAUX, MURS)
        End If

        For i = 0 To UBound(listes)
            With ActiveSheet.Range("B" & r)
                .Validation.Delete
                ' Liste déroulante seulement s'il y a plusieurs choix
                If InStr(listes(i), ",") > 0 Then
                    .Validation.Add xlValidateList, xlValidAlertStop, , listes(i)
                End If
                .Value = Split(listes(i), ",")(0)
                .Font.Bold = (i = 0)          ' Toiture, Étage n et Soubassement en gras
            End With
            If listes(i) = MURS Then MajEpaisseurMur ActiveSheet, r
            r = r + 1
        Next i
    Next e

    AfficherEtages ActiveSheet
End Sub

' Épaisseurs possibles (en cm) selon le type de mur ; "" si ce n'est pas un mur.
Public Function EpaisseursMur(typeMur As String) As String
    Select Case typeMur
        Case "Agglos béton creux", "Agglos béton plein": EpaisseursMur = "15 cm,20 cm"
        Case "Béton banché":    EpaisseursMur = "20 cm,25 cm,30 cm"
        Case "Briques creuses": EpaisseursMur = "15 cm,20 cm,25 cm"
        Case "Briques Monomur": EpaisseursMur = "30 cm,37 cm"
        Case "Briques pleines": EpaisseursMur = "10 cm,21.5 cm,33 cm"
        Case "Briques taillées": EpaisseursMur = "30 cm,40 cm,50 cm,60 cm"
    End Select
End Function

' Met en F la liste des épaisseurs du mur choisi en B, et la première en mètres.
Public Sub MajEpaisseurMur(ws As Worksheet, ligne As Long)
    Dim liste As String

    liste = EpaisseursMur(CStr(ws.Range("B" & ligne).Value))
    If liste = "" Then Exit Sub
    With ws.Range("F" & ligne)
        .Validation.Delete
        .Validation.Add xlValidateList, xlValidAlertStop, , liste
        .Value = Val(liste) / 100           ' "15 cm" -> 0.15 m
        .NumberFormat = "0.000"
    End With
End Sub

' Affiche la toiture et les étages 1 à niveau (C1), masque les autres.
' RDC -> toiture seule (lignes 5 à 8) ; 1 -> toiture + étage 1 ; ... ; 5 -> tout.
' Appelée aussi automatiquement quand C1 change (voir CodeFeuille.txt).
Public Sub AfficherEtages(ws As Worksheet)
    Dim niveau As Long, e As Long

    niveau = Val(ws.Range("C1").Value)     ' "RDC" donne 0
    For e = 1 To 5
        ' Bloc de l'étage e : 5 lignes à partir de la ligne 9 + 5 × (5 - e)
        ws.Rows(9 + 5 * (5 - e)).Resize(5).Hidden = (e > niveau)
    Next e
End Sub
