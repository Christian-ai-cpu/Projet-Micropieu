Attribute VB_Name = "RemplissageDesignation"
Option Explicit

' Listes déroulantes sous « Désignation de l'ouvrage », à partir de B5 :
'   B5 à B7   : Toiture / Toiture terrasse, puis Poutre et Mur
'   B8 à B12  : Étage 5, puis Plancher, Poutre, Poteau et Mur
'   ...
'   B28 à B32 : Étage 1, puis Plancher, Poutre, Poteau et Mur
'   B33 à B36 : RDC, puis Poutre, Poteau et Mur (pas de plancher, toujours affiché)
'   B37 à B40 : Soubassement, puis Poutre, Poteau et Mur (toujours affiché)
' Pour un mur : le type se choisit en B, l'épaisseur en F (liste adaptée au type),
' et le poids surfacique D est calculé : D = poids volumique apparent × épaisseur F.
' Pour une poutre ou un poteau : le poids volumique C est rempli selon le type.
' Les valeurs sont des ordres de grandeur à vérifier (normes, fabricants) :
' elles se modifient dans PoidsVolumique et PoidsVolumiqueMur.
' Seuls les étages jusqu'au niveau choisi en C1 sont affichés (voir AfficherEtages).

Private Const PLANCHERS As String = "Plancher dalle pleine,Plancher à poutrelles,Plancher bois traditionnel"
Private Const POUTRES As String = "Poutre BA,Poutre bois,Poutre métallique"
Private Const POTEAUX As String = "Poteau BA,Poteau bois,Poteau métallique"
Private Const MURS As String = "Agglos béton creux,Agglos béton plein,Béton banché," & _
    "Briques creuses,Briques Monomur,Briques pleines,Briques taillées"

Public Sub RemplirDesignationOuvrage()
    Dim listes As Variant, i As Long, e As Long, r As Long

    r = 5
    For e = 6 To -1 Step -1            ' 6 = toiture, étages 5 à 1, 0 = RDC, -1 = soubassement
        If e = 6 Then
            listes = Array("Toiture,Toiture terrasse", POUTRES, MURS)
        ElseIf e = 0 Then
            listes = Array("RDC", POUTRES, POTEAUX, MURS)
        ElseIf e = -1 Then
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
                .Font.Bold = (i = 0)          ' Toiture, Étage n, RDC et Soubassement en gras
            End With
            MajLigne ActiveSheet, r              ' poids volumique / surfacique
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

' Poids volumique (kN/m³) d'une poutre ou d'un poteau selon son type ; 0 sinon.
Public Function PoidsVolumique(typeElement As String) As Double
    Select Case typeElement
        Case "Poutre BA", "Poteau BA":                 PoidsVolumique = 25
        Case "Poutre bois", "Poteau bois":             PoidsVolumique = 6
        Case "Poutre métallique", "Poteau métallique": PoidsVolumique = 78.5
    End Select
End Function

' Poids volumique apparent (kN/m³) d'un mur, vides des blocs creux compris.
' Poids surfacique du mur = cette valeur × épaisseur (m).
Public Function PoidsVolumiqueMur(typeMur As String) As Double
    Select Case typeMur
        Case "Agglos béton creux": PoidsVolumiqueMur = 9      ' 15 cm -> 1,35 kN/m²
        Case "Agglos béton plein": PoidsVolumiqueMur = 20     ' 20 cm -> 4,00 kN/m²
        Case "Béton banché":       PoidsVolumiqueMur = 25     ' 20 cm -> 5,00 kN/m²
        Case "Briques creuses":    PoidsVolumiqueMur = 9      ' 20 cm -> 1,80 kN/m²
        Case "Briques Monomur":    PoidsVolumiqueMur = 8      ' 30 cm -> 2,40 kN/m²
        Case "Briques pleines":    PoidsVolumiqueMur = 18     ' 21,5 cm -> 3,87 kN/m²
        Case "Briques taillées":   PoidsVolumiqueMur = 22     ' pierre de taille, à préciser
    End Select
End Function

' Remplit les poids de la ligne selon l'élément choisi en B :
'   poutre / poteau -> poids volumique en C ; mur -> épaisseur F et poids surfacique D.
Public Sub MajLigne(ws As Worksheet, ligne As Long)
    Dim typeElement As String

    typeElement = CStr(ws.Range("B" & ligne).Value)
    If EpaisseursMur(typeElement) <> "" Then
        MajEpaisseurMur ws, ligne
    ElseIf PoidsVolumique(typeElement) > 0 Then
        ws.Range("C" & ligne).Value = PoidsVolumique(typeElement)
        ws.Range("D" & ligne).ClearContents
    End If
End Sub

' Met en F la liste des épaisseurs du mur choisi en B, et la première en mètres.
' Le poids surfacique D suit l'épaisseur F par formule ; C est vidé pour ne pas
' compter le poids du mur deux fois.
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
    ws.Range("C" & ligne).ClearContents
    ws.Range("D" & ligne).Formula = "=" & _
        Replace(CStr(PoidsVolumiqueMur(CStr(ws.Range("B" & ligne).Value))), ",", ".") & _
        "*F" & ligne
End Sub

' Affiche la toiture et les étages 1 à niveau (C1), masque les autres.
' RDC -> toiture seule (lignes 5 à 7) ; 1 -> toiture + étage 1 ; ... ; 5 -> tout.
' Appelée aussi automatiquement quand C1 change (voir CodeFeuille.txt).
Public Sub AfficherEtages(ws As Worksheet)
    Dim niveau As Long, e As Long

    niveau = Val(ws.Range("C1").Value)     ' "RDC" donne 0
    For e = 1 To 5
        ' Bloc de l'étage e : 5 lignes à partir de la ligne 8 + 5 × (5 - e)
        ws.Rows(8 + 5 * (5 - e)).Resize(5).Hidden = (e > niveau)
    Next e
End Sub
