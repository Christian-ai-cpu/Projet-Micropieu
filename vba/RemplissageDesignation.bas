Attribute VB_Name = "RemplissageDesignation"
Option Explicit

' Listes déroulantes sous « Désignation de l'ouvrage », à partir de B5 :
'   B5 à B7   : Toiture / Toiture terrasse, puis Plancher et Poutre
'               (la ligne Plancher n'est affichée que pour une toiture terrasse)
'   B8 à B12  : Étage 5, puis Plancher, Poutre, Poteau et Mur
'   ...
'   B28 à B32 : Étage 1, puis Plancher, Poutre, Poteau et Mur
'   B33 à B36 : RDC, puis Poutre, Poteau et Mur (pas de plancher, toujours affiché)
'   B37 à B40 : Soubassement, puis Poutre, Poteau et Mur (toujours affiché)
' Pour un mur : le type se choisit en B, l'épaisseur en F (liste adaptée au type),
' et le poids surfacique D est calculé : D = poids volumique × épaisseur F.
' Pour une poutre ou un poteau : le poids volumique C est lu selon le type.
' Les poids volumiques sont dans la feuille « Poids matériaux » (créée si absente) :
' on les modifie directement dans cette feuille, le tableau se met à jour.
' Seuls les étages jusqu'au niveau choisi en C1 sont affichés (voir AfficherEtages).
' Sous-total par étage : sur la ligne de titre (Toiture, Étage n, RDC, Soubassement),
' la colonne H fait la somme des éléments du bloc. Le TOTAL G (SOUS.TOTAL) ignore
' ces sous-totaux, il n'y a donc pas de double compte.

Private Const PLANCHERS As String = "Plancher dalle pleine,Plancher à poutrelles,Plancher bois traditionnel"
Private Const POUTRES As String = "Poutre BA,Poutre bois,Poutre métallique"
Private Const POTEAUX As String = "Poteau BA,Poteau bois,Poteau métallique"
Private Const MURS As String = "Agglos béton creux,Agglos béton plein,Béton banché," & _
    "Briques creuses,Briques Monomur,Briques pleines,Pierre de taille"

Private Const FEUILLE_POIDS As String = "Poids matériaux"

Public Sub RemplirDesignationOuvrage()
    Dim listes As Variant, i As Long, e As Long, r As Long, debut As Long

    CreerFeuillePoids ActiveSheet.Parent
    r = 5
    For e = 6 To -1 Step -1            ' 6 = toiture, étages 5 à 1, 0 = RDC, -1 = soubassement
        If e = 6 Then
            listes = Array("Toiture,Toiture terrasse", PLANCHERS, POUTRES)
        ElseIf e = 0 Then
            listes = Array("RDC", POUTRES, POTEAUX, MURS)
        ElseIf e = -1 Then
            listes = Array("Soubassement", POUTRES, POTEAUX, MURS)
        Else
            listes = Array("Étage " & e, PLANCHERS, POUTRES, POTEAUX, MURS)
        End If

        debut = r                            ' ligne de titre du bloc
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

        ' Sous-total du bloc sur la ligne de titre (G vidé : pas de nombre à saisir)
        With ActiveSheet
            .Range("G" & debut).ClearContents
            .Range("H" & debut).Formula = "=SUBTOTAL(109,H" & debut + 1 & ":H" & r - 1 & ")"
            .Range("H" & debut).Font.Bold = True
        End With
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
        Case "Pierre de taille": EpaisseursMur = "30 cm,40 cm,50 cm,60 cm"
    End Select
End Function

' Crée la feuille « Poids matériaux » avec les poids volumiques par défaut,
' seulement si elle n'existe pas encore (les valeurs modifiées sont conservées).
' Valeurs indicatives à vérifier (normes, fabricants). Pour les murs creux,
' c'est un poids volumique apparent (vides compris).
Public Sub CreerFeuillePoids(classeur As Workbook)
    Dim ws As Worksheet, actif As Worksheet
    Dim donnees As Variant, i As Long

    On Error Resume Next
    Set ws = classeur.Worksheets(FEUILLE_POIDS)
    On Error GoTo 0
    If Not ws Is Nothing Then Exit Sub

    Set actif = ActiveSheet
    Set ws = classeur.Worksheets.Add(After:=classeur.Worksheets(classeur.Worksheets.Count))
    ws.Name = FEUILLE_POIDS
    ws.Range("A1:B1").Value = Array("Élément", "Poids volumique (kN/m3)")
    donnees = Array("Poutre BA", 25, "Poteau BA", 25, _
        "Poutre bois", 6, "Poteau bois", 6, _
        "Poutre métallique", 78.5, "Poteau métallique", 78.5, _
        "Agglos béton creux", 9, "Agglos béton plein", 20, _
        "Béton banché", 25, "Briques creuses", 9, _
        "Briques Monomur", 8, "Briques pleines", 18, _
        "Pierre de taille", 22)
    For i = 0 To UBound(donnees) Step 2
        ws.Cells(2 + i / 2, 1).Value = donnees(i)
        ws.Cells(2 + i / 2, 2).Value = donnees(i + 1)
    Next i
    ws.Range("A1:B1").Font.Bold = True
    ws.Columns("A:B").AutoFit
    actif.Activate
End Sub

' Formule qui lit le poids volumique de l'élément en B dans « Poids matériaux ».
Private Function FormulePoids(ligne As Long) As String
    FormulePoids = "IFERROR(VLOOKUP(B" & ligne & ",'" & FEUILLE_POIDS & "'!$A:$B,2,FALSE),0)"
End Function

' Remplit les poids de la ligne selon l'élément choisi en B :
'   poutre / poteau -> poids volumique en C ; mur -> épaisseur F et poids surfacique D.
Public Sub MajLigne(ws As Worksheet, ligne As Long)
    Dim typeElement As String

    typeElement = CStr(ws.Range("B" & ligne).Value)
    If Left(typeElement, 7) = "Toiture" Then
        AfficherEtages ws                    ' plancher de toiture affiché ou non
    ElseIf EpaisseursMur(typeElement) <> "" Then
        MajEpaisseurMur ws, ligne
    ElseIf Left(typeElement, 6) = "Poutre" Or Left(typeElement, 6) = "Poteau" Then
        ws.Range("C" & ligne).Formula = "=" & FormulePoids(ligne)
        ws.Range("D" & ligne).ClearContents
        ' Rappel sur l'épaisseur h d'une poutre (commentaire sur la cellule F)
        If Left(typeElement, 6) = "Poutre" Then
            ws.Range("F" & ligne).ClearComments
            ws.Range("F" & ligne).AddComment _
                "Hauteur h de la poutre : si une dalle est présente, prendre " & _
                "h = hauteur totale de la poutre - épaisseur de la dalle " & _
                "(retombée seule), pour ne pas compter deux fois le béton de la dalle."
        End If
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
    ws.Range("D" & ligne).Formula = "=" & FormulePoids(ligne) & "*F" & ligne
End Sub

' Affiche la toiture et les étages 1 à niveau (C1), masque les autres.
' RDC -> toiture seule (lignes 5 à 7) ; 1 -> toiture + étage 1 ; ... ; 5 -> tout.
' La ligne Plancher de la toiture (6) n'est affichée que pour une toiture terrasse.
' Appelée aussi automatiquement quand C1 change (voir CodeFeuille.txt).
Public Sub AfficherEtages(ws As Worksheet)
    Dim niveau As Long, e As Long

    ws.Rows(6).Hidden = (ws.Range("B5").Value <> "Toiture terrasse")
    niveau = Val(ws.Range("C1").Value)     ' "RDC" donne 0
    For e = 1 To 5
        ' Bloc de l'étage e : 5 lignes à partir de la ligne 8 + 5 × (5 - e)
        ws.Rows(8 + 5 * (5 - e)).Resize(5).Hidden = (e > niveau)
    Next e
End Sub
