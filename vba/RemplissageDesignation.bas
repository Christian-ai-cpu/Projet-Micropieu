Attribute VB_Name = "RemplissageDesignation"
Option Explicit

' Listes déroulantes sous « Désignation de l'ouvrage », à partir de B5.
' Chaque bloc se termine par une ligne « Sous-total » (somme du bloc en H) :
'   B5 à B8   : Toiture / Toiture terrasse, Plancher, Poutre, Sous-total
'               (la ligne Plancher n'est affichée que pour une toiture terrasse)
'   B9 à B14  : Étage 5, Plancher, Poutre, Poteau, Mur, Sous-total
'   ...
'   B33 à B38 : Étage 1, Plancher, Poutre, Poteau, Mur, Sous-total
'   B39 à B43 : RDC, Poutre, Poteau, Mur, Sous-total (toujours affiché)
'   B44 à B48 : Soubassement, Poutre, Poteau, Mur, Sous-total (toujours affiché)
' Le TOTAL G (SOUS.TOTAL) ignore les sous-totaux : pas de double compte.
' Pour un mur : le type se choisit en B, l'épaisseur en F (liste adaptée au type),
' et le poids surfacique D est calculé : D = poids volumique × épaisseur F.
' Pour une poutre ou un poteau : le poids volumique C est rempli selon le type.
' Les valeurs sont des ordres de grandeur à vérifier (normes, fabricants) :
' elles se modifient dans PoidsVolumique et PoidsVolumiqueMur.
' Seuls les étages jusqu'au niveau choisi en C1 sont affichés (voir AfficherEtages).

Private Const PLANCHERS As String = "Plancher dalle pleine,Plancher à poutrelles,Plancher bois traditionnel"
Private Const POUTRES As String = "Poutre BA,Poutre bois,Poutre métallique"
Private Const POTEAUX As String = "Poteau BA,Poteau bois,Poteau métallique"
Private Const MURS As String = "Agglos béton creux,Agglos béton plein,Béton banché," & _
    "Briques creuses,Briques Monomur,Briques pleines,Pierre de taille"

Public Sub RemplirDesignationOuvrage()
    Dim listes As Variant, i As Long, e As Long, r As Long, debut As Long
    Dim nomBloc As String

    r = 5
    For e = 6 To -1 Step -1            ' 6 = toiture, étages 5 à 1, 0 = RDC, -1 = soubassement
        If e = 6 Then
            listes = Array("Toiture,Toiture terrasse", PLANCHERS, POUTRES)
            nomBloc = "toiture"
        ElseIf e = 0 Then
            listes = Array("RDC", POUTRES, POTEAUX, MURS)
            nomBloc = "RDC"
        ElseIf e = -1 Then
            listes = Array("Soubassement", POUTRES, POTEAUX, MURS)
            nomBloc = "soubassement"
        Else
            listes = Array("Étage " & e, PLANCHERS, POUTRES, POTEAUX, MURS)
            nomBloc = "étage " & e
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
            ' Nombre n laissé vide pour les titres et les poteaux (à saisir)
            If i = 0 Or listes(i) = POTEAUX Then ActiveSheet.Range("G" & r).ClearContents
            r = r + 1
        Next i
        ActiveSheet.Range("H" & debut).ClearContents   ' pas de charge sur le titre

        ' Ligne de sous-total sous le dernier élément du bloc
        With ActiveSheet
            .Range("B" & r).Validation.Delete
            .Range("B" & r).Value = "Sous-total " & nomBloc
            .Range("G" & r).ClearContents
            .Range("H" & r).Formula = "=SUBTOTAL(109,H" & debut + 1 & ":H" & r - 1 & ")"
            .Range("B" & r & ":H" & r).Font.Bold = True
            .Range("B" & r).Font.Italic = True
        End With
        r = r + 1
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
        Case "Pierre de taille":   PoidsVolumiqueMur = 22     ' 30 cm -> 6,60 kN/m²
    End Select
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
    ElseIf PoidsVolumique(typeElement) > 0 Then
        ws.Range("C" & ligne).Value = PoidsVolumique(typeElement)
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
    ws.Range("D" & ligne).Formula = "=" & _
        Replace(CStr(PoidsVolumiqueMur(CStr(ws.Range("B" & ligne).Value))), ",", ".") & _
        "*F" & ligne
End Sub

' Affiche la toiture et les étages 1 à niveau (C1), masque les autres.
' RDC -> toiture seule ; 1 -> toiture + étage 1 ; ... ; 5 -> tout.
' La ligne Plancher de la toiture (6) n'est affichée que pour une toiture terrasse.
' Appelée aussi automatiquement quand C1 change (voir CodeFeuille.txt).
Public Sub AfficherEtages(ws As Worksheet)
    Dim niveau As Long, e As Long

    ws.Rows(6).Hidden = (ws.Range("B5").Value <> "Toiture terrasse")
    niveau = Val(ws.Range("C1").Value)     ' "RDC" donne 0
    For e = 1 To 5
        ' Bloc de l'étage e (sous-total compris) : 6 lignes à partir de 9 + 6 × (5 - e)
        ws.Rows(9 + 6 * (5 - e)).Resize(6).Hidden = (e > niveau)
    Next e
End Sub
