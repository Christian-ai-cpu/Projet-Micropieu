Attribute VB_Name = "ChargePermanente"
Option Explicit

' Module « Charge permanente » : remplissage du tableau des charges permanentes G.
' Listes déroulantes sous « Désignation de l'ouvrage », à partir de B5.
' Chaque bloc se termine par une ligne « Sous-total » (somme du bloc en I).
' La charge d'exploitation Q (colonne J) est remplie par le module ChargeExploitation.
'   B5 à B9   : Toiture / Toiture terrasse, Revêtement, Plancher, Poutre, Sous-total
'               (Revêtement et Plancher affichés seulement pour une toiture terrasse)
'   B10 à B16 : Étage 5, Revêtement, Plancher, Poutre, Poteau, Mur, Sous-total
'   ...
'   B38 à B44 : Étage 1, Revêtement, Plancher, Poutre, Poteau, Mur, Sous-total
'   B45 à B49 : RDC, Poutre, Poteau, Mur, Sous-total (toujours affiché)
'   B50 à B53 : Soubassement, Poteau, Mur, Sous-total (toujours affiché)
' Revêtement (au-dessus de chaque plancher) : charge permanente du revêtement de sol,
' sa valeur en kN/m² est mise en D (poids surfacique), voir ChargeRevetement.
' Sa charge se calcule par G = (linéaire L / 2) × poids surfacique D × largeur b.
' Plancher (épaisseur e à saisir en F) : poids surfacique selon le type
'   dalle pleine      : 25 × e                    (C = 25, D vide)
'   à poutrelles      : 0,05 × 25 + (e - 0,05) × 11  (C vide, D = cette formule)
'   bois traditionnel : 1,5 kN/m²                 (C vide, D = 1,5)
' Ligne 54 : TOTAL G = somme des sous-totaux (I9 + I16 + ... + I53). Le sous-total
' d'un étage masqué vaut 0 (SOUS.TOTAL ignore les lignes masquées).
' Pour un mur : le type se choisit en B, l'épaisseur en F (liste adaptée au type),
' et le poids surfacique D est calculé : D = poids volumique × épaisseur F.
' Pour une poutre ou un poteau : le poids volumique C est rempli selon le type.
' Les valeurs sont des ordres de grandeur à vérifier (normes, fabricants) :
' elles se modifient dans PoidsVolumique et PoidsVolumiqueMur.
' Seuls les étages jusqu'au niveau choisi en C1 sont affichés (voir AfficherEtages).

Private Const PLANCHERS As String = "Plancher dalle pleine,Plancher à poutrelles,Plancher bois traditionnel"
Private Const POUTRES As String = "Poutre BA,Poutre bois,Poutre métallique"
Private Const POTEAUX As String = "Poteau BA,Poteau bois,Poteau métallique"
Private Const REVETEMENTS As String = "Carrelage grès cérame,Dalles épaisses (marbre),Parquet traditionnel,Sols minces"
Private Const MURS As String = "Agglos béton creux,Agglos béton plein,Béton banché," & _
    "Briques creuses,Briques Monomur,Briques pleines,Pierre de taille"

Public Sub RemplirChargePermanente()
    Dim listes As Variant, i As Long, e As Long, r As Long, debut As Long
    Dim nomBloc As String, sousTotaux As String

    r = 5
    For e = 6 To -1 Step -1            ' 6 = toiture, étages 5 à 1, 0 = RDC, -1 = soubassement
        If e = 6 Then
            listes = Array("Toiture,Toiture terrasse", REVETEMENTS, PLANCHERS, POUTRES)
            nomBloc = "toiture"
        ElseIf e = 0 Then
            listes = Array("RDC", POUTRES, POTEAUX, MURS)
            nomBloc = "RDC"
        ElseIf e = -1 Then
            listes = Array("Soubassement", POTEAUX, MURS)
            nomBloc = "soubassement"
        Else
            listes = Array("Étage " & e, REVETEMENTS, PLANCHERS, POUTRES, POTEAUX, MURS)
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
                .IndentLevel = IIf(listes(i) = REVETEMENTS, 1, 0)   ' revêtement rattaché au plancher
            End With
            MajLigne ActiveSheet, r              ' poids volumique / surfacique
            ' Nombre n laissé vide pour les titres et les poteaux (à saisir)
            If i = 0 Or listes(i) = POTEAUX Then ActiveSheet.Range("G" & r).ClearContents
            r = r + 1
        Next i
        ActiveSheet.Range("I" & debut).ClearContents   ' pas de charge sur le titre

        ' Ligne de sous-total sous le dernier élément du bloc
        With ActiveSheet
            .Range("B" & r).Validation.Delete
            .Range("B" & r).Value = "Sous-total " & nomBloc
            .Range("G" & r).ClearContents
            .Range("I" & r).Formula = "=SUBTOTAL(109,I" & debut + 1 & ":I" & r - 1 & ")"
            .Range("B" & r & ":M" & r).Font.Bold = True
            .Range("B" & r).Font.Italic = True
        End With
        sousTotaux = sousTotaux & ",I" & r
        r = r + 1
    Next e

    ' TOTAL G = somme des sous-totaux de tous les blocs
    With ActiveSheet
        .Range("B" & r).Value = "TOTAL (somme des sous-totaux)"
        .Range("I" & r).Formula = "=SUM(" & Mid(sousTotaux, 2) & ")"
        .Range("B" & r & ":M" & r).Font.Bold = True
    End With

    AfficherEtages ActiveSheet

    ' Colonnes I (G) et J (Q) ajustées à leur contenu, avec un minimum pour le titre
    With ActiveSheet.Columns("I:J")
        .AutoFit
        If .Columns(1).ColumnWidth < 12 Then .Columns(1).ColumnWidth = 12
        If .Columns(2).ColumnWidth < 12 Then .Columns(2).ColumnWidth = 12
    End With
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

' Charge permanente (kN/m²) du revêtement de sol posé sur le plancher ; 0 sinon.
Public Function ChargeRevetement(typeRevetement As String) As Double
    Select Case typeRevetement
        Case "Carrelage grès cérame":    ChargeRevetement = 0.5   ' valeur donnée : 50, à vérifier
        Case "Dalles épaisses (marbre)": ChargeRevetement = 1
        Case "Parquet traditionnel":     ChargeRevetement = 0.25
        Case "Sols minces":              ChargeRevetement = 0.1
    End Select
End Function

' Poids volumique (kN/m³) d'une poutre ou d'un poteau selon son type ; 0 sinon.
Public Function PoidsVolumique(typeElement As String) As Double
    Select Case typeElement
        Case "Poutre BA", "Poteau BA":                 PoidsVolumique = 25
        Case "Plancher dalle pleine", "Plancher à poutrelles": PoidsVolumique = 25
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
' dc : décalage de colonnes du tableau (0 pour le tableau d'origine, 13 × k pour la
' copie du mur k + 1, voir le module DuplicationMurs). Les formules sont écrites en
' références relatives (R1C1) pour rester justes dans les copies.
Public Sub MajLigne(ws As Worksheet, ligne As Long, Optional dc As Long = 0)
    Dim typeElement As String

    typeElement = CStr(ws.Cells(ligne, 2 + dc).Value)
    If Left(typeElement, 7) = "Toiture" Then
        AfficherEtages ws                    ' plancher de toiture affiché ou non
    ElseIf EpaisseursMur(typeElement) <> "" Then
        MajEpaisseurMur ws, ligne, dc
    ElseIf Left(typeElement, 8) = "Plancher" Then
        MajPlancher ws, ligne, typeElement, dc
    ElseIf ChargeRevetement(typeElement) > 0 Then
        ws.Cells(ligne, 3 + dc).ClearContents
        ws.Cells(ligne, 4 + dc).Value = ChargeRevetement(typeElement)
        ' Charge du revêtement : (linéaire L / 2) × poids surfacique D × largeur b
        ws.Cells(ligne, 9 + dc).FormulaR1C1 = "=RC[-1]/2*RC[-5]*RC[-4]"
    ElseIf PoidsVolumique(typeElement) > 0 Then
        ws.Cells(ligne, 3 + dc).Value = PoidsVolumique(typeElement)
        ws.Cells(ligne, 4 + dc).ClearContents
        ' Rappel sur l'épaisseur h d'une poutre (commentaire sur la cellule F)
        If Left(typeElement, 6) = "Poutre" Then
            ws.Cells(ligne, 6 + dc).ClearComments
            ws.Cells(ligne, 6 + dc).AddComment _
                "Hauteur h de la poutre : si une dalle est présente, prendre " & _
                "h = hauteur totale de la poutre - épaisseur de la dalle " & _
                "(retombée seule), pour ne pas compter deux fois le béton de la dalle."
        End If
    End If
End Sub

' Poids du plancher selon son type ; la charge se calcule avec C × F + D (F = épaisseur).
Public Sub MajPlancher(ws As Worksheet, ligne As Long, typePlancher As String, _
        Optional dc As Long = 0)
    Select Case typePlancher
        Case "Plancher dalle pleine"            ' 25 × e
            ws.Cells(ligne, 3 + dc).Value = PoidsVolumique(typePlancher)
            ws.Cells(ligne, 4 + dc).ClearContents
        Case "Plancher à poutrelles"            ' 0,05 × 25 + (e - 0,05) × 11
            ws.Cells(ligne, 3 + dc).ClearContents
            ws.Cells(ligne, 4 + dc).FormulaR1C1 = "=IF(RC[2]="""",0,0.05*" & _
                PoidsVolumique(typePlancher) & "+(RC[2]-0.05)*11)"
        Case "Plancher bois traditionnel"       ' 1,5 kN/m²
            ws.Cells(ligne, 3 + dc).ClearContents
            ws.Cells(ligne, 4 + dc).Value = 1.5
    End Select
End Sub

' Met en F la liste des épaisseurs du mur choisi en B, et la première en mètres.
' Le poids surfacique D suit l'épaisseur F par formule ; C est vidé pour ne pas
' compter le poids du mur deux fois.
Public Sub MajEpaisseurMur(ws As Worksheet, ligne As Long, Optional dc As Long = 0)
    Dim liste As String, typeMur As String

    typeMur = CStr(ws.Cells(ligne, 2 + dc).Value)
    liste = EpaisseursMur(typeMur)
    If liste = "" Then Exit Sub
    With ws.Cells(ligne, 6 + dc)
        .Validation.Delete
        .Validation.Add xlValidateList, xlValidAlertStop, , liste
        .Value = Val(liste) / 100           ' "15 cm" -> 0.15 m
        .NumberFormat = "0.000"
    End With
    ws.Cells(ligne, 3 + dc).ClearContents
    ws.Cells(ligne, 4 + dc).FormulaR1C1 = "=" & _
        Replace(CStr(PoidsVolumiqueMur(typeMur)), ",", ".") & "*RC[2]"
End Sub

' Affiche la toiture et les étages 1 à niveau (C1), masque les autres.
' RDC -> toiture seule ; 1 -> toiture + étage 1 ; ... ; 5 -> tout.
' Les lignes Revêtement et Plancher de la toiture (6 et 7) ne sont affichées que
' pour une toiture terrasse.
' Appelée aussi automatiquement quand C1 change (voir CodeFeuille.txt).
Public Sub AfficherEtages(ws As Worksheet)
    Dim niveau As Long, e As Long

    ws.Rows("6:7").Hidden = (ws.Range("B5").Value <> "Toiture terrasse")
    niveau = Val(ws.Range("C1").Value)     ' "RDC" donne 0
    For e = 1 To 5
        ' Bloc de l'étage e (sous-total compris) : 7 lignes à partir de 10 + 7 × (5 - e)
        ws.Rows(10 + 7 * (5 - e)).Resize(7).Hidden = (e > niveau)
    Next e
End Sub
