Attribute VB_Name = "ChargePermanente"
Option Explicit

' Module « Charge permanente » : remplissage du tableau des charges permanentes G.
' Listes déroulantes sous « Désignation de l'ouvrage », à partir de B5.
' Descente de charges en kN/ml pour un mur : G = b × (γ × h + g), b = largeur reprise (m).
' Chaque bloc se termine par une ligne « Sous-total » (somme du bloc en G).
' La charge d'exploitation Q (colonne H) est remplie par le module ChargeExploitation.
'   B5 à B10  : Toiture (liste), Charge, Revêtement, Plancher, Poutre, Sous-total
'               Toiture : fermettes industrielles + tuiles, fermettes industrielles +
'               ardoises, charpente traditionnelle bois massif, chacune avec la pente
'               du versant dans la désignation (18, 20, 25, 30 ou 45°), ou toiture
'               terrasse. Toiture en pente : ligne « Charge » (6) avec le poids
'               surfacique D = g / cos(pente) (voir ChargeToiture), b saisi en E et
'               G = b × D. Toiture terrasse : lignes Revêtement et Plancher (7 et 8)
'               à la place de la ligne Charge.
'   B11 à B17 : Étage 5, Revêtement, Plancher, Poutre, Poteau, Mur, Sous-total
'   ...
'   B39 à B45 : Étage 1, Revêtement, Plancher, Poutre, Poteau, Mur, Sous-total
'   B46 à B52 : RDC (liste), Revêtement, Plancher, Poutre, Poteau, Mur, Sous-total
'               (toujours affiché). Liste en B46 : RDC, RDC+1SS (1 sous-sol),
'               RDC+2SS (2 sous-sols), RDC+VS (vide sanitaire). Revêtement et
'               Plancher du RDC affichés sauf pour « RDC » seul.
'   B53 à B59 : Sous-sol 1, Revêtement, Plancher, Poutre, Poteau, Mur, Sous-total
'               (affiché pour RDC+1SS et RDC+2SS ; Revêtement et Plancher
'               seulement pour RDC+2SS)
'   B60 à B64 : Sous-sol 2, Poutre, Poteau, Mur, Sous-total (RDC+2SS seulement)
'   B65 à B68 : Vide sanitaire, Poteau, Mur, Sous-total (RDC+VS seulement)
'   B69 à B73 : Soubassement, Poteau, Semelle (liste par matériau), Mur, Sous-total
'               (toujours affiché)
' Une liste de plus de 255 caractères (limite d'Excel, ex. toitures) est écrite sur
' la feuille masquée « Listes » et la liste déroulante y renvoie (voir AjouterListe).
' Revêtement (au-dessus de chaque plancher) : charge permanente du revêtement de sol,
' sa valeur en kN/m² est mise en D (poids surfacique), voir ChargeRevetement.
' Sa charge se calcule par G = poids surfacique D × largeur b (kN/ml).
' Plancher (épaisseur e à saisir en F) : poids surfacique selon le type
'   dalle pleine      : 25 × e                    (C = 25, D vide)
'   à poutrelles      : 0,05 × 25 + (e - 0,05) × 11  (C vide, D = cette formule)
'   bois traditionnel : 1,5 kN/m²                 (C vide, D = 1,5)
' Ligne 74 : TOTAL G = somme des sous-totaux (G10 + G17 + ... + G73). Le sous-total
' d'un étage masqué vaut 0 (SOUS.TOTAL ignore les lignes masquées).
' Pour un mur : le type se choisit en B, l'épaisseur en E (liste adaptée au type),
' la hauteur h du mur se saisit en F, et :
'   D = poids volumique × épaisseur E + 0,2   (0,2 kN/m² d'enduits et d'isolant
'       thermique ou acoustique, voir ENDUIT_ISOLANT)
'   G = (100 - % d'ouvertures I1) / 100 × D × h   (kN/ml ; I1 pour le mur 1,
'       T1, AE1... pour les copies : chaque mur a son propre pourcentage)
' Pour une poutre ou un poteau : le poids volumique C est rempli selon le type.
' Les valeurs sont des ordres de grandeur à vérifier (normes, fabricants) :
' elles se modifient dans PoidsVolumique et PoidsVolumiqueMur.
' Seuls les étages jusqu'au niveau choisi en C1 sont affichés (voir AfficherEtages).

' Toitures en pente (type + pente du versant dans la désignation, voir ListeToitures)
Private Const TYPES_TOITURE As String = "Toiture fermettes industrielles + tuiles," & _
    "Toiture fermettes industrielles + ardoises,Toiture charpente traditionnelle bois massif"
Private Const PENTES As String = "18,20,25,30,45"   ' pentes des versants (°)

' Thème : saisies en vert, valeurs calculées en jaune (comme les autres feuilles)
Private Const VERT_SAISIE As Long = 13561798   ' RGB(198, 239, 206)
Private Const JAUNE_CALCUL As Long = 13434879  ' RGB(255, 255, 204)
Private Const NIVEAUX_RDC As String = "RDC,RDC+1SS,RDC+2SS,RDC+VS"   ' liste en B46
Private Const PLANCHERS As String = "Plancher dalle pleine,Plancher à poutrelles,Plancher bois traditionnel"
Private Const POUTRES As String = "Poutre BA,Poutre bois,Poutre métallique"
Private Const POTEAUX As String = "Poteau BA,Poteau bois,Poteau métallique"
Private Const SEMELLES As String = "Semelle BA,Semelle béton non armé,Semelle gros béton," & _
    "Semelle maçonnerie de moellons"
Private Const REVETEMENTS As String = "Carrelage grès cérame,Dalles épaisses (marbre),Parquet traditionnel,Sols minces"
Private Const ENDUIT_ISOLANT As Double = 0.2   ' kN/m², enduits + isolant (ajouté aux murs)
' Intitulé des murs : type + « + enduit et isolant » (0,2 kN/m² compris dans D).
' La liste fait 255 caractères, le maximum accepté par Excel pour une liste.
Private Const SUFFIXE_MUR As String = " + enduit et isolant"
Private Const MURS As String = "Agglos béton creux + enduit et isolant," & _
    "Agglos béton plein + enduit et isolant,Béton banché + enduit et isolant," & _
    "Briques creuses + enduit et isolant,Briques Monomur + enduit et isolant," & _
    "Briques pleines + enduit et isolant,Pierre de taille + enduit et isolant"

Public Sub RemplirChargePermanente()
    Dim listes As Variant, i As Long, e As Long, r As Long, debut As Long
    Dim nomBloc As String, sousTotaux As String

    Application.ScreenUpdating = False       ' pas de scintillement pendant l'exécution
    Application.EnableEvents = False
    r = 5
    ' 6 = toiture, étages 5 à 1, 0 = RDC, -1 = sous-sol 1, -2 = sous-sol 2,
    ' -3 = vide sanitaire, -4 = soubassement
    For e = 6 To -4 Step -1
        If e = 6 Then
            listes = Array(ListeToitures(), "Charge", REVETEMENTS, PLANCHERS, POUTRES)
            nomBloc = "toiture"
        ElseIf e = 0 Then
            listes = Array(NIVEAUX_RDC, REVETEMENTS, PLANCHERS, POUTRES, POTEAUX, MURS)
            nomBloc = "RDC"
        ElseIf e = -1 Then
            listes = Array("Sous-sol 1", REVETEMENTS, PLANCHERS, POUTRES, POTEAUX, MURS)
            nomBloc = "sous-sol 1"
        ElseIf e = -2 Then
            listes = Array("Sous-sol 2", POUTRES, POTEAUX, MURS)
            nomBloc = "sous-sol 2"
        ElseIf e = -3 Then
            listes = Array("Vide sanitaire", POTEAUX, MURS)
            nomBloc = "vide sanitaire"
        ElseIf e = -4 Then
            listes = Array("Soubassement", POTEAUX, SEMELLES, MURS)
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
                If InStr(listes(i), ",") > 0 Then AjouterListe ActiveSheet.Range("B" & r), _
                    CStr(listes(i)), "Liste " & nomBloc
                .Value = Split(listes(i), ",")(0)
                .Font.Bold = (i = 0)          ' Toiture, Étage n, RDC et Soubassement en gras
                .IndentLevel = IIf(listes(i) = REVETEMENTS, 1, 0)   ' revêtement rattaché au plancher
            End With
            MajLigne ActiveSheet, r              ' poids volumique / surfacique
            ' Code couleur : désignation (liste) et dimensions b, h à saisir en vert ;
            ' poids volumique / surfacique et charge G calculés en jaune
            With ActiveSheet
                If InStr(listes(i), ",") > 0 Then .Range("B" & r).Interior.Color = VERT_SAISIE
                If i > 0 Then
                    .Range("C" & r & ":D" & r & ",G" & r).Interior.Color = JAUNE_CALCUL
                    .Range("E" & r & ":F" & r).Interior.Color = VERT_SAISIE
                End If
            End With
            r = r + 1
        Next i
        ActiveSheet.Range("G" & debut).ClearContents   ' pas de charge sur le titre

        ' Ligne de sous-total sous le dernier élément du bloc
        With ActiveSheet
            .Range("B" & r).Validation.Delete
            .Range("B" & r).Value = "Sous-total " & nomBloc
            .Range("G" & r).Formula = "=SUBTOTAL(109,G" & debut & ":G" & r - 1 & ")"
            .Range("B" & r & ":K" & r).Font.Bold = True
            .Range("B" & r).Font.Italic = True
            .Range("G" & r & ":K" & r).Interior.Color = JAUNE_CALCUL   ' G, Q, ΣQ, αn, αn × ΣQ
        End With
        sousTotaux = sousTotaux & ",G" & r
        r = r + 1
    Next e

    ' TOTAL G = somme des sous-totaux de tous les blocs
    With ActiveSheet
        .Range("B" & r).Value = "TOTAL (somme des sous-totaux)"
        .Range("G" & r).Formula = "=SUM(" & Mid(sousTotaux, 2) & ")"
        .Range("B" & r & ":K" & r).Font.Bold = True
        .Range("G" & r & ":K" & r).Interior.Color = JAUNE_CALCUL
    End With

    AfficherEtages ActiveSheet

    ' Colonnes G (charges G) et H (Q) ajustées à leur contenu, avec un minimum pour le titre
    With ActiveSheet.Columns("G:H")
        .AutoFit
        If .Columns(1).ColumnWidth < 12 Then .Columns(1).ColumnWidth = 12
        If .Columns(2).ColumnWidth < 12 Then .Columns(2).ColumnWidth = 12
    End With
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
' Épaisseurs possibles (en cm) selon le type de mur ; "" si ce n'est pas un mur.
' (le suffixe « + enduit et isolant » de l'intitulé est ignoré)
Public Function EpaisseursMur(typeMur As String) As String
    Select Case Replace(typeMur, SUFFIXE_MUR, "")
        Case "Agglos béton creux", "Agglos béton plein": EpaisseursMur = "15 cm,20 cm"
        Case "Béton banché":    EpaisseursMur = "20 cm,25 cm,30 cm"
        Case "Briques creuses": EpaisseursMur = "15 cm,20 cm,25 cm"
        Case "Briques Monomur": EpaisseursMur = "30 cm,37 cm"
        Case "Briques pleines": EpaisseursMur = "10 cm,21.5 cm,33 cm"
        Case "Pierre de taille": EpaisseursMur = "30 cm,40 cm,50 cm,60 cm"
    End Select
End Function

' Liste des toitures pour la désignation : chaque type de toiture en pente avec
' chaque pente de versant (ex. « Toiture fermettes industrielles + tuiles 18° »),
' puis la toiture terrasse. Plus de 255 caractères : voir AjouterListe.
Public Function ListeToitures() As String
    Dim t As Variant, p As Variant, s As String
    For Each t In Split(TYPES_TOITURE, ",")
        For Each p In Split(PENTES, ",")
            s = s & "," & t & " " & p & "°"
        Next p
    Next t
    ListeToitures = Mid(s, 2) & ",Toiture terrasse"
End Function

' Pente du versant (°) lue à la fin de la désignation (« ... 25° » -> 25) ; 0 sinon.
Public Function PenteToiture(nomToiture As String) As Double
    PenteToiture = Val(Mid(nomToiture, InStrRev(nomToiture, " ") + 1))
End Function

' Poids surfacique (kN/m² de versant) d'une toiture en pente selon son type
' (charpente + couverture), la pente en fin de désignation étant ignorée ; ramené
' à l'horizontale sur la ligne « Charge » (÷ cos de la pente) puis multiplié par la
' largeur reprise b ; 0 pour la toiture terrasse (calculée avec ses lignes
' Revêtement et Plancher). Ordres de grandeur à vérifier.
Public Function ChargeToiture(nomToiture As String) As Double
    Dim typeToiture As String
    typeToiture = nomToiture
    If PenteToiture(nomToiture) > 0 Then typeToiture = Left(nomToiture, InStrRev(nomToiture, " ") - 1)
    Select Case typeToiture
        Case "Toiture fermettes industrielles + tuiles":      ChargeToiture = 0.7
        Case "Toiture fermettes industrielles + ardoises":    ChargeToiture = 0.55
        Case "Toiture charpente traditionnelle bois massif":  ChargeToiture = 1
    End Select
End Function

' Met une liste déroulante sur la cellule c. Au-delà de 255 caractères (limite
' d'Excel), les choix sont écrits dans une colonne de la feuille masquée « Listes »
' (titre nom en ligne 1) et la liste déroulante renvoie à cette colonne.
Public Sub AjouterListe(c As Range, liste As String, nom As String)
    Dim wl As Worksheet, wb As Workbook, choix As Variant, col As Long

    c.Validation.Delete
    If Len(liste) <= 255 Then
        c.Validation.Add xlValidateList, xlValidAlertStop, , liste
        Exit Sub
    End If
    Set wb = c.Worksheet.Parent
    On Error Resume Next
    Set wl = wb.Worksheets("Listes")
    On Error GoTo 0
    If wl Is Nothing Then
        Set wl = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wl.Name = "Listes"
        wl.Visible = xlSheetHidden
        c.Worksheet.Activate                 ' la feuille ajoutée était devenue active
    End If
    ' Colonne de cette liste (même titre), sinon la première colonne libre
    col = 1
    Do While wl.Cells(1, col).Value <> "" And wl.Cells(1, col).Value <> nom
        col = col + 1
    Loop
    choix = Split(liste, ",")
    wl.Columns(col).ClearContents
    wl.Cells(1, col).Value = nom
    wl.Cells(2, col).Resize(UBound(choix) + 1, 1).Value = Application.Transpose(choix)
    c.Validation.Add xlValidateList, xlValidAlertStop, , "=Listes!" & _
        wl.Cells(2, col).Resize(UBound(choix) + 1, 1).Address
End Sub

' Charge permanente (kN/m²) du revêtement de sol posé sur le plancher ; 0 sinon.
Public Function ChargeRevetement(typeRevetement As String) As Double
    Select Case typeRevetement
        Case "Carrelage grès cérame":    ChargeRevetement = 0.5   ' valeur donnée : 50, à vérifier
        Case "Dalles épaisses (marbre)": ChargeRevetement = 1
        Case "Parquet traditionnel":     ChargeRevetement = 0.25
        Case "Sols minces":              ChargeRevetement = 0.1
    End Select
End Function

' Poids volumique (kN/m³) d'une poutre, d'un poteau ou d'une semelle selon son type ; 0 sinon.
Public Function PoidsVolumique(typeElement As String) As Double
    Select Case typeElement
        Case "Poutre BA", "Poteau BA":                 PoidsVolumique = 25
        Case "Plancher dalle pleine", "Plancher à poutrelles": PoidsVolumique = 25
        Case "Semelle BA":                             PoidsVolumique = 25
        Case "Semelle béton non armé":                 PoidsVolumique = 24
        Case "Semelle gros béton":                     PoidsVolumique = 22
        Case "Semelle maçonnerie de moellons":         PoidsVolumique = 23
        Case "Poutre bois", "Poteau bois":             PoidsVolumique = 6
        Case "Poutre métallique", "Poteau métallique": PoidsVolumique = 78.5
    End Select
End Function

' Poids volumique apparent (kN/m³) d'un mur, vides des blocs creux compris.
' Poids surfacique du mur = cette valeur × épaisseur (m).
Public Function PoidsVolumiqueMur(typeMur As String) As Double
    Select Case Replace(typeMur, SUFFIXE_MUR, "")
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
'   poutre / poteau -> poids volumique en C ; mur -> épaisseur E et poids surfacique D.
' dc : décalage de colonnes du tableau (0 pour le tableau d'origine, 11 × k pour la
' copie du mur k + 1, voir le module DuplicationMurs). Les formules sont écrites en
' références relatives (R1C1) pour rester justes dans les copies.
Public Sub MajLigne(ws As Worksheet, ligne As Long, Optional dc As Long = 0)
    Dim typeElement As String

    typeElement = CStr(ws.Cells(ligne, 2 + dc).Value)
    If Left(typeElement, 7) = "Toiture" Then
        ' Rien sur la ligne de titre : la toiture en pente est calculée sur la ligne
        ' « Charge » juste en dessous, la toiture terrasse sur Revêtement et Plancher
        ws.Cells(ligne, 4 + dc).ClearContents
        ws.Cells(ligne, 6 + dc).Validation.Delete
        ws.Cells(ligne, 6 + dc).ClearContents
        ws.Cells(ligne, 7 + dc).ClearContents
        If ChargeToiture(typeElement) > 0 Then
            ' D = g / cos(pente de la désignation), G = b × D (kN/ml), b saisi en E
            ws.Cells(ligne + 1, 3 + dc).ClearContents
            ws.Cells(ligne + 1, 4 + dc).Formula = "=" & _
                Replace(CStr(ChargeToiture(typeElement)), ",", ".") & "/COS(RADIANS(" & _
                PenteToiture(typeElement) & "))"
            ws.Cells(ligne + 1, 7 + dc).FormulaR1C1 = _
                "=IF(RC[-5]="""","""",RC[-2]*(RC[-4]*RC[-1]+RC[-3]))"
        Else
            ws.Cells(ligne + 1, 4 + dc).ClearContents   ' terrasse : ligne Charge masquée
        End If
        AfficherEtages ws                    ' ligne Charge ou plancher de toiture affichés
    ElseIf Left(typeElement, 3) = "RDC" Then
        AfficherEtages ws                    ' sous-sols / vide sanitaire affichés ou non
    ElseIf EpaisseursMur(typeElement) <> "" Then
        MajEpaisseurMur ws, ligne, dc
    ElseIf Left(typeElement, 8) = "Plancher" Then
        MajPlancher ws, ligne, typeElement, dc
    ElseIf ChargeRevetement(typeElement) > 0 Then
        ws.Cells(ligne, 3 + dc).ClearContents
        ws.Cells(ligne, 4 + dc).Value = ChargeRevetement(typeElement)
        ' Charge du revêtement : formule du tableau, b × D (kN/ml)
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

' Met en E la liste des épaisseurs du mur choisi en B, et la première en mètres.
' Le poids surfacique D suit l'épaisseur E par formule (γ × E + 0,2 d'enduits et
' isolant) ; C est vidé pour ne pas compter le poids du mur deux fois.
' La hauteur h du mur se saisit en F. La charge G du mur tient compte du
' pourcentage d'ouvertures de son tableau (I1, T1, AE1...) : G = (100 - I1) / 100 × D × h.
Public Sub MajEpaisseurMur(ws As Worksheet, ligne As Long, Optional dc As Long = 0)
    Dim liste As String, typeMur As String

    typeMur = CStr(ws.Cells(ligne, 2 + dc).Value)
    liste = EpaisseursMur(typeMur)
    If liste = "" Then Exit Sub
    With ws.Cells(ligne, 5 + dc)            ' épaisseur en E
        .Validation.Delete
        .Validation.Add xlValidateList, xlValidAlertStop, , liste
        .Value = Val(liste) / 100           ' "15 cm" -> 0.15 m
        .NumberFormat = "0.000"
    End With
    ws.Cells(ligne, 6 + dc).Validation.Delete   ' F : hauteur h du mur, saisie libre
    ws.Cells(ligne, 3 + dc).ClearContents
    ws.Cells(ligne, 4 + dc).FormulaR1C1 = "=" & _
        Replace(CStr(PoidsVolumiqueMur(typeMur)), ",", ".") & "*RC[1]+" & _
        Replace(CStr(ENDUIT_ISOLANT), ",", ".")
    ' G = (100 - % d'ouvertures en ligne 1, colonne I du même tableau) / 100 × D × h
    ws.Cells(ligne, 7 + dc).FormulaR1C1 = "=IF(RC[-5]="""","""",(100-R1C[2])/100*RC[-3]*RC[-1])"
End Sub

' Affiche la toiture et les étages 1 à niveau (C1), masque les autres.
' RDC -> toiture seule ; 1 -> toiture + étage 1 ; ... ; 5 -> tout.
' Toiture en pente : ligne Charge (6) affichée ; toiture terrasse : lignes
' Revêtement et Plancher (7 et 8) affichées à la place. Selon le choix en B46 :
'   RDC     : ni plancher du RDC, ni sous-sol, ni vide sanitaire
'   RDC+1SS : revêtement et plancher du RDC (47-48) + sous-sol 1 (53-59, sans 54-55)
'   RDC+2SS : idem + revêtement et plancher du sous-sol 1 (54-55) + sous-sol 2 (60-64)
'   RDC+VS  : revêtement et plancher du RDC (47-48) + vide sanitaire (65-68)
' Appelée aussi automatiquement quand C1 change (voir CodeFeuille.txt).
Public Sub AfficherEtages(ws As Worksheet)
    Dim niveau As Long, e As Long, rdc As String, terrasse As Boolean

    terrasse = (ws.Range("B5").Value = "Toiture terrasse")
    ws.Rows(6).Hidden = terrasse
    ws.Rows("7:8").Hidden = Not terrasse
    rdc = CStr(ws.Range("B46").Value)
    ws.Rows("47:48").Hidden = (rdc = "RDC")
    ws.Rows("53:59").Hidden = (rdc <> "RDC+1SS" And rdc <> "RDC+2SS")   ' sous-sol 1
    ws.Rows("54:55").Hidden = (rdc <> "RDC+2SS")   ' plancher du sous-sol 1
    ws.Rows("60:64").Hidden = (rdc <> "RDC+2SS")   ' sous-sol 2
    ws.Rows("65:68").Hidden = (rdc <> "RDC+VS")    ' vide sanitaire
    niveau = Val(ws.Range("C1").Value)     ' "RDC" donne 0
    For e = 1 To 5
        ' Bloc de l'étage e (sous-total compris) : 7 lignes à partir de 11 + 7 × (5 - e)
        ws.Rows(11 + 7 * (5 - e)).Resize(7).Hidden = (e > niveau)
    Next e
End Sub
