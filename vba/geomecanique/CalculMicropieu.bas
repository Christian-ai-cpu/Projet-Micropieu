Attribute VB_Name = "CalculMicropieu"
Option Explicit

' Module « Calcul et vérification » : crée (ou recrée) la feuille « Calcul et
' vérification » des micropieux, à partir des feuilles « Données du projet »
' (module DonneesProjet) et « Tableaux NF P94-262 » (module TableauxNFP94262).
' Toutes les valeurs sont des formules : elles suivent les saisies de ces feuilles.
'   B3  : Données du projet (type de pieu, méthode de calcul, sol d'ancrage = dernière
'         couche du tableau des données géotechniques)
'   B8  : Caractéristiques du micropieu (abréviation, norme, classe, catégorie, mode
'         de fonctionnement, résistance de pointe, B, D, Ab, Pe, frottement négligé)
'   B22 : Caractéristiques de la couche porteuse (valeurs de la dernière couche,
'         courbe fsol et a, b, c du tableau F.5.2.2, alpha pieu-sol du tableau
'         F.5.2.1, fsol, alpha x fsol, qs max du tableau F.5.2.3, qs retenu)
'   B43 : Frottement axial unitaire par couche (mêmes calculs pour chaque couche)
' Méthode pressiométrique (annexe F) : fsol = (a.pl* + b)(1 - exp(-c.pl*)), pl* en
' MPa, fsol en kPa ; qs = min(alpha pieu-sol x fsol ; qs max).
' Catégories de sol de la feuille « Données du projet » -> colonnes des tableaux :
'   tableaux à 5 colonnes (F.5.2.1, F.5.2.2) : argiles et sols intermédiaires -> 1re
'   colonne, sables et graves -> 2e, craies -> 3e, marnes -> 4e, roches -> 5e ;
'   tableau qs max à 6 colonnes : dans l'ordre de la liste des catégories de sol.

Private Const NOM_FEUILLE As String = "Calcul et vérification"
Private Const DONNEES As String = "'Données du projet'!"
Private Const TABLEAUX As String = "Tableaux NF P94-262"

Public Sub CreerCalculVerification()
    Dim ws As Worksheet, wt As Worksheet, r As Long, i As Long, lc As Long
    Dim zAlpha As String, zABC As String, zQs As String, nb As String, cat As String
    Dim lg As Long, sol As String, pl As String, etape As String

    Application.ScreenUpdating = False
    On Error GoTo Erreur

    ' Feuille des tableaux de la norme : créée si elle manque
    etape = "feuille Tableaux NF P94-262"
    On Error Resume Next
    Set wt = ActiveWorkbook.Worksheets(TABLEAUX)
    On Error GoTo Erreur
    If wt Is Nothing Then
        Application.Run "CreerTableauxNFP94262"
        Application.ScreenUpdating = False
        Set wt = ActiveWorkbook.Worksheets(TABLEAUX)
    End If
    ' Plages de valeurs des tableaux utilisés (adresses absolues sur cette feuille)
    zAlpha = ZoneTableau(wt, "Tableau F.5.2.1", 0, 20, 5)   ' 20 catégories x 5 sols
    zABC = ZoneTableau(wt, "Tableau F.5.2.2", 0, 4, 5)      ' courbe, a, b, c x 5 sols
    zQs = ZoneTableau(wt, "Tableau F.5.2.3", 0, 20, 6)      ' 20 catégories x 6 sols

    etape = "création de la feuille"
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo Erreur
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = NOM_FEUILLE
    End If
    ws.Cells.Clear

    nb = DONNEES & "$" & Replace(CELLULE_NB_COUCHES, "C", "C$")   ' nombre de couches
    lc = LIGNE_COUCHES                       ' première couche dans « Données du projet »
    lg = LIGNE_COUCHES + MAX_COUCHES - 1     ' dernière ligne prévue

    With ws.Range("B1:I1")
        .Merge
        .Value = "CALCUL ET VÉRIFICATION DES MICROPIEUX"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With

    ' --- Données du projet ---
    etape = "données du projet"
    SousTitre ws, 3, "Données du projet"
    Ligne ws, 4, "Type de pieu", "=" & DONNEES & "D" & LIGNE_PIEU, ""
    Ligne ws, 5, "Méthode de calcul", "=" & DONNEES & "D" & LIGNE_COUCHES + MAX_COUCHES + 5, ""
    Ligne ws, 6, "Type de sol d'ancrage (dernière couche)", _
        "=INDEX(" & DONNEES & "$D$" & lc & ":$D$" & lg & "," & nb & ")", ""

    ' --- Caractéristiques du micropieu ---
    etape = "caractéristiques du micropieu"
    SousTitre ws, 8, "Caractéristiques du micropieu"
    Ligne ws, 9, "Abréviation", "=" & DONNEES & "D" & LIGNE_PIEU + 7, ""
    Ligne ws, 10, "Norme de référence", "=" & DONNEES & "D" & LIGNE_PIEU + 4, ""
    Ligne ws, 11, "Classe du micropieu", "=" & DONNEES & "D" & LIGNE_PIEU + 5, ""
    Ligne ws, 12, "Catégorie du micropieu", "=" & DONNEES & "D" & LIGNE_PIEU + 6, ""
    Ligne ws, 13, "Mode de fonctionnement", "=" & DONNEES & "D" & LIGNE_PIEU + 8, ""
    Ligne ws, 14, "Prise en compte de la résistance de pointe", "=" & DONNEES & "D" & LIGNE_PIEU + 9, ""
    Ligne ws, 15, "Largeur du micropieu B", "=" & DONNEES & "D" & LIGNE_PIEU + 10, "m"
    Ligne ws, 16, "Hauteur totale du micropieu dans la couche porteuse D", _
        "=" & DONNEES & "D" & LIGNE_PIEU + 12, "m"
    Ligne ws, 17, "Aire de la base du micropieu Ab = " & ChrW(960) & ".B² / 4", _
        "=IF(N(D15)=0,"""",PI()*D15^2/4)", "m²"
    Ligne ws, 18, "Périmètre du micropieu Pe = " & ChrW(960) & ".B", _
        "=IF(N(D15)=0,"""",PI()*D15)", "m"
    ' Épaisseur des couches dont le terrain est neutralisé (frottement négligé)
    Ligne ws, 19, "Frottement négligé (épaisseur des terrains neutralisés)", _
        "=IF(COUNTIF(OFFSET(" & DONNEES & "$I$" & lc & ",0,0," & nb & "),""Oui"")=0,""Non""," & _
        "SUMIF(OFFSET(" & DONNEES & "$I$" & lc & ",0,0," & nb & "),""Oui""," & _
        "OFFSET(" & DONNEES & "$F$" & lc & ",0,0," & nb & ")))", "m"
    ws.Range("D15:D19").NumberFormat = "0.000"
    cat = "$D$12"                            ' catégorie du micropieu (cette feuille)

    ' --- Caractéristiques de la couche porteuse (dernière couche) ---
    etape = "couche porteuse"
    SousTitre ws, 21, "Caractéristiques de la couche porteuse"
    r = 22
    Ligne ws, r, "N° de couche", "=" & nb, ""
    Ligne ws, r + 1, "Nom de couche (NF P11-300)", ValeurCouche("C", lc, lg, nb), ""
    Ligne ws, r + 2, "Catégorie de sol (NF P94-262)", ValeurCouche("D", lc, lg, nb), ""
    Ligne ws, r + 3, "Profondeur de la base", ValeurCouche("E", lc, lg, nb), "m/TN"
    Ligne ws, r + 4, "Épaisseur", ValeurCouche("F", lc, lg, nb), "m"
    Ligne ws, r + 5, "Pression limite pl*", ValeurCouche("G", lc, lg, nb), "MPa"
    Ligne ws, r + 6, "Module pressiométrique EM", ValeurCouche("H", lc, lg, nb), "MPa"
    Ligne ws, r + 7, "Terrain neutralisé", ValeurCouche("I", lc, lg, nb), ""
    sol = "$D$" & r + 2
    pl = "$D$" & r + 5
    Ligne ws, r + 8, "Courbe fsol (tableau F.5.2.2)", FormuleABC(zABC, 1, sol), ""
    Ligne ws, r + 9, "Paramètre a", FormuleABC(zABC, 2, sol), ""
    Ligne ws, r + 10, "Paramètre b", FormuleABC(zABC, 3, sol), ""
    Ligne ws, r + 11, "Paramètre c", FormuleABC(zABC, 4, sol), ""
    Ligne ws, r + 12, ChrW(945) & " pieu-sol (tableau F.5.2.1)", FormuleAlpha(zAlpha, cat, sol), ""
    Ligne ws, r + 13, "fsol = (a.pl* + b)(1 - exp(-c.pl*))", _
        FormuleFsol(pl, "$D$" & r + 9, "$D$" & r + 10, "$D$" & r + 11), "kPa"
    Ligne ws, r + 14, ChrW(945) & " pieu-sol x fsol", FormuleProduit("$D$" & r + 12, "$D$" & r + 13), "kPa"
    Ligne ws, r + 15, "qs max (tableau F.5.2.3)", FormuleQsMax(zQs, cat, sol), "kPa"
    Ligne ws, r + 16, "qs retenu = min(" & ChrW(945) & " x fsol ; qs max)", _
        FormuleQs("$D$" & r + 14, "$D$" & r + 15), "kPa"
    ws.Range("D" & r + 9 & ":D" & r + 11).NumberFormat = "0.0000"
    ws.Range("D" & r + 12).NumberFormat = "0.00"
    ws.Range("D" & r + 13 & ":D" & r + 16).NumberFormat = "0.0"

    ' --- Frottement axial unitaire par couche ---
    etape = "frottement par couche"
    r = 41
    SousTitre ws, r, "Frottement axial unitaire par couche (méthode pressiométrique)"
    ws.Range("B" & r + 1 & ":L" & r + 1).Value = Array("N°", "Nom de couche", "Catégorie de sol", _
        "Épaisseur (m)", "pl* (MPa)", "Terrain neutralisé", "Courbe", ChrW(945) & " pieu-sol", _
        "fsol (kPa)", ChrW(945) & " x fsol (kPa)", "qs max (kPa)")
    ws.Range("M" & r + 1).Value = "qs retenu (kPa)"
    For i = 1 To MAX_COUCHES
        CoucheLigne ws, r + 1 + i, i, lc + i - 1, nb, cat, zAlpha, zABC, zQs
    Next i

    ' --- Mise en forme ---
    etape = "mise en forme"
    With ws.Range("B4:E6,B9:E19,B22:E38,B" & r + 1 & ":M" & r + 1 + MAX_COUCHES)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("B" & r + 1 & ":M" & r + 1)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(217, 225, 242)
    End With
    ws.Range("D4:E38,B" & r + 2 & ":B" & r + 1 + MAX_COUCHES & ",E" & r + 2 & ":M" & _
        r + 1 + MAX_COUCHES).HorizontalAlignment = xlCenter
    ws.Range("D38").Font.Bold = True             ' qs retenu de la couche porteuse
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:M").ColumnWidth = 12
    ws.Activate

Fin:
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    MsgBox "Erreur pendant la création de la feuille « Calcul et vérification », partie « " & _
        etape & " » : " & Err.Description & " (n° " & Err.Number & ")", vbExclamation
    Resume Fin
End Sub

' Une ligne « intitulé (B:C fusionnées) | valeur ou formule (D) | unité (E) »
Private Sub Ligne(ws As Worksheet, r As Long, intitule As String, formule As String, unite As String)
    ws.Range("B" & r).Value = intitule
    ws.Range("B" & r & ":C" & r).Merge
    ws.Range("D" & r).Formula = formule
    ws.Range("E" & r).Value = unite
End Sub

' Sous-titre en gras, souligné, en colonne B
Private Sub SousTitre(ws As Worksheet, r As Long, texte As String)
    With ws.Range("B" & r)
        .Value = texte
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Size = 12
    End With
End Sub

' Valeur de la dernière couche (rang = nombre de couches) dans la colonne col du
' tableau des couches de la feuille « Données du projet »
Private Function ValeurCouche(col As String, lc As Long, lg As Long, nb As String) As String
    ValeurCouche = "=INDEX(" & DONNEES & "$" & col & "$" & lc & ":$" & col & "$" & lg & _
        "," & nb & ")&"""""
    If col >= "E" And col <= "H" Then        ' valeurs numériques
        ValeurCouche = "=IFERROR(INDEX(" & DONNEES & "$" & col & "$" & lc & ":$" & col & "$" & _
            lg & "," & nb & ")*1,"""")"
    End If
End Function

' Rang (1 à 6) de la catégorie de sol dans la liste CATEGORIES_SOL (feuille
' « Données du projet ») = colonne du tableau qs max
Private Function Rang6(sol As String) As String
    Rang6 = "MATCH(" & sol & ",{""" & Replace(CATEGORIES_SOL, ",", """,""") & """},0)"
End Function

' Colonne (1 à 5) des tableaux pressiométriques F.5.2.1 et F.5.2.2 : argiles et sols
' intermédiaires -> 1, sables et graves -> 2, craies -> 3, marnes -> 4, roches -> 5
Private Function Rang5(sol As String) As String
    Rang5 = "CHOOSE(" & Rang6(sol) & ",1,1,2,3,4,5)"
End Function

Private Function FormuleABC(zone As String, rangee As Long, sol As String) As String
    FormuleABC = "=IFERROR(INDEX(" & zone & "," & rangee & "," & Rang5(sol) & "),"""")"
End Function

Private Function FormuleAlpha(zone As String, cat As String, sol As String) As String
    FormuleAlpha = "=IFERROR(INDEX(" & zone & "," & cat & "," & Rang5(sol) & "),"""")"
End Function

Private Function FormuleQsMax(zone As String, cat As String, sol As String) As String
    FormuleQsMax = "=IFERROR(INDEX(" & zone & "," & cat & "," & Rang6(sol) & "),"""")"
End Function

' fsol en kPa (pl* en MPa, a, b, c du tableau F.5.2.2 donnant fsol en MPa)
Private Function FormuleFsol(pl As String, a As String, b As String, c As String) As String
    FormuleFsol = "=IF(OR(" & pl & "="""",NOT(ISNUMBER(" & a & "))),""""," & _
        "1000*(" & a & "*" & pl & "+" & b & ")*(1-EXP(-" & c & "*" & pl & ")))"
End Function

' alpha x fsol ; « - » si alpha n'est pas donné par la norme (catégories 17 et 18 :
' prendre la technique la plus proche)
Private Function FormuleProduit(alpha As String, fsol As String) As String
    FormuleProduit = "=IF(OR(" & fsol & "="""",ISBLANK(" & alpha & ")),""""," & _
        "IF(ISNUMBER(" & alpha & ")," & alpha & "*" & fsol & ",""-""))"
End Function

Private Function FormuleQs(produit As String, qsmax As String) As String
    FormuleQs = "=IF(AND(ISNUMBER(" & produit & "),ISNUMBER(" & qsmax & ")),MIN(" & _
        produit & "," & qsmax & "),IF(" & produit & "="""","""",""-""))"
End Function

' Ligne du tableau par couche (couche n° i, ligne ld dans « Données du projet ») ;
' vide au-delà du nombre de couches
Private Sub CoucheLigne(ws As Worksheet, r As Long, i As Long, ld As Long, nb As String, _
        cat As String, zAlpha As String, zABC As String, zQs As String)
    Dim cond As String, sol As String, pl As String
    cond = "=IF(" & i & ">" & nb & ","""","
    sol = "$D$" & r
    pl = "$F$" & r
    ws.Range("B" & r).Formula = cond & i & ")"
    ws.Range("C" & r).Formula = cond & DONNEES & "$C$" & ld & "&"""")"
    ws.Range("D" & r).Formula = cond & DONNEES & "$D$" & ld & "&"""")"
    ws.Range("E" & r).Formula = cond & "IF(ISNUMBER(" & DONNEES & "$F$" & ld & ")," & _
        DONNEES & "$F$" & ld & ",""""))"
    ws.Range("F" & r).Formula = cond & "IF(ISNUMBER(" & DONNEES & "$G$" & ld & ")," & _
        DONNEES & "$G$" & ld & ",""""))"
    ws.Range("G" & r).Formula = cond & DONNEES & "$I$" & ld & "&"""")"
    ws.Range("H" & r).Formula = Replace(FormuleABC(zABC, 1, sol), "=", cond, 1, 1) & ")"
    ws.Range("I" & r).Formula = Replace(FormuleAlpha(zAlpha, cat, sol), "=", cond, 1, 1) & ")"
    ' fsol avec a, b, c lus directement dans le tableau F.5.2.2
    ws.Range("J" & r).Formula = "=IF(OR(B" & r & "="""",F" & r & "="""",H" & r & "=""""),""""," & _
        "1000*(INDEX(" & zABC & ",2," & Rang5(sol) & ")*" & pl & "+INDEX(" & zABC & ",3," & _
        Rang5(sol) & "))*(1-EXP(-INDEX(" & zABC & ",4," & Rang5(sol) & ")*" & pl & ")))"
    ws.Range("K" & r).Formula = FormuleProduit("$I$" & r, "$J$" & r)
    ws.Range("L" & r).Formula = Replace(FormuleQsMax(zQs, cat, sol), "=", cond, 1, 1) & ")"
    ws.Range("M" & r).Formula = FormuleQs("$K$" & r, "$L$" & r)
    ws.Range("E" & r & ":F" & r).NumberFormat = "0.00"
    ws.Range("I" & r).NumberFormat = "0.00"
    ws.Range("J" & r & ":M" & r).NumberFormat = "0.0"
End Sub

' Adresse absolue (avec le nom de la feuille) des valeurs d'un tableau de la feuille
' « Tableaux NF P94-262 » : le titre commençant par titre est cherché, les valeurs
' commencent 2 lignes plus bas (sous l'en-tête) et 1 colonne à droite du titre
' (après la colonne des intitulés de ligne) ; nbLignes x nbCols valeurs.
Private Function ZoneTableau(wt As Worksheet, titre As String, decalage As Long, _
        nbLignes As Long, nbCols As Long) As String
    Dim c As Range
    Set c = wt.Cells.Find(What:=titre, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If c Is Nothing Then Err.Raise vbObjectError + 1, , "« " & titre & " » introuvable " & _
        "dans la feuille « " & wt.Name & " » : relancez CreerTableauxNFP94262."
    ZoneTableau = "'" & wt.Name & "'!" & c.Offset(2 + decalage, 1).Resize(nbLignes, nbCols).Address
End Function
