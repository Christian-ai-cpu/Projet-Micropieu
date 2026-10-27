Attribute VB_Name = "CalculFondation"
Option Explicit

' Module « Calcul et vérification » : crée (ou recrée) la feuille « Calcul et
' vérification » de la fondation, à partir des feuilles « Données du projet »
' (module DonneesProjet) et « Tableaux NF P94-262 » (module TableauxNFP94262).
' Toutes les valeurs sont des formules : elles suivent les saisies de ces feuilles.
'   B3  : Données du projet (type de pieu, méthode de calcul)
'   B7  : Caractéristiques de la fondation, en 3 groupes séparés par 2 lignes vides :
'         abréviation, norme, classe, catégorie ; mode de fonctionnement (compression
'         / traction), résistance de pointe, type de mise en oeuvre ; caractéristiques
'         mécaniques (B, D, hauteur de la couche porteuse, Ab, Pe, frottement négligé)
'   B26 : Coefficients de modèle gamma R;d1 et R;d2 (tableaux F.2.1 et G.2.1 recopiés
'         sur la feuille et utilisés dans le calcul de qs,i, selon la méthode de calcul)
'   B37 : Frottement axial unitaire et résistance par couche (calculés si le nom et la
'         catégorie de la couche sont renseignés), notations de la norme : profondeur,
'         hi, pl*,i, courbe et a, b, c (F.5.2.2), alpha pieu-sol (F.5.2.1), fsol,i,
'         qs,max (F.5.2.3), qs = min(alpha x fsol,i ; qs,max), As,i = Pe x hi,
'         qs,i = qs / (gamma R;d1 x gamma R;d2), Rs,i = qs,i x As,i (0 si le terrain
'         est neutralisé)
'   S3  : à droite, détermination du sol d'ancrage (dernière couche renseignée), de la
'         pression limite nette équivalente ple* = somme(pl*,i x dz) / (b + 3a) sur
'         [D-b ; D+3a] et de la hauteur d'encastrement équivalente Def = somme(pl*,i x
'         dz) / ple* sur [D-10B ; D], avec a = max(B/2 ; 0,5 m), b = min(a ; h),
'         D = somme des épaisseurs, h = épaisseur de la dernière couche
' Méthode pressiométrique (annexe F) : fsol = (a.pl* + b)(1 - exp(-c.pl*)), pl* en
' MPa, fsol en kPa ; qs = min(alpha pieu-sol x fsol ; qs max).
' Catégories de sol de la feuille « Données du projet » -> colonnes des tableaux :
'   tableaux à 5 colonnes (F.5.2.1, F.5.2.2) : argiles et sols intermédiaires -> 1re
'   colonne, sables et graves -> 2e, craies -> 3e, marnes -> 4e, roches -> 5e ;
'   tableau qs max à 6 colonnes : dans l'ordre de la liste des catégories de sol.

Private Const NOM_FEUILLE As String = "Calcul et vérification"
Private Const DONNEES As String = "'Données du projet'!"
Private Const TABLEAUX As String = "Tableaux NF P94-262"

' Thème vert : saisies en vert, valeurs non saisies (calculées) en jaune, en-têtes et
' titres en vert (couleurs RGB écrites en nombre : rouge + 256 x vert + 65536 x bleu)
Private Const VERT_SAISIE As Long = 13561798   ' RGB(198, 239, 206)
Private Const JAUNE_CALCUL As Long = 13434879  ' RGB(255, 255, 204), jaune clair
Private Const VERT_ENTETE As Long = 14348258   ' RGB(226, 239, 218)
Private Const VERT_FONCE As Long = 2315831     ' RGB(55, 86, 35)
Public Sub CreerCalculVerification()
    Dim ws As Worksheet, wt As Worksheet, r As Long, i As Long, lc As Long
    Dim zAlpha As String, zABC As String, zQs As String, nb As String, cat As String
    Dim lg As Long, etape As String, zF21 As String, ligneF21 As String, lt As Long, gamma As String

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
        .Value = "CALCUL ET VÉRIFICATION DES FONDATIONS"
        .Font.Bold = True
        .Font.Color = VERT_FONCE
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With

    ' --- Données du projet ---
    etape = "données du projet"
    SousTitre ws, 3, "Données du projet"
    Ligne ws, 4, "Type de pieu", "=" & DONNEES & "D" & LIGNE_PIEU, ""
    Ligne ws, 5, "Méthode de calcul", "=" & DONNEES & "D" & LIGNE_COUCHES + MAX_COUCHES + 5, ""

    ' --- Caractéristiques de la fondation : 3 groupes séparés par 2 lignes vides ---
    etape = "caractéristiques de la fondation"
    SousTitre ws, 7, "Caractéristiques de la fondation"
    Ligne ws, 8, "Abréviation", "=" & DONNEES & "D" & LIGNE_PIEU + 7, ""
    Ligne ws, 9, "Norme de référence", "=" & DONNEES & "D" & LIGNE_PIEU + 4, ""
    Ligne ws, 10, "Classe de la fondation", "=" & DONNEES & "D" & LIGNE_PIEU + 5, ""
    Ligne ws, 11, "Catégorie de la fondation", "=" & DONNEES & "D" & LIGNE_PIEU + 6, ""
    cat = "$D$11"                            ' catégorie de la fondation (cette feuille)

    Ligne ws, 14, "Mode de fonctionnement", "=" & DONNEES & "D" & LIGNE_PIEU + 8, ""
    Ligne ws, 15, "Prise en compte de la résistance de pointe", "=" & DONNEES & "D" & LIGNE_PIEU + 9, ""
    Ligne ws, 16, "Type de mise en oeuvre", "=" & DONNEES & "D" & LIGNE_PIEU + 10, ""

    Ligne ws, 19, "Largeur de la fondation B", "=" & DONNEES & "D" & LIGNE_PIEU + 11, "m"
    ' D = somme des épaisseurs des couches du tableau du frottement axial (lignes 39 à 53)
    Ligne ws, 20, "Hauteur totale de la fondation D (somme des épaisseurs)", _
        "=IF(COUNT($F$39:$F$53)=0,"""",SUM($F$39:$F$53))", "m"
    ' h = épaisseur de la dernière couche renseignée du tableau du frottement axial
    Ligne ws, 21, "Hauteur de la couche porteuse h", Derniere("F"), "m"
    Ligne ws, 22, "Aire de la base de la fondation Ab = " & ChrW(960) & ".B² / 4", _
        "=IF(N(D19)=0,"""",PI()*D19^2/4)", "m²"
    Ligne ws, 23, "Périmètre de la fondation Pe = " & ChrW(960) & ".B", _
        "=IF(N(D19)=0,"""",PI()*D19)", "m"
    ' Épaisseur des couches dont le terrain est neutralisé (frottement négligé)
    Ligne ws, 24, "Frottement négligé (épaisseur des terrains neutralisés)", _
        "=IF(COUNTIF(OFFSET(" & DONNEES & "$I$" & lc & ",0,0," & nb & "),""Oui"")=0,""Non""," & _
        "SUMIF(OFFSET(" & DONNEES & "$I$" & lc & ",0,0," & nb & "),""Oui""," & _
        "OFFSET(" & DONNEES & "$F$" & lc & ",0,0," & nb & ")))", "m"
    ws.Range("D19:D24").NumberFormat = "0.000"

    ' --- Coefficients de modèle (NF P94-262, tableaux F.2.1 et G.2.1) ---
    etape = "coefficients de modèle"
    SousTitre ws, 26, "Coefficients de modèle (NF P94-262, tableaux F.2.1 et G.2.1)"
    For i = 0 To 1                           ' 0 : pressiomètre (F.2.1), 1 : pénétromètre (G.2.1)
        r = 27 + 5 * i
        ws.Range("B" & r & ":F" & r).Value = Array(IIf(i = 0, "Méthode pressiométrique (F.2.1)", _
            "Méthode pénétrométrique (G.2.1)"), "", ChrW(947) & "R;d1 compression", _
            ChrW(947) & "R;d1 traction", ChrW(947) & "R;d2")
        ws.Range("B" & r + 1 & ":F" & r + 1).Value = Array("Classes 1 à 7 (hors catégories 10 et 15)", _
            "", IIf(i = 0, 1.15, 1.18), IIf(i = 0, 1.4, 1.45), 1.1)
        ws.Range("B" & r + 2 & ":F" & r + 2).Value = Array("Classes 1 à 7 ancrées dans la craie", _
            "", IIf(i = 0, 1.4, 1.45), IIf(i = 0, 1.7, 1.75), 1.1)
        ws.Range("B" & r + 3 & ":F" & r + 3).Value = Array("Catégories 10, 15, 17 à 20", "", 2, 2, 1.1)
        For lt = r To r + 3
            ws.Range("B" & lt & ":C" & lt).Merge
        Next lt
    Next i
    ws.Range("D28:F30,D33:F35").NumberFormat = "0.00"
    ' gamma R;d1 et R;d2 lus dans ces tableaux : tableau selon la méthode de calcul ;
    ' ligne 3 pour les catégories 10, 15 et 17 à 20, ligne 2 si la dernière couche est
    ' de la craie, ligne 1 sinon ; colonne compression / traction selon le mode
    zF21 = "IF($D$5=""Méthode pénétrométrique"",$D$33:$F$35,$D$28:$F$30)"
    ligneF21 = "IF(OR(" & cat & "=10," & cat & "=15," & cat & ">=17),3,IF($U$5=""Craies"",2,1))"
    gamma = "INDEX(" & zF21 & "," & ligneF21 & ",IF($D$14=""Traction"",2,1))*INDEX(" & zF21 & _
        "," & ligneF21 & ",3)"

    ' --- Frottement axial unitaire par couche (notations de la NF P94-262) ---
    etape = "frottement par couche"
    r = 37
    SousTitre ws, r, "Frottement axial unitaire et résistance par couche (méthode pressiométrique)"
    ws.Range("B" & r + 1 & ":Q" & r + 1).Value = Array("N°", "Nom de couche", "Catégorie de sol", _
        "Profondeur (m/TN)", "hi (m)", "pl*,i (MPa)", "Courbe fsol", ChrW(945) & "pieu-sol", _
        "fsol,i (kPa)", ChrW(945) & "pieu-sol x fsol,i (kPa)", "qs,max (kPa)", "qs = min (kPa)", _
        "As,i (m²)", "qs,i (kPa)", "Rs,i (kN)", "Terrain neutralisé")
    For i = 1 To MAX_COUCHES
        CoucheLigne ws, r + 1 + i, i, lc + i - 1, cat, gamma, zAlpha, zABC, zQs
    Next i
    lt = r + 1 + MAX_COUCHES                 ' dernière ligne du tableau

    ' --- À droite (colonnes S à X) : sol d'ancrage, pression limite nette équivalente
    ' ple* et hauteur d'encastrement équivalente Def (NF P94-262, annexe F), d'après les
    ' couches du tableau du frottement axial (lignes 39 à 53, dernière couche = ancrage)
    etape = "ancrage"
    SousTitre ws, 3, "Détermination du sol d'ancrage", "S"
    LigneD ws, 4, "Couche d'ancrage (dernière couche renseignée)", Derniere("C"), ""
    LigneD ws, 5, "Catégorie de sol", Derniere("D"), ""
    LigneD ws, 6, "Pression limite pl*", Derniere("G"), "MPa"
    LigneD ws, 7, "Hauteur de la couche porteuse h", "=$D$21", "m"
    ' Couches traversées par les intervalles [D-b ; D+3a] (ple*) et [D-10B ; D] (Def) ;
    ' la dernière couche est supposée se prolonger sous la base de la fondation
    SousTitre ws, 9, "Détermination de la pression limite nette équivalente ple*", "S"
    SousTitre ws, 28, "Détermination de la hauteur d'encastrement équivalente Def", "S"
    For i = 0 To 1
        r = IIf(i = 0, 10, 29)               ' ligne d'en-tête
        ws.Range("S" & r & ":X" & r).Value = Array("N°", "Couche", "z haut (m)", "z bas (m)", _
            IIf(i = 0, ChrW(916) & "z dans [D-b ; D+3a] (m)", ChrW(916) & "z dans [D-10B ; D] (m)"), _
            IIf(i = 0, "pl*,i x " & ChrW(916) & "z (MPa.m)", "pl*,i x " & ChrW(916) & "z / ple* (m)"))
        For lt = 1 To MAX_COUCHES
            AncrageLigne ws, r + lt, 38 + lt, i = 0
        Next lt
        lt = r + MAX_COUCHES + 1             ' ligne du résultat
        ws.Range("S" & lt).Value = IIf(i = 0, "ple* = " & ChrW(931) & " (pl*,i x " & ChrW(916) & _
            "z) / (b + 3a)  (MPa)", "Def = " & ChrW(931) & " (pl*,i x " & ChrW(916) & "z) / ple*  (m)")
        ws.Range("S" & lt & ":W" & lt).Merge
        ws.Range("X" & lt).Formula = IIf(i = 0, "=IF(OR(COUNT(X11:X25)=0,N(U52)=0),"""",SUM(X11:X25)/U52)", _
            "=IF(COUNT(X30:X44)=0,"""",SUM(X30:X44))")
        ws.Range("S" & lt & ":X" & lt).Font.Bold = True
    Next i
    ws.Range("U11:X26,U30:X45").NumberFormat = "0.00"

    ' Paramètres des intervalles (en tête : nom de la couche d'ancrage)
    LigneD ws, 47, "Paramètres de calcul de ple*", "=$U$4", ""
    LigneD ws, 48, "a = max(B/2 ; 0,5 m)", "=IF(N($D$19)=0,"""",MAX($D$19/2,0.5))", "m"
    LigneD ws, 49, "b = min(a ; h)", "=IF(OR(U48="""",N($D$21)=0),"""",MIN(U48,$D$21))", "m"
    LigneD ws, 50, "D - b", "=IF(OR(U49="""",$D$20=""""),"""",$D$20-U49)", "m"
    LigneD ws, 51, "D + 3a", "=IF(OR(U48="""",$D$20=""""),"""",$D$20+3*U48)", "m"
    LigneD ws, 52, "b + 3a", "=IF(U49="""","""",U49+3*U48)", "m"
    LigneD ws, 54, "Paramètres de calcul de Def", "=$U$4", ""
    LigneD ws, 55, "10 B", "=IF(N($D$19)=0,"""",10*$D$19)", "m"
    LigneD ws, 56, "D - 10 B", "=IF(OR(U55="""",$D$20=""""),"""",$D$20-U55)", "m"
    ws.Range("U48:U52,U55:U56").NumberFormat = "0.00"

    ' --- Mise en forme ---
    etape = "mise en forme"
    r = 37
    lt = r + 1 + MAX_COUCHES
    With ws.Range("B4:E5,B8:E11,B14:E16,B19:E24,B27:F30,B32:F35,B" & r + 1 & ":Q" & lt & _
            ",S4:V7,S10:X26,S29:X45,S47:V52,S54:V56")
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("B27:F27,B32:F32,B" & r + 1 & ":Q" & r + 1 & ",S10:X10,S29:X29,S47:V47,S54:V54")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = VERT_ENTETE
    End With
    ' Valeurs non saisies (reprises ou calculées) en jaune ; tableau du frottement axial
    ' sans couleur
    ws.Range("D4:D5,D8:D11,D14:D16,D19:D24,D28:F30,D33:F35,U4:U7,U48:U52,U55:U56,X26,X45").Interior.Color = _
        JAUNE_CALCUL
    ws.Range("D4:F35,B" & r + 2 & ":B" & lt & ",E" & r + 2 & ":Q" & lt & _
        ",U4:V56,S11:S25,U11:X25,S30:S44,U30:X44").HorizontalAlignment = xlCenter
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:Q").ColumnWidth = 12
    ws.Columns("R").ColumnWidth = 3
    ws.Columns("S:X").ColumnWidth = 16
    ws.Activate

Fin:
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    MsgBox "Erreur pendant la création de la feuille « Calcul et vérification », partie « " & _
        etape & " » : " & Err.Description & " (n° " & Err.Number & ")", vbExclamation
    Resume Fin
End Sub

' Valeur de la colonne col pour la dernière couche renseignée (nom en colonne C) du
' tableau du frottement axial (lignes 39 à 53)
Private Function Derniere(col As String) As String
    Derniere = "=IFERROR(LOOKUP(2,1/($C$39:$C$53<>""""),$" & col & "$39:$" & col & "$53),"""")"
End Function

' Une ligne « intitulé (S:T fusionnées) | valeur (U) | unité (V) » à droite
Private Sub LigneD(ws As Worksheet, r As Long, intitule As String, formule As String, unite As String)
    ws.Range("S" & r).Value = intitule
    ws.Range("S" & r & ":T" & r).Merge
    ws.Range("U" & r).Formula = formule
    ws.Range("V" & r).Value = unite
End Sub

' Ligne r d'un tableau ple* (ple = True) ou Def pour la couche de la ligne fr du
' tableau du frottement axial : z haut et z bas cumulés depuis le TN, épaisseur
' comprise dans l'intervalle [D-b ; D+3a] (ple*) ou [D-10B ; D] (Def)
Private Sub AncrageLigne(ws As Worksheet, r As Long, fr As Long, ple As Boolean)
    Dim cond As String, zbas As String
    cond = "=IF($C$" & fr & "="""","""","
    ws.Range("S" & r).Formula = cond & "$B$" & fr & ")"
    ws.Range("T" & r).Formula = cond & "$C$" & fr & ")"
    ws.Range("V" & r).Formula = cond & "SUM($F$39:$F$" & fr & "))"
    ws.Range("U" & r).Formula = cond & "V" & r & "-N($F$" & fr & "))"
    If ple Then
        ' dernière couche : prolongée sous la base (z bas très grand)
        If fr = 53 Then zbas = "1E+9" Else zbas = "IF(COUNTIF($C$" & fr + 1 & ":$C$53,""?*"")=0,1E+9,V" & r & ")"
        ws.Range("W" & r).Formula = cond & "IFERROR(MAX(0,MIN(" & zbas & ",$U$51)-MAX(U" & r & ",$U$50)),""""))"
        ws.Range("X" & r).Formula = "=IF(OR(W" & r & "="""",NOT(ISNUMBER($G$" & fr & "))),""""," & _
            "$G$" & fr & "*W" & r & ")"
    Else
        ws.Range("W" & r).Formula = cond & "IFERROR(MAX(0,MIN(V" & r & ",$D$20)-MAX(U" & r & ",$U$56)),""""))"
        ws.Range("X" & r).Formula = "=IF(OR(W" & r & "="""",NOT(ISNUMBER($G$" & fr & ")),N($X$26)=0)," & _
            """"",$G$" & fr & "*W" & r & "/$X$26)"
    End If
End Sub

' Une ligne « intitulé (B:C fusionnées) | valeur ou formule (D) | unité (E) »
Private Sub Ligne(ws As Worksheet, r As Long, intitule As String, formule As String, unite As String)
    ws.Range("B" & r).Value = intitule
    ws.Range("B" & r & ":C" & r).Merge
    ws.Range("D" & r).Formula = formule
    ws.Range("E" & r).Value = unite
End Sub

' Sous-titre en gras, souligné, en colonne B
Private Sub SousTitre(ws As Worksheet, r As Long, texte As String, Optional col As String = "B")
    With ws.Range(col & r)
        .Value = texte
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Size = 12
        .Font.Color = VERT_FONCE
    End With
End Sub

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

' Valeur d'un tableau de la norme (zone) en ligne lig et colonne col ; vide si erreur
Private Function FormuleIndex(zone As String, lig As String, col As String) As String
    FormuleIndex = "=IFERROR(INDEX(" & zone & "," & lig & "," & col & "),"""")"
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

' Ligne du tableau par couche (couche n° i, ligne ld dans « Données du projet ») :
' calculée seulement si le nom et la catégorie de la couche sont renseignés.
' Pe en D23 ; gamma = gamma R;d1 x gamma R;d2 (formule lue dans les tableaux des
' coefficients de modèle) ; terrain neutralisé : frottement négligé, Rs,i = 0.
Private Sub CoucheLigne(ws As Worksheet, r As Long, i As Long, ld As Long, cat As String, _
        gamma As String, zAlpha As String, zABC As String, zQs As String)
    Dim cond As String, sol As String, pl As String
    cond = "=IF(OR(" & DONNEES & "$C$" & ld & "=""""," & DONNEES & "$D$" & ld & "=""""),"""","
    sol = "$D$" & r
    pl = "$G$" & r
    ws.Range("B" & r).Formula = cond & i & ")"
    ws.Range("C" & r).Formula = cond & DONNEES & "$C$" & ld & ")"
    ws.Range("D" & r).Formula = cond & DONNEES & "$D$" & ld & ")"
    ws.Range("E" & r).Formula = cond & "IF(ISNUMBER(" & DONNEES & "$E$" & ld & ")," & _
        DONNEES & "$E$" & ld & ",""""))"
    ws.Range("F" & r).Formula = cond & "IF(ISNUMBER(" & DONNEES & "$F$" & ld & ")," & _
        DONNEES & "$F$" & ld & ",""""))"
    ws.Range("G" & r).Formula = cond & "IF(ISNUMBER(" & DONNEES & "$G$" & ld & ")," & _
        DONNEES & "$G$" & ld & ",""""))"
    ws.Range("H" & r).Formula = Replace(FormuleIndex(zABC, "1", Rang5(sol)), "=", cond, 1, 1) & ")"
    ws.Range("I" & r).Formula = Replace(FormuleIndex(zAlpha, cat, Rang5(sol)), "=", cond, 1, 1) & ")"
    ' fsol,i avec a, b, c lus directement dans le tableau F.5.2.2
    ws.Range("J" & r).Formula = "=IF(OR(B" & r & "="""",G" & r & "="""",H" & r & "=""""),""""," & _
        "1000*(INDEX(" & zABC & ",2," & Rang5(sol) & ")*" & pl & "+INDEX(" & zABC & ",3," & _
        Rang5(sol) & "))*(1-EXP(-INDEX(" & zABC & ",4," & Rang5(sol) & ")*" & pl & ")))"
    ws.Range("K" & r).Formula = FormuleProduit("$I$" & r, "$J$" & r)
    ws.Range("L" & r).Formula = Replace(FormuleIndex(zQs, cat, Rang6(sol)), "=", cond, 1, 1) & ")"
    ws.Range("M" & r).Formula = FormuleQs("$K$" & r, "$L$" & r)
    ws.Range("N" & r).Formula = "=IF(OR(B" & r & "="""",F" & r & "="""",N($D$23)=0),""""," & _
        "$D$23*F" & r & ")"
    ws.Range("O" & r).Formula = "=IF(B" & r & "="""","""",IFERROR(M" & r & "/(" & gamma & "),""-""))"
    ws.Range("P" & r).Formula = "=IF(AND(ISNUMBER(O" & r & "),ISNUMBER(N" & r & "))," & _
        "IF(Q" & r & "=""Oui"",0,O" & r & "*N" & r & "),IF(B" & r & "="""","""",""-""))"
    ws.Range("Q" & r).Formula = cond & DONNEES & "$I$" & ld & ")"
    ws.Range("E" & r & ":G" & r & ",I" & r).NumberFormat = "0.00"
    ws.Range("J" & r & ":M" & r & ",O" & r & ":P" & r).NumberFormat = "0.0"
    ws.Range("N" & r).NumberFormat = "0.000"
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
