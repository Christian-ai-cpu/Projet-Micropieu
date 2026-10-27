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
'   B26 : Facteurs partiels de résistance ELU (C.2.3.1, C.2.3.2) et ELS (14.2.1.1,
'         14.2.1.2) ; les coefficients de modèle gamma R;d1 et R;d2 du calcul de qs,i
'         sont lus dans la feuille « Tableaux NF P94-262 » (F.2.1 ou G.2.1)
'   B37 : Frottement axial unitaire et résistance par couche (calculés si le nom et la
'         catégorie de la couche sont renseignés), notations de la norme : profondeur,
'         hi, pl*,i, courbe et a, b, c (F.5.2.2), alpha pieu-sol (F.5.2.1), fsol,i,
'         qs,max (F.5.2.3), qs = min(alpha x fsol,i ; qs,max), As,i = Pe x hi,
'         qs,i = qs / (gamma R;d1 x gamma R;d2), Rs,i = qs,i x As,i (0 si le terrain
'         est neutralisé)
'   B56 : valeurs caractéristiques (Rb,k, Rs,k, Rc,k, Rt,k, Rc;cr,k, Rt;cr,k) puis
'         valeurs de calcul ELU fondamental / accidentel et ELS caractéristique /
'         quasi permanent (compression et traction)
'   U3  : à droite, détermination du sol d'ancrage (dernière couche renseignée), de la
'         pression limite nette équivalente ple* = somme(pl*,i x dz) / (b + 3a) sur
'         [D-b ; D+3a] et de la hauteur d'encastrement équivalente Def = somme(pl*,i x
'         dz) / ple* sur [D-10B ; D], avec a = max(B/2 ; 0,5 m), b = min(a ; h),
'         D = somme des épaisseurs, h = épaisseur de la dernière couche
'         puis kp (kp max du tableau F.4.2.1 et Def/B) et Rb = Ab x kp x ple* ;
'         colonnes masquées pour les micropieux (catégories 17 à 20)
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
    Dim zKp As String, c As String

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
    zKp = ZoneTableau(wt, "Tableau F.4.2.1", 0, 8, 5)       ' 8 classes x 5 sols

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

    ' --- Facteurs partiels de résistance (NF P94-262) : ELU (tableaux C.2.3.1 et
    ' C.2.3.2) et ELS (tableaux 14.2.1.1 et 14.2.1.2), utilisés pour les valeurs de calcul
    etape = "facteurs partiels"
    SousTitre ws, 26, "Facteurs partiels de résistance (NF P94-262)"
    ws.Range("B27:E27").Value = Array("ELU (tableaux C.2.3.1 et C.2.3.2)", "", _
        "Situations durables et transitoires", "Situations accidentelles")
    ws.Range("B28:E28").Value = Array("Pointe " & ChrW(947) & "b", "", 1.1, 1)
    ws.Range("B29:E29").Value = Array("Fût en compression " & ChrW(947) & "s", "", 1.1, 1)
    ws.Range("B30:E30").Value = Array("Total en compression " & ChrW(947) & "t", "", 1.1, 1)
    ws.Range("B31:E31").Value = Array("Fût en traction " & ChrW(947) & "s;t", "", 1.15, 1.05)
    ws.Range("B33:E33").Value = Array("ELS (tableaux 14.2.1.1 et 14.2.1.2)", "", _
        "Combinaisons caractéristiques", "Combinaisons quasi permanentes")
    ws.Range("B34:E34").Value = Array("Fût en compression " & ChrW(947) & "cr", "", 0.9, 1.1)
    ws.Range("B35:E35").Value = Array("Fût en traction " & ChrW(947) & "s;cr", "", 1.1, 1.5)
    For i = 27 To 35
        If i <> 32 Then ws.Range("B" & i & ":C" & i).Merge
    Next i
    ws.Range("D28:E31,D34:E35").NumberFormat = "0.00"
    ' Coefficients de modèle gamma R;d1 et R;d2 lus dans la feuille « Tableaux NF P94-262 »
    ' (tableau F.2.1 ou G.2.1 selon la méthode de calcul) : ligne 3 pour les catégories
    ' 10, 15 et 17 à 20, ligne 2 si le sol d'ancrage est de la craie, ligne 1 sinon ;
    ' colonnes compression / traction selon le mode de fonctionnement
    zF21 = "IF($D$5=""Méthode pénétrométrique""," & ZoneTableau(wt, "Tableau G.2.1", 0, 3, 4) & "," & _
        ZoneTableau(wt, "Tableau F.2.1", 0, 3, 4) & ")"
    ligneF21 = "IF(OR(" & cat & "=10," & cat & "=15," & cat & ">=17),3,IF($W$5=""Craies"",2,1))"
    gamma = "INDEX(" & zF21 & "," & ligneF21 & ",IF($D$14=""Traction"",2,1))*INDEX(" & zF21 & _
        "," & ligneF21 & ",IF($D$14=""Traction"",4,3))"

    ' --- Frottement axial unitaire par couche (notations de la NF P94-262) ---
    etape = "frottement par couche"
    r = 37
    SousTitre ws, r, "Frottement axial unitaire et résistance par couche"
    ws.Range("B" & r).Formula = "=""Frottement axial unitaire et résistance par couche (""&LOWER($D$5)&"")""
    ws.Range("B" & r + 1 & ":Q" & r + 1).Value = Array("N°", "Nom de couche", "Catégorie de sol", _
        "Profondeur (m/TN)", "hi (m)", "pl*,i (MPa)", "Courbe fsol", ChrW(945) & "pieu-sol", _
        "fsol,i (kPa)", ChrW(945) & "pieu-sol x fsol,i (kPa)", "qs,max (kPa)", "qs = min (kPa)", _
        "As,i (m²)", "qs,i (kPa)", "Rs,i (kN)", "Terrain neutralisé")
    For i = 1 To MAX_COUCHES
        CoucheLigne ws, r + 1 + i, i, lc + i - 1, cat, gamma, zAlpha, zABC, zQs
    Next i
    lt = r + 1 + MAX_COUCHES                 ' dernière ligne du tableau

    ' --- À droite (colonnes U à Z) : sol d'ancrage, pression limite nette équivalente
    ' ple*, hauteur d'encastrement équivalente Def, facteur de portance kp et résistance
    ' de pointe Rb (NF P94-262, annexe F), d'après les couches du tableau du frottement
    ' axial (lignes 39 à 53, dernière couche = ancrage). Colonnes masquées pour les
    ' micropieux (catégories 17 à 20, pointe non prise en compte) : voir AfficherPointe.
    etape = "ancrage"
    SousTitre ws, 3, "Détermination du sol d'ancrage", "U"
    LigneD ws, 4, "Couche d'ancrage (dernière couche renseignée)", Derniere("C"), ""
    LigneD ws, 5, "Catégorie de sol", Derniere("D"), ""
    LigneD ws, 6, "Pression limite pl*", Derniere("G"), "MPa"
    LigneD ws, 7, "Hauteur de la couche porteuse h", "=$D$21", "m"
    ' Couches traversées par les intervalles [D-b ; D+3a] (ple*) et [D-10B ; D] (Def) ;
    ' la dernière couche est supposée se prolonger sous la base de la fondation
    SousTitre ws, 9, "Détermination de la pression limite nette équivalente ple*", "U"
    SousTitre ws, 28, "Détermination de la hauteur d'encastrement équivalente Def", "U"
    For i = 0 To 1
        r = IIf(i = 0, 10, 29)               ' ligne d'en-tête
        ws.Range("U" & r & ":Z" & r).Value = Array("N°", "Couche", "z haut (m)", "z bas (m)", _
            IIf(i = 0, ChrW(916) & "z dans [D-b ; D+3a] (m)", ChrW(916) & "z dans [D-10B ; D] (m)"), _
            IIf(i = 0, "pl*,i x " & ChrW(916) & "z (MPa.m)", "pl*,i x " & ChrW(916) & "z / ple* (m)"))
        For lt = 1 To MAX_COUCHES
            AncrageLigne ws, r + lt, 38 + lt, i = 0
        Next lt
        lt = r + MAX_COUCHES + 1             ' ligne du résultat
        ws.Range("U" & lt).Value = IIf(i = 0, "ple* = " & ChrW(931) & " (pl*,i x " & ChrW(916) & _
            "z) / (b + 3a)  (MPa)", "Def = " & ChrW(931) & " (pl*,i x " & ChrW(916) & "z) / ple*  (m)")
        ws.Range("U" & lt & ":Y" & lt).Merge
        ws.Range("Z" & lt).Formula = IIf(i = 0, "=IF(OR(COUNT(Z11:Z25)=0,N(W52)=0),"""",SUM(Z11:Z25)/W52)", _
            "=IF(COUNT(Z30:Z44)=0,"""",SUM(Z30:Z44))")
        ws.Range("U" & lt & ":Z" & lt).Font.Bold = True
    Next i
    ws.Range("W11:Z26,W30:Z45").NumberFormat = "0.00"

    ' Paramètres des intervalles (en tête : nom de la couche d'ancrage)
    LigneD ws, 47, "Paramètres de calcul de ple*", "=$W$4", ""
    LigneD ws, 48, "a = max(B/2 ; 0,5 m)", "=IF(N($D$19)=0,"""",MAX($D$19/2,0.5))", "m"
    LigneD ws, 49, "b = min(a ; h)", "=IF(OR(W48="""",N($D$21)=0),"""",MIN(W48,$D$21))", "m"
    LigneD ws, 50, "D - b", "=IF(OR(W49="""",$D$20=""""),"""",$D$20-W49)", "m"
    LigneD ws, 51, "D + 3a", "=IF(OR(W48="""",$D$20=""""),"""",$D$20+3*W48)", "m"
    LigneD ws, 52, "b + 3a", "=IF(W49="""","""",W49+3*W48)", "m"
    LigneD ws, 54, "Paramètres de calcul de Def", "=$W$4", ""
    LigneD ws, 55, "10 B", "=IF(N($D$19)=0,"""",10*$D$19)", "m"
    LigneD ws, 56, "D - 10 B", "=IF(OR(W55="""",$D$20=""""),"""",$D$20-W55)", "m"
    ws.Range("W48:W52,W55:W56").NumberFormat = "0.00"

    ' Facteur de portance kp (F.4.2) et résistance de pointe Rb = Ab x kp x ple*
    SousTitre ws, 58, "Détermination du facteur de portance kp et de la résistance de pointe Rb", "U"
    LigneD ws, 59, "Classe de la fondation", "=$D$10", ""
    LigneD ws, 60, "kp max (tableau F.4.2.1, classe et sol d'ancrage)", _
        "=IFERROR(INDEX(" & zKp & ",VALUE(LEFT($D$10,1))," & Rang5("$W$5") & "),"""")", ""
    LigneD ws, 61, "Def / B", "=IF(OR($Z$45="""",N($D$19)=0),"""",$Z$45/$D$19)", ""
    LigneD ws, 62, "kp = kp max si Def/B > 5, sinon 1 + (kp max - 1) x (Def/B) / 5", _
        "=IF(OR(NOT(ISNUMBER(W60)),W61=""""),"""",IF(W61>5,W60,1+(W60-1)*W61/5))", ""
    LigneD ws, 63, "Aire de la base Ab", "=$D$22", "m²"
    LigneD ws, 64, "Pression limite nette équivalente ple*", "=$Z$26", "MPa"
    LigneD ws, 65, "Rb = Ab x kp x ple*", "=IF(OR(W62="""",N(W63)=0,W64=""""),"""",1000*W63*W62*W64)", "kN"
    ws.Range("W60:W65").NumberFormat = "0.00"
    ws.Range("U65:X65").Font.Bold = True

    ' --- Sous le frottement axial : valeurs caractéristiques puis valeurs de calcul ---
    etape = "valeurs caractéristiques"
    SousTitre ws, 56, "Valeurs caractéristiques des résistances"
    ' Rb,k = Rb / (gamma R;d1 x gamma R;d2) si la résistance de pointe est prise en compte
    Ligne ws, 57, "Résistance de pointe Rb,k = Rb / (" & ChrW(947) & "R;d1 x " & ChrW(947) & "R;d2)", _
        "=IF($D$15<>""Oui"",0,IFERROR($W$65/(" & gamma & "),""""))", "kN"
    Ligne ws, 58, "Résistance de frottement Rs,k = " & ChrW(931) & " Rs,i", _
        "=IF(COUNT($P$39:$P$53)=0,"""",SUM($P$39:$P$53))", "kN"
    Ligne ws, 60, "Portance Rc,k = Rb,k + Rs,k (compression)", _
        "=IF(AND($D$14=""Compression"",ISNUMBER(D58)),N(D57)+D58,"""")", "kN"
    Ligne ws, 61, "Résistance en traction Rt,k = Rs,k", "=IF(ISNUMBER(D58),D58,"""")", "kN"
    Ligne ws, 62, "Charge de fluage en compression Rc;cr,k = 0,5 Rb,k + 0,7 Rs,k", _
        "=IF(AND($D$14=""Compression"",ISNUMBER(D58)),0.5*N(D57)+0.7*D58,"""")", "kN"
    Ligne ws, 63, "Charge de fluage en traction Rt;cr,k = 0,7 Rt,k", _
        "=IF(ISNUMBER(D61),0.7*D61,"""")", "kN"
    ws.Range("D57:D63").NumberFormat = "0.0"

    etape = "valeurs de calcul"
    SousTitre ws, 65, "Valeurs de calcul des résistances"
    ws.Range("B66:E66").Value = Array("Combinaison", "", "Compression (kN)", "Traction (kN)")
    ' ELU : Rc,d = Rb,k / gamma b + Rs,k / gamma s ; Rt,d = Rt,k / gamma s;t
    ' ELS : Rc;cr,d = Rc;cr,k / gamma cr ; Rt;cr,d = Rt;cr,k / gamma s;cr
    For i = 0 To 3
        r = 67 + i
        c = Chr(Asc("D") + (i Mod 2))        ' colonne du facteur : D (durable / caract.), E
        ws.Range("B" & r).Value = Array("ELU fondamental (C.2.3.1)", "ELU accidentel (C.2.3.2)", _
            "ELS caractéristique (14.2.1.1)", "ELS quasi permanent (14.2.1.2)")(i)
        ws.Range("B" & r & ":C" & r).Merge
        If i < 2 Then
            ws.Range("D" & r).Formula = "=IF(D60="""","""",N(D57)/" & c & "28+D58/" & c & "29)"
            ws.Range("E" & r).Formula = "=IF(D61="""","""",D61/" & c & "31)"
        Else
            ws.Range("D" & r).Formula = "=IF(D62="""","""",D62/" & c & "34)"
            ws.Range("E" & r).Formula = "=IF(D63="""","""",D63/" & c & "35)"
        End If
    Next i
    ws.Range("D67:E70").NumberFormat = "0.0"


    ' --- Mise en forme ---
    etape = "mise en forme"
    r = 37
    lt = r + 1 + MAX_COUCHES
    With ws.Range("B4:E5,B8:E11,B14:E16,B19:E24,B27:E31,B33:E35,B" & r + 1 & ":Q" & lt & _
            ",B57:E58,B60:E63,B66:E70,U4:X7,U10:Z26,U29:Z45,U47:X52,U54:X56,U59:X65")
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("B27:E27,B33:E33,B66:E66,B" & r + 1 & ":Q" & r + 1 & ",U10:Z10,U29:Z29,U47:X47,U54:X54")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = VERT_ENTETE
    End With
    ' Valeurs non saisies (reprises ou calculées) en jaune ; tableau du frottement axial
    ' sans couleur
    ws.Range("D4:D5,D8:D11,D14:D16,D19:D24,D28:E31,D34:E35,D57:D58,D60:D63,D67:E70,W4:W7,W48:W52," & _
        "W55:W56,Z26,Z45,W59:W65").Interior.Color = _
        JAUNE_CALCUL
    ws.Range("D4:F35,D57:E70,B" & r + 2 & ":B" & lt & ",E" & r + 2 & ":Q" & lt & _
        ",W4:X65,U11:U25,W11:Z25,U30:U44,W30:Z44").HorizontalAlignment = xlCenter
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:Q").ColumnWidth = 12
    ws.Columns("R:T").ColumnWidth = 3
    ws.Columns("U:Z").ColumnWidth = 16
    AfficherPointe ws
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

' Une ligne « intitulé (U:V fusionnées) | valeur (W) | unité (X) » à droite
Private Sub LigneD(ws As Worksheet, r As Long, intitule As String, formule As String, unite As String)
    ws.Range("U" & r).Value = intitule
    ws.Range("U" & r & ":V" & r).Merge
    ws.Range("W" & r).Formula = formule
    ws.Range("X" & r).Value = unite
End Sub

' Colonnes de droite (sol d'ancrage, ple*, Def, kp, Rb) masquées pour les micropieux
' (catégories 17 à 20 : résistance de pointe non prise en compte). Appelée à la
' création et quand le type de pieu change (ChangementDonnees, module DonneesProjet).
Public Sub AfficherPointe(Optional ws As Worksheet)
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
    End If
    ws.Columns("T:Z").Hidden = (Val(ws.Range("D11").Value) >= 17)
End Sub

' Ligne r d'un tableau ple* (ple = True) ou Def pour la couche de la ligne fr du
' tableau du frottement axial : z haut et z bas cumulés depuis le TN, épaisseur
' comprise dans l'intervalle [D-b ; D+3a] (ple*) ou [D-10B ; D] (Def)
Private Sub AncrageLigne(ws As Worksheet, r As Long, fr As Long, ple As Boolean)
    Dim cond As String, zbas As String
    cond = "=IF($C$" & fr & "="""","""","
    ws.Range("U" & r).Formula = cond & "$B$" & fr & ")"
    ws.Range("V" & r).Formula = cond & "$C$" & fr & ")"
    ws.Range("X" & r).Formula = cond & "SUM($F$39:$F$" & fr & "))"
    ws.Range("W" & r).Formula = cond & "X" & r & "-N($F$" & fr & "))"
    If ple Then
        ' dernière couche : prolongée sous la base (z bas très grand)
        If fr = 53 Then zbas = "1E+9" Else zbas = "IF(COUNTIF($C$" & fr + 1 & ":$C$53,""?*"")=0,1E+9,X" & r & ")"
        ws.Range("Y" & r).Formula = cond & "IFERROR(MAX(0,MIN(" & zbas & ",$W$51)-MAX(W" & r & ",$W$50)),""""))"
        ws.Range("Z" & r).Formula = "=IF(OR(Y" & r & "="""",NOT(ISNUMBER($G$" & fr & "))),""""," & _
            "$G$" & fr & "*Y" & r & ")"
    Else
        ws.Range("Y" & r).Formula = cond & "IFERROR(MAX(0,MIN(X" & r & ",$D$20)-MAX(W" & r & ",$W$56)),""""))"
        ws.Range("Z" & r).Formula = "=IF(OR(Y" & r & "="""",NOT(ISNUMBER($G$" & fr & ")),N($Z$26)=0)," & _
            """"",$G$" & fr & "*Y" & r & "/$Z$26)"
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
    ' qs,i = qs / (gamma R;d1 x gamma R;d2) ; 0 si le terrain est neutralisé
    ws.Range("O" & r).Formula = "=IF(B" & r & "="""","""",IF(Q" & r & "=""Oui"",0," & _
        "IFERROR(M" & r & "/(" & gamma & "),""-"")))"
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
