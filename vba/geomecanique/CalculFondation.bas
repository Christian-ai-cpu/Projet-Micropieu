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
'   AB3 : tout à droite, mêmes tableaux par la méthode pénétrométrique (annexe G) :
'         qc,i, qcm, écrêtage qcc = min(qc ; 1,3 qcm), qce, Def, kc (G.4.2.1) et
'         Rb = Ab x kc x qce
' Méthode pressiométrique (annexe F) : fsol = (a.pl* + b)(1 - exp(-c.pl*)), pl* en
' MPa, fsol en kPa ; qs = min(alpha pieu-sol x fsol ; qs max).
' Catégories de sol de la feuille « Données du projet » -> colonnes des tableaux :
'   tableaux à 5 colonnes (F.5.2.1, F.5.2.2) : argiles et sols intermédiaires -> 1re
'   colonne, sables et graves -> 2e, craies -> 3e, marnes -> 4e, roches -> 5e ;
'   tableau qs max à 6 colonnes : dans l'ordre de la liste des catégories de sol.

Private Const NOM_FEUILLE As String = "Calcul et vérification"
Private Const DONNEES As String = "'Données du projet'!"
Private Const TABLEAUX As String = "Tableaux NF P94-262"
Private mCellule As String     ' cellule en cours d'écriture (indiquée en cas d'erreur)
' Méthode de calcul lue dans « Données du projet » à la création de la feuille : les
' formules sont écrites pour cette seule méthode (pl* et tableaux F, ou qc et tableaux G)
' et la feuille est recréée quand la méthode change (ActualiserCalculVerification,
' appelée par ChangementDonnees du module DonneesProjet)
Private mPenetro As Boolean

' Thème vert : saisies en vert, valeurs non saisies (calculées) en jaune, en-têtes et
' titres en vert (couleurs RGB écrites en nombre : rouge + 256 x vert + 65536 x bleu)
Private Const VERT_SAISIE As Long = 13561798   ' RGB(198, 239, 206)
Private Const JAUNE_CALCUL As Long = 13434879  ' RGB(255, 255, 204), jaune clair
Private Const VERT_ENTETE As Long = 14348258   ' RGB(226, 239, 218)
Private Const VERT_FONCE As Long = 2315831     ' RGB(55, 86, 35)
Public Sub CreerCalculVerification()
    Construire True
End Sub

' Recrée la feuille « Calcul et vérification » (si elle existe) sans l'activer : appelée
' quand la méthode de calcul change dans « Données du projet »
Public Sub ActualiserCalculVerification()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo 0
    If Not ws Is Nothing Then Construire False
End Sub

Private Sub Construire(activer As Boolean)
    Dim ws As Worksheet, wt As Worksheet, r As Long, i As Long, lc As Long
    Dim zAlpha As String, zABC As String, zQs As String, nb As String, cat As String
    Dim lg As Long, etape As String, zF21 As String, ligneF21 As String, lt As Long, gamma As String
    Dim zKp As String, c As String, zKc As String

    Application.ScreenUpdating = False
    On Error GoTo Erreur
    mPenetro = (ActiveWorkbook.Worksheets("Données du projet").Range("D" & _
        LIGNE_COUCHES + MAX_COUCHES + 5).Value = "Méthode pénétrométrique")

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
    ' Frottement : tableaux G (méthode pénétrométrique) et F, sous la forme « plage G|plage F »
    ' (lue par IndexNorme)
    zAlpha = ZoneTableau(wt, "Tableau G.5.2.1", 0, 20, 6) & "|" & _
        ZoneTableau(wt, "Tableau F.5.2.1", 0, 20, 5)        ' 20 catégories
    zABC = ZoneTableau(wt, "Tableau G.5.2.2", 0, 4, 6) & "|" & _
        ZoneTableau(wt, "Tableau F.5.2.2", 0, 4, 5)         ' courbe, a, b, c
    zQs = ZoneTableau(wt, "Tableau G.5.2.3", 0, 20, 6) & "|" & _
        ZoneTableau(wt, "Tableau F.5.2.3", 0, 20, 6)        ' qs max, 20 catégories
    zKp = ZoneTableau(wt, "Tableau F.4.2.1", 0, 8, 5)       ' 8 classes x 5 sols
    zKc = ZoneTableau(wt, "Tableau G.4.2.1", 0, 8, 6)       ' 8 classes x 6 sols

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
        "=IF(COUNTIF(OFFSET(" & DONNEES & "$J$" & lc & ",0,0," & nb & "),""Oui"")=0,""Non""," & _
        "SUMIF(OFFSET(" & DONNEES & "$J$" & lc & ",0,0," & nb & "),""Oui""," & _
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
    zF21 = ZoneTableau(wt, IIf(mPenetro, "Tableau G.2.1", "Tableau F.2.1"), 0, 3, 4)
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
    ' En-têtes selon la méthode de calcul : pl* et tableaux F, ou qc et tableaux G
    If mPenetro Then
        ws.Range("G" & r + 1 & ":J" & r + 1).Value = Array("qc,i (MPa)", "Courbe fsol (G.5.2.2)", _
            ChrW(945) & "pieu-sol (G.5.2.1)", "fsol,i = (a.qc + b)(1 - e^-c.qc) (kPa)")
        ws.Range("L" & r + 1).Value = "qs,max (G.5.2.3) (kPa)"
    Else
        ws.Range("G" & r + 1 & ":J" & r + 1).Value = Array("pl*,i (MPa)", "Courbe fsol (F.5.2.2)", _
            ChrW(945) & "pieu-sol (F.5.2.1)", "fsol,i = (a.pl* + b)(1 - e^-c.pl*) (kPa)")
        ws.Range("L" & r + 1).Value = "qs,max (F.5.2.3) (kPa)"
    End If
    For i = 1 To MAX_COUCHES
        CoucheLigne ws, r + 1 + i, i, lc + i - 1, cat, gamma, zAlpha, zABC, zQs
    Next i
    mCellule = ""
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
    LigneD ws, 6, "Pression limite pl*", "=IFERROR(LOOKUP(2,1/($C$39:$C$53<>""""),OFFSET(" & DONNEES & _
        "$G$" & lc & ",0,0," & MAX_COUCHES & ")),"""")", "MPa"
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

    ' --- Tout à droite (colonnes AB à AH) : mêmes tableaux par la méthode pénétrométrique
    etape = "ancrage (méthode pénétrométrique)"
    AncragePenetro ws, zKc

    ' --- Sous le frottement axial : valeurs caractéristiques puis valeurs de calcul ---
    ' Intitulé en B, symbole en C, valeur en D, unité en E
    etape = "valeurs caractéristiques"
    SousTitre ws, 56, "Valeurs caractéristiques des résistances"
    ' Rb,k = Rb / (gamma R;d1 x gamma R;d2) si la résistance de pointe est prise en compte ;
    ' 0 pour les micropieux (catégories 17 à 20)
    ' (Rb pénétrométrique en AD65 si la méthode est pénétrométrique, pressiométrique en W65 sinon)
    LigneS ws, 57, "Résistance de pointe = Rb / (" & ChrW(947) & "R;d1 x " & ChrW(947) & "R;d2)", "Rb,k", _
        "=IF(OR($D$15<>""Oui"",N($D$11)>=17),0,IFERROR(" & IIf(mPenetro, "$AD$65", "$W$65") & "/(" & _
        gamma & "),""""))", "kN"
    LigneS ws, 58, "Résistance de frottement = " & ChrW(931) & " Rs,i", "Rs,k", _
        "=IF(COUNT($P$39:$P$53)=0,"""",SUM($P$39:$P$53))", "kN"
    LigneS ws, 60, "Portance (compression) = Rb,k + Rs,k", "Rc,k", _
        "=IF(AND($D$14=""Compression"",ISNUMBER(D58)),N(D57)+D58,"""")", "kN"
    LigneS ws, 61, "Résistance en traction = Rs,k", "Rt,k", "=IF(ISNUMBER(D58),D58,"""")", "kN"
    LigneS ws, 62, "Charge de fluage en compression = 0,5 Rb,k + 0,7 Rs,k", "Rc;cr,k", _
        "=IF(AND($D$14=""Compression"",ISNUMBER(D58)),0.5*N(D57)+0.7*D58,"""")", "kN"
    LigneS ws, 63, "Charge de fluage en traction = 0,7 Rt,k", "Rt;cr,k", _
        "=IF(ISNUMBER(D61),0.7*D61,"""")", "kN"
    ws.Range("D57:D63").NumberFormat = "0.0"

    ' Valeurs de calcul : 2 lignes par combinaison (compression puis traction), avec
    ' combinaison en B (fusionnée), intitulé en C, symbole en D, valeur en E (avec l'unité)
    '   ELU : Rc,d = Rb,k / gamma b + Rs,k / gamma s ; Rt,d = Rt,k / gamma s;t
    '   ELS : Rc;cr,d = Rc;cr,k / gamma cr ; Rt;cr,d = Rt;cr,k / gamma s;cr
    etape = "valeurs de calcul"
    SousTitre ws, 65, "Valeurs de calcul des résistances"
    ws.Range("B66:E66").Value = Array("Combinaison", "Valeur de calcul", "Symbole", "Valeur")
    For i = 0 To 3
        r = 67 + 2 * i
        c = Chr(Asc("D") + (i Mod 2))        ' colonne du facteur : D (durable / caract.), E
        ws.Range("B" & r).Value = Array("ELU fondamental (C.2.3.1)", "ELU accidentel (C.2.3.2)", _
            "ELS caractéristique (14.2.1.1)", "ELS quasi permanent (14.2.1.2)")(i)
        ws.Range("B" & r & ":B" & r + 1).Merge
        If i < 2 Then
            ws.Range("C" & r & ":D" & r + 1).Value = Array2("Portance de calcul en compression", _
                "Rc,d", "Résistance de calcul en traction", "Rt,d")
            ws.Range("E" & r).Formula = "=IF(D60="""","""",N(D57)/" & c & "28+D58/" & c & "29)"
            ws.Range("E" & r + 1).Formula = "=IF(D61="""","""",D61/" & c & "31)"
        Else
            ws.Range("C" & r & ":D" & r + 1).Value = Array2("Charge de fluage de calcul en compression", _
                "Rc;cr,d", "Charge de fluage de calcul en traction", "Rt;cr,d")
            ws.Range("E" & r).Formula = "=IF(D62="""","""",D62/" & c & "34)"
            ws.Range("E" & r + 1).Formula = "=IF(D63="""","""",D63/" & c & "35)"
        End If
    Next i
    ws.Range("E67:E74").NumberFormat = "0.0"" kN""


    ' --- Mise en forme ---
    etape = "mise en forme"
    r = 37
    lt = r + 1 + MAX_COUCHES
    With ws.Range("B4:E5,B8:E11,B14:E16,B19:E24,B27:E31,B33:E35,B" & r + 1 & ":Q" & lt & _
            ",B57:E58,B60:E63,B66:E74,U4:X7,U10:Z26,U29:Z45,U47:X52,U54:X56,U59:X65")
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
    ws.Range("D4:D5,D8:D11,D14:D16,D19:D24,D28:E31,D34:E35,D57:D58,D60:D63,E67:E74,W4:W7,W48:W52," & _
        "W55:W56,Z26,Z45,W59:W65").Interior.Color = _
        JAUNE_CALCUL
    ws.Range("D4:F35,C57:E63,B67:B74,D67:E74,B" & r + 2 & ":B" & lt & ",E" & r + 2 & ":Q" & lt & _
        ",W4:X65,U11:U25,W11:Z25,U30:U44,W30:Z44").HorizontalAlignment = xlCenter
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:Q").ColumnWidth = 12
    ws.Columns("R:T").ColumnWidth = 3
    ws.Columns("U:Z").ColumnWidth = 16
    AfficherPointe ws
    If activer Then ws.Activate

Fin:
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    MsgBox "Erreur pendant la création de la feuille « Calcul et vérification », partie « " & _
        etape & " »" & IIf(mCellule = "", "", ", cellule " & mCellule) & " : " & Err.Description & _
        " (n° " & Err.Number & ")", vbExclamation
    Resume Fin
End Sub

' Valeur de la colonne col pour la dernière couche renseignée (nom en colonne C) du
' tableau du frottement axial (lignes 39 à 53)
Private Function Derniere(col As String) As String
    Derniere = "=IFERROR(LOOKUP(2,1/($C$39:$C$53<>""""),$" & col & "$39:$" & col & "$53),"""")"
End Function

' Une ligne « intitulé (U:V fusionnées) | valeur (W) | unité (X) » à droite
' (col = première colonne : U, ou AB pour la méthode pénétrométrique)
Private Sub LigneD(ws As Worksheet, r As Long, intitule As String, formule As String, unite As String, _
        Optional col As String = "U")
    With ws.Range(col & r)
        .Value = intitule
        .Resize(1, 2).Merge
        .Offset(0, 2).Formula = formule
        .Offset(0, 3).Value = unite
    End With
End Sub

' Colonnes de droite (sol d'ancrage, ple*, qce, Def, kp, kc, Rb) masquées pour les micropieux
' (catégories 17 à 20 : résistance de pointe non prise en compte). Appelée à la
' création et quand le type de pieu change (ChangementDonnees, module DonneesProjet).
Public Sub AfficherPointe(Optional ws As Worksheet)
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
    End If
    ws.Calculate                             ' catégorie (D11) à jour
    ws.Columns("T:AH").Hidden = (Val(ws.Range("D11").Value) >= 17)
End Sub

' Ligne r d'un tableau ple* (ple = True) ou Def pour la couche de la ligne fr du
' tableau du frottement axial : z haut et z bas cumulés depuis le TN, épaisseur
' comprise dans l'intervalle [D-b ; D+3a] (ple*) ou [D-10B ; D] (Def)
Private Sub AncrageLigne(ws As Worksheet, r As Long, fr As Long, ple As Boolean)
    Dim cond As String, zbas As String, pl As String
    pl = DONNEES & "$G$" & LIGNE_COUCHES + fr - 39     ' pl* de la couche (Données)
    cond = "=IF($C$" & fr & "="""","""","
    ws.Range("U" & r).Formula = cond & "$B$" & fr & ")"
    ws.Range("V" & r).Formula = cond & "$C$" & fr & ")"
    ws.Range("X" & r).Formula = cond & "SUM($F$39:$F$" & fr & "))"
    ws.Range("W" & r).Formula = cond & "X" & r & "-N($F$" & fr & "))"
    If ple Then
        ' dernière couche : prolongée sous la base (z bas très grand)
        If fr = 53 Then zbas = "1E+9" Else zbas = "IF(COUNTIF($C$" & fr + 1 & ":$C$53,""?*"")=0,1E+9,X" & r & ")"
        ws.Range("Y" & r).Formula = cond & "IFERROR(MAX(0,MIN(" & zbas & ",$W$51)-MAX(W" & r & ",$W$50)),""""))"
        ws.Range("Z" & r).Formula = "=IF(OR(Y" & r & "="""",NOT(ISNUMBER(" & pl & "))),""""," & _
            pl & "*Y" & r & ")"
    Else
        ws.Range("Y" & r).Formula = cond & "IFERROR(MAX(0,MIN(X" & r & ",$D$20)-MAX(W" & r & ",$W$56)),""""))"
        ws.Range("Z" & r).Formula = "=IF(OR(Y" & r & "="""",NOT(ISNUMBER(" & pl & ")),N($Z$26)=0)," & _
            """""," & pl & "*Y" & r & "/$Z$26)"
    End If
End Sub

' Une ligne « intitulé (B) | symbole (C) | valeur ou formule (D) | unité (E) »
Private Sub LigneS(ws As Worksheet, r As Long, intitule As String, symbole As String, _
        formule As String, unite As String)
    ws.Range("B" & r).Value = intitule
    ws.Range("C" & r).Value = symbole
    ws.Range("D" & r).Formula = formule
    ws.Range("E" & r).Value = unite
End Sub

' Tableau 2 x 2 (deux lignes « intitulé | symbole ») pour une plage de 2 x 2 cellules
Private Function Array2(a As String, b As String, c As String, d As String) As Variant
    Dim t(1 To 2, 1 To 2) As String
    t(1, 1) = a: t(1, 2) = b: t(2, 1) = c: t(2, 2) = d
    Array2 = t
End Function

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

' Valeur (ligne lig, colonne du sol sol) d'un tableau du frottement pour la méthode de
' calcul retenue : tableau G (6 colonnes de sols) si pénétrométrique, tableau F sinon
' (5 colonnes, ou 6 si f6) ; zones = « plage G|plage F »
Private Function IndexNorme(zones As String, lig As String, sol As String, _
        Optional f6 As Boolean = False) As String
    Dim z() As String
    z = Split(zones, "|")
    If mPenetro Then
        IndexNorme = "INDEX(" & z(0) & "," & lig & "," & Rang6(sol) & ")"
    Else
        IndexNorme = "INDEX(" & z(1) & "," & lig & "," & IIf(f6, Rang6(sol), Rang5(sol)) & ")"
    End If
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
    Dim cond As String, sol As String, pl As String, c As String
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
    ' pl* (Données, colonne G) ou qc (Données, colonne I) selon la méthode de calcul
    c = IIf(mPenetro, "$I$", "$G$")
    ws.Range("G" & r).Formula = cond & "IF(ISNUMBER(" & DONNEES & c & ld & ")," & _
        DONNEES & c & ld & ",""""))"
    mCellule = "H" & r
    ws.Range("H" & r).Formula = cond & "IFERROR(" & IndexNorme(zABC, "1", sol) & ",""""))"
    mCellule = "I" & r
    ws.Range("I" & r).Formula = cond & "IFERROR(" & IndexNorme(zAlpha, cat, sol) & ",""""))"
    ' fsol,i = (a.X + b)(1 - exp(-c.X)), X = pl* ou qc ; a, b, c du tableau F.5.2.2 ou G.5.2.2
    mCellule = "J" & r
    ws.Range("J" & r).Formula = "=IF(OR(B" & r & "="""",G" & r & "="""",H" & r & "=""""),"""",IFERROR(" & _
        "1000*(" & IndexNorme(zABC, "2", sol) & "*" & pl & "+" & IndexNorme(zABC, "3", sol) & _
        ")*(1-EXP(-" & IndexNorme(zABC, "4", sol) & "*" & pl & ")),""""))"
    ws.Range("K" & r).Formula = FormuleProduit("$I$" & r, "$J$" & r)
    mCellule = "L" & r
    ws.Range("L" & r).Formula = cond & "IFERROR(" & IndexNorme(zQs, cat, sol, True) & ",""""))"
    ws.Range("M" & r).Formula = FormuleQs("$K$" & r, "$L$" & r)
    ws.Range("N" & r).Formula = "=IF(OR(B" & r & "="""",F" & r & "="""",N($D$23)=0),""""," & _
        "$D$23*F" & r & ")"
    ' qs,i = qs / (gamma R;d1 x gamma R;d2) ; 0 si le terrain est neutralisé
    ws.Range("O" & r).Formula = "=IF(B" & r & "="""","""",IF(Q" & r & "=""Oui"",0," & _
        "IFERROR(M" & r & "/(" & gamma & "),""-"")))"
    ws.Range("P" & r).Formula = "=IF(AND(ISNUMBER(O" & r & "),ISNUMBER(N" & r & "))," & _
        "IF(Q" & r & "=""Oui"",0,O" & r & "*N" & r & "),IF(B" & r & "="""","""",""-""))"
    ws.Range("Q" & r).Formula = cond & DONNEES & "$J$" & ld & ")"
    ws.Range("E" & r & ":G" & r & ",I" & r).NumberFormat = "0.00"
    ws.Range("J" & r & ":M" & r & ",O" & r & ":P" & r).NumberFormat = "0.0"
    ws.Range("N" & r).NumberFormat = "0.000"
End Sub

' Valeurs d'un tableau de la feuille « Tableaux NF P94-262 », nommées d'après le numéro
' du tableau (ex. « Tableau F.5.2.1 » -> NF_F_5_2_1) : le titre du tableau est cherché
' exactement (cellule qui commence par « Tableau F.5.2.1 - », en respectant la casse,
' pour ne pas prendre un renvoi au tableau écrit dans une note), les valeurs commencent
' 2 lignes plus bas (sous l'en-tête) et 1 colonne à droite du titre (après la colonne
' des intitulés de ligne) ; nbLignes x nbCols valeurs. Renvoie le nom, utilisé dans les
' formules à la place de l'adresse.
Private Function ZoneTableau(wt As Worksheet, titre As String, decalage As Long, _
        nbLignes As Long, nbCols As Long) As String
    Dim c As Range, trouve As Range, premier As String, nom As String
    Set c = wt.Cells.Find(What:=titre & " - ", After:=wt.Cells(wt.Rows.Count, wt.Columns.Count), _
        LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=True)
    If Not c Is Nothing Then
        premier = c.Address
        Do
            If Left$(CStr(c.Value), Len(titre) + 3) = titre & " - " Then
                Set trouve = c
                Exit Do
            End If
            Set c = wt.Cells.FindNext(c)
            If c Is Nothing Then Exit Do
        Loop Until c.Address = premier
    End If
    If trouve Is Nothing Then Err.Raise vbObjectError + 1, , "« " & titre & " » introuvable " & _
        "dans la feuille « " & wt.Name & " » : relancez CreerTableauxNFP94262."
    nom = "NF_" & Replace(Replace(Trim$(Mid$(titre, Len("Tableau ") + 1)), ".", "_"), " ", "")
    ActiveWorkbook.Names.Add Name:=nom, RefersTo:="='" & wt.Name & "'!" & _
        trouve.Offset(2 + decalage, 1).Resize(nbLignes, nbCols).Address
    ZoneTableau = nom
End Function

' Tableaux d'ancrage par la méthode pénétrométrique (annexe G), à droite des tableaux
' pressiométriques (colonnes AB à AH), mêmes lignes : sol d'ancrage, résistance de pointe
' équivalente qce, hauteur d'encastrement équivalente Def, facteur de portance kc et
' Rb = Ab x kc x qce. qc,i lu dans la colonne qc (I) de « Données du projet ».
' Ecrêtage (G.4.2 (3)) : qcm = somme(qc,i x dz) / (b + 3a) sur [D-b ; D+3a],
' qcc = min(qc ; 1,3 qcm), qce = somme(qcc,i x dz) / (b + 3a) ; Def = somme(qc,i x dz)
' / qce sur [D-10B ; D]. a, b et les intervalles sont ceux de la méthode pressiométrique.
' kc = kc max (tableau G.4.2.1) si Def/B > 5, sinon k0 + (kc max - k0) x (Def/B) / 5
' avec k0 = 0,3 argiles, 0,2 sols intermédiaires, 0,1 sables et graves, 0,15 craies,
' marnes et roches.
Private Sub AncragePenetro(ws As Worksheet, zKc As String)
    Dim i As Long, r As Long, lt As Long, k0 As String
    SousTitre ws, 3, "Détermination du sol d'ancrage (méthode pénétrométrique)", "AB"
    LigneD ws, 4, "Couche d'ancrage (dernière couche renseignée)", "=$W$4", "", "AB"
    LigneD ws, 5, "Catégorie de sol", "=$W$5", "", "AB"
    LigneD ws, 6, "Résistance de pointe qc", "=IFERROR(LOOKUP(2,1/($AC$30:$AC$44<>""""),$AH$30:$AH$44),"""")", _
        "MPa", "AB"
    LigneD ws, 7, "Hauteur de la couche porteuse h", "=$D$21", "m", "AB"
    SousTitre ws, 9, "Détermination de la résistance de pointe équivalente qce", "AB"
    SousTitre ws, 28, "Détermination de la hauteur d'encastrement équivalente Def (méthode pénétrométrique)", "AB"
    For i = 0 To 1
        r = IIf(i = 0, 10, 29)               ' ligne d'en-tête
        ws.Range("AB" & r & ":AG" & r).Value = Array("N°", "Couche", "z haut (m)", "z bas (m)", _
            IIf(i = 0, ChrW(916) & "z dans [D-b ; D+3a] (m)", ChrW(916) & "z dans [D-10B ; D] (m)"), _
            IIf(i = 0, "qc,i x " & ChrW(916) & "z (MPa.m)", "qc,i x " & ChrW(916) & "z / qce (m)"))
        ws.Range("AH" & r).Value = IIf(i = 0, "qcc,i x " & ChrW(916) & "z (MPa.m)", "qc,i (MPa)")
        For lt = 1 To MAX_COUCHES
            AncrageLignePenetro ws, r + lt, 38 + lt, i = 0
        Next lt
    Next i
    ws.Range("AB26").Value = "qcm = " & ChrW(931) & " (qc,i x " & ChrW(916) & "z) / (b + 3a)  (MPa)"
    ws.Range("AB26:AF26").Merge
    ws.Range("AG26").Formula = "=IF(OR(COUNT(AG11:AG25)=0,N($W$52)=0),"""",SUM(AG11:AG25)/$W$52)"
    ws.Range("AB27").Value = "qce = " & ChrW(931) & " (qcc,i x " & ChrW(916) & _
        "z) / (b + 3a) avec qcc = min(qc ; 1,3 qcm)  (MPa)"
    ws.Range("AB27:AG27").Merge
    ws.Range("AH27").Formula = "=IF(OR(COUNT(AH11:AH25)=0,N($W$52)=0),"""",SUM(AH11:AH25)/$W$52)"
    ws.Range("AB45").Value = "Def = " & ChrW(931) & " (qc,i x " & ChrW(916) & "z) / qce  (m)"
    ws.Range("AB45:AF45").Merge
    ws.Range("AG45").Formula = "=IF(COUNT(AG30:AG44)=0,"""",SUM(AG30:AG44))"
    ws.Range("AB26:AH27,AB45:AG45").Font.Bold = True
    ws.Range("AD11:AH27,AD30:AH45").NumberFormat = "0.00"

    ' Paramètres des intervalles : identiques à la méthode pressiométrique
    LigneD ws, 47, "Paramètres de calcul de qce", "=$W$4", "", "AB"
    LigneD ws, 48, "a = max(B/2 ; 0,5 m)", "=$W$48", "m", "AB"
    LigneD ws, 49, "b = min(a ; h)", "=$W$49", "m", "AB"
    LigneD ws, 50, "D - b", "=$W$50", "m", "AB"
    LigneD ws, 51, "D + 3a", "=$W$51", "m", "AB"
    LigneD ws, 52, "b + 3a", "=$W$52", "m", "AB"
    LigneD ws, 54, "Paramètres de calcul de Def", "=$W$4", "", "AB"
    LigneD ws, 55, "10 B", "=$W$55", "m", "AB"
    LigneD ws, 56, "D - 10 B", "=$W$56", "m", "AB"
    ws.Range("AD48:AD52,AD55:AD56").NumberFormat = "0.00"

    SousTitre ws, 58, "Détermination du facteur de portance kc et de la résistance de pointe Rb", "AB"
    k0 = "CHOOSE(" & Rang6("$AD$5") & ",0.3,0.2,0.1,0.15,0.15,0.15)"
    LigneD ws, 59, "Classe de la fondation", "=$D$10", "", "AB"
    LigneD ws, 60, "kc max (tableau G.4.2.1, classe et sol d'ancrage)", _
        "=IFERROR(INDEX(" & zKc & ",VALUE(LEFT($D$10,1))," & Rang6("$AD$5") & "),"""")", "", "AB"
    LigneD ws, 61, "Def / B", "=IF(OR($AG$45="""",N($D$19)=0),"""",$AG$45/$D$19)", "", "AB"
    LigneD ws, 62, "kc = kc max si Def/B > 5, sinon k0 + (kc max - k0) x (Def/B) / 5", _
        "=IF(OR(NOT(ISNUMBER(AD60)),AD61=""""),"""",IF(AD61>5,AD60," & k0 & "+(AD60-" & k0 & ")*AD61/5))", _
        "", "AB"
    LigneD ws, 63, "Aire de la base Ab", "=$D$22", "m²", "AB"
    LigneD ws, 64, "Résistance de pointe équivalente qce", "=$AH$27", "MPa", "AB"
    LigneD ws, 65, "Rb = Ab x kc x qce", "=IF(OR(AD62="""",N(AD63)=0,AD64=""""),"""",1000*AD63*AD62*AD64)", _
        "kN", "AB"
    ws.Range("AD60:AD65").NumberFormat = "0.00"
    ws.Range("AB65:AE65").Font.Bold = True

    ' Mise en forme (comme les tableaux pressiométriques)
    With ws.Range("AB4:AE7,AB10:AH27,AB29:AH45,AB47:AE52,AB54:AE56,AB59:AE65")
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("AB10:AH10,AB29:AH29,AB47:AE47,AB54:AE54")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = VERT_ENTETE
    End With
    ws.Range("AD4:AD7,AD48:AD52,AD55:AD56,AG26,AH27,AG45,AD59:AD65").Interior.Color = JAUNE_CALCUL
    ws.Range("AD4:AE65,AB11:AB25,AD11:AH25,AB30:AB44,AD30:AH44").HorizontalAlignment = xlCenter
    ws.Columns("AA").ColumnWidth = 3
    ws.Columns("AB:AH").ColumnWidth = 16
End Sub

' Ligne r d'un tableau qce (qce = True) ou Def pénétrométrique pour la couche de la
' ligne fr du tableau du frottement axial ; qc,i lu dans « Données du projet »
' (colonne I, ligne de la couche)
Private Sub AncrageLignePenetro(ws As Worksheet, r As Long, fr As Long, qce As Boolean)
    Dim cond As String, zbas As String, qc As String
    cond = "=IF($C$" & fr & "="""","""","
    qc = DONNEES & "$I$" & LIGNE_COUCHES + fr - 39
    ws.Range("AB" & r).Formula = cond & "$B$" & fr & ")"
    ws.Range("AC" & r).Formula = cond & "$C$" & fr & ")"
    ws.Range("AE" & r).Formula = cond & "SUM($F$39:$F$" & fr & "))"
    ws.Range("AD" & r).Formula = cond & "AE" & r & "-N($F$" & fr & "))"
    If qce Then
        ' dernière couche : prolongée sous la base (z bas très grand)
        If fr = 53 Then zbas = "1E+9" Else zbas = "IF(COUNTIF($C$" & fr + 1 & ":$C$53,""?*"")=0,1E+9,AE" & r & ")"
        ws.Range("AF" & r).Formula = cond & "IFERROR(MAX(0,MIN(" & zbas & ",$W$51)-MAX(AD" & r & ",$W$50)),""""))"
        ws.Range("AG" & r).Formula = "=IF(OR(AF" & r & "="""",NOT(ISNUMBER(" & qc & "))),""""," & _
            qc & "*AF" & r & ")"
        ' qc écrêté à 1,3 qcm
        ws.Range("AH" & r).Formula = "=IF(OR(AF" & r & "="""",NOT(ISNUMBER(" & qc & ")),N($AG$26)=0),""""," & _
            "MIN(" & qc & ",1.3*$AG$26)*AF" & r & ")"
    Else
        ws.Range("AF" & r).Formula = cond & "IFERROR(MAX(0,MIN(AE" & r & ",$D$20)-MAX(AD" & r & ",$W$56)),""""))"
        ws.Range("AG" & r).Formula = "=IF(OR(AF" & r & "="""",NOT(ISNUMBER(" & qc & ")),N($AH$27)=0)," & _
            """"",(" & qc & ")*AF" & r & "/$AH$27)"
        ws.Range("AH" & r).Formula = cond & "IF(ISNUMBER(" & qc & ")," & qc & ",""""))"
    End If
End Sub
