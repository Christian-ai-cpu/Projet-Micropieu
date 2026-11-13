Attribute VB_Name = "VerificationArmature"
Option Explicit

' Module « Vérification de l'armature » : ajoute, sous les graphiques Y, T, M de la
' feuille « Vérification structurelle » (module VerificationStructurelle), la
' vérification de la section du tube du micropieu le plus long, après corrosion, selon
' la NF EN 1993-1-1 (6.2). Tableau en 3 colonnes : caractéristique (désignation en B
' et symbole en C), valeur (D), unité (E). Toutes les valeurs sont des formules.
'   1. Caractéristiques de la section : d0, t0 (nominaux), rc, d = d0 - 2 rc,
'      t = (d - di) / 2, A, I, Wel = 2 I / d, Wpl = (d³ - di³) / 6, fy, gamma M0 ;
'   2. Sollicitations (maximum des combinaisons du tableau des efforts en tête, issus
'      de la descente de charges) : NEd = Nc, VEd = H (= T0), MEd = M0 ;
'   3. Vérifications, en quatre volets (mini-titre, condition sous le titre, effort
'      résistant, rapport effort agissant / effort résistant, conclusion) :
'      3.1 classification du tube (tableau 5.2) : d/t <= 50 eps² (classe 1), 70 eps²
'          (classe 2), 90 eps² (classe 3), eps = racine(235 / fy) ;
'      3.2 compression (6.2.4) : NEd / Nc,Rd <= 1,0 (6.9), Nc,Rd = A fy / gamma M0 (6.10) ;
'      3.3 cisaillement (6.2.6) : VEd / Vpl,Rd <= 1,0 (6.17),
'          Vpl,Rd = Av (fy / racine 3) / gamma M0 (6.18), Av = 2 A / pi (6.2.6(3) g) ;
'      3.4 moment fléchissant (6.2.5) : MEd / Mc,Rd <= 1,0 (6.12),
'          Mc,Rd = Wpl fy / gamma M0 (classes 1 et 2, 6.13) ou Wel fy / gamma M0
'          (classe 3, 6.14) ;
'      conclusion : section vérifiée ou section à changer.
' Second tableau (par combinaison) : efforts le long du micropieu (NEd, VEd, MEd, y0,
' ymax), vérification des déplacements (ymax <= yadm), contraintes dans le tube
' (sigma = N/A + M.v/I, tau = V.Sy/(b.I), critère de von Mises pour les poutres
' sigma VM = racine(sigma² + 3 tau²) <= fy / gamma M0, tau e = fy / racine 3) et
' interaction NEd/Nc,Rd + MEd/Mc,Rd <= 1.
' Appelée à la fin de CreerVerificationStructurelle ; peut être relancée seule.

Private Const NOM_FEUILLE As String = "Vérification structurelle"
Private Const DONNEES As String = "'Données du projet'!"
Private Const TITRE As String = "VÉRIFICATION DE L'ARMATURE (NF EN 1993-1-1, 6.2) - TUBE APRÈS CORROSION"
Private Const TITRE2 As String = "EFFORTS SUR LE MICROPIEU - VÉRIFICATION DES DÉPLACEMENTS ET DES CONTRAINTES ADMISSIBLES"
Private Const LIB_YADM As String = "Déplacement horizontal admissible (à saisir)"
Private Const VERT_SAISIE As Long = 13561798   ' RGB(198, 239, 206)

Private Const JAUNE_CALCUL As Long = 13434879  ' RGB(255, 255, 204)
Private Const VERT_ENTETE As Long = 14348258   ' RGB(226, 239, 218)
Private Const VERT_FONCE As Long = 2315831     ' RGB(55, 86, 35)
Private Const GRIS_TITRE As Long = 15921906    ' RGB(242, 242, 242)

Private mR As Long                              ' ligne courante du tableau
Private mWs As Worksheet
Private mN As Long                              ' ligne de Nc (efforts en tête, colonnes E à H)
Private mP1 As Long, mPn As Long                ' 1re ligne et nombre de lignes du profil Y, T, M

Public Sub VerificationArmature()
    Dim ws As Worksheet, co As ChartObject, v As Variant, r0 As Long, etape As String
    Dim bas As Double, hP As Long, rN As Long, nP As Long, rFin As Long
    Dim rde As String, rdi As String, rAs As String, rI As String, rrc As String
    Dim d0 As String, t0 As String, rc As String, d As String, di As String, t As String
    Dim A As String, I As String, Wel As String, Wpl As String, fy As String, gM0 As String
    Dim NEd As String, VEd As String, MEd As String, eps As String, dt As String, cl As String
    Dim NcRd As String, Av As String, VplRd As String, McRd As String
    Dim rap1 As String, rap2 As String, rap3 As String, rap4 As String
    Dim ok1 As String, ok2 As String, ok3 As String, ok4 As String
    Dim r2 As Long, yAdm As Variant, vYa As String, xN As Long, xV As Long, xM As Long, xY0 As Long
    Dim xYm As Long, xRy As Long, xOkY As Long, xsN As Long, xsM As Long, xsx As Long, xtau As Long
    Dim xsa As Long, xta As Long, xcr As Long, xOkS As Long, xin As Long, xOkI As Long
    Dim xSy As Long, xb As Long, xsv As Long

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo Erreur
    If ws Is Nothing Then
        MsgBox "Créez d'abord la feuille « " & NOM_FEUILLE & " » (CreerVerificationStructurelle).", vbExclamation
        Exit Sub
    End If
    Set mWs = ws
    yAdm = LireD(ws, LIB_YADM)                ' déplacement admissible saisi (conservé)

    ' --- Cellules de la feuille « Vérification structurelle » (repérées par intitulé) ---
    etape = "lecture de la feuille"
    rde = Valeur(ws, "Diamètre après corrosion d - 2 rc")
    rdi = Valeur(ws, "Diamètre intérieur di (tube)")
    rAs = Valeur(ws, "Section après corrosion As")
    rI = Valeur(ws, "Inertie de la section après corrosion I")
    rrc = Valeur(ws, "Rayon de corrosion retenu rc (0 si protection efficace)")
    v = Application.Match("Nc", ws.Range("C1:C500"), 0)       ' efforts en tête (Nc, T0, H, M0)
    If rde = "" Or rAs = "" Or rI = "" Or IsError(v) Then
        MsgBox "Tableaux de la feuille « " & NOM_FEUILLE & " » introuvables : relancez " & _
            "CreerVerificationStructurelle.", vbExclamation
        Exit Sub
    End If
    rN = CLng(v)
    mN = rN
    v = Application.Match("z (m)", ws.Range("A1:A1000"), 0)    ' profil Y, T, M
    If Not IsError(v) Then
        hP = CLng(v)
        Do While Not IsEmpty(ws.Cells(hP + 2 + nP, 1).Value) Or ws.Cells(hP + 2 + nP, 1).HasFormula
            nP = nP + 1
            If nP > 5000 Then Exit Do
        Loop
        rFin = hP + 2 + nP
        mP1 = hP + 2: mPn = nP
    Else
        rFin = rN + 12
    End If

    ' --- Position : sous les graphiques (ou sous le profil) ; ancien tableau effacé
    ' (les graphiques ne sont pas touchés) ---
    etape = "position du tableau"
    v = Application.Match(TITRE, ws.Range("A1:A5000"), 0)
    If Not IsError(v) Then ws.Range(ws.Rows(CLng(v)), ws.Rows(CLng(v) + 160)).Clear
    bas = ws.Cells(rFin, 1).Top
    For Each co In ws.ChartObjects
        If co.Top + co.Height > bas Then bas = co.Top + co.Height
    Next co
    r0 = rFin
    Do While ws.Cells(r0, 1).Top < bas + 10
        r0 = r0 + 1
    Loop
    r0 = r0 + 1

    With ws.Range("A" & r0)
        .Value = TITRE
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
    mR = r0 + 1
    ws.Range("B" & mR & ":E" & mR).Value = Array("Caractéristique", "", "Valeur", "Unité")
    ws.Range("B" & mR & ":C" & mR).Merge
    With ws.Range("B" & mR & ":E" & mR)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .HorizontalAlignment = xlCenter
    End With

    ' --- 1. Caractéristiques de la section ---
    etape = "caractéristiques de la section"
    MiniTitre "1. Caractéristiques de la section (tube après corrosion)", ""
    d0 = Lig("Diamètre extérieur nominal", "d0", "=N(" & DONNEES & "$D$" & LIGNE_ACIER + 2 & ")", "mm", "0.0")
    t0 = Lig("Épaisseur nominale de la paroi", "t0", "=N(" & DONNEES & "$D$" & LIGNE_ACIER + 3 & ")", "mm", "0.0")
    rc = Lig("Rayon de corrosion retenu", "rc", "=" & IIf(rrc = "", "0", rrc), "mm", "0.00")
    d = Lig("Diamètre extérieur après corrosion", "d", "=" & rde, "mm", "0.0")
    di = Lig("Diamètre intérieur", "di", "=" & IIf(rdi = "", "0", rdi), "mm", "0.0")
    t = Lig("Épaisseur de paroi après corrosion", "t", "=(" & d & "-" & di & ")/2", "mm", "0.00")
    A = Lig("Aire de la section transversale", "A", "=" & rAs, "mm²", "0.0")
    I = Lig("Moment d'inertie de flexion", "I", "=" & rI, "mm4", "0.000E+00")
    Wel = Lig("Module de flexion élastique", "Wel", "=IF(N(" & d & ")>0,2*" & I & "/" & d & ","""")", "mm³", "0.000E+00")
    Wpl = Lig("Module de flexion plastique", "Wpl", "=(" & d & "^3-" & di & "^3)/6", "mm³", "0.000E+00")
    fy = Lig("Limite d'élasticité de l'acier", "fy", "=N(" & DONNEES & "$D$" & LIGNE_ACIER + 7 & ")", "MPa", "0")
    gM0 = Lig("Coefficient partiel (résistance des sections)", ChrW(947) & "M0", "=IF(N(" & DONNEES & "$D$" & _
        LIGNE_ACIER + 20 & ")>0,N(" & DONNEES & "$D$" & LIGNE_ACIER + 20 & "),1)", "", "0.00")

    ' --- 2. Sollicitations (maximum des combinaisons, descente de charges) ---
    etape = "sollicitations"
    MiniTitre "2. Sollicitations de calcul (maximum des combinaisons, descente de charges)", ""
    NEd = Lig("Effort normal de compression", "NEd", "=MAX(ABS(N(E" & rN & ")),ABS(N(F" & rN & ")),ABS(N(G" & rN & _
        ")),ABS(N(H" & rN & ")))", "kN", "0.0")
    VEd = Lig("Effort tranchant", "VEd", "=MAX(ABS(N(E" & rN + 2 & ")),ABS(N(F" & rN + 2 & ")),ABS(N(G" & rN + 2 & _
        ")),ABS(N(H" & rN + 2 & ")))", "kN", "0.0")
    MEd = Lig("Moment fléchissant", "MEd", "=MAX(ABS(N(E" & rN + 3 & ")),ABS(N(F" & rN + 3 & ")),ABS(N(G" & rN + 3 & _
        ")),ABS(N(H" & rN + 3 & ")))", "kN.m", "0.0")

    ' --- 3.1 Classification du tube (tableau 5.2) ---
    etape = "classification"
    MiniTitre "3.1 Classification de la section (tableau 5.2, section tubulaire)", _
        "Condition : classe 1 si d/t " & ChrW(8804) & " 50" & ChrW(949) & "² ; classe 2 si d/t " & ChrW(8804) & _
        " 70" & ChrW(949) & "² ; classe 3 si d/t " & ChrW(8804) & " 90" & ChrW(949) & "²"
    eps = Lig("Coefficient dépendant de fy", ChrW(949) & " = " & ChrW(8730) & "(235/fy)", "=IF(N(" & fy & ")>0,SQRT(235/" & _
        fy & "),"""")", "", "0.000")
    dt = Lig("Élancement de la paroi", "d/t", "=IF(N(" & t & ")>0," & d & "/" & t & ","""")", "", "0.00")
    Lig "Limite de la classe 3", "90" & ChrW(949) & "²", "=IF(" & eps & "="""",""""," & "90*" & eps & "^2)", "", "0.00"
    cl = Lig("Classe de la section", "Classe", "=IF(OR(" & dt & "=""""," & eps & "=""""),""-"",IF(" & dt & "<=50*" & _
        eps & "^2,1,IF(" & dt & "<=70*" & eps & "^2,2,IF(" & dt & "<=90*" & eps & "^2,3,4))))", "", "0")
    rap1 = Lig("Rapport d/t / 90" & ChrW(949) & "²", "d/t / 90" & ChrW(949) & "²", "=IF(OR(" & dt & "=""""," & eps & _
        "=""""),""""," & dt & "/(90*" & eps & "^2))", "", "0.00")
    ok1 = Lig("Vérification", "", "=IF(" & rap1 & "="""",""-"",IF(" & rap1 & "<=1,""OK (classe ""&" & cl & "&"")""," & _
        """Classe 4 : changer de section""))", "", "@")

    ' --- 3.2 Compression (6.2.4) ---
    etape = "compression"
    MiniTitre "3.2 Vérification en compression (6.2.4)", _
        "Condition : NEd / Nc,Rd " & ChrW(8804) & " 1,0 (6.9) avec Nc,Rd = A.fy / " & ChrW(947) & "M0 (6.10)"
    NcRd = Lig("Résistance de calcul à la compression", "Nc,Rd", "=IF(N(" & gM0 & ")>0," & A & "*" & fy & "/" & gM0 & _
        "/1000,"""")", "kN", "0.0")
    rap2 = Lig("Rapport effort agissant / effort résistant", "NEd / Nc,Rd", "=IF(N(" & NcRd & ")>0," & NEd & "/" & _
        NcRd & ","""")", "", "0.00")
    ok2 = Lig("Vérification", "", Conclusion(rap2), "", "@")

    ' --- 3.3 Cisaillement (6.2.6) ---
    etape = "cisaillement"
    MiniTitre "3.3 Vérification au cisaillement (6.2.6)", _
        "Condition : VEd / Vpl,Rd " & ChrW(8804) & " 1,0 (6.17) avec Vpl,Rd = Av (fy/" & ChrW(8730) & "3) / " & _
        ChrW(947) & "M0 (6.18) et Av = 2A/" & ChrW(960) & " (6.2.6(3) g)"
    Av = Lig("Aire de cisaillement (tube)", "Av = 2A/" & ChrW(960), "=2*" & A & "/PI()", "mm²", "0.0")
    VplRd = Lig("Résistance plastique de calcul au cisaillement", "Vpl,Rd", "=IF(N(" & gM0 & ")>0," & Av & "*(" & fy & _
        "/SQRT(3))/" & gM0 & "/1000,"""")", "kN", "0.0")
    rap3 = Lig("Rapport effort agissant / effort résistant", "VEd / Vpl,Rd", "=IF(N(" & VplRd & ")>0," & VEd & "/" & _
        VplRd & ","""")", "", "0.00")
    ok3 = Lig("Vérification", "", Conclusion(rap3), "", "@")

    ' --- 3.4 Moment fléchissant (6.2.5) ---
    etape = "moment fléchissant"
    MiniTitre "3.4 Vérification au moment fléchissant (6.2.5)", _
        "Condition : MEd / Mc,Rd " & ChrW(8804) & " 1,0 (6.12) avec Mc,Rd = Wpl.fy / " & ChrW(947) & _
        "M0 (classes 1 et 2, 6.13) ou Wel.fy / " & ChrW(947) & "M0 (classe 3, 6.14)"
    McRd = Lig("Moment résistant de calcul", "Mc,Rd", "=IF(OR(N(" & gM0 & ")<=0,N(" & cl & ")=0,N(" & cl & ")>3),""""," & _
        "IF(" & cl & "<=2," & Wpl & "," & Wel & ")*" & fy & "/" & gM0 & "/1000000)", "kN.m", "0.00")
    rap4 = Lig("Rapport effort agissant / effort résistant", "MEd / Mc,Rd", "=IF(N(" & McRd & ")>0," & MEd & "/" & _
        McRd & ","""")", "", "0.00")
    ok4 = Lig("Vérification", "", Conclusion(rap4), "", "@")

    ' --- Conclusion générale ---
    MiniTitre "Conclusion", ""
    Lig "Section du tube", "", "=IF(AND(LEFT(" & ok1 & ",2)=""OK"",LEFT(" & ok2 & ",2)=""OK"",LEFT(" & ok3 & _
        ",2)=""OK"",LEFT(" & ok4 & ",2)=""OK""),""Section vérifiée : OK"",""Section non vérifiée : changer de section"")", _
        "", "@"

    ' --- Mise en forme ---
    etape = "mise en forme"
    With ws.Range("B" & r0 + 1 & ":E" & mR)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("B" & r0 + 1 & ":E" & mR).FormatConditions
        .Delete
        .Add(xlTextString, String:="changer", TextOperator:=xlContains).Font.Color = RGB(192, 0, 0)
    End With

    ' === Second tableau : efforts le long du micropieu, vérification des déplacements et
    ' des contraintes admissibles du tube, pour chaque combinaison (colonnes D à G) ===
    etape = "efforts sur le micropieu"
    If hP = 0 Then GoTo Fin2
    mR = mR + 3
    r2 = mR
    With ws.Range("A" & r2)
        .Value = TITRE2
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
    mR = r2 + 1
    ws.Range("B" & mR & ":H" & mR).Value = Array("Caractéristique", "", "ELS CARA", "ELS QP", "ELU FOND", _
        "ELU ACC", "Unité")
    ws.Range("B" & mR & ":C" & mR).Merge
    With ws.Range("B" & mR & ":H" & mR)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .HorizontalAlignment = xlCenter
    End With

    ' 4. Efforts et déplacements le long du micropieu (profil Y, T, M)
    MiniTitre2 "4. Efforts et déplacements le long du micropieu (profil, par combinaison)", ""
    xN = LigC("Effort normal de compression (descente de charges)", "NEd", "=ABS(N(@" & rN & "))", "kN", "0.0")
    xV = LigC("Effort tranchant maximal le long du micropieu", "VEd = max|T(z)|", _
        "=IF(COUNT({T})=0,0,MAX(MAX({T}),-MIN({T})))", "kN", "0.0")
    xM = LigC("Moment fléchissant maximal le long du micropieu", "MEd = max|M(z)|", _
        "=IF(COUNT({M})=0,0,MAX(MAX({M}),-MIN({M})))", "kN.m", "0.00")
    xY0 = LigC("Déplacement horizontal en tête", "y0 = Y(z = 0)", "=N({Y0})", "mm", "0.00")
    xYm = LigC("Déplacement horizontal maximal", "ymax = max|Y(z)|", _
        "=IF(COUNT({Y})=0,0,MAX(MAX({Y}),-MIN({Y})))", "mm", "0.00")

    ' 5. Vérification des déplacements
    MiniTitre2 "5. Vérification des déplacements", _
        "Condition : ymax / yadm " & ChrW(8804) & " 1,0 ; yadm : déplacement horizontal admissible fixé par " & _
        "le projet (compatibilité avec l'ouvrage)"
    mR = mR + 1
    ws.Range("B" & mR).Value = LIB_YADM
    ws.Range("C" & mR).Value = "yadm"
    With ws.Range("D" & mR & ":G" & mR)
        .Merge
        .Value = IIf(N0(yAdm) > 0, N0(yAdm), 10)
        .Interior.Color = VERT_SAISIE
        .NumberFormat = "0.00"
    End With
    ws.Range("H" & mR).Value = "mm"
    ws.Range("C" & mR & ":H" & mR).HorizontalAlignment = xlCenter
    vYa = "$D$" & mR
    xRy = LigC("Rapport déplacement agissant / admissible", "ymax / yadm", "=IF(N(" & vYa & ")>0,#" & xYm & _
        "/" & vYa & ","""")", "", "0.00")
    xOkY = LigC("Vérification", "", ConclusionC(xRy), "", "@")

    ' 6. Contraintes dans le tube et critère de von Mises (théorie des poutres, élasticité) :
    '    contrainte normale de compression et de flexion à la fibre extrême
    '    sigma = N / A + M.v / I (v = d / 2, soit M / Wel avec Wel = 2 I / d) ;
    '    contrainte de cisaillement sous effort tranchant tau = V.Sy(A') / (b.I), au droit
    '    de l'axe neutre pour le tube : Sy = (d³ - di³) / 12 (moment statique de la
    '    demi-couronne), b = d - di (deux parois) ;
    '    critère de von Mises pour les poutres sigma VM = racine(sigma² + 3 tau²) <= fy
    '    (fy / gamma M0) ; limite élastique en cisaillement pur tau e = fy / racine 3.
    MiniTitre2 "6. Contraintes dans le tube et critère de von Mises (élasticité, théorie des poutres)", _
        "Condition : " & ChrW(963) & "VM = " & ChrW(8730) & "(" & ChrW(963) & "² + 3" & ChrW(964) & "²) " & _
        ChrW(8804) & " fy/" & ChrW(947) & "M0 avec " & ChrW(963) & " = N/A + M.v/I et " & ChrW(964) & _
        " = V.Sy/(b.I) ; " & ChrW(964) & " " & ChrW(8804) & " " & ChrW(964) & "e = fy/(" & ChrW(8730) & "3." & _
        ChrW(947) & "M0)"
    xsN = LigC("Contrainte normale de compression", ChrW(963) & "N = NEd / A", "=IF(N(" & A & ")>0,1000*#" & _
        xN & "/" & A & ","""")", "MPa", "0.0")
    xsM = LigC("Contrainte normale de flexion (fibre extrême, v = d/2)", ChrW(963) & "M = MEd.v / I", _
        "=IF(N(" & I & ")>0,1000000*#" & xM & "*(" & d & "/2)/" & I & ","""")", "MPa", "0.0")
    xsx = LigC("Contrainte normale totale", ChrW(963) & " = " & ChrW(963) & "N + " & ChrW(963) & "M", _
        "=IF(OR(#" & xsN & "="""",#" & xsM & "=""""),"""",#" & xsN & "+#" & xsM & ")", "MPa", "0.0")
    xSy = LigC("Moment statique de la demi-section (axe neutre)", "Sy = (d³ - di³) / 12", "=(" & d & "^3-" & di & _
        "^3)/12", "mm³", "0.000E+00")
    xb = LigC("Largeur de la section au droit de l'axe neutre", "b = d - di", "=" & d & "-" & di, "mm", "0.00")
    xtau = LigC("Contrainte de cisaillement sous effort tranchant", ChrW(964) & " = VEd.Sy / (b.I)", _
        "=IF(N(#" & xb & ")*N(" & I & ")>0,1000*#" & xV & "*#" & xSy & "/(#" & xb & "*" & I & "),"""")", "MPa", "0.00")
    xta = LigC("Limite élastique en cisaillement pur", ChrW(964) & "e = fy / (" & ChrW(8730) & "3." & ChrW(947) & "M0)", _
        "=IF(N(" & gM0 & ")>0," & fy & "/(SQRT(3)*" & gM0 & "),"""")", "MPa", "0.0")
    LigC "Rapport cisaillement agissant / limite", ChrW(964) & " / " & ChrW(964) & "e", "=IF(N(#" & xta & _
        ")>0,#" & xtau & "/#" & xta & ","""")", "", "0.00"
    xsv = LigC("Contrainte équivalente de von Mises", ChrW(963) & "VM = " & ChrW(8730) & "(" & ChrW(963) & "² + 3" & _
        ChrW(964) & "²)", "=IF(OR(#" & xsx & "="""",#" & xtau & "=""""),"""",SQRT(#" & xsx & "^2+3*#" & xtau & _
        "^2))", "MPa", "0.0")
    xsa = LigC("Contrainte admissible", "fy / " & ChrW(947) & "M0", "=IF(N(" & gM0 & ")>0," & fy & "/" & gM0 & _
        ","""")", "MPa", "0.0")
    xcr = LigC("Rapport contrainte agissante / admissible", ChrW(963) & "VM / (fy/" & ChrW(947) & "M0)", _
        "=IF(OR(#" & xsv & "="""",N(#" & xsa & ")=0),"""",#" & xsv & "/#" & xsa & ")", "", "0.00")
    xOkS = LigC("Vérification", "", ConclusionC(xcr), "", "@")

    ' 7. Interaction effort normal et moment (6.2.1(7))
    MiniTitre2 "7. Interaction effort normal et moment fléchissant (6.2.1(7))", _
        "Condition (6.2) : NEd / Nc,Rd + MEd / Mc,Rd " & ChrW(8804) & " 1 (Nc,Rd et Mc,Rd du tableau précédent, " & _
        "3.2 et 3.4)"
    xin = LigC("Somme des rapports effort agissant / effort résistant", "NEd/Nc,Rd + MEd/Mc,Rd", _
        "=IF(OR(N(" & NcRd & ")=0,N(" & McRd & ")=0),"""",#" & xN & "/" & NcRd & "+#" & xM & "/" & McRd & ")", _
        "", "0.00")
    xOkI = LigC("Vérification", "", ConclusionC(xin), "", "@")

    ' Conclusion par combinaison
    MiniTitre2 "Conclusion", ""
    LigC "Micropieu (déplacements, contraintes, interaction)", "", "=IF(AND(LEFT(#" & xOkY & ",2)=""OK"",LEFT(#" & _
        xOkS & ",2)=""OK"",LEFT(#" & xOkI & ",2)=""OK""),""OK"",""Non vérifié : changer de section"")", "", "@"

    With ws.Range("B" & r2 + 1 & ":H" & mR)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("B" & r2 + 1 & ":H" & mR).FormatConditions
        .Delete
        .Add(xlTextString, String:="changer", TextOperator:=xlContains).Font.Color = RGB(192, 0, 0)
    End With
Fin2:
    Exit Sub
Erreur:
    MsgBox "Erreur pendant la vérification de l'armature, partie « " & etape & " » : " & Err.Description & _
        " (n° " & Err.Number & ")", vbExclamation
End Sub

' Mini-titre d'un passage du tableau (B:E fusionnées), suivi éventuellement de la
' condition de vérification (ligne en italique sous le titre)
Private Sub MiniTitre(texte As String, condition As String)
    mR = mR + 1
    With mWs.Range("B" & mR & ":E" & mR)
        .Merge
        .Value = texte
        .Font.Bold = True
        .Font.Color = VERT_FONCE
        .Interior.Color = GRIS_TITRE
    End With
    If condition <> "" Then
        mR = mR + 1
        With mWs.Range("B" & mR & ":E" & mR)
            .Merge
            .Value = condition
            .Font.Italic = True
            .WrapText = True
        End With
        mWs.Rows(mR).RowHeight = 30
    End If
End Sub

' Ligne « désignation (B) | symbole (C) | valeur (D) | unité (E) » ; renvoie l'adresse
' absolue de la valeur
Private Function Lig(designation As String, symbole As String, formule As String, unite As String, _
        fmt As String) As String
    mR = mR + 1
    mWs.Range("B" & mR).Value = designation
    mWs.Range("C" & mR).Value = symbole
    mWs.Range("D" & mR).Formula = formule
    mWs.Range("E" & mR).Value = unite
    mWs.Range("C" & mR & ":E" & mR).HorizontalAlignment = xlCenter
    mWs.Range("D" & mR).Interior.Color = JAUNE_CALCUL
    If fmt <> "@" Then mWs.Range("D" & mR).NumberFormat = fmt
    If designation = "Vérification" Or Left(designation, 7) = "Section" Then mWs.Range("B" & mR & ":D" & mR).Font.Bold = True
    Lig = "$D$" & mR
End Function

' Conclusion d'une vérification d'après le rapport effort agissant / effort résistant
Private Function Conclusion(rapport As String) As String
    Conclusion = "=IF(" & rapport & "="""",""-"",IF(" & rapport & "<=1,""OK (taux ""&TEXT(" & rapport & _
        ",""0%"")&"")"",""Non vérifiée : changer de section""))"
End Function

' Mini-titre du second tableau (B:H fusionnées), suivi éventuellement de la condition
Private Sub MiniTitre2(texte As String, condition As String)
    mR = mR + 1
    With mWs.Range("B" & mR & ":H" & mR)
        .Merge
        .Value = texte
        .Font.Bold = True
        .Font.Color = VERT_FONCE
        .Interior.Color = GRIS_TITRE
    End With
    If condition <> "" Then
        mR = mR + 1
        With mWs.Range("B" & mR & ":H" & mR)
            .Merge
            .Value = condition
            .Font.Italic = True
            .WrapText = True
        End With
        mWs.Rows(mR).RowHeight = 30
    End If
End Sub

' Ligne par combinaison du second tableau : désignation (B), symbole (C), formules des
' colonnes D à G (combinaison k = 0 à 3), unité (H). Dans le modèle tpl : « # » = colonne
' de la combinaison dans ce tableau, « @ » = colonne de la combinaison dans le tableau
' des efforts en tête (E à H), {Y}, {T}, {M} = colonne du profil de la combinaison, {Y0}
' = déplacement en tête. Renvoie le numéro de la ligne.
Private Function LigC(designation As String, symbole As String, tpl As String, unite As String, _
        fmt As String) As Long
    Dim k As Long, f As String, cP As Long
    mR = mR + 1
    mWs.Range("B" & mR).Value = designation
    mWs.Range("C" & mR).Value = symbole
    mWs.Range("H" & mR).Value = unite
    For k = 0 To 3
        cP = 5 + 3 * k                          ' colonne Y de la combinaison dans le profil
        f = tpl
        f = Replace(f, "{Y0}", mWs.Cells(mP1, cP).Address)
        f = Replace(f, "{Y}", mWs.Range(mWs.Cells(mP1, cP), mWs.Cells(mP1 + mPn - 1, cP)).Address)
        f = Replace(f, "{T}", mWs.Range(mWs.Cells(mP1, cP + 1), mWs.Cells(mP1 + mPn - 1, cP + 1)).Address)
        f = Replace(f, "{M}", mWs.Range(mWs.Cells(mP1, cP + 2), mWs.Cells(mP1 + mPn - 1, cP + 2)).Address)
        f = Replace(f, "#", Chr(Asc("D") + k))
        f = Replace(f, "@", Chr(Asc("E") + k))
        mWs.Cells(mR, 4 + k).Formula = f
        If fmt <> "@" Then mWs.Cells(mR, 4 + k).NumberFormat = fmt
    Next k
    mWs.Range("C" & mR & ":H" & mR).HorizontalAlignment = xlCenter
    mWs.Range("D" & mR & ":G" & mR).Interior.Color = JAUNE_CALCUL
    If designation = "Vérification" Or Left(designation, 9) = "Micropieu" Then mWs.Range("B" & mR & ":G" & mR).Font.Bold = True
    LigC = mR
End Function

' Conclusion d'une vérification par combinaison (rapport en ligne r, même colonne)
Private Function ConclusionC(r As Long) As String
    ConclusionC = "=IF(#" & r & "="""",""-"",IF(#" & r & "<=1,""OK (taux ""&TEXT(#" & r & _
        ",""0%"")&"")"",""Non vérifiée : changer de section""))"
End Function

' Valeur (colonne D) de la ligne d'intitulé lib (colonne B) ; vide si absente
Private Function LireD(ws As Worksheet, lib As String) As Variant
    Dim v As Variant
    v = Application.Match(lib, ws.Range("B1:B5000"), 0)
    If Not IsError(v) Then LireD = ws.Range("D" & CLng(v)).Value
End Function

' Valeur numérique (0 si vide ou texte)
Private Function N0(v As Variant) As Double
    If Not IsEmpty(v) And IsNumeric(v) Then N0 = CDbl(v)
End Function

' Adresse ($D$n) de la valeur de la ligne d'intitulé lib (colonne B) ; vide si absente
Private Function Valeur(ws As Worksheet, lib As String) As String
    Dim v As Variant
    v = Application.Match(lib, ws.Range("B1:B500"), 0)
    If Not IsError(v) Then Valeur = "$D$" & CLng(v)
End Function
