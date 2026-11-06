Attribute VB_Name = "VerificationStructurelle"
Option Explicit

' Module « Vérification structurelle » : crée (ou recrée) la feuille « Vérification
' structurelle » des micropieux, à partir des feuilles « Données du projet » (module
' DonneesProjet), « Calcul et vérification » (module CalculFondation) et de la feuille de
' descente de charges (matrice des micropieux, module MatriceMicropieux). Toutes les
' valeurs sont des formules : elles suivent les saisies de ces feuilles.
'   B3  : caractéristiques du micropieu : armature (tube N80, barre HA ou GEWI), épaisseur
'         sacrificielle si nappe phréatique (NF EN 1993-5, tableaux 4.1 / 4.2, selon la
'         durée d'utilisation de l'ouvrage), section, inertie et raideur EI de l'armature
'         (après corrosion)
'   B22 : réaction du sol dans la couche molle (terrain neutralisé) : courbe de réaction
'         selon le type de sollicitation (a : permanentes en tête, Es = Kf ; b : poussées
'         latérales du sol, Kf puis Kf/2 ; c : courte durée en tête, Es = 2 Kf ; d :
'         accidentelles très brèves, 2 Kf puis Kf), paliers Pf et Pu ;
'           essai pressiométrique : Kf = EM x 18 rho / [4 (2,65 rho)^alpha + 3 alpha rho],
'                                   rho = max(B/B0 ; 1), Pf = pf* B, Pu = pl* B ;
'           essai CPT : Kf = beta qc / 2, Pf = B qc / beta1, Pu = B qc / beta2 ;
'           longueur de transfert l0 = (4 EpIp / Es)^(1/4), micropieu rigide (l0 >= D)
'           ou souple (D >= 3 l0)
'   B50 : poussée des terres au repos sur la fondation existante (soubassement de la
'         descente de charges) : h = hauteur du poteau du soubassement + hauteur de la
'         semelle, K0 = 1 - sin phi', sigma v = gamma h, sigma h = K0 sigma v,
'         H = K0 gamma h² / 2 (kN/ml), excentricité e = h / 3 au-dessus de la fibre
'         inférieure de la semelle (tête des micropieux)
'   A64 : efforts dans chaque micropieu et chaque combinaison (ELS caractéristique, ELS
'         quasi permanent, ELU fondamental, ELU accidentel) : compression Nc = Vmax de la
'         vérification de la profondeur, traction Nt = - Nc, effort tranchant en tête
'         V = gamma G,sup x H x L (L : longueur de mur reprise par le micropieu, matrice
'         de la descente de charges), moment en tête M = V x e
' Le nombre de micropieux et les adresses des autres feuilles sont lus à la création :
' relancer CreerVerificationStructurelle après un changement de ces feuilles.

Private Const NOM_FEUILLE As String = "Vérification structurelle"
Private Const DONNEES As String = "'Données du projet'!"
Private Const CALCUL As String = "'Calcul et vérification'!"
Private Const LIGNE_EFFORTS As Long = 64      ' titre du tableau des efforts

' Thème : saisies en vert, valeurs calculées en jaune, en-têtes en vert clair
Private Const VERT_SAISIE As Long = 13561798   ' RGB(198, 239, 206)
Private Const JAUNE_CALCUL As Long = 13434879  ' RGB(255, 255, 204)
Private Const VERT_ENTETE As Long = 14348258   ' RGB(226, 239, 218)
Private Const VERT_FONCE As Long = 2315831     ' RGB(55, 86, 35)

Public Sub CreerVerificationStructurelle()
    Dim ws As Worksheet, wc As Worksheet, wd As Worksheet, f As Worksheet, etape As String
    Dim a As Long, lp As Long, lm As Long, r As Long, i As Long, k As Long, m As Long, n As Long
    Dim r0 As Long, rv As Long, cC As Long, tC As Long, rd As Long, rP As Long, rS As Long
    Dim tube As String, dn As String, entete As Variant

    Application.ScreenUpdating = False
    On Error GoTo Erreur
    etape = "feuilles du projet"
    Set wc = ActiveWorkbook.Worksheets("Calcul et vérification")
    For Each f In ActiveWorkbook.Worksheets
        If CStr(f.Range("B80").Value) = "RÉPARTITION DES CHARGES DES MURS SUR LES MICROPIEUX" Then
            Set wd = f
            Exit For
        End If
    Next f

    etape = "création de la feuille"
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo Erreur
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = NOM_FEUILLE
    End If
    ws.Cells.Clear
    a = LIGNE_ACIER                           ' caractéristiques de l'acier (Données)
    lp = LIGNE_PIEU                           ' caractéristiques de la fondation (Données)
    lm = LIGNE_COUCHE_MOLLE                   ' frottement négligé et couche molle (Données)

    With ws.Range("B1:I1")
        .Merge
        .Value = "VÉRIFICATION STRUCTURELLE DES MICROPIEUX"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
        .HorizontalAlignment = xlCenter
    End With

    ' --- Caractéristiques du micropieu (armature, corrosion, section, inertie, EI) ---
    etape = "caractéristiques du micropieu"
    SousTitre ws, 3, "Caractéristiques du micropieu"
    Ligne ws, 4, "Type d'armature", "=" & DONNEES & "D" & a, ""
    Ligne ws, 5, "Désignation de l'armature", "=" & DONNEES & "D" & a + 1, ""
    Ligne ws, 6, "Diamètre extérieur (tube) ou nominal (barre) d", "=N(" & DONNEES & "D" & a + 2 & ")", "mm"
    Ligne ws, 7, "Épaisseur du tube e", "=N(" & DONNEES & "D" & a + 3 & ")", "mm"
    Ligne ws, 8, "Diamètre intérieur du tube di", "=N(" & DONNEES & "D" & a + 4 & ")", "mm"
    Ligne ws, 9, "Module de Young E", "=N(" & DONNEES & "D" & a + 6 & ")", "MPa"
    Ligne ws, 10, "Limite d'élasticité fyk", "=N(" & DONNEES & "D" & a + 7 & ")", "MPa"
    Ligne ws, 11, "Présence de nappe phréatique", "=" & DONNEES & "D34", ""
    Ligne ws, 12, "Durée d'utilisation de l'ouvrage", "=" & DONNEES & "D" & lp + 20, "ans"
    Ligne ws, 13, "Environnement de l'armature (NF EN 1993-5)", "=" & DONNEES & "D" & a + 12, ""
    ' Épaisseur sacrificielle : perte d'épaisseur du tableau 4.1 (sols) ou 4.2 (eau) de la
    ' NF EN 1993-5 pour la durée d'utilisation, seulement s'il y a une nappe (corrosion)
    Ligne ws, 14, "Corrosion prise en compte (nappe phréatique)", "=IF(D11=""Oui"",""Oui"",""Non"")", ""
    Ligne ws, 15, "Épaisseur sacrificielle " & ChrW(916) & "e (tableau 4.1 / 4.2)", _
        "=IF(D14=""Oui"",N(" & DONNEES & "D" & a + 13 & "),0)", "mm"
    ' Tube : corrosion de la face extérieure (intérieur rempli de coulis) ; barre : sur tout
    ' le pourtour
    tube = "D4=""Tube N80"""
    Ligne ws, 16, "Diamètre extérieur après corrosion d - 2 " & ChrW(916) & "e", "=MAX(0,D6-2*D15)", "mm"
    Ligne ws, 17, "Section d'acier As (après corrosion)", "=IF(" & tube & ",PI()/4*(D16^2-D8^2)," & _
        "PI()/4*D16^2)", "mm²"
    Ligne ws, 18, "Inertie I (après corrosion)", "=IF(" & tube & ",PI()/64*(D16^4-D8^4)," & _
        "PI()/64*D16^4)", "mm4"
    Ligne ws, 19, "Raideur en flexion EI", "=D9*D18*1E-9", "kN.m²"
    Ligne ws, 20, "Résistance plastique de la section As x fyk", "=D17*D10/1000", "kN"
    ws.Range("D6:D8,D16").NumberFormat = "0.0"
    ws.Range("D9:D10").NumberFormat = "0"
    ws.Range("D15").NumberFormat = "0.00"
    ws.Range("D17,D20").NumberFormat = "0.0"
    ws.Range("D18").NumberFormat = "0.000E+00"
    ws.Range("D19").NumberFormat = "0.00"

    ' --- Réaction du sol dans la couche molle et longueur de transfert ---
    ' Courbe de réaction (fonction linéaire par morceaux : pente Es, palier de fluage Pf,
    ' palier ultime Pu) selon le type de sollicitation (situations a à d) :
    '   a - permanentes en tête dominantes      : Es = Kf jusqu'à Pf
    '   b - poussées latérales du sol dominantes : Es = Kf jusqu'à Pf, Kf / 2 jusqu'à Pu
    '   c - courte durée en tête dominantes      : Es = 2 Kf jusqu'à Pf
    '   d - accidentelles très brèves en tête    : Es = 2 Kf jusqu'à Pf, Kf jusqu'à Pu
    ' Essai pressiométrique : Kf = EM x 18 rho / [4 (2,65 rho)^alpha + 3 alpha rho],
    ' rho = max(B / B0 ; 1), Pf = pf* B, Pu = pl* B ; essai CPT : Kf = beta qc / 2,
    ' Pf = B qc / beta1, Pu = B qc / beta2 (beta, beta1, beta2 : tableau I.4.1 / I.4.2).
    ' Longueur de transfert l0 = (4 EI / Es)^(1/4) ; micropieu rigide si l0 >= D, souple
    ' si D >= 3 l0.
    etape = "réaction du sol"
    SousTitre ws, 22, "Réaction du sol dans la couche molle et longueur de transfert (courbe de réaction)"
    Ligne ws, 23, "Méthode de calcul (type d'essai)", "=" & DONNEES & CELLULE_METHODE, ""
    Ligne ws, 24, "Type de sollicitation", "=" & DONNEES & "D" & lp + 21, ""
    Ligne ws, 25, "Diamètre du micropieu B (diamètre final retenu)", "=N(" & CALCUL & "$D$19)", "m"
    Ligne ws, 26, "Diamètre de référence B0", "=N(" & DONNEES & "D" & lp + 22 & ")", "m"
    Ligne ws, 27, ChrW(961) & " = max(B / B0 ; 1)", "=IF(D26<=0,"""",MAX(D25/D26,1))", ""
    Ligne ws, 28, "Épaisseur de la couche molle (terrain neutralisé)", "=N(" & DONNEES & "D" & lm + 2 & ")", "m"
    Ligne ws, 29, "=IF(D23=""Méthode pénétrométrique"",""Résistance de pointe qc de la couche molle""," & _
        """Module pressiométrique EM de la couche molle"")", "=N(" & DONNEES & "D" & lm + 1 & ")", "MPa"
    ' pl* de la 1re couche neutralisée (« Données du projet », colonne G)
    Ligne ws, 30, "Pression limite nette pl* de la couche molle", _
        "=IFERROR(INDEX(" & DONNEES & "$G$" & LIGNE_COUCHES & ":$G$" & LIGNE_COUCHES + MAX_COUCHES - 1 & _
        ",MATCH(""Oui""," & DONNEES & "$M$" & LIGNE_COUCHES & ":$M$" & LIGNE_COUCHES + MAX_COUCHES - 1 & _
        ",0)),0)", "MPa"
    Ligne ws, 31, "Pression de fluage nette pf* (à défaut d'essai : pl* / 1,5)", "=D30/1.5", "MPa"
    Ligne ws, 32, "Sol rocheux", "=" & DONNEES & "D" & lm + 3, ""
    Ligne ws, 33, "Catégorie du sol", "=" & DONNEES & "D" & lm + 4, ""
    Ligne ws, 34, "Nature du sol", "=" & DONNEES & "D" & lm + 5, ""
    Ligne ws, 35, "Coefficient rhéologique " & ChrW(945) & " (tableaux I.1.3.1 / I.1.3.2)", _
        "=N(" & DONNEES & "D" & lm + 6 & ")", ""
    ' beta, beta1, beta2 (CPT) : sols argileux (argile, tourbe) 12 / 5 / 3,5 ; sols
    ' intermédiaires (limon) 7,5 / 10 / 6 ; sols sableux (sable, grave) et craies, marnes,
    ' rochers 4,5 / 13 / 8
    Ligne ws, 36, ChrW(946) & " (essai CPT)", "=IF(OR(D34=""Tourbe"",D34=""Argile""),12,IF(D34=""Limon"",7.5,4.5))", ""
    Ligne ws, 37, ChrW(946) & "1 (essai CPT)", "=IF(OR(D34=""Tourbe"",D34=""Argile""),5,IF(D34=""Limon"",10,13))", ""
    Ligne ws, 38, ChrW(946) & "2 (essai CPT)", "=IF(OR(D34=""Tourbe"",D34=""Argile""),3.5,IF(D34=""Limon"",6,8))", ""
    Ligne ws, 39, "Module de réaction Kf", _
        "=IF(OR(D25<=0,D29<=0,D27=""""),"""",IF(D23=""Méthode pénétrométrique"",D36*D29/2," & _
        "D29*18*D27/(4*(2.65*D27)^D35+3*D35*D27)))", "MPa"
    Ligne ws, 40, "Palier de fluage Pf", "=IF(D23=""Méthode pénétrométrique"",D25*D29/D37,D31*D25)*1000", "kN/m"
    Ligne ws, 41, "Palier ultime Pu", "=IF(D23=""Méthode pénétrométrique"",D25*D29/D38,D30*D25)*1000", "kN/m"
    Ligne ws, 42, "Module de réaction Es (pente jusqu'au palier de fluage Pf)", _
        "=IF(D39="""","""",IF(OR(LEFT(D24,1)=""c"",LEFT(D24,1)=""d""),2*D39,D39))", "MPa"
    Ligne ws, 43, "Pente entre Pf et Pu", _
        "=IF(D39="""","""",IF(LEFT(D24,1)=""b"",D39/2,IF(LEFT(D24,1)=""d"",D39,0)))", "MPa"
    Ligne ws, 44, "Courbe de réaction appliquée", _
        "=IF(LEFT(D24,1)=""a"",""Es = Kf, P limité à Pf""," & _
        "IF(LEFT(D24,1)=""b"",""Es = Kf jusqu'à Pf, Kf/2 entre Pf et Pu, palier Pu""," & _
        "IF(LEFT(D24,1)=""c"",""Es = 2 Kf, P limité à Pf""," & _
        "IF(LEFT(D24,1)=""d"",""Es = 2 Kf jusqu'à Pf, Kf entre Pf et Pu, palier Pu"",""-""))))", ""
    Ligne ws, 45, "Raideur en flexion EpIp (= EI)", "=D19", "kN.m²"
    Ligne ws, 46, "Longueur de transfert l0 = (4 EpIp / Es)^(1/4)", _
        "=IF(OR(D42="""",N(D42)=0),"""",(4*D45/(D42*1000))^0.25)", "m"
    Ligne ws, 47, "Longueur du micropieu D", "=N(" & CALCUL & "$D$20)", "m"
    Ligne ws, 48, "Comportement du micropieu", _
        "=IF(D46="""",""-"",IF(D46>=D47,""Rigide (court) : l0 >= D"",IF(D47>=3*D46," & _
        """Souple (long) : D >= 3 l0"",""Intermédiaire"")))", ""
    ws.Range("D25:D28").NumberFormat = "0.000"
    ws.Range("D29:D31,D35:D39,D42:D43,D45:D47").NumberFormat = "0.00"
    ws.Range("D40:D41").NumberFormat = "0.0"

    ' --- Poussée des terres au repos sur la fondation existante (effort en tête) ---
    etape = "poussée des terres"
    SousTitre ws, 50, "Poussée des terres sur la fondation existante (effort en tête des micropieux)"
    If Not wd Is Nothing Then
        ' Lignes du soubassement dans la descente de charges (mur 1) : poteau et semelle
        For r = 5 To 79
            If CStr(wd.Range("B" & r).Value) = "Soubassement" Then
                For i = r + 1 To r + 6
                    If Left(CStr(wd.Range("B" & i).Value), 6) = "Poteau" And rP = 0 Then rP = i
                    If Left(CStr(wd.Range("B" & i).Value), 7) = "Semelle" And rS = 0 Then rS = i
                Next i
                Exit For
            End If
        Next r
        dn = "'" & wd.Name & "'!"
    End If
    Ligne ws, 51, "Hauteur du poteau du soubassement hp", IIf(rP > 0, "=N(" & dn & "$F$" & rP & ")", "0"), "m"
    Ligne ws, 52, "Hauteur de la semelle existante hs", IIf(rS > 0, "=N(" & dn & "$F$" & rS & ")", "0"), "m"
    Ligne ws, 53, "Hauteur de terre h = hp + hs", "=D51+D52", "m"
    Ligne ws, 54, "Poids volumique du sol " & ChrW(947) & " (couche 1)", _
        "=N(" & DONNEES & "J" & LIGNE_COUCHES & ")", "kN/m3"
    Ligne ws, 55, "Cohésion c' (couche 1)", "=N(" & DONNEES & "K" & LIGNE_COUCHES & ")", "kPa"
    Ligne ws, 56, "Angle de frottement " & ChrW(966) & "' (couche 1)", "=N(" & DONNEES & "L" & LIGNE_COUCHES & ")", "°"
    Ligne ws, 57, "Coefficient des terres au repos K0 = 1 - sin " & ChrW(966) & "'", _
        "=1-SIN(RADIANS(D56))", ""
    Ligne ws, 58, "Contrainte verticale à la base " & ChrW(963) & "v = " & ChrW(947) & " h", "=D54*D53", "kPa"
    Ligne ws, 59, "Contrainte horizontale au repos " & ChrW(963) & "h = K0 " & ChrW(963) & "v", "=D57*D58", "kPa"
    Ligne ws, 60, "Effort horizontal H = K0 " & ChrW(947) & " h² / 2", "=D57*D54*D53^2/2", "kN/ml"
    Ligne ws, 61, "Excentricité e = h / 3 (par rapport à la tête du micropieu)", "=D53/3", "m"
    Ligne ws, 62, "Feuille de descente de charges", IIf(wd Is Nothing, "introuvable", ""), ""
    If Not wd Is Nothing Then ws.Range("D62").Value = wd.Name
    ws.Range("D51:D61").NumberFormat = "0.00"

    ' --- Efforts dans les micropieux ---
    etape = "efforts dans les micropieux"
    ' Vérification de la profondeur (« Calcul et vérification ») : en-tête « N° MICROPIEU »
    ' en colonne A, combinaisons (Vmax) dans la colonne « COMBINAISONS ... », coefficients
    ' des combinaisons sous le tableau
    For r = 380 To 450
        If CStr(wc.Range("A" & r).Value) = "N° MICROPIEU" Then r0 = r: Exit For
    Next r
    If r0 > 0 Then
        For i = 1 To 60
            If CStr(wc.Cells(r0, i).Value) = "COMBINAISONS DE CHARGES VERTICALES (kN)" Then cC = i
        Next i
        Do While Left(CStr(wc.Range("A" & r0 + 3 + m).Value), 2) = "MP"
            m = m + 1
        Loop
        For r = r0 + 3 + m To r0 + 3 + m + 10
            If Left(CStr(wc.Range("A" & r).Value), 14) = "COEFFICIENTS D" Then tC = r: Exit For
        Next r
    End If
    If Not wd Is Nothing Then
        Do While Left(CStr(wd.Range("B" & 86 + n).Value), 4) = "Mur "
            n = n + 1
        Loop
        rd = 88 + n + 3                       ' MP1 dans la matrice des micropieux
    End If
    With ws.Range("A" & LIGNE_EFFORTS)
        .Value = "EFFORTS DANS LES MICROPIEUX"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
    If m = 0 Or cC = 0 Or wd Is Nothing Then
        ws.Range("A" & LIGNE_EFFORTS + 2).Value = "Créez la matrice des micropieux (descente de charges) " & _
            "et « Calcul et vérification », puis relancez CreerVerificationStructurelle."
        GoTo MiseEnForme
    End If
    ' En-têtes : A micropieu, B longueur de mur reprise, C à F compression Nc, G à J traction
    ' Nt, K à N effort tranchant V, O à R moment M (ELS CARA, ELS QP, ELU FOND, ELU ACC)
    rv = LIGNE_EFFORTS + 2
    ws.Range("A" & rv).Value = "N° MICROPIEU"
    ws.Range("B" & rv).Value = "Longueur de mur reprise L (m)"
    ws.Range("C" & rv).Value = "EFFORT DE COMPRESSION Nc (kN)"
    ws.Range("G" & rv).Value = "EFFORT DE TRACTION Nt (kN)"
    ws.Range("K" & rv).Value = "EFFORT TRANCHANT EN TÊTE V (kN)"
    ws.Range("O" & rv).Value = "MOMENT EN TÊTE M = V x e (kN.m)"
    ws.Range("A" & rv & ":A" & rv + 1).Merge
    ws.Range("B" & rv & ":B" & rv + 1).Merge
    For k = 0 To 3
        ws.Range(ws.Cells(rv, 3 + 4 * k), ws.Cells(rv, 6 + 4 * k)).Merge
        ws.Range(ws.Cells(rv + 1, 3 + 4 * k), ws.Cells(rv + 1, 6 + 4 * k)).Value = _
            Array("ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
    Next k
    Commentaire ws.Range("C" & rv), "Nc : Vmax de chaque combinaison (vérification de la profondeur, " & _
        "« Calcul et vérification »)"
    Commentaire ws.Range("G" & rv), "Nt : effort de traction, opposé de l'effort de compression"
    Commentaire ws.Range("K" & rv), "V = " & ChrW(947) & "G,sup x H x L : poussée des terres au repos " & _
        "sur la fondation existante (H en kN/ml) multipliée par la longueur de mur reprise L et par le " & _
        "coefficient " & ChrW(947) & "G,sup de la combinaison"
    Commentaire ws.Range("O" & rv), "M = V x e : e = h / 3, distance entre le point d'application de " & _
        "la poussée et la fibre inférieure de la semelle existante (tête du micropieu)"
    entete = Array("ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
    For i = 1 To m
        r = rv + 1 + i
        ws.Range("A" & r).Value = "MP" & i
        ' Longueur de mur reprise : somme de la ligne du micropieu dans la matrice
        ws.Range("B" & r).Formula = "=SUM(" & dn & wd.Range(wd.Cells(rd + i - 1, 3), _
            wd.Cells(rd + i - 1, 2 + n)).Address & ")"
        For k = 0 To 3
            ws.Cells(r, 3 + k).Formula = "=N(" & CALCUL & wc.Cells(r0 + 2 + i, cC + 2 * k).Address & ")"
            ws.Cells(r, 7 + k).Formula = "=-" & ws.Cells(r, 3 + k).Address(False, False)
            ' gamma G,sup de la combinaison (tableau des coefficients, colonne D) ; 1 sinon
            If tC > 0 Then
                ws.Cells(r, 11 + k).Formula = "=N(" & CALCUL & "$D$" & tC + 2 + k & ")*$D$60*B" & r
            Else
                ws.Cells(r, 11 + k).Formula = "=" & Choose(k + 1, "1", "1", "1.35", "1") & "*$D$60*B" & r
            End If
            ws.Cells(r, 15 + k).Formula = "=" & ws.Cells(r, 11 + k).Address(False, False) & "*$D$61"
        Next k
    Next i
    With ws.Range("A" & rv & ":R" & rv + 1 + m)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A" & rv & ":R" & rv + 1)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .WrapText = True
    End With
    ws.Range("A" & rv + 2 & ":A" & rv + 1 + m).Font.Bold = True
    ws.Range("B" & rv + 2 & ":R" & rv + 1 + m).Interior.Color = JAUNE_CALCUL
    ws.Range("B" & rv + 2 & ":R" & rv + 1 + m).NumberFormat = "0.0"

MiseEnForme:
    etape = "mise en forme"
    With ws.Range("B4:E20,B23:E48,B51:E62")
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("D4:E62").HorizontalAlignment = xlCenter
    ws.Range("D4:D20,D23:D48,D51:D62").Interior.Color = JAUNE_CALCUL
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:R").ColumnWidth = 12
    ws.Activate
Fin:
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    MsgBox "Erreur pendant la création de la feuille « " & NOM_FEUILLE & " », partie « " & etape & _
        " » : " & Err.Description & " (n° " & Err.Number & ")", vbExclamation
    Resume Fin
End Sub

' Une ligne « intitulé (B:C fusionnées) | valeur ou formule (D) | unité (E) » ; un
' intitulé commençant par « = » est une formule
Private Sub Ligne(ws As Worksheet, r As Long, intitule As String, formule As String, unite As String)
    If Left(intitule, 1) = "=" Then
        ws.Range("B" & r).Formula = intitule
    Else
        ws.Range("B" & r).Value = intitule
    End If
    ws.Range("B" & r & ":C" & r).Merge
    If Left(formule, 1) = "=" Then
        ws.Range("D" & r).Formula = formule
    Else
        ws.Range("D" & r).Value = formule
    End If
    ws.Range("E" & r).Value = unite
End Sub

' Sous-titre en gras, souligné
Private Sub SousTitre(ws As Worksheet, r As Long, texte As String)
    With ws.Range("B" & r)
        .Value = texte
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
End Sub

' Commentaire (signification) sur une cellule d'en-tête
Private Sub Commentaire(c As Range, texte As String)
    c.ClearComments
    c.AddComment texte
    c.Comment.Shape.TextFrame.AutoSize = True
End Sub
