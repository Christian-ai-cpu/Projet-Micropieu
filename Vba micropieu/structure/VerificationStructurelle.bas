Attribute VB_Name = "VerificationStructurelle"
Option Explicit

' Module « Vérification structurelle » : crée (ou recrée) la feuille « Vérification
' structurelle » des micropieux, à partir des feuilles « Données du projet » (module
' DonneesProjet), « Calcul et vérification » (module CalculFondation) et de la feuille de
' descente de charges (matrice des micropieux, module MatriceMicropieux). Toutes les
' valeurs sont des formules : elles suivent les saisies de ces feuilles.
'   A3  : tableau vertical du micropieu le plus long (plus grande profondeur réelle de
'         « Calcul et vérification ») : diamètre du pieu foré (Phi retenu ou autre
'         valeur saisie), enrobage et protection par le coulis (c >= 5 cm, C > 500
'         kg/m3, E/C < 0,5 : pas de corrosion), rayon de corrosion, armature après
'         corrosion, I, EI, EM et alpha (ou qc et beta), rho, Kf selon la méthode de
'         calcul, Es selon le type de sollicitation, l0 = (4 EI / Es)^(1/4), lambda,
'         profondeur réelle D, x = D / l0, inclinaison et angle
'   dessous : butée sur le sol environnant : composante horizontale des contraintes du
'         sol due à l'ancrage (couches d'ancrage de « Données du projet ») :
'         sigma h,k = qs sin(alpha), sigma h,d = qs,i sin(alpha), Fh = sigma h pi B hi ;
'         efforts en tête de la descente de charges (Nc, T0 = Nc sin(alpha), H = T0,
'         M = 0) et constantes A, B, C, E du pieu long à pointe encastrée, avec leurs
'         symboles et formules ;
'         profil le long du micropieu par pas de calcul dz (saisi) : z, pas, Y, T et M
'         de chaque combinaison, graphiques Y, T, M à droite du profil
' Le nombre de micropieux et les adresses des autres feuilles sont lus à la création :
' relancer CreerVerificationStructurelle après un changement de ces feuilles.

Private Const NOM_FEUILLE As String = "Vérification structurelle"
Private Const DONNEES As String = "'Données du projet'!"
' Intitulés des saisies conservées quand la feuille est recréée
Private Const LIB_AUTRE As String = "Autre diamètre (à saisir, vide : " & "Phi retenu)"
Private Const LIB_PAS As String = "Pas de calcul dz (à saisir)"
Private Const LIB_LIBRE As String = "Micropieu libre en tête (Oui / Non)"
Private Const CALCUL As String = "'Calcul et vérification'!"

' Thème : saisies en vert, valeurs calculées en jaune, en-têtes en vert clair
Private Const VERT_SAISIE As Long = 13561798   ' RGB(198, 239, 206)
Private Const JAUNE_CALCUL As Long = 13434879  ' RGB(255, 255, 204)
Private Const VERT_ENTETE As Long = 14348258   ' RGB(226, 239, 218)
Private Const VERT_FONCE As Long = 2315831     ' RGB(55, 86, 35)

Public Sub CreerVerificationStructurelle()
    Dim ws As Worksheet, wc As Worksheet, wd As Worksheet, f As Worksheet, etape As String
    Dim a As Long, lp As Long, lm As Long, r As Long, i As Long, k As Long, m As Long, n As Long
    Dim r0 As Long, cC As Long, tC As Long, rd As Long, rPo As Long, rS As Long, j As Long
    Dim tube As String, dn As String, entete As Variant, unites As Variant, x As String, cl As String
    Dim dA As String, dI As String, eA As String, pen As String, sol As String, b0 As String
    Dim emq As String, alp As String, nat As String, bet As String
    Dim hT As String, k0 As String, hH As String, eE As String
    Dim cPhi As Long, cDr As Long, cIn As Long, cAn As Long, iMax As Long, ci As Long
    Dim autre As Variant, condTete As Variant, rF As Long, rT As Long, rK As Long, rP As Long, rG As Long
    Dim nZ As Long, hZ As Variant, xs As String, zk As String, kA As String, kB As String
    Dim y0s As String, hs As String, ths As String, mhs As String, cond0 As String
    Dim vPhi As String, vAut As String, vB As String, vdA As String, vdi As String, vc As String
    Dim vPro As String, vrcD As String, vrc As String, vdc As String, vS0 As String, vS As String
    Dim vI As String, vE As String, vEI As String, vEM As String, vAl As String, vB0 As String
    Dim vRho As String, vKf As String, vEs As String, vl0 As String, vLam As String, vD As String
    Dim vx As String, vInc As String, vAng As String, vCond As String, vLm As String, vHp As String
    Dim vEx As String, vPas As String, pasZ As Variant, hF As Variant, nbC As Long, rA As Long
    Dim libre As Variant

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
    autre = LireValeur(ws, LIB_AUTRE)          ' saisies conservées : autre diamètre,
    pasZ = LireValeur(ws, LIB_PAS)             ' pas de calcul dz du profil
    libre = LireValeur(ws, LIB_LIBRE)          ' micropieu libre en tête (Oui / Non)
    ws.Cells.Clear
    Do While ws.ChartObjects.Count > 0       ' graphiques de la création précédente
        ws.ChartObjects(1).Delete
    Loop
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

    ' --- Paramètres lus dans les autres feuilles (formules directes) ---
    ' Armature (« Données du projet », caractéristiques de l'acier) : type, d, di, E
    tube = DONNEES & "$D$" & a & "=""Tube N80"""
    dA = "N(" & DONNEES & "$D$" & a + 2 & ")"
    dI = "N(" & DONNEES & "$D$" & a + 4 & ")"
    eA = "N(" & DONNEES & "$D$" & a + 6 & ")"
    ' Méthode de calcul, type de sollicitation, B0, couche molle : EM ou qc, alpha, beta
    ' (CPT : sols argileux 12, limons 7,5, sols sableux, craies, marnes, rochers 4,5)
    pen = "(" & DONNEES & CELLULE_METHODE & "=""Méthode pénétrométrique"")"
    sol = DONNEES & "$D$" & lp + 21
    b0 = "N(" & DONNEES & "$D$" & lp + 22 & ")"
    emq = "N(" & DONNEES & "$D$" & lm + 1 & ")"
    alp = "N(" & DONNEES & "$D$" & lm + 6 & ")"
    nat = DONNEES & "$D$" & lm + 5
    bet = "IF(OR(" & nat & "=""Tourbe""," & nat & "=""Argile""),12,IF(" & nat & "=""Limon"",7.5,4.5))"

    etape = "micropieux de la vérification de la profondeur"
    ' Vérification de la profondeur (« Calcul et vérification ») : en-tête « N° MICROPIEU »
    ' en colonne A, combinaisons (Vmax) dans la colonne « COMBINAISONS ... », coefficients
    ' des combinaisons sous le tableau
    ' (position variable : sous le tableau de la capacité portante par pas de calcul)
    On Error Resume Next
    r0 = Application.Match("N° MICROPIEU", wc.Range("A1:A5000"), 0)
    On Error GoTo Erreur
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
    If m = 0 Or cC = 0 Then
        ws.Range("A3").Value = "Créez la matrice des micropieux (descente de charges) " & _
            "et « Calcul et vérification », puis relancez CreerVerificationStructurelle."
        GoTo MiseEnForme
    End If
    ' --- Micropieu le plus long : plus grande profondeur réelle de la vérification de la
    ' profondeur (« Calcul et vérification ») ; colonnes repérées par leurs en-têtes ---
    etape = "micropieu le plus long"
    For i = 1 To 80
        If CStr(wc.Cells(r0 + 1, i).Value) = ChrW(934) & " RETENU" Then cPhi = i
        If CStr(wc.Cells(r0, i).Value) = "PROFONDEUR RÉELLE (m)" Then cDr = i
        If cDr = 0 And CStr(wc.Cells(r0, i).Value) = "PROFONDEUR RETENUE (m)" Then cDr = i
        If CStr(wc.Cells(r0 + 1, i).Value) = "Incliné" Then cIn = i
        If CStr(wc.Cells(r0 + 1, i).Value) = "Angle (°)" Then cAn = i
    Next i
    iMax = 1
    If cDr > 0 Then
        For i = 2 To m
            If N0(wc.Cells(r0 + 2 + i, cDr).Value) > N0(wc.Cells(r0 + 2 + iMax, cDr).Value) Then iMax = i
        Next i
    End If
    ci = r0 + 2 + iMax                        ' ligne du micropieu dans « Calcul et vérification »

    ' --- Tableau vertical : longueur de transfert l0, inclinaison, poussée des terres ---
    ' Intitulé en B:C, valeur en D, unité en E. Diamètre du pieu foré B = Phi retenu ou
    ' autre valeur saisie ; enrobage c = (B - d) / 2 ; protection par le coulis (acier en
    ' compression) si c >= 5 cm, dosage > 500 kg/m3 et E/C < 0,5 : pas de corrosion ;
    ' armature après corrosion (tube : face extérieure, barre : tout le pourtour) ;
    ' I et EI après corrosion ; Kf selon la méthode de calcul (pressiométrique :
    ' EM x 18 rho / [4 (2,65 rho)^alpha + 3 alpha rho], rho = max(B / B0 ; 1) ;
    ' pénétrométrique : beta qc / 2) ; Es selon le type de sollicitation ;
    ' l0 = (4 EI / Es)^(1/4) ; poussée des terres au repos H = K0 gamma h² / 2, e = h / 3.
    etape = "longueur de transfert"
    With ws.Range("A3")
        .Value = "LONGUEUR DE TRANSFERT l0 DU MICROPIEU LE PLUS LONG (MP" & iMax & ")"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
    r = 4
    r = r + 1: Ligne ws, r, "Micropieu le plus long (plus grande profondeur réelle)", "MP" & iMax, ""
    r = r + 1: Ligne ws, r, ChrW(934) & " retenu (vérification de la profondeur)", _
        IIf(cPhi > 0, "=IF(ISNUMBER(" & CALCUL & wc.Cells(ci, cPhi).Address & ")," & CALCUL & _
        wc.Cells(ci, cPhi).Address & ",N(" & CALCUL & "$D$19))", "=N(" & CALCUL & "$D$19)"), "m": vPhi = "$D$" & r
    r = r + 1: Ligne ws, r, LIB_AUTRE, "", "m": vAut = "$D$" & r
    If IsNumeric(autre) And Not IsEmpty(autre) Then ws.Range("D" & r).Value = autre
    ws.Range("D" & r).Interior.Color = VERT_SAISIE
    r = r + 1: Ligne ws, r, "Diamètre du pieu foré B", "=IF(N(" & vAut & ")>0," & vAut & "," & vPhi & ")", "m": vB = "$D$" & r
    r = r + 1: Ligne ws, r, "Section du pieu foré A = " & ChrW(960) & ".B²/4", "=PI()*" & vB & "^2/4", "m²"
    r = r + 1: Ligne ws, r, "Diamètre de l'armature d", "=" & dA, "mm": vdA = "$D$" & r
    r = r + 1: Ligne ws, r, "Diamètre intérieur di (tube)", "=IF(" & tube & "," & dI & ",0)", "mm": vdi = "$D$" & r
    r = r + 1: Ligne ws, r, "Enrobage c = (B - d) / 2", "=MAX(0,(1000*" & vB & "-" & vdA & ")/2)", "mm": vc = "$D$" & r
    r = r + 1: Ligne ws, r, "Protection par le coulis (c >= 50 mm, C > 500 kg/m3, E/C < 0,5, compression)", _
        "=IF(AND(" & vc & ">=50,N(" & DONNEES & "D" & lp + 28 & ")>500,N(" & DONNEES & "D" & lp + 27 & _
        ")>0,N(" & DONNEES & "D" & lp + 27 & ")<0.5," & DONNEES & "D" & lp + 12 & _
        "<>""Traction""),""Efficace"",""Non efficace"")", "": vPro = "$D$" & r
    r = r + 1: Ligne ws, r, "Rayon de corrosion (« Données du projet »)", "=N(" & DONNEES & "D" & lp + 30 & ")", _
        "mm": vrcD = "$D$" & r
    r = r + 1: Ligne ws, r, "Rayon de corrosion retenu rc (0 si protection efficace)", _
        "=IF(" & vPro & "=""Efficace"",0," & vrcD & ")", "mm": vrc = "$D$" & r
    r = r + 1: Ligne ws, r, "Diamètre après corrosion d - 2 rc", "=MAX(0," & vdA & "-2*" & vrc & ")", "mm": vdc = "$D$" & r
    r = r + 1: Ligne ws, r, "Section brute de l'armature", "=IF(" & tube & ",PI()/4*(" & vdA & "^2-" & vdi & _
        "^2),PI()/4*" & vdA & "^2)", "mm²": vS0 = "$D$" & r
    r = r + 1: Ligne ws, r, "Section après corrosion As", "=MAX(0,IF(" & tube & ",PI()/4*(" & vdc & "^2-" & vdi & _
        "^2),PI()/4*" & vdc & "^2))", "mm²": vS = "$D$" & r
    r = r + 1: Ligne ws, r, "Section corrodée", "=" & vS0 & "-" & vS, "mm²"
    r = r + 1: Ligne ws, r, "Inertie de la section brute", "=IF(" & tube & ",PI()/64*(" & vdA & "^4-" & vdi & _
        "^4),PI()/64*" & vdA & "^4)", "mm4"
    r = r + 1: Ligne ws, r, "Inertie de la section après corrosion I", "=MAX(0,IF(" & tube & ",PI()/64*(" & vdc & _
        "^4-" & vdi & "^4),PI()/64*" & vdc & "^4))", "mm4": vI = "$D$" & r
    r = r + 1: Ligne ws, r, "Module de Young de l'armature E", "=" & eA, "MPa": vE = "$D$" & r
    r = r + 1: Ligne ws, r, "Module de flexion EI = E x I", "=" & vE & "*" & vI & "*1E-9", "kN.m²": vEI = "$D$" & r
    r = r + 1: Ligne ws, r, "Méthode de calcul", "=" & DONNEES & CELLULE_METHODE, ""
    r = r + 1: Ligne ws, r, "=IF(" & pen & ",""Résistance de pointe qc de la couche molle""," & _
        """Module pressiométrique EM de la couche molle"")", "=" & emq, "MPa": vEM = "$D$" & r
    r = r + 1: Ligne ws, r, "=IF(" & pen & ",""" & ChrW(946) & " (CPT)"",""Coefficient rhéologique " & ChrW(945) & _
        """)", "=IF(" & pen & "," & bet & "," & alp & ")", "": vAl = "$D$" & r
    r = r + 1: Ligne ws, r, "Diamètre de référence B0", "=" & b0, "m": vB0 = "$D$" & r
    r = r + 1: Ligne ws, r, ChrW(961) & " = max(B / B0 ; 1)", "=IF(" & vB0 & "<=0,"""",MAX(" & vB & "/" & vB0 & ",1))", _
        "": vRho = "$D$" & r
    r = r + 1: Ligne ws, r, "Module de réaction Kf (selon la méthode de calcul)", "=IF(OR(N(" & vEM & ")<=0," & vRho & _
        "=""""),"""",IF(" & pen & "," & vAl & "*" & vEM & "/2," & vEM & "*18*" & vRho & "/(4*(2.65*" & vRho & ")^" & _
        vAl & "+3*" & vAl & "*" & vRho & ")))", "MPa": vKf = "$D$" & r
    r = r + 1: Ligne ws, r, "Type de sollicitation", "=" & sol, ""
    r = r + 1: Ligne ws, r, "Module Es (a, b : Kf ; c, d : 2 Kf)", "=IF(" & vKf & "="""",""""," & _
        "IF(OR(LEFT(" & sol & ",1)=""c"",LEFT(" & sol & ",1)=""d""),2,1)*" & vKf & ")", "MPa": vEs = "$D$" & r
    r = r + 1: Ligne ws, r, "Longueur de transfert l0 = (4 EI / Es)^(1/4)", "=IF(OR(" & vEs & "="""",N(" & vEs & _
        ")=0),"""",(4*" & vEI & "/(" & vEs & "*1000))^0.25)", "m": vl0 = "$D$" & r
    r = r + 1: Ligne ws, r, ChrW(955) & " = 1 / l0", "=IF(N(" & vl0 & ")<=0,"""",1/" & vl0 & ")", "1/m": vLam = "$D$" & r
    r = r + 1: Ligne ws, r, "Profondeur réelle du micropieu D", IIf(cDr > 0, "=N(" & CALCUL & _
        wc.Cells(ci, cDr).Address & ")", "=N(" & CALCUL & "$D$20)"), "m": vD = "$D$" & r
    r = r + 1: Ligne ws, r, "x = D / l0", "=IF(" & vLam & "="""",""""," & vLam & "*" & vD & ")", "": vx = "$D$" & r
    r = r + 1: Ligne ws, r, "Comportement du micropieu", "=IF(" & vx & "="""",""-"",IF(" & vx & "<=1," & _
        """Rigide (l0 >= D)"",IF(" & vx & ">=3,""Pieu long (D >= 3 l0)"",""Intermédiaire (calcul approché)"")))", ""
    r = r + 1: Ligne ws, r, "Micropieu incliné", IIf(cIn > 0, "=IF(" & CALCUL & wc.Cells(ci, cIn).Address & _
        "=""Oui"",""Oui"",""Non"")", "Non"), "": vInc = "$D$" & r
    r = r + 1: Ligne ws, r, "Angle d'inclinaison sur la verticale " & ChrW(945), IIf(cAn > 0, "=IF(" & vInc & _
        "=""Oui"",N(" & CALCUL & wc.Cells(ci, cAn).Address & "),0)", "0"), "°": vAng = "$D$" & r
    ' Micropieu libre en tête (liste Oui / Non, saisie conservée) : par défaut d'après
    ' « Données du projet » (conditions en tête de la fondation) ; Non = encastré en tête
    r = r + 1: Ligne ws, r, LIB_LIBRE, "", "": vCond = "$D$" & r
    With ws.Range("D" & r)
        .Validation.Delete
        .Validation.Add xlValidateList, xlValidAlertStop, , "Oui,Non"
        .Value = IIf(CStr(ActiveWorkbook.Worksheets("Données du projet").Range("D" & lp + 18).Value) = _
            "Encastré en tête", "Non", "Oui")
        If CStr(libre) = "Oui" Or CStr(libre) = "Non" Then .Value = libre
    End With
    Commentaire ws.Range("B" & r), "Oui : tête libre (rotation libre, M0 = 0) ; Non : tête encastrée " & _
        "dans la semelle (rotation nulle, moment d'encastrement M0 = -H.l0/2). Valeur par défaut : " & _
        "« Données du projet », conditions en tête de la fondation"
    rF = r
    With ws.Range("B5:E" & rF)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("D5:E" & rF).HorizontalAlignment = xlCenter
    ws.Range("D5:D" & rF).Interior.Color = JAUNE_CALCUL
    ws.Range(vAut & "," & vCond).Interior.Color = VERT_SAISIE
    ws.Range("D5:D" & rF).NumberFormat = "0.000"
    ws.Range(vI & "," & vI & ":" & vI).NumberFormat = "0.000E+00"
    ws.Range("D" & Mid(vI, 4) - 1).NumberFormat = "0.000E+00"
    ws.Range("B" & Mid(vl0, 4) & ":D" & Mid(vl0, 4) & ",B" & Mid(vx, 4) + 1 & ":D" & Mid(vx, 4) + 1).Font.Bold = True
    With ws.Range(vPro).FormatConditions
        .Delete
        .Add(xlCellValue, xlEqual, "=""Non efficace""").Font.Color = RGB(192, 0, 0)
    End With

    ' --- Butée sur le sol environnant : composante horizontale des contraintes du sol
    ' due à l'ancrage du micropieu (couches « prise comme couche d'ancrage » de « Données
    ' du projet ») ---
    ' Le frottement axial qs mobilisé le long de l'axe du micropieu incliné d'un angle
    ' alpha sur la verticale a une composante horizontale qs x sin(alpha) :
    '   caractéristique : sigma h,k = qs x sin(alpha)   (qs = min(alpha pieu-sol fsol ; qs,max))
    '   de calcul       : sigma h,d = qs,i x sin(alpha) (qs,i = qs / (gamma R;d1 x gamma R;d2))
    '   efforts horizontaux : Fh = sigma h x pi x B x hi (couche d'ancrage, non neutralisée)
    ' qs, qs,i et hi : tableau du frottement axial de « Calcul et vérification ».
    etape = "butée due à l'ancrage"
    hF = Application.Match("N°", wc.Range("B1:B120"), 0)  ' en-tête du frottement axial
    nbC = CLng(N0(ActiveWorkbook.Worksheets("Données du projet").Range(CELLULE_NB_COUCHES).Value))
    If nbC < 1 Then nbC = 1
    If nbC > MAX_COUCHES Then nbC = MAX_COUCHES
    rA = rF + 3
    With ws.Range("A" & rA)
        .Value = "BUTÉE SUR LE SOL ENVIRONNANT - COMPOSANTE HORIZONTALE DES CONTRAINTES DUE À L'ANCRAGE (MP" & _
            iMax & ")"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
    r = rA + 1
    ws.Range("A" & r & ":K" & r).Value = Array("N° de couche", "Nom de couche", "Épaisseur hi (m)", _
        "Couche d'ancrage", "qs caractéristique (kPa)", "qs,i de calcul (kPa)", "Angle " & ChrW(945) & " (°)", _
        ChrW(963) & "h,k = qs.sin " & ChrW(945) & " (kPa)", ChrW(963) & "h,d = qs,i.sin " & ChrW(945) & " (kPa)", _
        "Fh,k = " & ChrW(963) & "h,k." & ChrW(960) & ".B.hi (kN)", "Fh,d = " & ChrW(963) & "h,d." & ChrW(960) & ".B.hi (kN)")
    Commentaire ws.Range("D" & r), "Couche prise comme couche d'ancrage (« Données du projet », " & _
        "tableau des couches) ; terrain neutralisé : pas de frottement, donc pas de contrainte"
    Commentaire ws.Range("H" & r), "Composante horizontale du frottement axial caractéristique qs " & _
        "mobilisé le long du micropieu incliné (0 si le micropieu est vertical)"
    Commentaire ws.Range("I" & r), "Composante horizontale du frottement axial de calcul qs,i = qs / (" & _
        ChrW(947) & "R;d1 x " & ChrW(947) & "R;d2)"
    For i = 1 To nbC
        r = rA + 1 + i
        ws.Range("A" & r).Value = i
        If Not IsError(hF) Then
            x = CStr(CLng(hF) + i)            ' ligne de la couche dans le frottement axial
            ws.Range("B" & r).Formula = "=" & CALCUL & "$C$" & x & "&"""""
            ws.Range("C" & r).Formula = "=N(" & CALCUL & "$F$" & x & ")"
            ws.Range("D" & r).Formula = "=IF(AND(" & DONNEES & "$N$" & LIGNE_COUCHES + i - 1 & "=""Oui""," & _
                CALCUL & "$Q$" & x & "<>""Oui""),""Oui"",""Non"")"
            ws.Range("E" & r).Formula = "=IF(D" & r & "=""Oui"",N(" & CALCUL & "$M$" & x & "),""-"")"
            ws.Range("F" & r).Formula = "=IF(D" & r & "=""Oui"",N(" & CALCUL & "$O$" & x & "),""-"")"
        End If
        ws.Range("G" & r).Formula = "=" & vAng
        ws.Range("H" & r).Formula = "=IF(D" & r & "=""Oui"",E" & r & "*SIN(RADIANS(G" & r & ")),""-"")"
        ws.Range("I" & r).Formula = "=IF(D" & r & "=""Oui"",F" & r & "*SIN(RADIANS(G" & r & ")),""-"")"
        ws.Range("J" & r).Formula = "=IF(D" & r & "=""Oui"",H" & r & "*PI()*" & vB & "*C" & r & ",""-"")"
        ws.Range("K" & r).Formula = "=IF(D" & r & "=""Oui"",I" & r & "*PI()*" & vB & "*C" & r & ",""-"")"
    Next i
    r = rA + 2 + nbC
    ws.Range("A" & r).Value = "Total (couches d'ancrage)"
    ws.Range("A" & r & ":I" & r).Merge
    ws.Range("J" & r).Formula = "=SUM(J" & rA + 2 & ":J" & r - 1 & ")"
    ws.Range("K" & r).Formula = "=SUM(K" & rA + 2 & ":K" & r - 1 & ")"
    With ws.Range("A" & rA + 1 & ":K" & r)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A" & rA + 1 & ":K" & rA + 1)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
    End With
    ws.Range("B" & rA + 2 & ":K" & r).Interior.Color = JAUNE_CALCUL
    ws.Range("C" & rA + 2 & ":C" & r).NumberFormat = "0.00"
    ws.Range("E" & rA + 2 & ":K" & r).NumberFormat = "0.0"
    ws.Range("A" & r & ":K" & r).Font.Bold = True
    rF = r                                    ' dernière ligne écrite (pour la suite)

    ' --- Efforts en tête (descente de charges seulement) et constantes de la solution ---
    ' Symboles et formules écrits dans le tableau ; tête libre ou encastrée selon
    ' « Données du projet » (conditions en tête de la fondation).
    etape = "efforts en tête et constantes"
    rT = rF + 3
    With ws.Range("A" & rT)
        .Value = "EFFORTS EN TÊTE (DESCENTE DE CHARGES) ET CONSTANTES DE LA SOLUTION (PIEU LONG, " & _
            "POINTE ENCASTRÉE)"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
    ws.Range("B" & rT + 1).Formula = "=""Micropieu libre en tête : ""&" & vCond & "&IF(" & vCond & _
        "=""Oui"","" (tête libre)"","" (tête encastrée)"")"
    r = rT + 2
    ws.Range("B" & r & ":I" & r).Value = Array("Grandeur", "Symbole", "Formule", "ELS CARA", "ELS QP", _
        "ELU FOND", "ELU ACC", "Unité")
    entete = Array( _
        Array("Effort de compression de la descente de charges", "Nc", _
            "Vmax de la combinaison (vérification de la profondeur, « Calcul et vérification »)", "kN"), _
        Array("Effort horizontal dû à l'inclinaison", "T0", "T0 = Nc x sin " & ChrW(945) & " (0 si vertical)", "kN"), _
        Array("Effort transversal en tête", "H", "H = T0", "kN"), _
        Array("Moment en tête (selon « Micropieu libre en tête »)", "M0", "libre en tête = Oui : M0 = 0 (aucun moment de la " & _
            "descente de charges) ; libre en tête = Non (encastrée) : M0 = -H.l0/2 = -H/(2" & ChrW(955) & ") (moment d'encastrement)", _
            "kN.m"), _
        Array("Constante de l'onde issue de la tête", "A", "libre en tête = Oui : A = 2" & ChrW(955) & "H/Es + 2" & _
            ChrW(955) & "²M0/Es ; libre en tête = Non (encastrée) : A = " & ChrW(955) & "H/Es", "m"), _
        Array("Constante de l'onde issue de la tête", "B", "libre en tête = Oui : B = -2" & ChrW(955) & "²M0/Es ; " & _
            "libre en tête = Non (encastrée) : B = A", "m"), _
        Array("Constante de l'onde issue de la pointe", "C", "C = -e^(-" & ChrW(955) & "D) [A cos " & ChrW(955) & _
            "D + B sin " & ChrW(955) & "D]", "m"), _
        Array("Constante de l'onde issue de la pointe", "E", "E = e^(-" & ChrW(955) & "D) [(B - 2A) cos " & _
            ChrW(955) & "D - (A + 2B) sin " & ChrW(955) & "D]", "m"))
    For j = 0 To 7
        ws.Range("B" & r + 1 + j).Value = entete(j)(0)
        ws.Range("C" & r + 1 + j).Value = entete(j)(1)
        ws.Range("D" & r + 1 + j).Value = entete(j)(2)
        ws.Range("I" & r + 1 + j).Value = entete(j)(3)
    Next j
    For k = 0 To 3
        cl = Chr(Asc("E") + k)
        x = cl & "$"                          ' colonne de la combinaison (ligne absolue)
        ws.Range(cl & r + 1).Formula = "=N(" & CALCUL & wc.Cells(ci, cC + 2 * k).Address & ")"
        ws.Range(cl & r + 2).Formula = "=IF(" & vInc & "=""Oui"",ABS(" & x & r + 1 & ")*SIN(RADIANS(" & vAng & ")),0)"
        ws.Range(cl & r + 3).Formula = "=" & x & r + 2
        ' M0 : 0 en tête libre ; moment d'encastrement -H / (2 lambda) en tête encastrée
        ws.Range(cl & r + 4).Formula = "=IF(" & vCond & "=""Non"",IF(N(" & vLam & ")>0,-" & x & r + 3 & _
            "/(2*" & vLam & "),0),0)"
        ' constantes A, B (tête), C, E (pointe encastrée) ; Es en kPa
        cond0 = "=IF(OR(" & vLam & "="""",N(" & vEs & ")<=0),"""","
        ws.Range(cl & r + 5).Formula = cond0 & "IF(" & vCond & "=""Non""," & vLam & "*" & x & r + 3 & _
            "/(" & vEs & "*1000),2*" & vLam & "*" & x & r + 3 & "/(" & vEs & "*1000)+2*" & vLam & "^2*" & x & r + 4 & _
            "/(" & vEs & "*1000)))"
        ws.Range(cl & r + 6).Formula = cond0 & "IF(" & vCond & "=""Non""," & x & r + 5 & ",-2*" & vLam & _
            "^2*" & x & r + 4 & "/(" & vEs & "*1000)))"
        ws.Range(cl & r + 7).Formula = "=IF(" & x & r + 5 & "="""","""",-EXP(-" & vx & ")*(" & x & r + 5 & _
            "*COS(" & vx & ")+" & x & r + 6 & "*SIN(" & vx & ")))"
        ws.Range(cl & r + 8).Formula = "=IF(" & x & r + 5 & "="""","""",EXP(-" & vx & ")*((" & x & r + 6 & "-2*" & _
            x & r + 5 & ")*COS(" & vx & ")-(" & x & r + 5 & "+2*" & x & r + 6 & ")*SIN(" & vx & ")))"
    Next k
    With ws.Range("B" & r & ":I" & r + 8)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("B" & r + 1 & ":B" & r + 8 & ",D" & r + 1 & ":D" & r + 8).HorizontalAlignment = xlLeft
    With ws.Range("B" & r & ":I" & r)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
    End With
    ws.Range("C" & r + 1 & ":C" & r + 8).Font.Bold = True
    ws.Range("E" & r + 1 & ":H" & r + 8).Interior.Color = JAUNE_CALCUL
    ws.Range("E" & r + 1 & ":H" & r + 4).NumberFormat = "0.0"
    ws.Range("E" & r + 5 & ":H" & r + 8).NumberFormat = "0.000E+00"
    rK = r                                    ' en-tête : A, B, C, E en lignes rK + 5 à rK + 8

    ' --- Y, T et M le long du micropieu (z de 0 à la profondeur réelle D) ---
    ' z = 0, dz, 2 dz, ... (pas de calcul dz saisi) jusqu'à D ; pas = z - z précédent ;
    '   Y = e^(-lambda z)[A cos + B sin](lambda z) + e^(-lambda u)[C cos + E sin](lambda u)
    '   M = 2 EI lambda² {e^(-lambda z)[A sin - B cos] + e^(-lambda u)[C sin - E cos]}
    '   T = 2 EI lambda³ {e^(-lambda z)[(A + B) cos + (B - A) sin]
    '                   - e^(-lambda u)[(C + E) cos + (E - C) sin]}, u = D - z
    etape = "Y, T et M le long du micropieu"
    rP = rK + 8 + 5                           ' en-tête du profil
    ' Pas de calcul dz (saisi, conservé) : z = 0, dz, 2 dz, ... jusqu'à D
    ws.Range("B" & rP - 2).Value = LIB_PAS
    ws.Range("B" & rP - 2 & ":C" & rP - 2).Merge
    With ws.Range("D" & rP - 2)
        .Value = IIf(N0(pasZ) > 0, N0(pasZ), 0.25)
        .Interior.Color = VERT_SAISIE
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("E" & rP - 2).Value = "m"
    ws.Range("B" & rP - 2 & ":E" & rP - 2).Borders.LineStyle = xlContinuous
    vPas = "$D$" & rP - 2
    ' Nombre de lignes : D / dz arrondi au-dessus (+ la tête), à la création de la feuille
    ws.Calculate
    nZ = 1
    If N0(ws.Range(vD).Value) > 0 And N0(ws.Range(vPas).Value) > 0 Then _
        nZ = Application.RoundUp(N0(ws.Range(vD).Value) / N0(ws.Range(vPas).Value) - 0.000000001, 0)
    If nZ < 1 Then nZ = 1
    If nZ > 2000 Then nZ = 2000
    With ws.Range("A" & rP - 3)
        .Value = "DÉPLACEMENT Y, EFFORT TRANCHANT T ET MOMENT M LE LONG DU MICROPIEU MP" & iMax
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
    ws.Range("A" & rP & ":D" & rP).Value = Array("z (m)", "Pas de calcul (m)", ChrW(955) & "z", _
        ChrW(955) & "u = " & ChrW(955) & "(D - z)")
    For k = 0 To 3
        ws.Range(ws.Cells(rP, 1 + k), ws.Cells(rP + 1, 1 + k)).Merge
        ws.Cells(rP, 5 + 3 * k).Value = Choose(k + 1, "ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
        ws.Range(ws.Cells(rP, 5 + 3 * k), ws.Cells(rP, 7 + 3 * k)).Merge
        ws.Range(ws.Cells(rP + 1, 5 + 3 * k), ws.Cells(rP + 1, 7 + 3 * k)).Value = _
            Array("Y (mm)", "T (kN)", "M (kN.m)")
    Next k
    For j = 0 To nZ
        r = rP + 2 + j
        xs = "$A" & r
        ' z : 0, puis z précédent + dz, limité à D
        If j = 0 Then
            ws.Range("A" & r).Formula = "=IF(N(" & vD & ")>0,0,"""")"
            ws.Range("B" & r).Formula = "=IF(A" & r & "="""","""",0)"
        Else
            ws.Range("A" & r).Formula = "=IF(A" & r - 1 & "="""","""",IF(A" & r - 1 & ">=" & vD & "-1E-9,""""," & _
                "MIN(A" & r - 1 & "+" & vPas & "," & vD & ")))"
            ws.Range("B" & r).Formula = "=IF(A" & r & "="""","""",A" & r & "-A" & r - 1 & ")"
        End If
        ws.Range("C" & r).Formula = "=IF(OR(" & xs & "=""""," & vLam & "=""""),""""," & vLam & "*" & xs & ")"
        ws.Range("D" & r).Formula = "=IF(OR(" & xs & "=""""," & vLam & "=""""),""""," & vLam & "*(" & vD & "-" & xs & "))"
        kA = "$C" & r: kB = "$D" & r          ' lambda z, lambda u
        For k = 0 To 3
            cl = Chr(Asc("E") + k)
            y0s = cl & "$" & rK + 5: hs = cl & "$" & rK + 6     ' A, B
            ths = cl & "$" & rK + 7: mhs = cl & "$" & rK + 8    ' C, E
            cond0 = "=IF(OR(" & kA & "=""""," & y0s & "=""""),"""","
            ws.Cells(r, 5 + 3 * k).Formula = cond0 & "1000*(EXP(-" & kA & ")*(" & y0s & "*COS(" & kA & ")+" & hs & _
                "*SIN(" & kA & "))+EXP(-" & kB & ")*(" & ths & "*COS(" & kB & ")+" & mhs & "*SIN(" & kB & "))))"
            ws.Cells(r, 6 + 3 * k).Formula = cond0 & "2*" & vEI & "*" & vLam & "^3*(EXP(-" & kA & ")*((" & y0s & "+" & _
                hs & ")*COS(" & kA & ")+(" & hs & "-" & y0s & ")*SIN(" & kA & "))-EXP(-" & kB & ")*((" & ths & "+" & _
                mhs & ")*COS(" & kB & ")+(" & mhs & "-" & ths & ")*SIN(" & kB & "))))"
            ws.Cells(r, 7 + 3 * k).Formula = cond0 & "2*" & vEI & "*" & vLam & "^2*(EXP(-" & kA & ")*(" & y0s & _
                "*SIN(" & kA & ")-" & hs & "*COS(" & kA & "))+EXP(-" & kB & ")*(" & ths & "*SIN(" & kB & ")-" & _
                mhs & "*COS(" & kB & "))))"
        Next k
    Next j
    With ws.Range("A" & rP & ":P" & rP + 2 + nZ)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A" & rP & ":P" & rP + 1)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .WrapText = True
    End With
    ws.Range("A" & rP + 2 & ":P" & rP + 2 + nZ).Interior.Color = JAUNE_CALCUL
    ws.Range("A" & rP + 2 & ":D" & rP + 2 + nZ).NumberFormat = "0.000"
    ws.Range("E" & rP + 2 & ":P" & rP + 2 + nZ).NumberFormat = "0.00"
    ws.Range("A" & rP + 2 & ":A" & rP + 2 + nZ).Font.Bold = True
    rG = rP                                   ' graphiques à droite du profil

MiseEnForme:
    etape = "mise en forme"
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:Y").ColumnWidth = 12
    ' Graphiques Y, T et M le long du micropieu le plus long, à droite de son profil
    If rG > 0 Then
        etape = "graphiques Y, T et M"
        ws.Calculate
        GraphiquesProfil ws, iMax, rP + 2, nZ + 1, rG, 18
    End If
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

' Valeur saisie (colonne D) de la ligne d'intitulé lib (colonne B) de cette feuille,
' conservée quand la feuille est recréée ; vide si absente
Private Function LireValeur(ws As Worksheet, lib As String) As Variant
    Dim v As Variant
    v = Application.Match(lib, ws.Range("B1:B300"), 0)
    If IsError(v) Then Exit Function
    LireValeur = ws.Range("D" & CLng(v)).Value
End Function

' Graphiques Y (mm), T (kN) et M (kN.m) le long du micropieu n° i en fonction de z
' (croissante vers le bas), à partir de la cellule (rH, cH) : une courbe par combinaison ;
' profil : z en colonne A à partir de la ligne r1 (nb lignes au plus), Y, T, M de la
' combinaison k en colonnes 5 + 3 k, 6 + 3 k, 7 + 3 k ; lignes renseignées seulement
Private Sub GraphiquesProfil(ws As Worksheet, i As Long, r1 As Long, nb As Long, rH As Long, cH As Long)
    Dim co As ChartObject, g As Long, k As Long, n As Long, lg As Double, rz As Range
    Dim noms As Variant, titres As Variant
    Do While n < nb                           ' lignes renseignées (z numérique)
        If IsEmpty(ws.Cells(r1 + n, 1).Value) Or Not IsNumeric(ws.Cells(r1 + n, 1).Value) Then Exit Do
        If CStr(ws.Cells(r1 + n, 1).Value) = "" Then Exit Do
        n = n + 1
    Loop
    If n < 2 Then Exit Sub
    Set rz = ws.Range(ws.Cells(r1, 1), ws.Cells(r1 + n - 1, 1))
    noms = Array("ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
    titres = Array("Déplacement Y (mm)", "Effort tranchant T (kN)", "Moment M (kN.m)")
    lg = 380
    For g = 0 To 2
        With ws.Cells(rH, cH)
            Set co = ws.ChartObjects.Add(.Left + g * (lg + 10), .Top, lg, 360)
        End With
        co.Placement = xlFreeFloating
        With co.Chart
            .ChartType = xlXYScatterSmoothNoMarkers
            Do While .SeriesCollection.Count > 0
                .SeriesCollection(1).Delete
            Loop
            For k = 0 To 3
                With .SeriesCollection.NewSeries
                    .Name = noms(k)
                    .XValues = ws.Range(ws.Cells(r1, 5 + 3 * k + g), ws.Cells(r1 + n - 1, 5 + 3 * k + g))
                    .Values = rz
                    .ChartType = xlXYScatterSmoothNoMarkers
                    .Format.Line.Weight = 2
                End With
            Next k
            .HasTitle = True
            .ChartTitle.Text = "MP" & i & " - " & titres(g) & " le long du micropieu"
            .HasLegend = True
            .Legend.Position = xlLegendPositionBottom
            With .Axes(xlValue)                ' profondeur, croissante vers le bas
                .ReversePlotOrder = True
                .MinimumScale = 0
                .HasTitle = True
                .AxisTitle.Text = "Profondeur z (m)"
            End With
            With .Axes(xlCategory)
                .HasTitle = True
                .AxisTitle.Text = titres(g)
            End With
        End With
    Next g
End Sub

' Valeur numérique (0 si vide ou texte)
Private Function N0(v As Variant) As Double
    If Not IsEmpty(v) And IsNumeric(v) Then N0 = CDbl(v)
End Function

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
