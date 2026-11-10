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
'         profondeur réelle D, x = D / l0, inclinaison, condition en tête (saisie),
'         longueur de mur reprise, poussée des terres au repos H = K0 gamma h² / 2,
'         e = h / 3
'   dessous : efforts en tête par combinaison (Nc, V = gamma G,sup H L, M = V e,
'         T0 = Nc sin(angle), H = V + T0) et constantes A, B, C, E de la solution du
'         pieu long à pointe encastrée ;
'         profil le long du micropieu : z (0, pas de calcul de « Calcul et
'         vérification » jusqu'à D, puis D), pas de calcul, Y, T et M de chaque
'         combinaison ; graphiques Y, T, M à la fin de la feuille
' Le nombre de micropieux et les adresses des autres feuilles sont lus à la création :
' relancer CreerVerificationStructurelle après un changement de ces feuilles.

Private Const NOM_FEUILLE As String = "Vérification structurelle"
Private Const DONNEES As String = "'Données du projet'!"
' Intitulés des saisies conservées quand la feuille est recréée
Private Const LIB_AUTRE As String = "Autre diamètre (à saisir, vide : " & "Phi retenu)"
Private Const LIB_TETE As String = "Condition en tête (libre ou encastrée)"
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
    Dim vEx As String

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
    condTete = LireValeur(ws, LIB_TETE)        ' condition en tête
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
    ' Poussée des terres au repos sur la fondation existante : h = hauteur du poteau du
    ' soubassement + hauteur de la semelle (descente de charges, colonne F), gamma et phi'
    ' de la couche 1, K0 = 1 - sin phi', H = K0 gamma h² / 2 (kN/ml), e = h / 3
    If Not wd Is Nothing Then
        ' Lignes du soubassement dans la descente de charges (mur 1) : poteau et semelle
        For r = 5 To 79
            If CStr(wd.Range("B" & r).Value) = "Soubassement" Then
                For i = r + 1 To r + 6
                    If Left(CStr(wd.Range("B" & i).Value), 6) = "Poteau" And rPo = 0 Then rPo = i
                    If Left(CStr(wd.Range("B" & i).Value), 7) = "Semelle" And rS = 0 Then rS = i
                Next i
                Exit For
            End If
        Next r
        dn = "'" & wd.Name & "'!"
    End If
    hT = "(" & IIf(rPo > 0, "N(" & dn & "$F$" & rPo & ")", "0") & "+" & _
        IIf(rS > 0, "N(" & dn & "$F$" & rS & ")", "0") & ")"
    k0 = "(1-SIN(RADIANS(N(" & DONNEES & "$L$" & LIGNE_COUCHES & "))))"
    hH = "(" & k0 & "*N(" & DONNEES & "$J$" & LIGNE_COUCHES & ")*" & hT & "^2/2)"
    eE = "(" & hT & "/3)"

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
    If m = 0 Or cC = 0 Or wd Is Nothing Then
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
    r = r + 1: Ligne ws, r, LIB_TETE, "", "": vCond = "$D$" & r
    With ws.Range("D" & r)
        .Validation.Delete
        .Validation.Add xlValidateList, xlValidAlertStop, , "Libre en tête,Encastré en tête"
        .Value = IIf(CStr(ActiveWorkbook.Worksheets("Données du projet").Range("D" & lp + 18).Value) = _
            "Encastré en tête", "Encastré en tête", "Libre en tête")
        If CStr(condTete) = "Libre en tête" Or CStr(condTete) = "Encastré en tête" Then .Value = condTete
        .Interior.Color = VERT_SAISIE
    End With
    r = r + 1: Ligne ws, r, "Longueur de mur reprise L (matrice des micropieux)", "=SUM(" & dn & _
        wd.Range(wd.Cells(rd + iMax - 1, 3), wd.Cells(rd + iMax - 1, 2 + n)).Address & ")", "m": vLm = "$D$" & r
    r = r + 1: Ligne ws, r, "Hauteur de terre h = hp + hs (soubassement)", "=" & hT, "m"
    r = r + 1: Ligne ws, r, "Coefficient des terres au repos K0 = 1 - sin " & ChrW(966) & "'", "=" & k0, ""
    r = r + 1: Ligne ws, r, "Poussée des terres au repos H = K0 " & ChrW(947) & " h² / 2", "=" & hH, "kN/ml": vHp = "$D$" & r
    r = r + 1: Ligne ws, r, "Excentricité e = h / 3", "=" & eE, "m": vEx = "$D$" & r
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

    ' --- Efforts en tête et constantes de la solution, par combinaison ---
    ' Nc = Vmax (vérification de la profondeur) ; V = gamma G,sup x H x L ; M = V x e ;
    ' T0 = Nc x sin(angle) si incliné ; H = V + T0. Pieu long, pointe encastrée :
    '   tête libre    : A = 2 lambda H / Es + 2 lambda² M / Es, B = -2 lambda² M / Es ;
    '   tête encastrée : A = B = lambda H / Es ;
    '   C = -e^(-lambda D) [A cos lambda D + B sin lambda D] ;
    '   E = e^(-lambda D) [(B - 2A) cos lambda D - (A + 2B) sin lambda D].
    etape = "efforts en tête et constantes"
    rT = rF + 3
    With ws.Range("A" & rT)
        .Value = "EFFORTS EN TÊTE ET CONSTANTES DE LA SOLUTION (PIEU LONG, POINTE ENCASTRÉE)"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
    r = rT + 1
    ws.Range("B" & r).Value = "Grandeur"
    ws.Range("B" & r & ":C" & r).Merge
    ws.Range("D" & r & ":G" & r).Value = Array("ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
    entete = Array(ChrW(947) & "G,sup de la combinaison", "Effort de compression Nc (Vmax)", _
        "Effort tranchant de poussée V = " & ChrW(947) & "G,sup x H x L", "Moment en tête M = V x e", _
        "Effort horizontal T0 = Nc x sin " & ChrW(945) & " (micropieu incliné)", "Effort transversal en tête H = V + T0", _
        "Constante A", "Constante B", "Constante C", "Constante E")
    unites = Array("", "kN", "kN", "kN.m", "kN", "kN", "m", "m", "m", "m")
    For j = 0 To 9
        ws.Range("B" & r + 1 + j).Value = entete(j)
        ws.Range("B" & r + 1 + j & ":C" & r + 1 + j).Merge
        ws.Range("H" & r + 1 + j).Value = unites(j)
    Next j
    For k = 0 To 3
        cl = Chr(Asc("D") + k)
        x = cl & "$"                          ' colonne de la combinaison (ligne absolue)
        If tC > 0 Then
            ws.Range(cl & r + 1).Formula = "=N(" & CALCUL & "$D$" & tC + 2 + k & ")"
        Else
            ws.Range(cl & r + 1).Value = Choose(k + 1, 1, 1, 1.35, 1)
        End If
        ws.Range(cl & r + 2).Formula = "=N(" & CALCUL & wc.Cells(ci, cC + 2 * k).Address & ")"
        ws.Range(cl & r + 3).Formula = "=" & x & r + 1 & "*" & vHp & "*" & vLm
        ws.Range(cl & r + 4).Formula = "=" & x & r + 3 & "*" & vEx
        ws.Range(cl & r + 5).Formula = "=IF(" & vInc & "=""Oui"",ABS(" & x & r + 2 & ")*SIN(RADIANS(" & vAng & ")),0)"
        ws.Range(cl & r + 6).Formula = "=" & x & r + 3 & "+" & x & r + 5
        ' constantes A, B (tête), C, E (pointe encastrée) ; Es en kPa
        cond0 = "=IF(OR(" & vLam & "="""",N(" & vEs & ")<=0),"""","
        ws.Range(cl & r + 7).Formula = cond0 & "IF(" & vCond & "=""Encastré en tête""," & vLam & "*" & x & r + 6 & _
            "/(" & vEs & "*1000),2*" & vLam & "*" & x & r + 6 & "/(" & vEs & "*1000)+2*" & vLam & "^2*" & x & r + 4 & _
            "/(" & vEs & "*1000)))"
        ws.Range(cl & r + 8).Formula = cond0 & "IF(" & vCond & "=""Encastré en tête""," & x & r + 7 & ",-2*" & vLam & _
            "^2*" & x & r + 4 & "/(" & vEs & "*1000)))"
        ws.Range(cl & r + 9).Formula = "=IF(" & x & r + 7 & "="""","""",-EXP(-" & vx & ")*(" & x & r + 7 & _
            "*COS(" & vx & ")+" & x & r + 8 & "*SIN(" & vx & ")))"
        ws.Range(cl & r + 10).Formula = "=IF(" & x & r + 7 & "="""","""",EXP(-" & vx & ")*((" & x & r + 8 & "-2*" & _
            x & r + 7 & ")*COS(" & vx & ")-(" & x & r + 7 & "+2*" & x & r + 8 & ")*SIN(" & vx & ")))"
    Next k
    With ws.Range("B" & r & ":H" & r + 10)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("B" & r + 1 & ":B" & r + 10).HorizontalAlignment = xlLeft
    With ws.Range("B" & r & ":H" & r)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
    End With
    ws.Range("D" & r + 1 & ":G" & r + 10).Interior.Color = JAUNE_CALCUL
    ws.Range("D" & r + 1 & ":G" & r + 1).NumberFormat = "0.00"
    ws.Range("D" & r + 2 & ":G" & r + 6).NumberFormat = "0.0"
    ws.Range("D" & r + 7 & ":G" & r + 10).NumberFormat = "0.000E+00"
    rK = r                                    ' en-tête du tableau des constantes

    ' --- Y, T et M le long du micropieu (z de 0 à la profondeur réelle D) ---
    ' z = 0, profondeurs du tableau de la capacité portante par pas de calcul (« Calcul et
    ' vérification ») inférieures à D, puis D ; pas de calcul = z - z précédent ;
    '   Y = e^(-lambda z)[A cos + B sin](lambda z) + e^(-lambda u)[C cos + E sin](lambda u)
    '   M = 2 EI lambda² {e^(-lambda z)[A sin - B cos] + e^(-lambda u)[C sin - E cos]}
    '   T = 2 EI lambda³ {e^(-lambda z)[(A + B) cos + (B - A) sin]
    '                   - e^(-lambda u)[(C + E) cos + (E - C) sin]}, u = D - z
    etape = "Y, T et M le long du micropieu"
    hZ = Application.Match("Profondeur z (m)", wc.Range("C1:C200"), 0)
    nZ = 0
    If Not IsError(hZ) Then
        If IsNumeric(wc.Range("ZA" & CLng(hZ) - 2).Value) Then nZ = CLng(N0(wc.Range("ZA" & CLng(hZ) - 2).Value))
        If nZ <= 0 Then
            Do While IsNumeric(wc.Range("C" & CLng(hZ) + 1 + nZ).Value) And _
                Not IsEmpty(wc.Range("C" & CLng(hZ) + 1 + nZ).Value) And nZ < 5000
                nZ = nZ + 1
            Loop
        End If
    End If
    If nZ < 1 Then nZ = 1
    rP = rK + 10 + 4                          ' en-tête du profil
    With ws.Range("A" & rP - 1)
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
    For j = 0 To nZ + 1
        r = rP + 2 + j
        xs = "$A" & r
        ' z : 0, puis profondeurs du pas de calcul inférieures à D, puis D
        If j = 0 Then
            ws.Range("A" & r).Formula = "=IF(N(" & vD & ")>0,0,"""")"
            ws.Range("B" & r).Formula = "=IF(A" & r & "="""","""",0)"
        Else
            If j <= nZ And Not IsError(hZ) Then
                zk = "N(" & CALCUL & "$C$" & CLng(hZ) + j & ")"
            Else
                zk = "0"
            End If
            ws.Range("A" & r).Formula = "=IF(A" & r - 1 & "="""","""",IF(A" & r - 1 & ">=" & vD & "-1E-9,""""," & _
                "IF(AND(" & zk & ">A" & r - 1 & "," & zk & "<" & vD & "-1E-9)," & zk & "," & vD & ")))"
            ws.Range("B" & r).Formula = "=IF(A" & r & "="""","""",A" & r & "-A" & r - 1 & ")"
        End If
        ws.Range("C" & r).Formula = "=IF(OR(" & xs & "=""""," & vLam & "=""""),""""," & vLam & "*" & xs & ")"
        ws.Range("D" & r).Formula = "=IF(OR(" & xs & "=""""," & vLam & "=""""),""""," & vLam & "*(" & vD & "-" & xs & "))"
        kA = "$C" & r: kB = "$D" & r          ' lambda z, lambda u
        For k = 0 To 3
            cl = Chr(Asc("D") + k)
            y0s = cl & "$" & rK + 7: hs = cl & "$" & rK + 8     ' A, B
            ths = cl & "$" & rK + 9: mhs = cl & "$" & rK + 10   ' C, E
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
    With ws.Range("A" & rP & ":P" & rP + 3 + nZ)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A" & rP & ":P" & rP + 1)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .WrapText = True
    End With
    ws.Range("A" & rP + 2 & ":P" & rP + 3 + nZ).Interior.Color = JAUNE_CALCUL
    ws.Range("A" & rP + 2 & ":D" & rP + 3 + nZ).NumberFormat = "0.000"
    ws.Range("E" & rP + 2 & ":P" & rP + 3 + nZ).NumberFormat = "0.00"
    ws.Range("A" & rP + 2 & ":A" & rP + 3 + nZ).Font.Bold = True
    rG = rP + 3 + nZ + 3                      ' titre des graphiques

MiseEnForme:
    etape = "mise en forme"
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:Y").ColumnWidth = 12
    ' Graphiques Y, T et M le long du micropieu le plus long, à la fin de la feuille
    If rG > 0 Then
        etape = "graphiques Y, T et M"
        ws.Calculate
        With ws.Range("A" & rG)
            .Value = "GRAPHIQUES - DÉPLACEMENT Y, EFFORT TRANCHANT T ET MOMENT M LE LONG DU MICROPIEU MP" & iMax
            .Font.Bold = True
            .Font.Size = 14
            .Font.Color = VERT_FONCE
        End With
        GraphiquesProfil ws, iMax, rP + 2, nZ + 2, rG + 2
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
' (croissante vers le bas), à partir de la ligne rH : une courbe par combinaison ;
' profil : z en colonne A à partir de la ligne r1 (nb lignes au plus), Y, T, M de la
' combinaison k en colonnes 5 + 3 k, 6 + 3 k, 7 + 3 k ; lignes renseignées seulement
Private Sub GraphiquesProfil(ws As Worksheet, i As Long, r1 As Long, nb As Long, rH As Long)
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
        With ws.Cells(rH, 1)
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
