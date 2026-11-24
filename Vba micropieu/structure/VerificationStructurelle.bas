Attribute VB_Name = "VerificationStructurelle"
Option Explicit
' Couleurs de la charte graphique GEOSKILLS (bleu nuit, bleu océan, vert émeraude)
Private Const COUL_TITRE As Long = 5898259         ' RGB(19, 0, 90) : Bleu Nuit (charte GEOSKILLS)
Private Const COUL_ENTETE As Long = 15458502       ' RGB(198, 224, 235) : Bleu Océan éclairci
Private Const COUL_SAISIE As Long = 15201485       ' RGB(205, 244, 231) : Vert Émeraude éclairci
Private Const COUL_CALCUL As Long = 16315885       ' RGB(237, 245, 248) : Bleu Océan très clair
Private Const COUL_ALERTE As Long = 192            ' RGB(192, 0, 0) : messages « non vérifié »

' Module « Vérification structurelle » : crée (ou recrée) la feuille « Vérification
' structurelle » des micropieux, à partir des feuilles « Données du projet » (module
' DonneesProjet), « Calcul Portance » (module CalculFondation) et de la feuille de
' descente de charges (matrice des micropieux, module MatriceMicropieux). Toutes les
' valeurs sont des formules : elles suivent les saisies de ces feuilles.
'   A3  : tableau vertical du micropieu le plus long (plus grande profondeur réelle de
'         « Calcul Portance ») : diamètre du pieu foré (Phi retenu ou autre
'         valeur saisie), enrobage et protection par le coulis (c >= 5 cm, C > 500
'         kg/m3, E/C < 0,5 : pas de corrosion), rayon de corrosion, armature après
'         corrosion, I, EI, EM et alpha (ou qc et beta), rho, Kf selon la méthode de
'         calcul, Es selon le type de sollicitation, l0 = (4 EI / Es)^(1/4), lambda,
'         profondeur réelle D, x = D / l0, inclinaison et angle
'   dessous : butée des terres sur l'ancrage (h = poteau BA + semelle BA,
'         réparti sur les couches de « Données du projet ») : K0 = (1 + sin phi')/(1 - sin phi'),
'         sigma h = K0 sigma v + 2 c' racine(K0), Bp,i = (sigma h haut + bas) / 2 hi ;
'         efforts en tête de la descente de charges (Nc, H0 = Nc sin(alpha) + H talus,
'         M0 = M talus ou -H0/(2 lambda) si encastré) et constantes A, B, C, E du pieu long à pointe encastrée, avec leurs
'         symboles et formules ;
'         vérification de la résistance frontale de la fondation : H0 / Bp par
'         combinaison (Bp = butée totale x plus grande largeur de mur du micropieu) ;
'         si H0 / Bp < 1 pour toutes les combinaisons, profil Y, T, M et graphiques
'         non calculés ;
'         profil le long du micropieu par pas de calcul dz (saisi) : z, pas, Y, T et M
'         de chaque combinaison, graphiques Y, T, M sous le profil, puis vérification de
'         l'armature (module VerificationArmature) sous les graphiques
' Le nombre de micropieux et les adresses des autres feuilles sont lus à la création :
' relancer CreerVerificationStructurelle après un changement de ces feuilles.

Private Const NOM_FEUILLE As String = "Vérification structurelle"
Private Const DONNEES As String = "'Données du projet'!"
' Intitulés des saisies conservées quand la feuille est recréée
Private Const LIB_AUTRE As String = "Autre diamètre (à saisir, vide : " & "Phi retenu)"
Private Const LIB_PAS As String = "Pas de calcul dz (à saisir)"
Private Const LIB_LIBRE As String = "Condition en tête"
Private Const TETE_LIBRE As String = "Micropieu libre en tête"
Private Const TETE_ENC As String = "Micropieu encastré en tête"
Private Const CALCUL As String = "'Calcul Portance'!"

' Charte GEOSKILLS : saisies en vert émeraude clair, valeurs calculées en
' bleu très clair, en-têtes en bleu océan clair, titres en bleu nuit
Public Const TITRE_FRONTALE As String = "VÉRIFICATION DE LA RÉSISTANCE FRONTALE DE LA FONDATION"

Private mEnCours As Boolean                     ' création de la feuille en cours

' Recrée « Vérification structurelle » (si elle existe) sans changer de feuille active :
' appelée quand le nombre de micropieux change dans la descente de charges
' (ConstruireMatrice, module MatriceMicropieux)
Public Sub ActualiserVerificationStructurelle()
    Dim ws As Worksheet, actif As Object
    If mEnCours Then Exit Sub
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Set actif = ActiveSheet
    CreerVerificationStructurelle
    On Error Resume Next
    actif.Activate
    On Error GoTo 0
End Sub

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
    Dim vDur As String
    Dim vBut As String, rV As Long, calcY As Boolean, v As Variant
    Dim libre As Variant, cTal As Long, fH0 As String, fM0 As String

    Application.ScreenUpdating = False
    mEnCours = True
    On Error GoTo Erreur
    etape = "feuilles du projet"
    Application.Run "RenommerFeuilleCalcul"   ' ancienne feuille « Calcul et vérification »
    Set wc = ActiveWorkbook.Worksheets("Calcul Portance")
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
    libre = LireValeur(ws, LIB_LIBRE)          ' condition en tête (libre ou encastré)
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
        .Font.Color = COUL_TITRE
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
    ' Vérification de la profondeur (« Calcul Portance ») : en-tête « N° MICROPIEU »
    ' en colonne A, combinaisons (Vmax) dans la colonne « COMBINAISONS ... », coefficients
    ' des combinaisons sous le tableau
    ' (position variable : sous le tableau de la capacité portante par pas de calcul)
    ' Nombre de micropieux : celui de la descente de charges (K1) ; vérification de la
    ' profondeur de « Calcul Portance » (et matrice des micropieux) recréée si
    ' elle n'a pas ce nombre de micropieux
    If Not wd Is Nothing Then
        k = CLng(N0(wd.Range("K1").Value))   ' N0 : la variable vAl masque la fonction Val
        On Error Resume Next
        r0 = Application.Match("N° MICROPIEU", wc.Range("A1:A5000"), 0)
        On Error GoTo Erreur
        m = 0
        If r0 > 0 Then
            Do While Left(CStr(wc.Range("A" & r0 + 3 + m).Value), 2) = "MP"
                m = m + 1
            Loop
        End If
        If k >= 1 And k <> m Then
            etape = "mise à jour du nombre de micropieux (descente de charges)"
            Application.Run "ActualiserVerificationProfondeur"
            wc.Calculate
        End If
        m = 0: r0 = 0
        etape = "micropieux de la vérification de la profondeur"
    End If
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
        dn = "'" & wd.Name & "'!"
        ' Lignes du soubassement (descente de charges) : poteau et semelle (hauteurs en F)
        For r = 5 To 79
            If CStr(wd.Range("B" & r).Value) = "Soubassement" Then
                For i = r + 1 To r + 6
                    If Left(CStr(wd.Range("B" & i).Value), 6) = "Poteau" And rPo = 0 Then rPo = i
                    If Left(CStr(wd.Range("B" & i).Value), 7) = "Semelle" And rS = 0 Then rS = i
                Next i
                Exit For
            End If
        Next r
    End If
    If m = 0 Or cC = 0 Then
        ws.Range("A3").Value = "Créez la matrice des micropieux (descente de charges) " & _
            "et « Calcul Portance », puis relancez CreerVerificationStructurelle."
        GoTo MiseEnForme
    End If
    ' --- Micropieu le plus long : plus grande profondeur réelle de la vérification de la
    ' profondeur (« Calcul Portance ») ; colonnes repérées par leurs en-têtes ---
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
    ci = r0 + 2 + iMax                        ' ligne du micropieu dans « Calcul Portance »

    ' --- Tableau vertical : longueur de transfert l0, inclinaison, poussée des terres ---
    ' Intitulé en B:C, valeur en D, unité en E. Diamètre du pieu foré B = Phi retenu ou
    ' autre valeur saisie ; enrobage c = (B - d) / 2 ; protection par le coulis (acier en
    ' compression) si c >= 5 cm, dosage > 500 kg/m3 et E/C < 0,5 : pas de corrosion ;
    ' armature après corrosion (tube : face extérieure, barre : tout le pourtour) ;
    ' I et EI après corrosion ; Kf de courte durée selon la méthode de calcul
    ' (pressiométrique : 12 EM / [4/3 (B0/B)(2,65 B/B0)^alpha + alpha] si B > B0,
    ' 12 EM / [4/3 (2,65)^alpha + alpha] si B <= B0 ; pénétrométrique : beta qc) ;
    ' seul le Kf de la durée choisie dans « Données du projet » est calculé (longue
    ' durée : Kf / 2) ; Es = Kf ;
    ' l0 = (4 EI / Es)^(1/4).
    etape = "longueur de transfert"
    With ws.Range("A3")
        .Value = "LONGUEUR DE TRANSFERT l0 DU MICROPIEU LE PLUS LONG (MP" & iMax & ")"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COUL_TITRE
    End With
    r = 4
    r = r + 1: Ligne ws, r, "Micropieu le plus long (plus grande profondeur réelle)", "MP" & iMax, ""
    r = r + 1: Ligne ws, r, ChrW(934) & " retenu (vérification de la profondeur)", _
        IIf(cPhi > 0, "=IF(ISNUMBER(" & CALCUL & wc.Cells(ci, cPhi).Address & ")," & CALCUL & _
        wc.Cells(ci, cPhi).Address & ",N(" & CALCUL & "$D$19))", "=N(" & CALCUL & "$D$19)"), "m": vPhi = "$D$" & r
    r = r + 1: Ligne ws, r, LIB_AUTRE, "", "m": vAut = "$D$" & r
    If IsNumeric(autre) And Not IsEmpty(autre) Then ws.Range("D" & r).Value = autre
    ws.Range("D" & r).Interior.Color = COUL_SAISIE
    r = r + 1: Ligne ws, r, "Diamètre du pieu foré B", "=IF(N(" & vAut & ")>0," & vAut & "," & vPhi & ")", "m": vB = "$D$" & r
    r = r + 1: Ligne ws, r, "Section du pieu foré A = " & ChrW(960) & ".B²/4", "=PI()*" & vB & "^2/4", "m²"
    r = r + 1: Ligne ws, r, "Diamètre de l'armature d", "=" & dA, "mm": vdA = "$D$" & r
    r = r + 1: Ligne ws, r, "Diamètre intérieur di (tube)", "=IF(" & tube & "," & dI & ",0)", "mm": vdi = "$D$" & r
    r = r + 1: Ligne ws, r, "Enrobage c = (B - d) / 2", "=MAX(0,(1000*" & vB & "-" & vdA & ")/2)", "mm": vc = "$D$" & r
    r = r + 1: Ligne ws, r, "Protection par le coulis (c >= 50 mm, C > 500 kg/m3, E/C < 0,5, compression)", _
        "=IF(AND(" & vc & ">=50,N(" & DONNEES & "D" & lp + 28 & ")>500,N(" & DONNEES & "D" & lp + 27 & _
        ")>0,N(" & DONNEES & "D" & lp + 27 & ")<0.5," & DONNEES & "D" & lp + 12 & _
        "<>""Traction""),""Efficace"",""Non efficace"")", "": vPro = "$D$" & r
    r = r + 1: Ligne ws, r, "Rayon de corrosion (« Données du projet »)", "=N(" & DONNEES & "D" & a + 13 & ")", _
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
    ' Kf (annexe I) pour la seule durée des charges choisie dans « Données du projet » :
    ' courte durée : B <= B0 : 12 EM / [4/3 (2,65)^alpha + alpha] ;
    ' B > B0 : 12 EM / [4/3 (B0/B)(2,65 B/B0)^alpha + alpha] (rho = max(B/B0 ; 1)) ;
    ' pénétrométrique : beta qc ; longue durée : même expression divisée par 2
    r = r + 1: Ligne ws, r, "Type de sollicitation", "=" & sol, ""
    r = r + 1: Ligne ws, r, "Durée des charges pour le calcul de Kf (« Données du projet »)", _
        "=IF(" & DONNEES & "$D$" & lp + 23 & "=""Courte durée"",""Courte durée"",""Longue durée"")", "": _
        vDur = "$D$" & r
    x = "12 EM / [4/3 (B0/B)(2,65 B/B0)^" & ChrW(945) & " + " & ChrW(945) & "] (B > B0), 12 EM / [4/3 (2,65)^" & _
        ChrW(945) & " + " & ChrW(945) & "] (B " & ChrW(8804) & " B0)"
    r = r + 1: Ligne ws, r, "=IF(" & vDur & "=""Longue durée"",""Module de réaction Kf de longue durée = " & x & _
        " / 2"",""Module de réaction Kf de courte durée = " & x & """)", _
        "=IF(OR(N(" & vEM & ")<=0," & vRho & "=""""),"""",IF(" & pen & "," & vAl & "*" & vEM & _
        ",12*" & vEM & "/(4/3*(2.65*" & vRho & ")^" & vAl & "/" & vRho & "+" & vAl & "))/IF(" & vDur & _
        "=""Longue durée"",2,1))", "MPa": vKf = "$D$" & r
    r = r + 1: Ligne ws, r, "Module Es = Kf", "=" & vKf, "MPa": vEs = "$D$" & r
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
    ' « Données du projet » (conditions en tête de la fondation)
    r = r + 1: Ligne ws, r, LIB_LIBRE, "", "": vCond = "$D$" & r
    With ws.Range("D" & r)
        .Validation.Delete
        .Validation.Add xlValidateList, xlValidAlertStop, , TETE_LIBRE & "," & TETE_ENC
        .Value = IIf(CStr(ActiveWorkbook.Worksheets("Données du projet").Range("D" & lp + 18).Value) = _
            "Encastré en tête", TETE_ENC, TETE_LIBRE)
        If CStr(libre) = TETE_LIBRE Or CStr(libre) = TETE_ENC Then .Value = libre
        If CStr(libre) = "Oui" Then .Value = TETE_LIBRE   ' ancienne liste Oui / Non
        If CStr(libre) = "Non" Then .Value = TETE_ENC
    End With
    Commentaire ws.Range("B" & r), "Micropieu libre en tête : rotation libre, M0 = 0 ; micropieu " & _
        "encastré en tête : tête encastrée dans la semelle (rotation nulle, moment d'encastrement M0 = -H.l0/2). Valeur par défaut : " & _
        "« Données du projet », conditions en tête de la fondation"
    rF = r
    With ws.Range("B5:E" & rF)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("D5:E" & rF).HorizontalAlignment = xlCenter
    ws.Range("D5:D" & rF).Interior.Color = COUL_CALCUL
    ws.Range(vAut & "," & vCond).Interior.Color = COUL_SAISIE
    ws.Range("D5:D" & rF).NumberFormat = "0.000"
    ws.Range(vI & "," & vI & ":" & vI).NumberFormat = "0.000E+00"
    ws.Range("D" & Mid(vI, 4) - 1).NumberFormat = "0.000E+00"
    ws.Range("B" & Mid(vl0, 4) & ":D" & Mid(vl0, 4) & ",B" & Mid(vx, 4) + 1 & ":D" & Mid(vx, 4) + 1).Font.Bold = True
    With ws.Range(vPro).FormatConditions
        .Delete
        .Add(xlCellValue, xlEqual, "=""Non efficace""").Font.Color = COUL_ALERTE
    End With

    ' --- Butée des terres sur l'ancrage (poteau BA + semelle BA du
    ' soubassement) ---
    '   hauteur de l'ancrage : h = hp (poteau BA) + hs (semelle BA), descente de charges ;
    '   la hauteur h est répartie sur les couches du tableau de « Données du projet »
    '   (épaisseurs depuis le TN) : couche i entre zh = min(z haut ; h) et zb = min(z bas ; h) ;
    '   coefficient de poussée : K0 = (1 + sin phi') / (1 - sin phi') ;
    '   contrainte verticale : sigma v = somme(gamma j x hj) des couches au-dessus ;
    '   contrainte horizontale : sigma h = K0 x sigma v + 2 c' racine(K0) (haut et bas) ;
    '   effort par mètre linéaire : Bp,i = (sigma h haut + sigma h bas) / 2 x hi.
    etape = "butée des terres sur l'ancrage"
    nbC = CLng(N0(ActiveWorkbook.Worksheets("Données du projet").Range(CELLULE_NB_COUCHES).Value))
    If nbC < 1 Then nbC = 1
    If nbC > MAX_COUCHES Then nbC = MAX_COUCHES
    rA = rF + 3
    With ws.Range("A" & rA)
        .Value = "BUTÉE DES TERRES SUR L'ANCRAGE (POTEAU BA + SEMELLE BA)"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COUL_TITRE
    End With
    ' Hauteur de l'ancrage
    ws.Range("B" & rA + 1 & ":E" & rA + 1).Value = Array("Hauteur du poteau BA (soubassement)", "hp", "", "m")
    ws.Range("B" & rA + 2 & ":E" & rA + 2).Value = Array("Hauteur de la semelle BA (soubassement)", "hs", "", "m")
    ws.Range("B" & rA + 3 & ":E" & rA + 3).Value = Array("Hauteur de l'ancrage", "h", "", "m")
    ws.Range("D" & rA + 1).Formula = IIf(rPo > 0, "=N(" & dn & "$F$" & rPo & ")", "=0")
    ws.Range("D" & rA + 2).Formula = IIf(rS > 0, "=N(" & dn & "$F$" & rS & ")", "=0")
    ws.Range("D" & rA + 3).Formula = "=D" & rA + 1 & "+D" & rA + 2
    Commentaire ws.Range("C" & rA + 1), "Descente de charges, soubassement : hauteur (colonne F) de la ligne Poteau"
    Commentaire ws.Range("C" & rA + 2), "Descente de charges, soubassement : hauteur (colonne F) de la ligne Semelle"
    Commentaire ws.Range("C" & rA + 3), "h = hp + hs : hauteur de terrain en contact avec l'ancrage, " & _
        "répartie sur les couches de « Données du projet » à partir du TN"
    With ws.Range("B" & rA + 1 & ":E" & rA + 3)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With
    ws.Range("C" & rA + 1 & ":E" & rA + 3).HorizontalAlignment = xlCenter
    ws.Range("C" & rA + 1 & ":C" & rA + 3).Font.Bold = True
    ws.Range("D" & rA + 1 & ":D" & rA + 3).Interior.Color = COUL_CALCUL
    ws.Range("D" & rA + 1 & ":D" & rA + 3).NumberFormat = "0.00"
    x = "$D$" & rA + 3                        ' hauteur de l'ancrage
    ' Tableau par couche
    r = rA + 5
    ws.Range("A" & r & ":O" & r).Value = Array("N° de couche", "Nom de couche", "z haut (m)", "z bas (m)", _
        "Épaisseur dans l'ancrage hi (m)", ChrW(947) & " (kN/m3)", "c' (kPa)", ChrW(966) & "' (°)", _
        "K0 = (1 + sin " & ChrW(966) & "')/(1 - sin " & ChrW(966) & "')", ChrW(963) & "v haut (kPa)", ChrW(963) & "v bas (kPa)", _
        ChrW(963) & "h haut = K0." & ChrW(963) & "v + 2c'" & ChrW(8730) & "K0 (kPa)", _
        ChrW(963) & "h bas = K0." & ChrW(963) & "v + 2c'" & ChrW(8730) & "K0 (kPa)", _
        "Bp,i = (" & ChrW(963) & "h haut + " & ChrW(963) & "h bas)/2 . hi (kN/ml)", "Profondeur d'application (m)")
    Commentaire ws.Range("C" & r), "z haut = min(profondeur du haut de la couche depuis le TN ; h)"
    Commentaire ws.Range("D" & r), "z bas = min(profondeur du bas de la couche depuis le TN ; h) ; " & _
        "épaisseurs du tableau des couches de « Données du projet »"
    Commentaire ws.Range("I" & r), "K0 = (1 + sin " & ChrW(966) & "') / (1 - sin " & ChrW(966) & "')"
    Commentaire ws.Range("J" & r), ChrW(963) & "v haut = somme des " & ChrW(947) & "j.hj des couches au-dessus"
    Commentaire ws.Range("O" & r), "Centre de gravité du diagramme trapézoïdal de poussée de la couche : " & _
        "z haut + hi (" & ChrW(963) & "h haut + 2 " & ChrW(963) & "h bas) / (3 (" & ChrW(963) & "h haut + " & _
        ChrW(963) & "h bas))"
    For i = 1 To nbC
        r = rA + 5 + i
        k = LIGNE_COUCHES + i - 1             ' ligne de la couche dans « Données du projet »
        ws.Range("A" & r).Value = i
        ws.Range("B" & r).Formula = "=" & DONNEES & "$C$" & k & "&"""""
        If i = 1 Then
            ws.Range("C" & r).Formula = "=0"
            ws.Range("D" & r).Formula = "=MIN(N(" & DONNEES & "$F$" & k & ")," & x & ")"
            ws.Range("J" & r).Formula = "=0"
        Else
            ws.Range("C" & r).Formula = "=MIN(SUM(" & DONNEES & "$F$" & LIGNE_COUCHES & ":$F$" & k - 1 & ")," & x & ")"
            ws.Range("D" & r).Formula = "=MIN(SUM(" & DONNEES & "$F$" & LIGNE_COUCHES & ":$F$" & k & ")," & x & ")"
            ws.Range("J" & r).Formula = "=K" & r - 1
        End If
        ws.Range("E" & r).Formula = "=MAX(D" & r & "-C" & r & ",0)"
        ws.Range("F" & r).Formula = "=N(" & DONNEES & "$J$" & k & ")"
        ws.Range("G" & r).Formula = "=N(" & DONNEES & "$K$" & k & ")"
        ws.Range("H" & r).Formula = "=N(" & DONNEES & "$L$" & k & ")"
        ws.Range("I" & r).Formula = "=(1+SIN(RADIANS(H" & r & ")))/(1-SIN(RADIANS(H" & r & ")))"
        ws.Range("K" & r).Formula = "=J" & r & "+F" & r & "*E" & r
        ws.Range("L" & r).Formula = "=I" & r & "*J" & r & "+2*G" & r & "*SQRT(I" & r & ")"
        ws.Range("M" & r).Formula = "=I" & r & "*K" & r & "+2*G" & r & "*SQRT(I" & r & ")"
        ws.Range("N" & r).Formula = "=IF(E" & r & ">0,(L" & r & "+M" & r & ")/2*E" & r & ",0)"
        ws.Range("O" & r).Formula = "=IF(AND(E" & r & ">0,L" & r & "+M" & r & ">0),C" & r & "+E" & r & "*(L" & r & _
            "+2*M" & r & ")/(3*(L" & r & "+M" & r & ")),0)"
    Next i
    r = rA + 6 + nbC
    ws.Range("A" & r).Value = "Butée des terres totale sur l'ancrage " & ChrW(931) & "Bp,i et profondeur d'application"
    ws.Range("A" & r & ":M" & r).Merge
    ws.Range("N" & r).Formula = "=SUM(N" & rA + 6 & ":N" & r - 1 & ")"
    vBut = "$N$" & r                          ' butée totale (kN/ml)
    ws.Range("O" & r).Formula = "=IF(N" & r & ">0,SUMPRODUCT(N" & rA + 6 & ":N" & r - 1 & ",O" & rA + 6 & ":O" & _
        r - 1 & ")/N" & r & ",0)"
    With ws.Range("A" & rA + 5 & ":O" & r)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A" & rA + 5 & ":O" & rA + 5)
        .Font.Bold = True
        .Interior.Color = COUL_ENTETE
    End With
    ws.Range("B" & rA + 6 & ":O" & r).Interior.Color = COUL_CALCUL
    ws.Range("C" & rA + 6 & ":E" & r & ",O" & rA + 6 & ":O" & r).NumberFormat = "0.00"
    ws.Range("F" & rA + 6 & ":H" & r & ",J" & rA + 6 & ":N" & r).NumberFormat = "0.0"
    ws.Range("I" & rA + 6 & ":I" & r).NumberFormat = "0.000"
    ws.Range("A" & r & ":O" & r).Font.Bold = True
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
        .Font.Color = COUL_TITRE
    End With
    ws.Range("B" & rT + 1).Formula = "=""Condition en tête : ""&" & vCond
    r = rT + 2
    ws.Range("B" & r & ":I" & r).Value = Array("Grandeur", "Symbole", "Formule", "ELS CARA", "ELS QP", _
        "ELU FOND", "ELU ACC", "Unité")
    ' Charge du talus H0, M0 du micropieu (descente de charges, matrice des micropieux,
    ' colonnes « CHARGE DU TALUS ») : 0 si absente
    cTal = 0
    If Not wd Is Nothing Then
        For j = 3 To 80
            If CStr(wd.Cells(88 + n, j).Value) = "CHARGE DU TALUS" Then cTal = j: Exit For
        Next j
    End If
    If cTal > 0 Then
        fH0 = "N(" & dn & wd.Cells(rd + iMax - 1, cTal).Address & ")"
        fM0 = "N(" & dn & wd.Cells(rd + iMax - 1, cTal + 1).Address & ")"
    Else
        fH0 = "0": fM0 = "0"
    End If
    entete = Array( _
        Array("Effort de compression de la descente de charges", "Nc", _
            "Vmax de la combinaison (vérification de la profondeur, « Calcul Portance »)", "kN"), _
        Array("Effort horizontal en tête", "H0", "H0 = Nc x sin " & ChrW(945) & " (composante inclinée de la " & _
            "descente de charges, 0 si vertical) + H talus (descente de charges, charge du talus)", "kN"), _
        Array("Moment en tête (selon la condition en tête)", "M0", "micropieu libre en tête : M0 = M talus " & _
            "(descente de charges, charge du talus) ; micropieu encastré en tête : M0 = -H0.l0/2 = -H0/(2" & _
            ChrW(955) & ") (moment repris par la tête encastrée)", "kN.m"), _
        Array("Constante de l'onde issue de la tête", "A", "micropieu libre en tête : A = 2" & ChrW(955) & "H0/Es + 2" & _
            ChrW(955) & "²M0/Es ; micropieu encastré en tête : A = " & ChrW(955) & "H0/Es", "m"), _
        Array("Constante de l'onde issue de la tête", "B", "micropieu libre en tête : B = -2" & ChrW(955) & "²M0/Es ; " & _
            "micropieu encastré en tête : B = A", "m"), _
        Array("Constante de l'onde issue de la pointe", "C", "C = -e^(-" & ChrW(955) & "D) [A cos " & ChrW(955) & _
            "D + B sin " & ChrW(955) & "D]", "m"), _
        Array("Constante de l'onde issue de la pointe", "E", "E = e^(-" & ChrW(955) & "D) [(B - 2A) cos " & _
            ChrW(955) & "D - (A + 2B) sin " & ChrW(955) & "D]", "m"))
    For j = 0 To 6
        ws.Range("B" & r + 1 + j).Value = entete(j)(0)
        ws.Range("C" & r + 1 + j).Value = entete(j)(1)
        ws.Range("D" & r + 1 + j).Value = entete(j)(2)
        ws.Range("I" & r + 1 + j).Value = entete(j)(3)
    Next j
    Commentaire ws.Range("C" & r + 2), "H0 = Nc x sin " & ChrW(945) & " + H talus : composante horizontale de " & _
        "l'effort de la descente de charges (micropieu incliné) plus l'effort horizontal du talus (descente de " & _
        "charges, matrice des micropieux, colonne « CHARGE DU TALUS »)."
    Commentaire ws.Range("C" & r + 3), "Micropieu libre en tête : M0 = moment du talus (descente de charges, " & _
        "colonne « CHARGE DU TALUS »)." & vbLf & "Micropieu encastré en tête : M0 = -H0/(2" & ChrW(955) & _
        ") = -H0.l0/2, moment repris par la tête encastrée (le moment du talus est repris par l'encastrement)."
    For k = 0 To 3
        cl = Chr(Asc("E") + k)
        x = cl & "$"                          ' colonne de la combinaison (ligne absolue)
        ws.Range(cl & r + 1).Formula = "=N(" & CALCUL & wc.Cells(ci, cC + 2 * k).Address & ")"
        ' H0 = composante inclinée de la descente de charges + charge du talus
        ws.Range(cl & r + 2).Formula = "=IF(" & vInc & "=""Oui"",ABS(" & x & r + 1 & ")*SIN(RADIANS(" & vAng & _
            ")),0)+" & fH0
        ' M0 : moment du talus en tête libre ; moment repris par la tête encastrée
        ws.Range(cl & r + 3).Formula = "=IF(" & vCond & "=""" & TETE_ENC & """,IF(N(" & vLam & ")>0,-" & x & r + 2 & _
            "/(2*" & vLam & "),0)," & fM0 & ")"
        ' constantes A, B (tête), C, E (pointe encastrée) ; Es en kPa
        cond0 = "=IF(OR(" & vLam & "="""",N(" & vEs & ")<=0),"""","
        ws.Range(cl & r + 4).Formula = cond0 & "IF(" & vCond & "=""" & TETE_ENC & """," & vLam & "*" & x & r + 2 & _
            "/(" & vEs & "*1000),2*" & vLam & "*" & x & r + 2 & "/(" & vEs & "*1000)+2*" & vLam & "^2*" & x & r + 3 & _
            "/(" & vEs & "*1000)))"
        ws.Range(cl & r + 5).Formula = cond0 & "IF(" & vCond & "=""" & TETE_ENC & """," & x & r + 4 & ",-2*" & vLam & _
            "^2*" & x & r + 3 & "/(" & vEs & "*1000)))"
        ws.Range(cl & r + 6).Formula = "=IF(" & x & r + 4 & "="""","""",-EXP(-" & vx & ")*(" & x & r + 4 & _
            "*COS(" & vx & ")+" & x & r + 5 & "*SIN(" & vx & ")))"
        ws.Range(cl & r + 7).Formula = "=IF(" & x & r + 4 & "="""","""",EXP(-" & vx & ")*((" & x & r + 5 & "-2*" & _
            x & r + 4 & ")*COS(" & vx & ")-(" & x & r + 4 & "+2*" & x & r + 5 & ")*SIN(" & vx & ")))"
    Next k
    With ws.Range("B" & r & ":I" & r + 7)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("B" & r + 1 & ":B" & r + 7 & ",D" & r + 1 & ":D" & r + 7).HorizontalAlignment = xlLeft
    With ws.Range("B" & r & ":I" & r)
        .Font.Bold = True
        .Interior.Color = COUL_ENTETE
    End With
    ws.Range("C" & r + 1 & ":C" & r + 7).Font.Bold = True
    ws.Range("E" & r + 1 & ":H" & r + 7).Interior.Color = COUL_CALCUL
    ws.Range("E" & r + 1 & ":H" & r + 3).NumberFormat = "0.0"
    ws.Range("E" & r + 4 & ":H" & r + 7).NumberFormat = "0.000E+00"
    rK = r                                    ' en-tête : Nc, H0, M0 puis A, B, C, E en lignes rK + 4 à rK + 7

    ' --- Vérification de la résistance frontale de la fondation : effort horizontal en
    ' tête H0 de chaque combinaison comparé à la butée des terres reprise par l'ancrage
    ' du micropieu (butée totale par mètre linéaire x plus grande largeur de mur reprise
    ' par le micropieu). Rapport H0 / Bp < 1 pour toutes les combinaisons : la butée reprend
    ' l'effort horizontal, le calcul de Y, T et M du micropieu n'est pas nécessaire. ---
    etape = "résistance frontale de la fondation"
    rV = rK + 9
    With ws.Range("A" & rV)
        .Value = TITRE_FRONTALE & " (MP" & iMax & ")"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COUL_TITRE
    End With
    r = rV + 1
    ws.Range("B" & r & ":I" & r).Value = Array("Grandeur", "Symbole", "Formule", "ELS CARA", "ELS QP", _
        "ELU FOND", "ELU ACC", "Unité")
    ws.Range("B" & r + 1 & ":D" & r + 1).Value = Array("Butée des terres totale sur l'ancrage (par mètre linéaire)", _
        ChrW(931) & "Bp,i", "tableau « Butée des terres sur l'ancrage »")
    ws.Range("B" & r + 2 & ":D" & r + 2).Value = Array("Largeur de mur la plus grande reprise par le micropieu", _
        "L", "descente de charges : plus grande des longueurs de murs de la ligne du micropieu (matrice des " & _
        "micropieux)")
    ws.Range("B" & r + 3 & ":D" & r + 3).Value = Array("Butée des terres reprise par le micropieu", "Bp", _
        "Bp = " & ChrW(931) & "Bp,i x L")
    ws.Range("B" & r + 4 & ":D" & r + 4).Value = Array("Effort horizontal en tête", "H0", _
        "tableau des efforts en tête (H0 = Nc x sin " & ChrW(945) & " + H talus)")
    ws.Range("B" & r + 5 & ":D" & r + 5).Value = Array("Rapport effort horizontal / butée", "H0 / Bp", _
        "Condition : H0 / Bp < 1")
    ws.Range("B" & r + 6 & ":D" & r + 6).Value = Array("Vérification", "", _
        "H0 / Bp < 1 : la butée reprend l'effort horizontal ; sinon calcul de Y, T et M du micropieu")
    ws.Range("I" & r + 1 & ":I" & r + 5).Value = Application.Transpose(Array("kN/ml", "m", "kN", "kN", ""))
    For k = 0 To 3
        cl = Chr(Asc("E") + k)
        ws.Range(cl & r + 1).Formula = "=N(" & vBut & ")"
        If Not wd Is Nothing And n > 0 Then
            ws.Range(cl & r + 2).Formula = "=MAX(" & dn & wd.Range(wd.Cells(rd + iMax - 1, 3), _
                wd.Cells(rd + iMax - 1, 2 + n)).Address & ")"
        Else
            ws.Range(cl & r + 2).Formula = "=0"
        End If
        ws.Range(cl & r + 3).Formula = "=" & cl & r + 1 & "*" & cl & r + 2
        ws.Range(cl & r + 4).Formula = "=ABS(N(" & cl & "$" & rK + 2 & "))"
        ws.Range(cl & r + 5).Formula = "=IF(N(" & cl & r + 3 & ")>0," & cl & r + 4 & "/" & cl & r + 3 & _
            ",IF(N(" & cl & r + 4 & ")=0,0,""""))"
        ws.Range(cl & r + 6).Formula = "=IF(" & cl & r + 5 & "="""",""Butée nulle : calcul du micropieu""," & _
            "IF(" & cl & r + 5 & "<1,""OK : calcul non nécessaire"",""Butée insuffisante : calcul du micropieu""))"
    Next k
    With ws.Range("B" & r & ":I" & r + 6)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("B" & r + 1 & ":B" & r + 6 & ",D" & r + 1 & ":D" & r + 6).HorizontalAlignment = xlLeft
    With ws.Range("B" & r & ":I" & r)
        .Font.Bold = True
        .Interior.Color = COUL_ENTETE
    End With
    ws.Range("C" & r + 1 & ":C" & r + 6 & ",B" & r + 6 & ":H" & r + 6).Font.Bold = True
    ws.Range("E" & r + 1 & ":H" & r + 6).Interior.Color = COUL_CALCUL
    ws.Range("E" & r + 1 & ":H" & r + 4).NumberFormat = "0.0"
    ws.Range("E" & r + 5 & ":H" & r + 5).NumberFormat = "0.00"
    With ws.Range("E" & r + 6 & ":H" & r + 6).FormatConditions
        .Delete
        .Add(xlTextString, String:="calcul du micropieu", TextOperator:=xlContains).Font.Color = COUL_ALERTE
    End With
    ' Décision : calcul de Y, T et M seulement si une combinaison a H0 / Bp >= 1
    ws.Calculate
    calcY = False
    For k = 0 To 3
        v = ws.Cells(r + 5, 5 + k).Value
        If IsError(v) Then
            calcY = True
        ElseIf Not IsNumeric(v) Or IsEmpty(v) Or CStr(v) = "" Then
            calcY = True
        ElseIf CDbl(v) >= 1 Then
            calcY = True
        End If
    Next k
    rF = r + 6
    If Not calcY Then
        With ws.Range("B" & rF + 2 & ":I" & rF + 2)
            .Merge
            .Value = "H0 / Bp < 1 pour toutes les combinaisons : la butée des terres sur l'ancrage reprend " & _
                "l'effort horizontal en tête ; le calcul des déplacements Y, de l'effort tranchant T et du " & _
                "moment M du micropieu n'est pas nécessaire."
            .Font.Bold = True
            .Font.Color = COUL_TITRE
            .WrapText = True
        End With
        ws.Rows(rF + 2).RowHeight = 32
        GoTo MiseEnForme
    End If

    ' --- Y, T et M le long du micropieu (z de 0 à la profondeur réelle D) ---
    ' z = 0, dz, 2 dz, ... (pas de calcul dz saisi) jusqu'à D ; pas = z - z précédent ;
    '   Y = e^(-lambda z)[A cos + B sin](lambda z) + e^(-lambda u)[C cos + E sin](lambda u)
    '   M = 2 EI lambda² {e^(-lambda z)[A sin - B cos] + e^(-lambda u)[C sin - E cos]}
    '   T = 2 EI lambda³ {e^(-lambda z)[(A + B) cos + (B - A) sin]
    '                   - e^(-lambda u)[(C + E) cos + (E - C) sin]}, u = D - z
    etape = "Y, T et M le long du micropieu"
    rP = rF + 5                               ' en-tête du profil
    ' Pas de calcul dz (saisi, conservé) : z = 0, dz, 2 dz, ... jusqu'à D
    ws.Range("B" & rP - 2).Value = LIB_PAS
    ws.Range("B" & rP - 2 & ":C" & rP - 2).Merge
    With ws.Range("D" & rP - 2)
        .Value = IIf(N0(pasZ) > 0, N0(pasZ), 0.25)
        .Interior.Color = COUL_SAISIE
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
        .Font.Color = COUL_TITRE
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
            y0s = cl & "$" & rK + 4: hs = cl & "$" & rK + 5     ' A, B
            ths = cl & "$" & rK + 6: mhs = cl & "$" & rK + 7    ' C, E
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
        .Interior.Color = COUL_ENTETE
        .WrapText = True
    End With
    ws.Range("A" & rP + 2 & ":P" & rP + 2 + nZ).Interior.Color = COUL_CALCUL
    ws.Range("A" & rP + 2 & ":D" & rP + 2 + nZ).NumberFormat = "0.000"
    ws.Range("E" & rP + 2 & ":P" & rP + 2 + nZ).NumberFormat = "0.00"
    ws.Range("A" & rP + 2 & ":A" & rP + 2 + nZ).Font.Bold = True
    rG = rP + 2 + nZ + 3                      ' titre des graphiques, sous le profil

MiseEnForme:
    etape = "mise en forme"
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:Y").ColumnWidth = 12
    ' Graphiques Y, T et M le long du micropieu le plus long, sous son profil : trois
    ' graphiques côte à côte (une courbe par combinaison)
    If rG > 0 Then
        etape = "graphiques Y, T et M"
        ws.Calculate
        With ws.Range("A" & rG)
            .Value = "GRAPHIQUES - DÉPLACEMENT Y, EFFORT TRANCHANT T ET MOMENT M LE LONG DU MICROPIEU MP" & iMax
            .Font.Bold = True
            .Font.Size = 14
            .Font.Color = COUL_TITRE
        End With
        GraphiquesProfil ws, iMax, rP + 2, nZ + 1, rG + 2, 1
    End If
    ' Vérification de l'armature (module VerificationArmature), sous les graphiques (ou
    ' sous la vérification de la résistance frontale si le profil n'est pas calculé)
    ' puis flambement (Mandel) et platines, les uns sous les autres : les trois
    ' vérifications sont recréées à chaque création de la feuille (plus besoin de les
    ' relancer une à une) ; une vérification en échec est signalée sans arrêter les autres
    If rK > 0 Then
        LancerVerification "VerificationArmature", "vérification de l'armature"
        LancerVerification "VerificationMandel", "vérification au flambement (Mandel)"
        LancerVerification "VerificationPlatine", "dimensionnement des platines"
    End If
    ws.Cells.Font.Name = "CheGuevara Text Sans"   ' police principale de la charte graphique GEOSKILLS
    AjusterTextes ws                      ' colonnes élargies pour la police de la charte
    ' Nombre de micropieux de la descente de charges (K1) mémorisé, et code de la feuille
    ' installé : à l'activation de la feuille, recréation si ce nombre a changé
    If Not wd Is Nothing Then
        On Error Resume Next
        ws.Names.Add Name:="NbMicropieuxDDC", RefersTo:="=" & CLng(N0(wd.Range("K1").Value)), Visible:=False
        On Error GoTo Erreur
    End If
    InstallerCodeActivation ws
    ws.Activate
Fin:
    mEnCours = False
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
                    .Format.Line.ForeColor.RGB = Choose((k Mod 4) + 1, 5898259, 11371036, 8964355, 8139520)   ' palette de la charte
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
        .Font.Color = COUL_TITRE
    End With
End Sub

' Commentaire (signification) sur une cellule d'en-tête
' Vrai pendant la création de la feuille : les vérifications de l'armature, du
' flambement et des platines ne s'enchaînent alors pas elles-mêmes (elles sont lancées
' l'une après l'autre par CreerVerificationStructurelle)
Public Function CreationVerificationEnCours() As Boolean
    CreationVerificationEnCours = mEnCours
End Function

' Lance une vérification (module de la feuille) ; erreur signalée sans arrêter la suite
Private Sub LancerVerification(macro As String, libelle As String)
    On Error GoTo Echec
    Application.Run macro
    Exit Sub
Echec:
    MsgBox "« Vérification structurelle » : la " & libelle & " (macro " & macro & ") n'a pas pu " & _
        "être faite : " & Err.Description & " (n° " & Err.Number & "). Vérifiez que le module " & _
        macro & " est importé.", vbExclamation
End Sub

' À l'activation de la feuille (code de la feuille, Worksheet_Activate) : feuille
' recréée si le nombre de micropieux de la descente de charges (K1) a changé depuis sa
' création
Public Sub VerifierNombreMicropieux(ws As Worksheet)
    Dim f As Worksheet, k As Long, v As Variant
    If mEnCours Then Exit Sub
    For Each f In ActiveWorkbook.Worksheets
        If CStr(f.Range("B80").Value) = "RÉPARTITION DES CHARGES DES MURS SUR LES MICROPIEUX" Then
            k = CLng(N0(f.Range("K1").Value))
            Exit For
        End If
    Next f
    If k < 1 Then Exit Sub
    On Error Resume Next
    v = ws.Evaluate("NbMicropieuxDDC")
    On Error GoTo 0
    If IsError(v) Or IsEmpty(v) Then
        CreerVerificationStructurelle
    ElseIf CLng(N0(v)) <> k Then
        CreerVerificationStructurelle
    End If
End Sub

' Code de la feuille « Vérification structurelle » : Worksheet_Activate appelle
' VerifierNombreMicropieux (accès au projet VBA requis : Options > Centre de gestion
' de la confidentialité > Paramètres des macros > « Accès approuvé au modèle d'objet du
' projet VBA » ; sinon coller ce code à la main dans le module de la feuille)
Private Sub InstallerCodeActivation(ws As Worksheet)
    Dim cm As Object, texte As String
    On Error GoTo Fin
    Set cm = ws.Parent.VBProject.VBComponents(ws.CodeName).CodeModule
    If cm.CountOfLines > 0 Then texte = cm.Lines(1, cm.CountOfLines)
    If InStr(1, texte, "VerifierNombreMicropieux", vbTextCompare) = 0 Then
        cm.AddFromString "Private Sub Worksheet_Activate()" & vbCrLf & _
            "    VerifierNombreMicropieux Me" & vbCrLf & "End Sub"
    End If
Fin:
End Sub

Private Sub Commentaire(c As Range, texte As String)
    ' Commentaire en image (module CommentaireImage) ; à défaut, commentaire texte
    On Error Resume Next
    Application.Run "CommentaireImage", c, texte
    If Err.Number <> 0 Then
        Err.Clear
        c.ClearComments
        c.AddComment texte
        c.Comment.Shape.TextFrame.AutoSize = True
    End If
    On Error GoTo 0
End Sub

' Textes trop serrés avec la police de la charte (plus large que Calibri) : colonnes
' utilisées élargies de 15 % ; graphiques et images replacés sur leur cellule d'origine
Private Sub AjusterTextes(ws As Worksheet)
    Dim c As Range, sh As Shape, cel() As Range, k As Long
    On Error Resume Next
    If ws.Shapes.Count > 0 Then
        ReDim cel(1 To ws.Shapes.Count)
        For k = 1 To ws.Shapes.Count
            Set cel(k) = ws.Shapes(k).TopLeftCell
        Next k
    End If
    For Each c In ws.UsedRange.Columns
        If Not c.EntireColumn.Hidden Then c.ColumnWidth = Application.WorksheetFunction.Min(255, c.ColumnWidth * 1.15)
    Next c
    For k = 1 To ws.Shapes.Count
        If Not cel(k) Is Nothing Then ws.Shapes(k).Left = cel(k).Left
    Next k
    On Error GoTo 0
End Sub
