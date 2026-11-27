Attribute VB_Name = "VerificationStructurelle"
Option Explicit
Private Const COUL_TITRE As Long = 5898259
Private Const COUL_ENTETE As Long = 14337945
Private Const COUL_SAISIE As Long = 15201485
Private Const COUL_CALCUL As Long = 15590349
Private Const COUL_ALERTE As Long = 192
Private Const NOM_FEUILLE As String = "Vérification structurelle"
Private Const DONNEES As String = "'Données du projet'!"
Private Const LIB_AUTRE As String = "Autre diamètre (à saisir, vide : " & "Phi retenu)"
Private Const LIB_PAS As String = "Pas de calcul dz (à saisir)"
Private Const LIB_LIBRE As String = "Condition en tête"
Private Const TETE_LIBRE As String = "Micropieu libre en tête"
Private Const TETE_ENC As String = "Micropieu encastré en tête"
Private Const CALCUL As String = "'Calcul Portance'!"
Public Const TITRE_FRONTALE As String = "VÉRIFICATION DE LA RÉSISTANCE FRONTALE DE LA FONDATION"
Private mEnCours As Boolean
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
    Dim lMat As Long
    Dim libre As Variant, cTal As Long, fH0 As String, fM0 As String
    Application.ScreenUpdating = False
    mEnCours = True
    On Error GoTo Erreur
    etape = "feuilles du projet"
    Application.Run "RenommerFeuilleCalcul"
    Set wc = ActiveWorkbook.Worksheets("Calcul Portance")
    For Each f In ActiveWorkbook.Worksheets
        If TrouverLigneMatrice(f) > 0 Then
            Set wd = f
            lMat = TrouverLigneMatrice(f)
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
    autre = LireValeur(ws, LIB_AUTRE)
    pasZ = LireValeur(ws, LIB_PAS)
    libre = LireValeur(ws, LIB_LIBRE)
    ws.Cells.Clear
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop
    a = LIGNE_ACIER
    lp = LIGNE_PIEU
    lm = LIGNE_COUCHE_MOLLE
    With ws.Range("B1:I1")
        .Merge
        .Value = "VÉRIFICATION STRUCTURELLE DES MICROPIEUX"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COUL_TITRE
        .HorizontalAlignment = xlCenter
    End With
    tube = DONNEES & "$D$" & a & "=""Tube N80"""
    dA = "N(" & DONNEES & "$D$" & a + 2 & ")"
    dI = "N(" & DONNEES & "$D$" & a + 4 & ")"
    eA = "N(" & DONNEES & "$D$" & a + 6 & ")"
    pen = "(" & DONNEES & CELLULE_METHODE & "=""Méthode pénétrométrique"")"
    sol = DONNEES & "$D$" & lp + 21
    b0 = "N(" & DONNEES & "$D$" & lp + 22 & ")"
    emq = "N(" & DONNEES & "$D$" & lm + 1 & ")"
    alp = "N(" & DONNEES & "$D$" & lm + 6 & ")"
    nat = DONNEES & "$D$" & lm + 5
    bet = "IF(OR(" & nat & "=""Tourbe""," & nat & "=""Argile""),12,IF(" & nat & "=""Limon"",7.5,4.5))"
    etape = "micropieux de la vérification de la profondeur"
    If Not wd Is Nothing Then
        k = CLng(N0(wd.Range("K1").Value))
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
        Do While InStr("|Mur |Pot |", "|" & Left(CStr(wd.Range("B" & lMat + 6 + n).Value), 4) & "|") > 0
            n = n + 1
        Loop
        rd = lMat + 8 + n + 3
        dn = "'" & wd.Name & "'!"
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
        ws.Range("A3").Value = "Créez la matrice des micropieux (descente de charges) " & "et « Calcul Portance », puis relancez CreerVerificationStructurelle."
        GoTo MiseEnForme
    End If
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
    ci = r0 + 2 + iMax
    etape = "longueur de transfert"
    With ws.Range("A3")
        .Value = "LONGUEUR DE TRANSFERT l0 DU MICROPIEU LE PLUS LONG (MP" & iMax & ")"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COUL_TITRE
    End With
    r = 4
    r = r + 1: Ligne ws, r, "Micropieu le plus long (plus grande profondeur réelle)", "MP" & iMax, ""
    r = r + 1: Ligne ws, r, ChrW(934) & " retenu (vérification de la profondeur)", IIf(cPhi > 0, "=IF(ISNUMBER(" & CALCUL & wc.Cells(ci, cPhi).Address & ")," & CALCUL & wc.Cells(ci, cPhi).Address & ",N(" & CALCUL & "$D$19))", "=N(" & CALCUL & "$D$19)"), "m": vPhi = "$D$" & r
    r = r + 1: Ligne ws, r, LIB_AUTRE, "", "m": vAut = "$D$" & r
    If IsNumeric(autre) And Not IsEmpty(autre) Then ws.Range("D" & r).Value = autre
    ws.Range("D" & r).Interior.Color = COUL_SAISIE
    r = r + 1: Ligne ws, r, "Diamètre du pieu foré B", "=IF(N(" & vAut & ")>0," & vAut & "," & vPhi & ")", "m": vB = "$D$" & r
    r = r + 1: Ligne ws, r, "Section du pieu foré A = " & ChrW(960) & ".B²/4", "=PI()*" & vB & "^2/4", "m²"
    r = r + 1: Ligne ws, r, "Diamètre de l'armature d", "=" & dA, "mm": vdA = "$D$" & r
    r = r + 1: Ligne ws, r, "Diamètre intérieur di (tube)", "=IF(" & tube & "," & dI & ",0)", "mm": vdi = "$D$" & r
    r = r + 1: Ligne ws, r, "Enrobage c = (B - d) / 2", "=MAX(0,(1000*" & vB & "-" & vdA & ")/2)", "mm": vc = "$D$" & r
    r = r + 1: Ligne ws, r, "Protection par le coulis (c >= 50 mm, C > 500 kg/m3, E/C < 0,5, compression)", "=IF(AND(" & vc & ">=50,N(" & DONNEES & "D" & lp + 28 & ")>500,N(" & DONNEES & "D" & lp + 27 & ")>0,N(" & DONNEES & "D" & lp + 27 & ")<0.5," & DONNEES & "D" & lp + 12 & "<>""Traction""),""Efficace"",""Non efficace"")", "": vPro = "$D$" & r
    r = r + 1: Ligne ws, r, "Rayon de corrosion (« Données du projet »)", "=N(" & DONNEES & "D" & a + 13 & ")", "mm": vrcD = "$D$" & r
    r = r + 1: Ligne ws, r, "Rayon de corrosion retenu rc (0 si protection efficace)", "=IF(" & vPro & "=""Efficace"",0," & vrcD & ")", "mm": vrc = "$D$" & r
    r = r + 1: Ligne ws, r, "Diamètre après corrosion d - 2 rc", "=MAX(0," & vdA & "-2*" & vrc & ")", "mm": vdc = "$D$" & r
    r = r + 1: Ligne ws, r, "Section brute de l'armature", "=IF(" & tube & ",PI()/4*(" & vdA & "^2-" & vdi & "^2),PI()/4*" & vdA & "^2)", "mm²": vS0 = "$D$" & r
    r = r + 1: Ligne ws, r, "Section après corrosion As", "=MAX(0,IF(" & tube & ",PI()/4*(" & vdc & "^2-" & vdi & "^2),PI()/4*" & vdc & "^2))", "mm²": vS = "$D$" & r
    r = r + 1: Ligne ws, r, "Section corrodée", "=" & vS0 & "-" & vS, "mm²"
    r = r + 1: Ligne ws, r, "Inertie de la section brute", "=IF(" & tube & ",PI()/64*(" & vdA & "^4-" & vdi & "^4),PI()/64*" & vdA & "^4)", "mm4"
    r = r + 1: Ligne ws, r, "Inertie de la section après corrosion I", "=MAX(0,IF(" & tube & ",PI()/64*(" & vdc & "^4-" & vdi & "^4),PI()/64*" & vdc & "^4))", "mm4": vI = "$D$" & r
    r = r + 1: Ligne ws, r, "Module de Young de l'armature E", "=" & eA, "MPa": vE = "$D$" & r
    r = r + 1: Ligne ws, r, "Module de flexion EI = E x I", "=" & vE & "*" & vI & "*1E-9", "kN.m²": vEI = "$D$" & r
    r = r + 1: Ligne ws, r, "Méthode de calcul", "=" & DONNEES & CELLULE_METHODE, ""
    r = r + 1: Ligne ws, r, "=IF(" & pen & ",""Résistance de pointe qc de la couche molle""," & """Module pressiométrique EM de la couche molle"")", "=" & emq, "MPa": vEM = "$D$" & r
    r = r + 1: Ligne ws, r, "=IF(" & pen & ",""" & ChrW(946) & " (CPT)"",""Coefficient rhéologique " & ChrW(945) & """)", "=IF(" & pen & "," & bet & "," & alp & ")", "": vAl = "$D$" & r
    r = r + 1: Ligne ws, r, "Diamètre de référence B0", "=" & b0, "m": vB0 = "$D$" & r
    r = r + 1: Ligne ws, r, ChrW(961) & " = max(B / B0 ; 1)", "=IF(" & vB0 & "<=0,"""",MAX(" & vB & "/" & vB0 & ",1))", "": vRho = "$D$" & r
    r = r + 1: Ligne ws, r, "Type de sollicitation", "=" & sol, ""
    r = r + 1: Ligne ws, r, "Durée des charges pour le calcul de Kf (« Données du projet »)", "=IF(" & DONNEES & "$D$" & lp + 23 & "=""Courte durée"",""Courte durée"",""Longue durée"")", "": vDur = "$D$" & r
    x = "12 EM / [4/3 (B0/B)(2,65 B/B0)^" & ChrW(945) & " + " & ChrW(945) & "] (B > B0), 12 EM / [4/3 (2,65)^" & ChrW(945) & " + " & ChrW(945) & "] (B " & ChrW(8804) & " B0)"
    r = r + 1: Ligne ws, r, "=IF(" & vDur & "=""Longue durée"",""Module de réaction Kf de longue durée = " & x & " / 2"",""Module de réaction Kf de courte durée = " & x & """)", "=IF(OR(N(" & vEM & ")<=0," & vRho & "=""""),"""",IF(" & pen & "," & vAl & "*" & vEM & ",12*" & vEM & "/(4/3*(2.65*" & vRho & ")^" & vAl & "/" & vRho & "+" & vAl & "))/IF(" & vDur & "=""Longue durée"",2,1))", "MPa": vKf = "$D$" & r
    r = r + 1: Ligne ws, r, "Module Es = Kf", "=" & vKf, "MPa": vEs = "$D$" & r
    r = r + 1: Ligne ws, r, "Longueur de transfert l0 = (4 EI / Es)^(1/4)", "=IF(OR(" & vEs & "="""",N(" & vEs & ")=0),"""",(4*" & vEI & "/(" & vEs & "*1000))^0.25)", "m": vl0 = "$D$" & r
    r = r + 1: Ligne ws, r, ChrW(955) & " = 1 / l0", "=IF(N(" & vl0 & ")<=0,"""",1/" & vl0 & ")", "1/m": vLam = "$D$" & r
    r = r + 1: Ligne ws, r, "Profondeur réelle du micropieu D", IIf(cDr > 0, "=N(" & CALCUL & wc.Cells(ci, cDr).Address & ")", "=N(" & CALCUL & "$D$20)"), "m": vD = "$D$" & r
    r = r + 1: Ligne ws, r, "x = D / l0", "=IF(" & vLam & "="""",""""," & vLam & "*" & vD & ")", "": vx = "$D$" & r
    r = r + 1: Ligne ws, r, "Comportement du micropieu", "=IF(" & vx & "="""",""-"",IF(" & vx & "<=1," & """Rigide (l0 >= D)"",IF(" & vx & ">=3,""Pieu long (D >= 3 l0)"",""Intermédiaire (calcul approché)"")))", ""
    r = r + 1: Ligne ws, r, "Micropieu incliné", IIf(cIn > 0, "=IF(" & CALCUL & wc.Cells(ci, cIn).Address & "=""Oui"",""Oui"",""Non"")", "Non"), "": vInc = "$D$" & r
    r = r + 1: Ligne ws, r, "Angle d'inclinaison sur la verticale " & ChrW(945), IIf(cAn > 0, "=IF(" & vInc & "=""Oui"",N(" & CALCUL & wc.Cells(ci, cAn).Address & "),0)", "0"), "°": vAng = "$D$" & r
    r = r + 1: Ligne ws, r, LIB_LIBRE, "", "": vCond = "$D$" & r
    With ws.Range("D" & r)
        .Validation.Delete
        .Validation.Add xlValidateList, xlValidAlertStop, , TETE_LIBRE & "," & TETE_ENC
        .Value = IIf(CStr(ActiveWorkbook.Worksheets("Données du projet").Range("D" & lp + 18).Value) = "Encastré en tête", TETE_ENC, TETE_LIBRE)
        If CStr(libre) = TETE_LIBRE Or CStr(libre) = TETE_ENC Then .Value = libre
        If CStr(libre) = "Oui" Then .Value = TETE_LIBRE
        If CStr(libre) = "Non" Then .Value = TETE_ENC
    End With
    Commentaire ws.Range("B" & r), "Micropieu libre en tête : rotation libre, M0 = 0 ; micropieu " & "encastré en tête : tête encastrée dans la semelle (rotation nulle, moment d'encastrement M0 = -H.l0/2). Valeur par défaut : " & "« Données du projet », conditions en tête de la fondation"
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
    etape = "butée des terres sur l'ancrage"
    nbC = CLng(N0(ActiveWorkbook.Worksheets("Données du projet").Range(CELLULE_NB_COUCHES).Value))
    If nbC < 1 Then nbC = 1
    If nbC > MAX_COUCHES Then nbC = MAX_COUCHES
    rA = rF + 3
    With ws.Range("A" & rA)
        .Value = "BUTÉE DES TERRES SUR L'ANCRAGE (POTEAU BA + SEMELLE BA) - MICROPIEU MP" & iMax
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COUL_TITRE
    End With
    ws.Range("B" & rA + 1 & ":E" & rA + 1).Value = Array("Hauteur du poteau BA (soubassement)", "hp", "", "m")
    ws.Range("B" & rA + 2 & ":E" & rA + 2).Value = Array("Hauteur de la semelle BA (soubassement)", "hs", "", "m")
    ws.Range("B" & rA + 3 & ":E" & rA + 3).Value = Array("Hauteur de l'ancrage", "h", "", "m")
    ws.Range("D" & rA + 1).Formula = IIf(rPo > 0, "=N(" & dn & "$G$" & rPo & ")", "=0")
    ws.Range("D" & rA + 2).Formula = IIf(rS > 0, "=N(" & dn & "$G$" & rS & ")", "=0")
    ws.Range("D" & rA + 3).Formula = "=D" & rA + 1 & "+D" & rA + 2
    Commentaire ws.Range("C" & rA + 1), "Descente de charges, soubassement : profondeur e (colonne G) de la ligne Poteau"
    Commentaire ws.Range("C" & rA + 2), "Descente de charges, soubassement : profondeur e (colonne G) de la ligne Semelle"
    Commentaire ws.Range("C" & rA + 3), "h = hp + hs : hauteur de terrain en contact avec l'ancrage, " & "répartie sur les couches de « Données du projet » à partir du TN"
    With ws.Range("B" & rA + 1 & ":E" & rA + 3)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With
    ws.Range("C" & rA + 1 & ":E" & rA + 3).HorizontalAlignment = xlCenter
    ws.Range("C" & rA + 1 & ":C" & rA + 3).Font.Bold = True
    ws.Range("D" & rA + 1 & ":D" & rA + 3).Interior.Color = COUL_CALCUL
    ws.Range("D" & rA + 1 & ":D" & rA + 3).NumberFormat = "0.00"
    x = "$D$" & rA + 3
    r = rA + 5
    ws.Range("A" & r & ":O" & r).Value = Array("N° de couche", "Nom de couche", "z haut (m)", "z bas (m)", "Épaisseur dans l'ancrage hi (m)", ChrW(947) & " (kN/m3)", "c' (kPa)", ChrW(966) & "' (°)", "K0 = (1 + sin " & ChrW(966) & "')/(1 - sin " & ChrW(966) & "')", ChrW(963) & "v haut (kPa)", ChrW(963) & "v bas (kPa)", ChrW(963) & "h haut = K0." & ChrW(963) & "v + 2c'" & ChrW(8730) & "K0 (kPa)", ChrW(963) & "h bas = K0." & ChrW(963) & "v + 2c'" & ChrW(8730) & "K0 (kPa)", "Bp,i = (" & ChrW(963) & "h haut + " & ChrW(963) & "h bas)/2 . hi (kN/ml)", "Profondeur d'application (m)")
    Commentaire ws.Range("C" & r), "z haut = min(profondeur du haut de la couche depuis le TN ; h)"
    Commentaire ws.Range("D" & r), "z bas = min(profondeur du bas de la couche depuis le TN ; h) ; " & "épaisseurs du tableau des couches de « Données du projet »"
    Commentaire ws.Range("I" & r), "K0 = (1 + sin " & ChrW(966) & "') / (1 - sin " & ChrW(966) & "')"
    Commentaire ws.Range("J" & r), ChrW(963) & "v haut = somme des " & ChrW(947) & "j.hj des couches au-dessus"
    Commentaire ws.Range("O" & r), "Centre de gravité du diagramme trapézoïdal de poussée de la couche : " & "z haut + hi (" & ChrW(963) & "h haut + 2 " & ChrW(963) & "h bas) / (3 (" & ChrW(963) & "h haut + " & ChrW(963) & "h bas))"
    For i = 1 To nbC
        r = rA + 5 + i
        k = LIGNE_COUCHES + i - 1
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
        ws.Range("O" & r).Formula = "=IF(AND(E" & r & ">0,L" & r & "+M" & r & ">0),C" & r & "+E" & r & "*(L" & r & "+2*M" & r & ")/(3*(L" & r & "+M" & r & ")),0)"
    Next i
    r = rA + 6 + nbC
    ws.Range("A" & r).Value = "Butée des terres totale sur l'ancrage " & ChrW(931) & "Bp,i et profondeur d'application"
    ws.Range("A" & r & ":M" & r).Merge
    ws.Range("N" & r).Formula = "=SUM(N" & rA + 6 & ":N" & r - 1 & ")"
    vBut = "$N$" & r
    ws.Range("O" & r).Formula = "=IF(N" & r & ">0,SUMPRODUCT(N" & rA + 6 & ":N" & r - 1 & ",O" & rA + 6 & ":O" & r - 1 & ")/N" & r & ",0)"
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
    rF = r
    etape = "efforts en tête et constantes"
    rT = rF + 3
    With ws.Range("A" & rT)
        .Value = "EFFORTS EN TÊTE (DESCENTE DE CHARGES) ET CONSTANTES DE LA SOLUTION (PIEU LONG, " & "POINTE ENCASTRÉE) - MICROPIEU MP" & iMax
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COUL_TITRE
    End With
    ws.Range("B" & rT + 1).Formula = "=""Condition en tête : ""&" & vCond
    r = rT + 2
    ws.Range("B" & r & ":I" & r).Value = Array("Grandeur", "Symbole", "Formule", "ELS CARA", "ELS QP", "ELU FOND", "ELU ACC", "Unité")
    cTal = 0
    If Not wd Is Nothing Then
        For j = 3 To 250
            If CStr(wd.Cells(lMat + 8 + n, j).Value) = "CHARGE DU TALUS" Then cTal = j: Exit For
        Next j
    End If
    If cTal > 0 Then
        fH0 = "N(" & dn & wd.Cells(rd + iMax - 1, cTal).Address & ")"
        fM0 = "N(" & dn & wd.Cells(rd + iMax - 1, cTal + 1).Address & ")"
    Else
        fH0 = "0": fM0 = "0"
    End If
    entete = Array( Array("Effort de compression de la descente de charges", "Nc", "Vmax de la combinaison (vérification de la profondeur, « Calcul Portance »)", "kN"), Array("Effort horizontal en tête", "H0", "H0 = Nc x sin " & ChrW(945) & " (composante inclinée de la " & "descente de charges, 0 si vertical) + H talus (descente de charges, charge du talus)", "kN"), Array("Moment en tête (selon la condition en tête)", "M0", "micropieu libre en tête : M0 = M talus " & "(descente de charges, charge du talus) ; micropieu encastré en tête : M0 = -H0.l0/2 = -H0/(2" & ChrW(955) & ") (moment repris par la tête encastrée)", "kN.m"), Array("Constante de l'onde issue de la tête", "A", "micropieu libre en tête : A = 2" & ChrW(955) & "H0/Es + 2" & ChrW(955) & "²M0/Es ; micropieu encastré en tête : A = " & ChrW(955) & "H0/Es", "m"), Array("Constante de l'onde issue de la tête", "B", "micropieu libre en tête : B = -2" & ChrW(955) & "²M0/Es ; " & "micropieu encastré en tête : B = A", "m"), _
        Array("Constante de l'onde issue de la pointe", "C", "C = -e^(-" & ChrW(955) & "D) [A cos " & ChrW(955) & "D + B sin " & ChrW(955) & "D]", "m"), Array("Constante de l'onde issue de la pointe", "E", "E = e^(-" & ChrW(955) & "D) [(B - 2A) cos " & ChrW(955) & "D - (A + 2B) sin " & ChrW(955) & "D]", "m"))
    For j = 0 To 6
        ws.Range("B" & r + 1 + j).Value = entete(j)(0)
        ws.Range("C" & r + 1 + j).Value = entete(j)(1)
        ws.Range("D" & r + 1 + j).Value = entete(j)(2)
        ws.Range("I" & r + 1 + j).Value = entete(j)(3)
    Next j
    Commentaire ws.Range("C" & r + 2), "H0 = Nc x sin " & ChrW(945) & " + H talus : composante horizontale de " & "l'effort de la descente de charges (micropieu incliné) plus l'effort horizontal du talus (descente de " & "charges, matrice des micropieux, colonne « CHARGE DU TALUS »)."
    Commentaire ws.Range("C" & r + 3), "Micropieu libre en tête : M0 = moment du talus (descente de charges, " & "colonne « CHARGE DU TALUS »)." & vbLf & "Micropieu encastré en tête : M0 = -H0/(2" & ChrW(955) & ") = -H0.l0/2, moment repris par la tête encastrée (le moment du talus est repris par l'encastrement)."
    For k = 0 To 3
        cl = Chr(Asc("E") + k)
        x = cl & "$"
        ws.Range(cl & r + 1).Formula = "=N(" & CALCUL & wc.Cells(ci, cC + 2 * k).Address & ")"
        ws.Range(cl & r + 2).Formula = "=IF(" & vInc & "=""Oui"",ABS(" & x & r + 1 & ")*SIN(RADIANS(" & vAng & ")),0)+" & fH0
        ws.Range(cl & r + 3).Formula = "=IF(" & vCond & "=""" & TETE_ENC & """,IF(N(" & vLam & ")>0,-" & x & r + 2 & "/(2*" & vLam & "),0)," & fM0 & ")"
        cond0 = "=IF(OR(" & vLam & "="""",N(" & vEs & ")<=0),"""","
        ws.Range(cl & r + 4).Formula = cond0 & "IF(" & vCond & "=""" & TETE_ENC & """," & vLam & "*" & x & r + 2 & "/(" & vEs & "*1000),2*" & vLam & "*" & x & r + 2 & "/(" & vEs & "*1000)+2*" & vLam & "^2*" & x & r + 3 & "/(" & vEs & "*1000)))"
        ws.Range(cl & r + 5).Formula = cond0 & "IF(" & vCond & "=""" & TETE_ENC & """," & x & r + 4 & ",-2*" & vLam & "^2*" & x & r + 3 & "/(" & vEs & "*1000)))"
        ws.Range(cl & r + 6).Formula = "=IF(" & x & r + 4 & "="""","""",-EXP(-" & vx & ")*(" & x & r + 4 & "*COS(" & vx & ")+" & x & r + 5 & "*SIN(" & vx & ")))"
        ws.Range(cl & r + 7).Formula = "=IF(" & x & r + 4 & "="""","""",EXP(-" & vx & ")*((" & x & r + 5 & "-2*" & x & r + 4 & ")*COS(" & vx & ")-(" & x & r + 4 & "+2*" & x & r + 5 & ")*SIN(" & vx & ")))"
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
    rK = r
    etape = "résistance frontale de la fondation"
    rV = rK + 9
    With ws.Range("A" & rV)
        .Value = TITRE_FRONTALE & " (MP" & iMax & ")"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COUL_TITRE
    End With
    r = rV + 1
    ws.Range("B" & r & ":I" & r).Value = Array("Grandeur", "Symbole", "Formule", "ELS CARA", "ELS QP", "ELU FOND", "ELU ACC", "Unité")
    ws.Range("B" & r + 1 & ":D" & r + 1).Value = Array("Butée des terres totale sur l'ancrage (par mètre linéaire)", ChrW(931) & "Bp,i", "tableau « Butée des terres sur l'ancrage »")
    ws.Range("B" & r + 2 & ":D" & r + 2).Value = Array("Largeur de mur la plus grande reprise par le micropieu", "L", "descente de charges : plus grande des longueurs de murs de la ligne du micropieu (matrice des " & "micropieux)")
    ws.Range("B" & r + 3 & ":D" & r + 3).Value = Array("Butée des terres reprise par le micropieu", "Bp", "Bp = " & ChrW(931) & "Bp,i x L")
    ws.Range("B" & r + 4 & ":D" & r + 4).Value = Array("Effort horizontal en tête", "H0", "tableau des efforts en tête (H0 = Nc x sin " & ChrW(945) & " + H talus)")
    ws.Range("B" & r + 5 & ":D" & r + 5).Value = Array("Rapport effort horizontal / butée", "H0 / Bp", "Condition : H0 / Bp < 1")
    ws.Range("B" & r + 6 & ":D" & r + 6).Value = Array("Vérification", "", "H0 / Bp < 1 : la butée reprend l'effort horizontal ; sinon calcul de Y, T et M du micropieu")
    ws.Range("I" & r + 1 & ":I" & r + 5).Value = Application.Transpose(Array("kN/ml", "m", "kN", "kN", ""))
    For k = 0 To 3
        cl = Chr(Asc("E") + k)
        ws.Range(cl & r + 1).Formula = "=N(" & vBut & ")"
        If Not wd Is Nothing And n > 0 And Left(CStr(wd.Cells(lMat + 8 + n, 3).Value), 9) = "LONGUEURS" Then
            ws.Range(cl & r + 2).Formula = "=MAX(" & dn & wd.Range(wd.Cells(rd + iMax - 1, 3), wd.Cells(rd + iMax - 1, 2 + n)).Address & ")"
        Else
            ws.Range(cl & r + 2).Formula = "=0"
        End If
        ws.Range(cl & r + 3).Formula = "=" & cl & r + 1 & "*" & cl & r + 2
        ws.Range(cl & r + 4).Formula = "=ABS(N(" & cl & "$" & rK + 2 & "))"
        ws.Range(cl & r + 5).Formula = "=IF(N(" & cl & r + 3 & ")>0," & cl & r + 4 & "/" & cl & r + 3 & ",IF(N(" & cl & r + 4 & ")=0,0,""""))"
        ws.Range(cl & r + 6).Formula = "=IF(" & cl & r + 5 & "="""",""Butée nulle : calcul du micropieu""," & "IF(" & cl & r + 5 & "<1,""OK : calcul non nécessaire"",""Butée insuffisante : calcul du micropieu""))"
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
            .Value = "H0 / Bp < 1 pour toutes les combinaisons : la butée des terres sur l'ancrage reprend " & "l'effort horizontal en tête ; le calcul des déplacements Y, de l'effort tranchant T et du " & "moment M du micropieu n'est pas nécessaire."
            .Font.Bold = True
            .Font.Color = COUL_TITRE
            .WrapText = True
        End With
        ws.Rows(rF + 2).RowHeight = 32
        GoTo MiseEnForme
    End If
    etape = "Y, T et M le long du micropieu"
    rP = rF + 5
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
    ws.Calculate
    nZ = 1
    If N0(ws.Range(vD).Value) > 0 And N0(ws.Range(vPas).Value) > 0 Then nZ = Application.RoundUp(N0(ws.Range(vD).Value) / N0(ws.Range(vPas).Value) - 0.000000001, 0)
    If nZ < 1 Then nZ = 1
    If nZ > 2000 Then nZ = 2000
    With ws.Range("A" & rP - 3)
        .Value = "DÉPLACEMENT Y, EFFORT TRANCHANT T ET MOMENT M LE LONG DU MICROPIEU MP" & iMax
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COUL_TITRE
    End With
    ws.Range("A" & rP & ":D" & rP).Value = Array("z (m)", "Pas de calcul (m)", ChrW(955) & "z", ChrW(955) & "u = " & ChrW(955) & "(D - z)")
    For k = 0 To 3
        ws.Range(ws.Cells(rP, 1 + k), ws.Cells(rP + 1, 1 + k)).Merge
        ws.Cells(rP, 5 + 3 * k).Value = Choose(k + 1, "ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
        ws.Range(ws.Cells(rP, 5 + 3 * k), ws.Cells(rP, 7 + 3 * k)).Merge
        ws.Range(ws.Cells(rP + 1, 5 + 3 * k), ws.Cells(rP + 1, 7 + 3 * k)).Value = Array("Y (mm)", "T (kN)", "M (kN.m)")
    Next k
    For j = 0 To nZ
        r = rP + 2 + j
        xs = "$A" & r
        If j = 0 Then
            ws.Range("A" & r).Formula = "=IF(N(" & vD & ")>0,0,"""")"
            ws.Range("B" & r).Formula = "=IF(A" & r & "="""","""",0)"
        Else
            ws.Range("A" & r).Formula = "=IF(A" & r - 1 & "="""","""",IF(A" & r - 1 & ">=" & vD & "-1E-9,""""," & "MIN(A" & r - 1 & "+" & vPas & "," & vD & ")))"
            ws.Range("B" & r).Formula = "=IF(A" & r & "="""","""",A" & r & "-A" & r - 1 & ")"
        End If
        ws.Range("C" & r).Formula = "=IF(OR(" & xs & "=""""," & vLam & "=""""),""""," & vLam & "*" & xs & ")"
        ws.Range("D" & r).Formula = "=IF(OR(" & xs & "=""""," & vLam & "=""""),""""," & vLam & "*(" & vD & "-" & xs & "))"
        kA = "$C" & r: kB = "$D" & r
        For k = 0 To 3
            cl = Chr(Asc("E") + k)
            y0s = cl & "$" & rK + 4: hs = cl & "$" & rK + 5
            ths = cl & "$" & rK + 6: mhs = cl & "$" & rK + 7
            cond0 = "=IF(OR(" & kA & "=""""," & y0s & "=""""),"""","
            ws.Cells(r, 5 + 3 * k).Formula = cond0 & "1000*(EXP(-" & kA & ")*(" & y0s & "*COS(" & kA & ")+" & hs & "*SIN(" & kA & "))+EXP(-" & kB & ")*(" & ths & "*COS(" & kB & ")+" & mhs & "*SIN(" & kB & "))))"
            ws.Cells(r, 6 + 3 * k).Formula = cond0 & "2*" & vEI & "*" & vLam & "^3*(EXP(-" & kA & ")*((" & y0s & "+" & hs & ")*COS(" & kA & ")+(" & hs & "-" & y0s & ")*SIN(" & kA & "))-EXP(-" & kB & ")*((" & ths & "+" & mhs & ")*COS(" & kB & ")+(" & mhs & "-" & ths & ")*SIN(" & kB & "))))"
            ws.Cells(r, 7 + 3 * k).Formula = cond0 & "2*" & vEI & "*" & vLam & "^2*(EXP(-" & kA & ")*(" & y0s & "*SIN(" & kA & ")-" & hs & "*COS(" & kA & "))+EXP(-" & kB & ")*(" & ths & "*SIN(" & kB & ")-" & mhs & "*COS(" & kB & "))))"
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
    rG = rP + 2 + nZ + 3
MiseEnForme:
    etape = "mise en forme"
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:Y").ColumnWidth = 12
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
    If rK > 0 Then
        LancerVerification "VerificationArmature", "vérification de l'armature"
        LancerVerification "VerificationMandel", "vérification au flambement (Mandel)"
        LancerVerification "VerificationPlatine", "dimensionnement des platines"
    End If
    ws.Cells.Font.Name = "CheGuevara Text Sans"
    AjusterTextes ws
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
    MsgBox "Erreur pendant la création de la feuille « " & NOM_FEUILLE & " », partie « " & etape & " » : " & Err.Description & " (n° " & Err.Number & ")", vbExclamation
    Resume Fin
End Sub
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
Private Function LireValeur(ws As Worksheet, lib As String) As Variant
    Dim v As Variant
    v = Application.Match(lib, ws.Range("B1:B300"), 0)
    If IsError(v) Then Exit Function
    LireValeur = ws.Range("D" & CLng(v)).Value
End Function
Private Sub GraphiquesProfil(ws As Worksheet, i As Long, r1 As Long, nb As Long, rH As Long, cH As Long)
    Dim co As ChartObject, g As Long, k As Long, n As Long, lg As Double, rz As Range
    Dim noms As Variant, titres As Variant
    Do While n < nb
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
                    .Format.Line.ForeColor.RGB = Choose((k Mod 4) + 1, 5898259, 11371036, 8964355, 8139520)
                End With
            Next k
            .HasTitle = True
            .ChartTitle.Text = "MP" & i & " - " & titres(g) & " le long du micropieu"
            .HasLegend = True
            .Legend.Position = xlLegendPositionBottom
            With .Axes(xlValue)
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
Private Function N0(v As Variant) As Double
    If Not IsEmpty(v) And IsNumeric(v) Then N0 = CDbl(v)
End Function
Private Sub SousTitre(ws As Worksheet, r As Long, texte As String)
    With ws.Range("B" & r)
        .Value = texte
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Size = 14
        .Font.Color = COUL_TITRE
    End With
End Sub
Public Function CreationVerificationEnCours() As Boolean
    CreationVerificationEnCours = mEnCours
End Function
Private Sub LancerVerification(macro As String, libelle As String)
    On Error GoTo Echec
    Select Case macro
        Case "VerificationArmature": VerificationArmature.VerificationArmature
        Case "VerificationMandel": VerificationMandel.VerificationMandel
        Case "VerificationPlatine": VerificationPlatine.VerificationPlatine
    End Select
    Exit Sub
Echec:
    MsgBox "« Vérification structurelle » : la " & libelle & " (macro " & macro & ") n'a pas pu " & "être faite : " & Err.Description & " (n° " & Err.Number & "). Vérifiez que le module " & macro & " est importé.", vbExclamation
End Sub
Public Sub VerifierNombreMicropieux(ws As Worksheet)
    Dim f As Worksheet, k As Long, v As Variant
    If mEnCours Then Exit Sub
    For Each f In ActiveWorkbook.Worksheets
        If TrouverLigneMatrice(f) > 0 Then
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
Private Sub InstallerCodeActivation(ws As Worksheet)
    Dim cm As Object, texte As String
    On Error GoTo Fin
    Set cm = ws.Parent.VBProject.VBComponents(ws.CodeName).CodeModule
    If cm.CountOfLines > 0 Then texte = cm.Lines(1, cm.CountOfLines)
    If InStr(1, texte, "VerifierNombreMicropieux", vbTextCompare) = 0 Then
        cm.AddFromString "Private Sub Worksheet_Activate()" & vbCrLf & "    VerifierNombreMicropieux Me" & vbCrLf & "End Sub"
    End If
Fin:
End Sub
Private Sub Commentaire(c As Range, texte As String)
    On Error Resume Next
    Application.Run "CommentaireImage.CommentaireImage", c, texte
    If Err.Number <> 0 Then
        Err.Clear
        c.ClearComments
        c.AddComment texte
        c.Comment.Shape.TextFrame.AutoSize = True
    End If
    On Error GoTo 0
End Sub
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
