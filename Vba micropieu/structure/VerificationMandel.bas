Attribute VB_Name = "VerificationMandel"
Option Explicit
' Couleurs de la charte graphique GEOSKILLS (bleu nuit, bleu océan, vert émeraude)
Private Const COUL_TITRE As Long = 5898259         ' RGB(19, 0, 90) : Bleu Nuit (charte GEOSKILLS)
Private Const COUL_ENTETE As Long = 15458502       ' RGB(198, 224, 235) : Bleu Océan éclairci
Private Const COUL_SAISIE As Long = 15201485       ' RGB(205, 244, 231) : Vert Émeraude éclairci
Private Const COUL_CALCUL As Long = 16315885       ' RGB(237, 245, 248) : Bleu Océan très clair
Private Const COUL_ALERTE As Long = 192            ' RGB(192, 0, 0) : messages « non vérifié »

' Module « Vérification au flambement (Mandel) » : ajoute, sous la vérification de
' l'armature de la feuille « Vérification structurelle », l'étude du flambement de
' chaque micropieu de « Calcul Portance » selon l'abaque de Mandel (pieu dans
' un sol à raideur constante). Pointe toujours encastrée ; courbe selon la condition
' en tête (« Vérification structurelle », « Condition en tête ») :
'   tête libre    : courbe 4 (une extrémité encastrée, l'autre libre) ;
'   tête encastrée : courbe 2 (pieu encastré aux deux extrémités).
' Une ligne par micropieu, toutes les valeurs sont des formules :
'   armature : diamètre extérieur d, épaisseur t0 (« Données du projet ») ;
'   diamètre du forage B (Phi retenu de « Calcul Portance ») ;
'   enrobage c = (B - d) / 2, rayon de corrosion retenu (0 si protection par le coulis
'   efficace : c >= 50 mm, C > 500 kg/m3, E/C < 0,5, compression), d après corrosion,
'   inertie corrodée I, raideur de flexion EI = E x I ;
'   Kf (NF P94-262, annexe I, B0 = 0,60 m), seulement pour la durée des charges de
'   « Données du projet » : courte durée
'   12 EM / [4/3 (B0/B)(2,65 B/B0)^alpha + alpha] (B > B0) ou
'   12 EM / [4/3 (2,65)^alpha + alpha] (B <= B0) ; pénétrométrique beta qc ;
'   longue durée (« Données du projet », durée des charges) : Kf / 2 ;
'   demi-longueur l = D / 2 (D profondeur réelle), lambda = l (Kf / EI)^(1/4) ;
'   phi, courbe 4 : lambda voisin de 0 : pi²/(16 lambda²) + (16 lambda²/pi²)(3 - 8/pi) ;
'     lambda assez grand : 1 + (4/racine 3)(2 + cos 2 lambda racine 3) / e^(2 lambda) ;
'     formule de lambda voisin de 0 retenue jusqu'à lambda = 1 ;
'   phi, courbe 2 : lambda voisin de 0 : pi²/lambda² + 3 lambda²/pi² ;
'     lambda assez grand : 2 + (pi²/lambda²)(1 - |sin 2 lambda| / lambda)² ;
'     formule de lambda voisin de 0 retenue jusqu'à lambda = 1,75 ;
'   charge critique Ncr = phi racine(EI Kf), charge admissible Fcr = Ncr / F (F = 2,5
'   à l'ELS, saisi) ; effort ELS dans l'axe N = max(Vmax ELS CARA, ELS QP) / cos(alpha)
'   (micropieu incliné) ; condition N / Fcr <= 1.
' Appelée à la fin de VerificationArmature ; peut être relancée seule. Lance ensuite
' le dimensionnement des platines (module VerificationPlatine), en dessous.

Private Const NOM_FEUILLE As String = "Vérification structurelle"
Private Const DONNEES As String = "'Données du projet'!"
Private Const CALCUL As String = "'Calcul Portance'!"
Private Const TITRE As String = "VÉRIFICATION AU FLAMBEMENT DES MICROPIEUX - ABAQUE DE MANDEL"
Private Const LIB_F As String = "Coefficient de sécurité sur la charge critique à l'ELS (à saisir)"

Public Sub VerificationMandel()
    Dim ws As Worksheet, wc As Worksheet, co As ChartObject, v As Variant, etape As String
    Dim r0 As Long, m As Long, i As Long, k As Long, r As Long, rT As Long, rH As Long, rL As Long
    Dim cPhi As Long, cDr As Long, cIn As Long, cAn As Long, cC As Long, bas As Double, fSec As Variant
    Dim a As Long, lp As Long, lm As Long, tube As String, dA As String, dI As String, t0 As String
    Dim eA As String, rc As String, pen As String, emq As String, alp As String, nat As String
    Dim bet As String, b0 As String, sol As String, coulis As String, vF As String, x As String
    Dim entetes As Variant, formules As Variant, unites As Variant, ci As Long, fr As Range
    Dim enc As String, lt As String, dur As String, longue As Boolean

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    Application.Run "RenommerFeuilleCalcul"   ' ancienne feuille « Calcul et vérification »
    Set wc = ActiveWorkbook.Worksheets("Calcul Portance")
    On Error GoTo Erreur
    If ws Is Nothing Or wc Is Nothing Then
        MsgBox "Créez d'abord les feuilles « Calcul Portance » et « " & NOM_FEUILLE & " ».", vbExclamation
        Exit Sub
    End If
    a = LIGNE_ACIER: lp = LIGNE_PIEU: lm = LIGNE_COUCHE_MOLLE

    ' --- Micropieux de la vérification de la profondeur (« Calcul Portance ») ---
    etape = "micropieux de « Calcul Portance »"
    v = Application.Match("N° MICROPIEU", wc.Range("A1:A5000"), 0)
    If IsError(v) Then
        MsgBox "Tableau des micropieux introuvable dans « Calcul Portance ».", vbExclamation
        Exit Sub
    End If
    r0 = CLng(v)
    Do While Left(CStr(wc.Range("A" & r0 + 3 + m).Value), 2) = "MP"
        m = m + 1
    Loop
    For i = 1 To 80
        If CStr(wc.Cells(r0 + 1, i).Value) = ChrW(934) & " RETENU" Then cPhi = i
        If CStr(wc.Cells(r0, i).Value) = "PROFONDEUR RÉELLE (m)" Then cDr = i
        If cDr = 0 And CStr(wc.Cells(r0, i).Value) = "PROFONDEUR RETENUE (m)" Then cDr = i
        If CStr(wc.Cells(r0 + 1, i).Value) = "Incliné" Then cIn = i
        If CStr(wc.Cells(r0 + 1, i).Value) = "Angle (°)" Then cAn = i
        If CStr(wc.Cells(r0, i).Value) = "COMBINAISONS DE CHARGES VERTICALES (kN)" Then cC = i
    Next i
    If m = 0 Then Exit Sub

    ' --- Position : sous la vérification de l'armature (ancien tableau effacé) ---
    etape = "position du tableau"
    v = Application.Match(LIB_F, ws.Range("B1:B5000"), 0)
    If Not IsError(v) Then fSec = ws.Range("D" & CLng(v)).Value      ' saisie conservée
    v = Application.Match(TITRE, ws.Range("A1:A5000"), 0)
    If Not IsError(v) Then ws.Range(ws.Rows(CLng(v)), ws.Rows(CLng(v) + m + 20)).Clear
    ' tableau des platines (placé ensuite en dessous) effacé lui aussi
    v = Application.Match(TITRE_PLATINE, ws.Range("A1:A5000"), 0)
    If Not IsError(v) Then ws.Range(ws.Rows(CLng(v)), ws.Rows(CLng(v) + m + 30)).Clear
    Set fr = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    r = 1
    If Not fr Is Nothing Then r = fr.Row      ' dernière ligne remplie (vérification de l'armature)
    bas = ws.Cells(r, 1).Top
    For Each co In ws.ChartObjects
        If co.Top + co.Height > bas Then bas = co.Top + co.Height
    Next co
    rT = r
    Do While ws.Cells(rT, 1).Top < bas + 10
        rT = rT + 1
    Loop
    rT = rT + 2

    ' --- Paramètres (« Données du projet ») ---
    tube = DONNEES & "$D$" & a & "=""Tube N80"""
    dA = "N(" & DONNEES & "$D$" & a + 2 & ")"
    t0 = "N(" & DONNEES & "$D$" & a + 3 & ")"
    dI = "IF(" & tube & ",N(" & DONNEES & "$D$" & a + 4 & "),0)"
    eA = "N(" & DONNEES & "$D$" & a + 6 & ")"
    rc = "N(" & DONNEES & "$D$" & a + 13 & ")"
    pen = "(" & DONNEES & CELLULE_METHODE & "=""Méthode pénétrométrique"")"
    emq = "N(" & DONNEES & "$D$" & lm + 1 & ")"
    alp = "N(" & DONNEES & "$D$" & lm + 6 & ")"
    nat = DONNEES & "$D$" & lm + 5
    bet = "IF(OR(" & nat & "=""Tourbe""," & nat & "=""Argile""),12,IF(" & nat & "=""Limon"",7.5,4.5))"
    b0 = "IF(N(" & DONNEES & "$D$" & lp + 22 & ")>0,N(" & DONNEES & "$D$" & lp + 22 & "),0.6)"
    sol = DONNEES & "$D$" & lp + 21
    ' Condition en tête (« Vérification structurelle ») : encastrée -> courbe 2, libre ->
    ' courbe 4 ; à défaut, « Données du projet » (conditions en tête de la fondation)
    v = Application.Match("Condition en tête", ws.Range("B1:B500"), 0)
    If Not IsError(v) Then
        enc = "($D$" & CLng(v) & "=""Micropieu encastré en tête"")"
    Else
        enc = "(" & DONNEES & "$D$" & lp + 18 & "=""Encastré en tête"")"
    End If
    ' Durée des charges (« Données du projet ») : seul le Kf de cette durée est calculé
    longue = (CStr(ActiveWorkbook.Worksheets("Données du projet").Range("D" & lp + 23).Value) <> "Courte durée")
    dur = "IF(" & DONNEES & "$D$" & lp + 23 & "=""Courte durée"",1,2)"
    lt = "IF(" & enc & ",1.75,1)"             ' limite de la formule pour lambda voisin de 0
    coulis = "N(" & DONNEES & "D" & lp + 28 & ")>500,N(" & DONNEES & "D" & lp + 27 & ")>0,N(" & DONNEES & "D" & _
        lp + 27 & ")<0.5," & DONNEES & "D" & lp + 12 & "<>""Traction"""

    ' --- Titre, rappel de la méthode, coefficient de sécurité ---
    etape = "titre"
    With ws.Range("A" & rT)
        .Value = TITRE
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COUL_TITRE
    End With
    With ws.Range("B" & rT + 1 & ":L" & rT + 1)
        .Merge
        .Value = "Charge critique Ncr = " & ChrW(966) & "." & ChrW(8730) & "(EI.Kf) ; demi-longueur réduite " & _
            ChrW(955) & " = l.(Kf/EI)^(1/4), l = D/2 ; pointe toujours encastrée. Tête libre : courbe 4, " & _
            ChrW(955) & " voisin de 0 : " & ChrW(966) & " = " & ChrW(960) & "²/(16" & ChrW(955) & "²) + (16" & _
            ChrW(955) & "²/" & ChrW(960) & "²)(3 - 8/" & ChrW(960) & ") ; " & ChrW(955) & " assez grand : " & _
            ChrW(966) & " = 1 + (4/" & ChrW(8730) & "3)(2 + cos 2" & ChrW(955) & ChrW(8730) & "3)/e^(2" & _
            ChrW(955) & ") ; première formule jusqu'à " & ChrW(955) & " = 1. Tête encastrée : courbe 2, " & _
            ChrW(955) & " voisin de 0 : " & ChrW(966) & " = " & ChrW(960) & "²/" & ChrW(955) & "² + 3" & _
            ChrW(955) & "²/" & ChrW(960) & "² ; " & ChrW(955) & " assez grand : " & ChrW(966) & " = 2 + (" & _
            ChrW(960) & "²/" & ChrW(955) & "²)(1 - |sin 2" & ChrW(955) & "|/" & ChrW(955) & ")² ; première " & _
            "formule jusqu'à " & ChrW(955) & " = 1,75. Condition : N,ELS / Fcr " & ChrW(8804) & _
            " 1 avec Fcr = Ncr / F."
        .Font.Italic = True
        .WrapText = True
    End With
    ws.Rows(rT + 1).RowHeight = 60
    With ws.Range("M" & rT + 2 & ":S" & rT + 2)
        .Merge
        .Formula = "=IF(" & enc & ",""Micropieu encastré en tête, pointe encastrée : courbe 2""," & _
            """Micropieu libre en tête, pointe encastrée : courbe 4"")"
        .Font.Bold = True
        .Font.Color = COUL_TITRE
    End With
    rL = rT + 2
    ws.Range("B" & rL).Value = LIB_F
    ws.Range("C" & rL).Value = "F"
    With ws.Range("D" & rL)
        .Value = IIf(IsNumeric(fSec) And Not IsEmpty(fSec), fSec, 2.5)
        If N0(.Value) <= 0 Then .Value = 2.5
        .Interior.Color = COUL_SAISIE
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("B" & rL & ":D" & rL).Borders.LineStyle = xlContinuous
    vF = "$D$" & rL

    ' --- Tableau : une ligne par micropieu ---
    etape = "tableau des micropieux"
    rH = rL + 2
    entetes = Array("N° micropieu", "Diamètre extérieur de l'armature", "Épaisseur de l'armature", _
        "Diamètre du forage", "Enrobage", "Rayon de corrosion retenu", "Diamètre après corrosion", _
        "Inertie corrodée", "Raideur de flexion", "Module EM (ou qc) de la couche molle", _
        "Coefficient rhéologique " & ChrW(945) & " (ou " & ChrW(946) & ")", "Kf (durée des charges : " & _
        LCase(CStr(ActiveWorkbook.Worksheets("Données du projet").Range("D" & lp + 23).Value)) & ")", "Profondeur réelle", "Demi-longueur", _
        "Demi-longueur réduite", ChrW(966) & " (" & ChrW(955) & " voisin de 0)", ChrW(966) & " (" & ChrW(955) & _
        " assez grand)", ChrW(966) & " retenu", "Charge critique", "Charge critique admissible", _
        "Effort de compression ELS dans l'axe", "Rapport", "Vérification")
    formules = Array("", "d", "t0", "B (" & ChrW(934) & " retenu)", "c = (B - d)/2", _
        "rc (0 si c " & ChrW(8805) & " 50 mm et coulis conforme)", "dc = d - 2rc", _
        "I = " & ChrW(960) & "/64 (dc^4 - di^4)", "EI = E.I", "EM", ChrW(945), _
        IIf(longue, "Kf = 12 EM / [4/3 (B0/B)(2,65 B/B0)^" & ChrW(945) & " + " & ChrW(945) & "] / 2", _
        "Kf = 12 EM / [4/3 (B0/B)(2,65 B/B0)^" & ChrW(945) & " + " & ChrW(945) & "]"), "D", "l = D/2", _
        ChrW(955) & " = l (Kf/EI)^(1/4)", "courbe 4 ou 2 (voir ci-dessus)", "courbe 4 ou 2 (voir ci-dessus)", _
        ChrW(955) & " " & ChrW(8804) & " 1 (courbe 4) ou 1,75 (courbe 2) : " & ChrW(955) & " voisin de 0", "Ncr = " & ChrW(966) & ChrW(8730) & "(EI.Kf)", _
        "Fcr = Ncr / F", "N = max(Vmax ELS) / cos " & ChrW(945), "N / Fcr", "")
    unites = Array("", "mm", "mm", "m", "mm", "mm", "mm", "mm4", "kN.m²", "MPa", "", "MPa", "m", "m", _
        "", "", "", "", "kN", "kN", "kN", "", "")
    For k = 0 To 22
        ws.Cells(rH, 1 + k).Value = entetes(k)
        ws.Cells(rH + 1, 1 + k).Value = formules(k)
        ws.Cells(rH + 2, 1 + k).Value = unites(k)
    Next k
    For i = 1 To m
        r = rH + 2 + i
        ci = r0 + 2 + i                       ' ligne du micropieu dans « Calcul Portance »
        ws.Range("A" & r).Formula = "=" & CALCUL & "$A$" & ci & "&"""""
        ws.Range("B" & r).Formula = "=" & dA
        ws.Range("C" & r).Formula = "=IF(" & tube & "," & t0 & ",""-"")"
        If cPhi > 0 Then
            ws.Range("D" & r).Formula = "=N(" & CALCUL & wc.Cells(ci, cPhi).Address & ")"
        Else
            ws.Range("D" & r).Formula = "=0"
        End If
        ws.Range("E" & r).Formula = "=MAX(0,(1000*D" & r & "-B" & r & ")/2)"
        ws.Range("F" & r).Formula = "=IF(AND(E" & r & ">=50," & coulis & "),0," & rc & ")"
        ws.Range("G" & r).Formula = "=MAX(0,B" & r & "-2*F" & r & ")"
        ws.Range("H" & r).Formula = "=MAX(0,PI()/64*(G" & r & "^4-" & dI & "^4))"
        ws.Range("I" & r).Formula = "=" & eA & "*H" & r & "*1E-9"
        ws.Range("J" & r).Formula = "=" & emq
        ws.Range("K" & r).Formula = "=IF(" & pen & "," & bet & "," & alp & ")"
        x = "MAX(D" & r & "/" & b0 & ",1)"   ' rho = max(B / B0 ; 1)
        ' Kf de la seule durée choisie (« Données du projet ») : longue durée = Kf / 2
        ws.Range("L" & r).Formula = "=IF(OR(N(J" & r & ")<=0,N(D" & r & ")<=0),"""",IF(" & pen & ",K" & r & "*J" & r & _
            ",12*J" & r & "/(4/3*(2.65*" & x & ")^K" & r & "/" & x & "+K" & r & "))/" & dur & ")"
        If cDr > 0 Then
            ws.Range("M" & r).Formula = "=N(" & CALCUL & wc.Cells(ci, cDr).Address & ")"
        Else
            ws.Range("M" & r).Formula = "=0"
        End If
        ws.Range("N" & r).Formula = "=M" & r & "/2"
        ws.Range("O" & r).Formula = "=IF(OR(L" & r & "="""",N(I" & r & ")<=0),"""",N" & r & "*(1000*L" & r & "/I" & _
            r & ")^0.25)"
        x = "O" & r                           ' lambda
        ws.Range("P" & r).Formula = "=IF(N(" & x & ")<=0,"""",IF(" & enc & ",PI()^2/" & x & "^2+3*" & x & "^2/PI()^2," & _
            "PI()^2/(16*" & x & "^2)+16*" & x & "^2/PI()^2*(3-8/PI())))"
        ws.Range("Q" & r).Formula = "=IF(N(" & x & ")<=0,"""",IF(" & enc & ",2+PI()^2/" & x & "^2*(1-ABS(SIN(2*" & x & _
            "))/" & x & ")^2,1+4/SQRT(3)*(2+COS(2*" & x & "*SQRT(3)))/EXP(2*" & x & ")))"
        ws.Range("R" & r).Formula = "=IF(P" & r & "="""","""",IF(" & x & "<=" & lt & ",P" & r & ",Q" & r & "))"
        ws.Range("S" & r).Formula = "=IF(R" & r & "="""","""",R" & r & "*SQRT(I" & r & "*1000*L" & r & "))"
        ws.Range("T" & r).Formula = "=IF(OR(S" & r & "="""",N(" & vF & ")<=0),"""",S" & r & "/" & vF & ")"
        If cC > 0 Then
            ws.Range("U" & r).Formula = "=MAX(ABS(N(" & CALCUL & wc.Cells(ci, cC).Address & ")),ABS(N(" & CALCUL & _
                wc.Cells(ci, cC + 2).Address & ")))" & IIf(cIn > 0 And cAn > 0, "/IF(AND(" & CALCUL & _
                wc.Cells(ci, cIn).Address & "=""Oui"",N(" & CALCUL & wc.Cells(ci, cAn).Address & ")>0,N(" & CALCUL & _
                wc.Cells(ci, cAn).Address & ")<90),COS(RADIANS(" & CALCUL & wc.Cells(ci, cAn).Address & ")),1)", "")
        Else
            ws.Range("U" & r).Formula = "=0"
        End If
        ws.Range("V" & r).Formula = "=IF(N(T" & r & ")>0,U" & r & "/T" & r & ","""")"
        ws.Range("W" & r).Formula = "=IF(V" & r & "="""",""-"",IF(V" & r & "<=1,""OK (taux ""&TEXT(V" & r & _
            ",""0%"")&"")"",""Non vérifié : flambement""))"
    Next i

    ' --- Mise en forme ---
    etape = "mise en forme"
    r = rH + 2 + m
    With ws.Range("A" & rH & ":W" & r)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A" & rH & ":W" & rH + 2)
        .Font.Bold = True
        .Interior.Color = COUL_ENTETE
    End With
    ws.Range("A" & rH + 1 & ":W" & rH + 1).Font.Bold = False
    ws.Range("A" & rH + 1 & ":W" & rH + 1).Font.Italic = True
    ws.Range("B" & rH + 3 & ":W" & r).Interior.Color = COUL_CALCUL
    ws.Range("A" & rH + 3 & ":A" & r & ",W" & rH + 3 & ":W" & r).Font.Bold = True
    ws.Range("B" & rH + 3 & ":C" & r & ",E" & rH + 3 & ":G" & r).NumberFormat = "0.0"
    ws.Range("D" & rH + 3 & ":D" & r & ",M" & rH + 3 & ":R" & r & ",V" & rH + 3 & ":V" & r).NumberFormat = "0.00"
    ws.Range("H" & rH + 3 & ":H" & r).NumberFormat = "0.000E+00"
    ws.Range("I" & rH + 3 & ":J" & r & ",L" & rH + 3 & ":L" & r & ",S" & rH + 3 & ":U" & r).NumberFormat = "0.0"
    ws.Range("K" & rH + 3 & ":K" & r).NumberFormat = "0.000"
    With ws.Range("W" & rH + 3 & ":W" & r).FormatConditions
        .Delete
        .Add(xlTextString, String:="Non vérifié", TextOperator:=xlContains).Font.Color = COUL_ALERTE
    End With
    ' Dimensionnement des platines (module VerificationPlatine), juste en dessous
    On Error Resume Next
    Application.Run "VerificationPlatine"
    Exit Sub
Erreur:
    MsgBox "Erreur pendant la vérification au flambement (Mandel), partie « " & etape & " » : " & _
        Err.Description & " (n° " & Err.Number & ")", vbExclamation
End Sub

' Valeur numérique (0 si vide ou texte)
Private Function N0(v As Variant) As Double
    If Not IsEmpty(v) And IsNumeric(v) Then N0 = CDbl(v)
End Function
