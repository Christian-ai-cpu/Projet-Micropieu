Attribute VB_Name = "VerificationPlatine"
Option Explicit
' Couleurs de la charte graphique GEOSKILLS (voir module Charte)
Private Const COUL_TITRE As Long = 5898259         ' RGB(19, 0, 90) : Bleu Nuit (charte GEOSKILLS)
Private Const COUL_ENTETE As Long = 15458502       ' RGB(198, 224, 235) : Bleu Océan éclairci
Private Const COUL_SAISIE As Long = 15201485       ' RGB(205, 244, 231) : Vert Émeraude éclairci
Private Const COUL_CALCUL As Long = 16315885       ' RGB(237, 245, 248) : Bleu Océan très clair
Private Const COUL_ALERTE As Long = 192            ' RGB(192, 0, 0) : messages « non vérifié »

' Module « Dimensionnement des platines » : ajoute, sous la vérification au flambement
' (module VerificationMandel) de la feuille « Vérification structurelle », le
' dimensionnement de la platine carrée de liaison micropieu / massif de chaque micropieu
' de « Calcul et vérification » (charge N uniformément répartie par la platine).
'   Charges : N = Vmax des combinaisons de « Calcul et vérification » (ELS : maximum
'   des combinaisons caractéristique et quasi permanente ; ELU fondamental ; ELU
'   accidentel).
'   Côté (longueur = largeur) : contrainte dans le béton sigma b = N / A, A = a x a ;
'     a min = racine(N / sigma b,adm) avec sigma b,adm = 0,6 fc28 (ELS),
'     fc28 / gamma c (ELU, gamma c ELU fondamental), fc28 / gamma c (ELUA, gamma c
'     accidentel) ; et a min >= racine(N / fy) (contrainte dans l'acier de la platine,
'     fy de la platine) ; a retenu = maximum, au moins le diamètre du tube, arrondi au
'     cm supérieur.
'   Épaisseur : débord c = (a - d) / 2 (d diamètre extérieur du tube) ;
'     sigma b = N / (a x a) avec le côté retenu ; M = sigma b a c² / 2,
'     I/v = a e² / 6, sigma e = M / (I/v) ; e min = racine(3 sigma b c² / sigma e,adm)
'     avec sigma e,adm = 2/3 fy (ELS), 3/4 fy (ELU), fy (ELUA) ; e retenu = maximum,
'     arrondi au mm supérieur.
'   Platine retenue : dimensions calculées comparées au catalogue de « Données du
'   projet » (liste PLATINES) : platine correspondante ou la plus proche supérieure
'   (côté ap >= a et épaisseur tp >= e) ; au-delà du catalogue : platine calculée.
'   Vérification de la platine retenue : contrainte maximale (charge maximale des
'   combinaisons) sigma e,max = 3 sigma b cp² / tp², sigma b = N / ap²,
'   cp = (ap - d) / 2, comparée à la limite d'élasticité fy de l'acier de la platine.
' Appelée à la fin de VerificationMandel ; peut être relancée seule.

Public Const TITRE_PLATINE As String = "DIMENSIONNEMENT DES PLATINES DE LIAISON MICROPIEU / MASSIF (CHARGE DE CHAQUE MICROPIEU)"
Private Const NOM_FEUILLE As String = "Vérification structurelle"
Private Const DONNEES As String = "'Données du projet'!"
Private Const CALCUL As String = "'Calcul et vérification'!"
Private Const LIB_FC28 As String = "Résistance caractéristique du béton du massif fc28 (à saisir)"

Public Sub VerificationPlatine()
    Dim ws As Worksheet, wc As Worksheet, co As ChartObject, fr As Range, v As Variant, etape As String
    Dim r0 As Long, m As Long, i As Long, k As Long, r As Long, rT As Long, rP As Long, rH As Long
    Dim cC As Long, ci As Long, bas As Double, fcSaisi As Variant, a As Long, la As Long
    Dim vFc As String, vGu As String, vGa As String, vSb1 As String, vSb2 As String, vSb3 As String
    Dim vFy As String, vSe1 As String, vSe2 As String, vSe3 As String, vD As String
    Dim entetes As Variant, formules As Variant, unites As Variant, lib As Variant
    Dim cat As Variant, catD As String, catA As String, catT As String, x As String

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    Set wc = ActiveWorkbook.Worksheets("Calcul et vérification")
    On Error GoTo Erreur
    If ws Is Nothing Or wc Is Nothing Then
        MsgBox "Créez d'abord les feuilles « Calcul et vérification » et « " & NOM_FEUILLE & " ».", vbExclamation
        Exit Sub
    End If
    a = LIGNE_ACIER
    la = LIGNE_PIEU + 24                      ' partie matériaux (coulis) de « Données du projet »

    ' --- Micropieux et combinaisons (« Calcul et vérification ») ---
    etape = "micropieux de « Calcul et vérification »"
    v = Application.Match("N° MICROPIEU", wc.Range("A1:A5000"), 0)
    If IsError(v) Then Exit Sub
    r0 = CLng(v)
    Do While Left(CStr(wc.Range("A" & r0 + 3 + m).Value), 2) = "MP"
        m = m + 1
    Loop
    For i = 1 To 80
        If CStr(wc.Cells(r0, i).Value) = "COMBINAISONS DE CHARGES VERTICALES (kN)" Then cC = i
    Next i
    If m = 0 Or cC = 0 Then Exit Sub

    ' --- Position : sous le dernier tableau (ancien tableau effacé) ---
    etape = "position du tableau"
    v = Application.Match(LIB_FC28, ws.Range("B1:B5000"), 0)
    If Not IsError(v) Then fcSaisi = ws.Range("D" & CLng(v)).Value   ' saisie conservée
    v = Application.Match(TITRE_PLATINE, ws.Range("A1:A5000"), 0)
    If Not IsError(v) Then ws.Range(ws.Rows(CLng(v)), ws.Rows(CLng(v) + m + 30)).Clear
    Set fr = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    r = 1
    If Not fr Is Nothing Then r = fr.Row
    bas = ws.Cells(r, 1).Top
    For Each co In ws.ChartObjects
        If co.Top + co.Height > bas Then bas = co.Top + co.Height
    Next co
    rT = r
    Do While ws.Cells(rT, 1).Top < bas + 10
        rT = rT + 1
    Loop
    rT = rT + 2

    ' --- Titre et paramètres (intitulé B:C, valeur D, unité E) ---
    etape = "paramètres"
    With ws.Range("A" & rT)
        .Value = TITRE_PLATINE
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COUL_TITRE
    End With
    lib = Array( _
        Array(LIB_FC28, "", "MPa"), _
        Array("Coefficient " & ChrW(947) & "c - ELU fondamental (« Données du projet »)", _
            "=IF(N(" & DONNEES & "$D$" & la + 7 & ")>0,N(" & DONNEES & "$D$" & la + 7 & "),1.5)", ""), _
        Array("Coefficient " & ChrW(947) & "c - ELU accidentel (« Données du projet »)", _
            "=IF(N(" & DONNEES & "$D$" & la + 8 & ")>0,N(" & DONNEES & "$D$" & la + 8 & "),1.2)", ""), _
        Array(ChrW(963) & "b,adm ELS = 0,6 fc28", "", "MPa"), _
        Array(ChrW(963) & "b,adm ELU = fc28 / " & ChrW(947) & "c", "", "MPa"), _
        Array(ChrW(963) & "b,adm ELUA = fc28 / " & ChrW(947) & "c", "", "MPa"), _
        Array("Limite d'élasticité de la platine fy (« Données du projet »)", _
            "=N(" & DONNEES & "$D$" & a + 18 & ")", "MPa"), _
        Array(ChrW(963) & "e,adm ELS = 2/3 fy", "", "MPa"), _
        Array(ChrW(963) & "e,adm ELU = 3/4 fy", "", "MPa"), _
        Array(ChrW(963) & "e,adm ELUA = fy", "", "MPa"), _
        Array("Diamètre extérieur du tube d (« Données du projet »)", "=N(" & DONNEES & "$D$" & a + 2 & ")", "mm"))
    For k = 0 To UBound(lib)
        r = rT + 1 + k
        ws.Range("B" & r).Value = lib(k)(0)
        ws.Range("B" & r & ":C" & r).Merge
        If lib(k)(1) <> "" Then ws.Range("D" & r).Formula = lib(k)(1)
        ws.Range("E" & r).Value = lib(k)(2)
    Next k
    vFc = "$D$" & rT + 1: vGu = "$D$" & rT + 2: vGa = "$D$" & rT + 3
    vSb1 = "$D$" & rT + 4: vSb2 = "$D$" & rT + 5: vSb3 = "$D$" & rT + 6
    vFy = "$D$" & rT + 7: vSe1 = "$D$" & rT + 8: vSe2 = "$D$" & rT + 9: vSe3 = "$D$" & rT + 10
    vD = "$D$" & rT + 11
    With ws.Range(vFc)
        .Value = IIf(N0(fcSaisi) > 0, N0(fcSaisi), 25)
        .Interior.Color = COUL_SAISIE
    End With
    ws.Range(vSb1).Formula = "=0.6*" & vFc
    ws.Range(vSb2).Formula = "=" & vFc & "/" & vGu
    ws.Range(vSb3).Formula = "=" & vFc & "/" & vGa
    ws.Range(vSe1).Formula = "=2/3*" & vFy
    ws.Range(vSe2).Formula = "=3/4*" & vFy
    ws.Range(vSe3).Formula = "=" & vFy
    With ws.Range("B" & rT + 1 & ":E" & rT + 11)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("D" & rT + 1 & ":E" & rT + 11).HorizontalAlignment = xlCenter
    ws.Range("D" & rT + 2 & ":D" & rT + 11).Interior.Color = COUL_CALCUL
    ws.Range("D" & rT + 1 & ":D" & rT + 11).NumberFormat = "0.00"
    ws.Range(vD).NumberFormat = "0.0"

    ' --- Catalogue des platines (« Données du projet », liste PLATINES) : désignations,
    ' côtés a et épaisseurs t en constantes de tableau, dans l'ordre croissant ---
    etape = "catalogue des platines"
    cat = Split(PLATINES, ",")
    catD = "{": catA = "{": catT = "{"
    For k = 0 To UBound(cat)
        catD = catD & IIf(k > 0, ",", "") & """" & cat(k) & """"
        catA = catA & IIf(k > 0, ",", "") & Val(Mid(cat(k), InStr(cat(k), "carrée") + 7))
        catT = catT & IIf(k > 0, ",", "") & Val(Mid(cat(k), InStr(cat(k), "ép.") + 3))
    Next k
    catD = catD & "}": catA = catA & "}": catT = catT & "}"

    ' --- Tableau : une ligne par micropieu ---
    etape = "tableau des platines"
    rH = rT + 14
    entetes = Array("N° micropieu", "Charge ELS", "Charge ELU fondamentale", "Charge ELU accidentelle", _
        "Côté minimal ELS", "Côté minimal ELU", "Côté minimal ELUA", "Longueur = largeur calculée", "Débord", _
        ChrW(963) & "b ELS", ChrW(963) & "b ELU", ChrW(963) & "b ELUA", "Épaisseur minimale ELS", _
        "Épaisseur minimale ELU", "Épaisseur minimale ELUA", "Épaisseur calculée", "Platine calculée", _
        "Platine retenue (catalogue de « Données du projet »)", "Côté de la platine retenue", _
        "Épaisseur de la platine retenue", "Contrainte maximale dans la platine", _
        "Limite d'élasticité de l'acier de la platine", "Rapport", "Vérification")
    formules = Array("", "N = max(Vmax ELS CARA ; ELS QP)", "N = Vmax ELU FOND", "N = Vmax ELU ACC", _
        "a = " & ChrW(8730) & "(N / " & ChrW(963) & "b,adm)", "a = max(" & ChrW(8730) & "(N/" & ChrW(963) & _
        "b,adm) ; " & ChrW(8730) & "(N/fy))", "a = max(" & ChrW(8730) & "(N/" & ChrW(963) & "b,adm) ; " & _
        ChrW(8730) & "(N/fy))", "a = max(côtés minimaux ; d), au cm supérieur", "c = (a - d)/2", _
        ChrW(963) & "b = N / (a.a)", ChrW(963) & "b = N / (a.a)", ChrW(963) & "b = N / (a.a)", _
        "e = " & ChrW(8730) & "(3" & ChrW(963) & "b.c² / " & ChrW(963) & "e,adm)", "e = " & ChrW(8730) & _
        "(3" & ChrW(963) & "b.c² / " & ChrW(963) & "e,adm)", "e = " & ChrW(8730) & "(3" & ChrW(963) & _
        "b.c² / " & ChrW(963) & "e,adm)", "e = max, au mm supérieur", "a x a x e", _
        "platine correspondante ou la plus proche supérieure (ap " & ChrW(8805) & " a et tp " & ChrW(8805) & _
        " e) ; au-delà du catalogue : platine calculée", "ap", "tp", ChrW(963) & "e,max = 3" & ChrW(963) & _
        "b.cp² / tp², " & ChrW(963) & "b = Nmax / ap², cp = (ap - d)/2", "fy", ChrW(963) & "e,max / fy", _
        ChrW(963) & "e,max " & ChrW(8804) & " fy")
    unites = Array("", "kN", "kN", "kN", "mm", "mm", "mm", "mm", "mm", "MPa", "MPa", "MPa", "mm", "mm", "mm", _
        "mm", "mm", "", "mm", "mm", "MPa", "MPa", "", "")
    For k = 0 To 23
        ws.Cells(rH, 1 + k).Value = entetes(k)
        ws.Cells(rH + 1, 1 + k).Value = formules(k)
        ws.Cells(rH + 2, 1 + k).Value = unites(k)
    Next k
    For i = 1 To m
        r = rH + 2 + i
        ci = r0 + 2 + i                       ' ligne du micropieu dans « Calcul et vérification »
        ws.Range("A" & r).Formula = "=" & CALCUL & "$A$" & ci & "&"""""
        ws.Range("B" & r).Formula = "=MAX(ABS(N(" & CALCUL & wc.Cells(ci, cC).Address & ")),ABS(N(" & CALCUL & _
            wc.Cells(ci, cC + 2).Address & ")))"
        ws.Range("C" & r).Formula = "=ABS(N(" & CALCUL & wc.Cells(ci, cC + 4).Address & "))"
        ws.Range("D" & r).Formula = "=ABS(N(" & CALCUL & wc.Cells(ci, cC + 6).Address & "))"
        ' côtés minimaux (N en kN -> N x 1000 en N ; contraintes en MPa = N/mm²)
        ws.Range("E" & r).Formula = "=IF(N(" & vSb1 & ")>0,SQRT(1000*B" & r & "/" & vSb1 & "),0)"
        ws.Range("F" & r).Formula = "=MAX(IF(N(" & vSb2 & ")>0,SQRT(1000*C" & r & "/" & vSb2 & "),0),IF(N(" & vFy & _
            ")>0,SQRT(1000*C" & r & "/" & vFy & "),0))"
        ws.Range("G" & r).Formula = "=MAX(IF(N(" & vSb3 & ")>0,SQRT(1000*D" & r & "/" & vSb3 & "),0),IF(N(" & vFy & _
            ")>0,SQRT(1000*D" & r & "/" & vFy & "),0))"
        ws.Range("H" & r).Formula = "=CEILING(MAX(E" & r & ",F" & r & ",G" & r & "," & vD & ",10),10)"
        ws.Range("I" & r).Formula = "=MAX(0,(H" & r & "-" & vD & ")/2)"
        ws.Range("J" & r).Formula = "=1000*B" & r & "/H" & r & "^2"
        ws.Range("K" & r).Formula = "=1000*C" & r & "/H" & r & "^2"
        ws.Range("L" & r).Formula = "=1000*D" & r & "/H" & r & "^2"
        ws.Range("M" & r).Formula = "=IF(N(" & vSe1 & ")>0,SQRT(3*J" & r & "*I" & r & "^2/" & vSe1 & "),"""")"
        ws.Range("N" & r).Formula = "=IF(N(" & vSe2 & ")>0,SQRT(3*K" & r & "*I" & r & "^2/" & vSe2 & "),"""")"
        ws.Range("O" & r).Formula = "=IF(N(" & vSe3 & ")>0,SQRT(3*L" & r & "*I" & r & "^2/" & vSe3 & "),"""")"
        ws.Range("P" & r).Formula = "=CEILING(MAX(1,M" & r & ",N" & r & ",O" & r & "),1)"
        ws.Range("Q" & r).Formula = "=H" & r & "&"" x ""&H" & r & "&"" x ""&P" & r
        ' platine choisie : « Platine carrée a x a mm - ép. t mm »
        ' platine retenue : plus petite platine du catalogue couvrant a et e (correspondante
        ' ou la plus proche supérieure) ; au-delà du catalogue : platine calculée
        x = "MATCH(1,INDEX((" & catA & ">=H" & r & ")*(" & catT & ">=P" & r & "),0),0)"
        ws.Range("R" & r).Formula = "=IFERROR(INDEX(" & catD & "," & x & "),""Platine calculée ""&Q" & r & _
            "&"" mm (au-delà du catalogue)"")"
        ws.Range("S" & r).Formula = "=IFERROR(INDEX(" & catA & "," & x & "),H" & r & ")"
        ws.Range("T" & r).Formula = "=IFERROR(INDEX(" & catT & "," & x & "),P" & r & ")"
        ' contrainte maximale dans la platine retenue (charge maximale des combinaisons) :
        ' sigma e = M / (I/v) = 3 sigma b cp² / tp², sigma b = N / ap², cp = (ap - d) / 2
        ws.Range("U" & r).Formula = "=IF(OR(N(S" & r & ")<=0,N(T" & r & ")<=0),"""",3*(1000*MAX(B" & r & ",C" & r & _
            ",D" & r & ")/S" & r & "^2)*MAX(0,(S" & r & "-" & vD & ")/2)^2/T" & r & "^2)"
        ws.Range("V" & r).Formula = "=" & vFy
        ws.Range("W" & r).Formula = "=IF(OR(U" & r & "="""",N(V" & r & ")<=0),"""",U" & r & "/V" & r & ")"
        ws.Range("X" & r).Formula = "=IF(W" & r & "="""",""-"",IF(W" & r & "<=1,""OK (taux ""&TEXT(W" & r & _
            ",""0%"")&"")"",""Non vérifiée : augmenter la platine""))"
    Next i

    ' --- Mise en forme ---
    etape = "mise en forme"
    r = rH + 2 + m
    With ws.Range("A" & rH & ":X" & r)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A" & rH & ":X" & rH + 2)
        .Font.Bold = True
        .Interior.Color = COUL_ENTETE
    End With
    ws.Range("A" & rH + 1 & ":X" & rH + 1).Font.Bold = False
    ws.Range("A" & rH + 1 & ":X" & rH + 1).Font.Italic = True
    ws.Range("B" & rH + 3 & ":X" & r).Interior.Color = COUL_CALCUL
    ws.Range("A" & rH + 3 & ":A" & r & ",H" & rH + 3 & ":H" & r & ",P" & rH + 3 & ":T" & r & ",X" & rH + 3 & _
        ":X" & r).Font.Bold = True
    ws.Range("B" & rH + 3 & ":G" & r & ",I" & rH + 3 & ":I" & r & ",M" & rH + 3 & ":O" & r & ",U" & rH + 3 & _
        ":V" & r).NumberFormat = "0.0"
    ws.Range("J" & rH + 3 & ":L" & r & ",W" & rH + 3 & ":W" & r).NumberFormat = "0.00"
    ws.Range("H" & rH + 3 & ":H" & r & ",P" & rH + 3 & ":P" & r & ",S" & rH + 3 & ":T" & r).NumberFormat = "0"
    With ws.Range("X" & rH + 3 & ":X" & r).FormatConditions
        .Delete
        .Add(xlTextString, String:="Non vérifiée", TextOperator:=xlContains).Font.Color = COUL_ALERTE
    End With
    Exit Sub
Erreur:
    MsgBox "Erreur pendant le dimensionnement des platines, partie « " & etape & " » : " & _
        Err.Description & " (n° " & Err.Number & ")", vbExclamation
End Sub

' Valeur numérique (0 si vide ou texte)
Private Function N0(v As Variant) As Double
    If Not IsEmpty(v) And IsNumeric(v) Then N0 = CDbl(v)
End Function
