Attribute VB_Name = "VerificationPlatine"
Option Explicit
Private Const COUL_TITRE As Long = 5898259
Private Const COUL_ENTETE As Long = 14337945
Private Const COUL_SAISIE As Long = 15201485
Private Const COUL_CALCUL As Long = 15590349
Private Const COUL_ALERTE As Long = 192
Public Const TITRE_PLATINE As String = "DIMENSIONNEMENT DES PLATINES DE LIAISON MICROPIEU / MASSIF (CHARGE DE CHAQUE MICROPIEU)"
Private Const NOM_FEUILLE As String = "Vérification structurelle"
Private Const DONNEES As String = "'Données du projet'!"
Private Const CALCUL As String = "'Calcul Portance'!"
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
    Application.Run "RenommerFeuilleCalcul"
    Set wc = ActiveWorkbook.Worksheets("Calcul Portance")
    On Error GoTo Erreur
    If ws Is Nothing Or wc Is Nothing Then
        MsgBox "Créez d'abord les feuilles « Calcul Portance » et « " & NOM_FEUILLE & " ».", vbExclamation
        Exit Sub
    End If
    a = LIGNE_ACIER
    la = LIGNE_PIEU + 24
    etape = "micropieux de « Calcul Portance »"
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
    etape = "position du tableau"
    v = Application.Match(LIB_FC28, ws.Range("B1:B5000"), 0)
    If Not IsError(v) Then fcSaisi = ws.Range("D" & CLng(v)).Value
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
    etape = "paramètres"
    With ws.Range("A" & rT)
        .Value = TITRE_PLATINE
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COUL_TITRE
    End With
    lib = Array( Array(LIB_FC28, "", "MPa"), Array("Coefficient " & ChrW(947) & "c - ELU fondamental (« Données du projet »)", "=IF(N(" & DONNEES & "$D$" & la + 7 & ")>0,N(" & DONNEES & "$D$" & la + 7 & "),1.5)", ""), Array("Coefficient " & ChrW(947) & "c - ELU accidentel (« Données du projet »)", "=IF(N(" & DONNEES & "$D$" & la + 8 & ")>0,N(" & DONNEES & "$D$" & la + 8 & "),1.2)", ""), Array(ChrW(963) & "b,adm ELS = 0,6 fc28", "", "MPa"), Array(ChrW(963) & "b,adm ELU = fc28 / " & ChrW(947) & "c", "", "MPa"), Array(ChrW(963) & "b,adm ELUA = fc28 / " & ChrW(947) & "c", "", "MPa"), Array("Limite d'élasticité de la platine fy (« Données du projet »)", "=N(" & DONNEES & "$D$" & a + 18 & ")", "MPa"), Array(ChrW(963) & "e,adm ELS = 2/3 fy", "", "MPa"), Array(ChrW(963) & "e,adm ELU = 3/4 fy", "", "MPa"), Array(ChrW(963) & "e,adm ELUA = fy", "", "MPa"), Array("Diamètre extérieur du tube d (« Données du projet »)", "=N(" & DONNEES & "$D$" & a + 2 & ")", "mm"))
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
    etape = "catalogue des platines"
    cat = Split(PLATINES, ",")
    catD = "{": catA = "{": catT = "{"
    For k = 0 To UBound(cat)
        catD = catD & IIf(k > 0, ",", "") & """" & cat(k) & """"
        catA = catA & IIf(k > 0, ",", "") & Val(Mid(cat(k), InStr(cat(k), "carrée") + 7))
        catT = catT & IIf(k > 0, ",", "") & Val(Mid(cat(k), InStr(cat(k), "ép.") + 3))
    Next k
    catD = catD & "}": catA = catA & "}": catT = catT & "}"
    etape = "tableau des platines"
    rH = rT + 14
    entetes = Array("N° micropieu", "Charge ELS", "Charge ELU fondamentale", "Charge ELU accidentelle", "Côté minimal ELS", "Côté minimal ELU", "Côté minimal ELUA", "Longueur = largeur calculée", "Débord", ChrW(963) & "b ELS", ChrW(963) & "b ELU", ChrW(963) & "b ELUA", "Épaisseur minimale ELS", "Épaisseur minimale ELU", "Épaisseur minimale ELUA", "Épaisseur calculée", "Platine calculée", "Platine retenue (catalogue de « Données du projet »)", "Côté de la platine retenue", "Épaisseur de la platine retenue", "Contrainte maximale dans la platine", "Limite d'élasticité de l'acier de la platine", "Rapport", "Vérification")
    formules = Array("", "N = max(Vmax ELS CARA ; ELS QP)", "N = Vmax ELU FOND", "N = Vmax ELU ACC", "a = " & ChrW(8730) & "(N / " & ChrW(963) & "b,adm)", "a = max(" & ChrW(8730) & "(N/" & ChrW(963) & "b,adm) ; " & ChrW(8730) & "(N/fy))", "a = max(" & ChrW(8730) & "(N/" & ChrW(963) & "b,adm) ; " & ChrW(8730) & "(N/fy))", "a = max(côtés minimaux ; d), au cm supérieur", "c = (a - d)/2", ChrW(963) & "b = N / (a.a)", ChrW(963) & "b = N / (a.a)", ChrW(963) & "b = N / (a.a)", "e = " & ChrW(8730) & "(3" & ChrW(963) & "b.c² / " & ChrW(963) & "e,adm)", "e = " & ChrW(8730) & "(3" & ChrW(963) & "b.c² / " & ChrW(963) & "e,adm)", "e = " & ChrW(8730) & "(3" & ChrW(963) & "b.c² / " & ChrW(963) & "e,adm)", "e = max, au mm supérieur", "a x a x e", "platine correspondante ou la plus proche supérieure (ap " & ChrW(8805) & " a et tp " & ChrW(8805) & " e) ; au-delà du catalogue : platine calculée", "ap", "tp", ChrW(963) & "e,max = 3" & ChrW(963) & _
        "b.cp² / tp², " & ChrW(963) & "b = Nmax / ap², cp = (ap - d)/2", "fy", ChrW(963) & "e,max / fy", ChrW(963) & "e,max " & ChrW(8804) & " fy")
    unites = Array("", "kN", "kN", "kN", "mm", "mm", "mm", "mm", "mm", "MPa", "MPa", "MPa", "mm", "mm", "mm", "mm", "mm", "", "mm", "mm", "MPa", "MPa", "", "")
    For k = 0 To 23
        ws.Cells(rH, 1 + k).Value = entetes(k)
        ws.Cells(rH + 1, 1 + k).Value = formules(k)
        ws.Cells(rH + 2, 1 + k).Value = unites(k)
    Next k
    For i = 1 To m
        r = rH + 2 + i
        ci = r0 + 2 + i
        ws.Range("A" & r).Formula = "=IF(" & CALCUL & "$A$" & ci & "="""",""MP" & i & """," & CALCUL & "$A$" & ci & ")"
        ws.Range("B" & r).Formula = "=MAX(ABS(N(" & CALCUL & wc.Cells(ci, cC).Address & ")),ABS(N(" & CALCUL & wc.Cells(ci, cC + 2).Address & ")))"
        ws.Range("C" & r).Formula = "=ABS(N(" & CALCUL & wc.Cells(ci, cC + 4).Address & "))"
        ws.Range("D" & r).Formula = "=ABS(N(" & CALCUL & wc.Cells(ci, cC + 6).Address & "))"
        ws.Range("E" & r).Formula = "=IF(N(" & vSb1 & ")>0,SQRT(1000*B" & r & "/" & vSb1 & "),0)"
        ws.Range("F" & r).Formula = "=MAX(IF(N(" & vSb2 & ")>0,SQRT(1000*C" & r & "/" & vSb2 & "),0),IF(N(" & vFy & ")>0,SQRT(1000*C" & r & "/" & vFy & "),0))"
        ws.Range("G" & r).Formula = "=MAX(IF(N(" & vSb3 & ")>0,SQRT(1000*D" & r & "/" & vSb3 & "),0),IF(N(" & vFy & ")>0,SQRT(1000*D" & r & "/" & vFy & "),0))"
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
        x = "MATCH(1,INDEX((" & catA & ">=H" & r & ")*(" & catT & ">=P" & r & "),0),0)"
        ws.Range("R" & r).Formula = "=IFERROR(INDEX(" & catD & "," & x & "),""Platine calculée ""&Q" & r & "&"" mm (au-delà du catalogue)"")"
        ws.Range("S" & r).Formula = "=IFERROR(INDEX(" & catA & "," & x & "),H" & r & ")"
        ws.Range("T" & r).Formula = "=IFERROR(INDEX(" & catT & "," & x & "),P" & r & ")"
        ws.Range("U" & r).Formula = "=IF(OR(N(S" & r & ")<=0,N(T" & r & ")<=0),"""",3*(1000*MAX(B" & r & ",C" & r & ",D" & r & ")/S" & r & "^2)*MAX(0,(S" & r & "-" & vD & ")/2)^2/T" & r & "^2)"
        ws.Range("V" & r).Formula = "=" & vFy
        ws.Range("W" & r).Formula = "=IF(OR(U" & r & "="""",N(V" & r & ")<=0),"""",U" & r & "/V" & r & ")"
        ws.Range("X" & r).Formula = "=IF(W" & r & "="""",""-"",IF(W" & r & "<=1,""OK (taux ""&TEXT(W" & r & ",""0%"")&"")"",""Non vérifiée : augmenter la platine""))"
    Next i
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
    ws.Range("A" & rH + 3 & ":A" & r & ",H" & rH + 3 & ":H" & r & ",P" & rH + 3 & ":T" & r & ",X" & rH + 3 & ":X" & r).Font.Bold = True
    ws.Range("B" & rH + 3 & ":G" & r & ",I" & rH + 3 & ":I" & r & ",M" & rH + 3 & ":O" & r & ",U" & rH + 3 & ":V" & r).NumberFormat = "0.0"
    ws.Range("J" & rH + 3 & ":L" & r & ",W" & rH + 3 & ":W" & r).NumberFormat = "0.00"
    ws.Range("H" & rH + 3 & ":H" & r & ",P" & rH + 3 & ":P" & r & ",S" & rH + 3 & ":T" & r).NumberFormat = "0"
    With ws.Range("X" & rH + 3 & ":X" & r).FormatConditions
        .Delete
        .Add(xlTextString, String:="Non vérifiée", TextOperator:=xlContains).Font.Color = COUL_ALERTE
    End With
    Exit Sub
Erreur:
    MsgBox "Erreur pendant le dimensionnement des platines, partie « " & etape & " » : " & Err.Description & " (n° " & Err.Number & ")", vbExclamation
End Sub
Private Function N0(v As Variant) As Double
    If Not IsEmpty(v) And IsNumeric(v) Then N0 = CDbl(v)
End Function
