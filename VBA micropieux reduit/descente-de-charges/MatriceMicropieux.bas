Attribute VB_Name = "MatriceMicropieux"
Option Explicit
Private Const COUL_ENTETE As Long = 14337945
Private Const COUL_SAISIE As Long = 15201485
Private Const COUL_CALCUL As Long = 15590349
Public Const LIGNE_MATRICE As Long = 80
Public Const TITRE_MATRICE As String = "RÉPARTITION DES CHARGES DES MURS SUR LES MICROPIEUX"
Public Const CELLULE_NB_MICROPIEUX As String = "K1"
Public Sub ConstruireMatrice(Optional ws As Worksheet)
    Dim n As Long, m As Long, i As Long, j As Long, L As Long, r0 As Long, rG As Long
    Dim nOld As Long, mOld As Long, r0Old As Long, x0 As Long, r As Long, derniere As Long
    Dim li() As Variant, mat() As Variant, gc() As Variant, dims As Variant, dimsN() As Variant, eau() As Variant
    Dim ligne As String, cG As String, cQ As String, c As Range, evenements As Boolean
    Dim nappe As Boolean, f As Worksheet, xf As Long, xt As Long, talus() As Variant, cT As Long
    Dim pot() As Variant, x0Old As Long, cP As String, ligneP As String, cLo As Long, cIo As Long
    Dim typeF As String, isolee As Boolean, nom As String, nL As Long, nP As Long, fG As String, fQ As String
    Dim nE As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    L = LigneMatrice(ws)
    n = Val(ws.Range("G1").Value)
    If n < 1 Then n = 1
    m = Val(ws.Range(CELLULE_NB_MICROPIEUX).Value)
    If m < 1 Then m = 1
    evenements = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Fin
    ReDim li(1 To n)
    ReDim mat(1 To m, 1 To n)
    ReDim gc(1 To m)
    ReDim eau(1 To m, 1 To 3)
    ReDim talus(1 To m, 1 To 2)
    ReDim pot(1 To m, 1 To n)
    dims = ws.Range("C" & L + 1 & ":E" & L + 3).Value
    If Left(CStr(ws.Range("B" & L + 1).Value), 7) <> "Élément" Then
        ReDim dimsN(1 To 3, 1 To 3)
        dimsN(2, 1) = dims(1, 1): dimsN(2, 2) = dims(2, 1): dimsN(2, 3) = dims(2, 3)
        dimsN(3, 1) = dims(1, 1): dimsN(3, 2) = dims(3, 1): dimsN(3, 3) = dims(3, 2)
        dims = dimsN
    End If
    Do While Left(CStr(ws.Range("B" & L + 6 + nOld).Value), 4) = "Mur " Or Left(CStr(ws.Range("B" & L + 6 + nOld).Value), 4) = "Pot "
        nOld = nOld + 1
    Loop
    r0Old = L + 8 + nOld
    If CStr(ws.Range("B" & r0Old).Value) = "N° MICROPIEU" Then
        Do While Left(CStr(ws.Range("B" & r0Old + 3 + mOld).Value), 2) = "MP"
            mOld = mOld + 1
        Loop
        cLo = 0: cIo = 0: x0Old = 0
        For j = 3 To 3 + 2 * nOld
            If Left(CStr(ws.Cells(r0Old, j).Value), 9) = "LONGUEURS" Then cLo = j
            If Left(CStr(ws.Cells(r0Old, j).Value), 14) = "IDENTIFICATION" Then cIo = j
            If Left(CStr(ws.Cells(r0Old, j).Value), 11) = "CHARGES DES" Then x0Old = j: Exit For
        Next j
        If x0Old = 0 Then x0Old = 3 + nOld
        For j = 1 To Application.Min(n, nOld)
            If cLo > 0 Then li(j) = ws.Cells(r0Old + 2, cLo - 1 + j).Value
        Next j
        For i = 1 To Application.Min(m, mOld)
            For j = 1 To Application.Min(n, nOld)
                If cLo > 0 Then
                    Set c = ws.Cells(r0Old + 2 + i, cLo - 1 + j)
                    If Not c.HasFormula Then mat(i, j) = c.Value
                End If
                If cIo > 0 Then pot(i, j) = ws.Cells(r0Old + 2 + i, cIo - 1 + j).Value
            Next j
            gc(i) = ws.Cells(r0Old + 2 + i, x0Old + 2).Value
            For j = 1 To 3
                eau(i, j) = ws.Cells(r0Old + 2 + i, x0Old + 6 + j).Value
            Next j
        Next i
        For j = 3 To 250
            If CStr(ws.Cells(r0Old, j).Value) = "CHARGE DU TALUS" Then cT = j: Exit For
        Next j
        If cT > 0 Then
            For i = 1 To Application.Min(m, mOld)
                talus(i, 1) = ws.Cells(r0Old + 2 + i, cT).Value
                talus(i, 2) = ws.Cells(r0Old + 2 + i, cT + 1).Value
            Next i
        End If
    End If
    derniere = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If derniere < L + 200 Then derniere = L + 200
    ws.Range(ws.Rows(L), ws.Rows(derniere)).Clear
    For Each f In ws.Parent.Worksheets
        If f.Name = "Données du projet" Then nappe = (f.Range("D34").Value = "Oui")
    Next f
    For r = 5 To L - 1
        If CStr(ws.Range("B" & r).Value) = "G (kN/ml)" Or CStr(ws.Range("B" & r).Value) = "G (kN)" Then rG = r
    Next r
    With ws.Range("B" & L)
        .Value = TITRE_MATRICE
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.Range("B" & L + 1 & ":E" & L + 1).Value = Array("Élément", "Poids volumique (kN/m3)", "Base b (m)", "Hauteur h (m)")
    With ws.Range("B" & L + 1 & ":E" & L + 1)
        .Font.Bold = True
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = COUL_ENTETE
    End With
    ws.Range("B" & L + 2).Value = "Massif"
    ws.Range("B" & L + 3).Value = "Longrine"
    For i = 2 To 3
        ws.Range("C" & L + i & ":E" & L + i).Value = Array(IIf(IsEmpty(dims(i, 1)), 25, dims(i, 1)), dims(i, 2), dims(i, 3))
    Next i
    ws.Range("C" & L + 2 & ":E" & L + 3).Interior.Color = COUL_SAISIE
    ws.Range("B" & L + 1 & ":E" & L + 3).Borders.LineStyle = xlContinuous
    typeF = CStr(ws.Range("M1").Value)
    isolee = (typeF = "Semelle isolée")
    nom = IIf(isolee, "Pot ", "Mur ")
    nL = IIf(isolee, 0, n)
    nP = IIf(typeF = "Semelle filante", 0, n)
    nE = IIf(typeF = "Mixte", 5, 4)
    If isolee Then
        ws.Range("B" & L + 5 & ":D" & L + 5).Value = Array("POTEAU", "G (kN)", "Q (kN)")
    ElseIf nE = 5 Then
        ws.Range("B" & L + 5 & ":E" & L + 5).Value = Array("MUR", "G (kN/ml)", "Q (kN/ml)", "G poteaux (kN)")
    Else
        ws.Range("B" & L + 5 & ":D" & L + 5).Value = Array("MUR", "G (kN/ml)", "Q (kN/ml)")
    End If
    For j = 1 To n
        r = L + 5 + j
        ws.Range("B" & r).Value = nom & Format(j, "00")
        If rG > 0 Then
            ws.Range("C" & r).Formula = "=N(" & ws.Cells(rG, 3 + (j - 1) * PAS_MUR).Address & ")"
            ws.Range("D" & r).Formula = "=N(" & ws.Cells(rG + 1, 3 + (j - 1) * PAS_MUR).Address & ")"
            If nE = 5 Then ws.Range("E" & r).Formula = "=N(" & ws.Cells(rG + 2, 3 + (j - 1) * PAS_MUR).Address & ")"
        End If
    Next j
    ws.Range(ws.Cells(L + 6, 3), ws.Cells(L + 5 + n, nE)).Interior.Color = COUL_CALCUL
    cG = "$C$" & L + 6 & ":$C$" & L + 5 + n
    cQ = "$D$" & L + 6 & ":$D$" & L + 5 + n
    cP = "$E$" & L + 6 & ":$E$" & L + 5 + n
    r0 = L + 8 + n
    x0 = 3 + nL + nP
    ws.Cells(r0, 2).Value = "N° MICROPIEU"
    ws.Range(ws.Cells(r0, 2), ws.Cells(r0 + 1, 2)).Merge
    If nL > 0 Then
        ws.Cells(r0, 3).Value = "LONGUEURS DES MURS REPRISES PAR LE MICROPIEU (m)"
        If n > 1 Then ws.Range(ws.Cells(r0, 3), ws.Cells(r0, 2 + n)).Merge
    End If
    If nP > 0 Then
        ws.Cells(r0, 3 + nL).Value = "IDENTIFICATION POTEAU (0 / 1)"
        If n > 1 Then ws.Range(ws.Cells(r0, 3 + nL), ws.Cells(r0, 2 + nL + nP)).Merge
    End If
    ws.Cells(r0, x0).Value = IIf(isolee, "CHARGES DES POTEAUX", "CHARGES DES MURS")
    ws.Range(ws.Cells(r0, x0), ws.Cells(r0, x0 + 1)).Merge
    ws.Cells(r0, x0 + 2).Value = "STRUCTURE GC"
    ws.Cells(r0, x0 + 3).Value = "POIDS MASSIF (kN)"
    ws.Cells(r0, x0 + 4).Value = "POIDS LONGRINE (kN)"
    For i = 2 To 4
        ws.Range(ws.Cells(r0, x0 + i), ws.Cells(r0 + 1, x0 + i)).Merge
    Next i
    ws.Cells(r0, x0 + 5).Value = "POIDS TOTAL ÉLÉMENTAIRE DDC"
    ws.Range(ws.Cells(r0, x0 + 5), ws.Cells(r0, x0 + 6)).Merge
    xf = x0 + IIf(nappe, 9, 6)
    xt = xf + 1
    ws.Cells(r0, xt).Value = "CHARGE DU TALUS"
    ws.Range(ws.Cells(r0, xt), ws.Cells(r0, xt + 1)).Merge
    ws.Range(ws.Cells(r0 + 1, xt), ws.Cells(r0 + 1, xt + 1)).Value = Array("H0 (kN)", "M0 (kN.m)")
    If nappe Then
        ws.Cells(r0, x0 + 7).Value = "CHARGES DE L'EAU (kN)"
        ws.Range(ws.Cells(r0, x0 + 7), ws.Cells(r0, x0 + 9)).Merge
        ws.Range(ws.Cells(r0 + 1, x0 + 7), ws.Cells(r0 + 1, x0 + 9)).Value = Array("EAU EB", "EAU EH", "EAU EE")
    End If
    For j = 1 To n
        If nL > 0 Then ws.Cells(r0 + 1, 2 + j).Value = "Mur " & Format(j, "00")
        If nP > 0 Then ws.Cells(r0 + 1, 2 + nL + j).Value = nom & Format(j, "00")
    Next j
    ws.Cells(r0 + 1, x0).Value = "G (kN)"
    ws.Cells(r0 + 1, x0 + 1).Value = "Q (kN)"
    ws.Cells(r0 + 1, x0 + 5).Value = "G (kN)"
    ws.Cells(r0 + 1, x0 + 6).Value = "Q (kN)"
    If nL > 0 Then
        ws.Cells(r0 + 2, 2).Value = "Longueur du mur Li (m)"
        For j = 1 To n
            ws.Cells(r0 + 2, 2 + j).Value = li(j)
        Next j
    End If
    For i = 1 To m
        r = r0 + 2 + i
        ws.Cells(r, 2).Value = "MP" & i
        For j = 1 To nL
            If Not IsEmpty(mat(i, j)) Then
                ws.Cells(r, 2 + j).Value = mat(i, j)
            ElseIf i = j Then
                ws.Cells(r, 2 + j).Formula = "=N(" & ws.Cells(r0 + 2, 2 + j).Address & ")"
            Else
                ws.Cells(r, 2 + j).Value = 0
            End If
        Next j
        For j = 1 To nP
            With ws.Cells(r, 2 + nL + j)
                If Not IsEmpty(pot(i, j)) Then
                    .Value = pot(i, j)
                Else
                    .Value = IIf(isolee And i = j, 1, 0)
                End If
                .Validation.Delete
                .Validation.Add xlValidateList, xlValidAlertStop, , "0,1"
            End With
        Next j
        If nL > 0 Then ligne = ws.Range(ws.Cells(r, 3), ws.Cells(r, 2 + n)).Address(False, False)
        If nP > 0 Then ligneP = ws.Range(ws.Cells(r, 3 + nL), ws.Cells(r, 2 + nL + n)).Address(False, False)
        If isolee Then
            fG = "MMULT(" & ligneP & "," & cG & ")": fQ = "MMULT(" & ligneP & "," & cQ & ")"
        Else
            fG = "MMULT(" & ligne & "," & cG & ")": fQ = "MMULT(" & ligne & "," & cQ & ")"
            If nP > 0 Then fG = fG & "+MMULT(" & ligneP & "," & cP & ")"
        End If
        ws.Cells(r, x0).Formula = "=IFERROR(" & fG & ","""")"
        ws.Cells(r, x0 + 1).Formula = "=IFERROR(" & fQ & ","""")"
        With ws.Cells(r, x0 + 2)
            .Validation.Add xlValidateList, xlValidAlertStop, , "M,L"
            .Value = IIf(IsEmpty(gc(i)), "M", gc(i))
        End With
        ws.Cells(r, x0 + 3).Formula = "=IF(" & ws.Cells(r, x0 + 2).Address(False, False) & "=""M""," & "$C$" & L + 2 & "*$D$" & L + 2 & "^2*$E$" & L + 2 & ",0)"
        ws.Cells(r, x0 + 4).Formula = "=IF(" & ws.Cells(r, x0 + 2).Address(False, False) & "=""L""," & "$C$" & L + 3 & "*$D$" & L + 3 & "*$E$" & L + 3 & "*" & IIf(nL > 0, "SUM(" & ligne & ")", "0") & "/2,0)"
        ws.Cells(r, x0 + 5).Formula = "=N(" & ws.Cells(r, x0).Address(False, False) & ")+" & ws.Cells(r, x0 + 3).Address(False, False) & "+" & ws.Cells(r, x0 + 4).Address(False, False)
        ws.Cells(r, x0 + 6).Formula = "=N(" & ws.Cells(r, x0 + 1).Address(False, False) & ")"
        If nappe Then
            For j = 7 To 9
                ws.Cells(r, x0 + j).Value = IIf(IsEmpty(eau(i, j - 6)), 0, eau(i, j - 6))
            Next j
        End If
        ws.Cells(r, xt).Value = IIf(IsEmpty(talus(i, 1)), 0, talus(i, 1))
        ws.Cells(r, xt + 1).Value = IIf(IsEmpty(talus(i, 2)), 0, talus(i, 2))
    Next i
    r = r0 + 3 + m
    ws.Cells(r, 2).Value = "TOTAL"
    For i = 0 To xf - x0
        If i <> 2 Then ws.Cells(r, x0 + i).Formula = "=SUM(" & ws.Range(ws.Cells(r0 + 3, x0 + i), ws.Cells(r - 1, x0 + i)).Address(False, False) & ")"
    Next i
    For i = 0 To 1
        ws.Cells(r, xt + i).Formula = "=SUM(" & ws.Range(ws.Cells(r0 + 3, xt + i), ws.Cells(r - 1, xt + i)).Address(False, False) & ")"
    Next i
    ws.Range(ws.Cells(r, 2), ws.Cells(r, xt + 1)).Font.Bold = True
    With ws.Range(ws.Cells(L + 5, 2), ws.Cells(L + 5 + n, nE))
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(r0, 2), ws.Cells(r, xt + 1))
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With Union(ws.Range(ws.Cells(L + 5, 2), ws.Cells(L + 5, nE)), ws.Range(ws.Cells(r0, 2), ws.Cells(r0 + 1, xt + 1)))
        .Font.Bold = True
        .Interior.Color = COUL_ENTETE
        .WrapText = True
    End With
    If nL > 0 Then ws.Range(ws.Cells(r0 + 2, 3), ws.Cells(r - 1, 2 + nL)).Interior.Color = COUL_SAISIE
    If nP > 0 Then ws.Range(ws.Cells(r0 + 3, 3 + nL), ws.Cells(r - 1, 2 + nL + nP)).Interior.Color = COUL_SAISIE
    ws.Range(ws.Cells(r0 + 3, x0 + 2), ws.Cells(r - 1, x0 + 2)).Interior.Color = COUL_SAISIE
    ws.Range(ws.Cells(r0 + 3, x0), ws.Cells(r, x0 + 1)).Interior.Color = COUL_CALCUL
    ws.Range(ws.Cells(r0 + 3, x0 + 3), ws.Cells(r, x0 + 6)).Interior.Color = COUL_CALCUL
    If nappe Then
        ws.Range(ws.Cells(r0 + 3, x0 + 7), ws.Cells(r - 1, x0 + 9)).Interior.Color = COUL_SAISIE
        ws.Range(ws.Cells(r, x0 + 7), ws.Cells(r, x0 + 9)).Interior.Color = COUL_CALCUL
    End If
    ws.Range(ws.Cells(r0 + 2, 3), ws.Cells(r, xt + 1)).NumberFormat = "0.00"
    If nP > 0 Then ws.Range(ws.Cells(r0 + 3, 3 + nL), ws.Cells(r - 1, 2 + nL + nP)).NumberFormat = "0"
    ws.Range(ws.Cells(r0 + 3, xt), ws.Cells(r - 1, xt + 1)).Interior.Color = COUL_SAISIE
    ws.Range(ws.Cells(r, xt), ws.Cells(r, xt + 1)).Interior.Color = COUL_CALCUL
    ws.Range("C" & L + 2 & ":E" & L + 3).NumberFormat = "0.00"
    ws.Range(ws.Cells(L + 6, 3), ws.Cells(L + 5 + n, nE)).NumberFormat = "0.00"
    ws.Range("C" & L + 2 & ":E" & L + 3).HorizontalAlignment = xlCenter
    ws.Rows(r0).RowHeight = 45
    On Error Resume Next
    Application.Run "ActualiserVerificationProfondeur"
    If Not CBool(Application.Run("ReconstructionMatriceEnCours")) Then Application.Run "ActualiserVerificationStructurelle"
Fin:
    Application.EnableEvents = evenements
    Application.ScreenUpdating = True
End Sub
Public Function TrouverLigneMatrice(ws As Worksheet) As Long
    Dim v As Variant
    v = Application.Match(TITRE_MATRICE, ws.Range("B1:B3000"), 0)
    If Not IsError(v) Then TrouverLigneMatrice = CLng(v)
End Function
Public Function LigneMatrice(ws As Worksheet) As Long
    Dim r As Long
    r = TrouverLigneMatrice(ws)
    If r > 0 Then
        LigneMatrice = r
    Else
        r = ws.Cells(3000, "B").End(xlUp).Row + 2
        LigneMatrice = Application.WorksheetFunction.Max(LIGNE_MATRICE, r)
    End If
End Function
