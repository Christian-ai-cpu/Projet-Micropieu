Attribute VB_Name = "TableauChargesPermanentes_CodeEnDurArray"
Option Explicit
Public Sub CreerTableauChargesPermanentes()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    With ActiveSheet
        .Cells.Clear
        .Range("B1").Value = "Niveau d'étage"
        With .Range("C1")
            .Value = "RDC"
            .Validation.Delete
            .Validation.Add xlValidateList, xlValidAlertStop, , "RDC,1,2,3,4,5"
            .BorderAround xlContinuous, xlMedium
        End With
        .Range("D1").Value = "Catégorie du bâtiment"
        With .Range("E1")
            .Value = "A"
            .Validation.Delete
            .Validation.Add xlValidateList, xlValidAlertStop, , "A,B,C,D,E,F,G"
            .BorderAround xlContinuous, xlMedium
        End With
        .Range("F1").Value = "Entrez le nombre de murs"
        With .Range("G1")
            .Value = 1
            .Validation.Delete
            .Validation.Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, "1", "50"
            .BorderAround xlContinuous, xlMedium
        End With
        .Range("H1").Value = "Pourcentage des ouvertures des murs (%)"
        With .Range("I1")
            .Value = 0
            .Validation.Delete
            .Validation.Add xlValidateDecimal, xlValidAlertStop, xlBetween, "0", "100"
            .BorderAround xlContinuous, xlMedium
        End With
        .Range("B1:I1").Font.Bold = True
        .Range("B1:I1").WrapText = True
        .Range("B1:I1").VerticalAlignment = xlCenter
        .Range("B3:K3").Value = Array("Désignation de l'ouvrage", "Poids volumique (kN/m3)", "Poids Surfacique (kN/m²)", "Dimensions", "", "Charges permanentes G (kN/ml)", "Charge d'exploitation Q (kN/ml)", ChrW(931) & "Q (kN/ml)", ChrW(945) & "n", ChrW(945) & "n × " & ChrW(931) & "Q (kN/ml)")
        .Range("E4:F4").Value = Array("Largeur b (m)", "Hauteur h (m)")
        .Range("B3:B4,C3:C4,D3:D4,E3:F3,G3:G4,H3:H4,I3:I4,J3:J4,K3:K4").Merge
        With .Range("B3:K4")
            .Font.Bold = True
            .WrapText = True
            .RowHeight = 30
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Range("G5:G73").FormulaR1C1 = "=IF(RC[-5]="""","""",RC[-2]*(RC[-4]*RC[-1]+RC[-3]))"
        .Range("B74").Value = "TOTAL G"
        .Range("G74").Formula = "=SUBTOTAL(109,G5:G73)"
        .Range("B74:K74").Font.Bold = True
        .Range("C5:K74").NumberFormat = "0.00"
        .Range("C3:K74").HorizontalAlignment = xlCenter
        .Range("B3:K74").Borders.LineStyle = xlContinuous
        .Range("B3:K74").BorderAround xlContinuous, xlMedium
        .Columns("B").ColumnWidth = 48
        .Columns("C:K").ColumnWidth = 16
        .Columns("G").AutoFit
        If .Columns("G").ColumnWidth < 12 Then .Columns("G").ColumnWidth = 12
    End With
    ActiveSheet.Cells.Font.Name = "CheGuevara Text Sans"
    AjusterTextes ActiveSheet
    Application.EnableEvents = True
    Application.ScreenUpdating = True
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
