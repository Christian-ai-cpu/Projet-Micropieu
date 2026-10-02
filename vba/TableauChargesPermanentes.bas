Attribute VB_Name = "TableauChargesPermanentes"
Option Explicit

' En-tête B2:H3, saisie lignes 4 à 18, total ligne 19.
' G = n × b × (γ × h + g)
Public Sub CreerTableauChargesPermanentes()
    With ActiveSheet
        .Range("B2:H2").Value = Array("Désignation de l'ouvrage", "Poids volumique (kN/m3)", _
            "Poids Surfacique (kN/m²)", "Dimensions", "", "", "Charges permanentes G (kN/ml)")
        .Range("E3:G3").Value = Array("Largeur b (m)", "Épaisseur h (m)", "Nombre n")
        .Range("B2:B3,C2:C3,D2:D3,E2:G2,H2:H3").Merge
        With .Range("B2:H3")
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(31, 78, 121)
            .WrapText = True
            .RowHeight = 30
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Range("G4:G18").Value = 1
        .Range("H4:H18").FormulaR1C1 = "=IF(RC2="""","""",RC7*RC5*(RC3*RC6+RC4))"
        .Range("B19").Value = "TOTAL G"
        .Range("H19").Formula = "=SUM(H4:H18)"
        .Range("B19:H19").Font.Bold = True
        .Range("B19:H19").Interior.Color = RGB(221, 235, 247)
        .Range("C4:H19").NumberFormat = "0.00"
        .Range("G4:G18").NumberFormat = "0"
        .Range("C2:H19").HorizontalAlignment = xlCenter
        .Range("B2:H19").Borders.LineStyle = xlContinuous
        .Range("B2:H19").BorderAround xlContinuous, xlMedium
        .Columns("B").ColumnWidth = 35
        .Columns("C:H").ColumnWidth = 16
    End With
End Sub
