Attribute VB_Name = "TableauChargesPermanentes_CodeEnDurArray"
Option Explicit
' Version codée en dur avec Array, sans variable.

' Niveau d'étage (RDC à 5) en C1, catégorie du bâtiment (A à G) en E1, en-tête B3:H4, saisie lignes 5 à 48, total ligne 49.
' G = n × b × (γ × h + g)
Public Sub CreerTableauChargesPermanentes()
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
        .Range("B1:E1").Font.Bold = True
        .Range("B3:H3").Value = Array("Désignation de l'ouvrage", "Poids volumique (kN/m3)", _
            "Poids Surfacique (kN/m²)", "Dimensions", "", "", "Charges permanentes G (kN/ml)")
        .Range("E4:G4").Value = Array("Largeur b (m)", "Épaisseur h (m)", "Nombre n")
        .Range("B3:B4,C3:C4,D3:D4,E3:G3,H3:H4").Merge
        With .Range("B3:H4")
            .Font.Bold = True
            .WrapText = True
            .RowHeight = 30
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Range("G5:G48").Value = 1
        .Range("H5:H48").FormulaR1C1 = "=IF(RC2="""","""",RC7*RC5*(RC3*RC6+RC4))"
        .Range("B49").Value = "TOTAL G"
        .Range("H49").Formula = "=SUBTOTAL(109,H5:H48)"  ' ignore les lignes masquées
        .Range("B49:H49").Font.Bold = True
        .Range("C5:H49").NumberFormat = "0.00"
        .Range("G5:G48").NumberFormat = "0"
        .Range("C3:H49").HorizontalAlignment = xlCenter
        .Range("B3:H49").Borders.LineStyle = xlContinuous
        .Range("B3:H49").BorderAround xlContinuous, xlMedium
        .Columns("B").ColumnWidth = 35
        .Columns("C:H").ColumnWidth = 16
    End With
End Sub
