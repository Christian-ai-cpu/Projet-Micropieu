Attribute VB_Name = "TableauChargesPermanentes_CodeEnDurArray"
Option Explicit
' Version codée en dur avec Array, sans variable.

' Niveau d'étage (RDC à 5) en C1, catégorie du bâtiment (A à G) en E1, en-tête B3:J4 (H = Linéaire, I = Charges G, J = Charges Q, K = ΣQ, écrit ChrW(931), L = αn, écrit ChrW(945), M = αn × ΣQ), saisie lignes 5 à 54, total ligne 55.
' G = n × (L / 2) × b × (γ × h + g)   (L = linéaire en H, pris à moitié sauf pour les poteaux)
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
        .Range("B3:M3").Value = Array("Désignation de l'ouvrage", "Poids volumique (kN/m3)", _
            "Poids Surfacique (kN/m²)", "Dimensions", "", "", "Linéaire L (m)", "Charges permanentes G (kN)", _
            "Charge d'exploitation Q (kN)", ChrW(931) & "Q (kN)", ChrW(945) & "n", ChrW(945) & "n × " & ChrW(931) & "Q (kN)")
        .Range("E4:G4").Value = Array("Largeur b (m)", "Épaisseur h (m)", "Nombre n")
        .Range("B3:B4,C3:C4,D3:D4,E3:G3,H3:H4,I3:I4,J3:J4,K3:K4,L3:L4,M3:M4").Merge
        With .Range("B3:M4")
            .Font.Bold = True
            .WrapText = True
            .RowHeight = 30
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Range("G5:G54").Value = 1
        .Range("I5:I54").FormulaR1C1 = "=IF(RC2="""","""",RC7*RC8/IF(LEFT(RC2,6)=""Poteau"",1,2)*RC5*(RC3*RC6+RC4))"
        .Range("B55").Value = "TOTAL G"
        .Range("I55").Formula = "=SUBTOTAL(109,I5:I54)"  ' ignore les lignes masquées
        .Range("B55:M55").Font.Bold = True
        .Range("C5:M55").NumberFormat = "0.00"
        .Range("G5:G54").NumberFormat = "0"
        .Range("C3:M55").HorizontalAlignment = xlCenter
        .Range("B3:M55").Borders.LineStyle = xlContinuous
        .Range("B3:M55").BorderAround xlContinuous, xlMedium
        .Columns("B").ColumnWidth = 35
        .Columns("C:M").ColumnWidth = 16
        .Columns("I").AutoFit                        ' colonne des charges ajustée au contenu
        If .Columns("I").ColumnWidth < 12 Then .Columns("I").ColumnWidth = 12   ' place pour le titre
    End With
End Sub
