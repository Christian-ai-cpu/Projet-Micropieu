Attribute VB_Name = "TableauChargesPermanentes_CodeEnDurArray"
Option Explicit
' Version codée en dur avec Array, sans variable.

' Niveau d'étage (RDC à 5) en C1, catégorie du bâtiment (A à G) en E1, nombre de murs en G1, % d'ouvertures des murs en I1, en-tête B3:K4 (G = Charges G, H = Charges Q, I = ΣQ, écrit ChrW(931), J = αn, écrit ChrW(945), K = αn × ΣQ, tout en kN/ml), saisie lignes 5 à 73, total ligne 74.
' G = b × (γ × h + g)  en kN/ml pour un mur (b = largeur reprise en m)
Public Sub CreerTableauChargesPermanentes()
    Application.ScreenUpdating = False       ' pas de scintillement pendant l'exécution
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
        With .Range("G1")                            ' duplication du tableau (DuplicationMurs)
            .Value = 1
            .Validation.Delete
            .Validation.Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, "1", "50"
            .BorderAround xlContinuous, xlMedium
        End With
        .Range("H1").Value = "Pourcentage des ouvertures des murs (%)"
        With .Range("I1")                            ' G du mur × (100 - I1) / 100
            .Value = 0
            .Validation.Delete
            .Validation.Add xlValidateDecimal, xlValidAlertStop, xlBetween, "0", "100"
            .BorderAround xlContinuous, xlMedium
        End With
        .Range("B1:I1").Font.Bold = True
        .Range("B1:I1").WrapText = True
        .Range("B1:I1").VerticalAlignment = xlCenter
        .Range("B3:K3").Value = Array("Désignation de l'ouvrage", "Poids volumique (kN/m3)", _
            "Poids Surfacique (kN/m²)", "Dimensions", "", "Charges permanentes G (kN/ml)", _
            "Charge d'exploitation Q (kN/ml)", ChrW(931) & "Q (kN/ml)", ChrW(945) & "n", ChrW(945) & "n × " & ChrW(931) & "Q (kN/ml)")
        .Range("E4:F4").Value = Array("Largeur b (m)", "Hauteur h (m)")
        .Range("B3:B4,C3:C4,D3:D4,E3:F3,G3:G4,H3:H4,I3:I4,J3:J4,K3:K4").Merge
        With .Range("B3:K4")
            .Font.Bold = True
            .WrapText = True
            .RowHeight = 30
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Range("G5:G73").FormulaR1C1 = "=IF(RC[-5]="""","""",RC[-2]*(RC[-4]*RC[-1]+RC[-3]))"   ' références relatives (copies des murs)
        .Range("B74").Value = "TOTAL G"
        .Range("G74").Formula = "=SUBTOTAL(109,G5:G73)"  ' ignore les lignes masquées
        .Range("B74:K74").Font.Bold = True
        .Range("C5:K74").NumberFormat = "0.00"
        .Range("C3:K74").HorizontalAlignment = xlCenter
        .Range("B3:K74").Borders.LineStyle = xlContinuous
        .Range("B3:K74").BorderAround xlContinuous, xlMedium
        .Columns("B").ColumnWidth = 35
        .Columns("C:K").ColumnWidth = 16
        .Columns("G").AutoFit                        ' colonne des charges ajustée au contenu
        If .Columns("G").ColumnWidth < 12 Then .Columns("G").ColumnWidth = 12   ' place pour le titre
    End With
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
