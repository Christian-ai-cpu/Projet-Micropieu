Attribute VB_Name = "TableauChargesPermanentesComplet"
Option Explicit
Private Const COUL_SAISIE As Long = 15201485
Private Const COUL_CALCUL As Long = 15590349
Private Const LIGNE_NIVEAU As Long = 1
Private Const LIGNE_TITRE As Long = 3
Private Const LIGNE_SOUS_TITRE As Long = 4
Private Const PREMIERE_LIGNE As Long = 5
Private Const NB_LIGNES As Long = 69
Public Const TYPES_FONDATION As String = "Semelle filante,Semelle isolée,Mixte"
Public Sub CreerTableauChargesPermanentesComplet()
    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim r As Long, c As Long
    Dim Entete(1 To 2, 1 To 13) As String
    Dim titres As Variant, sousTitres As Variant
    Set ws = ActiveSheet
    derniereLigne = PREMIERE_LIGNE + NB_LIGNES - 1
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ws.Cells.Clear
    ws.Columns.ColumnWidth = ws.StandardWidth
    CreerLigneNiveau ws
    titres = Array("Désignation de l'ouvrage", "Poids volumique (kN/m3)", "Poids Surfacique (kN/m²)", "Dimensions", "", "", "", "Charges permanentes G (kN/ml)", "Charge des poteaux G (kN)", "Charge d'exploitation Q (kN/ml)", ChrW(931) & "Q (kN/ml)", ChrW(945) & "n", ChrW(945) & "n × " & ChrW(931) & "Q (kN/ml)")
    sousTitres = Array("", "", "", "Largeur b (m)", "Hauteur h (m)", "Profondeur e (m)", "Nombre n", "", "", "", "", "", "")
    For r = 1 To 2
        For c = 1 To 13
            Entete(r, c) = IIf(r = 1, titres(c - 1), sousTitres(c - 1))
        Next c
    Next r
    ws.Range("B" & LIGNE_TITRE).Resize(2, 13).Value = Entete
    ws.Range("J" & LIGNE_TITRE).Formula = "=IF($M$1=""Semelle isolée"",""Charges permanentes G (kN)""," & """Charge des poteaux G (kN)"")"
    ws.Range("K" & LIGNE_TITRE).Formula = "=IF($M$1=""Semelle isolée"",""Charge d'exploitation Q (kN)""," & """Charge d'exploitation Q (kN/ml)"")"
    ws.Range("L" & LIGNE_TITRE).Formula = "=IF($M$1=""Semelle isolée"",""" & ChrW(931) & "Q (kN)""," & """" & ChrW(931) & "Q (kN/ml)"")"
    ws.Range("N" & LIGNE_TITRE).Formula = "=IF($M$1=""Semelle isolée"",""" & ChrW(945) & "n × " & ChrW(931) & "Q (kN)"",""" & ChrW(945) & "n × " & ChrW(931) & "Q (kN/ml)"")"
    For c = 1 To 13
        If Entete(2, c) = "" Then ws.Cells(LIGNE_TITRE, c + 1).Resize(2, 1).Merge
    Next c
    ws.Range("E" & LIGNE_TITRE & ":H" & LIGNE_TITRE).Merge
    With ws.Range("B" & LIGNE_TITRE & ":N" & LIGNE_SOUS_TITRE)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    For r = PREMIERE_LIGNE To derniereLigne
        ws.Range("I" & r).Formula = "=IF(B" & r & "="""",""""," & "(C" & r & "*E" & r & "*F" & r & "+D" & r & "*E" & r & "))"
        ws.Range("J" & r).Formula = "=IF(B" & r & "="""","""",IF(LEFT(B" & r & ",6)=""Poteau""," & "N(C" & r & ")*N(E" & r & ")*N(F" & r & ")*N(G" & r & ")*IF(H" & r & "="""",1,N(H" & r & "))," & "IF($M$1=""Semelle isolée"",N(I" & r & ")*N(G" & r & ")*IF(H" & r & "="""",1,N(H" & r & ")),"""")))"
    Next r
    ws.Range("B" & derniereLigne + 1).Value = "TOTAL G"
    ws.Range("I" & derniereLigne + 1).Formula = "=SUBTOTAL(109,I" & PREMIERE_LIGNE & ":I" & derniereLigne & ")"
    ws.Range("J" & derniereLigne + 1).Formula = "=SUBTOTAL(109,J" & PREMIERE_LIGNE & ":J" & derniereLigne & ")"
    With ws.Range("B" & derniereLigne + 1 & ":N" & derniereLigne + 1)
        .Font.Bold = True
    End With
    With ws.Range("B" & LIGNE_TITRE & ":N" & derniereLigne + 1).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
    ws.Range("B" & LIGNE_TITRE & ":N" & derniereLigne + 1).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    ws.Range("G" & PREMIERE_LIGNE & ":G" & derniereLigne).NumberFormat = "0.00"
    ws.Range("H" & PREMIERE_LIGNE & ":H" & derniereLigne).NumberFormat = "0"
    ws.Range("I" & PREMIERE_LIGNE & ":J" & derniereLigne + 1).NumberFormat = "0.00"
    ws.Range("C" & LIGNE_TITRE & ":N" & derniereLigne + 1).HorizontalAlignment = xlCenter
    ws.Columns("C:N").ColumnWidth = 13
    With ws.Range("B" & LIGNE_TITRE & ":N" & derniereLigne + 1)
        .WrapText = True
        .Rows.AutoFit
    End With
    ws.Range("I" & PREMIERE_LIGNE & ":J" & derniereLigne + 1).WrapText = False
    ws.Columns("B").ColumnWidth = 48
    ws.Range("B" & LIGNE_TITRE & ":N" & derniereLigne + 1).Rows.AutoFit
    ws.Range("C1,G1,I1,K1,M1").Interior.Color = COUL_SAISIE
    ws.Range("E1").Interior.Color = IIf(ws.Range("E1").HasFormula, COUL_CALCUL, COUL_SAISIE)
    With ws.Range("B" & LIGNE_NIVEAU & ":K" & LIGNE_NIVEAU)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Rows.AutoFit
    End With
    ws.Cells.Font.Name = "CheGuevara Text Sans"
    AjusterTextes ws
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Tableau des charges permanentes créé.", vbInformation
End Sub
Public Sub CreerLigneNiveau(ws As Worksheet)
    Dim vC1 As Variant, vG1 As Variant, vI1 As Variant, vK1 As Variant, evenements As Boolean
    Dim vM1 As Variant
    evenements = Application.EnableEvents
    Application.EnableEvents = False
    vC1 = ws.Range("C1").Value: vG1 = ws.Range("G1").Value
    vI1 = ws.Range("I1").Value: vK1 = ws.Range("K1").Value: vM1 = ws.Range("M1").Value
    ws.Range("B" & LIGNE_NIVEAU).Value = "Niveau d'étage"
    ws.Range("B" & LIGNE_NIVEAU).Font.Bold = True
    With ws.Range("C" & LIGNE_NIVEAU)
        .Value = "RDC"
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=NiveauxEtage()
        .Validation.InCellDropdown = True
        .Validation.ErrorMessage = "Choisir un niveau d'étage : RDC à R+5, avec ou sans 1 ou 2 " & "niveaux sous le RDC (1SS, 2SS : sous-sol ou vide sanitaire)."
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With
    ws.Range("D" & LIGNE_NIVEAU).Value = "Catégorie du bâtiment"
    ws.Range("D" & LIGNE_NIVEAU).Font.Bold = True
    With ws.Range("E" & LIGNE_NIVEAU)
        .Value = "A"
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="A,B,C1,C2,C3,C4,C5,D1,D2,E1,E2,F,G"
        .Validation.InCellDropdown = True
        .Validation.ErrorMessage = "Choisir une catégorie de bâtiment (A à G et sous-catégories)."
        If FeuilleExiste("Données du projet") Then .Formula = "='Données du projet'!C14"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With
    ws.Range("F" & LIGNE_NIVEAU).Value = "Entrez le nombre de murs"
    ws.Range("F" & LIGNE_NIVEAU).Font.Bold = True
    With ws.Range("G" & LIGNE_NIVEAU)
        .Value = 1
        .Validation.Delete
        .Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="50"
        .Validation.ErrorMessage = "Entrer un nombre entier de murs (1 à 50)."
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With
    ws.Range("J" & LIGNE_NIVEAU).Value = "Nombre de micropieux"
    ws.Range("J" & LIGNE_NIVEAU).Font.Bold = True
    With ws.Range("K" & LIGNE_NIVEAU)
        .Value = 1
        .Validation.Delete
        .Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="200"
        .Validation.ErrorMessage = "Entrer un nombre entier de micropieux (1 à 200)."
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With
    ws.Range("H" & LIGNE_NIVEAU).Value = "Pourcentage des ouvertures des murs (%)"
    ws.Range("H" & LIGNE_NIVEAU).Font.Bold = True
    With ws.Range("I" & LIGNE_NIVEAU)
        .Value = 0
        .Validation.Delete
        .Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="100"
        .Validation.ErrorMessage = "Entrer un pourcentage d'ouvertures entre 0 et 100."
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With
    ws.Range("L" & LIGNE_NIVEAU).Value = "Type de fondation"
    ws.Range("L" & LIGNE_NIVEAU).Font.Bold = True
    With ws.Range("M" & LIGNE_NIVEAU & ":N" & LIGNE_NIVEAU)
        .UnMerge
        .Merge
        .Value = "Mixte"
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=TYPES_FONDATION
        .Validation.ErrorMessage = "Choisir : semelle filante, semelle isolée ou mixte."
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With
    If InStr("," & TYPES_FONDATION & ",", "," & CStr(vM1) & ",") > 0 Then ws.Range("M1").Value = vM1
    If InStr("," & NiveauxEtage() & ",", "," & CStr(vC1) & ",") > 0 Then ws.Range("C1").Value = vC1
    If Val(vG1) >= 1 Then ws.Range("G1").Value = Val(vG1)
    If IsNumeric(vI1) And Not IsEmpty(vI1) Then ws.Range("I1").Value = vI1
    If Val(vK1) >= 1 Then ws.Range("K1").Value = Val(vK1)
    ws.Range("C1,G1,I1,K1,M1").Interior.Color = COUL_SAISIE
    ws.Range("E1").Interior.Color = IIf(ws.Range("E1").HasFormula, COUL_CALCUL, COUL_SAISIE)
    Application.EnableEvents = evenements
End Sub
Private Function FeuilleExiste(nom As String) As Boolean
    Dim f As Worksheet
    On Error Resume Next
    Set f = ActiveWorkbook.Worksheets(nom)
    FeuilleExiste = Not f Is Nothing
End Function
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
