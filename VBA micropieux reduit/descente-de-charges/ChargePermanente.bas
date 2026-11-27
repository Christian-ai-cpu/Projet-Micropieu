Attribute VB_Name = "ChargePermanente"
Option Explicit
Private Const COUL_SAISIE As Long = 15201485
Private Const COUL_CALCUL As Long = 15590349
Private Const TYPES_TOITURE As String = "Toiture fermettes industrielles + tuiles," & "Toiture fermettes industrielles + ardoises,Toiture charpente traditionnelle bois massif"
Public Const LIGNE_POIDS_TOITURE As String = "Poids surfacique de la toiture"
Private Const PENTES As String = "18,20,25,30,45"
Private Const PLANCHERS As String = "Plancher dalle pleine,Plancher à poutrelles,Plancher bois traditionnel"
Private Const POUTRES As String = "Poutre BA,Poutre bois,Poutre métallique"
Private Const POTEAUX As String = "Poteau BA,Poteau bois,Poteau métallique"
Private Const SEMELLES As String = "Semelle BA,Semelle béton non armé,Semelle gros béton," & "Semelle maçonnerie de moellons"
Private Const REVETEMENTS As String = "Carrelage grès cérame,Dalles épaisses (marbre),Parquet traditionnel,Sols minces"
Private Const ENDUIT_ISOLANT As Double = 0.2
Private Const SUFFIXE_MUR As String = " + enduit et isolant"
Private Const MURS As String = "Agglos béton creux + enduit et isolant," & "Agglos béton plein + enduit et isolant,Béton banché + enduit et isolant," & "Briques creuses + enduit et isolant,Briques Monomur + enduit et isolant," & "Briques pleines + enduit et isolant,Pierre de taille + enduit et isolant"
Public Function NiveauxEtage() As String
    Dim k As Long, s As String, base As String
    For k = 0 To 5
        base = IIf(k = 0, "RDC", "R+" & k)
        s = s & "," & base & "," & base & "+1SS," & base & "+2SS"
    Next k
    NiveauxEtage = Mid(s, 2)
End Function
Public Sub RemplirChargePermanente()
    Dim listes As Variant, i As Long, e As Long, r As Long, debut As Long
    Dim nomBloc As String, sousTotaux As String, k As Long
    If CStr(ActiveSheet.Range("H4").Value) <> "Nombre n" Then
        CreerTableauChargesPermanentesComplet
    Else
        CreerLigneNiveau ActiveSheet
    End If
    k = TrouverLigneMatrice(ActiveSheet)
    If k > LIGNE_MATRICE Then ActiveSheet.Range(ActiveSheet.Rows(5), ActiveSheet.Rows(4 + k - LIGNE_MATRICE)).Delete
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    r = 5
    For e = 6 To -4 Step -1
        If e = 6 Then
            listes = Array(ListeToitures(), LIGNE_POIDS_TOITURE, REVETEMENTS, PLANCHERS, POUTRES)
            nomBloc = "toiture"
        ElseIf e = 0 Then
            listes = Array("RDC", REVETEMENTS, PLANCHERS, POUTRES, POTEAUX, MURS)
            nomBloc = "RDC"
        ElseIf e = -1 Then
            listes = Array("Sous-sol 1,Vide sanitaire 1", REVETEMENTS, PLANCHERS, POUTRES, POTEAUX, MURS)
            nomBloc = "sous-sol 1"
        ElseIf e = -2 Then
            listes = Array("Sous-sol 2,Vide sanitaire 2", POUTRES, POTEAUX, MURS)
            nomBloc = "sous-sol 2"
        ElseIf e = -3 Then
            listes = Array("Vide sanitaire", POTEAUX, MURS)
            nomBloc = "vide sanitaire"
        ElseIf e = -4 Then
            listes = Array("Soubassement", POTEAUX, SEMELLES, MURS)
            nomBloc = "soubassement"
        Else
            listes = Array("Étage " & e, REVETEMENTS, PLANCHERS, POUTRES, POTEAUX, MURS)
            nomBloc = "étage " & e
        End If
        debut = r
        For i = 0 To UBound(listes)
            With ActiveSheet.Range("B" & r)
                .Validation.Delete
                If InStr(listes(i), ",") > 0 Then AjouterListe ActiveSheet.Range("B" & r), CStr(listes(i)), "Liste " & nomBloc
                .Value = Split(listes(i), ",")(0)
                .Font.Bold = (i = 0)
                .IndentLevel = IIf(listes(i) = REVETEMENTS, 1, 0)
            End With
            MajLigne ActiveSheet, r
            With ActiveSheet
                If InStr(listes(i), ",") > 0 Then .Range("B" & r).Interior.Color = COUL_SAISIE
                If i > 0 Then
                    .Range("C" & r & ":D" & r & ",I" & r).Interior.Color = COUL_CALCUL
                    .Range("E" & r & ":F" & r).Interior.Color = COUL_SAISIE
                    If listes(i) = POTEAUX Then
                        .Range("G" & r & ":H" & r).Interior.Color = COUL_SAISIE
                        .Range("J" & r).Interior.Color = COUL_CALCUL
                    ElseIf e = -4 Then
                        .Range("G" & r).Interior.Color = COUL_SAISIE
                    End If
                End If
            End With
            r = r + 1
        Next i
        ActiveSheet.Range("I" & debut & ":J" & debut).ClearContents
        With ActiveSheet
            .Range("B" & r).Validation.Delete
            .Range("B" & r).Value = "Sous-total " & nomBloc
            .Range("I" & r).Formula = "=SUBTOTAL(109,I" & debut & ":INDEX(I:I,ROW()-1))"
            .Range("J" & r).ClearContents
            .Range("B" & r & ":N" & r).Font.Bold = True
            .Range("B" & r).Font.Italic = True
            .Range("I" & r & ":N" & r).Interior.Color = COUL_CALCUL
        End With
        sousTotaux = sousTotaux & ",I" & r
        r = r + 1
    Next e
    With ActiveSheet
        .Range("B" & r).Value = "TOTAL (somme des sous-totaux)"
        .Range("I" & r).Formula = "=SUM(" & Mid(sousTotaux, 2) & ")"
        .Range("J" & r).Formula = "=SUBTOTAL(109,J5:INDEX(J:J,ROW()-1))"
        .Range("B" & r & ":N" & r).Font.Bold = True
        .Range("I" & r & ":N" & r).Interior.Color = COUL_CALCUL
    End With
    AfficherEtages ActiveSheet
    With ActiveSheet.Columns("I:J")
        .AutoFit
        If .Columns(1).ColumnWidth < 12 Then .Columns(1).ColumnWidth = 12
        If .Columns(2).ColumnWidth < 12 Then .Columns(2).ColumnWidth = 12
    End With
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Public Function EpaisseursMur(typeMur As String) As String
    Select Case Replace(typeMur, SUFFIXE_MUR, "")
        Case "Agglos béton creux", "Agglos béton plein": EpaisseursMur = "15 cm,20 cm"
        Case "Béton banché":    EpaisseursMur = "20 cm,25 cm,30 cm"
        Case "Briques creuses": EpaisseursMur = "15 cm,20 cm,25 cm"
        Case "Briques Monomur": EpaisseursMur = "30 cm,37 cm"
        Case "Briques pleines": EpaisseursMur = "10 cm,21.5 cm,33 cm"
        Case "Pierre de taille": EpaisseursMur = "30 cm,40 cm,50 cm,60 cm"
    End Select
End Function
Public Function ListeToitures() As String
    Dim t As Variant, p As Variant, s As String
    For Each t In Split(TYPES_TOITURE, ",")
        For Each p In Split(PENTES, ",")
            s = s & "," & t & " " & p & "°"
        Next p
    Next t
    ListeToitures = Mid(s, 2) & ",Toiture terrasse"
End Function
Public Function PenteToiture(nomToiture As String) As Double
    PenteToiture = Val(Mid(nomToiture, InStrRev(nomToiture, " ") + 1))
End Function
Public Function ChargeToiture(nomToiture As String) As Double
    Dim typeToiture As String
    typeToiture = nomToiture
    If PenteToiture(nomToiture) > 0 Then typeToiture = Left(nomToiture, InStrRev(nomToiture, " ") - 1)
    Select Case typeToiture
        Case "Toiture fermettes industrielles + tuiles":      ChargeToiture = 0.7
        Case "Toiture fermettes industrielles + ardoises":    ChargeToiture = 0.55
        Case "Toiture charpente traditionnelle bois massif":  ChargeToiture = 1
    End Select
End Function
Public Sub AjouterListe(c As Range, liste As String, nom As String)
    Dim wl As Worksheet, wb As Workbook, choix As Variant, col As Long
    c.Validation.Delete
    If Len(liste) <= 255 Then
        c.Validation.Add xlValidateList, xlValidAlertStop, , liste
        c.Validation.ShowError = False
        Exit Sub
    End If
    Set wb = c.Worksheet.Parent
    On Error Resume Next
    Set wl = wb.Worksheets("Listes")
    On Error GoTo 0
    If wl Is Nothing Then
        Set wl = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wl.Name = "Listes"
        wl.Visible = xlSheetHidden
        c.Worksheet.Activate
    End If
    col = 1
    Do While wl.Cells(1, col).Value <> "" And wl.Cells(1, col).Value <> nom
        col = col + 1
    Loop
    choix = Split(liste, ",")
    wl.Columns(col).ClearContents
    wl.Cells(1, col).Value = nom
    wl.Cells(2, col).Resize(UBound(choix) + 1, 1).Value = Application.Transpose(choix)
    c.Validation.Add xlValidateList, xlValidAlertStop, , "=Listes!" & wl.Cells(2, col).Resize(UBound(choix) + 1, 1).Address
    c.Validation.ShowError = False
End Sub
Public Function ChargeRevetement(typeRevetement As String) As Double
    Select Case typeRevetement
        Case "Carrelage grès cérame":    ChargeRevetement = 0.5
        Case "Dalles épaisses (marbre)": ChargeRevetement = 1
        Case "Parquet traditionnel":     ChargeRevetement = 0.25
        Case "Sols minces":              ChargeRevetement = 0.1
    End Select
End Function
Public Function PoidsVolumique(typeElement As String) As Double
    Select Case typeElement
        Case "Poutre BA", "Poteau BA":                 PoidsVolumique = 25
        Case "Plancher dalle pleine", "Plancher à poutrelles": PoidsVolumique = 25
        Case "Semelle BA":                             PoidsVolumique = 25
        Case "Semelle béton non armé":                 PoidsVolumique = 24
        Case "Semelle gros béton":                     PoidsVolumique = 22
        Case "Semelle maçonnerie de moellons":         PoidsVolumique = 23
        Case "Poutre bois", "Poteau bois":             PoidsVolumique = 6
        Case "Poutre métallique", "Poteau métallique": PoidsVolumique = 78.5
    End Select
End Function
Public Function PoidsVolumiqueMur(typeMur As String) As Double
    Select Case Replace(typeMur, SUFFIXE_MUR, "")
        Case "Agglos béton creux": PoidsVolumiqueMur = 9
        Case "Agglos béton plein": PoidsVolumiqueMur = 20
        Case "Béton banché":       PoidsVolumiqueMur = 25
        Case "Briques creuses":    PoidsVolumiqueMur = 9
        Case "Briques Monomur":    PoidsVolumiqueMur = 8
        Case "Briques pleines":    PoidsVolumiqueMur = 18
        Case "Pierre de taille":   PoidsVolumiqueMur = 22
    End Select
End Function
Public Sub MajLigne(ws As Worksheet, ligne As Long, Optional dc As Long = 0)
    Dim typeElement As String, texteNote As String
    typeElement = CStr(ws.Cells(ligne, 2 + dc).Value)
    If Left(typeElement, 7) = "Toiture" Then
        ws.Range(ws.Cells(ligne, 3 + dc), ws.Cells(ligne, 4 + dc)).ClearContents
        ws.Range(ws.Cells(ligne, 9 + dc), ws.Cells(ligne, 10 + dc)).ClearContents
        ws.Cells(ligne + 1, 3 + dc).ClearContents
        If ChargeToiture(typeElement) > 0 Then
            ws.Cells(ligne + 1, 4 + dc).Formula = "=" & Replace(CStr(ChargeToiture(typeElement)), ",", ".") & "/COS(RADIANS(" & PenteToiture(typeElement) & "))"
        Else
            ws.Cells(ligne + 1, 4 + dc).ClearContents
        End If
        ws.Cells(ligne + 1, 9 + dc).FormulaR1C1 = "=IF(RC[-7]="""","""",RC[-4]*(RC[-6]*RC[-3]+RC[-5]))"
        AfficherEtages ws
    ElseIf Left(typeElement, 3) = "RDC" Then
        AfficherEtages ws
    ElseIf EpaisseursMur(typeElement) <> "" Then
        MajEpaisseurMur ws, ligne, dc
    ElseIf Left(typeElement, 8) = "Plancher" Then
        MajPlancher ws, ligne, typeElement, dc
    ElseIf ChargeRevetement(typeElement) > 0 Then
        ws.Cells(ligne, 3 + dc).ClearContents
        ws.Cells(ligne, 4 + dc).Value = ChargeRevetement(typeElement)
    ElseIf PoidsVolumique(typeElement) > 0 Then
        ws.Cells(ligne, 3 + dc).Value = PoidsVolumique(typeElement)
        ws.Cells(ligne, 4 + dc).ClearContents
        If Left(typeElement, 6) = "Poteau" Then ws.Cells(ligne, 9 + dc).FormulaR1C1 = "=IF(R1C13=""Mixte"","""",IF(RC[-7]="""","""",RC[-6]*RC[-4]*RC[-3]+RC[-5]*RC[-4]))"
        If Left(typeElement, 6) = "Poutre" Then
            texteNote = "Hauteur h de la poutre : si une dalle est présente, prendre " & "h = hauteur totale de la poutre - épaisseur de la dalle " & "(retombée seule), pour ne pas compter deux fois le béton de la dalle."
            On Error Resume Next
            Application.Run "CommentaireImage.CommentaireImage", ws.Cells(ligne, 6 + dc), texteNote
            If Err.Number <> 0 Then
                Err.Clear
                ws.Cells(ligne, 6 + dc).ClearComments
                ws.Cells(ligne, 6 + dc).AddComment texteNote
            End If
            On Error GoTo 0
        End If
    End If
End Sub
Public Sub MajPlancher(ws As Worksheet, ligne As Long, typePlancher As String, Optional dc As Long = 0)
    Select Case typePlancher
        Case "Plancher dalle pleine"
            ws.Cells(ligne, 3 + dc).Value = PoidsVolumique(typePlancher)
            ws.Cells(ligne, 4 + dc).ClearContents
        Case "Plancher à poutrelles"
            ws.Cells(ligne, 3 + dc).ClearContents
            ws.Cells(ligne, 4 + dc).FormulaR1C1 = "=IF(RC[2]="""",0,0.05*" & PoidsVolumique(typePlancher) & "+(RC[2]-0.05)*11)"
        Case "Plancher bois traditionnel"
            ws.Cells(ligne, 3 + dc).ClearContents
            ws.Cells(ligne, 4 + dc).Value = 1.5
    End Select
End Sub
Public Sub MajEpaisseurMur(ws As Worksheet, ligne As Long, Optional dc As Long = 0)
    Dim liste As String, typeMur As String
    typeMur = CStr(ws.Cells(ligne, 2 + dc).Value)
    liste = EpaisseursMur(typeMur)
    If liste = "" Then Exit Sub
    With ws.Cells(ligne, 5 + dc)
        .Validation.Delete
        .Validation.Add xlValidateList, xlValidAlertStop, , liste
        .Validation.ShowError = False
        .Value = Val(liste) / 100
        .NumberFormat = "0.000"
    End With
    ws.Cells(ligne, 6 + dc).Validation.Delete
    ws.Cells(ligne, 3 + dc).ClearContents
    ws.Cells(ligne, 4 + dc).FormulaR1C1 = "=" & Replace(CStr(PoidsVolumiqueMur(typeMur)), ",", ".") & "*RC[1]+" & Replace(CStr(ENDUIT_ISOLANT), ",", ".")
    ws.Cells(ligne, 9 + dc).FormulaR1C1 = "=IF(RC[-7]="""","""",(100-R1C)/100*RC[-5]*RC[-3])"
End Sub
Public Sub AfficherEtages(ws As Worksheet)
    Dim niveau As Long, rdc As String, terrasse As Boolean, nSS As Long, bloc As String
    Dim typeF As String, r As Long, k As Long, nb As Long, t As String, rFin As Long
    Dim cache As Boolean, numE As Long, revPl As Boolean
    rFin = LigneMatrice(ws) - 1
    ws.Range(ws.Rows(5), ws.Rows(rFin)).Hidden = False
    terrasse = (CStr(ws.Range("B5").Value) = "Toiture terrasse")
    rdc = CStr(ws.Range("C1").Value)
    nSS = 0
    If InStr(rdc, "+1SS") > 0 Then nSS = 1
    If InStr(rdc, "+2SS") > 0 Then nSS = 2
    niveau = 0
    If Left(rdc, 2) = "R+" Then niveau = Val(Mid(rdc, 3))
    typeF = CStr(ws.Range("M1").Value)
    nb = Val(ws.Range("G1").Value)
    If nb < 1 Then nb = 1
    For r = 5 To rFin
        t = CStr(ws.Range("B" & r).Value)
        If Left(t, 5) = "TOTAL" Then Exit For
        If Left(t, 7) = "Toiture" Then
            bloc = "T"
        ElseIf Left(t, 6) = "Étage " Then
            bloc = "E": numE = Val(Mid(t, 7))
        ElseIf t = "RDC" Then
            bloc = "R"
        ElseIf t = "Sous-sol 1" Or t = "Vide sanitaire 1" Then
            bloc = "S1"
        ElseIf t = "Sous-sol 2" Or t = "Vide sanitaire 2" Then
            bloc = "S2"
        ElseIf t = "Vide sanitaire" Then
            bloc = "V"
        ElseIf t = "Soubassement" Then
            bloc = "B"
        End If
        Select Case bloc
            Case "E": cache = (numE > niveau)
            Case "S1": cache = (nSS < 1)
            Case "S2": cache = (nSS < 2)
            Case "V": cache = True
            Case Else: cache = False
        End Select
        revPl = (InStr("," & REVETEMENTS & ",", "," & t & ",") > 0 Or Left(t, 8) = "Plancher")
        If revPl Then
            If bloc = "T" And Not terrasse Then cache = True
            If bloc = "R" And nSS = 0 Then cache = True
            If bloc = "S1" And nSS < 2 Then cache = True
        End If
        If t = LIGNE_POIDS_TOITURE And terrasse Then cache = True
        If typeF = "Semelle filante" And (Left(t, 6) = "Poutre" Or Left(t, 6) = "Poteau" Or Left(t, 7) = "Semelle") Then cache = True
        If cache Then ws.Rows(r).Hidden = True
        If ws.Range("C" & r).Interior.Color = COUL_CALCUL Then
            For k = 0 To nb - 1
                With ws.Range(ws.Cells(r, 7 + k * PAS_MUR), ws.Cells(r, 8 + k * PAS_MUR))
                    If typeF = "Semelle isolée" Or (typeF = "Mixte" And Left(t, 6) = "Poteau") Then
                        .Interior.Color = COUL_SAISIE
                    ElseIf bloc = "B" Then
                        .Cells(1, 1).Interior.Color = COUL_SAISIE
                        .Cells(1, 2).Interior.Pattern = xlNone
                    Else
                        .Interior.Pattern = xlNone
                    End If
                End With
            Next k
        End If
        If Left(t, 10) = "Sous-total" Then bloc = ""
    Next r
    For k = 0 To nb - 1
        ws.Columns(9 + k * PAS_MUR).Hidden = (typeF = "Semelle isolée")
        ws.Columns(10 + k * PAS_MUR).Hidden = (typeF = "Semelle filante")
    Next k
End Sub
Public Sub InsererElement()
    Dim ws As Worksheet, r As Long, t As String, tPrec As String
    Set ws = ActiveSheet
    r = ActiveCell.Row
    If r <= 5 Or r >= LigneMatrice(ws) - 1 Then
        MsgBox "Sélectionnez une ligne du tableau de descente de charges (dans un étage, sous " & "son titre), puis relancez InsererElement.", vbExclamation
        Exit Sub
    End If
    t = CStr(ws.Range("B" & r).Value)
    tPrec = CStr(ws.Range("B" & r - 1).Value)
    If Left(tPrec, 10) = "Sous-total" Or Left(t, 5) = "TOTAL" Then
        MsgBox "La ligne doit être insérée dans un étage, entre son titre et son sous-total.", vbExclamation
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ws.Rows(r).Insert Shift:=xlDown
    ws.Rows(r - 1).Copy Destination:=ws.Rows(r)
    If ws.Range("B" & r - 1).Font.Bold Then
        ws.Range("B" & r).Font.Bold = False
        ws.Range("B" & r).ClearContents
    End If
    Application.CutCopyMode = False
    Application.EnableEvents = True
    AfficherEtages ws
    Application.ScreenUpdating = True
    ws.Range("B" & r).Select
End Sub
