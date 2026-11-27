Attribute VB_Name = "DuplicationMurs"
Option Explicit
Public Const PAS_MUR As Long = 14
Public Sub DupliquerMurs(Optional ws As Worksheet)
    Dim nbMurs As Long, k As Long, derniere As Long, c As Long
    Dim source As Range, colB As String, ouvertures() As Variant, nom As String
    If ws Is Nothing Then Set ws = ActiveSheet
    nbMurs = Val(ws.Range("G1").Value)
    If nbMurs < 1 Then nbMurs = 1
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Fin
    ReDim ouvertures(1 To nbMurs)
    For k = 1 To nbMurs - 1
        ouvertures(k) = ws.Cells(1, 9 + k * PAS_MUR).Value
    Next k
    ws.Range(ws.Cells(1, 2 + PAS_MUR - 1), ws.Cells(LigneMatrice(ws) - 1, ws.Columns.Count)).Clear
    derniere = ws.Cells(LigneMatrice(ws) - 1, "B").End(xlUp).Row
    nom = IIf(CStr(ws.Range("M1").Value) = "Semelle isolée", "Pot ", "Mur ")
    ws.Range("B2").Value = nom & "1"
    ws.Range("B2").Font.Bold = True
    Set source = ws.Range("B2:N" & derniere)
    For k = 1 To nbMurs - 1
        source.Copy Destination:=ws.Cells(2, 2 + k * PAS_MUR)
        ws.Cells(2, 2 + k * PAS_MUR).Value = nom & k + 1
        ws.Range("H1:I1").Copy Destination:=ws.Cells(1, 8 + k * PAS_MUR)
        If Not IsEmpty(ouvertures(k)) Then ws.Cells(1, 9 + k * PAS_MUR).Value = ouvertures(k)
        colB = Split(ws.Cells(1, 2 + k * PAS_MUR).Address(True, True), "$")(1)
        source.Offset(0, k * PAS_MUR).Replace What:="$B$5", Replacement:="$" & colB & "$5", LookAt:=xlPart, MatchCase:=False
        For c = 2 To 14
            ws.Columns(c + k * PAS_MUR).ColumnWidth = ws.Columns(c).ColumnWidth
        Next c
    Next k
    ConstruireMatrice ws
    AfficherEtages ws
Fin:
    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Public Sub AppliquerTypeFondation()
    DupliquerMurs ActiveSheet
End Sub
