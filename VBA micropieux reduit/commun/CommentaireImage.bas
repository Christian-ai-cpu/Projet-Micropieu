Attribute VB_Name = "CommentaireImage"
Option Explicit
Private Const COUL_TITRE As Long = 5898259
Private Const LARGEUR As Double = 320
Private Const TAILLE_POLICE As Double = 10
Public Sub CommentaireImage(c As Range, texte As String)
    Dim co As ChartObject, tb As Shape, f As String, h As Double, n As Long, i As Long
    Dim lignes As Variant, ecran As Boolean
    c.ClearComments
    c.AddComment ""
    On Error GoTo Texte
    ecran = Application.ScreenUpdating
    lignes = Split(texte, vbLf)
    For i = 0 To UBound(lignes)
        n = n + Application.WorksheetFunction.Max(1, Application.WorksheetFunction.RoundUp(Len(lignes(i)) / 55, 0))
    Next i
    h = n * TAILLE_POLICE * 1.35 + 14
    Set co = c.Worksheet.ChartObjects.Add(0, 0, LARGEUR, h)
    With co.Chart
        .ChartArea.Format.Fill.Visible = msoTrue
        .ChartArea.Format.Fill.Solid
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Format.Line.Visible = msoTrue
        .ChartArea.Format.Line.ForeColor.RGB = COUL_TITRE
        .ChartArea.Format.Line.Weight = 1.5
        Set tb = .Shapes.AddTextbox(msoTextOrientationHorizontal, 3, 3, LARGEUR - 6, h - 6)
    End With
    With tb
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        With .TextFrame2
            .WordWrap = msoTrue
            .MarginLeft = 5
            .MarginRight = 5
            .MarginTop = 4
            .MarginBottom = 4
            .TextRange.Text = texte
            .TextRange.Font.Name = "CheGuevara Text Sans"
            .TextRange.Font.Size = TAILLE_POLICE
            .TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        End With
    End With
    f = Environ("TEMP") & Application.PathSeparator & "commentaire_" & Format(Now, "hhnnss") & "_" & Int(Rnd() * 1000000) & ".png"
    co.Chart.Export Filename:=f, FilterName:="PNG"
    co.Delete
    Set co = Nothing
    With c.Comment.Shape
        .Fill.UserPicture f
        .Line.Visible = msoFalse
        .Width = LARGEUR
        .Height = h
    End With
    c.Comment.Shape.AlternativeText = texte
    Kill f
    Application.ScreenUpdating = ecran
    Exit Sub
Texte:
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If f <> "" Then Kill f
    c.ClearComments
    c.AddComment texte
    With c.Comment.Shape
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = COUL_TITRE
        .Line.Weight = 1.5
        .TextFrame.Characters.Font.Name = "CheGuevara Text Sans"
        .TextFrame.Characters.Font.Size = TAILLE_POLICE
        .TextFrame.Characters.Font.Bold = False
        .TextFrame.AutoSize = True
    End With
    Application.ScreenUpdating = ecran
End Sub
Public Sub ConvertirCommentairesEnImages()
    Dim ws As Worksheet, cm As Comment, cellules As Collection, textes As Collection, i As Long
    Dim t As String
    Application.ScreenUpdating = False
    For Each ws In ActiveWorkbook.Worksheets
        Set cellules = New Collection
        Set textes = New Collection
        For Each cm In ws.Comments
            t = cm.Text
            If Trim(t) <> "" Then
                cellules.Add cm.Parent
                textes.Add t
            End If
        Next cm
        For i = 1 To cellules.Count
            CommentaireImage cellules(i), CStr(textes(i))
        Next i
    Next ws
    Application.ScreenUpdating = True
End Sub
