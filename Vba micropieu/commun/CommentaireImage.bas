Attribute VB_Name = "CommentaireImage"
Option Explicit

' Module « Commentaires en image » : les commentaires explicatifs des cellules sont
' affichés sous forme d'image (texte noir sur fond blanc, cadre vert, police lisible)
' au lieu de la note jaune d'Excel.
'   CommentaireImage c, texte : crée le commentaire de la cellule c ; le texte est
'     dessiné dans une zone de texte d'un graphique temporaire, exporté en image PNG
'     (dossier temporaire de Windows) puis posé comme fond du commentaire ; en cas
'     d'échec, commentaire texte mis en forme (fond blanc, cadre vert) ;
'   ConvertirCommentairesEnImages : convertit tous les commentaires du classeur
'     (toutes les feuilles, toutes les cellules) ; à lancer après la création des
'     feuilles, ou à tout moment.
' Utilisé par les modules CalculFondation, VerificationStructurelle, ChargePermanente
' (Application.Run « CommentaireImage »).

Private Const LARGEUR As Double = 320          ' largeur de l'image (points)
Private Const TAILLE_POLICE As Double = 10
Private Const VERT_FONCE As Long = 2315831     ' RGB(55, 86, 35)

Public Sub CommentaireImage(c As Range, texte As String)
    Dim co As ChartObject, tb As Shape, f As String, h As Double, n As Long, i As Long
    Dim lignes As Variant, ecran As Boolean
    c.ClearComments
    c.AddComment ""
    On Error GoTo Texte
    ecran = Application.ScreenUpdating
    ' Hauteur estimée : environ 55 caractères par ligne pour 320 points en Calibri 10
    lignes = Split(texte, vbLf)
    For i = 0 To UBound(lignes)
        n = n + Application.WorksheetFunction.Max(1, Application.WorksheetFunction.RoundUp(Len(lignes(i)) / 55, 0))
    Next i
    h = n * TAILLE_POLICE * 1.35 + 14
    ' Image : graphique temporaire (fond blanc, cadre vert) contenant une zone de texte
    Set co = c.Worksheet.ChartObjects.Add(0, 0, LARGEUR, h)
    With co.Chart
        .ChartArea.Format.Fill.Visible = msoTrue
        .ChartArea.Format.Fill.Solid
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Format.Line.Visible = msoTrue
        .ChartArea.Format.Line.ForeColor.RGB = VERT_FONCE
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
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = TAILLE_POLICE
            .TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        End With
    End With
    f = Environ("TEMP") & Application.PathSeparator & "commentaire_" & Format(Now, "hhnnss") & "_" & _
        Int(Rnd() * 1000000) & ".png"
    co.Chart.Export Filename:=f, FilterName:="PNG"
    co.Delete
    Set co = Nothing
    ' Image posée comme fond du commentaire, à sa taille
    With c.Comment.Shape
        .Fill.UserPicture f
        .Line.Visible = msoFalse
        .Width = LARGEUR
        .Height = h
    End With
    c.Comment.Shape.AlternativeText = texte     ' texte conservé (accessibilité, recherche)
    Kill f
    Application.ScreenUpdating = ecran
    Exit Sub
Texte:
    ' Échec de l'image : commentaire texte lisible (fond blanc, cadre vert, Calibri 10)
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If f <> "" Then Kill f
    c.ClearComments
    c.AddComment texte
    With c.Comment.Shape
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = VERT_FONCE
        .Line.Weight = 1.5
        .TextFrame.Characters.Font.Name = "Calibri"
        .TextFrame.Characters.Font.Size = TAILLE_POLICE
        .TextFrame.Characters.Font.Bold = False
        .TextFrame.AutoSize = True
    End With
    Application.ScreenUpdating = ecran
End Sub

' Convertit en image tous les commentaires (texte) du classeur actif
Public Sub ConvertirCommentairesEnImages()
    Dim ws As Worksheet, cm As Comment, cellules As Collection, textes As Collection, i As Long
    Dim t As String
    Application.ScreenUpdating = False
    For Each ws In ActiveWorkbook.Worksheets
        Set cellules = New Collection
        Set textes = New Collection
        For Each cm In ws.Comments
            t = cm.Text
            If Trim(t) <> "" Then              ' commentaire déjà en image : texte vide
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
