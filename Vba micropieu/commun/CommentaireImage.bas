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
'   CommentaireNorme c, fichier, texte : capture de la norme (dossier « images-normes »
'     placé à côté du classeur) en commentaire de la cellule c ; texte si l'image manque.
' Utilisé par les modules CalculFondation, VerificationStructurelle, VerificationArmature,
' ChargePermanente (Application.Run « CommentaireImage » / « CommentaireNorme »).

Private Const LARGEUR As Double = 320          ' largeur de l'image (points)
Private Const TAILLE_POLICE As Double = 10
Private Const VERT_FONCE As Long = 2315831     ' RGB(55, 86, 35)
Private Const LARGEUR_MAX As Double = 520      ' largeur maximale d'une capture (points)

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

' Commentaire « capture de la norme » : l'image fichier du dossier « images-normes »
' (à placer à côté du classeur) est posée comme commentaire de la cellule c, à sa
' taille (largeur limitée à LARGEUR_MAX points) ; si l'image est introuvable, texte
' (s'il est donné) en commentaire image.
Public Sub CommentaireNorme(c As Range, fichier As String, Optional texte As String = "")
    Dim chemin As String, w As Double, h As Double, l As Double
    chemin = DossierImages() & fichier
    If Dir(chemin) = "" Then
        If texte <> "" Then CommentaireImage c, texte
        Exit Sub
    End If
    TailleImage fichier, w, h
    l = w * 0.75
    If l > LARGEUR_MAX Then l = LARGEUR_MAX
    c.ClearComments
    c.AddComment ""
    With c.Comment.Shape
        .Fill.UserPicture chemin
        .Line.ForeColor.RGB = VERT_FONCE
        .Line.Weight = 1
        .Width = l
        .Height = l * h / w
        .AlternativeText = fichier & IIf(texte = "", "", " - " & texte)
    End With
End Sub

' Dossier des captures des normes : « images-normes » à côté du classeur actif
Public Function DossierImages() As String
    DossierImages = ActiveWorkbook.Path & Application.PathSeparator & "images-normes" & _
        Application.PathSeparator
End Function

' Taille en pixels des captures (largeur w, hauteur h)
Private Sub TailleImage(fichier As String, w As Double, h As Double)
    Select Case fichier
        Case "document_cisaillement_tube.png": w = 673: h = 791
        Case "ec3_6-1_gamma_M0.png": w = 1058: h = 468
        Case "ec3_6-2-1_eq_6-1.png": w = 973: h = 300
        Case "ec3_6-2-1_eq_6-2.png": w = 975: h = 211
        Case "ec3_6-2-4_eq_6-10.png": w = 974: h = 196
        Case "ec3_6-2-4_eq_6-9.png": w = 1058: h = 186
        Case "ec3_6-2-5_moment.png": w = 974: h = 601
        Case "ec3_6-2-6_Av_tube.png": w = 804: h = 45
        Case "ec3_6-2-6_cisaillement.png": w = 1058: h = 429
        Case "ec3_tableau_5-2_tube.png": w = 1054: h = 540
        Case "en1990_6-10.png": w = 1056: h = 130
        Case "en1990_6-11b.png": w = 1056: h = 286
        Case "en1990_6-14b.png": w = 975: h = 275
        Case "en1990_6-16b.png": w = 1058: h = 306
        Case "en1990_tableau_A1-1.png": w = 1058: h = 773
        Case "enpc_diapo_42.png": w = 1100: h = 688
        Case "enpc_diapo_43.png": w = 1100: h = 688
        Case "enpc_diapo_44.png": w = 1100: h = 688
        Case "fascicule_Kf_pressio_CPT.png": w = 551: h = 743
        Case "fascicule_courbes_reaction.png": w = 556: h = 783
        Case "fascicule_equation_yg.png": w = 610: h = 77
        Case "fascicule_methode_resolution_l0.png": w = 566: h = 763
        Case Else: w = 800: h = 500
    End Select
End Sub
