Attribute VB_Name = "DonneesProjet"
Option Explicit

' Module « Données du projet » : crée (ou recrée) la feuille « Données du projet »
' avec les hypothèses générales du projet.
'   B1        : HYPOTHÈSES GÉNÉRALES
'   B3 à C9   : Synthèse du projet (nom, lieu, maître d'ouvrage, date, numéro,
'               phase : liste G2-AVP, G2-PRO, G5)
'   B11 à C16 : Caractérisation générale de l'ouvrage (type d'ouvrage : Bâtiment
'               ou Pont ; type de fondation, affiché seulement en phase G5 ;
'               catégorie A à G ; altitude ; zone de sismicité 1 à 5)
'   B18 à D23 : Références normatives (EC0, EC2, EC3 partie 1, EC7)
' Les valeurs se saisissent en colonne C. Le type de fondation (ligne 13) s'affiche
' quand la phase (C9) vaut G5 : voir AfficherTypeFondation et CodeFeuilleDonnees.txt.

Private Const NOM_FEUILLE As String = "Données du projet"
Private Const CELLULE_PHASE As String = "C9"
Private Const LIGNE_FONDATION As Long = 13

Public Sub CreerDonneesProjet()
    Dim ws As Worksheet, r As Long, normes As Variant

    ' Feuille créée si elle n'existe pas, sinon vidée
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(Before:=ActiveWorkbook.Worksheets(1))
        ws.Name = NOM_FEUILLE
    End If
    ws.Cells.Clear
    ws.Rows.Hidden = False

    ' --- Titre ---
    With ws.Range("B1:D1")
        .Merge
        .Value = "HYPOTHÈSES GÉNÉRALES"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With

    ' --- Synthèse du projet ---
    SousTitre ws, 3, "Synthèse du projet"
    ws.Range("B4:B9").Value = Application.Transpose(Array("Nom du projet", "Lieu", _
        "Maître d'ouvrage", "Date", "Numéro du projet", "Phase du projet"))
    ws.Range("C7").Value = Date
    ws.Range("C7").NumberFormat = "dd/mm/yyyy"
    Liste ws.Range(CELLULE_PHASE), "G2-AVP,G2-PRO,G5"

    ' --- Caractérisation générale de l'ouvrage ---
    SousTitre ws, 11, "Caractérisation générale de l'ouvrage"
    ws.Range("B12:B16").Value = Application.Transpose(Array("Type d'ouvrage", _
        "Type de fondation", "Catégorie de bâtiment", "Altitude", "Zone de sismicité"))
    Liste ws.Range("C12"), "Bâtiment,Pont"
    Liste ws.Range("C13"), "Fondations superficielles,Fondations semi-profondes," & _
        "Fondations profondes (pieux),Micropieux,Radier"
    Liste ws.Range("C14"), "A,B,C,D,E,F,G"
    Liste ws.Range("C15"), "<=1000 m,>1000 m"
    Liste ws.Range("C16"), "1,2,3,4,5"

    ' --- Références normatives ---
    SousTitre ws, 18, "Références normatives"
    ws.Range("B19:D19").Value = Array("Eurocode", "Référence", "Intitulé")
    ws.Range("B19:D19").Font.Bold = True
    normes = Array( _
        Array("EC0", "NF EN 1990 et son annexe nationale", _
            "Eurocodes structuraux - Bases de calcul des structures"), _
        Array("EC2", "NF EN 1992-1-1 et son annexe nationale", _
            "Calcul des structures en béton - Partie 1-1 : règles générales et règles pour les bâtiments"), _
        Array("EC3 - Partie 1", "NF EN 1993-1-1 et son annexe nationale", _
            "Calcul des structures en acier - Partie 1-1 : règles générales et règles pour les bâtiments"), _
        Array("EC7", "NF EN 1997-1 et son annexe nationale", _
            "Calcul géotechnique - Partie 1 : règles générales"))
    For r = 0 To UBound(normes)
        ws.Range("B" & 20 + r & ":D" & 20 + r).Value = normes(r)
    Next r

    ' --- Mise en forme ---
    With ws.Range("B4:C9,B12:C16,B19:D23")
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("C4:C9,C12:C16").Interior.Color = RGB(255, 255, 204)   ' cases à remplir
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 40
    ws.Columns("D").ColumnWidth = 70
    ws.Rows("19:23").AutoFit

    AfficherTypeFondation ws
    ws.Activate
End Sub

' Affiche la ligne « Type de fondation » juste sous le type d'ouvrage seulement
' quand la phase du projet est G5. Appelée aussi quand la phase change
' (voir CodeFeuilleDonnees.txt).
Public Sub AfficherTypeFondation(ws As Worksheet)
    ws.Rows(LIGNE_FONDATION).Hidden = (ws.Range(CELLULE_PHASE).Value <> "G5")
End Sub

' Sous-titre en gras, souligné, en colonne B
Private Sub SousTitre(ws As Worksheet, ligne As Long, texte As String)
    With ws.Range("B" & ligne)
        .Value = texte
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Size = 12
    End With
End Sub

' Liste déroulante sur une cellule, premier choix sélectionné
Private Sub Liste(c As Range, choix As String)
    c.Validation.Delete
    c.Validation.Add xlValidateList, xlValidAlertStop, , choix
    c.Value = Split(choix, ",")(0)
    c.HorizontalAlignment = xlCenter
End Sub
