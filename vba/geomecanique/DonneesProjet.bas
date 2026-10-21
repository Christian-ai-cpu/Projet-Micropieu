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
'   B18 à I26 : Références normatives (référence, date, intitulé)
'   B28 à I45 : Données géotechniques : nombre de couches (C29, 1 à 15) et tableau
'               des couches de sol (seules les lignes jusqu'à ce nombre sont affichées)
'   B47 à D52 : Données des micropieux (arase supérieure, base, nappe, méthode,
'               niveau d'eau ; valeur en C, unité en D)
' Les valeurs se saisissent dans les cases jaunes. Le type de fondation (ligne 13)
' s'affiche quand la phase (C9) vaut G5 (AfficherTypeFondation) et le tableau des
' couches suit le nombre de couches (AfficherCouches) : voir CodeFeuilleDonnees.txt.
' Les listes de plus de 255 caractères (limite d'Excel, ex. noms des sols de la
' NF P11-300) sont écrites sur la feuille masquée « Listes » (voir ListeLongue).

Private Const NOM_FEUILLE As String = "Données du projet"
Private Const CELLULE_PHASE As String = "C9"
Private Const LIGNE_FONDATION As Long = 13
Private Const CELLULE_NB_COUCHES As String = "C29"   ' nombre de couches saisi
Private Const LIGNE_COUCHES As Long = 31       ' première couche de sol
Private Const MAX_COUCHES As Long = 15         ' lignes prévues (au-delà du nombre : masquées)
Private Const CATEGORIES_SOL As String = "Argiles et limons," & _
    "Sols intermédiaires (sable limoneux / sable argileux / argile sableuse)," & _
    "Sables et graves,Craies,Marnes et calcaires marneux,Roches"

Public Sub CreerDonneesProjet()
    Dim ws As Worksheet, r As Long, normes As Variant, lc As Long, lm As Long

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
    With ws.Range("B1:I1")
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

    ' --- Références normatives : référence (date), intitulé fusionné de D à I ---
    SousTitre ws, 18, "Références normatives"
    ws.Range("B19:D19").Value = Array("Référence", "Date", "Intitulé")
    normes = Array( _
        Array("NF EN 1990", "mars 2003", _
            "Eurocode 0 - Bases de calcul des structures"), _
        Array("NF EN 1992-1-1", "octobre 2005", _
            "Eurocode 2 - Calcul des structures en béton - Partie 1-1 : règles générales et règles pour les bâtiments"), _
        Array("NF P22-311-1", "octobre 2005", _
            "Eurocode 3 (NF EN 1993-1-1) - Calcul des structures en acier - Partie 1-1 : règles générales et règles pour les bâtiments"), _
        Array("NF P94-251-1", "juin 2005", _
            "Eurocode 7 (NF EN 1997-1) - Calcul géotechnique - Partie 1 : règles générales"), _
        Array("NF P94-262", "juillet 2012", _
            "Justification des ouvrages géotechniques - Norme d'application nationale de l'Eurocode 7 - Fondations profondes"), _
        Array("NF EN 1998-1", "septembre 2005", _
            "Eurocode 8 - Calcul des structures pour leur résistance aux séismes - Partie 1 : règles générales, actions sismiques et règles pour les bâtiments"), _
        Array("NF EN 1998-5", "septembre 2005", _
            "Eurocode 8 - Partie 5 : fondations, ouvrages de soutènement et aspects géotechniques"))
    For r = 0 To UBound(normes)
        ws.Range("B" & 20 + r & ":D" & 20 + r).Value = normes(r)
    Next r
    For r = 19 To 26
        ws.Range("D" & r & ":I" & r).Merge
    Next r

    ' --- Données géotechniques : nombre de couches puis tableau des couches ---
    SousTitre ws, 28, "Données géotechniques"
    ws.Range("B29").Value = "Nombre de couches"
    With ws.Range(CELLULE_NB_COUCHES)
        .Value = 3
        .Validation.Delete
        .Validation.Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, "1", CStr(MAX_COUCHES)
        .Validation.ErrorMessage = "Entrer un nombre de couches entre 1 et " & MAX_COUCHES & "."
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("B" & LIGNE_COUCHES - 1 & ":I" & LIGNE_COUCHES - 1).Value = Array("N° de couche", _
        "Nom de couche (NF P11-300)", "Catégorie de sol (NF P94-262)", "Prof. bas (m/TN)", _
        "Épaisseur (m)", "Pression limite pl* (MPa)", "Module pressiométrique EM (MPa)", _
        "Terrain neutralisé")
    For r = LIGNE_COUCHES To LIGNE_COUCHES + MAX_COUCHES - 1
        ws.Range("B" & r).Value = r - LIGNE_COUCHES + 1
        ListeLongue ws.Range("C" & r), SolsNFP11300(), "Sols NF P11-300"
        ws.Range("C" & r).ClearContents      ' nom de couche à choisir
        Liste ws.Range("D" & r), CATEGORIES_SOL
        ws.Range("D" & r).ClearContents      ' catégorie à choisir
        ' Épaisseur = différence entre la base de la couche et celle de la couche
        ' au-dessus (TN pour la première couche)
        If r = LIGNE_COUCHES Then
            ws.Range("F" & r).Formula = "=IF(E" & r & "="""","""",ABS(E" & r & "))"
        Else
            ws.Range("F" & r).Formula = "=IF(E" & r & "="""","""",ABS(E" & r & "-E" & r - 1 & "))"
        End If
        Liste ws.Range("I" & r), "Non,Oui"
    Next r
    lc = LIGNE_COUCHES + MAX_COUCHES - 1
    ws.Range("E" & LIGNE_COUCHES & ":F" & lc).NumberFormat = "0.00"
    ws.Range("G" & LIGNE_COUCHES & ":H" & lc).NumberFormat = "0.00"
    ws.Range("B" & LIGNE_COUCHES & ":B" & lc & ",E" & LIGNE_COUCHES & ":I" & lc).HorizontalAlignment = xlCenter

    ' --- Données des micropieux : valeur en C, unité à droite en D ---
    lm = lc + 2                              ' sous-titre
    SousTitre ws, lm, "Données des micropieux"
    ws.Range("B" & lm + 1 & ":B" & lm + 5).Value = Application.Transpose(Array( _
        "Cote de l'arase supérieure des micropieux", "Base des micropieux", _
        "Présence de nappe phréatique", "Méthode de calcul retenue pour les pieux", _
        "Niveau de l'eau"))
    ws.Range("D" & lm + 1 & ":D" & lm + 5).Value = Application.Transpose(Array( _
        "m/TN", "m/TN", "", "", "m/TN"))
    ws.Range("C" & lm + 1 & ":C" & lm + 2 & ",C" & lm + 5).NumberFormat = "0.00"
    ws.Range("C" & lm + 1 & ":C" & lm + 5).HorizontalAlignment = xlCenter
    Liste ws.Range("C" & lm + 3), "Non,Oui"
    Liste ws.Range("C" & lm + 4), "Méthode pressiométrique,Méthode pénétrométrique," & _
        "Essais de chargement statique"

    ' --- Mise en forme ---
    With ws.Range("B4:C9,B12:C16,B19:I26,B29:C29,B" & LIGNE_COUCHES - 1 & ":I" & lc & _
            ",B" & lm + 1 & ":D" & lm + 5)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("B19:I19,B" & LIGNE_COUCHES - 1 & ":I" & LIGNE_COUCHES - 1).Font.Bold = True
    ws.Range("B" & LIGNE_COUCHES - 1 & ":I" & LIGNE_COUCHES - 1).HorizontalAlignment = xlCenter
    ' Cases à remplir (données d'entrée) en jaune ; l'épaisseur est calculée
    ws.Range("C4:C9,C12:C16,C29,C" & LIGNE_COUCHES & ":E" & lc & ",G" & LIGNE_COUCHES & _
        ":I" & lc & ",C" & lm + 1 & ":C" & lm + 5).Interior.Color = RGB(255, 255, 204)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C:D").ColumnWidth = 40
    ws.Columns("E:I").ColumnWidth = 14
    ws.Rows("19:" & lm + 5).AutoFit

    AfficherTypeFondation ws
    AfficherCouches ws
    ws.Activate
End Sub

' Affiche autant de lignes de couches que le nombre saisi en C29 (1 à MAX_COUCHES),
' les autres lignes du tableau sont masquées. Appelée aussi quand C29 change
' (voir CodeFeuilleDonnees.txt).
Public Sub AfficherCouches(ws As Worksheet)
    Dim nb As Long
    nb = Val(ws.Range(CELLULE_NB_COUCHES).Value)
    If nb < 1 Then nb = 1
    If nb > MAX_COUCHES Then nb = MAX_COUCHES
    ws.Rows(LIGNE_COUCHES).Resize(MAX_COUCHES).Hidden = False
    If nb < MAX_COUCHES Then ws.Rows(LIGNE_COUCHES + nb).Resize(MAX_COUCHES - nb).Hidden = True
End Sub

' Affiche la ligne « Type de fondation » juste sous le type d'ouvrage seulement
' quand la phase du projet est G5. Appelée aussi quand la phase change
' (voir CodeFeuilleDonnees.txt).
Public Sub AfficherTypeFondation(ws As Worksheet)
    ws.Rows(LIGNE_FONDATION).Hidden = (ws.Range(CELLULE_PHASE).Value <> "G5")
End Sub

' Noms des sols et roches des classes A, B, C, D et R de la NF P11-300 (GTR).
' Désignations abrégées, à compléter si besoin.
Private Function SolsNFP11300() As String
    Dim s As String
    s = "A1 - Limons peu plastiques / loess / sables fins peu pollués"
    s = s & ",A2 - Sables fins argileux / limons / argiles et marnes peu plastiques"
    s = s & ",A3 - Argiles et argiles marneuses / limons très plastiques"
    s = s & ",A4 - Argiles et argiles marneuses très plastiques"
    s = s & ",B1 - Sables silteux"
    s = s & ",B2 - Sables argileux (peu argileux)"
    s = s & ",B3 - Graves silteuses"
    s = s & ",B4 - Graves argileuses (peu argileuses)"
    s = s & ",B5 - Sables et graves très silteux"
    s = s & ",B6 - Sables et graves argileux à très argileux"
    s = s & ",C1 - Sols à fines et gros éléments (argiles à silex / éboulis / moraines)"
    s = s & ",C2 - Sols à gros éléments dominants avec fines"
    s = s & ",D1 - Sables alluvionnaires propres / sables de dune"
    s = s & ",D2 - Graves alluvionnaires propres / sables"
    s = s & ",D3 - Graves alluvionnaires grossières propres / galets"
    s = s & ",R1 - Craies"
    s = s & ",R2 - Calcaires"
    s = s & ",R3 - Roches argileuses (marnes / argilites / pélites)"
    s = s & ",R4 - Roches siliceuses (grès / poudingues / brèches)"
    s = s & ",R5 - Roches salines (sel gemme / gypse)"
    s = s & ",R6 - Roches magmatiques et métamorphiques (granites / basaltes / schistes)"
    SolsNFP11300 = s
End Function

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

' Liste déroulante de plus de 255 caractères : les choix sont écrits dans une
' colonne de la feuille masquée « Listes » (titre nom en ligne 1) et la liste
' déroulante renvoie à cette colonne.
Private Sub ListeLongue(c As Range, choix As String, nom As String)
    Dim wl As Worksheet, wb As Workbook, t As Variant, col As Long

    If Len(choix) <= 255 Then
        Liste c, choix
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
        c.Worksheet.Activate                 ' la feuille ajoutée était devenue active
    End If
    ' Colonne de cette liste (même titre), sinon la première colonne libre
    col = 1
    Do While wl.Cells(1, col).Value <> "" And wl.Cells(1, col).Value <> nom
        col = col + 1
    Loop
    t = Split(choix, ",")
    wl.Columns(col).ClearContents
    wl.Cells(1, col).Value = nom
    wl.Cells(2, col).Resize(UBound(t) + 1, 1).Value = Application.Transpose(t)
    c.Validation.Delete
    c.Validation.Add xlValidateList, xlValidAlertStop, , "=Listes!" & _
        wl.Cells(2, col).Resize(UBound(t) + 1, 1).Address
    c.Value = t(0)
End Sub
