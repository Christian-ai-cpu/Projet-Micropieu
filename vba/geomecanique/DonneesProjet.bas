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
'   B28 à E29 : Référentiel altimétrique (liste) et niveau du terrain naturel (unité
'               m + référentiel)
'   B31 à I48 : Données géotechniques : nombre de couches (C32, 1 à 15) et tableau
'               des couches de sol (seules les lignes jusqu'à ce nombre sont affichées)
'   B50 à E55 : Données des micropieux (arase supérieure, base, nappe, méthode,
'               niveau d'eau ; intitulé sur B:C, valeur en D, unité en E)
'   B57 à D73 : Caractéristiques des micropieux (type de pieu du tableau A.1 de la
'               NF P94-262, diamètre constant, incliné, groupe, norme de référence,
'               classe, catégorie, mode de fonctionnement, résistance de pointe, B, hauteurs, nombres,
'               conditions en tête, essai de contrôle ; intitulé sur B:C, valeur en D)
'   B75       : HYPOTHÈSES RELATIVES AUX MATÉRIAUX CONSTITUTIFS DES FONDATIONS PROFONDES
'   B77 à E79 : Caractéristiques du béton (classe d'exposition, classe de résistance)
'   B81 à E91 : Caractéristiques de l'acier (type d'armature tube ou barre, classe,
'               diamètres et épaisseur du tube, E, fe, platine, nuance, contrainte max)
' Les valeurs se saisissent dans les cases jaunes. Le type de fondation (ligne 13)
' s'affiche quand la phase (C9) vaut G5 (AfficherTypeFondation) et le tableau des
' couches suit le nombre de couches (AfficherCouches) : voir CodeFeuilleDonnees.txt.
' Les listes de plus de 255 caractères (limite d'Excel, ex. noms des sols de la
' NF P11-300) sont écrites sur la feuille masquée « Listes » (voir ListeLongue).

Private Const NOM_FEUILLE As String = "Données du projet"
Private Const CELLULE_PHASE As String = "C9"
Private Const LIGNE_FONDATION As Long = 13
Private Const CELLULE_NB_COUCHES As String = "C32"   ' nombre de couches saisi
Private Const LIGNE_COUCHES As Long = 34       ' première couche de sol
Private Const MAX_COUCHES As Long = 15         ' lignes prévues (au-delà du nombre : masquées)
Private Const LIGNE_PIEU As Long = 58          ' type de pieu (caractéristiques des micropieux)
Private Const LIGNE_ACIER As Long = 82         ' type d'armature (caractéristiques de l'acier)
Private Const CLASSES_EXPOSITION As String = "X0,XC1,XC2,XC3,XC4,XD1,XD2,XD3," & _
    "XS1,XS2,XS3,XF1,XF2,XF3,XF4,XA1,XA2,XA3"
' Armatures des micropieux (dimensions courantes, à vérifier auprès du fournisseur).
' Tube : « Øde x e mm » est lu pour remplir les diamètres et l'épaisseur (MajAcier).
Private Const TUBES As String = "Tube N80 Ø73 x 5.5 mm,Tube N80 Ø88.9 x 6.5 mm," & _
    "Tube N80 Ø88.9 x 9.5 mm,Tube N80 Ø101.6 x 9 mm,Tube N80 Ø114.3 x 9 mm," & _
    "Tube N80 Ø127 x 9 mm,Tube N80 Ø139.7 x 9 mm,Tube N80 Ø168.3 x 10 mm," & _
    "Tube N80 Ø177.8 x 10.4 mm"
Private Const BARRES As String = "Barre GEWI Ø25 (S555/700),Barre GEWI Ø28 (S555/700)," & _
    "Barre GEWI Ø32 (S555/700),Barre GEWI Ø40 (S555/700),Barre GEWI Ø50 (S555/700)," & _
    "Barre HA B500 Ø25,Barre HA B500 Ø32,Barre HA B500 Ø40"
Private Const PLATINES As String = "Platine carrée soudée,Platine carrée avec raidisseurs," & _
    "Platine vissée sur manchon,Platine circulaire soudée"
Private Const CATEGORIES_SOL As String = "Argiles et limons," & _
    "Sols intermédiaires (sable limoneux / sable argileux / argile sableuse)," & _
    "Sables et graves,Craies,Marnes et calcaires marneux,Roches"

Public Sub CreerDonneesProjet()
    Dim ws As Worksheet, r As Long, normes As Variant, lc As Long, lm As Long, lp As Long
    Dim la As Long

    Application.ScreenUpdating = False       ' pas de scintillement pendant l'exécution
    Application.EnableEvents = False         ' pas de Worksheet_Change pendant la création

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

    ' --- Référentiel altimétrique et niveau du terrain naturel ---
    ' Intitulé sur B:C fusionnées, valeur en D, unité en E (m + référentiel choisi)
    ws.Range("B28").Value = "Référentiel altimétrique"
    ws.Range("B29").Value = "Niveau du terrain naturel"
    Liste ws.Range("D28"), "NGF-IGN69,NGF-IGN78,NVP,IGN,Local"
    ws.Range("D29").NumberFormat = "0.00"
    ws.Range("E29").Formula = "=""m ""&D28"
    ws.Range("D28:E29").HorizontalAlignment = xlCenter
    ws.Range("B28:C28").Merge
    ws.Range("B29:C29").Merge

    ' --- Données géotechniques : nombre de couches puis tableau des couches ---
    SousTitre ws, 31, "Données géotechniques"
    ws.Range("B32").Value = "Nombre de couches"
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

    ' --- Données des micropieux : intitulé sur B:C fusionnées, valeur en D, unité en E ---
    lm = lc + 2                              ' sous-titre
    SousTitre ws, lm, "Données des micropieux"
    ws.Range("B" & lm + 1 & ":B" & lm + 5).Value = Application.Transpose(Array( _
        "Cote de l'arase supérieure des micropieux", "Base des micropieux", _
        "Présence de nappe phréatique", "Méthode de calcul retenue pour les pieux", _
        "Niveau de l'eau"))
    ws.Range("E" & lm + 1 & ":E" & lm + 5).Value = Application.Transpose(Array( _
        "m/TN", "m/TN", "", "", "m"))
    ws.Range("D" & lm + 1 & ":D" & lm + 2 & ",D" & lm + 5).NumberFormat = "0.00"
    ws.Range("D" & lm + 1 & ":E" & lm + 5).HorizontalAlignment = xlCenter
    Liste ws.Range("D" & lm + 3), "Non,Oui"
    Liste ws.Range("D" & lm + 4), "Méthode pressiométrique,Méthode pénétrométrique," & _
        "Essais de chargement statique"
    For r = lm + 1 To lm + 5
        ws.Range("B" & r & ":C" & r).Merge
    Next r

    ' --- Caractéristiques des micropieux (sous le tableau précédent) ---
    ' Intitulé (unité comprise) sur B:C fusionnées, valeur en D. Le type de pieu
    ' (tableau A.1 de la NF P94-262) remplit la catégorie, la classe et la prise en
    ' compte de la résistance de pointe ; la norme de référence suit la catégorie
    ' (voir MajTypePieu). La hauteur totale saisie est multipliée par -1 (voir
    ' InverserHauteur).
    lp = LIGNE_PIEU
    SousTitre ws, lp - 1, "Caractéristiques des micropieux"
    ws.Range("B" & lp & ":B" & lp + 15).Value = Application.Transpose(Array( _
        "Type de pieu - technique de mise en oeuvre (NF P94-262, tableau A.1)", _
        "Diamètre constant", "Incliné", "Groupe", "Norme de référence", _
        "Classe du micropieu", "Catégorie du micropieu", "Mode de fonctionnement", _
        "Prise en compte de la résistance de pointe", "Diamètre du micropieu B (m)", _
        "Hauteur totale du micropieu (m)", "Hauteur du micropieu dans la couche porteuse h (m)", _
        "Nombre de micropieux / appui du bâtiment", "Nombre total de micropieux", _
        "Conditions en tête du micropieu", "Essai de contrôle"))
    ListeLongue ws.Range("D" & lp), TypesPieux(), "Pieux NF P94-262"
    ws.Range("D" & lp).Value = Split(TypesPieux(), ",")(16)   ' micropieu type I (cat. 17)
    Liste ws.Range("D" & lp + 1), "Oui,Non"
    Liste ws.Range("D" & lp + 2), "Non,Oui"
    Liste ws.Range("D" & lp + 3), "Non,Oui"
    ' Norme de référence selon la catégorie (D de la ligne Catégorie)
    ws.Range("D" & lp + 4).Formula = "=IF(D" & lp + 6 & "="""","""",IF(VALUE(D" & lp + 6 & _
        "&"""")<=6,""NF EN 1536"",IF(VALUE(D" & lp + 6 & "&"""")<=15,""NF EN 12699""," & _
        "IF(VALUE(D" & lp + 6 & "&"""")=16,""NF EN 12063"",""NF EN 14199""))))"
    Liste ws.Range("D" & lp + 5), "1,1bis,2,3,4,5,6,7,8"
    Liste ws.Range("D" & lp + 6), "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20"
    Liste ws.Range("D" & lp + 7), "Compression,Traction,Compression et traction"
    Liste ws.Range("D" & lp + 8), "Oui,Non"
    ws.Range("D" & lp + 9 & ":D" & lp + 11).NumberFormat = "0.00"
    ws.Range("D" & lp + 12 & ":D" & lp + 13).NumberFormat = "0"
    Liste ws.Range("D" & lp + 14), "Libre en tête,Encastré en tête"
    Liste ws.Range("D" & lp + 15), "Non,Oui"
    ws.Range("D" & lp & ":D" & lp + 15).HorizontalAlignment = xlCenter
    For r = lp To lp + 15
        ws.Range("B" & r & ":C" & r).Merge
    Next r
    MajTypePieu ws

    ' --- Hypothèses relatives aux matériaux constitutifs des fondations profondes ---
    ' Intitulé sur B:C fusionnées, valeur en D, unité en E
    la = lp + 17                             ' titre de la partie matériaux
    With ws.Range("B" & la & ":I" & la)
        .Merge
        .Value = "HYPOTHÈSES RELATIVES AUX MATÉRIAUX CONSTITUTIFS DES FONDATIONS PROFONDES"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    SousTitre ws, la + 2, "Caractéristiques du béton"
    ws.Range("B" & la + 3 & ":B" & la + 4).Value = Application.Transpose(Array( _
        "Classe d'exposition du béton", "Classe de résistance nominale du béton"))
    Liste ws.Range("D" & la + 3), CLASSES_EXPOSITION
    Liste ws.Range("D" & la + 4), "C20/25,C25/30,C30/37,C35/45,C40/50,C45/55,C50/60"
    ws.Range("D" & la + 4).Value = "C25/30"

    ' Acier : la liste des classes dépend du type d'armature, les dimensions du tube
    ' choisi sont remplies (voir MajArmature et MajAcier)
    SousTitre ws, LIGNE_ACIER - 1, "Caractéristiques de l'acier"
    ws.Range("B" & LIGNE_ACIER & ":B" & LIGNE_ACIER + 9).Value = Application.Transpose(Array( _
        "Type d'armature", "Classe de l'acier", "Diamètre extérieur", "Épaisseur e", _
        "Diamètre intérieur", "Module de Young E", "Contrainte de l'acier fe", _
        "Type de platine (liaison micropieu / structure)", "Nuance de la platine", _
        "Contrainte max de l'acier de la platine"))
    ws.Range("E" & LIGNE_ACIER + 2 & ":E" & LIGNE_ACIER + 9).Value = Application.Transpose(Array( _
        "mm", "mm", "mm", "MPa", "MPa", "", "", "MPa"))
    Liste ws.Range("D" & LIGNE_ACIER), "Tube,Barre"
    ws.Range("D" & LIGNE_ACIER + 2 & ":D" & LIGNE_ACIER + 4).NumberFormat = "0.0"
    ws.Range("D" & LIGNE_ACIER + 5).Value = 210000
    ws.Range("D" & LIGNE_ACIER + 6).Value = 500
    ws.Range("D" & LIGNE_ACIER + 5 & ":D" & LIGNE_ACIER + 6).NumberFormat = "0"
    Liste ws.Range("D" & LIGNE_ACIER + 7), PLATINES
    Liste ws.Range("D" & LIGNE_ACIER + 8), "S235,S275,S355,S420,S460"
    ' Limite d'élasticité de la nuance (S235 -> 235 MPa, épaisseur <= 40 mm)
    ws.Range("D" & LIGNE_ACIER + 9).Formula = "=VALUE(MID(D" & LIGNE_ACIER + 8 & ",2,3))"
    ws.Range("D" & la + 3 & ":E" & LIGNE_ACIER + 9).HorizontalAlignment = xlCenter
    For r = la + 3 To LIGNE_ACIER + 9
        If r <> LIGNE_ACIER - 1 And r <> LIGNE_ACIER - 2 Then ws.Range("B" & r & ":C" & r).Merge
    Next r

    ' --- Mise en forme ---
    With ws.Range("B4:C9,B12:C16,B19:I26,B28:E29,B32:C32,B" & LIGNE_COUCHES - 1 & ":I" & lc & _
            ",B" & lm + 1 & ":E" & lm + 5 & ",B" & lp & ":D" & lp + 15 & _
            ",B" & la + 3 & ":E" & la + 4 & ",B" & LIGNE_ACIER & ":E" & LIGNE_ACIER + 9)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("B19:I19,B" & LIGNE_COUCHES - 1 & ":I" & LIGNE_COUCHES - 1).Font.Bold = True
    ws.Range("B" & LIGNE_COUCHES - 1 & ":I" & LIGNE_COUCHES - 1).HorizontalAlignment = xlCenter
    ' Cases à remplir (données d'entrée) en jaune ; épaisseur et norme sont calculées
    ws.Range("C4:C9,C12:C16,D28:D29,C32,C" & LIGNE_COUCHES & ":E" & lc & ",G" & LIGNE_COUCHES & _
        ":I" & lc & ",D" & lm + 1 & ":D" & lm + 5 & ",D" & lp & ":D" & lp + 3 & _
        ",D" & lp + 5 & ":D" & lp + 15 & ",D" & la + 3 & ":D" & la + 4 & _
        ",D" & LIGNE_ACIER & ":D" & LIGNE_ACIER + 1 & ",D" & LIGNE_ACIER + 5 & ":D" & _
        LIGNE_ACIER + 8).Interior.Color = RGB(255, 255, 204)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 40
    ws.Columns("D").ColumnWidth = 28         ' valeurs des micropieux, catégorie de sol
    ws.Columns("E:I").ColumnWidth = 14
    ws.Rows("19:" & LIGNE_ACIER + 9).AutoFit

    AfficherTypeFondation ws
    AfficherCouches ws
    MajArmature ws
    ws.Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Affiche autant de lignes de couches que le nombre saisi en C32 (1 à MAX_COUCHES),
' les autres lignes du tableau sont masquées. Appelée aussi quand C32 change
' (voir CodeFeuilleDonnees.txt).
Public Sub AfficherCouches(ws As Worksheet)
    Dim nb As Long
    nb = Val(ws.Range(CELLULE_NB_COUCHES).Value)
    If nb < 1 Then nb = 1
    If nb > MAX_COUCHES Then nb = MAX_COUCHES
    ws.Rows(LIGNE_COUCHES).Resize(MAX_COUCHES).Hidden = False
    If nb < MAX_COUCHES Then ws.Rows(LIGNE_COUCHES + nb).Resize(MAX_COUCHES - nb).Hidden = True
End Sub

' Type de pieu choisi (D58) : la catégorie est le rang de la technique dans le
' tableau A.1 (TypesPieux, « Micropieu type I (M1) » -> 17), la classe s'en déduit
' (ClassePieu) et la résistance de pointe est prise en compte (Oui) sauf pour les
' micropieux, catégories 17 à 20 (Non) ; la norme de référence suit la catégorie
' par formule. Appelée aussi quand le type change (voir CodeFeuilleDonnees.txt).
Public Sub MajTypePieu(ws As Worksheet)
    Dim categorie As Long, types As Variant, i As Long
    types = Split(TypesPieux(), ",")
    For i = 0 To UBound(types)
        If types(i) = CStr(ws.Range("D" & LIGNE_PIEU).Value) Then categorie = i + 1
    Next i
    If categorie = 0 Then Exit Sub
    ws.Range("D" & LIGNE_PIEU + 6).Value = categorie
    ws.Range("D" & LIGNE_PIEU + 5).Value = ClassePieu(categorie)
    ws.Range("D" & LIGNE_PIEU + 8).Value = IIf(categorie >= 17, "Non", "Oui")
End Sub

' Hauteur totale du micropieu (D68) : la valeur saisie est multipliée par -1
' (ex. cote -12 saisie -> 12 m). Appelée quand la cellule change
' (voir CodeFeuilleDonnees.txt).
Public Sub InverserHauteur(c As Range)
    If IsNumeric(c.Value) And Not IsEmpty(c.Value) Then c.Value = -c.Value
End Sub

' Type d'armature choisi (D82) : la liste « Classe de l'acier » propose les tubes
' (TUBES) ou les barres (BARRES) ; diamètres et épaisseur ne sont affichés que pour
' un tube. Appelée aussi quand le type change (voir CodeFeuilleDonnees.txt).
Public Sub MajArmature(ws As Worksheet)
    Dim tube As Boolean, liste As String
    tube = (ws.Range("D" & LIGNE_ACIER).Value = "Tube")
    liste = IIf(tube, TUBES, BARRES)
    With ws.Range("D" & LIGNE_ACIER + 1)
        .Validation.Delete
        .Validation.Add xlValidateList, xlValidAlertStop, , liste
        If InStr("," & liste & ",", "," & .Value & ",") = 0 Then .Value = Split(liste, ",")(0)
    End With
    ws.Rows(LIGNE_ACIER + 2).Resize(3).Hidden = Not tube
    MajAcier ws
End Sub

' Classe de l'acier choisie (D83) : pour un tube « ... Øde x e mm », diamètre
' extérieur de et épaisseur e sont lus dans la désignation, diamètre intérieur =
' de - 2 e. Appelée aussi quand la classe change (voir CodeFeuilleDonnees.txt).
Public Sub MajAcier(ws As Worksheet)
    Dim nom As String, p As Long
    nom = CStr(ws.Range("D" & LIGNE_ACIER + 1).Value)
    p = InStr(nom, "Ø")
    If ws.Range("D" & LIGNE_ACIER).Value = "Tube" And p > 0 Then
        ws.Range("D" & LIGNE_ACIER + 2).Value = Val(Mid(nom, p + 1))
        ws.Range("D" & LIGNE_ACIER + 3).Value = Val(Mid(nom, InStr(p, nom, "x") + 1))
        ws.Range("D" & LIGNE_ACIER + 4).Formula = "=D" & LIGNE_ACIER + 2 & "-2*D" & LIGNE_ACIER + 3
    Else
        ws.Range("D" & LIGNE_ACIER + 2 & ":D" & LIGNE_ACIER + 4).ClearContents
    End If
End Sub

' Classe de pieu (NF P94-262, tableau A.1) selon la catégorie 1 à 20
Public Function ClassePieu(ByVal categorie As Long) As String
    Select Case categorie
        Case 1 To 5:   ClassePieu = "1"      ' pieux forés
        Case 6:        ClassePieu = "2"      ' tarière creuse
        Case 7, 8:     ClassePieu = "3"      ' vissés
        Case 9 To 12:  ClassePieu = "4"      ' battus fermés
        Case 13:       ClassePieu = "5"      ' battu acier ouvert
        Case 14, 15:   ClassePieu = "6"      ' profilés H
        Case 16:       ClassePieu = "7"      ' palplanches
        Case 17, 18:   ClassePieu = "1bis"   ' micropieux types I et II
        Case 19, 20:   ClassePieu = "8"      ' injectés haute pression (IGU, IRS)
    End Select
End Function

' Techniques de mise en oeuvre des pieux (NF P94-262, tableau A.1), dans l'ordre
' des catégories 1 à 20. Désignations abrégées, à vérifier sur la norme.
Private Function TypesPieux() As String
    Dim s As String
    s = "Foré simple (pieux et barrettes) (FS)"
    s = s & ",Foré boue (pieux et barrettes) (FB)"
    s = s & ",Foré tubé - virole perdue (FTP)"
    s = s & ",Foré tubé - virole récupérée (FTR)"
    s = s & ",Foré simple ou boue avec rainurage ou puits (FSR / FBR / PU)"
    s = s & ",Foré tarière creuse simple ou double rotation (FTC / FTCD)"
    s = s & ",Vissé moulé (VM)"
    s = s & ",Vissé tubé (VT)"
    s = s & ",Battu béton préfabriqué ou précontraint (BPF / BPR)"
    s = s & ",Battu enrobé - béton / mortier / coulis (BE)"
    s = s & ",Battu moulé (BM)"
    s = s & ",Battu acier fermé (BAF)"
    s = s & ",Battu acier ouvert (BAO)"
    s = s & ",Profilé H battu (HB)"
    s = s & ",Profilé H battu injecté (HBi)"
    s = s & ",Palplanches battues (PP)"
    s = s & ",Micropieu type I (M1)"
    s = s & ",Micropieu type II (M2)"
    s = s & ",Pieu ou micropieu injecté mode IGU - type III (PIGU / MIGU)"
    s = s & ",Pieu ou micropieu injecté mode IRS - type IV (PIRS / MIRS)"
    TypesPieux = s
End Function

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
    s = s & ",Terrain neutralisé"
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
