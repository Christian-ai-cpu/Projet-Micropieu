Attribute VB_Name = "DonneesProjet"
Option Explicit

' Module « Données du projet » : crée (ou recrée) la feuille des hypothèses générales.
'   B1 titre ; B3 synthèse du projet (phase en C9) ; B11 caractérisation de l'ouvrage
'   (type de fondation en ligne 13, affiché seulement en phase G5) ; B18 références
'   normatives ; B28 référentiel altimétrique et niveau du TN ; B31 données
'   géotechniques (nombre de couches en C32, couches en lignes 34 à 48) ; B50 données
'   des micropieux ; B57 caractéristiques des micropieux (type de pieu en D58, le reste
'   s'en déduit par formules, tableau A.1 de la NF P94-262) ; B77 matériaux : béton
'   (B79) et acier (B88, type d'armature en D89).
' Intitulé sur B:C fusionnées, valeur en D, unité en E ; saisies dans les cases jaunes.
' Réactions aux saisies : ChangementDonnees (appelée par CodeFeuilleDonnees.txt).
' Listes de plus de 255 caractères : feuille masquée « Listes » (ListeLongue).

Private Const NOM_FEUILLE As String = "Données du projet"
Private Const CELLULE_PHASE As String = "C9"
Private Const LIGNE_FONDATION As Long = 13
' Constantes publiques : utilisées aussi par le module CalculMicropieu
Public Const CELLULE_NB_COUCHES As String = "C32"
Public Const LIGNE_COUCHES As Long = 34       ' première couche de sol
Public Const MAX_COUCHES As Long = 15         ' lignes prévues (au-delà du nombre : masquées)
Public Const LIGNE_PIEU As Long = 58          ' type de pieu
Private Const LIGNE_ACIER As Long = 89         ' type d'armature
Public Const CATEGORIES_SOL As String = "Argiles et limons," & _
    "Sols intermédiaires (sable limoneux / sable argileux / argile sableuse)," & _
    "Sables et graves,Craies,Marnes et calcaires marneux,Roches"
' Tableau A.1 de la NF P94-262 : techniques (catégories 1 à 20) et abréviations
Private Const TYPES_PIEUX As String = "Foré simple (pieux et barrettes)," & _
    "Foré boue (pieux et barrettes),Foré tubé - virole perdue,Foré tubé - virole récupérée," & _
    "Foré simple ou boue avec rainurage ou puits,Foré tarière creuse simple ou double rotation," & _
    "Vissé moulé,Vissé tubé,Battu béton préfabriqué ou précontraint," & _
    "Battu enrobé - béton / mortier / coulis,Battu moulé,Battu acier fermé,Battu acier ouvert," & _
    "Profilé H battu,Profilé H battu injecté,Palplanches battues,Micropieu type I," & _
    "Micropieu type II,Pieu ou micropieu injecté mode IGU - type III," & _
    "Pieu ou micropieu injecté mode IRS - type IV"
Private Const ABREVIATIONS As String = "FS|FB|FTP|FTR|FSR, FBR, PU|FTC, FTCD|VM|VT|" & _
    "BPF, BPR|BE|BM|BAF|BAO|HB|HBi|PP|M1|M2|PIGU, MIGU|PIRS, MIRS"
' Sols et roches des classes A, B, C, D et R de la NF P11-300 (désignations abrégées)
Private Const SOLS_NFP11300 As String = "A1 - Limons peu plastiques / loess / sables fins peu pollués," & _
    "A2 - Sables fins argileux / limons / argiles et marnes peu plastiques," & _
    "A3 - Argiles et argiles marneuses / limons très plastiques," & _
    "A4 - Argiles et argiles marneuses très plastiques,B1 - Sables silteux," & _
    "B2 - Sables argileux (peu argileux),B3 - Graves silteuses,B4 - Graves argileuses (peu argileuses)," & _
    "B5 - Sables et graves très silteux,B6 - Sables et graves argileux à très argileux," & _
    "C1 - Sols à fines et gros éléments (argiles à silex / éboulis / moraines)," & _
    "C2 - Sols à gros éléments dominants avec fines,D1 - Sables alluvionnaires propres / sables de dune," & _
    "D2 - Graves alluvionnaires propres / sables,D3 - Graves alluvionnaires grossières propres / galets," & _
    "R1 - Craies,R2 - Calcaires,R3 - Roches argileuses (marnes / argilites / pélites)," & _
    "R4 - Roches siliceuses (grès / poudingues / brèches),R5 - Roches salines (sel gemme / gypse)," & _
    "R6 - Roches magmatiques et métamorphiques (granites / basaltes / schistes),Terrain neutralisé"
Private Const CLASSES_EXPOSITION As String = "X0,XC1,XC2,XC3,XC4,XD1,XD2,XD3,XS1,XS2,XS3," & _
    "XF1,XF2,XF3,XF4,XA1,XA2,XA3"
' Armatures (dimensions à vérifier auprès du fournisseur) ; tube : « Øde x e mm » lu
' par MajAcier ; barres filetées GEWI, GEWI Plus, DYWIDAG dans la liste des platines
Private Const TUBES As String = "Ø60.3 x 5.5 mm,Ø73 x 5.5 mm,Ø88.9 x 6.5 mm,Ø88.9 x 9.5 mm," & _
    "Ø101.6 x 9 mm,Ø114.3 x 9 mm,Ø127 x 9 mm,Ø139.7 x 9 mm,Ø168.3 x 10 mm,Ø177.8 x 10.4 mm"
Private Const BARRES As String = "Barre HA B500 Ø20,Barre HA B500 Ø25,Barre HA B500 Ø32,Barre HA B500 Ø40"
Private Const PLATINES As String = "GEWI Ø25 (S555/700),GEWI Ø28 (S555/700)," & _
    "GEWI Ø32 (S555/700),GEWI Ø40 (S555/700),GEWI Ø50 (S555/700),GEWI Ø63.5 (S555/700)," & _
    "GEWI Plus Ø18 (S670/800),GEWI Plus Ø22 (S670/800),GEWI Plus Ø25 (S670/800)," & _
    "GEWI Plus Ø28 (S670/800),GEWI Plus Ø30 (S670/800),GEWI Plus Ø35 (S670/800)," & _
    "GEWI Plus Ø43 (S670/800),GEWI Plus Ø57.5 (S670/800),GEWI Plus Ø63.5 (S670/800)," & _
    "GEWI Plus Ø75 (S670/800),DYWIDAG Ø18 (Y1050),DYWIDAG Ø26.5 (Y1050)," & _
    "DYWIDAG Ø32 (Y1050),DYWIDAG Ø36 (Y1050),DYWIDAG Ø40 (Y1050),DYWIDAG Ø47 (Y1050)," & _
    "DYWIDAG Ø57 (Y1050),DYWIDAG Ø65 (Y1050),DYWIDAG Ø75 (Y1050)"
' Nuances (NF EN 10025-2, -3 N/NL, -4 M/ML), fy et fu en MPa pour t <= 40 mm
' (NF EN 1993-1-1, tableau 3.1)
Private Const NUANCES As String = "S235,S275,S355,S450,S275N,S275NL,S355N,S355NL," & _
    "S420N,S420NL,S460N,S460NL,S275M,S275ML,S355M,S355ML,S420M,S420ML,S460M,S460ML"
Private Const NUANCES_FY As String = "235,275,355,440,275,275,355,355,420,420,460,460," & _
    "275,275,355,355,420,420,460,460"
Private Const NUANCES_FU As String = "360,430,510,550,390,390,490,490,520,520,540,540," & _
    "370,370,470,470,520,520,540,540"

Public Sub CreerDonneesProjet()
    Dim ws As Worksheet, r As Long, n As Variant, lc As Long, lm As Long, lp As Long
    Dim la As Long, c As String, g As String

    Application.ScreenUpdating = False
    Application.EnableEvents = False         ' pas de Worksheet_Change pendant la création
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(Before:=ActiveWorkbook.Worksheets(1))
        ws.Name = NOM_FEUILLE
    End If
    ws.Cells.Clear
    ws.Rows.Hidden = False
    g = ChrW(947)                            ' gamma
    Titre ws.Range("B1:I1"), "HYPOTHÈSES GÉNÉRALES"

    ' --- Synthèse du projet et caractérisation de l'ouvrage (valeur en C) ---
    SousTitre ws, 3, "Synthèse du projet"
    Colonne ws.Range("B4"), Array("Nom du projet", "Lieu", "Maître d'ouvrage", "Date", _
        "Numéro du projet", "Phase du projet")
    ws.Range("C7").Value = Date
    ws.Range("C7").NumberFormat = "dd/mm/yyyy"
    Liste ws.Range(CELLULE_PHASE), "G2-AVP,G2-PRO,G5"
    SousTitre ws, 11, "Caractérisation générale de l'ouvrage"
    Colonne ws.Range("B12"), Array("Type d'ouvrage", "Type de fondation", _
        "Catégorie de bâtiment", "Altitude", "Zone de sismicité")
    Liste ws.Range("C12"), "Bâtiment,Pont"
    Liste ws.Range("C13"), "Fondations superficielles,Fondations semi-profondes," & _
        "Fondations profondes (pieux),Micropieux,Radier"
    Liste ws.Range("C14"), "A,B,C,D,E,F,G"
    Liste ws.Range("C15"), "<=1000 m,>1000 m"
    Liste ws.Range("C16"), "1,2,3,4,5"

    ' --- Références normatives : référence, date, intitulé (fusionné de D à I) ---
    SousTitre ws, 18, "Références normatives"
    n = Array("Référence|Date|Intitulé", _
        "NF EN 1990|mars 2003|Eurocode 0 - Bases de calcul des structures", _
        "NF EN 1992-1-1|octobre 2005|Eurocode 2 - Calcul des structures en béton - Partie 1-1 : " & _
            "règles générales et règles pour les bâtiments", _
        "NF P22-311-1|octobre 2005|Eurocode 3 (NF EN 1993-1-1) - Calcul des structures en acier - " & _
            "Partie 1-1 : règles générales et règles pour les bâtiments", _
        "NF P94-251-1|juin 2005|Eurocode 7 (NF EN 1997-1) - Calcul géotechnique - Partie 1 : règles générales", _
        "NF P94-262|juillet 2012|Justification des ouvrages géotechniques - Norme d'application " & _
            "nationale de l'Eurocode 7 - Fondations profondes", _
        "NF EN 1998-1|septembre 2005|Eurocode 8 - Calcul des structures pour leur résistance aux " & _
            "séismes - Partie 1 : règles générales, actions sismiques et règles pour les bâtiments", _
        "NF EN 1998-5|septembre 2005|Eurocode 8 - Partie 5 : fondations, ouvrages de soutènement " & _
            "et aspects géotechniques")
    For r = 0 To UBound(n)
        ws.Range("B" & 19 + r & ":D" & 19 + r).Value = Split(n(r), "|")
        ws.Range("D" & 19 + r & ":I" & 19 + r).Merge
    Next r

    ' --- Référentiel altimétrique et niveau du terrain naturel (m + référentiel) ---
    Colonne ws.Range("B28"), Array("Référentiel altimétrique", "Niveau du terrain naturel"), True
    Liste ws.Range("D28"), "NGF-IGN69,NGF-IGN78,NVP,IGN,Local"
    ws.Range("D29").NumberFormat = "0.00"
    ws.Range("E29").Formula = "=""m ""&D28"
    ws.Range("D28:E29").HorizontalAlignment = xlCenter

    ' --- Données géotechniques : nombre de couches puis une ligne par couche ---
    SousTitre ws, 31, "Données géotechniques"
    ws.Range("B32").Value = "Nombre de couches"
    With ws.Range(CELLULE_NB_COUCHES)
        .Value = 3
        .Validation.Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, "1", CStr(MAX_COUCHES)
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("B33:I33").Value = Array("N° de couche", "Nom de couche (NF P11-300)", _
        "Catégorie de sol (NF P94-262)", "Prof. bas (m/TN)", "Épaisseur (m)", _
        "Pression limite pl* (MPa)", "Module pressiométrique EM (MPa)", "Terrain neutralisé")
    lc = LIGNE_COUCHES + MAX_COUCHES - 1
    For r = LIGNE_COUCHES To lc
        ws.Range("B" & r).Value = r - LIGNE_COUCHES + 1
        ListeLongue ws.Range("C" & r), SOLS_NFP11300, "Sols NF P11-300"
        Liste ws.Range("D" & r), CATEGORIES_SOL
        ws.Range("C" & r & ":D" & r).ClearContents      ' nom et catégorie à choisir
        Liste ws.Range("I" & r), "Non,Oui"
        ' Épaisseur = différence avec la base de la couche au-dessus (TN pour la 1re)
        ws.Range("F" & r).Formula = "=IF(E" & r & "="""","""",ABS(E" & r & _
            IIf(r = LIGNE_COUCHES, "", "-E" & r - 1) & "))"
    Next r
    ws.Range("E" & LIGNE_COUCHES & ":H" & lc).NumberFormat = "0.00"
    ws.Range("B" & LIGNE_COUCHES & ":B" & lc & ",E" & LIGNE_COUCHES & ":I" & lc).HorizontalAlignment = xlCenter

    ' --- Données des micropieux ---
    lm = lc + 2
    SousTitre ws, lm, "Données des micropieux"
    Colonne ws.Range("B" & lm + 1), Array("Cote de l'arase supérieure des micropieux", _
        "Base des micropieux", "Présence de nappe phréatique", _
        "Méthode de calcul retenue pour les pieux", "Niveau de l'eau"), True
    ws.Range("E" & lm + 1 & ":E" & lm + 5).Value = Application.Transpose(Array("m/TN", "m/TN", "", "", "m"))
    ws.Range("D" & lm + 1 & ":D" & lm + 5).NumberFormat = "0.00"
    Liste ws.Range("D" & lm + 3), "Non,Oui"
    Liste ws.Range("D" & lm + 4), "Méthode pressiométrique,Méthode pénétrométrique," & _
        "Essais de chargement statique"

    ' --- Caractéristiques des micropieux : tout se déduit du type de pieu choisi
    ' (catégorie = rang dans TYPES_PIEUX), sauf le mode de fonctionnement ---
    lp = LIGNE_PIEU
    SousTitre ws, lp - 1, "Caractéristiques des micropieux"
    Colonne ws.Range("B" & lp), Array("Type de pieu - technique de mise en oeuvre " & _
        "(NF P94-262, tableau A.1)", "Diamètre constant", "Incliné", "Groupe", _
        "Norme de référence", "Classe du micropieu", "Catégorie du micropieu", _
        "Abréviation du type de pieu (NF P94-262, tableau A.1)", "Mode de fonctionnement", _
        "Prise en compte de la résistance de pointe", "Type de mise en oeuvre", _
        "Diamètre du micropieu B (m)", "Hauteur totale du micropieu (m)", _
        "Hauteur totale du micropieu dans la couche porteuse D (m)", _
        "Nombre de micropieux / appui du bâtiment", "Nombre total de micropieux", _
        "Conditions en tête du micropieu", "Essai de contrôle"), True
    ListeLongue ws.Range("D" & lp), TYPES_PIEUX, "Pieux NF P94-262"
    ws.Range("D" & lp).Value = Split(TYPES_PIEUX, ",")(16)   ' micropieu type I (cat. 17)
    Liste ws.Range("D" & lp + 1), "Oui,Non"
    Liste ws.Range("D" & lp + 2), "Non,Oui"
    Liste ws.Range("D" & lp + 3), "Non,Oui"
    c = "D" & lp + 6                         ' catégorie
    ws.Range(c).Formula = "=IFERROR(MATCH(D" & lp & "," & Tableau(TYPES_PIEUX, ",") & ",0),"""")"
    ' Norme : 1 à 6 NF EN 1536, 7 à 16 NF EN 12699, 17 à 20 les trois
    ws.Range("D" & lp + 4).Formula = "=IF(" & c & "="""","""",IF(" & c & "<=6,""NF EN 1536""," & _
        "IF(" & c & "<=16,""NF EN 12699"",""NF EN 1536 / 14199 / 12699"")))"
    ws.Range("D" & lp + 5).Formula = "=IF(" & c & "="""","""",INDEX(" & _
        Tableau("1,1,1,1,1,2,3,3,4,4,4,4,5,6,6,7,1 bis,1 bis,8,8", ",") & "," & c & "))"
    ws.Range("D" & lp + 7).Formula = "=IF(" & c & "="""","""",INDEX(" & Tableau(ABREVIATIONS, "|") & _
        "," & c & "))"
    Liste ws.Range("D" & lp + 8), "Compression,Traction"
    ' Résistance de pointe : non prise en compte pour les micropieux (catégories 17 à 20)
    ws.Range("D" & lp + 9).Formula = "=IF(" & c & "="""","""",IF(" & c & ">=17,""Non"",""Oui""))"
    ' Type de mise en oeuvre (colonnes des tableaux C.2.3 de la norme)
    ws.Range("D" & lp + 10).Formula = "=IF(" & c & "="""","""",IF(" & c & "=6," & _
        """Pieu à la tarière continue (CFA)"",IF(AND(" & c & ">=7," & c & "<=16)," & _
        """Pieu foncé"",""Pieu foré"")))"
    ws.Range("D" & lp + 11 & ":D" & lp + 13).NumberFormat = "0.00"
    ws.Range("D" & lp + 14 & ":D" & lp + 15).NumberFormat = "0"
    Liste ws.Range("D" & lp + 16), "Libre en tête,Encastré en tête"
    Liste ws.Range("D" & lp + 17), "Non,Oui"

    ' --- Matériaux : béton puis acier ---
    la = lp + 19
    Titre ws.Range("B" & la & ":I" & la), "HYPOTHÈSES RELATIVES AUX MATÉRIAUX CONSTITUTIFS DES " & _
        "FONDATIONS PROFONDES"
    SousTitre ws, la + 2, "Caractéristiques du béton"
    Colonne ws.Range("B" & la + 3), Array("Classe d'exposition du béton", "Classe structurale", _
        "Classe de résistance nominale du béton", "Résistance du béton à 28 jours fc28", _
        "Coefficient " & g & "c - combinaison ELU fondamentale", _
        "Coefficient " & g & "c - combinaison accidentelle", _
        "Coefficient " & g & "c - combinaison sismique"), True
    ws.Range("E" & la + 6).Value = "MPa"
    Liste ws.Range("D" & la + 3), CLASSES_EXPOSITION
    Liste ws.Range("D" & la + 4), "S1,S2,S3,S4,S5,S6", "S4"
    Liste ws.Range("D" & la + 5), "C20/25,C25/30,C30/37,C35/45,C40/50,C45/55,C50/60", "C25/30"
    ' fc28 = résistance sur cylindre de la classe (C25/30 -> 25 MPa)
    ws.Range("D" & la + 6).Formula = "=VALUE(MID(D" & la + 5 & ",2,FIND(""/"",D" & la + 5 & ")-2))"
    ws.Range("D" & la + 7 & ":D" & la + 9).Value = Application.Transpose(Array(1.5, 1.2, 1.3))
    ws.Range("D" & la + 6 & ":D" & la + 9).NumberFormat = "0.00"

    SousTitre ws, LIGNE_ACIER - 1, "Caractéristiques de l'acier"
    Colonne ws.Range("B" & LIGNE_ACIER), Array("Type d'armature", "Classe de l'acier", _
        "Diamètre extérieur", "Épaisseur e", "Diamètre intérieur", "Module de Young E", _
        "Contrainte de l'acier fe", "Coefficient " & g & "s - combinaison ELU fondamentale", _
        "Coefficient " & g & "s - combinaison accidentelle", _
        "Coefficient " & g & "s - combinaison sismique", _
        "Type de platine (liaison micropieu / structure)", "Nuance de l'acier de la platine", _
        "Limite élastique de l'acier fy", "Limite de rupture de l'acier fu"), True
    ws.Range("E" & LIGNE_ACIER + 2 & ":E" & LIGNE_ACIER + 13).Value = Application.Transpose(Array( _
        "mm", "mm", "mm", "MPa", "MPa", "", "", "", "", "", "MPa", "MPa"))
    Liste ws.Range("D" & LIGNE_ACIER), "Tube,Barre"
    ws.Range("D" & LIGNE_ACIER + 2 & ":D" & LIGNE_ACIER + 4).NumberFormat = "0.0"
    ws.Range("D" & LIGNE_ACIER + 5 & ":D" & LIGNE_ACIER + 9).Value = _
        Application.Transpose(Array(210000, 500, 1.15, 1, 1))
    ws.Range("D" & LIGNE_ACIER + 7 & ":D" & LIGNE_ACIER + 9).NumberFormat = "0.00"
    ListeLongue ws.Range("D" & LIGNE_ACIER + 10), PLATINES, "Platines"
    Liste ws.Range("D" & LIGNE_ACIER + 11), NUANCES, "S355"
    ' fy et fu de la nuance choisie
    For r = 12 To 13
        ws.Range("D" & LIGNE_ACIER + r).Formula = "=INDEX({" & IIf(r = 12, NUANCES_FY, NUANCES_FU) & _
            "},MATCH(D" & LIGNE_ACIER + 11 & "," & Tableau(NUANCES, ",") & ",0))"
    Next r

    ' --- Mise en forme ---
    With ws.Range("B4:C9,B12:C16,B19:I26,B28:E29,B32:C32,B33:I" & lc & ",B" & lm + 1 & ":E" & _
            lm + 5 & ",B" & lp & ":D" & lp + 17 & ",B" & la + 3 & ":E" & la + 9 & ",B" & _
            LIGNE_ACIER & ":E" & LIGNE_ACIER + 13)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("B19:I19,B33:I33").Font.Bold = True
    ws.Range("B33:I33,D" & lm + 1 & ":E" & LIGNE_ACIER + 13).HorizontalAlignment = xlCenter
    ' Cases à remplir en jaune ; les valeurs calculées restent blanches
    ws.Range("C4:C9,C12:C16,D28:D29,C32,C" & LIGNE_COUCHES & ":E" & lc & ",G" & LIGNE_COUCHES & _
        ":I" & lc & ",D" & lm + 1 & ":D" & lm + 5 & ",D" & lp & ":D" & lp + 3 & ",D" & lp + 8 & _
        ",D" & lp + 11 & ":D" & lp + 17 & ",D" & la + 3 & ":D" & la + 5 & ",D" & la + 7 & ":D" & _
        la + 9 & ",D" & LIGNE_ACIER & ":D" & LIGNE_ACIER + 1 & ",D" & LIGNE_ACIER + 5 & ":D" & _
        LIGNE_ACIER + 11).Interior.Color = RGB(255, 255, 204)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 40
    ws.Columns("D").ColumnWidth = 28
    ws.Columns("E:I").ColumnWidth = 14
    ws.Rows("19:" & LIGNE_ACIER + 13).AutoFit

    AfficherTypeFondation ws
    AfficherCouches ws
    MajArmature ws
    ws.Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Appelée par le code de la feuille (CodeFeuilleDonnees.txt) à chaque modification :
' phase -> type de fondation ; nombre de couches -> lignes de couches ; hauteur totale
' -> x (-1) ; type d'armature -> classes d'acier ; classe de l'acier -> dimensions.
Public Sub ChangementDonnees(ws As Worksheet, Target As Range)
    On Error GoTo Fin
    Application.EnableEvents = False
    If Not Intersect(Target, ws.Range(CELLULE_PHASE)) Is Nothing Then AfficherTypeFondation ws
    If Not Intersect(Target, ws.Range(CELLULE_NB_COUCHES)) Is Nothing Then AfficherCouches ws
    With ws.Range("D" & LIGNE_PIEU + 12)     ' hauteur totale : valeur saisie x (-1)
        If Not Intersect(Target, .Cells) Is Nothing And IsNumeric(.Value) And Not IsEmpty(.Value) Then _
            .Value = -.Value
    End With
    If Not Intersect(Target, ws.Range("D" & LIGNE_ACIER)) Is Nothing Then MajArmature ws
    If Not Intersect(Target, ws.Range("D" & LIGNE_ACIER + 1)) Is Nothing Then MajAcier ws
Fin:
    Application.EnableEvents = True
End Sub

' Ligne « Type de fondation » affichée seulement en phase G5
Public Sub AfficherTypeFondation(ws As Worksheet)
    ws.Rows(LIGNE_FONDATION).Hidden = (ws.Range(CELLULE_PHASE).Value <> "G5")
End Sub

' Autant de lignes de couches affichées que le nombre saisi (1 à MAX_COUCHES)
Public Sub AfficherCouches(ws As Worksheet)
    Dim nb As Long
    nb = Application.Max(1, Application.Min(MAX_COUCHES, Val(ws.Range(CELLULE_NB_COUCHES).Value)))
    ws.Rows(LIGNE_COUCHES).Resize(MAX_COUCHES).Hidden = False
    If nb < MAX_COUCHES Then ws.Rows(LIGNE_COUCHES + nb).Resize(MAX_COUCHES - nb).Hidden = True
End Sub

' Type d'armature : liste des tubes ou des barres HA ; dimensions affichées pour un tube
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

' Tube « Øde x e mm » : diamètre extérieur, épaisseur, diamètre intérieur = de - 2 e
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

' Constante de tableau pour une formule : "a,b" -> {"a","b"} (séparateur sep)
Private Function Tableau(liste As String, sep As String) As String
    Tableau = "{""" & Replace(liste, sep, """,""") & """}"
End Function

' Intitulés écrits en colonne à partir de la cellule c ; fusion avec la colonne
' suivante si fusion (intitulé sur B:C, valeur en D)
Private Sub Colonne(c As Range, intitules As Variant, Optional fusion As Boolean = False)
    Dim k As Long
    For k = 0 To UBound(intitules)
        c.Offset(k).Value = intitules(k)
        If fusion Then c.Offset(k).Resize(1, 2).Merge
    Next k
End Sub

' Grand titre centré sur la plage
Private Sub Titre(plage As Range, texte As String)
    plage.Merge
    plage.Value = texte
    plage.Font.Bold = True
    plage.Font.Size = 14
    plage.HorizontalAlignment = xlCenter
End Sub

' Sous-titre en gras, souligné, en colonne B
Private Sub SousTitre(ws As Worksheet, ligne As Long, texte As String)
    With ws.Range("B" & ligne).Font
        .Bold = True
        .Underline = xlUnderlineStyleSingle
        .Size = 12
    End With
    ws.Range("B" & ligne).Value = texte
End Sub

' Liste déroulante sur une cellule ; valeur de départ = premier choix ou valeur
Private Sub Liste(c As Range, choix As String, Optional valeur As String = "")
    c.Validation.Delete
    c.Validation.Add xlValidateList, xlValidAlertStop, , choix
    c.Value = IIf(valeur = "", Split(choix, ",")(0), valeur)
    c.HorizontalAlignment = xlCenter
End Sub

' Liste de plus de 255 caractères : choix écrits dans une colonne de la feuille masquée
' « Listes » (titre nom en ligne 1), la liste déroulante y renvoie
Private Sub ListeLongue(c As Range, choix As String, nom As String)
    Dim wl As Worksheet, t As Variant, col As Long
    If Len(choix) <= 255 Then Liste c, choix: Exit Sub
    On Error Resume Next
    Set wl = c.Worksheet.Parent.Worksheets("Listes")
    On Error GoTo 0
    If wl Is Nothing Then
        Set wl = c.Worksheet.Parent.Worksheets.Add(After:=c.Worksheet.Parent.Worksheets(c.Worksheet.Parent.Worksheets.Count))
        wl.Name = "Listes"
        wl.Visible = xlSheetHidden
        c.Worksheet.Activate                 ' la feuille ajoutée était devenue active
    End If
    col = 1
    Do While wl.Cells(1, col).Value <> "" And wl.Cells(1, col).Value <> nom
        col = col + 1
    Loop
    t = Split(choix, ",")
    wl.Columns(col).ClearContents
    wl.Cells(1, col).Value = nom
    wl.Cells(2, col).Resize(UBound(t) + 1).Value = Application.Transpose(t)
    c.Validation.Delete
    c.Validation.Add xlValidateList, xlValidAlertStop, , "=Listes!" & wl.Cells(2, col).Resize(UBound(t) + 1).Address
    c.Value = t(0)
End Sub
