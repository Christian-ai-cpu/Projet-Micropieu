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
'               m + référentiel, reprise pour les profondeurs et les cotes)
'   B31 à E36 : Données de la fondation (arase supérieure, profondeur de
'               reconnaissance, nappe, méthode de calcul, niveau d'eau ; intitulé sur
'               B:C, valeur en D, unité en E)
'   B38 à K55 : Données géotechniques : nombre de couches (C39, 1 à 15) et tableau
'               des couches de sol (seules les lignes jusqu'à ce nombre sont affichées) ;
'               colonne pl* masquée en méthode pénétrométrique, colonne qc masquée en
'               méthode pressiométrique (AfficherMethode) ; couche d'ancrage en K
'   B57 à D77 : Caractéristiques de la fondation (type de pieu du tableau A.1 de la
'               NF P94-262, diamètre constant, incliné et angle, effet de groupe et m,
'               n, entraxe d (lignes affichées si « Oui »), norme de référence,
'               classe, catégorie, abréviation, mode de fonctionnement (compression / traction),
'               résistance de pointe, type de mise en oeuvre, B, nombres,
'               conditions en tête, essai de contrôle ; intitulé sur B:C, valeur en D)
'   B79       : HYPOTHÈSES RELATIVES AUX MATÉRIAUX CONSTITUTIFS DES FONDATIONS PROFONDES
'   B81 à E89 : Caractéristiques du béton (classe d'exposition, classe de résistance
'               minimale associée (NF EN 206/CN) et conformité de la classe nominale,
'               classe structurale, classe de résistance, fc28, coefficients gamma c ELU /
'               accidentel / sismique)
'   B91 à E110: Caractéristiques de l'acier (type d'armature Tube N80, Barre HA ou Barre
'               GEWI, désignation, dimensions, section, E et fyk associés, coefficients
'               gamma s, durée d'utilisation du projet, environnement et perte
'               d'épaisseur due à la corrosion (NF EN 1993-5, tableaux 4.1 et 4.2),
'               épaisseur ou diamètre résiduel, platine, nuance, fy et fu)
' Les valeurs se saisissent dans les cases vertes. Le type de fondation (ligne 13)
' s'affiche quand la phase (C9) vaut G5 (AfficherTypeFondation) et le tableau des
' couches suit le nombre de couches (AfficherCouches) : voir CodeFeuilleDonnees.txt.
' Les listes de plus de 255 caractères (limite d'Excel, ex. noms des sols de la
' NF P11-300) sont écrites sur la feuille masquée « Listes » (voir ListeLongue).

' Constantes publiques : utilisées aussi par le module CalculFondation.
Private Const NOM_FEUILLE As String = "Données du projet"
Private Const CELLULE_PHASE As String = "C9"
Private Const LIGNE_FONDATION As Long = 13
Private Const LIGNE_DONNEES_FONDATION As Long = 31   ' sous-titre « Données de la fondation »
Public Const CELLULE_METHODE As String = "D35"      ' méthode de calcul retenue
Public Const CELLULE_NB_COUCHES As String = "C39"   ' nombre de couches saisi
Public Const LIGNE_COUCHES As Long = 41       ' première couche de sol
Public Const MAX_COUCHES As Long = 15         ' lignes prévues (au-delà du nombre : masquées)
Public Const LIGNE_PIEU As Long = 58          ' type de pieu (caractéristiques de la fondation)
Private Const ABREVIATIONS As String = "FS|FB|FTP|FTR|FSR, FBR, PU|FTC, FTCD|VM|VT|" & _
    "BPF, BPR|BE|BM|BAF|BAO|HB|HBi|PP|M1|M2|PIGU, MIGU|PIRS, MIRS"   ' tableau A.1, catégories 1 à 20
Private Const LIGNE_ACIER As Long = 92         ' type d'armature (caractéristiques de l'acier)
Private Const CLASSES_EXPOSITION As String = "X0,XC1,XC2,XC3,XC4,XD1,XD2,XD3," & _
    "XS1,XS2,XS3,XF1,XF2,XF3,XF4,XA1,XA2,XA3"
' Classe de résistance minimale du béton selon la classe d'exposition (NF EN 206/CN,
' tableau NA.F.1, valeurs à vérifier ; « - » : pas d'exigence), même ordre que
' CLASSES_EXPOSITION
Private Const CLASSES_MIN As String = "-,C20/25,C20/25,C25/30,C25/30,C25/30,C30/37,C35/45," & _
    "C30/37,C30/37,C35/45,C25/30,C25/30,C30/37,C30/37,C30/37,C35/45,C40/50"
' Armatures des micropieux : 3 types, caractéristiques mécaniques associées
' automatiquement (MajAcier), à confirmer selon le produit retenu :
'   Tube N80   : fyk = 560 MPa, E = 210 000 MPa ; « Øde x e mm » donne les dimensions
'   Barre HA   : fyk = 400 (B400) ou 500 MPa (B500), E = 200 000 MPa
'   Barre GEWI : barres filetées DSI (catalogue DSI-France du 27/10/2017) GEWI, GEWI Plus
'                et DYWIDAG, E = 205 000 MPa ; « nom|Ø nominal (mm)|section (mm²)|fyk (MPa) »
Private Const TYPES_ARMATURE As String = "Tube N80,Barre HA,Barre GEWI"
Private Const TUBES As String = "Ø60.3 x 5.5 mm,Ø73 x 5.5 mm,Ø88.9 x 6.5 mm," & _
    "Ø88.9 x 9.5 mm,Ø101.6 x 9 mm,Ø114.3 x 9 mm,Ø127 x 9 mm,Ø139.7 x 9 mm," & _
    "Ø168.3 x 10 mm,Ø177.8 x 10.4 mm"
Private Const BARRES As String = "Barre HA B400 Ø20,Barre HA B400 Ø25,Barre HA B400 Ø32," & _
    "Barre HA B400 Ø40,Barre HA B500 Ø20,Barre HA B500 Ø25,Barre HA B500 Ø32,Barre HA B500 Ø40"
Private Const GEWI As String = "GEWI Ø16|16|201|500;GEWI Ø20|20|314|500;GEWI Ø25|25|491|500;" & _
    "GEWI Ø28|28|616|500;GEWI Ø32|32|804|500;GEWI Ø40|40|1257|500;GEWI Ø50|50|1963|500;" & _
    "GEWI Ø63.5|63.5|3167|555;GEWI Ø75|75|4418|500;" & _
    "GEWI Plus Ø18|18|254|670;GEWI Plus Ø22|22|380|670;GEWI Plus Ø25|25|491|670;" & _
    "GEWI Plus Ø28|28|616|670;GEWI Plus Ø30|30|707|670;GEWI Plus Ø35|35|962|670;" & _
    "GEWI Plus Ø43|43|1452|670;GEWI Plus Ø57.5|57.5|2597|670;GEWI Plus Ø63.5|63.5|3167|670;" & _
    "GEWI Plus Ø75|75|4418|670;" & _
    "DYWIDAG Ø26.5|26.5|552|950;DYWIDAG Ø32|32|804|950;DYWIDAG Ø36|36|1018|950;" & _
    "DYWIDAG Ø40|40|1257|950;DYWIDAG Ø47.5|47.5|1735|950;DYWIDAG Ø57.5|57.5|2581|835;" & _
    "DYWIDAG Ø65|65|3331|835;DYWIDAG Ø75|75|4418|835"
' Perte d'épaisseur due à la corrosion (mm) selon l'environnement et la durée
' d'utilisation du projet (5, 25, 50, 75, 100 ans) : NF EN 1993-5, tableau 4.1 (sols)
' et tableau 4.2 (eau), valeurs recommandées ; remblais compactés : valeurs / 2
Private Const DUREES As String = "5,25,50,75,100"
Private Const ENVIRONNEMENTS As String = "Sols naturels non remaniés (sable / limon / argile / schiste)," & _
    "Sols naturels pollués et sites industriels,Sols naturels agressifs (marécages / tourbe)," & _
    "Remblais non compactés et non agressifs,Remblais non compactés et agressifs (cendres / scories)," & _
    "Eau douce (rivière / canal) - zone de fort impact,Eau douce très polluée - zone de fort impact," & _
    "Eau de mer en climat tempéré - zone de fort impact," & _
    "Eau de mer en climat tempéré - zone d'immersion permanente ou de marnage"
Private Const PERTES As String = "0,0.3,0.6,0.9,1.2;0.15,0.75,1.5,2.25,3;0.2,1,1.75,2.5,3.25;" & _
    "0.18,0.7,1.2,1.7,2.2;0.5,2,3.25,4.5,5.75;0.15,0.55,0.9,1.15,1.4;0.3,1.3,2.3,3.3,4.3;" & _
    "0.55,1.9,3.75,5.6,7.5;0.25,0.9,1.75,2.6,3.5"
' Type de platine : catalogue des barres filetées GEWI (S555/700), GEWI Plus
' (S670/800) et DYWIDAG (Y1050), diamètres à vérifier sur les fiches fabricant.
Private Const PLATINES As String = "GEWI Ø25 (S555/700),GEWI Ø28 (S555/700)," & _
    "GEWI Ø32 (S555/700),GEWI Ø40 (S555/700),GEWI Ø50 (S555/700),GEWI Ø63.5 (S555/700)," & _
    "GEWI Plus Ø18 (S670/800),GEWI Plus Ø22 (S670/800),GEWI Plus Ø25 (S670/800)," & _
    "GEWI Plus Ø28 (S670/800),GEWI Plus Ø30 (S670/800),GEWI Plus Ø35 (S670/800)," & _
    "GEWI Plus Ø43 (S670/800),GEWI Plus Ø57.5 (S670/800),GEWI Plus Ø63.5 (S670/800)," & _
    "GEWI Plus Ø75 (S670/800),DYWIDAG Ø18 (Y1050),DYWIDAG Ø26.5 (Y1050)," & _
    "DYWIDAG Ø32 (Y1050),DYWIDAG Ø36 (Y1050),DYWIDAG Ø40 (Y1050),DYWIDAG Ø47 (Y1050)," & _
    "DYWIDAG Ø57 (Y1050),DYWIDAG Ø65 (Y1050),DYWIDAG Ø75 (Y1050)"
' Nuances d'acier (NF EN 10025-2, -3 N/NL, -4 M/ML) avec limite élastique fy et
' limite de rupture fu en MPa pour une épaisseur <= 40 mm (NF EN 1993-1-1, tab. 3.1)
Private Const NUANCES As String = "S235,S275,S355,S450,S275N,S275NL,S355N,S355NL," & _
    "S420N,S420NL,S460N,S460NL,S275M,S275ML,S355M,S355ML,S420M,S420ML,S460M,S460ML"
Private Const NUANCES_FY As String = "235,275,355,440,275,275,355,355,420,420,460,460," & _
    "275,275,355,355,420,420,460,460"
Private Const NUANCES_FU As String = "360,430,510,550,390,390,490,490,520,520,540,540," & _
    "370,370,470,470,520,520,540,540"
Public Const CATEGORIES_SOL As String = "Argiles et limons," & _
    "Sols intermédiaires (sable limoneux / sable argileux / argile sableuse)," & _
    "Sables et graves,Craies,Marnes et calcaires marneux,Roches"

' Thème vert : saisies en vert, valeurs non saisies (calculées) en jaune, en-têtes et
' titres en vert (couleurs RGB écrites en nombre : rouge + 256 x vert + 65536 x bleu)
Private Const VERT_SAISIE As Long = 13561798   ' RGB(198, 239, 206)
Private Const JAUNE_CALCUL As Long = 13434879  ' RGB(255, 255, 204), jaune clair
Private Const VERT_ENTETE As Long = 14348258   ' RGB(226, 239, 218)
Private Const VERT_FONCE As Long = 2315831     ' RGB(55, 86, 35)
Public Sub CreerDonneesProjet()
    Dim ws As Worksheet, r As Long, normes As Variant, lc As Long, lm As Long, lp As Long
    Dim la As Long, c As String

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
        .Font.Color = VERT_FONCE
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
    ' Catégories d'usage (NF EN 1991-1-1) avec sous-catégories
    Liste ws.Range("C14"), "A,B,C1,C2,C3,C4,C5,D1,D2,E1,E2,F,G"
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
    ' Unité : m + référentiel, ou m/TN pour un référentiel local (cotes depuis le TN)
    ws.Range("E29").Formula = "=IF(D28=""Local"",""m/TN"",""m ""&D28)"
    ws.Range("D28:E29").HorizontalAlignment = xlCenter
    ws.Range("B28:C28").Merge
    ws.Range("B29:C29").Merge

    ' --- Données de la fondation : intitulé sur B:C fusionnées, valeur en D, unité en E ---
    lm = LIGNE_DONNEES_FONDATION             ' sous-titre
    SousTitre ws, lm, "Données de la fondation"
    ws.Range("B" & lm + 1 & ":B" & lm + 5).Value = Application.Transpose(Array( _
        "Cote de l'arase supérieure de la fondation", "Profondeur de reconnaissance", _
        "Présence de nappe phréatique", "Méthode de calcul retenue pour la fondation", _
        "Niveau de l'eau"))
    ' Unités des cotes et profondeurs : celle du niveau du terrain naturel (E29, m +
    ' référentiel altimétrique)
    ws.Range("E" & lm + 1 & ":E" & lm + 2).Formula = "=$E$29"
    ws.Range("E" & lm + 5).Value = "m"
    ws.Range("D" & lm + 1 & ":D" & lm + 2 & ",D" & lm + 5).NumberFormat = "0.00"
    ws.Range("D" & lm + 1 & ":E" & lm + 5).HorizontalAlignment = xlCenter
    Liste ws.Range("D" & lm + 3), "Non,Oui"
    Liste ws.Range("D" & lm + 4), "Méthode pressiométrique,Méthode pénétrométrique," & _
        "Essais de chargement statique"
    For r = lm + 1 To lm + 5
        ws.Range("B" & r & ":C" & r).Merge
    Next r

    ' --- Données géotechniques : nombre de couches puis tableau des couches ---
    SousTitre ws, LIGNE_COUCHES - 3, "Données géotechniques"
    ws.Range("B" & LIGNE_COUCHES - 2).Value = "Nombre de couches"
    With ws.Range(CELLULE_NB_COUCHES)
        .Value = 3
        .Validation.Delete
        .Validation.Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, "1", CStr(MAX_COUCHES)
        .Validation.ErrorMessage = "Entrer un nombre de couches entre 1 et " & MAX_COUCHES & "."
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("B" & LIGNE_COUCHES - 1 & ":K" & LIGNE_COUCHES - 1).Value = Array("N° de couche", _
        "Nom de couche (NF P11-300)", "Catégorie de sol (NF P94-262)", "Prof. bas", _
        "Épaisseur (m)", "Pression limite pl* (MPa)", "Module pressiométrique EM (MPa)", _
        "Résistance de pointe qc (MPa)", "Terrain neutralisé", "Prise comme couche d'ancrage")
    ' Unité de la profondeur : celle du niveau du terrain naturel (m + référentiel)
    ws.Range("E" & LIGNE_COUCHES - 1).Formula = "=""Prof. bas (""&$E$29&"")"""
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
        Liste ws.Range("J" & r), "Non,Oui"
        Liste ws.Range("K" & r), "Non,Oui"
    Next r
    lc = LIGNE_COUCHES + MAX_COUCHES - 1
    ' Formats des valeurs d'entrée : profondeur, épaisseur, pl*, EM et qc à 2 décimales
    ws.Range("E" & LIGNE_COUCHES & ":I" & lc).NumberFormat = "0.00"
    ws.Range("B" & LIGNE_COUCHES & ":B" & lc & ",E" & LIGNE_COUCHES & ":K" & lc).HorizontalAlignment = xlCenter

    ' --- Caractéristiques de la fondation (sous le tableau précédent) ---
    ' Intitulé (unité comprise) sur B:C fusionnées, valeur en D. Norme de référence,
    ' classe, catégorie, abréviation, résistance de pointe et type de mise en oeuvre
    ' sont des formules qui suivent le type de pieu choisi (tableau A.1 de la
    ' NF P94-262) ; le mode de fonctionnement (compression / traction) se choisit.
    ' Les hauteurs de la fondation sont calculées dans « Calcul et vérification ».
    lp = LIGNE_PIEU
    SousTitre ws, lp - 1, "Caractéristiques de la fondation"
    ws.Range("B" & lp & ":B" & lp + 19).Value = Application.Transpose(Array( _
        "Type de pieu - technique de mise en oeuvre (NF P94-262, tableau A.1)", _
        "Diamètre constant", "Incliné", "Angle d'inclinaison sur la verticale (°)", _
        "Effet de groupe", "Nombre de lignes de pieux m", "Nombre de pieux par ligne n", _
        "Entraxe des pieux d (m)", "Norme de référence", _
        "Classe de la fondation", "Catégorie de la fondation", _
        "Abréviation du type de pieu (NF P94-262, tableau A.1)", "Mode de fonctionnement", _
        "Prise en compte de la résistance de pointe", "Type de mise en oeuvre", _
        "Diamètre de la fondation B (m)", _
        "Nombre de fondations / appui du bâtiment", "Nombre total de fondations", _
        "Conditions en tête de la fondation", "Essai de contrôle"))
    ListeLongue ws.Range("D" & lp), TypesPieux(), "Pieux NF P94-262"
    ws.Range("D" & lp).Value = Split(TypesPieux(), ",")(16)   ' micropieu type I (cat. 17)
    Liste ws.Range("D" & lp + 1), "Oui,Non"
    Liste ws.Range("D" & lp + 2), "Non,Oui"
    Liste ws.Range("D" & lp + 4), "Non,Oui"
    ' Angle (si incliné) et m, n, d (si effet de groupe) : lignes affichées seulement
    ' quand le choix vaut « Oui » (AfficherOptionsPieu)
    ws.Range("D" & lp + 3).NumberFormat = "0.0"
    ws.Range("D" & lp + 5 & ":D" & lp + 6).NumberFormat = "0"
    ws.Range("D" & lp + 7).NumberFormat = "0.00"
    ' Tout ce qui suit dépend du type de pieu choisi (tableau A.1 de la NF P94-262) :
    ' la catégorie est le rang du type dans TypesPieux, le reste s'en déduit
    c = "D" & lp + 10                        ' cellule de la catégorie
    ws.Range(c).Formula = "=IFERROR(MATCH(D" & lp & ",{""" & _
        Replace(TypesPieux(), ",", """,""") & """},0),"""")"
    ' Norme de référence : 1 à 6 NF EN 1536, 7 à 16 NF EN 12699, 17 à 20 les trois
    ws.Range("D" & lp + 8).Formula = "=IF(" & c & "="""","""",IF(" & c & "<=6,""NF EN 1536""," & _
        "IF(" & c & "<=16,""NF EN 12699"",""NF EN 1536 / 14199 / 12699"")))"
    ' Classe de pieu selon la catégorie
    ws.Range("D" & lp + 9).Formula = "=IF(" & c & "="""","""",CHOOSE(" & c & _
        ",""1"",""1"",""1"",""1"",""1"",""2"",""3"",""3"",""4"",""4"",""4"",""4"",""5""," & _
        """6"",""6"",""7"",""1 bis"",""1 bis"",""8"",""8""))"
    ws.Range("D" & lp + 11).Formula = FormuleAbreviation(lp)
    ' Mode de fonctionnement de la fondation : compression ou traction
    Liste ws.Range("D" & lp + 12), "Compression,Traction"
    ' Résistance de pointe : non prise en compte pour les micropieux (catégories 17 à 20)
    ws.Range("D" & lp + 13).Formula = "=IF(" & c & "="""","""",IF(" & c & ">=17,""Non"",""Oui""))"
    ' Type de mise en oeuvre (colonnes des tableaux C.2.3 de la norme)
    ws.Range("D" & lp + 14).Formula = "=IF(" & c & "="""","""",IF(" & c & "=6," & _
        """Pieu à la tarière continue (CFA)"",IF(AND(" & c & ">=7," & c & "<=16)," & _
        """Pieu foncé"",""Pieu foré"")))"
    ws.Range("D" & lp + 15).NumberFormat = "0.000"          ' diamètre B (m)
    ws.Range("D" & lp + 16 & ":D" & lp + 17).NumberFormat = "0"
    Liste ws.Range("D" & lp + 18), "Libre en tête,Encastré en tête"
    Liste ws.Range("D" & lp + 19), "Non,Oui"
    ws.Range("D" & lp & ":D" & lp + 19).HorizontalAlignment = xlCenter
    For r = lp To lp + 19
        ws.Range("B" & r & ":C" & r).Merge
    Next r

    ' --- Hypothèses relatives aux matériaux constitutifs des fondations profondes ---
    ' Intitulé sur B:C fusionnées, valeur en D, unité en E
    la = lp + 21                             ' titre de la partie matériaux
    With ws.Range("B" & la & ":I" & la)
        .Merge
        .Value = "HYPOTHÈSES RELATIVES AUX MATÉRIAUX CONSTITUTIFS DES FONDATIONS PROFONDES"
        .Font.Bold = True
        .Font.Color = VERT_FONCE
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    SousTitre ws, la + 2, "Caractéristiques du béton"
    ws.Range("B" & la + 3 & ":B" & la + 10).Value = Application.Transpose(Array( _
        "Classe d'exposition du béton", _
        "Classe de résistance minimale (NF EN 206/CN, tableau NA.F.1)", "Classe structurale", _
        "Classe de résistance nominale du béton", "Résistance du béton à 28 jours fc28", _
        "Coefficient " & ChrW(947) & "c - combinaison ELU fondamentale", _
        "Coefficient " & ChrW(947) & "c - combinaison accidentelle", _
        "Coefficient " & ChrW(947) & "c - combinaison sismique"))
    ws.Range("E" & la + 7).Value = "MPa"
    Liste ws.Range("D" & la + 3), CLASSES_EXPOSITION
    ' Classe minimale correspondant à la classe d'exposition (CLASSES_MIN)
    ws.Range("D" & la + 4).Formula = "=IFERROR(INDEX({""" & Replace(CLASSES_MIN, ",", """,""") & _
        """},MATCH(D" & la + 3 & ",{""" & Replace(CLASSES_EXPOSITION, ",", """,""") & """},0)),"""")"
    Liste ws.Range("D" & la + 5), "S1,S2,S3,S4,S5,S6"
    ws.Range("D" & la + 5).Value = "S4"
    Liste ws.Range("D" & la + 6), "C20/25,C25/30,C30/37,C35/45,C40/50,C45/55,C50/60"
    ws.Range("D" & la + 6).Value = "C25/30"
    ' Classe nominale comparée à la classe minimale (résistance fck de la classe)
    ws.Range("E" & la + 6).Formula = "=IF(OR(D" & la + 4 & "="""",D" & la + 4 & "=""-"",D" & la + 6 & _
        "=""""),"""",IF(VALUE(MID(D" & la + 6 & ",2,FIND(""/"",D" & la + 6 & ")-2))>=VALUE(MID(D" & la + 4 & _
        ",2,FIND(""/"",D" & la + 4 & ")-2)),""Conforme"",""Inférieure à la classe minimale""))"
    ' fc28 = résistance caractéristique sur cylindre de la classe (C25/30 -> 25 MPa)
    ws.Range("D" & la + 7).Formula = "=VALUE(MID(D" & la + 6 & ",2,FIND(""/"",D" & la + 6 & ")-2))"
    ws.Range("D" & la + 8 & ":D" & la + 10).Value = Application.Transpose(Array(1.5, 1.2, 1.3))
    ws.Range("D" & la + 7 & ":D" & la + 10).NumberFormat = "0.00"

    ' Acier : la liste des désignations dépend du type d'armature, les dimensions et
    ' caractéristiques mécaniques sont remplies (voir MajArmature et MajAcier) ; perte
    ' d'épaisseur due à la corrosion selon l'environnement et la durée d'utilisation
    SousTitre ws, LIGNE_ACIER - 1, "Caractéristiques de l'acier"
    ws.Range("B" & LIGNE_ACIER & ":B" & LIGNE_ACIER + 18).Value = Application.Transpose(Array( _
        "Type d'armature", "Désignation de l'armature", "Diamètre extérieur (tube) ou nominal (barre)", _
        "Épaisseur e", "Diamètre intérieur", "Section d'acier As", "Module de Young E", _
        "Limite d'élasticité fyk", _
        "Coefficient " & ChrW(947) & "s - combinaison ELU fondamentale", _
        "Coefficient " & ChrW(947) & "s - combinaison accidentelle", _
        "Coefficient " & ChrW(947) & "s - combinaison sismique", _
        "Durée d'utilisation du projet", _
        "Environnement de l'armature (NF EN 1993-5, tableaux 4.1 et 4.2)", _
        "Perte d'épaisseur due à la corrosion", _
        "Épaisseur résiduelle du tube (e - perte) ou diamètre résiduel de la barre (d - 2 x perte)", _
        "Type de platine (liaison fondation / structure)", "Nuance de l'acier de la platine", _
        "Limite élastique de l'acier fy", "Limite de rupture de l'acier fu"))
    ws.Range("E" & LIGNE_ACIER + 2 & ":E" & LIGNE_ACIER + 18).Value = Application.Transpose(Array( _
        "mm", "mm", "mm", "mm²", "MPa", "MPa", "", "", "", "ans", "", "mm", "mm", "", "", "MPa", "MPa"))
    Liste ws.Range("D" & LIGNE_ACIER), TYPES_ARMATURE
    ws.Range("D" & LIGNE_ACIER + 2 & ":D" & LIGNE_ACIER + 4).NumberFormat = "0.0"
    ws.Range("D" & LIGNE_ACIER + 5 & ":D" & LIGNE_ACIER + 7).NumberFormat = "0"
    ws.Range("D" & LIGNE_ACIER + 8 & ":D" & LIGNE_ACIER + 10).Value = Application.Transpose(Array(1.15, 1, 1))
    ws.Range("D" & LIGNE_ACIER + 8 & ":D" & LIGNE_ACIER + 10).NumberFormat = "0.00"
    Liste ws.Range("D" & LIGNE_ACIER + 11), DUREES
    ws.Range("D" & LIGNE_ACIER + 11).Value = 50
    ListeLongue ws.Range("D" & LIGNE_ACIER + 12), ENVIRONNEMENTS, "Environnements NF EN 1993-5"
    ' Perte d'épaisseur : tableau PERTES (environnement en ligne, durée en colonne)
    ws.Range("D" & LIGNE_ACIER + 13).Formula = "=IFERROR(INDEX({" & PERTES & "},MATCH(D" & _
        LIGNE_ACIER + 12 & ",{""" & Replace(ENVIRONNEMENTS, ",", """,""") & """},0),MATCH(D" & _
        LIGNE_ACIER + 11 & ",{" & DUREES & "},0)),"""")"
    ws.Range("D" & LIGNE_ACIER + 14).Formula = "=IF(OR(D" & LIGNE_ACIER + 13 & "="""",D" & _
        LIGNE_ACIER + 2 & "=""""),"""",IF(D" & LIGNE_ACIER & "=""Tube N80"",D" & LIGNE_ACIER + 3 & _
        "-D" & LIGNE_ACIER + 13 & ",D" & LIGNE_ACIER + 2 & "-2*D" & LIGNE_ACIER + 13 & "))"
    ws.Range("D" & LIGNE_ACIER + 13 & ":D" & LIGNE_ACIER + 14).NumberFormat = "0.00"
    ListeLongue ws.Range("D" & LIGNE_ACIER + 15), PLATINES, "Platines"
    Liste ws.Range("D" & LIGNE_ACIER + 16), NUANCES
    ws.Range("D" & LIGNE_ACIER + 16).Value = "S355"
    ' fy et fu de la nuance choisie (tables NUANCES_FY et NUANCES_FU)
    ws.Range("D" & LIGNE_ACIER + 17).Formula = FormuleNuance(NUANCES_FY, LIGNE_ACIER + 16)
    ws.Range("D" & LIGNE_ACIER + 18).Formula = FormuleNuance(NUANCES_FU, LIGNE_ACIER + 16)
    ws.Range("D" & la + 3 & ":E" & LIGNE_ACIER + 18).HorizontalAlignment = xlCenter
    For r = la + 3 To LIGNE_ACIER + 18
        If r <> LIGNE_ACIER - 1 And r <> LIGNE_ACIER - 2 Then ws.Range("B" & r & ":C" & r).Merge
    Next r

    ' --- Mise en forme ---
    With ws.Range("B4:C9,B12:C16,B19:I26,B28:E29,B" & LIGNE_COUCHES - 2 & ":C" & LIGNE_COUCHES - 2 & _
            ",B" & LIGNE_COUCHES - 1 & ":K" & lc & _
            ",B" & lm + 1 & ":E" & lm + 5 & ",B" & lp & ":D" & lp + 19 & _
            ",B" & la + 3 & ":E" & la + 10 & ",B" & LIGNE_ACIER & ":E" & LIGNE_ACIER + 18)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("B19:I19,B" & LIGNE_COUCHES - 1 & ":K" & LIGNE_COUCHES - 1).Font.Bold = True
    ws.Range("B19:I19,B" & LIGNE_COUCHES - 1 & ":K" & LIGNE_COUCHES - 1).Interior.Color = VERT_ENTETE
    ws.Range("B" & LIGNE_COUCHES - 1 & ":K" & LIGNE_COUCHES - 1).HorizontalAlignment = xlCenter
    ' Cases à remplir en vert, valeurs calculées en jaune (caractéristiques de la
    ' fondation comprises) ; sans couleur : données géotechniques (laboratoire) et
    ' données de la fondation
    ws.Range("D" & lp & ":D" & lp + 7 & ",D" & lp + 12 & ",D" & lp + 15 & ":D" & lp + 19 _
        ).Interior.Color = VERT_SAISIE
    ws.Range("D" & lp + 8 & ":D" & lp + 11 & ",D" & lp + 13 & ":D" & lp + 14).Interior.Color = JAUNE_CALCUL
    ws.Range("C4:C9,C12:C16,D28:D29," & CELLULE_NB_COUCHES & ",D" & la + 3 & ",D" & la + 5 & ":D" & _
        la + 6 & ",D" & la + 8 & ":D" & la + 10 & _
        ",D" & LIGNE_ACIER & ":D" & LIGNE_ACIER + 1 & ",D" & LIGNE_ACIER + 8 & ":D" & _
        LIGNE_ACIER + 12 & ",D" & LIGNE_ACIER + 15 & ":D" & LIGNE_ACIER + 16).Interior.Color = VERT_SAISIE
    ws.Range("D" & la + 4 & ",D" & la + 7 & ",D" & LIGNE_ACIER + 2 & ":D" & LIGNE_ACIER + 7 & _
        ",D" & LIGNE_ACIER + 13 & ":D" & LIGNE_ACIER + 14 & _
        ",D" & LIGNE_ACIER + 17 & ":D" & LIGNE_ACIER + 18).Interior.Color = JAUNE_CALCUL
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 40
    ws.Columns("D").ColumnWidth = 28         ' valeurs de la fondation, catégorie de sol
    ws.Columns("E:K").ColumnWidth = 14
    ws.Rows("19:" & LIGNE_ACIER + 18).AutoFit

    AfficherTypeFondation ws
    AfficherCouches ws
    AfficherMethode ws
    AfficherOptionsPieu ws
    MajArmature ws
    ws.Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Appelée par le code de la feuille « Données du projet » (CodeFeuilleDonnees.txt)
' à chaque modification : les adresses des cellules suivies sont tenues ici, d'après
' les constantes du module, pour ne plus avoir à modifier le code de la feuille quand
' la disposition change.
'   phase (C9) -> type de fondation ; nombre de couches -> lignes de couches ;
'   type d'armature -> liste des classes d'acier ; classe de l'acier -> dimensions
'   du tube ; incliné / effet de groupe -> angle / m, n, d affichés ; type de pieu ->
'   résistance de pointe affichée ou non dans « Calcul et vérification » ; méthode de
'   calcul -> colonne pl* ou qc masquée et « Calcul et vérification » recréée
Public Sub ChangementDonnees(ws As Worksheet, Target As Range)
    On Error GoTo Fin
    Application.EnableEvents = False
    If Not Intersect(Target, ws.Range(CELLULE_PHASE)) Is Nothing Then AfficherTypeFondation ws
    If Not Intersect(Target, ws.Range(CELLULE_NB_COUCHES)) Is Nothing Then AfficherCouches ws
    If Not Intersect(Target, ws.Range("D" & LIGNE_ACIER)) Is Nothing Then MajArmature ws
    If Not Intersect(Target, ws.Range("D" & LIGNE_ACIER + 1)) Is Nothing Then MajAcier ws
    If Not Intersect(Target, Union(ws.Range("D" & LIGNE_PIEU + 2), ws.Range("D" & LIGNE_PIEU + 4))) _
        Is Nothing Then AfficherOptionsPieu ws
    ' Effet de groupe : bloc Ce de « Calcul et vérification » affiché ou masqué
    If Not Intersect(Target, ws.Range("D" & LIGNE_PIEU + 4)) Is Nothing Then
        On Error Resume Next
        Application.Run "AfficherGroupe"
        On Error GoTo Fin
    End If
    ' Type de pieu : colonnes de la résistance de pointe de « Calcul et vérification »
    ' affichées ou masquées (module CalculFondation, s'il est présent)
    If Not Intersect(Target, ws.Range("D" & LIGNE_PIEU)) Is Nothing Then
        ws.Calculate
        On Error Resume Next
        Application.Run "AfficherPointe"
    End If
    ' Méthode de calcul : colonne pl* ou qc des données géotechniques masquée, puis
    ' « Calcul et vérification » recréée avec les formules de la méthode choisie (pl* et
    ' tableaux F, ou qc et tableaux G)
    If Not Intersect(Target, ws.Range(CELLULE_METHODE)) Is Nothing Then
        AfficherMethode ws
        On Error Resume Next
        Application.Run "ActualiserCalculVerification"
    End If
Fin:
    Application.EnableEvents = True
End Sub

' Lignes de l'angle d'inclinaison (si « Incliné » vaut Oui) et du nombre de lignes m,
' du nombre de pieux par ligne n et de l'entraxe d (si « Effet de groupe » vaut Oui)
Public Sub AfficherOptionsPieu(ws As Worksheet)
    ws.Rows(LIGNE_PIEU + 3).Hidden = (ws.Range("D" & LIGNE_PIEU + 2).Value <> "Oui")
    ws.Rows(LIGNE_PIEU + 5).Resize(3).Hidden = (ws.Range("D" & LIGNE_PIEU + 4).Value <> "Oui")
End Sub

' Données géotechniques selon la méthode de calcul retenue : colonne de la pression
' limite pl* (G) masquée en méthode pénétrométrique, colonne de la résistance de
' pointe qc (I) masquée en méthode pressiométrique ; les deux affichées sinon.
Public Sub AfficherMethode(ws As Worksheet)
    Dim m As String
    m = CStr(ws.Range(CELLULE_METHODE).Value)
    ws.Columns("G").Hidden = (m = "Méthode pénétrométrique")
    ws.Columns("I").Hidden = (m = "Méthode pressiométrique")
End Sub

' Affiche autant de lignes de couches que le nombre saisi en C39 (1 à MAX_COUCHES),
' les autres lignes du tableau sont masquées. Appelée aussi quand C39 change
' (voir CodeFeuilleDonnees.txt).
Public Sub AfficherCouches(ws As Worksheet)
    Dim nb As Long
    nb = Val(ws.Range(CELLULE_NB_COUCHES).Value)
    If nb < 1 Then nb = 1
    If nb > MAX_COUCHES Then nb = MAX_COUCHES
    ws.Rows(LIGNE_COUCHES).Resize(MAX_COUCHES).Hidden = False
    If nb < MAX_COUCHES Then ws.Rows(LIGNE_COUCHES + nb).Resize(MAX_COUCHES - nb).Hidden = True
End Sub

' Formule qui donne la valeur de la table valeurs (même ordre que NUANCES) pour
' la nuance choisie en D de la ligne indiquée.
Private Function FormuleNuance(valeurs As String, ligne As Long) As String
    FormuleNuance = "=INDEX({" & valeurs & "},MATCH(D" & ligne & ",{""" & _
        Replace(NUANCES, ",", """,""") & """},0))"
End Function

' Type d'armature choisi (D92) : la liste « Désignation de l'armature » propose les
' tubes N80 (TUBES), les barres HA (BARRES) ou les barres GEWI, GEWI Plus et DYWIDAG
' (GEWI) ; épaisseur et diamètre intérieur ne sont affichés que pour un tube.
' Appelée aussi quand le type change (voir CodeFeuilleDonnees.txt).
Public Sub MajArmature(ws As Worksheet)
    Dim typ As String, liste As String, t As Variant, k As Long
    typ = CStr(ws.Range("D" & LIGNE_ACIER).Value)
    Select Case typ
        Case "Tube N80": liste = TUBES
        Case "Barre HA": liste = BARRES
        Case Else
            t = Split(GEWI, ";")
            For k = 0 To UBound(t)
                liste = liste & IIf(k > 0, ",", "") & Split(t(k), "|")(0)
            Next k
    End Select
    ListeLongue ws.Range("D" & LIGNE_ACIER + 1), liste, "Armatures " & typ
    ws.Rows(LIGNE_ACIER + 3).Resize(2).Hidden = (typ <> "Tube N80")
    MajAcier ws
End Sub

' Désignation de l'armature choisie (D93) : dimensions, section As, module E et limite
' d'élasticité fyk associés au type d'armature :
'   Tube N80   : « Øde x e mm » -> de et e, di = de - 2 e, As = pi (de² - di²) / 4,
'                E = 210 000 MPa, fyk = 560 MPa
'   Barre HA   : Ø et nuance (B400 / B500) lus dans la désignation, As = pi Ø² / 4,
'                E = 200 000 MPa, fyk = 400 ou 500 MPa
'   Barre GEWI : Ø, As et fyk du catalogue (GEWI), E = 205 000 MPa
' Appelée aussi quand la désignation change (voir CodeFeuilleDonnees.txt).
Public Sub MajAcier(ws As Worksheet)
    Dim typ As String, nom As String, p As Long, t As Variant, k As Long, a As Long
    a = LIGNE_ACIER
    typ = CStr(ws.Range("D" & a).Value)
    nom = CStr(ws.Range("D" & a + 1).Value)
    p = InStr(nom, "Ø")
    ws.Range("D" & a + 2 & ":D" & a + 7).ClearContents
    If p = 0 Then Exit Sub
    Select Case typ
        Case "Tube N80"
            ws.Range("D" & a + 2).Value = Val(Mid(nom, p + 1))
            ws.Range("D" & a + 3).Value = Val(Mid(nom, InStr(p, nom, "x") + 1))
            ws.Range("D" & a + 4).Formula = "=D" & a + 2 & "-2*D" & a + 3
            ws.Range("D" & a + 5).Formula = "=PI()*(D" & a + 2 & "^2-D" & a + 4 & "^2)/4"
            ws.Range("D" & a + 6).Value = 210000
            ws.Range("D" & a + 7).Value = 560
        Case "Barre HA"
            ws.Range("D" & a + 2).Value = Val(Mid(nom, p + 1))
            ws.Range("D" & a + 5).Formula = "=PI()*D" & a + 2 & "^2/4"
            ws.Range("D" & a + 6).Value = 200000
            ws.Range("D" & a + 7).Value = IIf(InStr(nom, "B400") > 0, 400, 500)
        Case Else
            t = Split(GEWI, ";")
            For k = 0 To UBound(t)
                If Split(t(k), "|")(0) = nom Then
                    ws.Range("D" & a + 2).Value = Val(Split(t(k), "|")(1))
                    ws.Range("D" & a + 5).Value = Val(Split(t(k), "|")(2))
                    ws.Range("D" & a + 6).Value = 205000
                    ws.Range("D" & a + 7).Value = Val(Split(t(k), "|")(3))
                    Exit For
                End If
            Next k
    End Select
End Sub

' Formule qui donne l'abréviation (ABREVIATIONS) du type de pieu choisi en D de la
' ligne lp, d'après son rang dans la liste TypesPieux ; vide si aucun type.
Private Function FormuleAbreviation(lp As Long) As String
    FormuleAbreviation = "=IFERROR(INDEX({""" & Replace(ABREVIATIONS, "|", """,""") & _
        """},MATCH(D" & lp & ",{""" & Replace(TypesPieux(), ",", """,""") & """},0)),"""")"
End Function

' Techniques de mise en oeuvre des pieux (NF P94-262, tableau A.1), dans l'ordre
' des catégories 1 à 20. Désignations abrégées, à vérifier sur la norme.
Private Function TypesPieux() As String
    Dim s As String
    s = "Foré simple (pieux et barrettes)"
    s = s & ",Foré boue (pieux et barrettes)"
    s = s & ",Foré tubé - virole perdue"
    s = s & ",Foré tubé - virole récupérée"
    s = s & ",Foré simple ou boue avec rainurage ou puits"
    s = s & ",Foré tarière creuse simple ou double rotation"
    s = s & ",Vissé moulé"
    s = s & ",Vissé tubé"
    s = s & ",Battu béton préfabriqué ou précontraint"
    s = s & ",Battu enrobé - béton / mortier / coulis"
    s = s & ",Battu moulé"
    s = s & ",Battu acier fermé"
    s = s & ",Battu acier ouvert"
    s = s & ",Profilé H battu"
    s = s & ",Profilé H battu injecté"
    s = s & ",Palplanches battues"
    s = s & ",Micropieu type I"
    s = s & ",Micropieu type II"
    s = s & ",Pieu ou micropieu injecté mode IGU - type III"
    s = s & ",Pieu ou micropieu injecté mode IRS - type IV"
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
        .Font.Color = VERT_FONCE
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
