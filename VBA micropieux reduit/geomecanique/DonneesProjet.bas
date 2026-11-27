Attribute VB_Name = "DonneesProjet"
Option Explicit
Private Const COUL_TITRE As Long = 5898259
Private Const COUL_ENTETE As Long = 14337945
Private Const COUL_SAISIE As Long = 15201485
Private Const COUL_CALCUL As Long = 15590349
Private Const NOM_FEUILLE As String = "Données du projet"
Private Const CELLULE_PHASE As String = "C9"
Private Const LIGNE_FONDATION As Long = 13
Private Const LIGNE_DONNEES_FONDATION As Long = 31
Public Const CELLULE_METHODE As String = "D35"
Public Const LIGNE_COUCHE_MOLLE As Long = 37
Public Const CELLULE_NB_COUCHES As String = "C46"
Public Const LIGNE_COUCHES As Long = 48
Public Const MAX_COUCHES As Long = 15
Public Const LIGNE_PIEU As Long = 65
Public Const NB_DIAMETRES As Long = 4
Private Const ABREVIATIONS As String = "M1|M2|PIGU, MIGU|PIRS, MIRS"
Public Const LIGNE_ACIER As Long = 102
Private Const TYPES_ARMATURE As String = "Tube N80,Barre HA,Barre GEWI"
Private Const TUBES As String = "Ø60.3 x 5.5 mm,Ø73 x 5.5 mm,Ø88.9 x 6.5 mm," & "Ø88.9 x 9.5 mm,Ø101.6 x 9 mm,Ø114.3 x 9 mm,Ø127 x 9 mm,Ø139.7 x 9 mm," & "Ø168.3 x 10 mm,Ø177.8 x 10.4 mm"
Private Const BARRES As String = "Barre HA B400 Ø20,Barre HA B400 Ø25,Barre HA B400 Ø32," & "Barre HA B400 Ø40,Barre HA B500 Ø20,Barre HA B500 Ø25,Barre HA B500 Ø32,Barre HA B500 Ø40"
Private Const GEWI As String = "GEWI Ø16|16|201|500;GEWI Ø20|20|314|500;GEWI Ø25|25|491|500;" & "GEWI Ø28|28|616|500;GEWI Ø32|32|804|500;GEWI Ø40|40|1257|500;GEWI Ø50|50|1963|500;" & "GEWI Ø63.5|63.5|3167|555;GEWI Ø75|75|4418|500;" & "GEWI Plus Ø18|18|254|670;GEWI Plus Ø22|22|380|670;GEWI Plus Ø25|25|491|670;" & "GEWI Plus Ø28|28|616|670;GEWI Plus Ø30|30|707|670;GEWI Plus Ø35|35|962|670;" & "GEWI Plus Ø43|43|1452|670;GEWI Plus Ø57.5|57.5|2597|670;GEWI Plus Ø63.5|63.5|3167|670;" & "GEWI Plus Ø75|75|4418|670;" & "DYWIDAG Ø26.5|26.5|552|950;DYWIDAG Ø32|32|804|950;DYWIDAG Ø36|36|1018|950;" & "DYWIDAG Ø40|40|1257|950;DYWIDAG Ø47.5|47.5|1735|950;DYWIDAG Ø57.5|57.5|2581|835;" & "DYWIDAG Ø65|65|3331|835;DYWIDAG Ø75|75|4418|835"
Private Const DUREES As String = "5,25,50,75,100"
Private Const ENVIRONNEMENTS As String = "Sols naturels non remaniés (sable / limon / argile / schiste)," & "Sols naturels pollués et sites industriels,Sols naturels agressifs (marécages / tourbe)," & "Remblais non compactés et non agressifs,Remblais non compactés et agressifs (cendres / scories)," & "Eau douce (rivière / canal) - zone de fort impact,Eau douce très polluée - zone de fort impact," & "Eau de mer en climat tempéré - zone de fort impact," & "Eau de mer en climat tempéré - zone d'immersion permanente ou de marnage," & "Remblais compactés et non agressifs,Remblais compactés et agressifs (cendres / scories)"
Private Const PERTES As String = "0,0.3,0.6,0.9,1.2;0.15,0.75,1.5,2.25,3;0.2,1,1.75,2.5,3.25;" & "0.18,0.7,1.2,1.7,2.2;0.5,2,3.25,4.5,5.75;0.15,0.55,0.9,1.15,1.4;0.3,1.3,2.3,3.3,4.3;" & "0.55,1.9,3.75,5.6,7.5;0.25,0.9,1.75,2.6,3.5;0.09,0.35,0.6,0.85,1.1;0.25,1,1.63,2.25,2.88"
Private Const RHO_CIMENT As Double = 3.1
Public Const PLATINES As String = "Platine carrée 150 x 150 mm - ép. 15 mm," & "Platine carrée 150 x 150 mm - ép. 20 mm,Platine carrée 200 x 200 mm - ép. 20 mm," & "Platine carrée 200 x 200 mm - ép. 25 mm,Platine carrée 250 x 250 mm - ép. 25 mm," & "Platine carrée 250 x 250 mm - ép. 30 mm,Platine carrée 300 x 300 mm - ép. 30 mm," & "Platine carrée 300 x 300 mm - ép. 40 mm,Platine carrée 350 x 350 mm - ép. 40 mm," & "Platine carrée 350 x 350 mm - ép. 50 mm,Platine carrée 400 x 400 mm - ép. 50 mm," & "Platine carrée 400 x 400 mm - ép. 60 mm,Platine carrée 500 x 500 mm - ép. 80 mm"
Private Const NUANCES As String = "S235,S275,S355,S450,S275N,S275NL,S355N,S355NL," & "S420N,S420NL,S460N,S460NL,S275M,S275ML,S355M,S355ML,S420M,S420ML,S460M,S460ML"
Private Const NUANCES_FY As String = "235,275,355,440,275,275,355,355,420,420,460,460," & "275,275,355,355,420,420,460,460"
Private Const NUANCES_FU As String = "360,430,510,550,390,390,490,490,520,520,540,540," & "370,370,470,470,520,520,540,540"
Private Const NUANCES_FY80 As String = "215,255,335,410,255,255,335,335,390,390,430,430," & "255,255,335,335,390,390,430,430"
Private Const NUANCES_FU80 As String = "360,410,470,550,370,370,470,470,520,520,540,540," & "360,360,450,450,500,500,530,530"
Public Const TYPES_SOLLICITATION As String = "a - Sollicitations permanentes en tête dominantes," & "b - Poussées latérales du sol dominantes,c - Sollicitations de courte durée en tête dominantes," & "d - Sollicitations accidentelles très brèves en tête dominantes"
Private Const CATEGORIES_RHEO As String = "I - surconsolidé ou très serré," & "II - normalement consolidé ou normalement serré,III - sous-consolidé altéré et remanié ou lâche"
Private Const NATURES_SOL As String = "Tourbe,Argile,Limon,Sable,Grave,Rocher très peu fracturé," & "Rocher normal,Rocher très fracturé,Rocher très altéré"
Public Const CATEGORIES_SOL As String = "Argiles et limons," & "Sols intermédiaires (sable limoneux / sable argileux / argile sableuse)," & "Sables et graves,Craies,Marnes et calcaires marneux,Roches"
Public Sub CreerDonneesProjet()
    Dim ws As Worksheet, r As Long, normes As Variant, lc As Long, lm As Long, lp As Long
    Dim la As Long, c As String
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(Before:=ActiveWorkbook.Worksheets(1))
        ws.Name = NOM_FEUILLE
    End If
    ws.Cells.Clear
    ws.Rows.Hidden = False
    ws.Rows("1:2").RowHeight = 30
    With ws.Range("H1:I1")
        .Merge
        .Value = "Numéro du projet"
        .Font.Italic = True
        .Font.Size = 9
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
    End With
    With ws.Range("H2:I2")
        .Merge
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = COUL_SAISIE
        .BorderAround xlContinuous, xlMedium
    End With
    With ws.Range("B3:I3")
        .Merge
        .Value = "HYPOTHÈSES GÉNÉRALES"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 14
        .Interior.Color = COUL_TITRE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(3).RowHeight = 24
    ws.Range("G4").Value = "Date :"
    ws.Range("G4").HorizontalAlignment = xlRight
    With ws.Range("H4:I4")
        .Merge
        .Formula = "=TODAY()"
        .NumberFormat = "dd/mm/yyyy"
        .HorizontalAlignment = xlCenter
        .Interior.Color = COUL_SAISIE
        .Borders(xlEdgeBottom).LineStyle = xlDot
    End With
    ws.Rows(4).RowHeight = 22
    With ws.Range("B5:I5")
        .Merge
        .Value = "SYNTHÈSE DU PROJET"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = COUL_TITRE
        .Interior.Color = COUL_ENTETE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .BorderAround xlContinuous, xlThin
    End With
    ws.Rows(5).RowHeight = 22
    ws.Range("B6:B9").Value = Application.Transpose(Array("Nom du projet :", "Lieu :", "Maître d'ouvrage :", "Phase du projet :"))
    For r = 6 To 9
        With ws.Range("C" & r & ":F" & r)
            .Merge
            .Interior.Color = COUL_SAISIE
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Borders(xlEdgeBottom).LineStyle = xlDot
        End With
    Next r
    ws.Rows("6:9").RowHeight = 22
    Liste ws.Range(CELLULE_PHASE), "G2-AVP,G2-PRO,G5"
    ws.Range(CELLULE_PHASE).HorizontalAlignment = xlLeft
    SousTitre ws, 11, "Caractérisation générale de l'ouvrage"
    ws.Range("B12:B16").Value = Application.Transpose(Array("Type d'ouvrage", "Type de fondation", "Catégorie de bâtiment", "Altitude", "Zone de sismicité"))
    Liste ws.Range("C12"), "Bâtiment,Pont"
    Liste ws.Range("C13"), "Micropieux"
    ws.Range("C13").Value = "Micropieux"
    Liste ws.Range("C14"), "A,B,C1,C2,C3,C4,C5,D1,D2,E1,E2,F,G"
    Liste ws.Range("C15"), "<=1000 m,>1000 m"
    Liste ws.Range("C16"), "1,2,3,4,5"
    SousTitre ws, 18, "Références normatives"
    ws.Range("B19:D19").Value = Array("Référence", "Date", "Intitulé")
    normes = Array( Array("NF EN 1990", "mars 2003", "Eurocode 0 - Bases de calcul des structures"), Array("NF EN 1992-1-1", "octobre 2005", "Eurocode 2 - Calcul des structures en béton - Partie 1-1 : règles générales et règles pour les bâtiments"), Array("NF P22-311-1", "octobre 2005", "Eurocode 3 (NF EN 1993-1-1) - Calcul des structures en acier - Partie 1-1 : règles générales et règles pour les bâtiments"), Array("NF P94-251-1", "juin 2005", "Eurocode 7 (NF EN 1997-1) - Calcul géotechnique - Partie 1 : règles générales"), Array("NF P94-262", "juillet 2012", "Justification des ouvrages géotechniques - Norme d'application nationale de l'Eurocode 7 - Fondations profondes"), Array("NF EN 1998-1", "septembre 2005", "Eurocode 8 - Calcul des structures pour leur résistance aux séismes - Partie 1 : règles générales, actions sismiques et règles pour les bâtiments"), Array("NF EN 1998-5", "septembre 2005", _
            "Eurocode 8 - Partie 5 : fondations, ouvrages de soutènement et aspects géotechniques"))
    For r = 0 To UBound(normes)
        ws.Range("B" & 20 + r & ":D" & 20 + r).Value = normes(r)
    Next r
    For r = 19 To 26
        ws.Range("D" & r & ":I" & r).Merge
    Next r
    ws.Range("B28").Value = "Référentiel altimétrique"
    ws.Range("B29").Value = "Niveau du terrain naturel"
    Liste ws.Range("D28"), "NGF-IGN69,NGF-IGN78,NVP,IGN,Local"
    ws.Range("D29").NumberFormat = "0.00"
    ws.Range("E29").Formula = "=IF(D28=""Local"",""m/TN"",""m ""&D28)"
    ws.Range("D28:E29").HorizontalAlignment = xlCenter
    ws.Range("B28:C28").Merge
    ws.Range("B29:C29").Merge
    lm = LIGNE_DONNEES_FONDATION
    SousTitre ws, lm, "Données de la fondation"
    ws.Range("B" & lm + 1 & ":B" & lm + 5).Value = Application.Transpose(Array( "Cote de l'arase supérieure de la fondation", "Profondeur de reconnaissance", "Présence de nappe phréatique", "Méthode de calcul retenue pour la fondation", "Niveau de l'eau"))
    ws.Range("E" & lm + 1 & ":E" & lm + 2).Formula = "=$E$29"
    ws.Range("E" & lm + 5).Value = "m"
    ws.Range("D" & lm + 1 & ":D" & lm + 2 & ",D" & lm + 5).NumberFormat = "0.00"
    ws.Range("D" & lm + 1 & ":E" & lm + 5).HorizontalAlignment = xlCenter
    Liste ws.Range("D" & lm + 3), "Non,Oui"
    Liste ws.Range("D" & lm + 4), "Méthode pressiométrique,Méthode pénétrométrique," & "Essais de chargement statique"
    r = LIGNE_COUCHE_MOLLE
    ws.Range("B" & r & ":B" & r + 6).Value = Application.Transpose(Array( "Frottement négligé : épaisseur des terrains neutralisés", "Couche molle (terrain neutralisé) : module", "Couche molle : épaisseur D", "Sol rocheux", "Catégorie du sol", "Nature du sol", "Coefficient rhéologique " & ChrW(945)))
    ws.Range("B" & r + 1).Formula = "=IF(" & CELLULE_METHODE & "=""Méthode pénétrométrique""," & """Couche molle (terrain neutralisé) : résistance de pointe qc""," & """Couche molle (terrain neutralisé) : module pressiométrique EM"")"
    ws.Range("E" & r & ":E" & r + 6).Value = Application.Transpose(Array("m", "MPa", "m", "", "", "", ""))
    ws.Range("D" & r).Formula = "=SUMIF($M$" & LIGNE_COUCHES & ":$M$" & LIGNE_COUCHES + MAX_COUCHES - 1 & ",""Oui"",$F$" & LIGNE_COUCHES & ":$F$" & LIGNE_COUCHES + MAX_COUCHES - 1 & ")"
    ws.Range("D" & r + 1).Formula = "=IFERROR(INDEX(IF(" & CELLULE_METHODE & "=""Méthode pénétrométrique""," & "$I$" & LIGNE_COUCHES & ":$I$" & LIGNE_COUCHES + MAX_COUCHES - 1 & ",$H$" & LIGNE_COUCHES & ":$H$" & LIGNE_COUCHES + MAX_COUCHES - 1 & "),MATCH(""Oui"",$M$" & LIGNE_COUCHES & ":$M$" & LIGNE_COUCHES + MAX_COUCHES - 1 & ",0)),"""")"
    ws.Range("D" & r + 2).Formula = "=D" & r
    Liste ws.Range("D" & r + 3), "Non,Oui"
    Liste ws.Range("D" & r + 4), CATEGORIES_RHEO
    ws.Range("D" & r + 4).Value = Split(CATEGORIES_RHEO, ",")(1)
    Liste ws.Range("D" & r + 5), NATURES_SOL
    ws.Range("D" & r + 5).Value = "Argile"
    ws.Range("D" & r + 6).Formula = "=IFERROR(IF(OR(D" & r + 3 & "=""Oui"",LEFT(D" & r + 5 & ",6)=""Rocher"")," & "INDEX(NF_RheoRoc_Alpha,MATCH(D" & r + 5 & ",NF_RheoRoc_Nature,0)),INDEX(CHOOSE(MATCH(D" & r + 4 & ",{""" & Replace(CATEGORIES_RHEO, ",", """,""") & """},0),NF_Rheo_Alpha1,NF_Rheo_Alpha2,NF_Rheo_Alpha3)," & "MATCH(D" & r + 5 & ",NF_Rheo_Nature,0))),IFERROR(IF(OR(D" & r + 3 & "=""Oui"",LEFT(D" & r + 5 & ",6)=""Rocher""),CHOOSE(MATCH(D" & r + 5 & ",{""Rocher très peu fracturé"",""Rocher normal""," & """Rocher très fracturé"",""Rocher très altéré""},0),2/3,1/2,1/3,2/3)," & "INDEX({1,1,1;1,0.6667,0.5;0.6667,0.5,0.5;0.5,0.3333,0.3333;0.3333,0.25,0.25},MATCH(D" & r + 5 & ",{""Tourbe"",""Argile"",""Limon"",""Sable"",""Grave""},0),MATCH(D" & r + 4 & ",{""" & Replace(CATEGORIES_RHEO, ",", """,""") & """},0))),""""))"
    ws.Range("D" & r & ":D" & r + 2 & ",D" & r + 6).NumberFormat = "0.00"
    ws.Range("D" & r & ":E" & r + 6).HorizontalAlignment = xlCenter
    ws.Range("D" & r & ":D" & r + 2 & ",D" & r + 6).Interior.Color = COUL_CALCUL
    ws.Range("D" & r + 3 & ":D" & r + 5).Interior.Color = COUL_SAISIE
    For r = lm + 1 To LIGNE_COUCHE_MOLLE + 6
        ws.Range("B" & r & ":C" & r).Merge
    Next r
    SousTitre ws, LIGNE_COUCHES - 3, "Données géotechniques"
    ws.Range("B" & LIGNE_COUCHES - 2).Value = "Nombre de couches"
    With ws.Range(CELLULE_NB_COUCHES)
        .Value = 3
        .Validation.Delete
        .Validation.Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, "1", CStr(MAX_COUCHES)
        .Validation.ErrorMessage = "Entrer un nombre de couches entre 1 et " & MAX_COUCHES & "."
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("B" & LIGNE_COUCHES - 1 & ":N" & LIGNE_COUCHES - 1).Value = Array("N° de couche", "Nom de couche", "Catégorie de sol (NF P94-262)", "Prof. bas", "Épaisseur (m)", "Pression limite pl* (MPa)", "Module pressiométrique EM (MPa)", "Résistance de pointe qc (MPa)", "Poids volumique " & ChrW(947) & " (kN/m3)", "Cohésion c' (kPa)", "Angle de frottement " & ChrW(966) & "' (°)", "Terrain neutralisé", "Prise comme couche d'ancrage")
    ws.Range("E" & LIGNE_COUCHES - 1).Formula = "=""Prof. bas (""&$E$29&"")"""
    ws.Range("H" & LIGNE_COUCHES - 1).Formula = "=IF(" & CELLULE_METHODE & "=""Méthode pénétrométrique""," & """Module pénétrométrique (MPa)"",""Module pressiométrique EM (MPa)"")"
    For r = LIGNE_COUCHES To LIGNE_COUCHES + MAX_COUCHES - 1
        ws.Range("B" & r).Value = r - LIGNE_COUCHES + 1
        ws.Range("C" & r).Validation.Delete
        Liste ws.Range("D" & r), CATEGORIES_SOL & ",Neutralisé"
        ws.Range("D" & r).ClearContents
        If r = LIGNE_COUCHES Then
            ws.Range("F" & r).Formula = "=IF(E" & r & "="""","""",ABS(E" & r & "))"
        Else
            ws.Range("F" & r).Formula = "=IF(E" & r & "="""","""",ABS(E" & r & "-E" & r - 1 & "))"
        End If
        ListeLibre ws.Range("J" & r), "18,19,20", "Poids volumique " & ChrW(947) & " (kN/m3)"
        ListeLibre ws.Range("K" & r), "0,5,15,20,25", "Cohésion c' (kPa)"
        ListeLibre ws.Range("L" & r), "15,25,30,35,40,45", "Angle de frottement " & ChrW(966) & "' (°)"
        Liste ws.Range("M" & r), "Non,Oui"
        Liste ws.Range("N" & r), "Non,Oui"
    Next r
    lc = LIGNE_COUCHES + MAX_COUCHES - 1
    ws.Range("E" & LIGNE_COUCHES & ":I" & lc).NumberFormat = "0.00"
    ws.Range("J" & LIGNE_COUCHES & ":L" & lc).NumberFormat = "0.0"
    ws.Range("M" & LIGNE_COUCHES & ":N" & lc).NumberFormat = "General"
    ws.Range("B" & LIGNE_COUCHES & ":N" & lc).HorizontalAlignment = xlCenter
    lp = LIGNE_PIEU
    SousTitre ws, lp - 1, "Caractéristiques de la fondation"
    ws.Range("B" & lp & ":B" & lp + 23).Value = Application.Transpose(Array( "Type de pieu - technique de mise en oeuvre (NF P94-262, tableau A.1)", "Diamètre constant", "Incliné", "Angle d'inclinaison sur la verticale (°)", "Effet de groupe", "Nombre de lignes de pieux m", "Nombre de pieux par ligne n", "Entraxe des pieux d (m)", "Norme de référence", "Classe de la fondation", "Catégorie de la fondation", "Abréviation du type de pieu (NF P94-262, tableau A.1)", "Mode de fonctionnement", "Prise en compte de la résistance de pointe", "Type de mise en oeuvre", "Diamètre de référence B = B0 (m), tableau des diamètres en F:G", "Nombre de fondations / appui du bâtiment", "Nombre total de fondations", "Conditions en tête de la fondation", "Essai de contrôle", "Durée d'utilisation de l'ouvrage (ans)", "Type de sollicitation (courbe de réaction)", "Largeur de référence B0 pour la réaction du sol (m)", "Durée des charges pour le calcul de Kf (courte durée : Kf ; longue durée : Kf / 2)"))
    ListeLongue ws.Range("D" & lp), TypesPieux(), "Pieux NF P94-262"
    ws.Range("D" & lp).Value = Split(TypesPieux(), ",")(0)
    Liste ws.Range("D" & lp + 1), "Oui,Non"
    Liste ws.Range("D" & lp + 2), "Non,Oui"
    Liste ws.Range("D" & lp + 4), "Non,Oui"
    ws.Range("D" & lp + 3).NumberFormat = "0.0"
    ws.Range("D" & lp + 5 & ":D" & lp + 6).NumberFormat = "0"
    ws.Range("D" & lp + 7).NumberFormat = "0.00"
    c = "D" & lp + 10
    ws.Range(c).Formula = "=IFERROR(MATCH(D" & lp & ",{""" & Replace(TypesPieux(), ",", """,""") & """},0)+16,"""")"
    ws.Range("D" & lp + 8).Formula = "=IF(" & c & "="""","""",""NF EN 1536 / 14199 / 12699"")"
    ws.Range("D" & lp + 9).Formula = "=IF(" & c & "="""","""",IF(" & c & "<=18,""1 bis"",""8""))"
    ws.Range("D" & lp + 11).Formula = FormuleAbreviation(lp)
    Liste ws.Range("D" & lp + 12), "Compression,Traction"
    ws.Range("D" & lp + 13).Formula = "=IF(" & c & "="""","""",""Non"")"
    ws.Range("D" & lp + 14).Formula = "=IF(" & c & "="""","""",""Pieu foré"")"
    ws.Range("F" & lp + 8 & ":G" & lp + 8).Value = Array("Diamètre", "B (m)")
    For r = 1 To NB_DIAMETRES
        ws.Range("F" & lp + 8 + r).Value = "B" & r - 1
        With ws.Range("G" & lp + 8 + r)
            .NumberFormat = "0.000"
            .Interior.Color = COUL_SAISIE
            If r > 1 Then
                .Validation.Add xlValidateCustom, xlValidAlertStop, , "=OR(G" & lp + 8 + r & "="""",G" & lp + 8 + r & ">G" & lp + 7 + r & ")"
                .Validation.ErrorMessage = "Chaque diamètre doit être supérieur au précédent."
            End If
        End With
    Next r
    With ws.Range("F" & lp + 8 & ":G" & lp + 8 + NB_DIAMETRES)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("F" & lp + 8 & ":G" & lp + 8)
        .Font.Bold = True
        .Interior.Color = COUL_ENTETE
    End With
    ws.Range("D" & lp + 15).Formula = "=IF(G" & lp + 9 & "="""","""",G" & lp + 9 & ")"
    ws.Range("D" & lp + 15).NumberFormat = "0.000"
    ws.Range("D" & lp + 16 & ":D" & lp + 17).NumberFormat = "0"
    Liste ws.Range("D" & lp + 18), "Libre en tête,Encastré en tête"
    Liste ws.Range("D" & lp + 19), "Non,Oui"
    Liste ws.Range("D" & lp + 20), DUREES
    ws.Range("D" & lp + 20).Value = 50
    Liste ws.Range("D" & lp + 21), TYPES_SOLLICITATION
    ws.Range("D" & lp + 22).Value = 0.6
    ws.Range("D" & lp + 22).NumberFormat = "0.000"
    Liste ws.Range("D" & lp + 23), "Longue durée,Courte durée"
    ws.Range("D" & lp + 23).Value = "Longue durée"
    ws.Range("D" & lp & ":D" & lp + 23).HorizontalAlignment = xlCenter
    For r = lp To lp + 23
        ws.Range("B" & r & ":C" & r).Merge
    Next r
    la = lp + 24
    With ws.Range("B" & la & ":I" & la)
        .Merge
        .Value = "HYPOTHÈSES RELATIVES AUX MATÉRIAUX CONSTITUTIFS DES FONDATIONS PROFONDES"
        .Font.Bold = True
        .Font.Color = COUL_TITRE
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    SousTitre ws, la + 2, "Caractéristiques du coulis de ciment"
    ws.Range("B" & la + 3 & ":B" & la + 9).Value = Application.Transpose(Array( "Rapport eau / ciment E/C du coulis", "Dosage en ciment du coulis C", "Masse volumique du coulis = C (1 + E/C)", "Résistance du coulis à 28 jours fc28", "Coefficient " & ChrW(947) & "c - combinaison ELU fondamentale", "Coefficient " & ChrW(947) & "c - combinaison accidentelle", "Coefficient " & ChrW(947) & "c - combinaison sismique"))
    ws.Range("E" & la + 3 & ":E" & la + 6).Value = Application.Transpose(Array("", "kg/m3", "kN/m3", "MPa"))
    ws.Range("D" & la + 3).Value = 0.45
    ws.Range("D" & la + 4).Value = Round(1000 / (1 / RHO_CIMENT + 0.45), 0)
    ws.Range("D" & la + 5).Formula = "=IF(OR(N(D" & la + 3 & ")=0,N(D" & la + 4 & ")=0),""""," & "D" & la + 4 & "*(1+D" & la + 3 & ")*9.81/1000)"
    ws.Range("D" & la + 6).Value = 25
    ws.Range("D" & la + 7 & ":D" & la + 9).Value = Application.Transpose(Array(1.5, 1.2, 1.3))
    ws.Range("D" & la + 3).NumberFormat = "0.00"
    ws.Range("D" & la + 4).NumberFormat = "0"
    ws.Range("D" & la + 5 & ":D" & la + 9).NumberFormat = "0.00"
    SousTitre ws, LIGNE_ACIER - 1, "Caractéristiques de l'acier"
    ws.Range("B" & LIGNE_ACIER & ":B" & LIGNE_ACIER + 21).Value = Application.Transpose(Array( "Type d'armature", "Désignation de l'armature", "Diamètre extérieur (tube) ou nominal (barre)", "Épaisseur e", "Diamètre intérieur", "Section d'acier As", "Module de Young E", "Limite d'élasticité fyk", "Coefficient " & ChrW(947) & "s - combinaison ELU fondamentale", "Coefficient " & ChrW(947) & "s - combinaison accidentelle", "Coefficient " & ChrW(947) & "s - combinaison sismique", "Durée d'utilisation du projet", "Environnement de l'armature (NF EN 1993-5, tableaux 4.1 et 4.2)", "Rayon de corrosion (perte d'épaisseur due à la corrosion, selon la durée et l'environnement)", "Épaisseur résiduelle du tube (e - perte) ou diamètre résiduel de la barre (d - 2 x perte)", "Platine d'appui (liaison fondation / structure, NF EN 1993-1-8)", "Épaisseur t de la platine", "Nuance de l'acier de la platine (NF EN 10025-2, -3, -4)", "Limite d'élasticité fy (NF EN 1993-1-1, tableau 3.1, selon t)", _
        "Résistance ultime à la traction fu (NF EN 1993-1-1, tableau 3.1, selon t)", "Coefficient partiel " & ChrW(947) & "M0 (NF EN 1993-1-1, 6.1)", "Coefficient partiel " & ChrW(947) & "M2 (NF EN 1993-1-8, tableau 2.1)"))
    ws.Range("E" & LIGNE_ACIER + 2 & ":E" & LIGNE_ACIER + 21).Value = Application.Transpose(Array( "mm", "mm", "mm", "mm²", "MPa", "MPa", "", "", "", "ans", "", "mm", "mm", "", "mm", "", "MPa", "MPa", "", ""))
    Liste ws.Range("D" & LIGNE_ACIER), TYPES_ARMATURE
    ws.Range("D" & LIGNE_ACIER + 2 & ":D" & LIGNE_ACIER + 4).NumberFormat = "0.0"
    ws.Range("D" & LIGNE_ACIER + 5 & ":D" & LIGNE_ACIER + 7).NumberFormat = "0"
    ws.Range("D" & LIGNE_ACIER + 8 & ":D" & LIGNE_ACIER + 10).Value = Application.Transpose(Array(1.15, 1, 1))
    ws.Range("D" & LIGNE_ACIER + 8 & ":D" & LIGNE_ACIER + 10).NumberFormat = "0.00"
    ws.Range("D" & LIGNE_ACIER + 11).Formula = "=D" & LIGNE_PIEU + 20
    ListeLongue ws.Range("D" & LIGNE_ACIER + 12), ENVIRONNEMENTS, "Environnements NF EN 1993-5"
    ws.Range("D" & LIGNE_ACIER + 13).Formula = "=IFERROR(INDEX({" & PERTES & "},MATCH(D" & LIGNE_ACIER + 12 & ",{""" & Replace(ENVIRONNEMENTS, ",", """,""") & """},0),MATCH(D" & LIGNE_ACIER + 11 & ",{" & DUREES & "},0)),"""")"
    ws.Range("D" & LIGNE_ACIER + 14).Formula = "=IF(OR(D" & LIGNE_ACIER + 13 & "="""",D" & LIGNE_ACIER + 2 & "=""""),"""",IF(D" & LIGNE_ACIER & "=""Tube N80"",D" & LIGNE_ACIER + 3 & "-D" & LIGNE_ACIER + 13 & ",D" & LIGNE_ACIER + 2 & "-2*D" & LIGNE_ACIER + 13 & "))"
    ws.Range("D" & LIGNE_ACIER + 13 & ":D" & LIGNE_ACIER + 14).NumberFormat = "0.00"
    ListeLongue ws.Range("D" & LIGNE_ACIER + 15), PLATINES, "Platines NF EN 1993-1-8"
    ws.Range("D" & LIGNE_ACIER + 15).Value = Split(PLATINES, ",")(2)
    ws.Range("D" & LIGNE_ACIER + 16).Formula = "=IFERROR(--SUBSTITUTE(TRIM(MID(D" & LIGNE_ACIER + 15 & ",FIND(""ép."",D" & LIGNE_ACIER + 15 & ")+3,10)),"" mm"",""""),"""")"
    Liste ws.Range("D" & LIGNE_ACIER + 17), NUANCES
    ws.Range("D" & LIGNE_ACIER + 17).Value = "S355"
    ws.Range("D" & LIGNE_ACIER + 18).Formula = FormuleNuanceEp(NUANCES_FY, NUANCES_FY80, LIGNE_ACIER + 17, LIGNE_ACIER + 16)
    ws.Range("D" & LIGNE_ACIER + 19).Formula = FormuleNuanceEp(NUANCES_FU, NUANCES_FU80, LIGNE_ACIER + 17, LIGNE_ACIER + 16)
    ws.Range("D" & LIGNE_ACIER + 20 & ":D" & LIGNE_ACIER + 21).Value = Application.Transpose(Array(1, 1.25))
    ws.Range("D" & LIGNE_ACIER + 16).NumberFormat = "0"
    ws.Range("D" & LIGNE_ACIER + 20 & ":D" & LIGNE_ACIER + 21).NumberFormat = "0.00"
    ws.Range("D" & la + 3 & ":E" & LIGNE_ACIER + 21).HorizontalAlignment = xlCenter
    For r = la + 3 To LIGNE_ACIER + 21
        If r <> LIGNE_ACIER - 1 And r <> LIGNE_ACIER - 2 Then ws.Range("B" & r & ":C" & r).Merge
    Next r
    With ws.Range("B12:C16,B19:I26,B28:E29,B" & LIGNE_COUCHES - 2 & ":C" & LIGNE_COUCHES - 2 & ",B" & LIGNE_COUCHES - 1 & ":N" & lc & ",B" & lm + 1 & ":E" & LIGNE_COUCHE_MOLLE + 6 & ",B" & lp & ":D" & lp + 23 & ",B" & la + 3 & ":E" & la + 9 & ",B" & LIGNE_ACIER & ":E" & LIGNE_ACIER + 21)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("B19:I19,B" & LIGNE_COUCHES - 1 & ":N" & LIGNE_COUCHES - 1).Font.Bold = True
    ws.Range("B19:I19,B" & LIGNE_COUCHES - 1 & ":N" & LIGNE_COUCHES - 1).Interior.Color = COUL_ENTETE
    ws.Range("B" & LIGNE_COUCHES - 1 & ":N" & LIGNE_COUCHES - 1).HorizontalAlignment = xlCenter
    ws.Range("D" & lp & ":D" & lp + 7 & ",D" & lp + 12 & ",D" & lp + 16 & ":D" & lp + 23 ).Interior.Color = COUL_SAISIE
    ws.Range("D" & lp + 8 & ":D" & lp + 11 & ",D" & lp + 13 & ":D" & lp + 15).Interior.Color = COUL_CALCUL
    ws.Range("C12:C16,D28:D29," & CELLULE_NB_COUCHES & ",D" & la + 3 & ":D" & la + 4 & ",D" & la + 6 & ":D" & la + 9 & ",D" & LIGNE_ACIER & ":D" & LIGNE_ACIER + 1 & ",D" & LIGNE_ACIER + 8 & ":D" & LIGNE_ACIER + 10 & ",D" & LIGNE_ACIER + 12 & ",D" & LIGNE_ACIER + 15 & ",D" & LIGNE_ACIER + 17 & ",D" & LIGNE_ACIER + 20 & ":D" & LIGNE_ACIER + 21).Interior.Color = COUL_SAISIE
    ws.Range("D" & la + 5 & ",D" & LIGNE_ACIER + 2 & ":D" & LIGNE_ACIER + 7 & ",D" & LIGNE_ACIER + 11 & ":D" & LIGNE_ACIER + 14 & ",D" & LIGNE_ACIER + 16 & ",D" & LIGNE_ACIER + 18 & ":D" & LIGNE_ACIER + 19).Interior.Color = COUL_CALCUL
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 40
    ws.Columns("D").ColumnWidth = 28
    ws.Columns("E:N").ColumnWidth = 14
    ws.Rows("19:" & LIGNE_ACIER + 21).AutoFit
    ws.Cells.Font.Name = "CheGuevara Text Sans"
    AjusterTextes ws
    InsererLogo ws
    AfficherTypeFondation ws
    AfficherCouches ws
    AfficherMethode ws
    AfficherOptionsPieu ws
    MajArmature ws
    ws.Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Private Sub MajCouches(ws As Worksheet, zone As Range)
    Dim c As Range, t As String
    For Each c In zone.Cells
        If c.Column = 3 And VarType(c.Value) = vbString Then
            t = Trim(c.Value)
            If t <> "" Then c.Value = UCase(Left(t, 1)) & Mid(t, 2)
            c.HorizontalAlignment = xlCenter
        ElseIf c.Column = 4 And CStr(c.Value) = "Neutralisé" Then
            ws.Range("M" & c.Row).Value = "Oui"
        End If
    Next c
End Sub
Public Sub ChangementDonnees(ws As Worksheet, Target As Range)
    Dim f As Worksheet
    On Error GoTo Fin
    Application.EnableEvents = False
    If Not Intersect(Target, ws.Range(CELLULE_PHASE)) Is Nothing Then AfficherTypeFondation ws
    If Not Intersect(Target, ws.Range(CELLULE_NB_COUCHES)) Is Nothing Then AfficherCouches ws
    If Not Intersect(Target, ws.Range("C" & LIGNE_COUCHES & ":D" & LIGNE_COUCHES + MAX_COUCHES - 1)) Is Nothing Then MajCouches ws, Intersect(Target, ws.Range("C" & LIGNE_COUCHES & ":D" & LIGNE_COUCHES + MAX_COUCHES - 1))
    If Not Intersect(Target, ws.Range("D" & LIGNE_ACIER)) Is Nothing Then MajArmature ws
    If Not Intersect(Target, ws.Range("D" & LIGNE_ACIER + 1)) Is Nothing Then MajAcier ws
    If Not Intersect(Target, Union(ws.Range("D" & LIGNE_PIEU + 2), ws.Range("D" & LIGNE_PIEU + 4))) Is Nothing Then AfficherOptionsPieu ws
    If Not Intersect(Target, ws.Range("D" & LIGNE_PIEU + 2 & ":D" & LIGNE_PIEU + 3)) Is Nothing Then
        On Error Resume Next
        Application.Run "AppliquerInclinaisonDonnees"
        On Error GoTo Fin
    End If
    If Not Intersect(Target, Union(ws.Range(CELLULE_NB_COUCHES), ws.Range("E" & LIGNE_COUCHES & ":E" & LIGNE_COUCHES + MAX_COUCHES - 1), ws.Range("D" & LIGNE_COUCHES & ":D" & LIGNE_COUCHES + MAX_COUCHES - 1), ws.Range("M" & LIGNE_COUCHES & ":M" & LIGNE_COUCHES + MAX_COUCHES - 1))) Is Nothing Then
        On Error Resume Next
        Application.Run "ActualiserPas"
        On Error GoTo Fin
    End If
    If Not Intersect(Target, ws.Range("D" & LIGNE_PIEU + 4)) Is Nothing Then
        On Error Resume Next
        Application.Run "AfficherGroupe"
        On Error GoTo Fin
    End If
    If Not Intersect(Target, ws.Range(CELLULE_METHODE)) Is Nothing Then
        AfficherMethode ws
        On Error Resume Next
        Application.Run "ActualiserCalculVerification"
        On Error GoTo Fin
    End If
    If Not Intersect(Target, ws.Range("G" & LIGNE_PIEU + 9 & ":G" & LIGNE_PIEU + 8 + NB_DIAMETRES)) Is Nothing Then
        On Error Resume Next
        Application.Run "ActualiserCalculVerification"
        On Error GoTo Fin
    End If
    If Not Intersect(Target, ws.Range("D" & LIGNE_DONNEES_FONDATION + 3)) Is Nothing Then
        On Error Resume Next
        For Each f In ws.Parent.Worksheets
            If TrouverLigneMatrice(f) > 0 Then Application.Run "ConstruireMatrice", f
        Next f
        Application.Run "ActualiserVerificationProfondeur"
    End If
Fin:
    Application.EnableEvents = True
End Sub
Public Sub AfficherOptionsPieu(ws As Worksheet)
    ws.Rows(LIGNE_PIEU + 3).Hidden = (ws.Range("D" & LIGNE_PIEU + 2).Value <> "Oui")
    ws.Rows(LIGNE_PIEU + 5).Resize(3).Hidden = (ws.Range("D" & LIGNE_PIEU + 4).Value <> "Oui")
End Sub
Public Sub AfficherMethode(ws As Worksheet)
    Dim m As String
    m = CStr(ws.Range(CELLULE_METHODE).Value)
    ws.Columns("G").Hidden = (m = "Méthode pénétrométrique")
    ws.Columns("I").Hidden = (m = "Méthode pressiométrique")
End Sub
Public Sub AfficherCouches(ws As Worksheet)
    Dim nb As Long
    nb = Val(ws.Range(CELLULE_NB_COUCHES).Value)
    If nb < 1 Then nb = 1
    If nb > MAX_COUCHES Then nb = MAX_COUCHES
    ws.Rows(LIGNE_COUCHES).Resize(MAX_COUCHES).Hidden = False
    If nb < MAX_COUCHES Then ws.Rows(LIGNE_COUCHES + nb).Resize(MAX_COUCHES - nb).Hidden = True
End Sub
Private Function FormuleNuance(valeurs As String, ligne As Long) As String
    FormuleNuance = "=INDEX({" & valeurs & "},MATCH(D" & ligne & ",{""" & Replace(NUANCES, ",", """,""") & """},0))"
End Function
Private Function FormuleNuanceEp(v40 As String, v80 As String, ligne As Long, ligneT As Long) As String
    Dim t As String
    t = "D" & ligneT
    FormuleNuanceEp = "=IF(" & t & "="""",""""," & "IF(" & t & ">80,""t > 80 mm"",IF(" & t & "<=40," & Mid(FormuleNuance(v40, ligne), 2) & "," & Mid(FormuleNuance(v80, ligne), 2) & ")))"
End Function
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
Private Function FormuleAbreviation(lp As Long) As String
    FormuleAbreviation = "=IFERROR(INDEX({""" & Replace(ABREVIATIONS, "|", """,""") & """},MATCH(D" & lp & ",{""" & Replace(TypesPieux(), ",", """,""") & """},0)),"""")"
End Function
Private Function TypesPieux() As String
    Dim s As String
    s = "Micropieu type I"
    s = s & ",Micropieu type II"
    s = s & ",Pieu ou micropieu injecté mode IGU - type III"
    s = s & ",Pieu ou micropieu injecté mode IRS - type IV"
    TypesPieux = s
End Function
Public Sub AfficherTypeFondation(ws As Worksheet)
    ws.Rows(LIGNE_FONDATION).Hidden = (ws.Range(CELLULE_PHASE).Value <> "G5")
End Sub
Public Sub InsererLogo(ws As Worksheet)
    Dim zone As Range, f As String, sh As Shape, ext As Variant, k As Long
    Set zone = ws.Range("B1:B2")
    On Error Resume Next
    For k = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(k).Name = "LogoGeoskills" Then ws.Shapes(k).Delete
    Next k
    zone.UnMerge
    zone.ClearContents
    zone.Borders.LineStyle = xlNone
    If ActiveWorkbook.Path <> "" And LCase(Left(ActiveWorkbook.Path, 4)) <> "http" Then
        For Each ext In Array("png", "jpg", "jpeg")
            f = ActiveWorkbook.Path & Application.PathSeparator & "logo_geoskills." & ext
            If Len(Dir(f)) > 0 Then Exit For
            f = ""
        Next ext
    End If
    If f <> "" Then
        Set sh = Nothing
        Set sh = ws.Shapes.AddPicture(f, msoFalse, msoTrue, zone.Left, zone.Top, -1, -1)
        If sh Is Nothing Then f = ""
    End If
    If f <> "" Then
        sh.Name = "LogoGeoskills"
        sh.LockAspectRatio = msoTrue
        sh.Height = zone.Height - 6
        If sh.Width > zone.Width - 6 Then sh.Width = zone.Width - 6
        sh.Left = zone.Left + 3
        sh.Top = zone.Top + (zone.Height - sh.Height) / 2
        sh.Placement = xlMove
    Else
        With zone
            .Merge
            .Value = "Logo GEOSKILLS" & vbLf & "(logo_geoskills.png dans le dossier du classeur)"
            .Font.Italic = True
            .Font.Size = 8
            .Font.Color = COUL_TITRE
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .BorderAround xlDash, xlThin, , COUL_TITRE
        End With
    End If
    On Error GoTo 0
End Sub
Public Sub InsererLogoDonnees()
    InsererLogo ActiveWorkbook.Worksheets(NOM_FEUILLE)
End Sub
Private Sub ListeLibre(c As Range, choix As String, titre As String)
    c.Validation.Delete
    c.Validation.Add xlValidateList, xlValidAlertInformation, , choix
    With c.Validation
        .ShowError = False
        .InputTitle = Left(titre, 32)
        .InputMessage = "Choisir dans la liste ou saisir une valeur. Ordres de grandeur : " & "« Tableaux NF P94-262 », paramètres de sols pour la vérification de la butée."
        .ShowInput = True
    End With
End Sub
Private Sub SousTitre(ws As Worksheet, ligne As Long, texte As String)
    With ws.Range("B" & ligne)
        .Value = texte
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Size = 12
        .Font.Color = COUL_TITRE
    End With
End Sub
Private Sub Liste(c As Range, choix As String)
    c.Validation.Delete
    c.Validation.Add xlValidateList, xlValidAlertStop, , choix
    c.Value = Split(choix, ",")(0)
    c.HorizontalAlignment = xlCenter
End Sub
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
        c.Worksheet.Activate
    End If
    col = 1
    Do While wl.Cells(1, col).Value <> "" And wl.Cells(1, col).Value <> nom
        col = col + 1
    Loop
    t = Split(choix, ",")
    wl.Columns(col).ClearContents
    wl.Cells(1, col).Value = nom
    wl.Cells(2, col).Resize(UBound(t) + 1, 1).Value = Application.Transpose(t)
    c.Validation.Delete
    c.Validation.Add xlValidateList, xlValidAlertStop, , "=Listes!" & wl.Cells(2, col).Resize(UBound(t) + 1, 1).Address
    c.Value = t(0)
End Sub
Private Sub AjusterTextes(ws As Worksheet)
    Dim c As Range, sh As Shape, cel() As Range, k As Long
    On Error Resume Next
    If ws.Shapes.Count > 0 Then
        ReDim cel(1 To ws.Shapes.Count)
        For k = 1 To ws.Shapes.Count
            Set cel(k) = ws.Shapes(k).TopLeftCell
        Next k
    End If
    For Each c In ws.UsedRange.Columns
        If Not c.EntireColumn.Hidden Then c.ColumnWidth = Application.WorksheetFunction.Min(255, c.ColumnWidth * 1.15)
    Next c
    For k = 1 To ws.Shapes.Count
        If Not cel(k) Is Nothing Then ws.Shapes(k).Left = cel(k).Left
    Next k
    On Error GoTo 0
End Sub
