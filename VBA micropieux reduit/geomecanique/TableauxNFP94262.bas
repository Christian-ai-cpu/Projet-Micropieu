Attribute VB_Name = "TableauxNFP94262"
Option Explicit
Private Const COUL_TITRE As Long = 5898259
Private Const COUL_ENTETE As Long = 14337945
Private Const NOM_FEUILLE As String = "Tableaux NF P94-262"
Private Const SOLS5 As String = "Argile % CaCO3 < 30 % / Limon / Sols intermédiaires|" & "Sols intermédiaires / Sable / Grave|Craie|Marne et calcaire-marneux|" & "Roche altérée ou fragmentée"
Private Const SOLS6 As String = "Argile % CaCO3 < 30 % / Limon|Sols intermédiaires|" & "Sable / Grave|Craie|Marne et calcaire-marneux|Roche altérée ou fragmentée"
Private Const LARGEUR_EMPL As Long = 9
Private Const NB_PAR_LIGNE As Long = 3
Private mHaut As Long
Private mBas As Long
Private mEmpl As Long
Private mEntete As Range
Private mDonnees As Range
Public Sub CreerTableauxNFP94262()
    Dim ws As Worksheet, s As String, g As String, c As Long
    Application.ScreenUpdating = False
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo Erreur
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = NOM_FEUILLE
    End If
    ws.Cells.Clear
    g = ChrW(947)
    With ws.Range(ws.Cells(1, 2), ws.Cells(1, 1 + LARGEUR_EMPL * NB_PAR_LIGNE - 1))
        .Merge
        .Value = "TABLEAUX DE LA NORME NF P94-262 (juillet 2012) - MICROPIEUX"
        .Font.Bold = True
        .Font.Color = COUL_TITRE
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    mHaut = 3
    mBas = 3
    mEmpl = 0
    NouveauGroupe ws, "1. GÉNÉRALITÉS", 1
    s = "1 bis|17|Micropieu type I|M1|NF EN 1536 / 14199 / 12699;" & "1 bis|18|Micropieu type II|M2|NF EN 1536 / 14199 / 12699;" & "8|19|Pieu ou micropieu injecté mode IGU (type III)|PIGU, MIGU|NF EN 1536 / 14199 / 12699;" & "8|20|Pieu ou micropieu injecté mode IRS (type IV)|PIRS, MIRS|NF EN 1536 / 14199 / 12699"
    Placer ws, "Tableau A.1 - Classes et catégories de pieux (micropieux)", "Classe|Catégorie|Technique de mise en oeuvre|Abréviation|Norme de référence", s, "La classe 1 bis complète la classe 1. Les micropieux ne correspondant pas exactement " & "aux types I à IV sont considérés par défaut comme de classe 1 bis (A.9.1)."
    NommerColonnes "NF_A1", Array("Classe", "Categorie", "Technique", "Abreviation", "Norme")
    s = "Graves|1,5 Vs|1,7 à 2,4;Graves sableuses|1,5 Vs|1,7 à 2,4;Sables graveleux|1,5 Vs|1,7 à 2,4;" & "Sables grossiers|1,5 Vs|1,7 à 2,4;Sables moyens|1,5 Vs|1,7 à 2,4;Sables fins|1,5 Vs|1,7 à 2,4;" & "Sables limoneux|1,5 à 2 Vs pour IRS - 1,5 Vs pour IGU|1,7 à 2,4;" & "Limons|2 Vs pour IRS - 1,5 Vs pour IGU|1,7 à 2,4;" & "Argiles|2,5 à 3 Vs pour IRS - 1,5 à 2 Vs pour IGU|1,7 à 2,4;" & "Marnes|1,5 à 2 Vs pour couche compacte|1,7 à 2,4;" & "Marno-calcaires|2 à 6 Vs ou plus si couche fracturée|1,7 à 2,4;" & "Craie altérée ou fragmentée|1,1 à 1,5 Vs si couche finement fissurée|1,7 à 2,4;" & "Rocher altéré ou fragmenté|2 Vs ou plus si couche fracturée|1,7 à 2,4"
    Placer ws, "Tableau A.9.1 - Quantité de coulis recommandée pour le scellement " & "des micropieux", "Sols|Quantité usuelle de coulis à injecter Vi|Dosage coulis C/E", s, "Vs : volume du scellement calculé à partir du diamètre de forage."
    NouveauGroupe ws, "2. FACTEURS PARTIELS, DE CORRÉLATION ET DE MODÈLE", 1
    s = "Pointe|" & g & "b|1.1|1;Fût (compression)|" & g & "s|1.1|1;" & "Total / combiné (compression)|" & g & "t|1.1|1;Fût en traction|" & g & "s;t|1.15|1.05"
    Placer ws, "Tableaux C.2.3.1 et C.2.3.2 - Facteurs partiels de résistance " & g & "R pour les pieux (ELU)", "Résistance|Symbole|Situations durables et transitoires|" & "Situations accidentelles", s, "Valeurs identiques pour les pieux foncés, forés et à la tarière continue (jeu R2)."
    s = "Fût en compression|" & g & "cr|0.9|1.1;Fût en traction (y compris micropieux)|" & g & "s;cr|1.1|1.5"
    Placer ws, "Tableaux 14.2.1.1 et 14.2.1.2 - Facteurs partiels de résistance " & g & "R pour les fondations profondes (ELS)", "Résistance|Symbole|" & "Combinaisons caractéristiques|Combinaisons quasi permanentes", s, ""
    s = "XI1'|1.4|1.3|1.2|1.1|1;XI2'|1.4|1.2|1.05|1|1"
    Placer ws, "Tableau C.2.4.1 - Facteurs de corrélation " & ChrW(958) & "' (essais de chargement statique, n = nombre de pieux testés)", Replace("XI'|n = 1|n = 2|n = 3|n = 4|n >= 5", "XI", ChrW(958)), Replace(s, "XI", ChrW(958)), ""
    s = "XI3'|1.4|1.35|1.33|1.31|1.29|1.27|1.25;XI4'|1.4|1.27|1.23|1.2|1.15|1.12|1.08"
    Placer ws, "Tableau C.2.4.2 - Facteurs de corrélation " & ChrW(958) & "' (essais sur les sols, n = nombre de profils d'essais)", Replace("XI'|n = 1|n = 2|n = 3|n = 4|n = 5|n = 7|n = 10", "XI", ChrW(958)), Replace(s, "XI", ChrW(958)), ""
    s = "Pieux de catégories 10, 15, 17, 18, 19 et 20|2|2|1.1|1.1"
    Placer ws, "Tableau F.2.1 - Coefficients de modèle - Méthode pressiométrique", "Pieux|" & g & "R;d1 compression|" & g & "R;d1 traction|" & g & "R;d2 compression|" & g & "R;d2 traction", s, g & "R;d1 : procédure du « pieu modèle » ; " & g & "R;d2 : procédure du « modèle de terrain » (non utilisé pour le « pieu modèle »)."
    s = "Pieux de catégories 10, 15, 17, 18, 19 et 20|2|2|1.1|1.1"
    Placer ws, "Tableau G.2.1 - Coefficients de modèle - Méthode pénétrométrique", "Pieux|" & g & "R;d1 compression|" & g & "R;d1 traction|" & g & "R;d2 compression|" & g & "R;d2 traction", s, ""
    NouveauGroupe ws, "3. MÉTHODES DE CALCUL", 1
    NouveauGroupe ws, "Méthode pressiométrique (annexe F)", 2
    s = "17 - M1|1.1|1|1.8|1.5|1.6;18 - M2|1.1|1|1.8|1.5|1.6;" & "19 - PIGU, MIGU|2.7|2.9|2.4|2.4|2.4;20 - PIRS, MIRS|3.4|3.8|3.1|3.1|3.1"
    Placer ws, "Tableau F.5.2.1 - Choix des valeurs de " & ChrW(945) & " pieu-sol - Méthode pressiométrique", "Catégorie|" & SOLS5, s, "Catégories 17 et 18 : valeurs de la catégorie 1 (FS, foré simple), technique la plus proche. " & "BE, HBi, MIGU, PIGU, PIRS, MIRS : valeurs pour une injection stricte et soignée, " & "à confirmer par les essais de conformité (section 8.9)."
    s = "Courbe|Q1|Q2|Q3|Q4|Q5;a|0.003|0.01|0.007|0.008|0.01;b|0.04|0.06|0.07|0.08|0.08;" & "c|3.5|1.2|1.3|3|3"
    Placer ws, "Tableau F.5.2.2 - Paramètres a, b et c des courbes fsol - " & "Méthode pressiométrique", "Type de sol|" & SOLS5, s, "fsol = (a.pl* + b)(1 - exp(-c.pl*)), pl* en MPa."
    Placer ws, "Tableau F.5.2.3 - Valeurs maximales du frottement axial unitaire " & "limite qs (kPa) - Méthode pressiométrique", "Catégorie|" & SOLS6, QsMax(), "Catégories 17 et 18 : valeurs de la catégorie 1 (FS, foré simple), technique la plus proche. " & "BAO, HB, PP vibrofoncés : abattement de 30 % sur qs."
    ActiveWorkbook.Names.Add Name:="NF_Sols", RefersTo:="='" & ws.Name & "'!" & mEntete.Offset(0, 1).Resize(1, mEntete.Columns.Count - 1).Address
    NouveauGroupe ws, "Méthode pénétrométrique (annexe G)", 2
    s = "17 - M1|0.55|0.65|0.7|0.8|1.4|1.5;18 - M2|0.55|0.65|0.7|0.8|1.4|1.5;" & "19 - PIGU, MIGU|1.35|1.6|2|1.1|2.25|2.25;20 - PIRS, MIRS|1.7|2.05|2.65|1.4|2.9|2.9"
    Placer ws, "Tableau G.5.2.1 - Choix des valeurs de " & ChrW(945) & " pieu-sol - Méthode pénétrométrique", "Catégorie|" & SOLS6, s, "Catégories 17 et 18 : valeurs de la catégorie 1 (FS, foré simple), technique la plus proche. " & "Lignes des micropieux (catégories 17 à 20) seulement."
    s = "Courbe|Q1|Q2|Q3|Q2|Q2|Q2;a|0.0018|0.0015|0.0012|0.0015|0.0015|0.0015;" & "b|0.1|0.1|0.1|0.1|0.1|0.1;c|0.4|0.25|0.15|0.25|0.25|0.25"
    Placer ws, "Tableau G.5.2.2 - Paramètres a, b et c des courbes fsol - " & "Méthode pénétrométrique", "Type de sol|" & SOLS6, s, "fsol = (a.qc + b)(1 - exp(-c.qc)), qc en MPa."
    Placer ws, "Tableau G.5.2.3 - Valeurs maximales du frottement axial unitaire " & "limite qs (kPa) - Méthode pénétrométrique", "Catégorie|" & SOLS6, QsMax(), "Mêmes valeurs que le tableau F.5.2.3."
    NouveauGroupe ws, "4. DURABILITÉ DES ARMATURES (NF EN 1993-5)", 1
    s = "Sols naturels non remaniés (sable, limon, argile, schiste)|0|0.3|0.6|0.9|1.2;" & "Sols naturels pollués et sites industriels|0.15|0.75|1.5|2.25|3;" & "Sols naturels agressifs (marécages, tourbe)|0.2|1|1.75|2.5|3.25;" & "Remblais non compactés et non agressifs|0.18|0.7|1.2|1.7|2.2;" & "Remblais non compactés et agressifs (cendres, scories)|0.5|2|3.25|4.5|5.75"
    Placer ws, "Tableau 4.1 - Perte d'épaisseur (mm) due à la corrosion des pieux et " & "palplanches dans les sols, avec ou sans nappe phréatique", "Durée d'utilisation requise|5 ans|25 ans|50 ans|75 ans|100 ans", s, "Remblais compactés : valeurs divisées par 2. Valeurs à 5 et 25 ans mesurées, " & "autres valeurs extrapolées. Valeurs recommandées, à vérifier avec l'annexe nationale."
    s = "Eau douce commune (rivière, canal) - zone de fort impact (ligne de flottaison)|0.15|0.55|0.9|1.15|1.4;" & "Eau douce très polluée (égouts, effluents industriels) - zone de fort impact|0.3|1.3|2.3|3.3|4.3;" & "Eau de mer en climat tempéré - zone de fort impact (basses eaux, éclaboussures)|0.55|1.9|3.75|5.6|7.5;" & "Eau de mer en climat tempéré - zone d'immersion permanente ou de marnage|0.25|0.9|1.75|2.6|3.5"
    Placer ws, "Tableau 4.2 - Perte d'épaisseur (mm) due à la corrosion des pieux et " & "palplanches en eau douce ou en eau de mer", "Durée d'utilisation requise|5 ans|25 ans|50 ans|75 ans|100 ans", s, "Valeurs à 5 et 25 ans mesurées, autres valeurs extrapolées. Valeurs recommandées, " & "à vérifier avec l'annexe nationale."
    s = "Sols naturels non remaniés (sable, limon, argile, schiste)|0|0.3|0.6|0.9|1.2;" & "Sols naturels pollués et sites industriels|0.15|0.75|1.5|2.25|3;" & "Sols naturels agressifs (marécages, tourbe)|0.2|1|1.75|2.5|3.25;" & "Remblais non compactés et non agressifs|0.18|0.7|1.2|1.7|2.2;" & "Remblais non compactés et agressifs (cendres, scories)|0.5|2|3.25|4.5|5.75;" & "Remblais compactés et non agressifs|0.09|0.35|0.6|0.85|1.1;" & "Remblais compactés et agressifs (cendres, scories)|0.25|1|1.63|2.25|2.88;" & "Eau douce (rivière, canal) - zone de fort impact|0.15|0.55|0.9|1.15|1.4;" & "Eau douce très polluée - zone de fort impact|0.3|1.3|2.3|3.3|4.3;" & "Eau de mer en climat tempéré - zone de fort impact|0.55|1.9|3.75|5.6|7.5;" & "Eau de mer en climat tempéré - zone d'immersion permanente ou de marnage|0.25|0.9|1.75|2.6|3.5"
    Placer ws, "Valeurs recommandées de la perte d'épaisseur (mm) due à la corrosion - " & "rayon de corrosion selon la nature du terrain et la durée d'utilisation " & "(NF EN 1993-5, tableaux 4.1 et 4.2)", "Nature du terrain / environnement|5 ans|25 ans|50 ans|75 ans|100 ans", s, "Remblais compactés : valeurs des remblais non compactés divisées par 2 (note du " & "tableau 4.1). Valeur retenue dans « Données du projet » (caractéristiques de " & "l'acier) : durée d'utilisation du projet x environnement de l'armature."
    NouveauGroupe ws, "5. RÉACTION DU SOL - MODULE DE RÉACTION Kf (ANNEXE I)", 1
    g = "0.666666666666667"
    s = "Tourbe|-|1|-|1|-|1;" & "Argile|> 16|1|9 à 16|" & g & "|7 à 9|0.5;" & "Limon|> 14|" & g & "|8 à 14|0.5|-|0.5;" & "Sable|> 12|0.5|7 à 12|0.333333333333333|-|0.333333333333333;" & "Grave|> 10|0.333333333333333|6 à 10|0.25|-|0.25"
    Placer ws, "Tableau I.1.3.1 - Coefficient rhéologique " & ChrW(945) & " des sols", "Type de sol|I - surconsolidé ou très serré : EM/pl*|" & ChrW(945) & "|II - normalement consolidé ou normalement serré : EM/pl*|" & ChrW(945) & "|III - sous-consolidé altéré et remanié ou lâche : EM/pl*|" & ChrW(945), s, "Catégorie du sol choisie d'après le rapport EM/pl* ; " & ChrW(945) & " retenu dans « Données du " & "projet » (couche molle : catégorie et nature du sol)."
    mDonnees.Columns(3).NumberFormat = "# ?/?"
    mDonnees.Columns(5).NumberFormat = "# ?/?"
    mDonnees.Columns(7).NumberFormat = "# ?/?"
    NommerColonnes "NF_Rheo", Array("Nature", "EMpl1", "Alpha1", "EMpl2", "Alpha2", "EMpl3", "Alpha3")
    s = "Rocher très peu fracturé|" & g & ";Rocher normal|0.5;Rocher très fracturé|0.333333333333333;" & "Rocher très altéré|" & g
    Placer ws, "Tableau I.1.3.2 - Coefficient rhéologique " & ChrW(945) & " des rochers", "Type de rocher|" & ChrW(945), s, ""
    mDonnees.Columns(2).NumberFormat = "# ?/?"
    NommerColonnes "NF_RheoRoc", Array("Nature", "Alpha")
    s = "B " & ChrW(8804) & " B0|Kf = 12 EM / [4/3 x (2,65)^" & ChrW(945) & " + " & ChrW(945) & "];" & "B > B0|Kf = 12 EM / [4/3 x (B0/B) x (2,65 B/B0)^" & ChrW(945) & " + " & ChrW(945) & "];" & "Diamètre de référence|B0 = 0,60 m;" & "Charges de courte durée|Kf;" & "Charges de longue durée|Kf / 2;" & "Méthode pénétrométrique|Kf = " & ChrW(946) & " qc (courte durée), " & ChrW(946) & " qc / 2 (longue durée)"
    Placer ws, "Module linéique de réaction du sol Kf (annexe I)", "Condition|Expression", s, "EM en MPa, B et B0 en m ; Kf en MPa. Durée des charges choisie dans « Données du projet » " & "(caractéristiques de la fondation)."
    NouveauGroupe ws, "6. PARAMÈTRES DE SOLS", 1
    s = "Roche altérée / Enrochements|45|0|20;Gravier propre|40|0|19;" & "Gravier limoneux ou argileux|35|0|20;Sable propre|35|0|19;Sable limoneux|30|0|20;" & "Sable argileux|25|5|20;Limon|35|0|19;Limon argileux|30|15|19;Argile limoneuse|25|20|19;" & "Argile|15|25|18"
    Placer ws, "Ordre de grandeur des paramètres de sols pour la vérification de la butée", "Type de sol|" & ChrW(966) & "' (°)|c' (kPa)|" & ChrW(947) & " (kN/m3)", s, "Valeurs indicatives, à remplacer par celles des essais quand elles existent. " & "Proposées en listes déroulantes dans « Données du projet » (tableau des couches)."
    NommerColonnes "NF_ParamSol", Array("Type", "Phi", "C", "Gamma")
    ws.Columns("A").ColumnWidth = 3
    For c = 0 To NB_PAR_LIGNE - 1
        ws.Columns(2 + c * LARGEUR_EMPL).ColumnWidth = 32
        ws.Columns(3 + c * LARGEUR_EMPL).Resize(, LARGEUR_EMPL - 2).ColumnWidth = 12
        ws.Columns(1 + (c + 1) * LARGEUR_EMPL).ColumnWidth = 3
    Next c
    ws.Cells.Font.Name = "CheGuevara Text Sans"
    AjusterTextes ws
    ws.Activate
Fin:
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    MsgBox "Erreur pendant la création de la feuille « Tableaux NF P94-262 » : " & Err.Description & " (n° " & Err.Number & ")", vbExclamation
    Resume Fin
End Sub
Private Sub NouveauGroupe(ws As Worksheet, titre As String, niveau As Long)
    If mEmpl > 0 Then mHaut = mBas
    mEmpl = 0
    With ws.Range(ws.Cells(mHaut, 2), ws.Cells(mHaut, 1 + LARGEUR_EMPL * NB_PAR_LIGNE - 1))
        .Merge
        .Value = titre
        .Font.Bold = True
        If niveau = 1 Then
            .Font.Size = 13
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = COUL_TITRE
        Else
            .Font.Size = 11
            .Font.Underline = xlUnderlineStyleSingle
            .Font.Color = COUL_TITRE
        End If
    End With
    mHaut = mHaut + 2
    mBas = mHaut
End Sub
Private Sub Placer(ws As Worksheet, titre As String, entetes As String, donnees As String, note As String)
    Dim fin As Long
    fin = EcrireTableau(ws, mHaut, 2 + mEmpl * LARGEUR_EMPL, titre, entetes, donnees, note)
    If fin > mBas Then mBas = fin
    mEmpl = mEmpl + 1
    If mEmpl = NB_PAR_LIGNE Then
        mEmpl = 0
        mHaut = mBas
    End If
End Sub
Private Sub NommerColonnes(prefixe As String, noms As Variant)
    Dim k As Long
    For k = 0 To UBound(noms)
        ActiveWorkbook.Names.Add Name:=prefixe & "_" & noms(k), RefersTo:="='" & mDonnees.Worksheet.Name & "'!" & mDonnees.Columns(k + 1).Address
    Next k
End Sub
Private Function QsMax() As String
    QsMax = "17 - M1|90|90|90|200|170|200;18 - M2|90|90|90|200|170|200;19 - PIGU, MIGU|200|200|380|320|320|320;" & "20 - PIRS, MIRS|200|200|440|440|440|500"
End Function
Private Function EcrireTableau(ws As Worksheet, ByVal r As Long, c0 As Long, titre As String, entetes As String, donnees As String, note As String) As Long
    Dim e As Variant, lignes As Variant, cols As Variant, i As Long, j As Long, debut As Long
    With ws.Range(ws.Cells(r, c0), ws.Cells(r, c0 + LARGEUR_EMPL - 2))
        .Merge
        .Value = titre
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Color = COUL_TITRE
        .WrapText = True
        .VerticalAlignment = xlBottom
    End With
    If ws.Rows(r).RowHeight < 30 Then ws.Rows(r).RowHeight = 30
    r = r + 1
    debut = r
    e = Split(entetes, "|")
    For j = 0 To UBound(e)
        ws.Cells(r, c0 + j).Value = e(j)
    Next j
    With ws.Range(ws.Cells(r, c0), ws.Cells(r, c0 + UBound(e)))
        .Font.Bold = True
        .Interior.Color = COUL_ENTETE
        .HorizontalAlignment = xlCenter
    End With
    lignes = Split(donnees, ";")
    For i = 0 To UBound(lignes)
        r = r + 1
        cols = Split(lignes(i), "|")
        For j = 0 To UBound(cols)
            If j > 0 And EstNombre(CStr(cols(j))) Then
                ws.Cells(r, c0 + j).Value = Val(cols(j))
            Else
                ws.Cells(r, c0 + j).NumberFormat = "@"
                ws.Cells(r, c0 + j).Value = cols(j)
            End If
        Next j
    Next i
    With ws.Range(ws.Cells(debut, c0), ws.Cells(r, c0 + UBound(e)))
        .Borders.LineStyle = xlContinuous
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(debut + 1, c0 + 1), ws.Cells(r, c0 + UBound(e))).HorizontalAlignment = xlCenter
    Set mEntete = ws.Range(ws.Cells(debut, c0), ws.Cells(debut, c0 + UBound(e)))
    Set mDonnees = ws.Range(ws.Cells(debut + 1, c0), ws.Cells(r, c0 + UBound(e)))
    If note <> "" Then
        r = r + 1
        With ws.Range(ws.Cells(r, c0), ws.Cells(r, c0 + LARGEUR_EMPL - 2))
            .Merge
            .Value = note
            .Font.Italic = True
            .Font.Size = 9
            .WrapText = True
            .RowHeight = 24
        End With
    End If
    EcrireTableau = r + 2
End Function
Private Function EstNombre(t As String) As Boolean
    Dim i As Long, c As String
    If t = "" Then Exit Function
    For i = 1 To Len(t)
        c = Mid(t, i, 1)
        If Not (c Like "[0-9]" Or c = ".") Then Exit Function
    Next i
    EstNombre = True
End Function
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
