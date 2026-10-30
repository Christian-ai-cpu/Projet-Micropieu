Attribute VB_Name = "TableauxNFP94262"
Option Explicit

' Module « Tableaux NF P94-262 » : crée (ou recrée) la feuille « Tableaux NF P94-262 »
' avec les tableaux de la norme NF P94-262 (juillet 2012) utiles au calcul des
' micropieux. Version MICROPIEUX : seules les lignes des catégories 17 à 20 (M1, M2,
' PIGU / MIGU, PIRS / MIRS) sont reprises ; les tableaux kp max (F.4.2.1) et kc max
' (G.4.2.1) ne sont pas repris (résistance de pointe non prise en compte) :
'   A.1               : classes et catégories de pieux, norme de référence
'   A.9.1             : quantité de coulis recommandée pour le scellement des micropieux
'   C.2.3.1 / C.2.3.2 : facteurs partiels de résistance ELU (durable / accidentel)
'   14.2.1.1 / 14.2.1.2 : facteurs partiels de résistance ELS
'   C.2.4.1 / C.2.4.2 : facteurs de corrélation xi'
'   F.2.1 / G.2.1     : coefficients de modèle (pressiomètre / pénétromètre), ligne
'                       des catégories 10, 15 et 17 à 20
'   F.5.2.1, F.5.2.2, F.5.2.3 : méthode pressiométrique (alpha pieu-sol, paramètres
'                       de fsol, qs max)
'   G.5.2.1, G.5.2.2, G.5.2.3 : méthode pénétrométrique (alpha pieu-sol, paramètres
'                       de fsol, qs max)
' Disposition : 3 groupes (généralités ; facteurs partiels, de corrélation et de
' modèle ; méthodes de calcul) et, dans chaque groupe, 3 tableaux côte à côte.
' « - » : pas de valeur dans la norme. Les notes des tableaux sont résumées sous
' chaque tableau ; se reporter à la norme pour leur texte complet.
' Plages nommées créées pour la feuille « Données du projet » : NF_A1_Classe,
' NF_A1_Categorie, NF_A1_Technique, NF_A1_Abreviation, NF_A1_Norme (tableau A.1) et
' NF_Sols (catégories de sol, en-têtes des tableaux qs max).

Private Const NOM_FEUILLE As String = "Tableaux NF P94-262"
' Colonnes de sols : 5 colonnes (pressiomètre, tableaux F.5.2.1, F.5.2.2)
' et 6 colonnes (qs max et méthode pénétrométrique)
Private Const SOLS5 As String = "Argile % CaCO3 < 30 % / Limon / Sols intermédiaires|" & _
    "Sols intermédiaires / Sable / Grave|Craie|Marne et calcaire-marneux|" & _
    "Roche altérée ou fragmentée"
Private Const SOLS6 As String = "Argile % CaCO3 < 30 % / Limon|Sols intermédiaires|" & _
    "Sable / Grave|Craie|Marne et calcaire-marneux|Roche altérée ou fragmentée"

' Disposition des tableaux : par groupes (titre de groupe), 3 tableaux côte à côte
' par ligne. Chaque emplacement fait LARGEUR_EMPL colonnes (tableau de 8 colonnes
' au plus + 1 colonne vide), à partir de la colonne B.
Private Const LARGEUR_EMPL As Long = 9
Private Const NB_PAR_LIGNE As Long = 3
Private mHaut As Long          ' première ligne de la rangée de tableaux en cours
Private mBas As Long           ' ligne libre sous le plus long tableau de la rangée
Private mEmpl As Long          ' emplacement suivant dans la rangée (0 à 2)
Private mEntete As Range       ' ligne d'en-tête du dernier tableau écrit
Private mDonnees As Range      ' lignes de données du dernier tableau écrit

' Thème vert : saisies en vert, valeurs non saisies (calculées) en jaune, en-têtes et
' titres en vert (couleurs RGB écrites en nombre : rouge + 256 x vert + 65536 x bleu)
Private Const VERT_SAISIE As Long = 13561798   ' RGB(198, 239, 206)
Private Const JAUNE_CALCUL As Long = 13434879  ' RGB(255, 255, 204), jaune clair
Private Const VERT_ENTETE As Long = 14348258   ' RGB(226, 239, 218)
Private Const VERT_FONCE As Long = 2315831     ' RGB(55, 86, 35)
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
    g = ChrW(947)                            ' gamma

    With ws.Range(ws.Cells(1, 2), ws.Cells(1, 1 + LARGEUR_EMPL * NB_PAR_LIGNE - 1))
        .Merge
        .Value = "TABLEAUX DE LA NORME NF P94-262 (juillet 2012) - MICROPIEUX"
        .Font.Bold = True
        .Font.Color = VERT_FONCE
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    mHaut = 3
    mBas = 3
    mEmpl = 0

    NouveauGroupe ws, "1. GÉNÉRALITÉS", 1
    ' --- Tableau A.1 ---
    ' Version micropieux : catégories 17 à 20 seulement
    s = "1 bis|17|Micropieu type I|M1|NF EN 1536 / 14199 / 12699;" & _
        "1 bis|18|Micropieu type II|M2|NF EN 1536 / 14199 / 12699;" & _
        "8|19|Pieu ou micropieu injecté mode IGU (type III)|PIGU, MIGU|NF EN 1536 / 14199 / 12699;" & _
        "8|20|Pieu ou micropieu injecté mode IRS (type IV)|PIRS, MIRS|NF EN 1536 / 14199 / 12699"
    Placer ws, "Tableau A.1 - Classes et catégories de pieux (micropieux)", _
        "Classe|Catégorie|Technique de mise en oeuvre|Abréviation|Norme de référence", s, _
        "La classe 1 bis complète la classe 1. Les micropieux ne correspondant pas exactement " & _
        "aux types I à IV sont considérés par défaut comme de classe 1 bis (A.9.1)."
    ' Colonnes nommées, utilisées par la feuille « Données du projet »
    NommerColonnes "NF_A1", Array("Classe", "Categorie", "Technique", "Abreviation", "Norme")

    ' --- Tableau A.9.1 ---
    s = "Graves|1,5 Vs|1,7 à 2,4;Graves sableuses|1,5 Vs|1,7 à 2,4;Sables graveleux|1,5 Vs|1,7 à 2,4;" & _
        "Sables grossiers|1,5 Vs|1,7 à 2,4;Sables moyens|1,5 Vs|1,7 à 2,4;Sables fins|1,5 Vs|1,7 à 2,4;" & _
        "Sables limoneux|1,5 à 2 Vs pour IRS - 1,5 Vs pour IGU|1,7 à 2,4;" & _
        "Limons|2 Vs pour IRS - 1,5 Vs pour IGU|1,7 à 2,4;" & _
        "Argiles|2,5 à 3 Vs pour IRS - 1,5 à 2 Vs pour IGU|1,7 à 2,4;" & _
        "Marnes|1,5 à 2 Vs pour couche compacte|1,7 à 2,4;" & _
        "Marno-calcaires|2 à 6 Vs ou plus si couche fracturée|1,7 à 2,4;" & _
        "Craie altérée ou fragmentée|1,1 à 1,5 Vs si couche finement fissurée|1,7 à 2,4;" & _
        "Rocher altéré ou fragmenté|2 Vs ou plus si couche fracturée|1,7 à 2,4"
    Placer ws, "Tableau A.9.1 - Quantité de coulis recommandée pour le scellement " & _
        "des micropieux", "Sols|Quantité usuelle de coulis à injecter Vi|Dosage coulis C/E", s, _
        "Vs : volume du scellement calculé à partir du diamètre de forage."

    NouveauGroupe ws, "2. FACTEURS PARTIELS, DE CORRÉLATION ET DE MODÈLE", 1
    ' --- Facteurs partiels de résistance ---
    s = "Pointe|" & g & "b|1.1|1;Fût (compression)|" & g & "s|1.1|1;" & _
        "Total / combiné (compression)|" & g & "t|1.1|1;Fût en traction|" & g & "s;t|1.15|1.05"
    Placer ws, "Tableaux C.2.3.1 et C.2.3.2 - Facteurs partiels de résistance " & _
        g & "R pour les pieux (ELU)", "Résistance|Symbole|Situations durables et transitoires|" & _
        "Situations accidentelles", s, _
        "Valeurs identiques pour les pieux foncés, forés et à la tarière continue (jeu R2)."
    s = "Fût en compression|" & g & "cr|0.9|1.1;Fût en traction (y compris micropieux)|" & _
        g & "s;cr|1.1|1.5"
    Placer ws, "Tableaux 14.2.1.1 et 14.2.1.2 - Facteurs partiels de résistance " & _
        g & "R pour les fondations profondes (ELS)", "Résistance|Symbole|" & _
        "Combinaisons caractéristiques|Combinaisons quasi permanentes", s, ""

    ' --- Facteurs de corrélation ---
    s = "XI1'|1.4|1.3|1.2|1.1|1;XI2'|1.4|1.2|1.05|1|1"
    Placer ws, "Tableau C.2.4.1 - Facteurs de corrélation " & ChrW(958) & _
        "' (essais de chargement statique, n = nombre de pieux testés)", _
        Replace("XI'|n = 1|n = 2|n = 3|n = 4|n >= 5", "XI", ChrW(958)), _
        Replace(s, "XI", ChrW(958)), ""
    s = "XI3'|1.4|1.35|1.33|1.31|1.29|1.27|1.25;XI4'|1.4|1.27|1.23|1.2|1.15|1.12|1.08"
    Placer ws, "Tableau C.2.4.2 - Facteurs de corrélation " & ChrW(958) & _
        "' (essais sur les sols, n = nombre de profils d'essais)", _
        Replace("XI'|n = 1|n = 2|n = 3|n = 4|n = 5|n = 7|n = 10", "XI", ChrW(958)), _
        Replace(s, "XI", ChrW(958)), ""

    ' --- Coefficients de modèle ---
    s = "Pieux de catégories 10, 15, 17, 18, 19 et 20|2|2|1.1|1.1"
    Placer ws, "Tableau F.2.1 - Coefficients de modèle - Méthode pressiométrique", _
        "Pieux|" & g & "R;d1 compression|" & g & "R;d1 traction|" & g & "R;d2 compression|" & _
        g & "R;d2 traction", s, g & "R;d1 : procédure du « pieu modèle » ; " & g & _
        "R;d2 : procédure du « modèle de terrain » (non utilisé pour le « pieu modèle »)."
    s = "Pieux de catégories 10, 15, 17, 18, 19 et 20|2|2|1.1|1.1"
    Placer ws, "Tableau G.2.1 - Coefficients de modèle - Méthode pénétrométrique", _
        "Pieux|" & g & "R;d1 compression|" & g & "R;d1 traction|" & g & "R;d2 compression|" & _
        g & "R;d2 traction", s, ""

    NouveauGroupe ws, "3. MÉTHODES DE CALCUL", 1
    NouveauGroupe ws, "Méthode pressiométrique (annexe F)", 2
    ' --- Méthode pressiométrique ---
    s = "17 - M1|-|-|-|-|-;18 - M2|-|-|-|-|-;" & _
        "19 - PIGU, MIGU|2.7|2.9|2.4|2.4|2.4;20 - PIRS, MIRS|3.4|3.8|3.1|3.1|3.1"
    Placer ws, "Tableau F.5.2.1 - Choix des valeurs de " & ChrW(945) & _
        " pieu-sol - Méthode pressiométrique", "Catégorie|" & SOLS5, s, _
        "Catégories 17 et 18 : prendre les valeurs des techniques les plus proches. " & _
        "BE, HBi, MIGU, PIGU, PIRS, MIRS : valeurs pour une injection stricte et soignée, " & _
        "à confirmer par les essais de conformité (section 8.9)."
    s = "Courbe|Q1|Q2|Q3|Q4|Q5;a|0.003|0.01|0.007|0.008|0.01;b|0.04|0.06|0.07|0.08|0.08;" & _
        "c|3.5|1.2|1.3|3|3"
    Placer ws, "Tableau F.5.2.2 - Paramètres a, b et c des courbes fsol - " & _
        "Méthode pressiométrique", "Type de sol|" & SOLS5, s, _
        "fsol = (a.pl* + b)(1 - exp(-c.pl*)), pl* en MPa."
    Placer ws, "Tableau F.5.2.3 - Valeurs maximales du frottement axial unitaire " & _
        "limite qs (kPa) - Méthode pressiométrique", "Catégorie|" & SOLS6, QsMax(), _
        "Catégories 17 et 18 : prendre les valeurs des techniques les plus proches. " & _
        "BAO, HB, PP vibrofoncés : abattement de 30 % sur qs."
    ' Catégories de sol (en-têtes du tableau qs max), utilisées par « Données du projet »
    ActiveWorkbook.Names.Add Name:="NF_Sols", RefersTo:="='" & ws.Name & "'!" & _
        mEntete.Offset(0, 1).Resize(1, mEntete.Columns.Count - 1).Address

    NouveauGroupe ws, "Méthode pénétrométrique (annexe G)", 2
    ' --- Méthode pénétrométrique ---
    s = "17 - M1|-|-|-|-|-|-;18 - M2|-|-|-|-|-|-;" & _
        "19 - PIGU, MIGU|1.35|1.6|2|1.1|2.25|2.25;20 - PIRS, MIRS|1.7|2.05|2.65|1.4|2.9|2.9"
    Placer ws, "Tableau G.5.2.1 - Choix des valeurs de " & ChrW(945) & _
        " pieu-sol - Méthode pénétrométrique", "Catégorie|" & SOLS6, s, _
        "Catégories 17 et 18 : prendre les valeurs des techniques les plus proches. " & _
        "Lignes des micropieux (catégories 17 à 20) seulement."
    s = "Courbe|Q1|Q2|Q3|Q2|Q2|Q2;a|0.0018|0.0015|0.0012|0.0015|0.0015|0.0015;" & _
        "b|0.1|0.1|0.1|0.1|0.1|0.1;c|0.4|0.25|0.15|0.25|0.25|0.25"
    Placer ws, "Tableau G.5.2.2 - Paramètres a, b et c des courbes fsol - " & _
        "Méthode pénétrométrique", "Type de sol|" & SOLS6, s, _
        "fsol = (a.qc + b)(1 - exp(-c.qc)), qc en MPa."
    Placer ws, "Tableau G.5.2.3 - Valeurs maximales du frottement axial unitaire " & _
        "limite qs (kPa) - Méthode pénétrométrique", "Catégorie|" & SOLS6, QsMax(), _
        "Mêmes valeurs que le tableau F.5.2.3."

    ' Largeurs : 1re colonne de chaque emplacement large (intitulés), colonnes de
    ' valeurs étroites, colonne vide entre deux tableaux
    ws.Columns("A").ColumnWidth = 3
    For c = 0 To NB_PAR_LIGNE - 1
        ws.Columns(2 + c * LARGEUR_EMPL).ColumnWidth = 32
        ws.Columns(3 + c * LARGEUR_EMPL).Resize(, LARGEUR_EMPL - 2).ColumnWidth = 12
        ws.Columns(1 + (c + 1) * LARGEUR_EMPL).ColumnWidth = 3
    Next c
    ws.Activate
Fin:
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    MsgBox "Erreur pendant la création de la feuille « Tableaux NF P94-262 » : " & _
        Err.Description & " (n° " & Err.Number & ")", vbExclamation
    Resume Fin
End Sub


' Termine la rangée en cours et écrit un titre de groupe (niveau 1, grand, sur fond
' bleu) ou de sous-groupe (niveau 2, souligné) sur toute la largeur.
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
            .Interior.Color = VERT_FONCE
        Else
            .Font.Size = 11
            .Font.Underline = xlUnderlineStyleSingle
            .Font.Color = VERT_FONCE
        End If
    End With
    mHaut = mHaut + 2
    mBas = mHaut
End Sub

' Place un tableau dans l'emplacement suivant de la rangée (3 par rangée)
Private Sub Placer(ws As Worksheet, titre As String, entetes As String, donnees As String, _
        note As String)
    Dim fin As Long
    fin = EcrireTableau(ws, mHaut, 2 + mEmpl * LARGEUR_EMPL, titre, entetes, donnees, note)
    If fin > mBas Then mBas = fin
    mEmpl = mEmpl + 1
    If mEmpl = NB_PAR_LIGNE Then
        mEmpl = 0
        mHaut = mBas
    End If
End Sub

' Nomme chaque colonne de données du dernier tableau écrit : prefixe_nom
' (ex. NF_A1_Technique), pour que les autres feuilles lisent les valeurs de la norme.
Private Sub NommerColonnes(prefixe As String, noms As Variant)
    Dim k As Long
    For k = 0 To UBound(noms)
        ActiveWorkbook.Names.Add Name:=prefixe & "_" & noms(k), RefersTo:="='" & _
            mDonnees.Worksheet.Name & "'!" & mDonnees.Columns(k + 1).Address
    Next k
End Sub

' Valeurs maximales du frottement axial unitaire limite qs en kPa (tableaux F.5.2.3
' et G.5.2.3, identiques), lignes des micropieux (catégories 17 à 20)
Private Function QsMax() As String
    QsMax = "17 - M1|-|-|-|-|-|-;18 - M2|-|-|-|-|-|-;19 - PIGU, MIGU|200|200|380|320|320|320;" & _
        "20 - PIRS, MIRS|200|200|440|440|440|500"
End Function

' Écrit un tableau à partir de la ligne r et de la colonne c0 : titre (gras,
' souligné, sur la largeur de l'emplacement), ligne d'en-tête
' (colonnes séparées par « | »), lignes de données (lignes séparées par « ; »,
' colonnes par « | », nombres avec un point décimal) puis une note éventuelle en
' italique. Renvoie la ligne libre suivante (une ligne vide après le tableau).
Private Function EcrireTableau(ws As Worksheet, ByVal r As Long, c0 As Long, titre As String, _
        entetes As String, donnees As String, note As String) As Long
    Dim e As Variant, lignes As Variant, cols As Variant, i As Long, j As Long, debut As Long

    With ws.Range(ws.Cells(r, c0), ws.Cells(r, c0 + LARGEUR_EMPL - 2))
        .Merge
        .Value = titre
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Color = VERT_FONCE
        .WrapText = True
        .VerticalAlignment = xlBottom
    End With
    If ws.Rows(r).RowHeight < 30 Then ws.Rows(r).RowHeight = 30   ' titre sur 2 lignes
    r = r + 1
    debut = r
    e = Split(entetes, "|")
    For j = 0 To UBound(e)
        ws.Cells(r, c0 + j).Value = e(j)
    Next j
    With ws.Range(ws.Cells(r, c0), ws.Cells(r, c0 + UBound(e)))
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .HorizontalAlignment = xlCenter
    End With
    lignes = Split(donnees, ";")
    For i = 0 To UBound(lignes)
        r = r + 1
        cols = Split(lignes(i), "|")
        For j = 0 To UBound(cols)
            If j > 0 And EstNombre(CStr(cols(j))) Then
                ws.Cells(r, c0 + j).Value = Val(cols(j))      ' Val lit le point décimal
            Else
                ws.Cells(r, c0 + j).NumberFormat = "@"         ' texte (« 1 bis », « - »)
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

' Vrai si le texte est un nombre écrit avec un point décimal (« 1.15 », « 90 »)
Private Function EstNombre(t As String) As Boolean
    Dim i As Long, c As String
    If t = "" Then Exit Function
    For i = 1 To Len(t)
        c = Mid(t, i, 1)
        If Not (c Like "[0-9]" Or c = ".") Then Exit Function
    Next i
    EstNombre = True
End Function
