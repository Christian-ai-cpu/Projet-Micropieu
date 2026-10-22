Attribute VB_Name = "TableauxNFP94262"
Option Explicit

' Module « Tableaux NF P94-262 » : crée (ou recrée) la feuille « Tableaux NF P94-262 »
' avec les tableaux de la norme NF P94-262 (juillet 2012) utiles au calcul des
' micropieux :
'   A.1               : classes et catégories de pieux, norme de référence
'   A.9.1             : quantité de coulis recommandée pour le scellement des micropieux
'   C.2.3.1 / C.2.3.2 : facteurs partiels de résistance ELU (durable / accidentel)
'   14.2.1.1 / 14.2.1.2 : facteurs partiels de résistance ELS
'   C.2.4.1 / C.2.4.2 : facteurs de corrélation xi'
'   F.2.1 / G.2.1     : coefficients de modèle (pressiomètre / pénétromètre)
'   F.4.2.1, F.5.2.1, F.5.2.2, F.5.2.3 : méthode pressiométrique (kp max,
'                       alpha pieu-sol, paramètres de fsol, qs max)
'   G.4.2.1, G.5.2.1, G.5.2.2, G.5.2.3 : méthode pénétrométrique (kc max,
'                       alpha pieu-sol, paramètres de fsol, qs max)
' Disposition : 3 groupes (généralités ; facteurs partiels, de corrélation et de
' modèle ; méthodes de calcul) et, dans chaque groupe, 3 tableaux côte à côte.
' « - » : pas de valeur dans la norme. Les notes des tableaux sont résumées sous
' chaque tableau ; se reporter à la norme pour leur texte complet.
' Plages nommées créées pour la feuille « Données du projet » : NF_A1_Classe,
' NF_A1_Categorie, NF_A1_Technique, NF_A1_Abreviation, NF_A1_Norme (tableau A.1) et
' NF_Sols (catégories de sol, en-têtes des tableaux qs max).

Private Const NOM_FEUILLE As String = "Tableaux NF P94-262"
' Colonnes de sols : 5 colonnes (pressiomètre, tableaux F.4.2.1, F.5.2.1, F.5.2.2)
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

Public Sub CreerTableauxNFP94262()
    Dim ws As Worksheet, s As String, g As String, c As Long

    Application.ScreenUpdating = False
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo 0
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
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    mHaut = 3
    mBas = 3
    mEmpl = 0

    NouveauGroupe ws, "1. GÉNÉRALITÉS", 1
    ' --- Tableau A.1 ---
    s = "1|1|Foré simple (pieux et barrettes)|FS|NF EN 1536;" & _
        "1|2|Foré boue (pieux et barrettes)|FB|NF EN 1536;" & _
        "1|3|Foré tubé (virole perdue)|FTP|NF EN 1536;" & _
        "1|4|Foré tubé (virole récupérée)|FTR|NF EN 1536;" & _
        "1|5|Foré simple ou boue avec rainurage ou puits|FSR, FBR, PU|NF EN 1536;" & _
        "2|6|Foré tarière continue simple rotation ou double rotation|FTC, FTCD|NF EN 1536;" & _
        "3|7|Vissé moulé|VM|NF EN 12699;3|8|Vissé tubé|VT|NF EN 12699;" & _
        "4|9|Battu béton préfabriqué ou précontraint|BPF, BPR|NF EN 12699;" & _
        "4|10|Battu enrobé (béton - mortier - coulis)|BE|NF EN 12699;" & _
        "4|11|Battu moulé|BM|NF EN 12699;4|12|Battu acier fermé|BAF|NF EN 12699;" & _
        "5|13|Battu acier ouvert|BAO|NF EN 12699;" & _
        "6|14|Profilé H battu|HB|NF EN 12699;6|15|Profilé H battu injecté|HBi|NF EN 12699;" & _
        "7|16|Palplanches battues|PP|NF EN 12699;" & _
        "1 bis|17|Micropieu type I|M1|NF EN 1536 / 14199 / 12699;" & _
        "1 bis|18|Micropieu type II|M2|NF EN 1536 / 14199 / 12699;" & _
        "8|19|Pieu ou micropieu injecté mode IGU (type III)|PIGU, MIGU|NF EN 1536 / 14199 / 12699;" & _
        "8|20|Pieu ou micropieu injecté mode IRS (type IV)|PIRS, MIRS|NF EN 1536 / 14199 / 12699"
    Placer ws, "Tableau A.1 - Classes et catégories de pieux", _
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
    s = "Pieux non ancrés dans la craie, classes 1 à 7 hors catégories 10 et 15|1.15|1.4|1.1|1.1;" & _
        "Pieux ancrés dans la craie, classes 1 à 7 hors catégories 10, 15, 17, 18, 19 et 20|1.4|1.7|1.1|1.1;" & _
        "Pieux de catégories 10, 15, 17, 18, 19 et 20|2|2|1.1|1.1"
    Placer ws, "Tableau F.2.1 - Coefficients de modèle - Méthode pressiométrique", _
        "Pieux|" & g & "R;d1 compression|" & g & "R;d1 traction|" & g & "R;d2 compression|" & _
        g & "R;d2 traction", s, g & "R;d1 : procédure du « pieu modèle » ; " & g & _
        "R;d2 : procédure du « modèle de terrain » (non utilisé pour le « pieu modèle »)."
    s = "Pieux non ancrés dans la craie, classes 1 à 7 hors catégories 10 et 15|1.18|1.45|1.1|1.1;" & _
        "Pieux ancrés dans la craie, classes 1 à 7 hors catégories 10, 15, 17, 18, 19 et 20|1.45|1.75|1.1|1.1;" & _
        "Pieux de catégories 10, 15, 17, 18, 19 et 20|2|2|1.1|1.1"
    Placer ws, "Tableau G.2.1 - Coefficients de modèle - Méthode pénétrométrique", _
        "Pieux|" & g & "R;d1 compression|" & g & "R;d1 traction|" & g & "R;d2 compression|" & _
        g & "R;d2 traction", s, ""

    NouveauGroupe ws, "3. MÉTHODES DE CALCUL", 1
    NouveauGroupe ws, "Méthode pressiométrique (annexe F)", 2
    ' --- Méthode pressiométrique ---
    s = "1|1.15|1.1|1.45|1.45|1.45;2|1.3|1.65|1.6|1.6|2;3|1.55|3.2|2.35|2.1|2.1;" & _
        "4|1.35|3.1|2.3|2.3|2.3;5|1|1.9|1.4|1.4|1.2;6|1.2|3.1|1.7|2.2|1.5;" & _
        "7|1|1|1|1|1.2;8|1.15|1.1|1.45|1.45|1.45"
    Placer ws, "Tableau F.4.2.1 - Facteur de portance pressiométrique kp max " & _
        "(encastrement relatif Def/B > 5)", "Classe de pieu|" & SOLS5, s, _
        "Pour les micropieux, la résistance de pointe n'est normalement pas prise en compte. " & _
        "Classes 5 à 7 vibrofoncées (BAO, HB, PP) : abattement de 50 % sur kp."
    s = "1 - FS|1.1|1|1.8|1.5|1.6;2 - FB|1.25|1.4|1.8|1.5|1.6;3 - FTP|0.7|0.6|0.5|0.9|-;" & _
        "4 - FTR|1.25|1.4|1.7|1.4|-;5 - FSR, FBR, PU|1.3|-|-|-|-;6 - FTC, FTCD|1.5|1.8|2.1|1.6|1.6;" & _
        "7 - VM|1.9|2.1|1.7|1.7|-;8 - VT|0.6|0.6|1|0.7|-;9 - BPF, BPR|1.1|1.4|1|0.9|-;" & _
        "10 - BE|2|2.1|1.9|1.6|-;11 - BM|1.2|1.4|2.1|1|-;12 - BAF|0.8|1.2|0.4|0.9|-;" & _
        "13 - BAO|1.2|0.7|0.5|1|1;14 - HB|1.1|1|0.4|1|0.9;15 - HBi|2.7|2.9|2.4|2.4|2.4;" & _
        "16 - PP|0.9|0.8|0.4|1.2|1.2;17 - M1|-|-|-|-|-;18 - M2|-|-|-|-|-;" & _
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
    s = "1|0.4|0.3|0.2|0.3|0.3|0.3;2|0.45|0.3|0.25|0.3|0.3|0.3;3|0.5|0.5|0.5|0.4|0.35|0.35;" & _
        "4|0.45|0.4|0.4|0.4|0.4|0.4;5|0.35|0.3|0.25|0.15|0.15|0.15;6|0.4|0.4|0.4|0.35|0.2|0.2;" & _
        "7|0.35|0.25|0.15|0.15|0.15|0.15;8|0.45|0.3|0.2|0.3|0.3|0.25"
    Placer ws, "Tableau G.4.2.1 - Facteur de portance pénétrométrique kc max " & _
        "(encastrement relatif Def/B > 5)", "Classe de pieu|" & SOLS6, s, _
        "Pour les micropieux, la résistance de pointe n'est normalement pas prise en compte. " & _
        "Classes 5 à 7 vibrofoncées (BAO, HB, PP) : abattement de 50 % sur kc."
    s = "1 - FS|0.55|0.65|0.7|0.8|1.4|1.5;2 - FB|0.65|0.8|1|0.8|1.4|1.5;" & _
        "3 - FTP|0.35|0.4|0.4|0.25|0.85|-;4 - FTR|0.65|0.8|1|0.75|0.13|-;" & _
        "5 - FSR, FBR, PU|0.7|0.85|-|-|-|-;6 - FTC, FTCD|0.75|0.9|1.25|0.95|1.5|1.5;" & _
        "7 - VM|0.95|1.15|1.45|0.75|1.6|-;8 - VT|0.3|0.35|0.4|0.45|0.65|-;" & _
        "9 - BPF, BPR|0.55|0.65|1|0.45|0.85|-;10 - BE|1|1.2|1.45|0.85|1.5|-;" & _
        "11 - BM|0.6|0.7|1|0.95|0.95|-;12 - BAF|0.4|0.5|0.85|0.2|0.85|-;" & _
        "13 - BAO|0.6|0.7|0.5|0.25|0.95|0.95;14 - HB|0.55|0.65|0.7|0.2|0.95|0.85;" & _
        "15 - HBi|1.35|1.6|2|1.1|2.25|2.25;16 - PP|0.45|0.55|0.55|0.2|1.25|1.15;" & _
        "17 - M1|-|-|-|-|-|-;18 - M2|-|-|-|-|-|-;19 - PIGU, MIGU|1.35|1.6|2|1.1|2.25|2.25;" & _
        "20 - PIRS, MIRS|1.7|2.05|2.65|1.4|2.9|2.9"
    Placer ws, "Tableau G.5.2.1 - Choix des valeurs de " & ChrW(945) & _
        " pieu-sol - Méthode pénétrométrique", "Catégorie|" & SOLS6, s, _
        "Catégories 17 et 18 : prendre les valeurs des techniques les plus proches. " & _
        "Valeur 0,13 (FTR, marne) reproduite telle qu'imprimée dans la norme."
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
    Application.ScreenUpdating = True
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
            .Interior.Color = RGB(47, 84, 150)
        Else
            .Font.Size = 11
            .Font.Underline = xlUnderlineStyleSingle
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
' et G.5.2.3, identiques)
Private Function QsMax() As String
    QsMax = "1 - FS|90|90|90|200|170|200;2 - FB|90|90|90|200|170|200;" & _
        "3 - FTP|50|50|50|50|90|-;4 - FTR|90|90|90|170|170|-;5 - FSR, FBR, PU|90|90|-|-|-|-;" & _
        "6 - FTC, FTCD|90|90|170|200|200|200;7 - VM|130|130|200|170|170|-;" & _
        "8 - VT|50|50|90|90|90|-;9 - BPF, BPR|130|130|130|90|90|-;" & _
        "10 - BE|170|170|260|200|200|-;11 - BM|90|90|130|260|200|-;" & _
        "12 - BAF|90|90|90|50|90|-;13 - BAO|90|90|50|50|90|90;14 - HB|90|90|130|50|90|90;" & _
        "15 - HBi|200|200|380|320|320|320;16 - PP|90|90|50|50|90|90;" & _
        "17 - M1|-|-|-|-|-|-;18 - M2|-|-|-|-|-|-;19 - PIGU, MIGU|200|200|380|320|320|320;" & _
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
        .Interior.Color = RGB(217, 225, 242)
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
