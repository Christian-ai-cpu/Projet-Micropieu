Attribute VB_Name = "CalculFondation"
Option Explicit

' Module « Calcul et vérification » : crée (ou recrée) la feuille « Calcul et
' vérification » de la fondation, à partir des feuilles « Données du projet »
' (module DonneesProjet) et « Tableaux NF P94-262 » (module TableauxNFP94262).
' Toutes les valeurs sont des formules : elles suivent les saisies de ces feuilles.
'   B3  : Données du projet (type de pieu, méthode de calcul)
'   B7  : Caractéristiques de la fondation, en 3 groupes séparés par 2 lignes vides :
'         abréviation, norme, classe, catégorie ; mode de fonctionnement (compression
'         / traction), résistance de pointe, type de mise en oeuvre ; caractéristiques
'         mécaniques (B, D, hauteur de la couche porteuse, Ab, Pe, frottement négligé)
'   B26 : Facteurs partiels de résistance ELU (C.2.3.1, C.2.3.2) et ELS (14.2.1.1,
'         14.2.1.2) ; les coefficients de modèle gamma R;d1 et R;d2 du calcul de qs,i
'         sont lus dans la feuille « Tableaux NF P94-262 » (F.2.1 ou G.2.1)
'   B37 : Frottement axial unitaire et résistance par couche (calculés si le nom et la
'         catégorie de la couche sont renseignés), notations de la norme : profondeur,
'         hi, pl*,i, courbe et a, b, c (F.5.2.2), alpha pieu-sol (F.5.2.1), fsol,i,
'         qs,max (F.5.2.3), qs = min(alpha x fsol,i ; qs,max), As,i = Pe x hi,
'         qs,i = qs / (gamma R;d1 x gamma R;d2), Rs,i = qs,i x As,i (0 si le terrain
'         est neutralisé)
'   B56 : valeurs caractéristiques (Rb,k, Rs,k, Rc,k, Rt,k, Rc;cr,k, Rt;cr,k) puis
'         valeurs de calcul ELU fondamental / accidentel et ELS caractéristique /
'         quasi permanent (compression et traction)
' Version MICROPIEUX (NF P94-262, catégories 17 à 20 : M1, M2, PIGU / MIGU, PIRS / MIRS) :
'   - résistance de pointe non prise en compte : Rb,k = 0, pas de tableaux d'ancrage
'     (ple*, qce, Def, kp, kc) ;
'   - tableaux de la norme limités aux micropieux : ligne 1 à 4 des tableaux alpha et
'     qs max = catégories 17 à 20 (rang = catégorie - 16), une seule ligne pour les
'     coefficients de modèle F.2.1 / G.2.1 (catégories 10, 15 et 17 à 20).
'   - pas de calcul en saisie libre en colonne A, à gauche du frottement axial (k = arrondi
'     supérieur de hi / pas ; vide : k = 1) ; lignes
'     du frottement axial limitées au nombre de couches de « Données du projet » ;
'     tableau de la capacité portante par pas de calcul (une ligne par pas, total des k)
'     sous les valeurs de calcul (couche, profondeur, cote, pas, qs,i, Rs,k, Rt,k,
'     ple*, h, kp max, Def, Rb,k puis Rc;cr,d / Rt;cr,d ELS et Rc,d / Rt,d ELU), valeurs
'     cumulées en profondeur (ligne précédente + qs,i x Pe x pas), et graphique de la
'     capacité portante (ELS caractéristique, ELS quasi permanent, ELU fondamental) en
'     fonction de la profondeur, à droite de ce tableau.
' Méthode pressiométrique (annexe F) : fsol = (a.pl* + b)(1 - exp(-c.pl*)), pl* en
' MPa, fsol en kPa ; qs = min(alpha pieu-sol x fsol ; qs max).
' Catégories de sol de la feuille « Données du projet » -> colonnes des tableaux :
'   tableaux à 5 colonnes (F.5.2.1, F.5.2.2) : argiles et sols intermédiaires -> 1re
'   colonne, sables et graves -> 2e, craies -> 3e, marnes -> 4e, roches -> 5e ;
'   tableau qs max à 6 colonnes : dans l'ordre de la liste des catégories de sol.

Private Const NOM_FEUILLE As String = "Calcul et vérification"
Private Const DONNEES As String = "'Données du projet'!"
' Pas de calcul : saisie libre pour chaque couche (colonne A, en vert, sans valeur
' imposée) ; nombre de subdivisions k calculé d'après le pas et l'épaisseur hi :
' k = arrondi supérieur de hi / pas (dernier pas de la couche réduit au reste) ;
' pas vide : k = 1 (pas = hi). NB_SUBDIV ne sert plus qu'aux anciennes feuilles.
Private Const NB_SUBDIV As Long = 20
' Lignes du tableau par couche du frottement axial (après déplacement de l'effet de
' groupe) : couche 1 en LIGNE_FROT, pas de calcul saisis en colonne A
Private Const LIGNE_FROT As Long = 48
' Le tableau de la capacité portante par pas de calcul a autant de lignes que le total
' des subdivisions k (nombre écrit en ZA, ligne LIGNE_PROFIL - 3) ; la vérification
' de la profondeur commence au moins NB_MIN_LIGNES lignes plus bas (graphiques)
Private Const NB_MIN_LIGNES As Long = 40
Private Const LIGNE_PROFIL As Long = 91        ' 1re ligne du tableau par pas de calcul
' Tableaux par pas de calcul des diamètres B0, B1, ... côte à côte : tableau de Bk à
' partir de la colonne 2 + k x BLOC_PROFIL (24 colonnes), graphique de B0 à droite du sien
Private Const BLOC_PROFIL As Long = 40
Private Const TABLEAUX As String = "Tableaux NF P94-262"
Private mCellule As String     ' cellule en cours d'écriture (indiquée en cas d'erreur)
Private mPas As Variant        ' pas de calcul saisis, conservés quand la feuille est recréée
Private mPhiManuel As Variant  ' Phi manuels saisis, conservés quand la vérification est recréée
Private mIncl As Variant, mAngle As Variant   ' inclinaison (Oui / Non) et angle saisis par micropieu
Private mGeo As Variant, mGsn As Variant, mAmax As Variant, mAmin As Variant   ' idem : Geo,
                               ' frottement négatif Gsn, action accidentelle A max et min
' Méthode de calcul lue dans « Données du projet » à la création de la feuille : les
' formules sont écrites pour cette seule méthode (pl* et tableaux F, ou qc et tableaux G)
' et la feuille est recréée quand la méthode change (ActualiserCalculVerification,
' appelée par ChangementDonnees du module DonneesProjet)
Private mPenetro As Boolean

' Thème vert : saisies en vert, valeurs non saisies (calculées) en jaune, en-têtes et
' titres en vert (couleurs RGB écrites en nombre : rouge + 256 x vert + 65536 x bleu)
Private Const VERT_SAISIE As Long = 13561798   ' RGB(198, 239, 206)
Private Const JAUNE_CALCUL As Long = 13434879  ' RGB(255, 255, 204), jaune clair
Private Const VERT_ENTETE As Long = 14348258   ' RGB(226, 239, 218)
Private Const VERT_FONCE As Long = 2315831     ' RGB(55, 86, 35)
Public Sub CreerCalculVerification()
    Construire True
End Sub

' Recrée la feuille « Calcul et vérification » (si elle existe) sans l'activer : appelée
' quand la méthode de calcul change dans « Données du projet »
Public Sub ActualiserCalculVerification()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo 0
    If Not ws Is Nothing Then Construire False
End Sub

Private Sub Construire(activer As Boolean)
    Dim ws As Worksheet, wt As Worksheet, r As Long, i As Long, lc As Long
    Dim zAlpha As String, zABC As String, zQs As String, nb As String, cat As String
    Dim lg As Long, etape As String, zF21 As String, ligneF21 As String, lt As Long, gamma As String
    Dim c As String, kd As Long, vb As Variant, evenements As Boolean, nPas As Long

    Application.ScreenUpdating = False
    evenements = Application.EnableEvents
    Application.EnableEvents = False         ' pas de ChangementCalcul pendant la création
    On Error GoTo Erreur
    mPenetro = (ActiveWorkbook.Worksheets("Données du projet").Range(CELLULE_METHODE).Value = _
        "Méthode pénétrométrique")

    ' Feuille des tableaux de la norme : créée si elle manque
    etape = "feuille Tableaux NF P94-262"
    On Error Resume Next
    Set wt = ActiveWorkbook.Worksheets(TABLEAUX)
    On Error GoTo Erreur
    If wt Is Nothing Then
        Application.Run "CreerTableauxNFP94262"
        Application.ScreenUpdating = False
        Set wt = ActiveWorkbook.Worksheets(TABLEAUX)
    End If
    ' Plages de valeurs des tableaux utilisés (adresses absolues sur cette feuille)
    ' Frottement : tableaux G (méthode pénétrométrique) et F, sous la forme « plage G|plage F »
    ' (lue par IndexNorme)
    ' Feuille créée avant le remplissage des catégories 17 et 18 (« - ») : recréée
    zAlpha = ZoneTableau(wt, "Tableau F.5.2.1", 0, 4, 5)
    If CStr(wt.Range(zAlpha).Cells(1, 1).Value) = "-" Then
        Application.Run "CreerTableauxNFP94262"
        Application.ScreenUpdating = False
    End If
    zAlpha = ZoneTableau(wt, "Tableau G.5.2.1", 0, 4, 6) & "|" & _
        ZoneTableau(wt, "Tableau F.5.2.1", 0, 4, 5)         ' catégories 17 à 20
    zABC = ZoneTableau(wt, "Tableau G.5.2.2", 0, 4, 6) & "|" & _
        ZoneTableau(wt, "Tableau F.5.2.2", 0, 4, 5)         ' courbe, a, b, c
    zQs = ZoneTableau(wt, "Tableau G.5.2.3", 0, 4, 6) & "|" & _
        ZoneTableau(wt, "Tableau F.5.2.3", 0, 4, 6)         ' qs max, catégories 17 à 20

    etape = "création de la feuille"
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo Erreur
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = NOM_FEUILLE
    End If
    mPhiManuel = LirePhiManuel(ws)           ' Phi manuels de la vérification, conservés
    LireActions ws
    mPas = ws.Range("A" & LIGNE_FROT & ":A" & LIGNE_FROT + MAX_COUCHES - 1).Formula   ' pas saisis
    ws.Cells.Clear
    Do While ws.ChartObjects.Count > 0       ' graphique de la création précédente
        ws.ChartObjects(1).Delete
    Loop

    nb = DONNEES & "$" & Replace(CELLULE_NB_COUCHES, "C", "C$")   ' nombre de couches
    lc = LIGNE_COUCHES                       ' première couche dans « Données du projet »
    lg = LIGNE_COUCHES + MAX_COUCHES - 1     ' dernière ligne prévue

    With ws.Range("B1:I1")
        .Merge
        .Value = "CALCUL ET VÉRIFICATION DES MICROPIEUX"
        .Font.Bold = True
        .Font.Color = VERT_FONCE
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With

    ' --- Données du projet ---
    etape = "données du projet"
    SousTitre ws, 3, "Données du projet"
    Ligne ws, 4, "Type de pieu", "=" & DONNEES & "D" & LIGNE_PIEU, ""
    Ligne ws, 5, "Méthode de calcul", "=" & DONNEES & CELLULE_METHODE, ""

    ' --- Caractéristiques de la fondation : 3 groupes séparés par 2 lignes vides ---
    etape = "caractéristiques de la fondation"
    SousTitre ws, 7, "Caractéristiques de la fondation"
    Ligne ws, 8, "Abréviation", "=" & DONNEES & "D" & LIGNE_PIEU + 11, ""
    Ligne ws, 9, "Norme de référence", "=" & DONNEES & "D" & LIGNE_PIEU + 8, ""
    Ligne ws, 10, "Classe de la fondation", "=" & DONNEES & "D" & LIGNE_PIEU + 9, ""
    Ligne ws, 11, "Catégorie de la fondation", "=" & DONNEES & "D" & LIGNE_PIEU + 10, ""
    cat = "$D$11"                            ' catégorie de la fondation (cette feuille)

    Ligne ws, 14, "Mode de fonctionnement", "=" & DONNEES & "D" & LIGNE_PIEU + 12, ""
    Ligne ws, 15, "Prise en compte de la résistance de pointe", "=" & DONNEES & "D" & LIGNE_PIEU + 13, ""
    Ligne ws, 16, "Type de mise en oeuvre", "=" & DONNEES & "D" & LIGNE_PIEU + 14, ""

    Ligne ws, 19, "Diamètre du micropieu B (B0 de « Données du projet »)", "=" & DONNEES & "D" & _
        LIGNE_PIEU + 15, "m"
    ' D = somme des épaisseurs des couches du tableau du frottement axial (lignes 39 à 53)
    Ligne ws, 20, "Hauteur totale de la fondation D (somme des épaisseurs)", _
        "=IF(COUNT($F$39:$F$53)=0,"""",SUM($F$39:$F$53))", "m"
    ' h = épaisseur de la dernière couche renseignée du tableau du frottement axial
    Ligne ws, 21, "Hauteur de la couche porteuse h", Derniere("F"), "m"
    Ligne ws, 22, "Aire de la base de la fondation Ab = " & ChrW(960) & ".B² / 4", _
        "=IF(N(D19)=0,"""",PI()*D19^2/4)", "m²"
    Ligne ws, 23, "Périmètre de la fondation Pe = " & ChrW(960) & ".B", _
        "=IF(N(D19)=0,"""",PI()*D19)", "m"
    ' Épaisseur des couches dont le terrain est neutralisé (frottement négligé)
    Ligne ws, 24, "Frottement négligé (épaisseur des terrains neutralisés)", _
        "=IF(COUNTIF(OFFSET(" & DONNEES & "$M$" & lc & ",0,0," & nb & "),""Oui"")=0,""Non""," & _
        "SUMIF(OFFSET(" & DONNEES & "$M$" & lc & ",0,0," & nb & "),""Oui""," & _
        "OFFSET(" & DONNEES & "$F$" & lc & ",0,0," & nb & ")))", "m"
    ws.Range("D19:D24").NumberFormat = "0.000"

    ' --- Facteurs partiels de résistance (NF P94-262) : ELU (tableaux C.2.3.1 et
    ' C.2.3.2) et ELS (tableaux 14.2.1.1 et 14.2.1.2), utilisés pour les valeurs de calcul
    etape = "facteurs partiels"
    SousTitre ws, 26, "Facteurs partiels de résistance (NF P94-262)"
    ws.Range("B27:E27").Value = Array("ELU (tableaux C.2.3.1 et C.2.3.2)", "", _
        "Situations durables et transitoires", "Situations accidentelles")
    ws.Range("B28:E28").Value = Array("Pointe " & ChrW(947) & "b", "", 1.1, 1)
    ws.Range("B29:E29").Value = Array("Fût en compression " & ChrW(947) & "s", "", 1.1, 1)
    ws.Range("B30:E30").Value = Array("Total en compression " & ChrW(947) & "t", "", 1.1, 1)
    ws.Range("B31:E31").Value = Array("Fût en traction " & ChrW(947) & "s;t", "", 1.15, 1.05)
    ws.Range("B33:E33").Value = Array("ELS (tableaux 14.2.1.1 et 14.2.1.2)", "", _
        "Combinaisons caractéristiques", "Combinaisons quasi permanentes")
    ws.Range("B34:E34").Value = Array("Fût en compression " & ChrW(947) & "cr", "", 0.9, 1.1)
    ws.Range("B35:E35").Value = Array("Fût en traction " & ChrW(947) & "s;cr", "", 1.1, 1.5)
    For i = 27 To 35
        If i <> 32 Then ws.Range("B" & i & ":C" & i).Merge
    Next i
    ws.Range("D28:E31,D34:E35").NumberFormat = "0.00"
    ' Coefficients de modèle gamma R;d1 et R;d2 lus dans la feuille « Tableaux NF P94-262 »
    ' (tableau F.2.1 ou G.2.1 selon la méthode de calcul) : ligne unique des micropieux
    ' (catégories 10, 15 et 17 à 20) ; colonnes compression / traction selon le mode de
    ' fonctionnement
    zF21 = ZoneTableau(wt, IIf(mPenetro, "Tableau G.2.1", "Tableau F.2.1"), 0, 1, 4)
    ligneF21 = "1"
    gamma = "INDEX(" & zF21 & "," & ligneF21 & ",IF($D$14=""Traction"",2,1))*INDEX(" & zF21 & _
        "," & ligneF21 & ",IF($D$14=""Traction"",4,3))"

    ' --- Frottement axial unitaire par couche (notations de la NF P94-262) ---
    etape = "frottement par couche"
    r = 37
    SousTitre ws, r, "Frottement axial unitaire et résistance par couche"
    ws.Range("B" & r).Formula = "=""Frottement axial unitaire et résistance par couche (""&LOWER($D$5)&"")""
    ws.Range("B" & r + 1 & ":Q" & r + 1).Value = Array("N°", "Nom de couche", "Catégorie de sol", _
        "Profondeur (m/TN)", "hi (m)", "pl*,i (MPa)", "Courbe fsol", ChrW(945) & "pieu-sol", _
        "fsol,i (kPa)", ChrW(945) & "pieu-sol x fsol,i (kPa)", "qs,max (kPa)", "qs = min (kPa)", _
        "As,i = " & ChrW(960) & ".B.hi (m²)", "qs,i (kPa)", "Rs,i (kN)", "Terrain neutralisé")
    ' Unité de la profondeur : celle de « Données du projet » (m + référentiel altimétrique)
    ws.Range("E" & r + 1).Formula = "=""Profondeur (""&" & DONNEES & "$E$29&"")"""
    ' En-têtes selon la méthode de calcul : pl* et tableaux F, ou qc et tableaux G
    If mPenetro Then
        ws.Range("G" & r + 1 & ":J" & r + 1).Value = Array("qc,i (MPa)", "Courbe fsol (G.5.2.2)", _
            ChrW(945) & "pieu-sol (G.5.2.1)", "fsol,i = (a.qc + b)(1 - e^-c.qc) (kPa)")
        ws.Range("L" & r + 1).Value = "qs,max (G.5.2.3) (kPa)"
    Else
        ws.Range("G" & r + 1 & ":J" & r + 1).Value = Array("pl*,i (MPa)", "Courbe fsol (F.5.2.2)", _
            ChrW(945) & "pieu-sol (F.5.2.1)", "fsol,i = (a.pl* + b)(1 - e^-c.pl*) (kPa)")
        ws.Range("L" & r + 1).Value = "qs,max (F.5.2.3) (kPa)"
    End If
    For i = 1 To MAX_COUCHES
        CoucheLigne ws, r + 1 + i, i, lc + i - 1, "(" & cat & "-16)", gamma, zAlpha, zABC, zQs
    Next i
    mCellule = ""
    lt = r + 1 + MAX_COUCHES                 ' dernière ligne du tableau

    ' --- Sous le frottement axial : valeurs caractéristiques puis valeurs de calcul ---
    ' Intitulé en B, symbole en C, valeur en D, unité en E
    etape = "valeurs caractéristiques"
    SousTitre ws, 56, "Valeurs caractéristiques des résistances"
    ' Rb,k = 0 : résistance de pointe non prise en compte pour les micropieux
    LigneS ws, 57, "Résistance de pointe (non prise en compte pour les micropieux)", "Rb,k", "=0", "kN"
    LigneS ws, 58, "Résistance de frottement = " & ChrW(931) & " Rs,i", "Rs,k", _
        "=IF(COUNT($P$39:$P$53)=0,"""",SUM($P$39:$P$53))", "kN"
    LigneS ws, 60, "Portance (compression) = Rb,k + Rs,k", "Rc,k", _
        "=IF(AND($D$14=""Compression"",ISNUMBER(D58)),N(D57)+D58,"""")", "kN"
    LigneS ws, 61, "Résistance en traction = Rs,k", "Rt,k", "=IF(ISNUMBER(D58),D58,"""")", "kN"
    LigneS ws, 62, "Charge de fluage en compression = 0,5 Rb,k + 0,7 Rs,k", "Rc;cr,k", _
        "=IF(AND($D$14=""Compression"",ISNUMBER(D58)),0.5*N(D57)+0.7*D58,"""")", "kN"
    LigneS ws, 63, "Charge de fluage en traction = 0,7 Rt,k", "Rt;cr,k", _
        "=IF(ISNUMBER(D61),0.7*D61,"""")", "kN"
    ws.Range("D57:D63").NumberFormat = "0.0"

    ' Valeurs de calcul : 2 lignes par combinaison (compression puis traction), avec
    ' combinaison en B (fusionnée), intitulé en C, symbole en D, valeur en E (avec l'unité)
    '   ELU : Rc,d = Rb,k / gamma b + Rs,k / gamma s ; Rt,d = Rt,k / gamma s;t
    '   ELS : Rc;cr,d = Rc;cr,k / gamma cr ; Rt;cr,d = Rt;cr,k / gamma s;cr
    etape = "valeurs de calcul"
    SousTitre ws, 65, "Valeurs de calcul des résistances"
    ws.Range("B66:E66").Value = Array("Combinaison", "Valeur de calcul", "Symbole", "Valeur")
    For i = 0 To 3
        r = 67 + 2 * i
        c = Chr(Asc("D") + (i Mod 2))        ' colonne du facteur : D (durable / caract.), E
        ws.Range("B" & r).Value = Array("ELU fondamental (C.2.3.1)", "ELU accidentel (C.2.3.2)", _
            "ELS caractéristique (14.2.1.1)", "ELS quasi permanent (14.2.1.2)")(i)
        ws.Range("B" & r & ":B" & r + 1).Merge
        If i < 2 Then
            ws.Range("C" & r & ":D" & r + 1).Value = Array2("Portance de calcul en compression", _
                "Rc,d", "Résistance de calcul en traction", "Rt,d")
            ws.Range("E" & r).Formula = "=IF(D60="""","""",N(D57)/" & c & "28+D58/" & c & "29)"
            ws.Range("E" & r + 1).Formula = "=IF(D61="""","""",D61/" & c & "31)"
        Else
            ws.Range("C" & r & ":D" & r + 1).Value = Array2("Charge de fluage de calcul en compression", _
                "Rc;cr,d", "Charge de fluage de calcul en traction", "Rt;cr,d")
            ws.Range("E" & r).Formula = "=IF(D62="""","""",D62/" & c & "34)"
            ' sans essai de contrôle : Rt;cr,d x 0,15
            ws.Range("E" & r + 1).Formula = "=IF(D63="""","""",D63/" & c & "35*IF($D$76=""Non"",0.15,1))"
        End If
    Next i
    ws.Range("E67:E74").NumberFormat = "0.0"" kN""
    Ligne ws, 76, "Essai de contrôle (sinon Rt;cr,d x 0,15)", "=" & DONNEES & "D" & LIGNE_PIEU + 19, ""

    ' --- Effet de groupe lié au rapprochement des pieux (NF P94-262, J.2) ---
    ' Ce = 1 si d >= 3B ; sinon Ce = 1 - Cd (2 - (1/m + 1/n)) avec Cd = 1 - (1 + d/B) / 4
    ' pour 1 <= d/B <= 3 (« - » si d/B < 1) ; Ce = 1 sans effet de groupe
    etape = "effet de groupe"
    SousTitre ws, 78, "Effet de groupe - coefficient d'efficacité Ce (NF P94-262, J.2)"
    Ligne ws, 79, "Effet de groupe", "=" & DONNEES & "D" & LIGNE_PIEU + 4, ""
    Ligne ws, 80, "Nombre de lignes de pieux m", "=IF(" & DONNEES & "D" & LIGNE_PIEU + 5 & "=""""," & _
        """""," & DONNEES & "D" & LIGNE_PIEU + 5 & ")", ""
    Ligne ws, 81, "Nombre de pieux par ligne n", "=IF(" & DONNEES & "D" & LIGNE_PIEU + 6 & "=""""," & _
        """""," & DONNEES & "D" & LIGNE_PIEU + 6 & ")", ""
    Ligne ws, 82, "Entraxe des pieux d", "=IF(" & DONNEES & "D" & LIGNE_PIEU + 7 & "=""""," & _
        """""," & DONNEES & "D" & LIGNE_PIEU + 7 & ")", "m"
    Ligne ws, 83, "d / B", "=IF(OR(N(D82)=0,N($D$19)=0),"""",D82/$D$19)", ""
    Ligne ws, 84, "Cd = 1 - (1 + d/B) / 4  (1 <= d/B <= 3)", _
        "=IF(D83="""","""",IF(D83>=3,0,IF(D83>=1,1-(1+D83)/4,""-"")))", ""
    Ligne ws, 85, "Ce = 1 si d >= 3B, sinon 1 - Cd (2 - (1/m + 1/n))", _
        "=IF(D79<>""Oui"",1,IF(D83="""","""",IF(D83>=3,1,IF(D83<1,""-""," & _
        "IF(OR(N(D80)=0,N(D81)=0),"""",1-D84*(2-(1/D80+1/D81)))))))", ""
    ws.Range("D80:D81").NumberFormat = "0"
    ws.Range("D82:D85").NumberFormat = "0.00"
    ws.Range("B85:D85").Font.Bold = True


    ' --- Pas de calcul (colonne A, à gauche du frottement axial) et capacité portante
    ' en fonction de la profondeur (tableau sous les valeurs de calcul) ---
    ' Pas saisi librement (colonne A) ; k = arrondi supérieur de hi / pas, pour toutes
    ' les couches (1 si le pas est vide). Colonnes d'aide masquées (ZA à ZC, ZF) :
    ' cumul de k, profondeur du bas de la couche, cumul de Rs,i, pas effectif.
    ' Chaque ligne du tableau correspond à un pas : couche (ZD) et rang du pas dans la
    ' couche (ZE), masqués ; z = bas de la couche précédente + rang x pas ;
    ' Rs,k(z) = Rs,i des couches au-dessus + qs,i x Pe x (z - haut de la couche) ;
    ' Rt,k = Rs,k ; micropieux : ple*, kp max et Def sans objet (« - »), Rb,k = 0 ;
    ' valeurs de calcul comme dans le tableau des valeurs de calcul
    etape = "pas de calcul"
    ws.Range("A38").Value = "Pas de calcul (m) (saisie libre)"
    ws.Range("R38").Value = "Nombre de subdivisions k = hi / pas"
    ws.Range("ZA38:ZC38").Value = Array(ChrW(931) & " k", "z bas (m)", ChrW(931) & " Rs,i (kN)")
    ws.Range("ZF38").Value = "Pas effectif (m)"
    For r = 39 To 53
        ' Pas saisi conservé (valeur seulement, pas les anciennes formules)
        If IsArray(mPas) Then
            vb = mPas(r - 38, 1)
            If CStr(vb) <> "" And Left(CStr(vb), 1) <> "=" Then
                If ValeurPas(vb) > 0 Then ws.Range("A" & r).Value = ValeurPas(vb)
            End If
        End If
        ' Pas saisi lu en nombre (ZG) même s'il est écrit en texte (« 0.5 » ou « 0,5 »)
        ws.Range("ZG" & r).Formula = "=IF(ISNUMBER(A" & r & "),A" & r & ",IFERROR(--SUBSTITUTE(A" & r & _
            ","","",""."")," & "IFERROR(--SUBSTITUTE(A" & r & ",""."","",""),0)))"
        ' k = arrondi supérieur de hi / pas (1 si le pas est vide)
        ws.Range("R" & r).Formula = "=IF(B" & r & "="""","""",IF(N(ZG" & r & ")>0," & _
            "MAX(1,ROUNDUP(ROUND(N(F" & r & ")/ZG" & r & ",6),0)),1))"
        ' Pas effectif : pas saisi (au plus hi), hi si le pas est vide
        ws.Range("ZF" & r).Formula = "=IF(B" & r & "="""","""",IF(N(ZG" & r & ")>0,MIN(ZG" & r & _
            ",N(F" & r & ")),N(F" & r & ")))"
        ws.Range("ZA" & r).Formula = "=IF(B" & r & "="""","""",SUM($R$39:R" & r & "))"
        ws.Range("ZB" & r).Formula = "=IF(B" & r & "="""","""",SUM($F$39:F" & r & "))"
        ws.Range("ZC" & r).Formula = "=IF(B" & r & "="""","""",SUM($P$39:P" & r & "))"
    Next r
    ws.Range("A39:A53").NumberFormat = "General"   ' saisie libre
    ws.Range("A39:A53").Validation.Delete
    ws.Range("R39:R53").NumberFormat = "0"
    With ws.Range("A38:A53,R38:R53")
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A38,R38")
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
    End With
    ws.Range("A39:A53").Interior.Color = VERT_SAISIE
    ws.Range("R39:R53").Interior.Color = JAUNE_CALCUL
    mPas = Empty

    ' Nombre total de pas (somme des k) : nombre de lignes du tableau de la capacité
    ' portante, écrit en ZA (ligne LIGNE_PROFIL - 3) pour les autres procédures
    ws.Calculate
    nPas = Application.Max(ws.Range("ZA39:ZA53"))
    If nPas < 1 Then nPas = 1
    ws.Range("ZA" & LIGNE_PROFIL - 3).Value = nPas

    etape = "capacité portante par pas de calcul"
    ' Valeurs cumulées en profondeur, pas par pas : ligne précédente + apport du pas
    '   Rs,k(z) = Rs,k(z - pas) + qs x As,i / (gamma R;d1 compression x gamma R;d2),
    '   Rt,k(z) = Rt,k(z - pas) + qs x As,i / (gamma R;d1 traction x gamma R;d2),
    '   coefficients de modèle repris dans le tableau après le pas de calcul,
    '   As,i = pi x Phi x pas (Phi = B0), 0 si le terrain est neutralisé ;
    '   valeurs de calcul déduites de Rs,k et Rt,k cumulés (mêmes relations que le
    '   tableau des valeurs de calcul)
    lt = LIGNE_PROFIL + nPas - 1             ' dernière ligne : une ligne par pas
    SousTitre ws, LIGNE_PROFIL - 3, "Capacité portante des micropieux en fonction de la " & _
        "profondeur (par pas de calcul, valeurs cumulées) - diamètre B0"
    r = LIGNE_PROFIL - 2
    ws.Range("B" & r).Value = "Références suivant couche"
    ws.Range("R" & r).Value = "ELS caractéristique (14.2.1.1)"
    ws.Range("T" & r).Value = "ELS quasi permanent (14.2.1.2)"
    ws.Range("V" & r).Value = "ELU fondamental (C.2.3.1)"
    ws.Range("X" & r).Value = "ELU accidentel (C.2.3.2)"
    ws.Range("B" & r & ":Q" & r).Merge
    For i = 0 To 3
        ws.Range(Chr(Asc("R") + 2 * i) & r & ":" & Chr(Asc("S") + 2 * i) & r).Merge
    Next i
    ws.Range("B" & r + 1 & ":Y" & r + 1).Value = Array("N° de couche", "Profondeur z (m)", "Cote", _
        ChrW(934) & " (m)", "Pas de calcul (m)", ChrW(947) & "R;d1 compression", _
        ChrW(947) & "R;d1 traction", ChrW(947) & "R;d2", "qs (kPa)", "Rs,k (kN)", "Rt,k (kN)", _
        "ple* (MPa)", "h (m)", "kp max", "Def (m)", "Rb,k (kN)", "Rc;cr,d (kN)", "Rt;cr,d (kN)", _
        "Rc;cr,d (kN)", "Rt;cr,d (kN)", "Rc,d (kN)", "Rt,d (kN)", "Rc,d (kN)", "Rt,d (kN)")
    ws.Range("D" & r + 1).Formula = "=""Cote (""&" & DONNEES & "$E$29&"")"""
    ' Formules de la 1re ligne (j = rang du pas = ligne - LIGNE_PROFIL + 1), recopiées
    ' vers le bas ; la ligne précédente de la 1re ligne est l'en-tête (texte -> N() = 0)
    r = LIGNE_PROFIL
    c = "(ROW()-" & LIGNE_PROFIL - 1 & ")"                   ' rang j du pas
    ws.Range("$ZD" & r).Formula = "=IF(OR(COUNT($ZA$39:$ZA$53)=0," & c & ">MAX($ZA$39:$ZA$53)),""""," & _
        "COUNTIF($ZA$39:$ZA$53,""<""&" & c & ")+1)"
    ws.Range("$ZE" & r).Formula = "=IF($ZD" & r & "="""",""""," & c & "-IF($ZD" & r & "=1,0," & _
        "INDEX($ZA$39:$ZA$53,$ZD" & r & "-1)))"
    ws.Range("B" & r).Formula = "=IF($ZD" & r & "="""","""",INDEX($B$39:$B$53,$ZD" & r & "))"
    ' Diamètre Phi du micropieu : B0 de « Données du projet » (Bk pour les tableaux copiés)
    ws.Range("E" & r).Formula = "=IF($ZD" & r & "="""",""""," & DONNEES & "$G$" & LIGNE_PIEU + 9 & ")"
    ' Pas de la ligne : h(z) - h(z précédent) dans la couche (dernier pas réduit au reste)
    ws.Range("F" & r).Formula = "=IF($ZD" & r & "="""","""",N" & r & "-IF($ZE" & r & "=1,0,N(N" & r - 1 & ")))"
    ' Coefficients de modèle (tableau F.2.1 ou G.2.1, ligne des micropieux) : gamma R;d1
    ' compression (colonne 1), traction (colonne 2), gamma R;d2 (colonne 3)
    ws.Range("G" & r).Formula = "=IF($ZD" & r & "="""","""",INDEX(" & zF21 & ",1,1))"
    ws.Range("H" & r).Formula = "=IF($ZD" & r & "="""","""",INDEX(" & zF21 & ",1,2))"
    ws.Range("I" & r).Formula = "=IF($ZD" & r & "="""","""",INDEX(" & zF21 & ",1,3))"
    ' h = profondeur dans la couche = min(rang x pas effectif ; hi)
    ws.Range("N" & r).Formula = "=IF($ZD" & r & "="""","""",MIN($ZE" & r & "*INDEX($ZF$39:$ZF$53,$ZD" & r & _
        "),INDEX($F$39:$F$53,$ZD" & r & ")))"
    ws.Range("C" & r).Formula = "=IF($ZD" & r & "="""","""",IF($ZD" & r & "=1,0,INDEX($ZB$39:$ZB$53,$ZD" & _
        r & "-1))+N" & r & ")"
    ws.Range("D" & r).Formula = "=IF(C" & r & "="""","""",IF(ISNUMBER(" & DONNEES & "$D$29)," & _
        DONNEES & "$D$29-C" & r & ",-C" & r & "))"
    ' qs de la couche (frottement axial unitaire limite, colonne « qs = min »)
    ws.Range("J" & r).Formula = "=IF($ZD" & r & "="""","""",INDEX($M$39:$M$53,$ZD" & r & "))"
    ' Rs,k et Rt,k cumulés : ligne précédente + qs x As,i / (gamma R;d1 x gamma R;d2), avec
    ' As,i = pi x Phi x pas, gamma R;d1 de compression pour Rs,k, de traction pour Rt,k
    ' (0 si le terrain est neutralisé)
    ws.Range("K" & r).Formula = "=IF($ZD" & r & "="""","""",N(K" & r - 1 & ")+IF(INDEX($Q$39:$Q$53,$ZD" & _
        r & ")=""Oui"",0,N(J" & r & ")*PI()*E" & r & "*F" & r & "/(G" & r & "*I" & r & ")))"
    ws.Range("L" & r).Formula = "=IF($ZD" & r & "="""","""",N(L" & r - 1 & ")+IF(INDEX($Q$39:$Q$53,$ZD" & _
        r & ")=""Oui"",0,N(J" & r & ")*PI()*E" & r & "*F" & r & "/(H" & r & "*I" & r & ")))"
    ws.Range("M" & r).Formula = "=IF($ZD" & r & "="""","""",""-"")"
    ws.Range("O" & r).Formula = "=M" & r
    ws.Range("P" & r).Formula = "=M" & r
    ws.Range("Q" & r).Formula = "=IF($ZD" & r & "="""","""",0)"
    ' ELS : Rc;cr,d = (0,5 Rb,k + 0,7 Rs,k) / gamma cr ; Rt;cr,d = 0,7 Rt,k / gamma s;cr
    ' (x 0,15 sans essai de contrôle) ; ELU : Rc,d = Rb,k / gamma b + Rs,k / gamma s ;
    ' Rt,d = Rt,k / gamma s;t
    For i = 0 To 1
        c = Chr(Asc("D") + i)                ' colonne du facteur : D ou E
        ws.Range(Chr(Asc("R") + 2 * i) & r).Formula = "=IF(K" & r & "="""","""",(0.5*Q" & r & _
            "+0.7*K" & r & ")/$" & c & "$34)"
        ws.Range(Chr(Asc("S") + 2 * i) & r).Formula = "=IF(L" & r & "="""","""",0.7*L" & r & "/$" & c & _
            "$35*IF($D$76=""Non"",0.15,1))"
        ws.Range(Chr(Asc("V") + 2 * i) & r).Formula = "=IF(K" & r & "="""","""",Q" & r & "/$" & c & _
            "$28+K" & r & "/$" & c & "$29)"
        ws.Range(Chr(Asc("W") + 2 * i) & r).Formula = "=IF(L" & r & "="""","""",L" & r & "/$" & c & "$31)"
    Next i
    ws.Range("B" & r & ":$ZE" & lt).FillDown
    ws.Range("C" & r & ":D" & lt & ",N" & r & ":N" & lt).NumberFormat = "0.00"
    ws.Range("E" & r & ":F" & lt).NumberFormat = "0.000"
    ws.Range("G" & r & ":I" & lt).NumberFormat = "0.00"
    ws.Range("J" & r & ":L" & lt & ",Q" & r & ":Y" & lt).NumberFormat = "0.0"
    With ws.Range("B" & LIGNE_PROFIL - 2 & ":Y" & lt)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("B" & LIGNE_PROFIL - 2 & ":Y" & LIGNE_PROFIL - 1)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .WrapText = True
    End With
    ws.Columns("ZA:ZG").Hidden = True

    ' Tableaux des diamètres B1, B2, B3 de « Données du projet » : copie du
    ' tableau de B0 décalée de k x BLOC_PROFIL colonnes, mêmes calculs avec Phi = Bk
    etape = "capacité portante des autres diamètres"
    For kd = 1 To NB_DIAMETRES - 1
        If True Then                         ' tableau toujours créé (vide si Bk n'est pas saisi)
            ws.Range(ws.Cells(LIGNE_PROFIL - 3, 2), ws.Cells(lt, 25)).Copy _
                Destination:=ws.Cells(LIGNE_PROFIL - 3, 2 + kd * BLOC_PROFIL)
            ws.Range(ws.Cells(LIGNE_PROFIL, 2 + kd * BLOC_PROFIL), ws.Cells(lt, 25 + kd * BLOC_PROFIL)).Replace _
                What:="$G$" & LIGNE_PIEU + 9, Replacement:="$G$" & LIGNE_PIEU + 9 + kd, _
                LookAt:=xlPart, MatchCase:=False
            ws.Cells(LIGNE_PROFIL - 3, 2 + kd * BLOC_PROFIL).Value = "Capacité portante des micropieux " & _
                "en fonction de la profondeur (par pas de calcul, valeurs cumulées) - diamètre B" & kd
            ws.Range(ws.Columns(2 + kd * BLOC_PROFIL), ws.Columns(25 + kd * BLOC_PROFIL)).ColumnWidth = 12
        End If
    Next kd
    Application.CutCopyMode = False

    ' --- Mise en forme ---
    etape = "mise en forme"
    r = 37
    lt = r + 1 + MAX_COUCHES
    With ws.Range("B4:E5,B8:E11,B14:E16,B19:E24,B27:E31,B33:E35,B" & r + 1 & ":Q" & lt & _
            ",B57:E58,B60:E63,B66:E74,B76:E76,B79:E85")
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("B27:E27,B33:E33,B66:E66,B" & r + 1 & ":Q" & r + 1)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = VERT_ENTETE
    End With
    ' Valeurs non saisies (reprises ou calculées) en jaune ; tableau du frottement axial
    ' sans couleur
    ws.Range("D4:D5,D8:D11,D14:D16,D19:D24,D28:E31,D34:E35,D57:D58,D60:D63,E67:E74,D76,D79:D85").Interior.Color = _
        JAUNE_CALCUL
    ws.Range("D4:F35,C57:E63,B67:B74,D67:E74,D76:E85,B" & r + 2 & ":B" & lt & ",E" & r + 2 & ":Q" & lt _
        ).HorizontalAlignment = xlCenter
    ws.Columns("A").ColumnWidth = 12        ' pas de calcul
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:Y").ColumnWidth = 12
    ' Bloc de l'effet de groupe (écrit en lignes 78 à 86) déplacé avant le tableau du
    ' frottement axial par couche (stratigraphie) : il occupe les lignes 37 à 45 et la
    ' suite des colonnes A à ZE descend de 9 lignes (jusqu'à la ligne 86 ; le tableau par
    ' pas de calcul, plus bas, ne bouge pas), les formules suivent le déplacement
    etape = "déplacement de l'effet de groupe"
    ws.Range("A78:ZG86").Cut
    ws.Range("A37").Insert Shift:=xlDown
    Application.CutCopyMode = False

    etape = "vérification de la profondeur des micropieux"
    VerificationProfondeur ws

    ' Graphiques créés une fois les largeurs de colonnes fixées
    etape = "graphique de la capacité portante"
    GraphiqueCapacite ws, 0
    For kd = 1 To NB_DIAMETRES - 1           ' graphique de chaque diamètre B1, B2, ... créé
        If ws.Cells(LIGNE_PROFIL, 3 + kd * BLOC_PROFIL).HasFormula Then GraphiqueCapacite ws, kd
    Next kd
    AfficherGroupe ws
    AfficherCouchesCalcul ws
    InstallerCodeFeuille ws
    If activer Then ws.Activate

Fin:
    Application.EnableEvents = evenements
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    MsgBox "Erreur pendant la création de la feuille « Calcul et vérification », partie « " & _
        etape & " »" & IIf(mCellule = "", "", ", cellule " & mCellule) & " : " & Err.Description & _
        " (n° " & Err.Number & ")", vbExclamation
    Resume Fin
End Sub

' Bloc de l'effet de groupe (lignes 37 à 45, avant le frottement axial) masqué quand
' « Effet de groupe » vaut Non dans « Données du projet ». Appelée à la création
' et quand ce choix change (ChangementDonnees, module DonneesProjet).
Public Sub AfficherGroupe(Optional ws As Worksheet)
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
    End If
    ws.Rows("37:45").Hidden = (ActiveWorkbook.Worksheets("Données du projet").Range("D" & _
        LIGNE_PIEU + 4).Value <> "Oui")
End Sub

' Vérification de la profondeur des micropieux, sous le tableau de la capacité portante
' par pas de calcul : pour chaque micropieu de la feuille de descente de charges (matrice
' des micropieux, module MatriceMicropieux, titre en B80) :
'   descente de charges : Gmax = G, Gmin = G ou 0 (choix), Qmax = Q, Qmin = 0 ; actions
'   saisies : poussée des terres Geo (min = max), frottement négatif Gsn, action
'   accidentelle A max et min ; eau EB, EH, EE (kN, avec nappe seulement) ;
'   combinaisons verticales max / min :
'     ELS caractéristique : G + Q ;  ELS quasi permanent : G + psi2 Q ;
'     ELU fondamental : 1,35 Gmax + 1,5 Qmax (max), 1,0 Gmin + 1,5 Qmin (min) ;
'     ELU accidentel : G + psi1 Q ; psi1, psi2 selon la catégorie de bâtiment ;
'   diamètres : Phi de chaque combinaison (plus petit diamètre du tableau des diamètres
'   de « Données du projet » dont la portance couvre la charge, portance proportionnelle
'   au diamètre), Phi pieu (le plus grand), Phi manuel (saisi), Phi retenu (le plus grand
'   des deux) ;
'   profondeur minimale requise pour chaque cas : 1re profondeur du tableau par pas de
'   calcul où la portance de la même combinaison (Rc;cr,d ELS, Rc,d ELU), ramenée au
'   diamètre Phi retenu, est supérieure ou égale à la charge (« Non atteinte » sinon) ;
'   profondeur retenue : la plus grande des profondeurs requises du micropieu ;
'   vérification de la longueur : profondeur de reconnaissance (« Données du projet »,
'   D33) - profondeur retenue >= 5 Phi retenu (NF P94-262, annexe M, M.1 (2) Note 2).
' Le nombre de micropieux et de murs est lu à la création : recréer cette feuille
' après un changement dans la descente de charges.
' Recrée le tableau de vérification de la profondeur des micropieux (si la feuille
' « Calcul et vérification » existe), le choix de Gmin étant conservé. Appelée quand le
' nombre de murs ou de micropieux change dans la descente de charges
' (ConstruireMatrice, module MatriceMicropieux).
Public Sub ActualiserVerificationProfondeur()
    Dim ws As Worksheet, L As Long, gmin As String
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    L = LigneVerif(ws)
    gmin = CStr(ws.Range("D" & L + 2).Value)
    mPhiManuel = LirePhiManuel(ws)
    LireActions ws
    ws.Range(ws.Rows(L), ws.Rows(L + 300)).Clear
    VerificationProfondeur ws, gmin
End Sub

' Phi manuels saisis dans la vérification de la profondeur (colonne « Phi MANUEL »,
' lignes MP1...)
Private Function LirePhiManuel(ws As Worksheet) As Variant
    LirePhiManuel = LireSaisie(ws, ChrW(934) & " MANUEL")
End Function

' Valeurs saisies d'une colonne du tableau de vérification de la profondeur (en-tête
' entete sur l'une des 2 lignes d'en-tête au-dessus de MP1), lignes MP1... ; vide si la
' colonne n'existe pas
Private Function LireSaisie(ws As Worksheet, entete As String) As Variant
    Dim r As Long, i As Long, t() As Variant, c As Long, cm As Long
    r = LigneVerif(ws) + 10                  ' ligne de MP1
    For c = 1 To 60
        If CStr(ws.Cells(r - 2, c).Value) = entete Or CStr(ws.Cells(r - 1, c).Value) = entete Then cm = c
    Next c
    If cm = 0 Then Exit Function
    Do While Left(CStr(ws.Range("A" & r + i).Value), 2) = "MP"
        i = i + 1
    Loop
    If i = 0 Then Exit Function
    ReDim t(1 To i)
    For i = 1 To UBound(t)
        t(i) = ws.Cells(r + i - 1, cm).Value
    Next i
    LireSaisie = t
End Function

' Actions saisies de la vérification de la profondeur (Geo max, Gsn, A max, A min),
' conservées quand elle est recréée
Private Sub LireActions(ws As Worksheet)
    mGeo = LireSaisie(ws, "Geo max")
    mGsn = LireSaisie(ws, "Gsn")
    mAmax = LireSaisie(ws, "A max")
    mAmin = LireSaisie(ws, "A min")
    mIncl = LireSaisie(ws, "Incliné")
    mAngle = LireSaisie(ws, "Angle (°)")
End Sub

' Valeurs saisies de la ligne i (1 = MP1) d'une colonne lue par LireSaisie ; 0 sinon
Private Function Saisie(t As Variant, i As Long) As Variant
    Saisie = 0
    If IsArray(t) Then
        If i <= UBound(t) Then
            If Not IsEmpty(t(i)) Then Saisie = t(i)
        End If
    End If
End Function

Private Sub VerificationProfondeur(ws As Worksheet, Optional gmin As String = "G")
    Dim wd As Worksheet, f As Worksheet, L As Long, n As Long, m As Long, i As Long, k As Long
    Dim r0 As Long, rd As Long, rt As Long, r As Long, dn As String, cap As String, z As String
    Dim charge As Variant, colCap As Variant, entete As Variant, phi As String, c As Long
    Dim eau As Boolean, cC As Long, cD As Long, cP As Long, cR As Long, cL As Long, cF As Long
    Dim reco As String, eMax As String, eMin As String, eAcc As String, phiT As String
    Dim incl As Boolean, cRe As Long, kb As String, capk As String, zk As String, cZ As Long
    Dim cI As Long, cA As Long, angle0 As Variant
    Dim fx As String, kd As Long, cE As Long, tC As Long, tP As Long

    L = LigneVerif(ws)                       ' titre du tableau
    For Each f In ActiveWorkbook.Worksheets
        If CStr(f.Range("B80").Value) = "RÉPARTITION DES CHARGES DES MURS SUR LES MICROPIEUX" Then
            Set wd = f
            Exit For
        End If
    Next f

    With ws.Range("A" & L)
        .Value = "VÉRIFICATION DE LA PROFONDEUR DES MICROPIEUX (DESCENTE DE CHARGES)"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = VERT_FONCE
    End With
    Ligne ws, L + 1, "Feuille de descente de charges", IIf(wd Is Nothing, "introuvable", ""), ""
    If Not wd Is Nothing Then ws.Range("D" & L + 1).Value = wd.Name
    Ligne ws, L + 2, "Gmin pris égal à (G ou 0)", "", ""
    With ws.Range("D" & L + 2)
        .Validation.Add xlValidateList, xlValidAlertStop, , "G,0"
        .Value = IIf(gmin = "0", "0", "G")
        .Interior.Color = VERT_SAISIE
    End With
    Ligne ws, L + 3, "Catégorie du bâtiment", "=" & DONNEES & "C14", ""
    Ligne ws, L + 4, ChrW(968) & "1 (combinaison accidentelle)", "=IFERROR(CHOOSE(MATCH(LEFT(D" & _
        L + 3 & ",1),{""A"",""B"",""C"",""D"",""E"",""F"",""G""},0),0.5,0.5,0.7,0.7,0.9,0.7,0.5),0.5)", ""
    Ligne ws, L + 5, ChrW(968) & "2 (combinaison quasi permanente)", "=IFERROR(CHOOSE(MATCH(LEFT(D" & _
        L + 3 & ",1),{""A"",""B"",""C"",""D"",""E"",""F"",""G""},0),0.3,0.3,0.6,0.6,0.8,0.6,0.3),0.3)", ""
    ws.Range("D" & L + 1 & ":D" & L + 5).HorizontalAlignment = xlCenter
    ws.Range("D" & L + 3 & ":D" & L + 5).Interior.Color = JAUNE_CALCUL
    ws.Range("D" & L + 4 & ":D" & L + 5).NumberFormat = "0.00"
    With ws.Range("B" & L + 1 & ":E" & L + 5)
        .Borders.LineStyle = xlContinuous
        .Font.Bold = True
    End With

    reco = DONNEES & "$D$33"
    colCap = Array("R", "T", "V", "X")
    phiT = DONNEES & "$G$" & LIGNE_PIEU + 9 & ":$G$" & LIGNE_PIEU + 8 + NB_DIAMETRES   ' B0, B1, ...
    If wd Is Nothing Then
        ws.Range("A" & L + 7).Value = "Créez la matrice des micropieux dans la feuille de descente de " & _
            "charges, puis relancez CreerCalculVerification."
        Exit Sub
    End If

    ' Matrice des micropieux : murs (Mur 01...) à partir de la ligne 86, en-tête à 88 + n,
    ' micropieux (MP1...) à partir de 91 + n, POIDS TOTAL ÉLÉMENTAIRE DDC G et Q en
    ' colonnes 8 + n et 9 + n, eau EB, EH, EE en colonnes 10 + n à 12 + n
    Do While Left(CStr(wd.Range("B" & 86 + n).Value), 4) = "Mur "
        n = n + 1
    Loop
    rd = 88 + n + 3
    Do While Left(CStr(wd.Range("B" & rd + m).Value), 2) = "MP"
        m = m + 1
    Loop
    dn = "'" & wd.Name & "'!"

    ' Charges de l'eau (EB basses eaux, EH hautes eaux, EE eaux exceptionnelles) : colonnes
    ' créées seulement s'il y a une nappe phréatique (« Données du projet », D34) et si la
    ' matrice de la descente de charges les contient (colonnes 10 + n à 12 + n)
    eau = (CStr(ActiveWorkbook.Worksheets("Données du projet").Range("D34").Value) = "Oui") And _
        (Left(CStr(wd.Cells(88 + n, 10 + n).Value), 15) = "CHARGES DE L'EA")
    ' Colonnes : A micropieu ; B à E Gmax, Gmin, Qmax, Qmin ; F à J Geo max, Geo min, Gsn,
    ' A max, A min (saisis) (+ 3 colonnes d'eau) ; cC
    ' combinaisons Vmax / Vmin ; cD diamètres (Phi CARA, QP, FOND, ACC, Phi pieu, Phi
    ' manuel, Phi retenu) ; cP profondeurs Dmax / Dmin ; cR profondeur retenue D ; cL
    ' vérification de la longueur (profondeur de reconnaissance - D >= 5 Phi retenu)
    ' Actions : Geo max, Geo min (= Geo max) en F, G ; frottement négatif Gsn en H ;
    ' action accidentelle A max, A min en I, J ; eau éventuelle en K à M
    cE = 11                                   ' 1re colonne de l'eau
    cC = cE + IIf(eau, 3, 0)
    cD = cC + 8
    cP = cD + 7
    ' Inclinaison de chaque micropieu (cI, liste Oui / Non) avant la profondeur retenue
    ' (cR) ; après celle-ci, angle sur la verticale (cA, à saisir) et profondeur réelle
    ' D / cos(angle) (cRe), colonnes affichées seulement si un micropieu est incliné
    ' (AfficherInclinaison). Valeurs par défaut : « Incliné » et angle de « Données du
    ' projet ».
    cI = cP + 8
    cR = cI + 1
    cA = cR + 1
    cRe = cR + 2
    cZ = cR + 3                               ' altitudes Ztête et Zpointe
    cL = cR + 5
    cF = cL + 3                               ' dernière colonne
    incl = (CStr(ActiveWorkbook.Worksheets("Données du projet").Range("D" & LIGNE_PIEU + 2).Value) = "Oui")
    angle0 = ActiveWorkbook.Worksheets("Données du projet").Range("D" & LIGNE_PIEU + 3).Value
    ws.Range(ws.Columns(cP), ws.Columns(cF + 3)).Hidden = False
    r0 = L + 7
    ws.Cells(r0, 1).Value = "N° MICROPIEU"
    ws.Cells(r0, 2).Value = "DESCENTE DE CHARGES ET ACTIONS (kN)"
    ws.Cells(r0, cC).Value = "COMBINAISONS DE CHARGES VERTICALES (kN)"
    ws.Cells(r0, cD).Value = "DIAMÈTRE DU MICROPIEU (m)"
    ws.Cells(r0, cP).Value = "PROFONDEUR MINIMALE REQUISE (m)"
    ws.Cells(r0, cR).Value = "PROFONDEUR RETENUE (m)"
    ws.Cells(r0, cL).Value = "VÉRIFICATION DE LA LONGUEUR (profondeur de reconnaissance - D >= 5 " & _
        ChrW(934) & ")"
    ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + 2, 1)).Merge
    ws.Range(ws.Cells(r0, 2), ws.Cells(r0, cC - 1)).Merge
    ws.Range(ws.Cells(r0, cC), ws.Cells(r0, cD - 1)).Merge
    ws.Range(ws.Cells(r0, cD), ws.Cells(r0, cP - 1)).Merge
    ws.Range(ws.Cells(r0, cP), ws.Cells(r0, cR - 1)).Merge
    ws.Range(ws.Cells(r0, cR), ws.Cells(r0 + 1, cR)).Merge
    ws.Cells(r0, cI).Value = "INCLINAISON"
    ws.Cells(r0 + 1, cI).Value = "Incliné"
    ws.Range(ws.Cells(r0 + 1, cI), ws.Cells(r0 + 2, cI)).Merge
    ws.Cells(r0, cA).Value = "INCLINAISON"
    ws.Cells(r0 + 1, cA).Value = "Angle (°)"
    ws.Range(ws.Cells(r0 + 1, cA), ws.Cells(r0 + 2, cA)).Merge
    ws.Cells(r0, cRe).Value = "PROFONDEUR RÉELLE (m)"
    ws.Range(ws.Cells(r0, cRe), ws.Cells(r0 + 1, cRe)).Merge
    ws.Cells(r0 + 2, cRe).Value = "D / cos " & ChrW(945)
    ws.Cells(r0, cZ).Formula = "=""ALTITUDES (""&" & DONNEES & "$E$29&"")"""
    ws.Range(ws.Cells(r0, cZ), ws.Cells(r0, cZ + 1)).Merge
    ws.Cells(r0 + 1, cZ).Value = "Ztête"
    ws.Cells(r0 + 1, cZ + 1).Value = "Zpointe"
    ws.Range(ws.Cells(r0 + 1, cZ), ws.Cells(r0 + 2, cZ)).Merge
    ws.Range(ws.Cells(r0 + 1, cZ + 1), ws.Cells(r0 + 2, cZ + 1)).Merge
    ws.Range(ws.Cells(r0, cL), ws.Cells(r0, cF)).Merge
    ws.Range("B" & r0 + 1).Value = "Gstr"
    ws.Range("D" & r0 + 1).Value = "Q"
    ws.Range("B" & r0 + 1 & ":C" & r0 + 1).Merge
    ws.Range("D" & r0 + 1 & ":E" & r0 + 1).Merge
    ws.Range("B" & r0 + 2 & ":E" & r0 + 2).Value = Array("Gstr max", "Gstr min", "Qmax", "Qmin")
    ws.Range("F" & r0 + 1).Value = "Geo (poussée des terres)"
    ws.Range("F" & r0 + 1 & ":G" & r0 + 1).Merge
    ws.Range("F" & r0 + 2 & ":G" & r0 + 2).Value = Array("Geo max", "Geo min")
    ws.Range("H" & r0 + 1).Value = "Gsn"
    ws.Range("H" & r0 + 1 & ":H" & r0 + 2).Merge
    ws.Range("I" & r0 + 1).Value = "A (accidentel)"
    ws.Range("I" & r0 + 1 & ":J" & r0 + 1).Merge
    ws.Range("I" & r0 + 2 & ":J" & r0 + 2).Value = Array("A max", "A min")
    If eau Then
        ws.Range("K" & r0 + 1).Value = "EAU"
        ws.Range("K" & r0 + 1 & ":M" & r0 + 1).Merge
        ws.Range("K" & r0 + 2 & ":M" & r0 + 2).Value = Array("Eau EB", "Eau EH", "Eau EE")
    End If
    ws.Range(ws.Cells(r0 + 1, cD), ws.Cells(r0 + 1, cD + 6)).Value = Array(ChrW(934) & " CARA", _
        ChrW(934) & " QP", ChrW(934) & " FOND", ChrW(934) & " ACC", ChrW(934) & " PIEU", _
        ChrW(934) & " MANUEL", ChrW(934) & " RETENU")
    ws.Cells(r0 + 2, cR).Value = "D"
    ws.Range(ws.Cells(r0 + 1, cL), ws.Cells(r0 + 1, cF)).Value = Array("Profondeur de reconnaissance (m)", _
        "Profondeur de reconnaissance - D réelle (m)", _
        "5 x " & ChrW(934) & " retenu (m)", "Vérification")
    For k = 0 To 6
        ws.Range(ws.Cells(r0 + 1, cD + k), ws.Cells(r0 + 2, cD + k)).Merge
    Next k
    For k = 0 To 3
        ws.Range(ws.Cells(r0 + 1, cL + k), ws.Cells(r0 + 2, cL + k)).Merge
    Next k
    entete = Array("ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
    For k = 0 To 3
        For i = 0 To 1                       ' combinaisons (V), profondeurs (D)
            c = IIf(i = 0, cC, cP) + 2 * k
            ws.Cells(r0 + 1, c).Value = entete(k)
            ws.Range(ws.Cells(r0 + 1, c), ws.Cells(r0 + 1, c + 1)).Merge
            ws.Range(ws.Cells(r0 + 2, c), ws.Cells(r0 + 2, c + 1)).Value = _
                IIf(i = 0, Array("Vmax", "Vmin"), Array("Dmax", "Dmin"))
        Next i
    Next k

    ' Commentaires des en-têtes : signification des abréviations
    Commentaire ws.Range("B" & r0 + 1), "Gstr : charges permanentes de la structure (descente de " & _
        "charges : poids des murs, planchers, massifs, longrines...)"
    Commentaire ws.Range("B" & r0 + 2), "Gstr max : charge permanente de la structure (valeur " & _
        "défavorable, coefficient " & ChrW(947) & "G,sup)"
    Commentaire ws.Range("C" & r0 + 2), "Gstr min : charge permanente favorable (G ou 0 selon le " & _
        "choix en tête du tableau, coefficient " & ChrW(947) & "G,inf)"
    Commentaire ws.Range("D" & r0 + 1), "Q : charges d'exploitation (descente de charges)"
    Commentaire ws.Range("D" & r0 + 2), "Qmax : charge d'exploitation défavorable"
    Commentaire ws.Range("E" & r0 + 2), "Qmin : charge d'exploitation favorable, prise égale à 0"
    Commentaire ws.Range("F" & r0 + 1), "Geo : action géotechnique permanente (poussée des terres) " & _
        "transmise au micropieu"
    Commentaire ws.Range("F" & r0 + 2), "Geo max : valeur de la poussée des terres (à saisir)"
    Commentaire ws.Range("G" & r0 + 2), "Geo min : prise égale à Geo max"
    Commentaire ws.Range("H" & r0 + 1), "Gsn : frottement négatif (action permanente due au " & _
        "tassement du sol autour du micropieu), compté aux combinaisons max"
    Commentaire ws.Range("I" & r0 + 1), "A : action accidentelle (choc, explosion...), " & _
        "combinaison ELU accidentelle seulement"
    Commentaire ws.Range("I" & r0 + 2), "A max : action accidentelle de la combinaison max"
    Commentaire ws.Range("J" & r0 + 2), "A min : action accidentelle de la combinaison min"
    If eau Then
        Commentaire ws.Range("K" & r0 + 1), "Charges de l'eau (nappe phréatique)"
        Commentaire ws.Range("K" & r0 + 2), "Eau EB : niveau des basses eaux"
        Commentaire ws.Range("L" & r0 + 2), "Eau EH : niveau des hautes eaux"
        Commentaire ws.Range("M" & r0 + 2), "Eau EE : niveau des eaux exceptionnelles"
    End If
    entete = Array("ELS caractéristique (NF EN 1990, 6.14b)", "ELS quasi permanent (6.16b)", _
        "ELU fondamental (6.10)", "ELU accidentel (6.11b)")
    For k = 0 To 3
        ' capture de l'expression de la combinaison (NF EN 1990)
        NoteNorme ws.Cells(r0 + 1, cC + 2 * k), CStr(Choose(k + 1, "en1990_6-14b.png", "en1990_6-16b.png", _
            "en1990_6-10.png", "en1990_6-11b.png")), "Combinaison " & entete(k) & " : coefficients " & _
            "dans le tableau des combinaisons sous la vérification"
        Commentaire ws.Cells(r0 + 2, cC + 2 * k), "Vmax : charge verticale de calcul maximale " & _
            "(actions défavorables)"
        Commentaire ws.Cells(r0 + 2, cC + 2 * k + 1), "Vmin : charge verticale de calcul minimale " & _
            "(actions favorables)"
        Commentaire ws.Cells(r0 + 1, cD + k), ChrW(934) & " " & Split("CARA,QP,FOND,ACC", ",")(k) & _
            " : plus petit diamètre (B0 à B3) dont la portance maximale couvre la charge de la " & _
            "combinaison " & entete(k)
        Commentaire ws.Cells(r0 + 2, cP + 2 * k), "Dmax : profondeur minimale pour reprendre Vmax"
        Commentaire ws.Cells(r0 + 2, cP + 2 * k + 1), "Dmin : profondeur minimale pour reprendre Vmin"
    Next k
    Commentaire ws.Cells(r0 + 1, cD + 4), ChrW(934) & " PIEU : le plus grand des diamètres des " & _
        "combinaisons"
    Commentaire ws.Cells(r0 + 1, cD + 5), ChrW(934) & " MANUEL : diamètre imposé (à saisir)"
    Commentaire ws.Cells(r0 + 1, cD + 6), ChrW(934) & " RETENU : max(" & ChrW(934) & " pieu ; " & _
        ChrW(934) & " manuel)"
    Commentaire ws.Cells(r0 + 2, cR), "D : profondeur retenue, la plus grande des profondeurs requises"
    Commentaire ws.Cells(r0 + 1, cI), "Incliné : Oui si le micropieu est incliné (liste) ; les " & _
        "colonnes de l'angle et de la profondeur réelle s'affichent alors"
    Commentaire ws.Cells(r0 + 1, cA), "Angle d'inclinaison du micropieu sur la verticale (à saisir)"
    Commentaire ws.Cells(r0 + 2, cRe), "Profondeur réelle : longueur du micropieu " & _
        "incliné = D / cos " & ChrW(945) & " (" & ChrW(945) & " : angle sur la verticale) ; D si " & _
        "le micropieu n'est pas incliné"
    Commentaire ws.Cells(r0 + 1, cZ), "Ztête : altitude de la tête du micropieu = cote de l'arase " & _
        "supérieure de la fondation (« Données du projet »)"
    Commentaire ws.Cells(r0 + 1, cZ + 1), "Zpointe : altitude de la pointe du micropieu = Ztête - D " & _
        "(profondeur retenue)"
    Commentaire ws.Cells(r0 + 1, cL), "Profondeur de reconnaissance (« Données du projet »)"
    Commentaire ws.Cells(r0 + 1, cL + 2), "5 " & ChrW(934) & " : longueur reconnue sous la base " & _
        "(NF P94-262, annexe M, M.1)"

    ' Portance de chaque combinaison dans le tableau par pas de calcul : Rc;cr,d ELS
    ' caractéristique (R), ELS quasi permanent (T), Rc,d ELU fondamental (V), accidentel (X),
    ' lue dans le tableau de chaque diamètre B0, B1, B2, B3.
    ' Phi de chaque combinaison : plus petit diamètre (B0 < B1 < B2 < B3) dont la portance
    ' maximale (tableau de ce diamètre) atteint la plus grande charge (Vmax ou Vmin) ;
    ' Phi pieu = le plus grand des 4 ; Phi retenu = max(Phi pieu ; Phi manuel saisi).
    ' Profondeur Dmax / Dmin de chaque cas : 1re profondeur du tableau de Phi retenu où
    ' la portance est >= la charge ; D = la plus grande ; profondeur réelle D / cos(angle)
    ' pour un micropieu incliné.
    ' Eau ajoutée avec un coefficient 1 : EH aux combinaisons max, EB aux combinaisons
    ' min, EE à l'ELU accidentel max
    ' --- Coefficients des combinaisons (NF EN 1990), sous le tableau de vérification ---
    ' Tableau des combinaisons (lignes tC + 2 à tC + 5 : ELS CARA, ELS QP, ELU FOND, ELU ACC)
    '   Vmax = gamma G,sup x (Gmax + Geo + Gsn) + gamma Q x psi x Qmax + gamma A x A max
    '   Vmin = gamma G,inf x (Gmin + Geo min) + gamma Q x psi x Qmin + gamma A x A min
    ' coefficients gamma des tableaux A1.2(B) (6.10), A1.3 (6.11b) et A1.4 (6.14b, 6.16b) ;
    ' psi du tableau A1.1 (valeurs recommandées) selon la catégorie du bâtiment
    tC = r0 + 2 + Application.Max(m, 1) + 3
    tP = tC + 8
    With ws.Range("A" & tC)
        .Value = "COEFFICIENTS DES COMBINAISONS D'ACTIONS (NF EN 1990)"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
    ws.Range("B" & tC + 1 & ":I" & tC + 1).Value = Array("Combinaison", "Expression (NF EN 1990)", _
        ChrW(947) & "G,sup (Gstr max, Geo, Gsn)", ChrW(947) & "G,inf (Gstr min, Geo min)", ChrW(947) & "Q", _
        ChrW(968) & " appliqué à Q", "Valeur de " & ChrW(968), ChrW(947) & "A (accidentel)")
    ws.Range("B" & tC + 2 & ":G" & tC + 2).Value = Array("ELS caractéristique", "6.14b (tableau A1.4)", _
        1, 1, 1, "- (Q dominante)")
    ws.Range("B" & tC + 3 & ":G" & tC + 3).Value = Array("ELS quasi permanent", "6.16b (tableau A1.4)", _
        1, 1, 1, ChrW(968) & "2")
    ws.Range("B" & tC + 4 & ":G" & tC + 4).Value = Array("ELU fondamental", "6.10 (tableau A1.2(B))", _
        1.35, 1, 1.5, "- (Q dominante)")
    ws.Range("B" & tC + 5 & ":G" & tC + 5).Value = Array("ELU accidentel", "6.11b (tableau A1.3)", _
        1, 1, 1, ChrW(968) & "1")
    ws.Range("H" & tC + 2).Value = 1
    ws.Range("H" & tC + 3).Formula = "=IFERROR(VLOOKUP(LEFT($D$" & L + 3 & ",1),$B$" & tP + 2 & ":$F$" & _
        tP + 9 & ",5,FALSE),0.3)"
    ws.Range("H" & tC + 4).Value = 1
    ws.Range("H" & tC + 5).Formula = "=IFERROR(VLOOKUP(LEFT($D$" & L + 3 & ",1),$B$" & tP + 2 & ":$F$" & _
        tP + 9 & ",4,FALSE),0.5)"
    ws.Range("I" & tC + 2 & ":I" & tC + 5).Value = Application.Transpose(Array(0, 0, 0, 1))
    ' Tableau A1.1 : coefficients psi pour les bâtiments (charges d'exploitation)
    ws.Range("A" & tP).Value = "Tableau A1.1 - Valeurs recommandées des coefficients " & ChrW(968) & _
        " pour les bâtiments (NF EN 1990)"
    ws.Range("A" & tP).Font.Bold = True
    NoteNorme ws.Range("A" & tP), "en1990_tableau_A1-1.png", ""
    ws.Range("B" & tP + 1 & ":F" & tP + 1).Value = Array("Catégorie", "Charges d'exploitation", _
        ChrW(968) & "0", ChrW(968) & "1", ChrW(968) & "2")
    ws.Range("B" & tP + 2 & ":F" & tP + 9).Value = Array2D(Array( _
        Array("A", "Habitation, zones résidentielles", 0.7, 0.5, 0.3), _
        Array("B", "Bureaux", 0.7, 0.5, 0.3), _
        Array("C", "Lieux de réunion", 0.7, 0.7, 0.6), _
        Array("D", "Commerces", 0.7, 0.7, 0.6), _
        Array("E", "Stockage", 1, 0.9, 0.8), _
        Array("F", "Zone de trafic, véhicules <= 30 kN", 0.7, 0.7, 0.6), _
        Array("G", "Zone de trafic, véhicules de 30 à 160 kN", 0.7, 0.5, 0.3), _
        Array("H", "Toits", 0, 0, 0)))
    With ws.Range("B" & tC + 1 & ":I" & tC + 5 & ",B" & tP + 1 & ":F" & tP + 9)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("B" & tC + 1 & ":I" & tC + 1 & ",B" & tP + 1 & ":F" & tP + 1)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
    End With
    ws.Range("D" & tC + 2 & ":I" & tC + 5 & ",D" & tP + 2 & ":F" & tP + 9).NumberFormat = "0.00"
    ws.Range("D" & tC + 2 & ":I" & tC + 5).Interior.Color = JAUNE_CALCUL
    ' psi en tête de la vérification : repris du tableau des combinaisons
    ws.Range("D" & L + 4).Formula = "=$H$" & tC + 5
    ws.Range("D" & L + 5).Formula = "=$H$" & tC + 3

    z = "$C$" & LIGNE_PROFIL & ":$C$" & LIGNE_PROFIL + NbPas(ws) - 1
    For i = 1 To m
        r = r0 + 2 + i
        ws.Range("A" & r).Value = "MP" & i
        ws.Range("B" & r).Formula = "=N(" & dn & wd.Cells(rd + i - 1, 8 + n).Address & ")"
        ws.Range("C" & r).Formula = "=IF($D$" & L + 2 & "=""G"",B" & r & ",0)"
        ws.Range("D" & r).Formula = "=N(" & dn & wd.Cells(rd + i - 1, 9 + n).Address & ")"
        ws.Range("E" & r).Value = 0
        ' Actions saisies (0 par défaut, conservées quand le tableau est recréé)
        ws.Range("F" & r).Value = Saisie(mGeo, i)
        ws.Range("G" & r).Formula = "=F" & r                    ' Geo min = Geo max
        ws.Range("H" & r).Value = Saisie(mGsn, i)
        ws.Range("I" & r).Value = Saisie(mAmax, i)
        ws.Range("J" & r).Value = Saisie(mAmin, i)
        If eau Then
            For k = 0 To 2                   ' eau EB, EH, EE de la descente de charges
                ws.Cells(r, cE + k).Formula = "=N(" & dn & wd.Cells(rd + i - 1, 10 + n + k).Address & ")"
            Next k
            eMax = "+L" & r: eMin = "+K" & r: eAcc = "+M" & r
        Else
            eMax = "": eMin = "": eAcc = ""
        End If
        ' Combinaisons (NF EN 1990 et NF P94-262) : permanentes G + Geo + Gsn (Gsn
        ' défavorable, compté seulement aux combinaisons max)
        '   ELS caractéristique : G + Geo + Gsn + Q
        '   ELS quasi permanent : G + Geo + Gsn + psi2 Q
        '   ELU fondamental     : 1,35 (Gmax + Geo + Gsn) + 1,5 Q (max) ;
        '                         1,0 (Gmin + Geo) + 1,5 Qmin (min)
        '   ELU accidentel      : G + Geo + Gsn + A + psi1 Q (A max au max, A min au min)
        For k = 0 To 3                       ' coefficients : ligne tC + 2 + k
            c = tC + 2 + k
            ws.Cells(r, cC + 2 * k).Formula = "=$D$" & c & "*(B" & r & "+F" & r & "+H" & r & ")+$F$" & c & _
                "*$H$" & c & "*D" & r & "+$I$" & c & "*I" & r & IIf(k = 3, eAcc, eMax)
            ws.Cells(r, cC + 2 * k + 1).Formula = "=$E$" & c & "*(C" & r & "+G" & r & ")+$F$" & c & _
                "*$H$" & c & "*E" & r & "+$I$" & c & "*J" & r & eMin
        Next k
        For k = 0 To 3
            ' Phi de la combinaison : charge verticale (plus grande de Vmax et Vmin) comparée
            ' à la portance maximale de chaque diamètre (tableau de B0, B1, ...) ; le plus
            ' petit diamètre qui vérifie la condition est retenu
            charge = "MAX(" & Ad(ws, r, cC + 2 * k) & "," & Ad(ws, r, cC + 2 * k + 1) & ")"
            fx = """Insuffisant"""
            For kd = NB_DIAMETRES - 1 To 0 Step -1
                cap = ws.Range(ws.Cells(LIGNE_PROFIL, ws.Range(colCap(k) & "1").Column + kd * BLOC_PROFIL), _
                    ws.Cells(LIGNE_PROFIL + NbPas(ws) - 1, _
                    ws.Range(colCap(k) & "1").Column + kd * BLOC_PROFIL)).Address
                fx = "IF(AND(ISNUMBER(" & DONNEES & "$G$" & LIGNE_PIEU + 9 + kd & "),MAX(" & cap & ")>=" & _
                    charge & ")," & DONNEES & "$G$" & LIGNE_PIEU + 9 + kd & "," & fx & ")"
            Next kd
            ws.Cells(r, cD + k).Formula = "=IF(COUNT(" & phiT & ")=0,""""," & fx & ")"
        Next k
        ws.Cells(r, cD + 4).Formula = "=IF(COUNTIF(" & Ad(ws, r, cD) & ":" & Ad(ws, r, cD + 3) & _
            ",""Insuffisant"")>0,""Insuffisant"",IF(COUNT(" & Ad(ws, r, cD) & ":" & Ad(ws, r, cD + 3) & _
            ")=0,"""",MAX(" & Ad(ws, r, cD) & ":" & Ad(ws, r, cD + 3) & ")))"
        If IsArray(mPhiManuel) Then
            If i <= UBound(mPhiManuel) Then ws.Cells(r, cD + 5).Value = mPhiManuel(i)
        End If
        ws.Cells(r, cD + 6).Formula = "=IF(ISNUMBER(" & Ad(ws, r, cD + 4) & "),MAX(" & Ad(ws, r, cD + 4) & _
            ",N(" & Ad(ws, r, cD + 5) & ")),IF(N(" & Ad(ws, r, cD + 5) & ")>0," & Ad(ws, r, cD + 5) & _
            ",IF(" & Ad(ws, r, cD + 4) & "="""","""",""Insuffisant"")))"
        ' Profondeurs Dmax / Dmin avec Phi retenu : dans le tableau de ce diamètre (B0, B1,
        ' ...), sinon (Phi manuel hors du tableau des diamètres) tableau de B0 x Phi / B0
        kb = "(MATCH(" & Ad(ws, r, cD + 6) & "," & phiT & ",0)-1)*" & BLOC_PROFIL
        For k = 0 To 7
            cap = "$" & colCap(k \ 2) & "$" & LIGNE_PROFIL & ":$" & colCap(k \ 2) & "$" & _
                LIGNE_PROFIL + NbPas(ws) - 1
            capk = "OFFSET(" & cap & ",0," & kb & ")"
            zk = "OFFSET(" & z & ",0," & kb & ")"
            ws.Cells(r, cP + k).Formula = "=IF(NOT(ISNUMBER(" & Ad(ws, r, cD + 6) & ")),""Non atteinte""," & _
                "IFERROR(IF(ISNUMBER(MATCH(" & Ad(ws, r, cD + 6) & "," & phiT & ",0)),INDEX(" & zk & _
                ",MATCH(1,INDEX(ISNUMBER(" & capk & ")*(" & capk & ">=" & Ad(ws, r, cC + k) & "),0),0))," & _
                "INDEX(" & z & ",MATCH(1,INDEX(ISNUMBER(" & cap & ")*(" & cap & ">=" & Ad(ws, r, cC + k) & _
                "*" & DONNEES & "$G$" & LIGNE_PIEU + 9 & "/" & Ad(ws, r, cD + 6) & "),0),0))),""Non atteinte""))"
        Next k
        ws.Cells(r, cR).Formula = "=IF(COUNTIF(" & Ad(ws, r, cP) & ":" & Ad(ws, r, cP + 7) & _
            ",""Non atteinte"")>0,""Non atteinte"",MAX(" & Ad(ws, r, cP) & ":" & Ad(ws, r, cP + 7) & "))"
        ' Inclinaison (Oui / Non) et angle : saisies conservées, sinon « Données du projet »
        With ws.Cells(r, cI)
            .Validation.Delete
            .Validation.Add xlValidateList, xlValidAlertStop, , "Non,Oui"
            .Value = IIf(incl, "Oui", "Non")
            If IsArray(mIncl) Then
                If i <= UBound(mIncl) Then
                    If CStr(mIncl(i)) = "Oui" Or CStr(mIncl(i)) = "Non" Then .Value = mIncl(i)
                End If
            End If
        End With
        ws.Cells(r, cA).Value = IIf(IsNumeric(angle0) And Not IsEmpty(angle0), angle0, 0)
        If IsArray(mAngle) Then
            If i <= UBound(mAngle) Then
                If IsNumeric(mAngle(i)) And Not IsEmpty(mAngle(i)) Then ws.Cells(r, cA).Value = mAngle(i)
            End If
        End If
        ' Profondeur réelle : D / cos(angle) si incliné, D sinon
        ws.Cells(r, cRe).Formula = "=IF(ISNUMBER(" & Ad(ws, r, cR) & "),IF(" & Ad(ws, r, cI) & "=""Oui""," & _
            Ad(ws, r, cR) & "/COS(RADIANS(N(" & Ad(ws, r, cA) & ")))," & Ad(ws, r, cR) & "),"""")"
        ' Altitudes : Ztête = cote de l'arase supérieure (« Données du projet », D32),
        ' Zpointe = Ztête - D (profondeur retenue)
        ws.Cells(r, cZ).Formula = "=IF(ISNUMBER(" & DONNEES & "$D$32)," & DONNEES & "$D$32,"""")"
        ws.Cells(r, cZ + 1).Formula = "=IF(AND(ISNUMBER(" & Ad(ws, r, cZ) & "),ISNUMBER(" & Ad(ws, r, cR) & "))," & _
            Ad(ws, r, cZ) & "-" & Ad(ws, r, cR) & ","""")"
        ' Vérification de la longueur : profondeur de reconnaissance - D >= 5 Phi retenu
        ws.Cells(r, cL).Formula = "=IF(ISNUMBER(" & reco & ")," & reco & ","""")"
        ' D = profondeur réelle (D / cos alpha) si le micropieu est incliné
        ws.Cells(r, cL + 1).Formula = "=IF(AND(ISNUMBER(" & Ad(ws, r, cL) & "),ISNUMBER(" & _
            Ad(ws, r, cRe) & "))," & Ad(ws, r, cL) & "-" & Ad(ws, r, cRe) & _
            ","""")"
        ws.Cells(r, cL + 2).Formula = "=IF(ISNUMBER(" & Ad(ws, r, cD + 6) & "),5*" & Ad(ws, r, cD + 6) & ","""")"
        ws.Cells(r, cF).Formula = "=IF(OR(" & Ad(ws, r, cL + 1) & "=""""," & Ad(ws, r, cL + 2) & _
            "=""""),""-"",IF(" & Ad(ws, r, cL + 1) & ">=" & Ad(ws, r, cL + 2) & _
            ",""Vérifiée"",""Non vérifiée""))"
    Next i
    mPhiManuel = Empty
    mGeo = Empty: mGsn = Empty: mAmax = Empty: mAmin = Empty: mIncl = Empty: mAngle = Empty
    rt = r0 + 2 + Application.Max(m, 1)

    ' Diamètre du micropieu B (D19) = diamètre final retenu : le plus grand Phi retenu
    ' des micropieux (B0 de « Données du projet » s'il n'y en a pas) ; il sert à évaluer
    ' la section As,i = pi x B x hi du tableau du frottement axial (et Ab, Pe)
    If m > 0 Then
        ws.Range("B19").Value = "Diamètre final retenu du micropieu B (plus grand " & ChrW(934) & _
            " retenu, sinon B0)"
        ws.Range("D19").Formula = "=IF(COUNT(" & ws.Range(ws.Cells(r0 + 3, cD + 6), ws.Cells(rt, cD + 6)).Address & _
            ")>0,MAX(" & ws.Range(ws.Cells(r0 + 3, cD + 6), ws.Cells(rt, cD + 6)).Address & ")," & _
            DONNEES & "D" & LIGNE_PIEU + 15 & ")"
    End If

    ' Mise en forme
    With ws.Range(ws.Cells(r0, 1), ws.Cells(rt, cF))
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + 2, cF))
        .Font.Bold = True
        .Font.Size = 11
        .Interior.Color = VERT_ENTETE
        .WrapText = True
    End With
    If m > 0 Then
        ws.Range("A" & r0 + 3 & ":A" & rt).Font.Bold = True
        ws.Range(ws.Cells(r0 + 3, 2), ws.Cells(rt, cF)).Interior.Color = JAUNE_CALCUL
        ws.Range(ws.Cells(r0 + 3, cD + 5), ws.Cells(rt, cD + 5)).Interior.Color = VERT_SAISIE ' Phi manuel
        ws.Range("F" & r0 + 3 & ":F" & rt & ",H" & r0 + 3 & ":J" & rt).Interior.Color = VERT_SAISIE ' actions
        ws.Range(ws.Cells(r0 + 3, 2), ws.Cells(rt, cD - 1)).NumberFormat = "0.0"
        ws.Range(ws.Cells(r0 + 3, cD), ws.Cells(rt, cP - 1)).NumberFormat = "0.000"
        ws.Range(ws.Cells(r0 + 3, cP), ws.Cells(rt, cL + 2)).NumberFormat = "0.00"
        ws.Range(ws.Cells(r0 + 3, cD + 6), ws.Cells(rt, cD + 6)).Font.Bold = True
        ws.Range(ws.Cells(r0 + 3, cR), ws.Cells(rt, cL - 1)).Font.Bold = True
        ws.Range(ws.Cells(r0 + 3, cI), ws.Cells(rt, cI)).Interior.Color = VERT_SAISIE  ' inclinaison
        ws.Range(ws.Cells(r0 + 3, cA), ws.Cells(rt, cA)).Interior.Color = VERT_SAISIE  ' angle
        ws.Range(ws.Cells(r0 + 3, cA), ws.Cells(rt, cA)).NumberFormat = "0.0"
        ws.Range(ws.Cells(r0 + 3, cI), ws.Cells(rt, cI)).Font.Bold = False
        ws.Range(ws.Cells(r0 + 3, cF), ws.Cells(rt, cF)).Font.Bold = True
        ' Vérification en rouge si non vérifiée
        With ws.Range(ws.Cells(r0 + 3, cF), ws.Cells(rt, cF)).FormatConditions
            .Delete
            .Add(xlCellValue, xlEqual, "=""Non vérifiée""").Font.Color = RGB(192, 0, 0)
        End With
    End If
    ws.Range(ws.Columns(cD), ws.Columns(cF)).ColumnWidth = 12
    ws.Rows(r0).RowHeight = 30
    AfficherInclinaison ws
End Sub

' Colonnes de l'angle et de la profondeur réelle de la vérification de la profondeur
' affichées seulement si au moins un micropieu est incliné (colonne « Incliné » = Oui).
' Appelée à la création et quand une inclinaison change (ChangementCalcul).
Public Sub AfficherInclinaison(Optional ws As Worksheet)
    Dim r0 As Long, c As Long, cI As Long, r As Long, n As Long
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
    End If
    cI = ColonneInclinaison(ws, r0)
    If cI = 0 Then Exit Sub
    r = r0 + 3
    Do While Left(CStr(ws.Range("A" & r).Value), 2) = "MP"
        If CStr(ws.Cells(r, cI).Value) = "Oui" Then n = n + 1
        r = r + 1
    Loop
    ' angle (cI + 2) et profondeur réelle (cI + 3), de part et d'autre de D (cI + 1)
    ws.Range(ws.Columns(cI + 2), ws.Columns(cI + 3)).Hidden = (n = 0)
End Sub

' « Incliné » et angle de « Données du projet » recopiés sur tous les micropieux de la
' vérification de la profondeur (appelée par ChangementDonnees, module DonneesProjet)
Public Sub AppliquerInclinaisonDonnees()
    Dim ws As Worksheet, wp As Worksheet, cI As Long, r0 As Long, r As Long
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Set wp = ActiveWorkbook.Worksheets("Données du projet")
    cI = ColonneInclinaison(ws, r0)
    If cI = 0 Then Exit Sub
    r = r0 + 3
    Do While Left(CStr(ws.Range("A" & r).Value), 2) = "MP"
        ws.Cells(r, cI).Value = IIf(CStr(wp.Range("D" & LIGNE_PIEU + 2).Value) = "Oui", "Oui", "Non")
        ws.Cells(r, cI + 2).Value = N0(wp.Range("D" & LIGNE_PIEU + 3).Value)
        r = r + 1
    Loop
    AfficherInclinaison ws
End Sub

' Valeur numérique (0 si vide ou texte)
Private Function N0(v As Variant) As Double
    If Not IsEmpty(v) And IsNumeric(v) Then N0 = CDbl(v)
End Function

' Colonne « Incliné » de la vérification de la profondeur (0 si absente) ; r0 : ligne
' de l'en-tête « N° MICROPIEU »
Private Function ColonneInclinaison(ws As Worksheet, r0 As Long) As Long
    Dim c As Long, v As Variant
    v = Application.Match("N° MICROPIEU", ws.Range("A1:A5000"), 0)
    If IsError(v) Then Exit Function
    r0 = CLng(v)
    For c = 2 To 80
        If CStr(ws.Cells(r0 + 1, c).Value) = "Incliné" Then ColonneInclinaison = c: Exit Function
    Next c
End Function

' Nombre de lignes (pas) du tableau de la capacité portante, écrit en ZA à la création ;
' ancienne feuille sans ce nombre : MAX_COUCHES x NB_SUBDIV
Private Function NbPas(ws As Worksheet) As Long
    Dim v As Variant
    v = ws.Range("ZA" & LIGNE_PROFIL - 3).Value
    If Not IsEmpty(v) And IsNumeric(v) Then NbPas = CLng(v)
    If NbPas < 1 Then NbPas = MAX_COUCHES * NB_SUBDIV
End Function

' Ligne du titre de la vérification de la profondeur des micropieux : sous le tableau
' de la capacité portante (et sous ses graphiques)
Private Function LigneVerif(ws As Worksheet) As Long
    LigneVerif = LIGNE_PROFIL + Application.Max(NbPas(ws), NB_MIN_LIGNES) + 2
End Function

' Lignes du tableau du frottement axial limitées au nombre de couches de « Données du
' projet » (lignes des couches suivantes masquées)
Public Sub AfficherCouchesCalcul(Optional ws As Worksheet)
    Dim nb As Variant
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
    End If
    nb = ActiveWorkbook.Worksheets("Données du projet").Range(CELLULE_NB_COUCHES).Value
    If IsEmpty(nb) Or Not IsNumeric(nb) Then nb = MAX_COUCHES
    If nb < 1 Then nb = 1
    If nb > MAX_COUCHES Then nb = MAX_COUCHES
    ws.Rows(LIGNE_FROT).Resize(MAX_COUCHES).Hidden = False
    If nb < MAX_COUCHES Then ws.Rows(LIGNE_FROT + nb).Resize(MAX_COUCHES - nb).Hidden = True
End Sub

' Pas de calcul, épaisseurs, terrains neutralisés ou nombre de couches modifiés : la
' feuille est recréée si le nombre total de pas change (lignes du tableau de la
' capacité portante), sinon les lignes du frottement axial sont seulement ajustées.
' Appelée par ChangementCalcul (code de la feuille, CodeFeuilleCalcul.txt) et par
' ChangementDonnees (module DonneesProjet).
Public Sub ActualiserPas()
    Dim ws As Worksheet, n As Long
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ws.Calculate
    n = Application.Max(ws.Range("ZA" & LIGNE_FROT & ":ZA" & LIGNE_FROT + MAX_COUCHES - 1))
    If n < 1 Then n = 1
    If n <> NbPas(ws) Then
        Construire False
    Else
        AfficherCouchesCalcul ws
    End If
End Sub

' Appelée par le code de la feuille « Calcul et vérification » (CodeFeuilleCalcul.txt) :
' pas de calcul saisi (colonne A du frottement axial) -> tableau de la capacité
' portante adapté au nombre total de subdivisions
Public Sub ChangementCalcul(ws As Worksheet, Target As Range)
    Dim c As Range, zone As Range, cI As Long, r0 As Long
    ' Inclinaison d'un micropieu (vérification de la profondeur) : colonnes angle et
    ' profondeur réelle affichées ou masquées
    cI = ColonneInclinaison(ws, r0)
    If cI > 0 Then
        If Not Intersect(Target, ws.Range(ws.Cells(r0 + 3, cI), ws.Cells(r0 + 300, cI))) Is Nothing Then
            Application.EnableEvents = False
            AfficherInclinaison ws
            Application.EnableEvents = True
        End If
    End If
    Set zone = Intersect(Target, ws.Range("A" & LIGNE_FROT & ":A" & LIGNE_FROT + MAX_COUCHES - 1))
    If zone Is Nothing Then Exit Sub
    Application.EnableEvents = False
    On Error Resume Next
    ' Pas écrit en texte (« 0.5 » ou « 0,5 ») : converti en nombre
    For Each c In zone.Cells
        If Not IsEmpty(c.Value) And VarType(c.Value) = vbString Then
            If ValeurPas(c.Value) > 0 Then c.Value = ValeurPas(c.Value)
        End If
    Next c
    On Error GoTo 0
    ActualiserPas
    Application.EnableEvents = True
End Sub

' Valeur d'un pas saisi (nombre, ou texte avec point ou virgule décimale) ; 0 sinon
Private Function ValeurPas(v As Variant) As Double
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbInteger Or VarType(v) = vbLong Or _
        VarType(v) = vbSingle Or VarType(v) = vbCurrency Then
        ValeurPas = CDbl(v)
    Else
        ValeurPas = Val(Replace(Trim$(CStr(v)), ",", "."))
    End If
    If ValeurPas < 0 Then ValeurPas = 0
End Function

' Code de la feuille « Calcul et vérification » (Worksheet_Change -> ChangementCalcul)
' installé automatiquement si Excel autorise l'accès au projet VBA (Fichier > Options >
' Centre de gestion de la confidentialité > Paramètres des macros > « Accès approuvé au
' modèle d'objet du projet VBA ») ; sinon, coller CodeFeuilleCalcul.txt à la main.
Private Sub InstallerCodeFeuille(ws As Worksheet)
    Dim cm As Object, texte As String
    On Error GoTo Fin
    texte = ws.Parent.VBProject.Name         ' accès au projet (CodeName à jour)
    texte = ""
    Set cm = ws.Parent.VBProject.VBComponents(ws.CodeName).CodeModule
    If cm.CountOfLines > 0 Then texte = cm.Lines(1, cm.CountOfLines)
    If InStr(1, texte, "ChangementCalcul", vbTextCompare) = 0 Then
        cm.AddFromString "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
            "    ChangementCalcul Me, Target" & vbCrLf & "End Sub"
    End If
    ws.Range("A" & LIGNE_FROT - 1).ClearComments
    Exit Sub
Fin:
    ' Accès refusé : rappel sur l'en-tête du pas de calcul
    On Error Resume Next
    Commentaire ws.Range("A" & LIGNE_FROT - 1), "Pour que le tableau de la capacité portante et la " & _
        "vérification de la profondeur s'ajustent dès la saisie d'un pas : coller CodeFeuilleCalcul.txt " & _
        "dans le code de cette feuille, ou autoriser l'accès au projet VBA puis relancer " & _
        "CreerCalculVerification."
End Sub

' Adresse relative (ex. « F395 ») de la cellule ligne r, colonne c
Private Function Ad(ws As Worksheet, r As Long, c As Long) As String
    Ad = ws.Cells(r, c).Address(False, False)
End Function

' Graphique de la capacité portante en compression en fonction de la profondeur pour
' le tableau par pas de calcul du diamètre Bk (k = 0 à NB_DIAMETRES - 1, tableau de Bk
' décalé de k x BLOC_PROFIL colonnes), à droite de ce tableau : Rc;cr,d ELS
' caractéristique, Rc;cr,d ELS quasi permanent et Rc,d ELU fondamental en abscisse,
' profondeur z en ordonnée, croissante vers le bas. Les séries renvoient à des noms de
' la feuille (CPk_z, CPk_1 à CPk_3) limités aux pas renseignés (OFFSET sur le nombre de
' z), pour que le graphique suive le nombre de couches et leurs épaisseurs.
Private Sub GraphiqueCapacite(ws As Worksheet, k As Long)
    Dim co As ChartObject, f As String, nb As String, j As Long, cols As Variant, noms As Variant
    Dim d As Long, col As String, lg As Double
    f = "'" & ws.Name & "'!"
    d = k * BLOC_PROFIL
    nb = "MAX(1,COUNT(" & f & ws.Cells(LIGNE_PROFIL, 3 + d).Address & ":" & _
        ws.Cells(LIGNE_PROFIL + NbPas(ws) - 1, 3 + d).Address & "))"
    cols = Array(3, 18, 20, 22)              ' z (C), ELS cara. (R), ELS QP (T), ELU fond. (V)
    noms = Array("Profondeur z", "ELS caractéristique - Rc;cr,d", "ELS quasi permanent - Rc;cr,d", _
        "ELU fondamental - Rc,d")
    For j = 0 To 3
        ws.Names.Add Name:="CP" & k & "_" & IIf(j = 0, "z", CStr(j)), _
            RefersTo:="=OFFSET(" & f & ws.Cells(LIGNE_PROFIL, cols(j) + d).Address & ",0,0," & nb & ")"
    Next j
    ' Position : à droite du tableau (colonne AA + décalage), largeur jusqu'au tableau du
    ' diamètre suivant (620 points au plus) ; taille fixe si les colonnes changent ensuite
    lg = 620
    If k < NB_DIAMETRES - 1 Then
        lg = ws.Cells(1, 2 + (k + 1) * BLOC_PROFIL).Left - ws.Cells(1, 27 + d).Left - 15
        If lg > 620 Then lg = 620
        If lg < 300 Then lg = 300
    End If
    With ws.Cells(LIGNE_PROFIL - 3, 27 + d)
        Set co = ws.ChartObjects.Add(.Left, .Top, lg, 480)
    End With
    co.Placement = xlFreeFloating
    With co.Chart
        .ChartType = xlXYScatterSmoothNoMarkers
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        For j = 1 To 3
            With .SeriesCollection.NewSeries
                .Name = noms(j)
                .XValues = "=" & f & "CP" & k & "_" & j
                .Values = "=" & f & "CP" & k & "_z"
                ' Courbe lissée, sans marqueurs, épaisseur 3 pt
                .ChartType = xlXYScatterSmoothNoMarkers
                .Smooth = True
                .MarkerStyle = xlMarkerStyleNone
                .Format.Line.Weight = 3
            End With
        Next j
        .HasTitle = True
        .ChartTitle.Text = "Capacité portante en compression en fonction de la profondeur - " & _
            "diamètre B" & k & " = " & Format(ws.Parent.Worksheets("Données du projet").Range("G" & _
            LIGNE_PIEU + 9 + k).Value, "0.000") & " m"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        With .Axes(xlValue)                  ' profondeur, croissante vers le bas
            .ReversePlotOrder = True
            .MinimumScale = 0
            .HasTitle = True
            .AxisTitle.Text = "Profondeur z (m)"
        End With
        With .Axes(xlCategory)
            .MinimumScale = 0
            .HasTitle = True
            .AxisTitle.Text = "Capacité portante (kN)"
        End With
    End With
End Sub

' Valeur de la colonne col pour la dernière couche renseignée (nom en colonne C) du
' tableau du frottement axial (lignes 39 à 53)
Private Function Derniere(col As String) As String
    Derniere = "=IFERROR(LOOKUP(2,1/($C$39:$C$53<>""""),$" & col & "$39:$" & col & "$53),"""")"
End Function

' Une ligne « intitulé (B) | symbole (C) | valeur ou formule (D) | unité (E) »
Private Sub LigneS(ws As Worksheet, r As Long, intitule As String, symbole As String, _
        formule As String, unite As String)
    ws.Range("B" & r).Value = intitule
    ws.Range("C" & r).Value = symbole
    ws.Range("D" & r).Formula = formule
    ws.Range("E" & r).Value = unite
End Sub

' Tableau 2 x 2 (deux lignes « intitulé | symbole ») pour une plage de 2 x 2 cellules
Private Function Array2(a As String, b As String, c As String, d As String) As Variant
    Dim t(1 To 2, 1 To 2) As String
    t(1, 1) = a: t(1, 2) = b: t(2, 1) = c: t(2, 2) = d
    Array2 = t
End Function

' Une ligne « intitulé (B:C fusionnées) | valeur ou formule (D) | unité (E) »
Private Sub Ligne(ws As Worksheet, r As Long, intitule As String, formule As String, unite As String)
    ws.Range("B" & r).Value = intitule
    ws.Range("B" & r & ":C" & r).Merge
    ws.Range("D" & r).Formula = formule
    ws.Range("E" & r).Value = unite
End Sub

' Sous-titre en gras, souligné, en colonne B
Private Sub SousTitre(ws As Worksheet, r As Long, texte As String, Optional col As String = "B")
    With ws.Range(col & r)
        .Value = texte
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
End Sub

' Rang (1 à 6) de la catégorie de sol dans la liste CATEGORIES_SOL (feuille
' « Données du projet ») = colonne du tableau qs max
Private Function Rang6(sol As String) As String
    Rang6 = "MATCH(" & sol & ",{""" & Replace(CATEGORIES_SOL, ",", """,""") & """},0)"
End Function

' Colonne (1 à 5) des tableaux pressiométriques F.5.2.1 et F.5.2.2 : argiles et sols
' intermédiaires -> 1, sables et graves -> 2, craies -> 3, marnes -> 4, roches -> 5
Private Function Rang5(sol As String) As String
    Rang5 = "CHOOSE(" & Rang6(sol) & ",1,1,2,3,4,5)"
End Function

' Valeur (ligne lig, colonne du sol sol) d'un tableau du frottement pour la méthode de
' calcul retenue : tableau G (6 colonnes de sols) si pénétrométrique, tableau F sinon
' (5 colonnes, ou 6 si f6) ; zones = « plage G|plage F »
Private Function IndexNorme(zones As String, lig As String, sol As String, _
        Optional f6 As Boolean = False) As String
    Dim z() As String
    z = Split(zones, "|")
    If mPenetro Then
        IndexNorme = "INDEX(" & z(0) & "," & lig & "," & Rang6(sol) & ")"
    Else
        IndexNorme = "INDEX(" & z(1) & "," & lig & "," & IIf(f6, Rang6(sol), Rang5(sol)) & ")"
    End If
End Function

' Capture de la norme fichier (module CommentaireImage, dossier « images-normes ») en
' commentaire de la cellule c ; à défaut, commentaire texte
Private Sub NoteNorme(c As Range, fichier As String, texte As String)
    On Error Resume Next
    Application.Run "CommentaireNorme", c, fichier, texte
    If Err.Number <> 0 Then
        Err.Clear
        If texte <> "" Then Commentaire c, texte
    End If
    On Error GoTo 0
End Sub

' Commentaire (signification) sur une cellule d'en-tête
Private Sub Commentaire(c As Range, texte As String)
    ' Commentaire en image (module CommentaireImage) ; à défaut, commentaire texte
    On Error Resume Next
    Application.Run "CommentaireImage", c, texte
    If Err.Number <> 0 Then
        Err.Clear
        c.ClearComments
        c.AddComment texte
        c.Comment.Shape.TextFrame.AutoSize = True
    End If
    On Error GoTo 0
End Sub

' Tableau 2D (lignes x colonnes) à partir d'un tableau de lignes (Array d'Array)
Private Function Array2D(lignes As Variant) As Variant
    Dim t() As Variant, i As Long, j As Long
    ReDim t(1 To UBound(lignes) + 1, 1 To UBound(lignes(0)) + 1)
    For i = 0 To UBound(lignes)
        For j = 0 To UBound(lignes(i))
            t(i + 1, j + 1) = lignes(i)(j)
        Next j
    Next i
    Array2D = t
End Function

' Valeur d'un tableau de la norme (zone) en ligne lig et colonne col ; vide si erreur
Private Function FormuleIndex(zone As String, lig As String, col As String) As String
    FormuleIndex = "=IFERROR(INDEX(" & zone & "," & lig & "," & col & "),"""")"
End Function

' alpha x fsol ; « - » si alpha n'est pas donné par la norme (catégories 17 et 18 :
' valeurs de la catégorie 1, foré simple, dans « Tableaux NF P94-262 »)
Private Function FormuleProduit(alpha As String, fsol As String) As String
    FormuleProduit = "=IF(OR(" & fsol & "="""",ISBLANK(" & alpha & ")),""""," & _
        "IF(ISNUMBER(" & alpha & ")," & alpha & "*" & fsol & ",""-""))"
End Function

Private Function FormuleQs(produit As String, qsmax As String) As String
    FormuleQs = "=IF(AND(ISNUMBER(" & produit & "),ISNUMBER(" & qsmax & ")),MIN(" & _
        produit & "," & qsmax & "),IF(" & produit & "="""","""",""-""))"
End Function

' Ligne du tableau par couche (couche n° i, ligne ld dans « Données du projet ») :
' calculée seulement si le nom et la catégorie de la couche sont renseignés.
' Pe en D23 ; gamma = gamma R;d1 x gamma R;d2 (formule lue dans les tableaux des
' coefficients de modèle) ; terrain neutralisé : frottement négligé, Rs,i = 0.
Private Sub CoucheLigne(ws As Worksheet, r As Long, i As Long, ld As Long, cat As String, _
        gamma As String, zAlpha As String, zABC As String, zQs As String)
    Dim cond As String, sol As String, pl As String, c As String
    cond = "=IF(OR(" & DONNEES & "$C$" & ld & "=""""," & DONNEES & "$D$" & ld & "=""""),"""","
    sol = "$D$" & r
    pl = "$G$" & r
    ws.Range("B" & r).Formula = cond & i & ")"
    ws.Range("C" & r).Formula = cond & DONNEES & "$C$" & ld & ")"
    ws.Range("D" & r).Formula = cond & DONNEES & "$D$" & ld & ")"
    ws.Range("E" & r).Formula = cond & "IF(ISNUMBER(" & DONNEES & "$E$" & ld & ")," & _
        DONNEES & "$E$" & ld & ",""""))"
    ws.Range("F" & r).Formula = cond & "IF(ISNUMBER(" & DONNEES & "$F$" & ld & ")," & _
        DONNEES & "$F$" & ld & ",""""))"
    ' pl* (Données, colonne G) ou qc (Données, colonne I) selon la méthode de calcul
    c = IIf(mPenetro, "$I$", "$G$")
    ws.Range("G" & r).Formula = cond & "IF(ISNUMBER(" & DONNEES & c & ld & ")," & _
        DONNEES & c & ld & ",""""))"
    mCellule = "H" & r
    ws.Range("H" & r).Formula = cond & "IFERROR(" & IndexNorme(zABC, "1", sol) & ",""""))"
    mCellule = "I" & r
    ws.Range("I" & r).Formula = cond & "IFERROR(" & IndexNorme(zAlpha, cat, sol) & ",""""))"
    ' fsol,i = (a.X + b)(1 - exp(-c.X)), X = pl* ou qc ; a, b, c du tableau F.5.2.2 ou G.5.2.2
    mCellule = "J" & r
    ws.Range("J" & r).Formula = "=IF(OR(B" & r & "="""",G" & r & "="""",H" & r & "=""""),"""",IFERROR(" & _
        "1000*(" & IndexNorme(zABC, "2", sol) & "*" & pl & "+" & IndexNorme(zABC, "3", sol) & _
        ")*(1-EXP(-" & IndexNorme(zABC, "4", sol) & "*" & pl & ")),""""))"
    ws.Range("K" & r).Formula = FormuleProduit("$I$" & r, "$J$" & r)
    mCellule = "L" & r
    ws.Range("L" & r).Formula = cond & "IFERROR(" & IndexNorme(zQs, cat, sol, True) & ",""""))"
    ws.Range("M" & r).Formula = FormuleQs("$K$" & r, "$L$" & r)
    ' As,i = pi x B x hi, B = diamètre final retenu (D19)
    ws.Range("N" & r).Formula = "=IF(OR(B" & r & "="""",F" & r & "="""",N($D$19)=0),""""," & _
        "PI()*$D$19*F" & r & ")"
    ' qs,i = qs / (gamma R;d1 x gamma R;d2) ; 0 si le terrain est neutralisé
    ws.Range("O" & r).Formula = "=IF(B" & r & "="""","""",IF(Q" & r & "=""Oui"",0," & _
        "IFERROR(M" & r & "/(" & gamma & "),""-"")))"
    ws.Range("P" & r).Formula = "=IF(AND(ISNUMBER(O" & r & "),ISNUMBER(N" & r & "))," & _
        "IF(Q" & r & "=""Oui"",0,O" & r & "*N" & r & "),IF(B" & r & "="""","""",""-""))"
    ws.Range("Q" & r).Formula = cond & DONNEES & "$M$" & ld & ")"
    ws.Range("E" & r & ":G" & r & ",I" & r).NumberFormat = "0.00"
    ws.Range("J" & r & ":M" & r & ",O" & r & ":P" & r).NumberFormat = "0.0"
    ws.Range("N" & r).NumberFormat = "0.000"
End Sub

' Valeurs d'un tableau de la feuille « Tableaux NF P94-262 », nommées d'après le numéro
' du tableau (ex. « Tableau F.5.2.1 » -> NF_F_5_2_1) :
'   - le titre du tableau est cherché exactement (cellule qui commence par « Tableau
'     F.5.2.1 - », en respectant la casse, pour ne pas prendre un renvoi écrit dans une
'     note) ; le titre étant fusionné sur la largeur du tableau, on repart de la
'     première cellule de la fusion ;
'   - dans les lignes sous le titre, la première colonne de valeurs est repérée par son
'     en-tête : 1re colonne de sol (en-tête commençant par « Argile ») ou, pour les
'     coefficients de modèle F.2.1 et G.2.1, 1re colonne gamma R;d1 ;
'   - les valeurs commencent sur la ligne sous cet en-tête ; nbLignes x nbCols valeurs.
' Renvoie le nom, utilisé dans les formules à la place de l'adresse.
Private Function ZoneTableau(wt As Worksheet, titre As String, decalage As Long, _
        nbLignes As Long, nbCols As Long) As String
    Dim c As Range, trouve As Range, ent As Range, premier As String, nom As String
    Dim cle As String, lig As Long, col As Long
    Set c = wt.Cells.Find(What:=titre & " - ", After:=wt.Cells(wt.Rows.Count, wt.Columns.Count), _
        LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=True)
    If Not c Is Nothing Then
        premier = c.Address
        Do
            If Left$(CStr(c.MergeArea.Cells(1, 1).Value), Len(titre) + 3) = titre & " - " Then
                Set trouve = c.MergeArea.Cells(1, 1)
                Exit Do
            End If
            Set c = wt.Cells.FindNext(c)
            If c Is Nothing Then Exit Do
        Loop Until c.Address = premier
    End If
    If trouve Is Nothing Then Err.Raise vbObjectError + 1, , "« " & titre & " » introuvable " & _
        "dans la feuille « " & wt.Name & " » : relancez CreerTableauxNFP94262."
    ' En-tête de la 1re colonne de valeurs, dans les 4 lignes sous le titre
    If titre = "Tableau F.2.1" Or titre = "Tableau G.2.1" Then
        cle = "R;d1"
    Else
        cle = "Argile"
    End If
    For lig = trouve.Row + 1 To trouve.Row + 4
        For col = trouve.Column To trouve.Column + 15
            If InStr(1, CStr(wt.Cells(lig, col).Value), cle, vbTextCompare) > 0 Then
                Set ent = wt.Cells(lig, col)
                Exit For
            End If
        Next col
        If Not ent Is Nothing Then Exit For
    Next lig
    If ent Is Nothing Then Err.Raise vbObjectError + 2, , "En-tête des valeurs du « " & titre & _
        " » introuvable dans la feuille « " & wt.Name & " » : relancez CreerTableauxNFP94262."
    nom = "NF_" & Replace(Replace(Trim$(Mid$(titre, Len("Tableau ") + 1)), ".", "_"), " ", "")
    ActiveWorkbook.Names.Add Name:=nom, RefersTo:="='" & wt.Name & "'!" & _
        ent.Offset(1 + decalage, 0).Resize(nbLignes, nbCols).Address
    ZoneTableau = nom
End Function
