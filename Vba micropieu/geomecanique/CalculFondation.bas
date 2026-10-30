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
'   - pas de calcul en colonne A, à gauche du frottement axial (épaisseur / k, k = 20
'     ou 1 si le terrain est neutralisé), et tableau de la capacité portante par pas de
'     calcul sous les valeurs de calcul (couche, profondeur, cote, pas, qs,i, Rs,k, Rt,k,
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
' Pas de calcul : nombre de subdivisions des couches non neutralisées (entier >= 20) ;
' une couche neutralisée n'est pas subdivisée (pas = épaisseur)
Private Const NB_SUBDIV As Long = 20
Private Const LIGNE_PROFIL As Long = 91        ' 1re ligne du tableau par pas de calcul
Private Const TABLEAUX As String = "Tableaux NF P94-262"
Private mCellule As String     ' cellule en cours d'écriture (indiquée en cas d'erreur)
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
    Dim c As String

    Application.ScreenUpdating = False
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

    Ligne ws, 19, "Largeur de la fondation B", "=" & DONNEES & "D" & LIGNE_PIEU + 15, "m"
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
        "=IF(COUNTIF(OFFSET(" & DONNEES & "$J$" & lc & ",0,0," & nb & "),""Oui"")=0,""Non""," & _
        "SUMIF(OFFSET(" & DONNEES & "$J$" & lc & ",0,0," & nb & "),""Oui""," & _
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
        "As,i (m²)", "qs,i (kPa)", "Rs,i (kN)", "Terrain neutralisé")
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
    ' Pas = épaisseur / k, k = 1 si le terrain est neutralisé, NB_SUBDIV sinon. Colonnes
    ' d'aide masquées (AX à AZ) : cumul de k, profondeur du bas de la couche, cumul de Rs,i.
    ' Chaque ligne du tableau correspond à un pas : couche (BA) et rang du pas dans la
    ' couche (BB), masqués ; z = bas de la couche précédente + rang x pas ;
    ' Rs,k(z) = Rs,i des couches au-dessus + qs,i x Pe x (z - haut de la couche) ;
    ' Rt,k = Rs,k ; micropieux : ple*, kp max et Def sans objet (« - »), Rb,k = 0 ;
    ' valeurs de calcul comme dans le tableau des valeurs de calcul
    etape = "pas de calcul"
    ws.Range("A38").Value = "Pas de calcul (m)"
    ws.Range("R38").Value = "Nombre de subdivisions k"
    ws.Range("AX38:AZ38").Value = Array(ChrW(931) & " k", "z bas (m)", ChrW(931) & " Rs,i (kN)")
    For r = 39 To 53
        ws.Range("R" & r).Formula = "=IF(B" & r & "="""","""",IF(Q" & r & "=""Oui"",1," & NB_SUBDIV & "))"
        ws.Range("A" & r).Formula = "=IF(OR(B" & r & "="""",N(F" & r & ")=0),"""",F" & r & "/R" & r & ")"
        ws.Range("AX" & r).Formula = "=IF(B" & r & "="""","""",SUM($R$39:R" & r & "))"
        ws.Range("AY" & r).Formula = "=IF(B" & r & "="""","""",SUM($F$39:F" & r & "))"
        ws.Range("AZ" & r).Formula = "=IF(B" & r & "="""","""",SUM($P$39:P" & r & "))"
    Next r
    ws.Range("A39:A53").NumberFormat = "0.000"
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
    ws.Range("A39:A53,R39:R53").Interior.Color = JAUNE_CALCUL

    etape = "capacité portante par pas de calcul"
    ' Valeurs cumulées en profondeur, pas par pas : ligne précédente + apport du pas
    '   Rs,k(z) = Rs,k(z - pas) + qs,i x Pe x pas (0 si le terrain est neutralisé), soit
    '   Rs,i = qs,i x As,i du tableau du frottement avec As,i = Pe x pas ; Rt,k = Rs,k ;
    '   valeurs de calcul déduites de Rs,k et Rt,k cumulés (mêmes relations que le
    '   tableau des valeurs de calcul)
    lt = LIGNE_PROFIL + MAX_COUCHES * NB_SUBDIV - 1        ' dernière ligne prévue
    SousTitre ws, LIGNE_PROFIL - 3, "Capacité portante des micropieux en fonction de la " & _
        "profondeur (par pas de calcul, valeurs cumulées)"
    r = LIGNE_PROFIL - 2
    ws.Range("B" & r).Value = "Références suivant couche"
    ws.Range("N" & r).Value = "ELS caractéristique (14.2.1.1)"
    ws.Range("P" & r).Value = "ELS quasi permanent (14.2.1.2)"
    ws.Range("R" & r).Value = "ELU fondamental (C.2.3.1)"
    ws.Range("T" & r).Value = "ELU accidentel (C.2.3.2)"
    ws.Range("B" & r & ":M" & r).Merge
    For i = 0 To 3
        ws.Range(Chr(Asc("N") + 2 * i) & r & ":" & Chr(Asc("O") + 2 * i) & r).Merge
    Next i
    ws.Range("B" & r + 1 & ":U" & r + 1).Value = Array("N° de couche", "Profondeur z (m)", "Cote", _
        "Pas de calcul (m)", "qs,i (kPa)", "Rs,k (kN)", "Rt,k (kN)", "ple* (MPa)", "h (m)", "kp max", _
        "Def (m)", "Rb,k (kN)", "Rc;cr,d (kN)", "Rt;cr,d (kN)", "Rc;cr,d (kN)", "Rt;cr,d (kN)", _
        "Rc,d (kN)", "Rt,d (kN)", "Rc,d (kN)", "Rt,d (kN)")
    ws.Range("D" & r + 1).Formula = "=""Cote (""&" & DONNEES & "$E$29&"")"""
    ' Formules de la 1re ligne (j = rang du pas = ligne - LIGNE_PROFIL + 1), recopiées
    ' vers le bas ; la ligne précédente de la 1re ligne est l'en-tête (texte -> N() = 0)
    r = LIGNE_PROFIL
    c = "(ROW()-" & LIGNE_PROFIL - 1 & ")"                   ' rang j du pas
    ws.Range("BA" & r).Formula = "=IF(OR(COUNT($AX$39:$AX$53)=0," & c & ">MAX($AX$39:$AX$53)),""""," & _
        "COUNTIF($AX$39:$AX$53,""<""&" & c & ")+1)"
    ws.Range("BB" & r).Formula = "=IF(BA" & r & "="""",""""," & c & "-IF(BA" & r & "=1,0," & _
        "INDEX($AX$39:$AX$53,BA" & r & "-1)))"
    ws.Range("B" & r).Formula = "=IF(BA" & r & "="""","""",INDEX($B$39:$B$53,BA" & r & "))"
    ws.Range("E" & r).Formula = "=IF(BA" & r & "="""","""",INDEX($A$39:$A$53,BA" & r & "))"
    ws.Range("J" & r).Formula = "=IF(BA" & r & "="""","""",BB" & r & "*E" & r & ")"
    ws.Range("C" & r).Formula = "=IF(BA" & r & "="""","""",IF(BA" & r & "=1,0,INDEX($AY$39:$AY$53,BA" & _
        r & "-1))+J" & r & ")"
    ws.Range("D" & r).Formula = "=IF(C" & r & "="""","""",IF(ISNUMBER(" & DONNEES & "$D$29)," & _
        DONNEES & "$D$29-C" & r & ",-C" & r & "))"
    ws.Range("F" & r).Formula = "=IF(BA" & r & "="""","""",INDEX($O$39:$O$53,BA" & r & "))"
    ' Rs,k cumulé : valeur de la ligne précédente + qs,i x Pe x pas
    ws.Range("G" & r).Formula = "=IF(BA" & r & "="""","""",N(G" & r - 1 & ")+IF(INDEX($Q$39:$Q$53,BA" & _
        r & ")=""Oui"",0,N(F" & r & ")*N($D$23)*E" & r & "))"
    ws.Range("H" & r).Formula = "=IF(G" & r & "="""","""",G" & r & ")"
    ws.Range("I" & r).Formula = "=IF(BA" & r & "="""","""",""-"")"
    ws.Range("K" & r).Formula = "=I" & r
    ws.Range("L" & r).Formula = "=I" & r
    ws.Range("M" & r).Formula = "=IF(BA" & r & "="""","""",0)"
    ' ELS : Rc;cr,d = (0,5 Rb,k + 0,7 Rs,k) / gamma cr ; Rt;cr,d = 0,7 Rt,k / gamma s;cr
    ' (x 0,15 sans essai de contrôle) ; ELU : Rc,d = Rb,k / gamma b + Rs,k / gamma s ;
    ' Rt,d = Rt,k / gamma s;t
    For i = 0 To 1
        c = Chr(Asc("D") + i)                ' colonne du facteur : D ou E
        ws.Range(Chr(Asc("N") + 2 * i) & r).Formula = "=IF(G" & r & "="""","""",(0.5*M" & r & _
            "+0.7*G" & r & ")/$" & c & "$34)"
        ws.Range(Chr(Asc("O") + 2 * i) & r).Formula = "=IF(H" & r & "="""","""",0.7*H" & r & "/$" & c & _
            "$35*IF($D$76=""Non"",0.15,1))"
        ws.Range(Chr(Asc("R") + 2 * i) & r).Formula = "=IF(G" & r & "="""","""",M" & r & "/$" & c & _
            "$28+G" & r & "/$" & c & "$29)"
        ws.Range(Chr(Asc("S") + 2 * i) & r).Formula = "=IF(H" & r & "="""","""",H" & r & "/$" & c & "$31)"
    Next i
    ws.Range("B" & r & ":BB" & lt).FillDown
    ws.Range("C" & r & ":D" & lt & ",J" & r & ":J" & lt).NumberFormat = "0.00"
    ws.Range("E" & r & ":E" & lt).NumberFormat = "0.000"
    ws.Range("F" & r & ":H" & lt & ",M" & r & ":U" & lt).NumberFormat = "0.0"
    With ws.Range("B" & LIGNE_PROFIL - 2 & ":U" & lt)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("B" & LIGNE_PROFIL - 2 & ":U" & LIGNE_PROFIL - 1)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .WrapText = True
    End With
    ws.Columns("AX:BB").Hidden = True

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
    ws.Columns("E:U").ColumnWidth = 12
    ' Bloc de l'effet de groupe (écrit en lignes 78 à 86) déplacé avant le tableau du
    ' frottement axial par couche (stratigraphie) : il occupe les lignes 37 à 45 et la
    ' suite des colonnes A à BB descend de 9 lignes (jusqu'à la ligne 86 ; le tableau par
    ' pas de calcul, plus bas, ne bouge pas), les formules suivent le déplacement
    etape = "déplacement de l'effet de groupe"
    ws.Range("A78:BB86").Cut
    ws.Range("A37").Insert Shift:=xlDown
    Application.CutCopyMode = False

    etape = "graphique de la capacité portante"
    GraphiqueCapacite ws

    etape = "vérification de la profondeur des micropieux"
    VerificationProfondeur ws
    AfficherGroupe ws
    If activer Then ws.Activate

Fin:
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
'   descente de charges : Gmax = G, Gmin = G ou 0 (choix), Qmax = Q, Qmin = 0, eau
'   EB, EH, EE (kN, 0 sans nappe) ;
'   combinaisons verticales max / min :
'     ELS caractéristique : G + Q ;  ELS quasi permanent : G + psi2 Q ;
'     ELU fondamental : 1,35 Gmax + 1,5 Qmax (max), 1,0 Gmin + 1,5 Qmin (min) ;
'     ELU accidentel : G + psi1 Q ; psi1, psi2 selon la catégorie de bâtiment ;
'   diamètre B du micropieu pour chaque cas max / min (celui du tableau de la capacité
'   portante) ;
'   profondeur minimale requise pour chaque cas : 1re profondeur du tableau par pas de
'   calcul où la portance de la même combinaison (Rc;cr,d ELS, Rc,d ELU) est
'   supérieure ou égale à la charge (« Non atteinte » sinon) ;
'   profondeur retenue : la plus grande des profondeurs requises du micropieu ;
'   vérification de la longueur : profondeur de reconnaissance (« Données du projet »,
'   D33) - profondeur retenue >= 5 B (NF P94-262, annexe M, M.1 (2) Note 2).
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
    L = LIGNE_PROFIL + MAX_COUCHES * NB_SUBDIV + 2
    gmin = CStr(ws.Range("D" & L + 2).Value)
    ws.Range(ws.Rows(L), ws.Rows(L + 300)).Clear
    VerificationProfondeur ws, gmin
End Sub

Private Sub VerificationProfondeur(ws As Worksheet, Optional gmin As String = "G")
    Dim wd As Worksheet, f As Worksheet, L As Long, n As Long, m As Long, i As Long, k As Long
    Dim r0 As Long, rd As Long, rt As Long, r As Long, dn As String, cap As String, z As String
    Dim charge As Variant, colCap As Variant, entete As Variant

    L = LIGNE_PROFIL + MAX_COUCHES * NB_SUBDIV + 2          ' titre du tableau
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

    ' En-têtes sur 3 lignes (titres en gras). Colonnes : A micropieu ; B à H descente de
    ' charges (Gmax, Gmin, Qmax, Qmin, eau EB, EH, EE) ; I à P combinaisons max / min ;
    ' Q à X diamètre B pour chaque cas ; Y à AF profondeurs requises ; AG profondeur
    ' retenue ; AH à AK vérification de la longueur (profondeur de reconnaissance -
    ' longueur >= 5 B)
    r0 = L + 7
    ws.Range("A" & r0).Value = "N° MICROPIEU"
    ws.Range("B" & r0).Value = "DESCENTE DE CHARGES (kN)"
    ws.Range("I" & r0).Value = "COMBINAISONS DE CHARGES VERTICALES (kN)"
    ws.Range("Q" & r0).Value = "DIAMÈTRE B (m)"
    ws.Range("Y" & r0).Value = "PROFONDEUR MINIMALE REQUISE (m)"
    ws.Range("AG" & r0).Value = "PROFONDEUR RETENUE (m)"
    ws.Range("AH" & r0).Value = "VÉRIFICATION DE LA LONGUEUR (profondeur de reconnaissance - longueur >= 5 B)"
    ws.Range("A" & r0 & ":A" & r0 + 2).Merge
    ws.Range("B" & r0 & ":H" & r0).Merge
    ws.Range("I" & r0 & ":P" & r0).Merge
    ws.Range("Q" & r0 & ":X" & r0).Merge
    ws.Range("Y" & r0 & ":AF" & r0).Merge
    ws.Range("AG" & r0 & ":AG" & r0 + 2).Merge
    ws.Range("AH" & r0 & ":AK" & r0).Merge
    ws.Range("B" & r0 + 1).Value = "G"
    ws.Range("D" & r0 + 1).Value = "Q"
    ws.Range("F" & r0 + 1).Value = "EAU"
    ws.Range("B" & r0 + 1 & ":C" & r0 + 1).Merge
    ws.Range("D" & r0 + 1 & ":E" & r0 + 1).Merge
    ws.Range("F" & r0 + 1 & ":H" & r0 + 1).Merge
    ws.Range("B" & r0 + 2 & ":H" & r0 + 2).Value = Array("Gmax", "Gmin", "Qmax", "Qmin", _
        "Eau EB", "Eau EH", "Eau EE")
    ws.Range("AH" & r0 + 1 & ":AK" & r0 + 1).Value = Array("Profondeur de reconnaissance (m)", _
        "Profondeur de reconnaissance - longueur (m)", "5 x B (m)", "Vérification")
    For k = 0 To 3
        ws.Range(ws.Cells(r0 + 1, 34 + k), ws.Cells(r0 + 2, 34 + k)).Merge
    Next k
    entete = Array("ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
    For k = 0 To 3
        For i = 0 To 2                       ' combinaisons, diamètres, profondeurs
            ws.Cells(r0 + 1, 9 + 8 * i + 2 * k).Value = entete(k)
            ws.Range(ws.Cells(r0 + 1, 9 + 8 * i + 2 * k), ws.Cells(r0 + 1, 10 + 8 * i + 2 * k)).Merge
            ws.Range(ws.Cells(r0 + 2, 9 + 8 * i + 2 * k), ws.Cells(r0 + 2, 10 + 8 * i + 2 * k)).Value = _
                Array("max", "min")
        Next i
    Next k

    ' Portance de chaque combinaison dans le tableau par pas de calcul : Rc;cr,d ELS
    ' caractéristique (N), ELS quasi permanent (P), Rc,d ELU fondamental (R), accidentel (T)
    ' Eau (charges de l'eau EB basses eaux, EH hautes eaux, EE eaux exceptionnelles, 0 sans
    ' nappe) ajoutée avec un coefficient 1 : EH aux combinaisons max, EB aux combinaisons
    ' min, EE à l'ELU accidentel max
    colCap = Array("N", "P", "R", "T")
    z = "$C$" & LIGNE_PROFIL & ":$C$" & LIGNE_PROFIL + MAX_COUCHES * NB_SUBDIV - 1
    For i = 1 To m
        r = r0 + 2 + i
        ws.Range("A" & r).Value = "MP" & i
        ws.Range("B" & r).Formula = "=N(" & dn & wd.Cells(rd + i - 1, 8 + n).Address & ")"
        ws.Range("C" & r).Formula = "=IF($D$" & L + 2 & "=""G"",B" & r & ",0)"
        ws.Range("D" & r).Formula = "=N(" & dn & wd.Cells(rd + i - 1, 9 + n).Address & ")"
        ws.Range("E" & r).Value = 0
        For k = 0 To 2                       ' eau EB, EH, EE de la descente de charges
            ws.Cells(r, 6 + k).Formula = "=N(" & dn & wd.Cells(rd + i - 1, 10 + n + k).Address & ")"
        Next k
        ws.Range("I" & r).Formula = "=B" & r & "+D" & r & "+G" & r
        ws.Range("J" & r).Formula = "=C" & r & "+E" & r & "+F" & r
        ws.Range("K" & r).Formula = "=B" & r & "+$D$" & L + 5 & "*D" & r & "+G" & r
        ws.Range("L" & r).Formula = "=C" & r & "+$D$" & L + 5 & "*E" & r & "+F" & r
        ws.Range("M" & r).Formula = "=1.35*B" & r & "+1.5*D" & r & "+G" & r
        ws.Range("N" & r).Formula = "=1*C" & r & "+1.5*E" & r & "+F" & r
        ws.Range("O" & r).Formula = "=B" & r & "+$D$" & L + 4 & "*D" & r & "+H" & r
        ws.Range("P" & r).Formula = "=C" & r & "+$D$" & L + 4 & "*E" & r & "+F" & r
        For k = 0 To 7
            ' Diamètre B du micropieu pour chaque cas (celui du tableau par pas de calcul)
            ws.Cells(r, 17 + k).Formula = "=$D$19"
            ' Profondeur minimale : 1re ligne où la portance (nombre) >= la charge
            charge = ws.Cells(r, 9 + k).Address(False, False)
            cap = "$" & colCap(k \ 2) & "$" & LIGNE_PROFIL & ":$" & colCap(k \ 2) & "$" & _
                LIGNE_PROFIL + MAX_COUCHES * NB_SUBDIV - 1
            ws.Cells(r, 25 + k).Formula = "=IFERROR(INDEX(" & z & ",MATCH(1,INDEX(ISNUMBER(" & cap & _
                ")*(" & cap & ">=" & charge & "),0),0)),""Non atteinte"")"
        Next k
        ws.Range("AG" & r).Formula = "=IF(COUNTIF(Y" & r & ":AF" & r & ",""Non atteinte"")>0," & _
            """Non atteinte"",MAX(Y" & r & ":AF" & r & "))"
        ' Vérification de la longueur : profondeur de reconnaissance (« Données du
        ' projet ») - longueur retenue >= 5 B
        ws.Range("AH" & r).Formula = "=IF(ISNUMBER(" & DONNEES & "$D$33)," & DONNEES & "$D$33,"""")"
        ws.Range("AI" & r).Formula = "=IF(AND(ISNUMBER(AH" & r & "),ISNUMBER(AG" & r & ")),AH" & r & _
            "-AG" & r & ","""")"
        ws.Range("AJ" & r).Formula = "=5*$D$19"
        ws.Range("AK" & r).Formula = "=IF(AI" & r & "="""",""-"",IF(AI" & r & ">=AJ" & r & _
            ",""Vérifiée"",""Non vérifiée""))"
    Next i
    rt = r0 + 2 + Application.Max(m, 1)

    ' Mise en forme
    With ws.Range("A" & r0 & ":AK" & rt)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A" & r0 & ":AK" & r0 + 2)
        .Font.Bold = True
        .Font.Size = 11
        .Interior.Color = VERT_ENTETE
        .WrapText = True
    End With
    If m > 0 Then
        ws.Range("A" & r0 + 3 & ":A" & rt).Font.Bold = True
        ws.Range("B" & r0 + 3 & ":AK" & rt).Interior.Color = JAUNE_CALCUL
        ws.Range("B" & r0 + 3 & ":P" & rt).NumberFormat = "0.0"
        ws.Range("Q" & r0 + 3 & ":X" & rt).NumberFormat = "0.000"
        ws.Range("Y" & r0 + 3 & ":AJ" & rt).NumberFormat = "0.00"
        ws.Range("AG" & r0 + 3 & ":AG" & rt & ",AK" & r0 + 3 & ":AK" & rt).Font.Bold = True
        ' Vérification en rouge si non vérifiée
        With ws.Range("AK" & r0 + 3 & ":AK" & rt).FormatConditions
            .Delete
            .Add(xlCellValue, xlEqual, "=""Non vérifiée""").Font.Color = RGB(192, 0, 0)
        End With
    End If
    ws.Columns("V:AK").ColumnWidth = 12
    ws.Rows(r0).RowHeight = 30
End Sub

' Graphique de la capacité portante en compression en fonction de la profondeur, à
' droite du tableau par pas de calcul : Rc;cr,d ELS caractéristique (colonne N),
' Rc;cr,d ELS quasi permanent (P) et Rc,d ELU fondamental (R) en abscisse, profondeur z
' (C) en ordonnée, croissante vers le bas. Les séries renvoient à des noms de la feuille
' (CP_z, CP_N, CP_P, CP_R) limités aux pas renseignés (OFFSET sur le nombre de z), pour
' que le graphique suive le nombre de couches et leurs épaisseurs.
Private Sub GraphiqueCapacite(ws As Worksheet)
    Dim co As ChartObject, f As String, nb As String, k As Long, cols As Variant, noms As Variant
    f = "'" & ws.Name & "'!"
    nb = "MAX(1,COUNT(" & f & "$C$" & LIGNE_PROFIL & ":$C$" & _
        LIGNE_PROFIL + MAX_COUCHES * NB_SUBDIV - 1 & "))"
    cols = Array("C", "N", "P", "R")
    noms = Array("Profondeur z", "ELS caractéristique - Rc;cr,d", "ELS quasi permanent - Rc;cr,d", _
        "ELU fondamental - Rc,d")
    For k = 0 To 3
        ws.Names.Add Name:=IIf(k = 0, "CP_z", "CP_" & cols(k)), _
            RefersTo:="=OFFSET(" & f & "$" & cols(k) & "$" & LIGNE_PROFIL & ",0,0," & nb & ")"
    Next k
    With ws.Range("AD" & LIGNE_PROFIL - 3)
        Set co = ws.ChartObjects.Add(.Left, .Top, 620, 480)
    End With
    With co.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        For k = 1 To 3
            With .SeriesCollection.NewSeries
                .Name = noms(k)
                .XValues = "=" & f & "CP_" & cols(k)
                .Values = "=" & f & "CP_z"
            End With
        Next k
        .HasTitle = True
        .ChartTitle.Text = "Capacité portante en compression en fonction de la profondeur"
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

' Valeur d'un tableau de la norme (zone) en ligne lig et colonne col ; vide si erreur
Private Function FormuleIndex(zone As String, lig As String, col As String) As String
    FormuleIndex = "=IFERROR(INDEX(" & zone & "," & lig & "," & col & "),"""")"
End Function

' alpha x fsol ; « - » si alpha n'est pas donné par la norme (catégories 17 et 18 :
' prendre la technique la plus proche)
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
    ws.Range("N" & r).Formula = "=IF(OR(B" & r & "="""",F" & r & "="""",N($D$23)=0),""""," & _
        "$D$23*F" & r & ")"
    ' qs,i = qs / (gamma R;d1 x gamma R;d2) ; 0 si le terrain est neutralisé
    ws.Range("O" & r).Formula = "=IF(B" & r & "="""","""",IF(Q" & r & "=""Oui"",0," & _
        "IFERROR(M" & r & "/(" & gamma & "),""-"")))"
    ws.Range("P" & r).Formula = "=IF(AND(ISNUMBER(O" & r & "),ISNUMBER(N" & r & "))," & _
        "IF(Q" & r & "=""Oui"",0,O" & r & "*N" & r & "),IF(B" & r & "="""","""",""-""))"
    ws.Range("Q" & r).Formula = cond & DONNEES & "$J$" & ld & ")"
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
