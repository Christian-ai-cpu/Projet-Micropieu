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
' Méthode pressiométrique (annexe F) : fsol = (a.pl* + b)(1 - exp(-c.pl*)), pl* en
' MPa, fsol en kPa ; qs = min(alpha pieu-sol x fsol ; qs max).
' Catégories de sol de la feuille « Données du projet » -> colonnes des tableaux :
'   tableaux à 5 colonnes (F.5.2.1, F.5.2.2) : argiles et sols intermédiaires -> 1re
'   colonne, sables et graves -> 2e, craies -> 3e, marnes -> 4e, roches -> 5e ;
'   tableau qs max à 6 colonnes : dans l'ordre de la liste des catégories de sol.

Private Const NOM_FEUILLE As String = "Calcul et vérification"
Private Const DONNEES As String = "'Données du projet'!"
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
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:Q").ColumnWidth = 12
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

' Bloc de l'effet de groupe (lignes 78 à 85) masqué quand « Effet de groupe » vaut Non
' dans « Données du projet ». Appelée à la création et quand ce choix change
' (ChangementDonnees, module DonneesProjet).
Public Sub AfficherGroupe(Optional ws As Worksheet)
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
    End If
    ws.Rows("78:85").Hidden = (ActiveWorkbook.Worksheets("Données du projet").Range("D" & _
        LIGNE_PIEU + 4).Value <> "Oui")
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
        .Font.Size = 12
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
