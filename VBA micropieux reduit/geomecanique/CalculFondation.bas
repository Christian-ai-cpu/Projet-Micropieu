Attribute VB_Name = "CalculFondation"
Option Explicit
Private Const COUL_TITRE As Long = 5898259
Private Const COUL_ENTETE As Long = 14337945
Private Const COUL_SAISIE As Long = 15201485
Private Const COUL_CALCUL As Long = 15590349
Private Const COUL_ALERTE As Long = 192
Private Const NOM_FEUILLE As String = "Calcul Portance"
Private Const DONNEES As String = "'Données du projet'!"
Private Const NB_SUBDIV As Long = 20
Private Const LIGNE_FROT As Long = 48
Private Const NB_MIN_LIGNES As Long = 40
Private Const LIGNE_PROFIL As Long = 91
Private Const BLOC_PROFIL As Long = 40
Private Const TABLEAUX As String = "Tableaux NF P94-262"
Private mCellule As String
Private mPas As Variant
Private mPhiManuel As Variant
Private mIncl As Variant, mAngle As Variant
Private mGeo As Variant, mGsn As Variant, mAmax As Variant, mAmin As Variant
Private mPenetro As Boolean
Private mReconstruction As Boolean
Public Sub CreerCalculVerification()
    Construire True
End Sub
Public Function ReconstructionMatriceEnCours() As Boolean
    ReconstructionMatriceEnCours = mReconstruction
End Function
Public Sub CreerCalculPortance()
    Construire True
End Sub
Public Sub RenommerFeuilleCalcul()
    Dim f As Worksheet, ancien As Worksheet
    For Each f In ActiveWorkbook.Worksheets
        If f.Name = NOM_FEUILLE Then Exit Sub
        If f.Name = "Calcul et v" & ChrW(233) & "rification" Then Set ancien = f
    Next f
    If Not ancien Is Nothing Then ancien.Name = NOM_FEUILLE
End Sub
Public Sub ActualiserCalculVerification()
    Dim ws As Worksheet
    On Error Resume Next
    RenommerFeuilleCalcul
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
    Application.EnableEvents = False
    On Error GoTo Erreur
    mPenetro = (ActiveWorkbook.Worksheets("Données du projet").Range(CELLULE_METHODE).Value = "Méthode pénétrométrique")
    etape = "feuille Tableaux NF P94-262"
    On Error Resume Next
    Set wt = ActiveWorkbook.Worksheets(TABLEAUX)
    On Error GoTo Erreur
    If wt Is Nothing Then
        Application.Run "CreerTableauxNFP94262"
        Application.ScreenUpdating = False
        Set wt = ActiveWorkbook.Worksheets(TABLEAUX)
    End If
    zAlpha = ZoneTableau(wt, "Tableau F.5.2.1", 0, 4, 5)
    If CStr(wt.Range(zAlpha).Cells(1, 1).Value) = "-" Then
        Application.Run "CreerTableauxNFP94262"
        Application.ScreenUpdating = False
    End If
    zAlpha = ZoneTableau(wt, "Tableau G.5.2.1", 0, 4, 6) & "|" & ZoneTableau(wt, "Tableau F.5.2.1", 0, 4, 5)
    zABC = ZoneTableau(wt, "Tableau G.5.2.2", 0, 4, 6) & "|" & ZoneTableau(wt, "Tableau F.5.2.2", 0, 4, 5)
    zQs = ZoneTableau(wt, "Tableau G.5.2.3", 0, 4, 6) & "|" & ZoneTableau(wt, "Tableau F.5.2.3", 0, 4, 6)
    etape = "création de la feuille"
    On Error Resume Next
    RenommerFeuilleCalcul
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo Erreur
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = NOM_FEUILLE
    End If
    mPhiManuel = LirePhiManuel(ws)
    LireActions ws
    mPas = ws.Range("A" & LIGNE_FROT & ":A" & LIGNE_FROT + MAX_COUCHES - 1).Formula
    ws.Cells.Clear
    ws.Columns.ColumnWidth = ws.StandardWidth
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop
    nb = DONNEES & "$" & Replace(CELLULE_NB_COUCHES, "C", "C$")
    lc = LIGNE_COUCHES
    lg = LIGNE_COUCHES + MAX_COUCHES - 1
    With ws.Range("B1:I1")
        .Merge
        .Value = "CALCUL ET VÉRIFICATION DES MICROPIEUX"
        .Font.Bold = True
        .Font.Color = COUL_TITRE
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    etape = "données du projet"
    SousTitre ws, 3, "Données du projet"
    Ligne ws, 4, "Type de pieu", "=" & DONNEES & "D" & LIGNE_PIEU, ""
    Ligne ws, 5, "Méthode de calcul", "=" & DONNEES & CELLULE_METHODE, ""
    etape = "caractéristiques de la fondation"
    SousTitre ws, 7, "Caractéristiques de la fondation"
    Ligne ws, 8, "Abréviation", "=" & DONNEES & "D" & LIGNE_PIEU + 11, ""
    Ligne ws, 9, "Norme de référence", "=" & DONNEES & "D" & LIGNE_PIEU + 8, ""
    Ligne ws, 10, "Classe de la fondation", "=" & DONNEES & "D" & LIGNE_PIEU + 9, ""
    Ligne ws, 11, "Catégorie de la fondation", "=" & DONNEES & "D" & LIGNE_PIEU + 10, ""
    cat = "$D$11"
    Ligne ws, 14, "Mode de fonctionnement", "=" & DONNEES & "D" & LIGNE_PIEU + 12, ""
    Ligne ws, 15, "Prise en compte de la résistance de pointe", "=" & DONNEES & "D" & LIGNE_PIEU + 13, ""
    Ligne ws, 16, "Type de mise en oeuvre", "=" & DONNEES & "D" & LIGNE_PIEU + 14, ""
    Ligne ws, 19, "Diamètre du micropieu B (B0 de « Données du projet »)", "=" & DONNEES & "D" & LIGNE_PIEU + 15, "m"
    Ligne ws, 20, "Hauteur totale de la fondation D (somme des épaisseurs)", "=IF(COUNT($F$39:$F$53)=0,"""",SUM($F$39:$F$53))", "m"
    Ligne ws, 21, "Hauteur de la couche porteuse h", Derniere("F"), "m"
    Ligne ws, 22, "Aire de la base de la fondation Ab = " & ChrW(960) & ".B² / 4", "=IF(N(D19)=0,"""",PI()*D19^2/4)", "m²"
    Ligne ws, 23, "Périmètre de la fondation Pe = " & ChrW(960) & ".B", "=IF(N(D19)=0,"""",PI()*D19)", "m"
    Ligne ws, 24, "Frottement négligé (épaisseur des terrains neutralisés)", "=IF(COUNTIF(OFFSET(" & DONNEES & "$M$" & lc & ",0,0," & nb & "),""Oui"")=0,""Non""," & "SUMIF(OFFSET(" & DONNEES & "$M$" & lc & ",0,0," & nb & "),""Oui""," & "OFFSET(" & DONNEES & "$F$" & lc & ",0,0," & nb & ")))", "m"
    ws.Range("D19:D24").NumberFormat = "0.000"
    etape = "facteurs partiels"
    SousTitre ws, 26, "Facteurs partiels de résistance (NF P94-262)"
    ws.Range("B27:E27").Value = Array("ELU (tableaux C.2.3.1 et C.2.3.2)", "", "Situations durables et transitoires", "Situations accidentelles")
    ws.Range("B28:E28").Value = Array("Pointe " & ChrW(947) & "b", "", 1.1, 1)
    ws.Range("B29:E29").Value = Array("Fût en compression " & ChrW(947) & "s", "", 1.1, 1)
    ws.Range("B30:E30").Value = Array("Total en compression " & ChrW(947) & "t", "", 1.1, 1)
    ws.Range("B31:E31").Value = Array("Fût en traction " & ChrW(947) & "s;t", "", 1.15, 1.05)
    ws.Range("B33:E33").Value = Array("ELS (tableaux 14.2.1.1 et 14.2.1.2)", "", "Combinaisons caractéristiques", "Combinaisons quasi permanentes")
    ws.Range("B34:E34").Value = Array("Fût en compression " & ChrW(947) & "cr", "", 0.9, 1.1)
    ws.Range("B35:E35").Value = Array("Fût en traction " & ChrW(947) & "s;cr", "", 1.1, 1.5)
    For i = 27 To 35
        If i <> 32 Then ws.Range("B" & i & ":C" & i).Merge
    Next i
    ws.Range("D28:E31,D34:E35").NumberFormat = "0.00"
    zF21 = ZoneTableau(wt, IIf(mPenetro, "Tableau G.2.1", "Tableau F.2.1"), 0, 1, 4)
    ligneF21 = "1"
    gamma = "INDEX(" & zF21 & "," & ligneF21 & ",IF($D$14=""Traction"",2,1))*INDEX(" & zF21 & "," & ligneF21 & ",IF($D$14=""Traction"",4,3))"
    etape = "frottement par couche"
    r = 37
    SousTitre ws, r, "Frottement axial unitaire et résistance par couche"
    ws.Range("B" & r).Formula = "=""Frottement axial unitaire et résistance par couche (""&LOWER($D$5)&"")""
    ws.Range("B" & r + 1 & ":Q" & r + 1).Value = Array("N°", "Nom de couche", "Catégorie de sol", "Profondeur (m/TN)", "hi (m)", "pl*,i (MPa)", "Courbe fsol", ChrW(945) & "pieu-sol", "fsol,i (kPa)", ChrW(945) & "pieu-sol x fsol,i (kPa)", "qs,max (kPa)", "qs = min (kPa)", "As,i = " & ChrW(960) & ".B.hi (m²)", "qs,i (kPa)", "Rs,i (kN)", "Terrain neutralisé")
    ws.Range("E" & r + 1).Formula = "=""Profondeur (""&" & DONNEES & "$E$29&"")"""
    If mPenetro Then
        ws.Range("G" & r + 1 & ":J" & r + 1).Value = Array("qc,i (MPa)", "Courbe fsol (G.5.2.2)", ChrW(945) & "pieu-sol (G.5.2.1)", "fsol,i = (a.qc + b)(1 - e^-c.qc) (kPa)")
        ws.Range("L" & r + 1).Value = "qs,max (G.5.2.3) (kPa)"
    Else
        ws.Range("G" & r + 1 & ":J" & r + 1).Value = Array("pl*,i (MPa)", "Courbe fsol (F.5.2.2)", ChrW(945) & "pieu-sol (F.5.2.1)", "fsol,i = (a.pl* + b)(1 - e^-c.pl*) (kPa)")
        ws.Range("L" & r + 1).Value = "qs,max (F.5.2.3) (kPa)"
    End If
    For i = 1 To MAX_COUCHES
        CoucheLigne ws, r + 1 + i, i, lc + i - 1, "(" & cat & "-16)", gamma, zAlpha, zABC, zQs
    Next i
    mCellule = ""
    lt = r + 1 + MAX_COUCHES
    With ws.Range("C" & r + 2 & ":D" & lt)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    etape = "valeurs caractéristiques"
    SousTitre ws, 56, "Valeurs caractéristiques des résistances"
    LigneS ws, 57, "Résistance de pointe (non prise en compte pour les micropieux)", "Rb,k", "=0", "kN"
    LigneS ws, 58, "Résistance de frottement = " & ChrW(931) & " Rs,i", "Rs,k", "=IF(COUNT($P$39:$P$53)=0,"""",SUM($P$39:$P$53))", "kN"
    LigneS ws, 60, "Portance (compression) = Rb,k + Rs,k", "Rc,k", "=IF(AND($D$14=""Compression"",ISNUMBER(D58)),N(D57)+D58,"""")", "kN"
    LigneS ws, 61, "Résistance en traction = Rs,k", "Rt,k", "=IF(ISNUMBER(D58),D58,"""")", "kN"
    LigneS ws, 62, "Charge de fluage en compression = 0,5 Rb,k + 0,7 Rs,k", "Rc;cr,k", "=IF(AND($D$14=""Compression"",ISNUMBER(D58)),0.5*N(D57)+0.7*D58,"""")", "kN"
    LigneS ws, 63, "Charge de fluage en traction = 0,7 Rt,k", "Rt;cr,k", "=IF(ISNUMBER(D61),0.7*D61,"""")", "kN"
    ws.Range("D57:D63").NumberFormat = "0.0"
    etape = "valeurs de calcul"
    SousTitre ws, 65, "Valeurs de calcul des résistances"
    ws.Range("B66:E66").Value = Array("Combinaison", "Valeur de calcul", "Symbole", "Valeur")
    For i = 0 To 3
        r = 67 + 2 * i
        c = Chr(Asc("D") + (i Mod 2))
        ws.Range("B" & r).Value = Array("ELU fondamental (C.2.3.1)", "ELU accidentel (C.2.3.2)", "ELS caractéristique (14.2.1.1)", "ELS quasi permanent (14.2.1.2)")(i)
        ws.Range("B" & r & ":B" & r + 1).Merge
        If i < 2 Then
            ws.Range("C" & r & ":D" & r + 1).Value = Array2("Portance de calcul en compression", "Rc,d", "Résistance de calcul en traction", "Rt,d")
            ws.Range("E" & r).Formula = "=IF(D60="""","""",N(D57)/" & c & "28+D58/" & c & "29)"
            ws.Range("E" & r + 1).Formula = "=IF(D61="""","""",D61/" & c & "31)"
        Else
            ws.Range("C" & r & ":D" & r + 1).Value = Array2("Charge de fluage de calcul en compression", "Rc;cr,d", "Charge de fluage de calcul en traction", "Rt;cr,d")
            ws.Range("E" & r).Formula = "=IF(D62="""","""",D62/" & c & "34)"
            ws.Range("E" & r + 1).Formula = "=IF(D63="""","""",D63/" & c & "35*IF($D$76=""Non"",0.15,1))"
        End If
    Next i
    ws.Range("E67:E74").NumberFormat = "0.0"" kN""
    Ligne ws, 76, "Essai de contrôle (sinon Rt;cr,d x 0,15)", "=" & DONNEES & "D" & LIGNE_PIEU + 19, ""
    etape = "effet de groupe"
    SousTitre ws, 78, "Effet de groupe - coefficient d'efficacité Ce (NF P94-262, J.2)"
    Ligne ws, 79, "Effet de groupe", "=" & DONNEES & "D" & LIGNE_PIEU + 4, ""
    Ligne ws, 80, "Nombre de lignes de pieux m", "=IF(" & DONNEES & "D" & LIGNE_PIEU + 5 & "=""""," & """""," & DONNEES & "D" & LIGNE_PIEU + 5 & ")", ""
    Ligne ws, 81, "Nombre de pieux par ligne n", "=IF(" & DONNEES & "D" & LIGNE_PIEU + 6 & "=""""," & """""," & DONNEES & "D" & LIGNE_PIEU + 6 & ")", ""
    Ligne ws, 82, "Entraxe des pieux d", "=IF(" & DONNEES & "D" & LIGNE_PIEU + 7 & "=""""," & """""," & DONNEES & "D" & LIGNE_PIEU + 7 & ")", "m"
    Ligne ws, 83, "d / B", "=IF(OR(N(D82)=0,N($D$19)=0),"""",D82/$D$19)", ""
    Ligne ws, 84, "Cd = 1 - (1 + d/B) / 4  (1 <= d/B <= 3)", "=IF(D83="""","""",IF(D83>=3,0,IF(D83>=1,1-(1+D83)/4,""-"")))", ""
    Ligne ws, 85, "Ce = 1 si d >= 3B, sinon 1 - Cd (2 - (1/m + 1/n))", "=IF(D79<>""Oui"",1,IF(D83="""","""",IF(D83>=3,1,IF(D83<1,""-""," & "IF(OR(N(D80)=0,N(D81)=0),"""",1-D84*(2-(1/D80+1/D81)))))))", ""
    ws.Range("D80:D81").NumberFormat = "0"
    ws.Range("D82:D85").NumberFormat = "0.00"
    ws.Range("B85:D85").Font.Bold = True
    etape = "pas de calcul"
    ws.Range("A38").Value = "Pas de calcul (m) (saisie libre)"
    ws.Range("R38").Value = "Nombre de subdivisions k = hi / pas"
    ws.Range("ZA38:ZC38").Value = Array(ChrW(931) & " k", "z bas (m)", ChrW(931) & " Rs,i (kN)")
    ws.Range("ZF38").Value = "Pas effectif (m)"
    For r = 39 To 53
        If IsArray(mPas) Then
            vb = mPas(r - 38, 1)
            If CStr(vb) <> "" And Left(CStr(vb), 1) <> "=" Then
                If ValeurPas(vb) > 0 Then ws.Range("A" & r).Value = ValeurPas(vb)
            End If
        End If
        ws.Range("ZG" & r).Formula = "=IF(ISNUMBER(A" & r & "),A" & r & ",IFERROR(--SUBSTITUTE(A" & r & ","","",""."")," & "IFERROR(--SUBSTITUTE(A" & r & ",""."","",""),0)))"
        ws.Range("R" & r).Formula = "=IF(B" & r & "="""","""",IF(N(ZG" & r & ")>0," & "MAX(1,ROUNDUP(ROUND(N(F" & r & ")/ZG" & r & ",6),0)),1))"
        ws.Range("ZF" & r).Formula = "=IF(B" & r & "="""","""",IF(N(ZG" & r & ")>0,MIN(ZG" & r & ",N(F" & r & ")),N(F" & r & ")))"
        ws.Range("ZA" & r).Formula = "=IF(B" & r & "="""","""",SUM($R$39:R" & r & "))"
        ws.Range("ZB" & r).Formula = "=IF(B" & r & "="""","""",SUM($F$39:F" & r & "))"
        ws.Range("ZC" & r).Formula = "=IF(B" & r & "="""","""",SUM($P$39:P" & r & "))"
    Next r
    ws.Range("A39:A53").NumberFormat = "General"
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
        .Interior.Color = COUL_ENTETE
    End With
    ws.Range("A39:A53").Interior.Color = COUL_SAISIE
    ws.Range("R39:R53").Interior.Color = COUL_CALCUL
    mPas = Empty
    ws.Calculate
    nPas = Application.Max(ws.Range("ZA39:ZA53"))
    If nPas < 1 Then nPas = 1
    ws.Range("ZA" & LIGNE_PROFIL - 3).Value = nPas
    etape = "capacité portante par pas de calcul"
    lt = LIGNE_PROFIL + nPas - 1
    SousTitre ws, LIGNE_PROFIL - 3, "Capacité portante des micropieux en fonction de la " & "profondeur (par pas de calcul, valeurs cumulées) - diamètre B0"
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
    ws.Range("B" & r + 1 & ":Y" & r + 1).Value = Array("N° de couche", "Profondeur z (m)", "Cote", ChrW(934) & " (m)", "Pas de calcul (m)", ChrW(947) & "R;d1 compression", ChrW(947) & "R;d1 traction", ChrW(947) & "R;d2", "qs (kPa)", "Rs,k (kN)", "Rt,k (kN)", "ple* (MPa)", "h (m)", "kp max", "Def (m)", "Rb,k (kN)", "Rc;cr,d (kN)", "Rt;cr,d (kN)", "Rc;cr,d (kN)", "Rt;cr,d (kN)", "Rc,d (kN)", "Rt,d (kN)", "Rc,d (kN)", "Rt,d (kN)")
    ws.Range("D" & r + 1).Formula = "=""Cote (""&" & DONNEES & "$E$29&"")"""
    r = LIGNE_PROFIL
    c = "(ROW()-" & LIGNE_PROFIL - 1 & ")"
    ws.Range("$ZD" & r).Formula = "=IF(OR(COUNT($ZA$39:$ZA$53)=0," & c & ">MAX($ZA$39:$ZA$53)),""""," & "COUNTIF($ZA$39:$ZA$53,""<""&" & c & ")+1)"
    ws.Range("$ZE" & r).Formula = "=IF($ZD" & r & "="""",""""," & c & "-IF($ZD" & r & "=1,0," & "INDEX($ZA$39:$ZA$53,$ZD" & r & "-1)))"
    ws.Range("B" & r).Formula = "=IF($ZD" & r & "="""","""",INDEX($B$39:$B$53,$ZD" & r & "))"
    ws.Range("E" & r).Formula = "=IF($ZD" & r & "="""",""""," & DONNEES & "$G$" & LIGNE_PIEU + 9 & ")"
    ws.Range("F" & r).Formula = "=IF($ZD" & r & "="""","""",N" & r & "-IF($ZE" & r & "=1,0,N(N" & r - 1 & ")))"
    ws.Range("G" & r).Formula = "=IF($ZD" & r & "="""","""",INDEX(" & zF21 & ",1,1))"
    ws.Range("H" & r).Formula = "=IF($ZD" & r & "="""","""",INDEX(" & zF21 & ",1,2))"
    ws.Range("I" & r).Formula = "=IF($ZD" & r & "="""","""",INDEX(" & zF21 & ",1,3))"
    ws.Range("N" & r).Formula = "=IF($ZD" & r & "="""","""",MIN($ZE" & r & "*INDEX($ZF$39:$ZF$53,$ZD" & r & "),INDEX($F$39:$F$53,$ZD" & r & ")))"
    ws.Range("C" & r).Formula = "=IF($ZD" & r & "="""","""",IF($ZD" & r & "=1,0,INDEX($ZB$39:$ZB$53,$ZD" & r & "-1))+N" & r & ")"
    ws.Range("D" & r).Formula = "=IF(C" & r & "="""","""",IF(ISNUMBER(" & DONNEES & "$D$29)," & DONNEES & "$D$29-C" & r & ",-C" & r & "))"
    ws.Range("J" & r).Formula = "=IF($ZD" & r & "="""","""",INDEX($M$39:$M$53,$ZD" & r & "))"
    ws.Range("K" & r).Formula = "=IF($ZD" & r & "="""","""",N(K" & r - 1 & ")+IF(INDEX($Q$39:$Q$53,$ZD" & r & ")=""Oui"",0,N(J" & r & ")*PI()*E" & r & "*F" & r & "/(G" & r & "*I" & r & ")))"
    ws.Range("L" & r).Formula = "=IF($ZD" & r & "="""","""",N(L" & r - 1 & ")+IF(INDEX($Q$39:$Q$53,$ZD" & r & ")=""Oui"",0,N(J" & r & ")*PI()*E" & r & "*F" & r & "/(H" & r & "*I" & r & ")))"
    ws.Range("M" & r).Formula = "=IF($ZD" & r & "="""","""",""-"")"
    ws.Range("O" & r).Formula = "=M" & r
    ws.Range("P" & r).Formula = "=M" & r
    ws.Range("Q" & r).Formula = "=IF($ZD" & r & "="""","""",0)"
    For i = 0 To 1
        c = Chr(Asc("D") + i)
        ws.Range(Chr(Asc("R") + 2 * i) & r).Formula = "=IF(K" & r & "="""","""",(0.5*Q" & r & "+0.7*K" & r & ")/$" & c & "$34)"
        ws.Range(Chr(Asc("S") + 2 * i) & r).Formula = "=IF(L" & r & "="""","""",0.7*L" & r & "/$" & c & "$35*IF($D$76=""Non"",0.15,1))"
        ws.Range(Chr(Asc("V") + 2 * i) & r).Formula = "=IF(K" & r & "="""","""",Q" & r & "/$" & c & "$28+K" & r & "/$" & c & "$29)"
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
        .Interior.Color = COUL_ENTETE
        .WrapText = True
    End With
    ws.Columns("ZA:ZG").Hidden = True
    etape = "capacité portante des autres diamètres"
    For kd = 1 To NB_DIAMETRES - 1
        If True Then
            ws.Range(ws.Cells(LIGNE_PROFIL - 3, 2), ws.Cells(lt, 25)).Copy Destination:=ws.Cells(LIGNE_PROFIL - 3, 2 + kd * BLOC_PROFIL)
            ws.Range(ws.Cells(LIGNE_PROFIL, 2 + kd * BLOC_PROFIL), ws.Cells(lt, 25 + kd * BLOC_PROFIL)).Replace What:="$G$" & LIGNE_PIEU + 9, Replacement:="$G$" & LIGNE_PIEU + 9 + kd, LookAt:=xlPart, MatchCase:=False
            ws.Cells(LIGNE_PROFIL - 3, 2 + kd * BLOC_PROFIL).Value = "Capacité portante des micropieux " & "en fonction de la profondeur (par pas de calcul, valeurs cumulées) - diamètre B" & kd
            ws.Range(ws.Columns(2 + kd * BLOC_PROFIL), ws.Columns(25 + kd * BLOC_PROFIL)).ColumnWidth = 12
        End If
    Next kd
    Application.CutCopyMode = False
    etape = "mise en forme"
    r = 37
    lt = r + 1 + MAX_COUCHES
    With ws.Range("B4:E5,B8:E11,B14:E16,B19:E24,B27:E31,B33:E35,B" & r + 1 & ":Q" & lt & ",B57:E58,B60:E63,B66:E74,B76:E76,B79:E85")
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("B27:E27,B33:E33,B66:E66,B" & r + 1 & ":Q" & r + 1)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = COUL_ENTETE
    End With
    ws.Range("D4:D5,D8:D11,D14:D16,D19:D24,D28:E31,D34:E35,D57:D58,D60:D63,E67:E74,D76,D79:D85").Interior.Color = COUL_CALCUL
    ws.Range("D4:F35,C57:E63,B67:B74,D67:E74,D76:E85,B" & r + 2 & ":B" & lt & ",E" & r + 2 & ":Q" & lt ).HorizontalAlignment = xlCenter
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:Y").ColumnWidth = 12
    etape = "déplacement de l'effet de groupe"
    ws.Range("A78:ZG86").Cut
    ws.Range("A37").Insert Shift:=xlDown
    Application.CutCopyMode = False
    etape = "vérification de la profondeur des micropieux"
    VerificationProfondeur ws
    etape = "graphique de la capacité portante"
    GraphiqueCapacite ws, 0
    For kd = 1 To NB_DIAMETRES - 1
        If ws.Cells(LIGNE_PROFIL, 3 + kd * BLOC_PROFIL).HasFormula Then GraphiqueCapacite ws, kd
    Next kd
    AfficherGroupe ws
    AfficherCouchesCalcul ws
    InstallerCodeFeuille ws
    ws.Cells.Font.Name = "CheGuevara Text Sans"
    AjusterTextes ws
    If activer Then ws.Activate
Fin:
    Application.EnableEvents = evenements
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    MsgBox "Erreur pendant la création de la feuille « Calcul Portance », partie « " & etape & " »" & IIf(mCellule = "", "", ", cellule " & mCellule) & " : " & Err.Description & " (n° " & Err.Number & ")", vbExclamation
    Resume Fin
End Sub
Public Sub AfficherGroupe(Optional ws As Worksheet)
    If ws Is Nothing Then
        On Error Resume Next
        RenommerFeuilleCalcul
        Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
    End If
    ws.Rows("37:45").Hidden = (ActiveWorkbook.Worksheets("Données du projet").Range("D" & LIGNE_PIEU + 4).Value <> "Oui")
End Sub
Public Sub ActualiserVerificationProfondeur()
    Dim ws As Worksheet, L As Long, gmin As String
    On Error Resume Next
    RenommerFeuilleCalcul
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If mReconstruction Then Exit Sub
    L = LigneVerif(ws)
    gmin = CStr(ws.Range("D" & L + 2).Value)
    mPhiManuel = LirePhiManuel(ws)
    LireActions ws
    ws.Range(ws.Rows(L), ws.Rows(L + 300)).Clear
    VerificationProfondeur ws, gmin
End Sub
Private Function LirePhiManuel(ws As Worksheet) As Variant
    LirePhiManuel = LireSaisie(ws, ChrW(934) & " MANUEL")
End Function
Private Function LireSaisie(ws As Worksheet, entete As String) As Variant
    Dim r As Long, i As Long, t() As Variant, c As Long, cm As Long
    r = LigneVerif(ws) + 10
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
Private Sub LireActions(ws As Worksheet)
    mGeo = LireSaisie(ws, "Geo max")
    mGsn = LireSaisie(ws, "Gsn")
    mAmax = LireSaisie(ws, "A max")
    mAmin = LireSaisie(ws, "A min")
    mIncl = LireSaisie(ws, "Incliné")
    mAngle = LireSaisie(ws, "Angle (°)")
End Sub
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
    Dim fx As String, kd As Long, cE As Long, tC As Long, tP As Long, cGd As Long, cEd As Long
    Dim lMat As Long
    L = LigneVerif(ws)
    For Each f In ActiveWorkbook.Worksheets
        If TrouverLigneMatrice(f) > 0 Then
            Set wd = f
            lMat = TrouverLigneMatrice(f)
            Exit For
        End If
    Next f
    With ws.Range("A" & L)
        .Value = "VÉRIFICATION DE LA PROFONDEUR DES MICROPIEUX (DESCENTE DE CHARGES)"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = COUL_TITRE
    End With
    Ligne ws, L + 1, "Feuille de descente de charges", IIf(wd Is Nothing, "introuvable", ""), ""
    If Not wd Is Nothing Then ws.Range("D" & L + 1).Value = wd.Name
    Ligne ws, L + 2, "Gmin pris égal à (G ou 0)", "", ""
    With ws.Range("D" & L + 2)
        .Validation.Add xlValidateList, xlValidAlertStop, , "G,0"
        .Value = IIf(gmin = "0", "0", "G")
        .Interior.Color = COUL_SAISIE
    End With
    Ligne ws, L + 3, "Catégorie du bâtiment", "=" & DONNEES & "C14", ""
    Ligne ws, L + 4, ChrW(968) & "1 (combinaison accidentelle)", "=IFERROR(CHOOSE(MATCH(LEFT(D" & L + 3 & ",1),{""A"",""B"",""C"",""D"",""E"",""F"",""G""},0),0.5,0.5,0.7,0.7,0.9,0.7,0.5),0.5)", ""
    Ligne ws, L + 5, ChrW(968) & "2 (combinaison quasi permanente)", "=IFERROR(CHOOSE(MATCH(LEFT(D" & L + 3 & ",1),{""A"",""B"",""C"",""D"",""E"",""F"",""G""},0),0.3,0.3,0.6,0.6,0.8,0.6,0.3),0.3)", ""
    ws.Range("D" & L + 1 & ":D" & L + 5).HorizontalAlignment = xlCenter
    ws.Range("D" & L + 3 & ":D" & L + 5).Interior.Color = COUL_CALCUL
    ws.Range("D" & L + 4 & ":D" & L + 5).NumberFormat = "0.00"
    With ws.Range("B" & L + 1 & ":E" & L + 5)
        .Borders.LineStyle = xlContinuous
        .Font.Bold = True
    End With
    reco = DONNEES & "$D$33"
    colCap = Array("R", "T", "V", "X")
    phiT = DONNEES & "$G$" & LIGNE_PIEU + 9 & ":$G$" & LIGNE_PIEU + 8 + NB_DIAMETRES
    If wd Is Nothing Then
        ws.Range("A" & L + 7).Value = "Créez la matrice des micropieux dans la feuille de descente de " & "charges, puis relancez CreerCalculVerification."
        Exit Sub
    End If
    Do While InStr("|Mur |Pot |", "|" & Left(CStr(wd.Range("B" & lMat + 6 + n).Value), 4) & "|") > 0
        n = n + 1
    Loop
    rd = lMat + 8 + n + 3
    Do While Left(CStr(wd.Range("B" & rd + m).Value), 2) = "MP"
        m = m + 1
    Loop
    k = Val(wd.Range("K1").Value)
    If k >= 1 And k <> m And Not mReconstruction Then
        mReconstruction = True
        On Error Resume Next
        Application.Run "ConstruireMatrice", wd
        On Error GoTo 0
        mReconstruction = False
        n = 0: m = 0
        Do While InStr("|Mur |Pot |", "|" & Left(CStr(wd.Range("B" & lMat + 6 + n).Value), 4) & "|") > 0
            n = n + 1
        Loop
        rd = lMat + 8 + n + 3
        Do While Left(CStr(wd.Range("B" & rd + m).Value), 2) = "MP"
            m = m + 1
        Loop
    End If
    dn = "'" & wd.Name & "'!"
    cGd = 0: cEd = 0
    For c = 3 To 250
        If CStr(wd.Cells(lMat + 8 + n, c).Value) = "POIDS TOTAL ÉLÉMENTAIRE DDC" Then cGd = c
        If Left(CStr(wd.Cells(lMat + 8 + n, c).Value), 15) = "CHARGES DE L'EA" Then cEd = c
    Next c
    If cGd = 0 Then cGd = 8 + n
    eau = (CStr(ActiveWorkbook.Worksheets("Données du projet").Range("D34").Value) = "Oui") And (cEd > 0)
    cE = 8
    cC = 14
    cD = cC + 8
    cP = cD + 7
    cI = cP + 8
    cR = cI + 1
    cA = cR + 1
    cRe = cR + 2
    cZ = cR + 3
    cL = cR + 5
    cF = cL + 3
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
    ws.Cells(r0, cL).Value = "VÉRIFICATION DE LA LONGUEUR (profondeur de reconnaissance - D >= 5 " & ChrW(934) & ")"
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
    ws.Range("H" & r0 + 1).Value = "EAU"
    ws.Range("H" & r0 + 1 & ":J" & r0 + 1).Merge
    ws.Range("H" & r0 + 2 & ":J" & r0 + 2).Value = Array("EB", "EH", "EE")
    ws.Range("K" & r0 + 1).Value = "Gsn"
    ws.Range("K" & r0 + 1 & ":K" & r0 + 2).Merge
    ws.Range("L" & r0 + 1).Value = "A (accidentel)"
    ws.Range("L" & r0 + 1 & ":M" & r0 + 1).Merge
    ws.Range("L" & r0 + 2 & ":M" & r0 + 2).Value = Array("A max", "A min")
    ws.Range(ws.Cells(r0 + 1, cD), ws.Cells(r0 + 1, cD + 6)).Value = Array(ChrW(934) & " CARA", ChrW(934) & " QP", ChrW(934) & " FOND", ChrW(934) & " ACC", ChrW(934) & " PIEU", ChrW(934) & " MANUEL", ChrW(934) & " RETENU")
    ws.Cells(r0 + 2, cR).Value = "D"
    ws.Range(ws.Cells(r0 + 1, cL), ws.Cells(r0 + 1, cF)).Value = Array("Profondeur de reconnaissance (m)", "Profondeur de reconnaissance - D réelle (m)", "5 x " & ChrW(934) & " retenu (m)", "Vérification")
    For k = 0 To 6
        ws.Range(ws.Cells(r0 + 1, cD + k), ws.Cells(r0 + 2, cD + k)).Merge
    Next k
    For k = 0 To 3
        ws.Range(ws.Cells(r0 + 1, cL + k), ws.Cells(r0 + 2, cL + k)).Merge
    Next k
    entete = Array("ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
    For k = 0 To 3
        For i = 0 To 1
            c = IIf(i = 0, cC, cP) + 2 * k
            ws.Cells(r0 + 1, c).Value = entete(k)
            ws.Range(ws.Cells(r0 + 1, c), ws.Cells(r0 + 1, c + 1)).Merge
            ws.Range(ws.Cells(r0 + 2, c), ws.Cells(r0 + 2, c + 1)).Value = IIf(i = 0, Array("Vmax", "Vmin"), Array("Dmax", "Dmin"))
        Next i
    Next k
    Commentaire ws.Range("B" & r0 + 1), "Gstr : charges permanentes de la structure (descente de " & "charges : poids des murs, planchers, massifs, longrines...)"
    Commentaire ws.Range("B" & r0 + 2), "Gstr max : charge permanente de la structure (valeur " & "défavorable, coefficient " & ChrW(947) & "G,sup)"
    Commentaire ws.Range("C" & r0 + 2), "Gstr min : charge permanente favorable (G ou 0 selon le " & "choix en tête du tableau, coefficient " & ChrW(947) & "G,inf)"
    Commentaire ws.Range("D" & r0 + 1), "Q : charges d'exploitation (descente de charges)"
    Commentaire ws.Range("D" & r0 + 2), "Qmax : charge d'exploitation défavorable"
    Commentaire ws.Range("E" & r0 + 2), "Qmin : charge d'exploitation favorable, prise égale à 0"
    Commentaire ws.Range("F" & r0 + 1), "Geo : action géotechnique permanente (poussée des terres) " & "transmise au micropieu"
    Commentaire ws.Range("F" & r0 + 2), "Geo max : valeur de la poussée des terres (à saisir)"
    Commentaire ws.Range("G" & r0 + 2), "Geo min : prise égale à Geo max"
    Commentaire ws.Range("H" & r0 + 1), "Charges de l'eau (descente de charges, 0 sans nappe " & "phréatique), comptées seulement dans les combinaisons min (Vmin) : EH pour l'ELS " & "caractéristique et l'ELU fondamental, EB pour l'ELS quasi permanent, EE pour l'ELU accidentel"
    Commentaire ws.Range("H" & r0 + 2), "EB : niveau des basses eaux"
    Commentaire ws.Range("I" & r0 + 2), "EH : niveau des hautes eaux"
    Commentaire ws.Range("J" & r0 + 2), "EE : niveau des eaux exceptionnelles"
    Commentaire ws.Range("K" & r0 + 1), "Gsn : frottement négatif (action permanente due au " & "tassement du sol autour du micropieu), compté aux combinaisons max"
    Commentaire ws.Range("L" & r0 + 1), "A : action accidentelle (choc, explosion...), " & "combinaison ELU accidentelle seulement"
    Commentaire ws.Range("L" & r0 + 2), "A max : action accidentelle de la combinaison max"
    Commentaire ws.Range("M" & r0 + 2), "A min : action accidentelle de la combinaison min"
    entete = Array("ELS caractéristique (NF EN 1990, 6.14b)", "ELS quasi permanent (6.16b)", "ELU fondamental (6.10)", "ELU accidentel (6.11b)")
    For k = 0 To 3
        Commentaire ws.Cells(r0 + 1, cC + 2 * k), "Combinaison " & entete(k) & " : coefficients " & "dans le tableau des combinaisons sous la vérification"
        Commentaire ws.Cells(r0 + 2, cC + 2 * k), "Vmax : charge verticale de calcul maximale " & "(actions défavorables)"
        Commentaire ws.Cells(r0 + 2, cC + 2 * k + 1), "Vmin : charge verticale de calcul minimale " & "(actions favorables)"
        Commentaire ws.Cells(r0 + 1, cD + k), ChrW(934) & " " & Split("CARA,QP,FOND,ACC", ",")(k) & " : plus petit diamètre (B0 à B3) dont la portance maximale couvre la charge de la " & "combinaison " & entete(k)
        Commentaire ws.Cells(r0 + 2, cP + 2 * k), "Dmax : profondeur minimale pour reprendre Vmax"
        Commentaire ws.Cells(r0 + 2, cP + 2 * k + 1), "Dmin : profondeur minimale pour reprendre Vmin"
    Next k
    Commentaire ws.Cells(r0 + 1, cD + 4), ChrW(934) & " PIEU : le plus grand des diamètres des " & "combinaisons"
    Commentaire ws.Cells(r0 + 1, cD + 5), ChrW(934) & " MANUEL : diamètre imposé (à saisir)"
    Commentaire ws.Cells(r0 + 1, cD + 6), ChrW(934) & " RETENU : max(" & ChrW(934) & " pieu ; " & ChrW(934) & " manuel)"
    Commentaire ws.Cells(r0 + 2, cR), "D : profondeur retenue, la plus grande des profondeurs requises"
    Commentaire ws.Cells(r0 + 1, cI), "Incliné : Oui si le micropieu est incliné (liste) ; les " & "colonnes de l'angle et de la profondeur réelle s'affichent alors"
    Commentaire ws.Cells(r0 + 1, cA), "Angle d'inclinaison du micropieu sur la verticale (à saisir)"
    Commentaire ws.Cells(r0 + 2, cRe), "Profondeur réelle : longueur du micropieu " & "incliné = D / cos " & ChrW(945) & " (" & ChrW(945) & " : angle sur la verticale) ; D si " & "le micropieu n'est pas incliné"
    Commentaire ws.Cells(r0 + 1, cZ), "Ztête : altitude de la tête du micropieu = cote de l'arase " & "supérieure de la fondation (« Données du projet »)"
    Commentaire ws.Cells(r0 + 1, cZ + 1), "Zpointe : altitude de la pointe du micropieu = Ztête - D " & "(profondeur retenue)"
    Commentaire ws.Cells(r0 + 1, cL), "Profondeur de reconnaissance (« Données du projet »)"
    Commentaire ws.Cells(r0 + 1, cL + 2), "5 " & ChrW(934) & " : longueur reconnue sous la base " & "(NF P94-262, annexe M, M.1)"
    tC = r0 + 2 + Application.Max(m, 1) + 3
    tP = tC + 8
    With ws.Range("A" & tC)
        .Value = "COEFFICIENTS DES COMBINAISONS D'ACTIONS (NF EN 1990)"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COUL_TITRE
    End With
    ws.Range("B" & tC + 1 & ":I" & tC + 1).Value = Array("Combinaison", "Expression (NF EN 1990)", ChrW(947) & "G,sup (Gstr max, Geo, Gsn)", ChrW(947) & "G,inf (Gstr min, Geo min)", ChrW(947) & "Q", ChrW(968) & " appliqué à Q", "Valeur de " & ChrW(968), ChrW(947) & "A (accidentel)")
    ws.Range("B" & tC + 2 & ":G" & tC + 2).Value = Array("ELS caractéristique", "6.14b (tableau A1.4)", 1, 1, 1, "- (Q dominante)")
    ws.Range("B" & tC + 3 & ":G" & tC + 3).Value = Array("ELS quasi permanent", "6.16b (tableau A1.4)", 1, 1, 1, ChrW(968) & "2")
    ws.Range("B" & tC + 4 & ":G" & tC + 4).Value = Array("ELU fondamental", "6.10 (tableau A1.2(B))", 1.35, 1, 1.5, "- (Q dominante)")
    ws.Range("B" & tC + 5 & ":G" & tC + 5).Value = Array("ELU accidentel", "6.11b (tableau A1.3)", 1, 1, 1, ChrW(968) & "1")
    ws.Range("H" & tC + 2).Value = 1
    ws.Range("H" & tC + 3).Formula = "=IFERROR(VLOOKUP(LEFT($D$" & L + 3 & ",1),$B$" & tP + 2 & ":$F$" & tP + 9 & ",5,FALSE),0.3)"
    ws.Range("H" & tC + 4).Value = 1
    ws.Range("H" & tC + 5).Formula = "=IFERROR(VLOOKUP(LEFT($D$" & L + 3 & ",1),$B$" & tP + 2 & ":$F$" & tP + 9 & ",4,FALSE),0.5)"
    ws.Range("I" & tC + 2 & ":I" & tC + 5).Value = Application.Transpose(Array(0, 0, 0, 1))
    ws.Range("A" & tP).Value = "Tableau A1.1 - Valeurs recommandées des coefficients " & ChrW(968) & " pour les bâtiments (NF EN 1990)"
    ws.Range("A" & tP).Font.Bold = True
    ws.Range("B" & tP + 1 & ":F" & tP + 1).Value = Array("Catégorie", "Charges d'exploitation", ChrW(968) & "0", ChrW(968) & "1", ChrW(968) & "2")
    ws.Range("B" & tP + 2 & ":F" & tP + 9).Value = Array2D(Array( Array("A", "Habitation, zones résidentielles", 0.7, 0.5, 0.3), Array("B", "Bureaux", 0.7, 0.5, 0.3), Array("C", "Lieux de réunion", 0.7, 0.7, 0.6), Array("D", "Commerces", 0.7, 0.7, 0.6), Array("E", "Stockage", 1, 0.9, 0.8), Array("F", "Zone de trafic, véhicules <= 30 kN", 0.7, 0.7, 0.6), Array("G", "Zone de trafic, véhicules de 30 à 160 kN", 0.7, 0.5, 0.3), Array("H", "Toits", 0, 0, 0)))
    With ws.Range("B" & tC + 1 & ":I" & tC + 5 & ",B" & tP + 1 & ":F" & tP + 9)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("B" & tC + 1 & ":I" & tC + 1 & ",B" & tP + 1 & ":F" & tP + 1)
        .Font.Bold = True
        .Interior.Color = COUL_ENTETE
    End With
    ws.Range("D" & tC + 2 & ":I" & tC + 5 & ",D" & tP + 2 & ":F" & tP + 9).NumberFormat = "0.00"
    ws.Range("D" & tC + 2 & ":I" & tC + 5).Interior.Color = COUL_CALCUL
    ws.Range("D" & L + 4).Formula = "=$H$" & tC + 5
    ws.Range("D" & L + 5).Formula = "=$H$" & tC + 3
    z = "$C$" & LIGNE_PROFIL & ":$C$" & LIGNE_PROFIL + NbPas(ws) - 1
    For i = 1 To m
        r = r0 + 2 + i
        ws.Range("A" & r).Value = "MP" & i
        ws.Range("B" & r).Formula = "=N(" & dn & wd.Cells(rd + i - 1, cGd).Address & ")"
        ws.Range("C" & r).Formula = "=IF($D$" & L + 2 & "=""G"",B" & r & ",0)"
        ws.Range("D" & r).Formula = "=N(" & dn & wd.Cells(rd + i - 1, cGd + 1).Address & ")"
        ws.Range("E" & r).Value = 0
        ws.Range("F" & r).Value = Saisie(mGeo, i)
        ws.Range("G" & r).Formula = "=F" & r
        ws.Range("K" & r).Value = Saisie(mGsn, i)
        ws.Range("L" & r).Value = Saisie(mAmax, i)
        ws.Range("M" & r).Value = Saisie(mAmin, i)
        For k = 0 To 2
            If eau Then
                ws.Cells(r, cE + k).Formula = "=N(" & dn & wd.Cells(rd + i - 1, cEd + k).Address & ")"
            Else
                ws.Cells(r, cE + k).Value = 0
            End If
        Next k
        For k = 0 To 3
            c = tC + 2 + k
            ws.Cells(r, cC + 2 * k).Formula = "=$D$" & c & "*(B" & r & "+F" & r & "+K" & r & ")+$F$" & c & "*$H$" & c & "*D" & r & "+$I$" & c & "*L" & r
            ws.Cells(r, cC + 2 * k + 1).Formula = "=$E$" & c & "*(C" & r & "+G" & r & ")+$F$" & c & "*$H$" & c & "*E" & r & "+$I$" & c & "*M" & r & "+" & Choose(k + 1, "I", "H", "I", "J") & r
        Next k
        For k = 0 To 3
            charge = "MAX(" & Ad(ws, r, cC + 2 * k) & "," & Ad(ws, r, cC + 2 * k + 1) & ")"
            fx = """Insuffisant"""
            For kd = NB_DIAMETRES - 1 To 0 Step -1
                cap = ws.Range(ws.Cells(LIGNE_PROFIL, ws.Range(colCap(k) & "1").Column + kd * BLOC_PROFIL), ws.Cells(LIGNE_PROFIL + NbPas(ws) - 1, ws.Range(colCap(k) & "1").Column + kd * BLOC_PROFIL)).Address
                fx = "IF(AND(ISNUMBER(" & DONNEES & "$G$" & LIGNE_PIEU + 9 + kd & "),MAX(" & cap & ")>=" & charge & ")," & DONNEES & "$G$" & LIGNE_PIEU + 9 + kd & "," & fx & ")"
            Next kd
            ws.Cells(r, cD + k).Formula = "=IF(COUNT(" & phiT & ")=0,""""," & fx & ")"
        Next k
        ws.Cells(r, cD + 4).Formula = "=IF(COUNTIF(" & Ad(ws, r, cD) & ":" & Ad(ws, r, cD + 3) & ",""Insuffisant"")>0,""Insuffisant"",IF(COUNT(" & Ad(ws, r, cD) & ":" & Ad(ws, r, cD + 3) & ")=0,"""",MAX(" & Ad(ws, r, cD) & ":" & Ad(ws, r, cD + 3) & ")))"
        If IsArray(mPhiManuel) Then
            If i <= UBound(mPhiManuel) Then ws.Cells(r, cD + 5).Value = mPhiManuel(i)
        End If
        ws.Cells(r, cD + 6).Formula = "=IF(ISNUMBER(" & Ad(ws, r, cD + 4) & "),MAX(" & Ad(ws, r, cD + 4) & ",N(" & Ad(ws, r, cD + 5) & ")),IF(N(" & Ad(ws, r, cD + 5) & ")>0," & Ad(ws, r, cD + 5) & ",IF(" & Ad(ws, r, cD + 4) & "="""","""",""Insuffisant"")))"
        kb = "(MATCH(" & Ad(ws, r, cD + 6) & "," & phiT & ",0)-1)*" & BLOC_PROFIL
        For k = 0 To 7
            cap = "$" & colCap(k \ 2) & "$" & LIGNE_PROFIL & ":$" & colCap(k \ 2) & "$" & LIGNE_PROFIL + NbPas(ws) - 1
            capk = "OFFSET(" & cap & ",0," & kb & ")"
            zk = "OFFSET(" & z & ",0," & kb & ")"
            ws.Cells(r, cP + k).Formula = "=IF(NOT(ISNUMBER(" & Ad(ws, r, cD + 6) & ")),""Non atteinte""," & "IFERROR(IF(ISNUMBER(MATCH(" & Ad(ws, r, cD + 6) & "," & phiT & ",0)),INDEX(" & zk & ",MATCH(1,INDEX(ISNUMBER(" & capk & ")*(" & capk & ">=" & Ad(ws, r, cC + k) & "),0),0))," & "INDEX(" & z & ",MATCH(1,INDEX(ISNUMBER(" & cap & ")*(" & cap & ">=" & Ad(ws, r, cC + k) & "*" & DONNEES & "$G$" & LIGNE_PIEU + 9 & "/" & Ad(ws, r, cD + 6) & "),0),0))),""Non atteinte""))"
        Next k
        ws.Cells(r, cR).Formula = "=IF(COUNTIF(" & Ad(ws, r, cP) & ":" & Ad(ws, r, cP + 7) & ",""Non atteinte"")>0,""Non atteinte"",MAX(" & Ad(ws, r, cP) & ":" & Ad(ws, r, cP + 7) & "))"
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
        ws.Cells(r, cRe).Formula = "=IF(ISNUMBER(" & Ad(ws, r, cR) & "),IF(" & Ad(ws, r, cI) & "=""Oui""," & Ad(ws, r, cR) & "/COS(RADIANS(N(" & Ad(ws, r, cA) & ")))," & Ad(ws, r, cR) & "),"""")"
        ws.Cells(r, cZ).Formula = "=IF(ISNUMBER(" & DONNEES & "$D$32)," & DONNEES & "$D$32,"""")"
        ws.Cells(r, cZ + 1).Formula = "=IF(AND(ISNUMBER(" & Ad(ws, r, cZ) & "),ISNUMBER(" & Ad(ws, r, cR) & "))," & Ad(ws, r, cZ) & "-" & Ad(ws, r, cR) & ","""")"
        ws.Cells(r, cL).Formula = "=IF(ISNUMBER(" & reco & ")," & reco & ","""")"
        ws.Cells(r, cL + 1).Formula = "=IF(AND(ISNUMBER(" & Ad(ws, r, cL) & "),ISNUMBER(" & Ad(ws, r, cRe) & "))," & Ad(ws, r, cL) & "-" & Ad(ws, r, cRe) & ","""")"
        ws.Cells(r, cL + 2).Formula = "=IF(ISNUMBER(" & Ad(ws, r, cD + 6) & "),5*" & Ad(ws, r, cD + 6) & ","""")"
        ws.Cells(r, cF).Formula = "=IF(OR(" & Ad(ws, r, cL + 1) & "=""""," & Ad(ws, r, cL + 2) & "=""""),""-"",IF(" & Ad(ws, r, cL + 1) & ">=" & Ad(ws, r, cL + 2) & ",""Vérifiée"",""Non vérifiée""))"
    Next i
    mPhiManuel = Empty
    mGeo = Empty: mGsn = Empty: mAmax = Empty: mAmin = Empty: mIncl = Empty: mAngle = Empty
    rt = r0 + 2 + Application.Max(m, 1)
    If m > 0 Then
        ws.Range("B19").Value = "Diamètre final retenu du micropieu B (plus grand " & ChrW(934) & " retenu, sinon B0)"
        ws.Range("D19").Formula = "=IF(COUNT(" & ws.Range(ws.Cells(r0 + 3, cD + 6), ws.Cells(rt, cD + 6)).Address & ")>0,MAX(" & ws.Range(ws.Cells(r0 + 3, cD + 6), ws.Cells(rt, cD + 6)).Address & ")," & DONNEES & "D" & LIGNE_PIEU + 15 & ")"
    End If
    With ws.Range(ws.Cells(r0, 1), ws.Cells(rt, cF))
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + 2, cF))
        .Font.Bold = True
        .Font.Size = 11
        .Interior.Color = COUL_ENTETE
        .WrapText = True
    End With
    If m > 0 Then
        ws.Range("A" & r0 + 3 & ":A" & rt).Font.Bold = True
        ws.Range(ws.Cells(r0 + 3, 2), ws.Cells(rt, cF)).Interior.Color = COUL_CALCUL
        ws.Range(ws.Cells(r0 + 3, cD + 5), ws.Cells(rt, cD + 5)).Interior.Color = COUL_SAISIE
        ws.Range("F" & r0 + 3 & ":F" & rt & ",K" & r0 + 3 & ":M" & rt).Interior.Color = COUL_SAISIE
        ws.Range(ws.Cells(r0 + 3, 2), ws.Cells(rt, cD - 1)).NumberFormat = "0.0"
        ws.Range(ws.Cells(r0 + 3, cD), ws.Cells(rt, cP - 1)).NumberFormat = "0.000"
        ws.Range(ws.Cells(r0 + 3, cP), ws.Cells(rt, cL + 2)).NumberFormat = "0.00"
        ws.Range(ws.Cells(r0 + 3, cD + 6), ws.Cells(rt, cD + 6)).Font.Bold = True
        ws.Range(ws.Cells(r0 + 3, cR), ws.Cells(rt, cL - 1)).Font.Bold = True
        ws.Range(ws.Cells(r0 + 3, cI), ws.Cells(rt, cI)).Interior.Color = COUL_SAISIE
        ws.Range(ws.Cells(r0 + 3, cA), ws.Cells(rt, cA)).Interior.Color = COUL_SAISIE
        ws.Range(ws.Cells(r0 + 3, cA), ws.Cells(rt, cA)).NumberFormat = "0.0"
        ws.Range(ws.Cells(r0 + 3, cI), ws.Cells(rt, cI)).Font.Bold = False
        ws.Range(ws.Cells(r0 + 3, cF), ws.Cells(rt, cF)).Font.Bold = True
        With ws.Range(ws.Cells(r0 + 3, cF), ws.Cells(rt, cF)).FormatConditions
            .Delete
            .Add(xlCellValue, xlEqual, "=""Non vérifiée""").Font.Color = COUL_ALERTE
        End With
    End If
    ws.Range(ws.Columns(cD), ws.Columns(cF)).ColumnWidth = 12
    ws.Rows(r0).RowHeight = 30
    AfficherInclinaison ws
End Sub
Public Sub AfficherInclinaison(Optional ws As Worksheet)
    Dim r0 As Long, c As Long, cI As Long, r As Long, n As Long
    If ws Is Nothing Then
        On Error Resume Next
        RenommerFeuilleCalcul
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
    ws.Range(ws.Columns(cI + 2), ws.Columns(cI + 3)).Hidden = (n = 0)
End Sub
Public Sub AppliquerInclinaisonDonnees()
    Dim ws As Worksheet, wp As Worksheet, cI As Long, r0 As Long, r As Long
    On Error Resume Next
    RenommerFeuilleCalcul
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
Private Function N0(v As Variant) As Double
    If Not IsEmpty(v) And IsNumeric(v) Then N0 = CDbl(v)
End Function
Private Function ColonneInclinaison(ws As Worksheet, r0 As Long) As Long
    Dim c As Long, v As Variant
    v = Application.Match("N° MICROPIEU", ws.Range("A1:A5000"), 0)
    If IsError(v) Then Exit Function
    r0 = CLng(v)
    For c = 2 To 80
        If CStr(ws.Cells(r0 + 1, c).Value) = "Incliné" Then ColonneInclinaison = c: Exit Function
    Next c
End Function
Private Function NbPas(ws As Worksheet) As Long
    Dim v As Variant
    v = ws.Range("ZA" & LIGNE_PROFIL - 3).Value
    If Not IsEmpty(v) And IsNumeric(v) Then NbPas = CLng(v)
    If NbPas < 1 Then NbPas = MAX_COUCHES * NB_SUBDIV
End Function
Private Function LigneVerif(ws As Worksheet) As Long
    LigneVerif = LIGNE_PROFIL + Application.Max(NbPas(ws), NB_MIN_LIGNES) + 2
End Function
Public Sub AfficherCouchesCalcul(Optional ws As Worksheet)
    Dim nb As Variant
    If ws Is Nothing Then
        On Error Resume Next
        RenommerFeuilleCalcul
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
Public Sub ActualiserPas()
    Dim ws As Worksheet, n As Long
    On Error Resume Next
    RenommerFeuilleCalcul
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
Public Sub ChangementCalcul(ws As Worksheet, Target As Range)
    Dim c As Range, zone As Range, cI As Long, r0 As Long
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
    For Each c In zone.Cells
        If Not IsEmpty(c.Value) And VarType(c.Value) = vbString Then
            If ValeurPas(c.Value) > 0 Then c.Value = ValeurPas(c.Value)
        End If
    Next c
    On Error GoTo 0
    ActualiserPas
    Application.EnableEvents = True
End Sub
Private Function ValeurPas(v As Variant) As Double
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbInteger Or VarType(v) = vbLong Or VarType(v) = vbSingle Or VarType(v) = vbCurrency Then
        ValeurPas = CDbl(v)
    Else
        ValeurPas = Val(Replace(Trim$(CStr(v)), ",", "."))
    End If
    If ValeurPas < 0 Then ValeurPas = 0
End Function
Private Sub InstallerCodeFeuille(ws As Worksheet)
    Dim cm As Object, texte As String
    On Error GoTo Fin
    texte = ws.Parent.VBProject.Name
    texte = ""
    Set cm = ws.Parent.VBProject.VBComponents(ws.CodeName).CodeModule
    If cm.CountOfLines > 0 Then texte = cm.Lines(1, cm.CountOfLines)
    If InStr(1, texte, "ChangementCalcul", vbTextCompare) = 0 Then
        cm.AddFromString "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & "    ChangementCalcul Me, Target" & vbCrLf & "End Sub"
    End If
    ws.Range("A" & LIGNE_FROT - 1).ClearComments
    Exit Sub
Fin:
    On Error Resume Next
    Commentaire ws.Range("A" & LIGNE_FROT - 1), "Pour que le tableau de la capacité portante et la " & "vérification de la profondeur s'ajustent dès la saisie d'un pas : coller CodeFeuilleCalcul.txt " & "dans le code de cette feuille, ou autoriser l'accès au projet VBA puis relancer " & "CreerCalculVerification."
End Sub
Private Function Ad(ws As Worksheet, r As Long, c As Long) As String
    Ad = ws.Cells(r, c).Address(False, False)
End Function
Private Sub GraphiqueCapacite(ws As Worksheet, k As Long)
    Dim co As ChartObject, f As String, nb As String, j As Long, cols As Variant, noms As Variant
    Dim d As Long, col As String, lg As Double
    f = "'" & ws.Name & "'!"
    d = k * BLOC_PROFIL
    nb = "MAX(1,COUNT(" & f & ws.Cells(LIGNE_PROFIL, 3 + d).Address & ":" & ws.Cells(LIGNE_PROFIL + NbPas(ws) - 1, 3 + d).Address & "))"
    cols = Array(3, 18, 20, 22)
    noms = Array("Profondeur z", "ELS caractéristique - Rc;cr,d", "ELS quasi permanent - Rc;cr,d", "ELU fondamental - Rc,d")
    For j = 0 To 3
        ws.Names.Add Name:="CP" & k & "_" & IIf(j = 0, "z", CStr(j)), RefersTo:="=OFFSET(" & f & ws.Cells(LIGNE_PROFIL, cols(j) + d).Address & ",0,0," & nb & ")"
    Next j
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
                .ChartType = xlXYScatterSmoothNoMarkers
                .Smooth = True
                .MarkerStyle = xlMarkerStyleNone
                .Format.Line.Weight = 3
                .Format.Line.ForeColor.RGB = Choose(((j - 1) Mod 4) + 1, 5898259, 11371036, 8964355, 8139520)
            End With
        Next j
        .HasTitle = True
        .ChartTitle.Text = "Capacité portante en compression en fonction de la profondeur - " & "diamètre B" & k & " = " & Format(ws.Parent.Worksheets("Données du projet").Range("G" & LIGNE_PIEU + 9 + k).Value, "0.000") & " m"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        With .Axes(xlValue)
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
Private Function Derniere(col As String) As String
    Derniere = "=IFERROR(LOOKUP(2,1/($C$39:$C$53<>""""),$" & col & "$39:$" & col & "$53),"""")"
End Function
Private Sub LigneS(ws As Worksheet, r As Long, intitule As String, symbole As String, formule As String, unite As String)
    ws.Range("B" & r).Value = intitule
    ws.Range("C" & r).Value = symbole
    ws.Range("D" & r).Formula = formule
    ws.Range("E" & r).Value = unite
End Sub
Private Function Array2(a As String, b As String, c As String, d As String) As Variant
    Dim t(1 To 2, 1 To 2) As String
    t(1, 1) = a: t(1, 2) = b: t(2, 1) = c: t(2, 2) = d
    Array2 = t
End Function
Private Sub Ligne(ws As Worksheet, r As Long, intitule As String, formule As String, unite As String)
    ws.Range("B" & r).Value = intitule
    ws.Range("B" & r & ":C" & r).Merge
    ws.Range("D" & r).Formula = formule
    ws.Range("E" & r).Value = unite
End Sub
Private Sub SousTitre(ws As Worksheet, r As Long, texte As String, Optional col As String = "B")
    With ws.Range(col & r)
        .Value = texte
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Size = 14
        .Font.Color = COUL_TITRE
    End With
End Sub
Private Function Rang6(sol As String) As String
    Rang6 = "IFERROR(MATCH(" & sol & ",{""" & Replace(CATEGORIES_SOL, ",", """,""") & """},0),1)"
End Function
Private Function Rang5(sol As String) As String
    Rang5 = "CHOOSE(" & Rang6(sol) & ",1,1,2,3,4,5)"
End Function
Private Function IndexNorme(zones As String, lig As String, sol As String, Optional f6 As Boolean = False) As String
    Dim z() As String
    z = Split(zones, "|")
    If mPenetro Then
        IndexNorme = "INDEX(" & z(0) & "," & lig & "," & Rang6(sol) & ")"
    Else
        IndexNorme = "INDEX(" & z(1) & "," & lig & "," & IIf(f6, Rang6(sol), Rang5(sol)) & ")"
    End If
End Function
Private Sub Commentaire(c As Range, texte As String)
    On Error Resume Next
    Application.Run "CommentaireImage.CommentaireImage", c, texte
    If Err.Number <> 0 Then
        Err.Clear
        c.ClearComments
        c.AddComment texte
        c.Comment.Shape.TextFrame.AutoSize = True
    End If
    On Error GoTo 0
End Sub
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
Private Function FormuleIndex(zone As String, lig As String, col As String) As String
    FormuleIndex = "=IFERROR(INDEX(" & zone & "," & lig & "," & col & "),"""")"
End Function
Private Function FormuleProduit(alpha As String, fsol As String) As String
    FormuleProduit = "=IF(OR(" & fsol & "="""",ISBLANK(" & alpha & ")),""""," & "IF(ISNUMBER(" & alpha & ")," & alpha & "*" & fsol & ",""-""))"
End Function
Private Function FormuleQs(produit As String, qsmax As String) As String
    FormuleQs = "=IF(AND(ISNUMBER(" & produit & "),ISNUMBER(" & qsmax & ")),MIN(" & produit & "," & qsmax & "),IF(" & produit & "="""","""",""-""))"
End Function
Private Sub CoucheLigne(ws As Worksheet, r As Long, i As Long, ld As Long, cat As String, gamma As String, zAlpha As String, zABC As String, zQs As String)
    Dim cond As String, sol As String, pl As String, c As Variant
    cond = "=IF(OR(" & DONNEES & "$C$" & ld & "=""""," & DONNEES & "$D$" & ld & "=""""),"""","
    sol = "$D$" & r
    pl = "$G$" & r
    ws.Range("B" & r).Formula = cond & i & ")"
    ws.Range("C" & r).Formula = cond & DONNEES & "$C$" & ld & ")"
    ws.Range("D" & r).Formula = cond & DONNEES & "$D$" & ld & ")"
    ws.Range("E" & r).Formula = cond & "IF(ISNUMBER(" & DONNEES & "$E$" & ld & ")," & DONNEES & "$E$" & ld & ",""""))"
    ws.Range("F" & r).Formula = cond & "IF(ISNUMBER(" & DONNEES & "$F$" & ld & ")," & DONNEES & "$F$" & ld & ",""""))"
    c = IIf(mPenetro, "$I$", "$G$")
    ws.Range("G" & r).Formula = cond & "IF(ISNUMBER(" & DONNEES & c & ld & ")," & DONNEES & c & ld & ",""""))"
    mCellule = "H" & r
    ws.Range("H" & r).Formula = cond & "IFERROR(" & IndexNorme(zABC, "1", sol) & ",""""))"
    mCellule = "I" & r
    ws.Range("I" & r).Formula = cond & "IFERROR(" & IndexNorme(zAlpha, cat, sol) & ",""""))"
    mCellule = "J" & r
    ws.Range("J" & r).Formula = "=IF(OR(B" & r & "="""",G" & r & "="""",H" & r & "=""""),"""",IFERROR(" & "1000*(" & IndexNorme(zABC, "2", sol) & "*" & pl & "+" & IndexNorme(zABC, "3", sol) & ")*(1-EXP(-" & IndexNorme(zABC, "4", sol) & "*" & pl & ")),""""))"
    ws.Range("K" & r).Formula = FormuleProduit("$I$" & r, "$J$" & r)
    mCellule = "L" & r
    ws.Range("L" & r).Formula = cond & "IFERROR(" & IndexNorme(zQs, cat, sol, True) & ",""""))"
    ws.Range("M" & r).Formula = FormuleQs("$K$" & r, "$L$" & r)
    ws.Range("N" & r).Formula = "=IF(OR(B" & r & "="""",F" & r & "="""",N($D$19)=0),""""," & "PI()*$D$19*F" & r & ")"
    ws.Range("O" & r).Formula = "=IF(B" & r & "="""","""",IF(Q" & r & "=""Oui"",0," & "IFERROR(M" & r & "/(" & gamma & "),""-"")))"
    ws.Range("P" & r).Formula = "=IF(AND(ISNUMBER(O" & r & "),ISNUMBER(N" & r & "))," & "IF(Q" & r & "=""Oui"",0,O" & r & "*N" & r & "),IF(B" & r & "="""","""",""-""))"
    ws.Range("Q" & r).Formula = cond & "IF(" & DONNEES & "$D$" & ld & "=""Neutralisé"",""Oui""," & DONNEES & "$M$" & ld & "))"
    For Each c In Array("G", "H", "I", "J", "K", "L", "M", "N", "O", "P")
        ws.Range(c & r).Formula = "=IF($Q$" & r & "=""Oui"",0," & Mid(ws.Range(c & r).Formula, 2) & ")"
    Next c
    With ws.Range("G" & r & ":P" & r).FormatConditions
        .Delete
        .Add(xlExpression, Formula1:="=$Q$" & r & "=""Oui""").NumberFormat = """-"""
    End With
    ws.Range("E" & r & ":G" & r & ",I" & r).NumberFormat = "0.00"
    ws.Range("J" & r & ":M" & r & ",O" & r & ":P" & r).NumberFormat = "0.0"
    ws.Range("N" & r).NumberFormat = "0.000"
End Sub
Private Function ZoneTableau(wt As Worksheet, titre As String, decalage As Long, nbLignes As Long, nbCols As Long) As String
    Dim c As Range, trouve As Range, ent As Range, premier As String, nom As String
    Dim cle As String, lig As Long, col As Long
    Set c = wt.Cells.Find(What:=titre & " - ", After:=wt.Cells(wt.Rows.Count, wt.Columns.Count), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True)
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
    If trouve Is Nothing Then Err.Raise vbObjectError + 1, , "« " & titre & " » introuvable " & "dans la feuille « " & wt.Name & " » : relancez CreerTableauxNFP94262."
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
    If ent Is Nothing Then Err.Raise vbObjectError + 2, , "En-tête des valeurs du « " & titre & " » introuvable dans la feuille « " & wt.Name & " » : relancez CreerTableauxNFP94262."
    nom = "NF_" & Replace(Replace(Trim$(Mid$(titre, Len("Tableau ") + 1)), ".", "_"), " ", "")
    ActiveWorkbook.Names.Add Name:=nom, RefersTo:="='" & wt.Name & "'!" & ent.Offset(1 + decalage, 0).Resize(nbLignes, nbCols).Address
    ZoneTableau = nom
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
