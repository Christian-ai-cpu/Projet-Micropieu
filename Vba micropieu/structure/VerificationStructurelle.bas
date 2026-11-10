Attribute VB_Name = "VerificationStructurelle"
Option Explicit

' Module « Vérification structurelle » : crée (ou recrée) la feuille « Vérification
' structurelle » des micropieux, à partir des feuilles « Données du projet » (module
' DonneesProjet), « Calcul et vérification » (module CalculFondation) et de la feuille de
' descente de charges (matrice des micropieux, module MatriceMicropieux). Toutes les
' valeurs sont des formules : elles suivent les saisies de ces feuilles.
'   A3  : longueur de transfert l0 de chaque micropieu : diamètre du pieu foré (Phi
'         retenu ou autre valeur saisie), section du pieu foré, enrobage et protection
'         par le coulis (c >= 5 cm, C > 500 kg/m3, E/C < 0,5 : pas de corrosion), rayon
'         de corrosion (« Données du projet »), section de l'armature (tube N80, barre HA
'         ou GEWI) brute, corrodée et après corrosion, inertie, EI, EM et alpha (ou qc et
'         beta) et Kf selon la méthode de calcul (pressiométrique : Kf = EM x 18 rho /
'         [4 (2,65 rho)^alpha + 3 alpha rho], rho = max(B/B0 ; 1) ; CPT : Kf = beta qc
'         / 2), Es selon le type de sollicitation (a, b : Kf ; c, d : 2 Kf),
'         l0 = (4 EI / Es)^(1/4), D réelle, comportement rigide (l0 >= D) ou souple
'         (D >= 3 l0)
'   dessous : déplacement en tête (voir plus bas), puis efforts dans chaque micropieu et
'         chaque combinaison (ELS caractéristique,
'         ELS quasi permanent, ELU fondamental, ELU accidentel) : compression Nc = Vmax
'         de la vérification de la profondeur, traction Nt = - Nc, effort tranchant en
'         tête V = gamma G,sup x H x L (L : longueur de mur reprise par le micropieu,
'         matrice de la descente de charges), moment en tête M = V x e ; poussée des
'         terres au repos sur la fondation existante : h = hauteur du poteau du
'         soubassement + hauteur de la semelle, K0 = 1 - sin phi' (couche 1),
'         H = K0 gamma h² / 2 (kN/ml), e = h / 3 (rappel au-dessus du tableau)
'   déplacement en tête y0 de chaque micropieu et chaque combinaison, entre le tableau
'         de l0 et celui des efforts (poutre de longueur D sur appuis élastiques, solution
'         de Hetényi : tête libre y0 = a H + b M, tête encastrée y0 = (a - b²/c) H et
'         moment d'encastrement) ; H = V + T0, T0 = Nc x sin(angle) si le micropieu est
'         incliné (effort horizontal provenant de la descente de charges)
'   Y, T et M le long de chaque micropieu, de la tête à la profondeur réelle D (pieu long
'         à pointe encastrée : ondes issues de la tête et de la pointe, constantes A, B,
'         C, E par combinaison), profondeurs = pas de calcul de « Calcul et
'         vérification » ; graphiques Y, T, M à la fin de la feuille
' Le nombre de micropieux et les adresses des autres feuilles sont lus à la création :
' relancer CreerVerificationStructurelle après un changement de ces feuilles.

Private Const NOM_FEUILLE As String = "Vérification structurelle"
Private Const DONNEES As String = "'Données du projet'!"
Private Const TITRE_DEPL As String = "DÉPLACEMENT EN TÊTE DES MICROPIEUX"
' Profils Y, T, M le long du micropieu : LARG_PROFIL colonnes par micropieu (15 + 1
' vide) ; graphiques à la fin de la feuille, HAUT_GRAPHE lignes par micropieu
Private Const LARG_PROFIL As Long = 16
Private Const HAUT_GRAPHE As Long = 26
Private Const CALCUL As String = "'Calcul et vérification'!"

' Thème : saisies en vert, valeurs calculées en jaune, en-têtes en vert clair
Private Const VERT_SAISIE As Long = 13561798   ' RGB(198, 239, 206)
Private Const JAUNE_CALCUL As Long = 13434879  ' RGB(255, 255, 204)
Private Const VERT_ENTETE As Long = 14348258   ' RGB(226, 239, 218)
Private Const VERT_FONCE As Long = 2315831     ' RGB(55, 86, 35)

Public Sub CreerVerificationStructurelle()
    Dim ws As Worksheet, wc As Worksheet, wd As Worksheet, f As Worksheet, etape As String
    Dim a As Long, lp As Long, lm As Long, r As Long, i As Long, k As Long, m As Long, n As Long
    Dim r0 As Long, rv As Long, cC As Long, tC As Long, rd As Long, rP As Long, rS As Long
    Dim tube As String, dn As String, entete As Variant
    Dim rL As Long, cPhi As Long, cDr As Long, autre As Variant, x As String, j As Long
    Dim dA As String, dI As String, eA As String, pen As String, sol As String, b0 As String
    Dim emq As String, alp As String, nat As String, bet As String
    Dim hT As String, k0 As String, hH As String, eE As String
    Dim rY As Long, condTete As Variant, ry0 As String, rl0 As String, xs As String
    Dim cIn As Long, cAn As Long, rT As Long, nZ As Long, hZ As Variant, cb As Long, rd0 As String
    Dim kA As String, kB As String, kC As String, kD As String, y0s As String, hs As String
    Dim ths As String, mhs As String, cond0 As String, rP As Long, rG As Long, pr As String, zk As String

    Application.ScreenUpdating = False
    On Error GoTo Erreur
    etape = "feuilles du projet"
    Set wc = ActiveWorkbook.Worksheets("Calcul et vérification")
    For Each f In ActiveWorkbook.Worksheets
        If CStr(f.Range("B80").Value) = "RÉPARTITION DES CHARGES DES MURS SUR LES MICROPIEUX" Then
            Set wd = f
            Exit For
        End If
    Next f

    etape = "création de la feuille"
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo Erreur
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = NOM_FEUILLE
    End If
    autre = LireAutreDiametre(ws)             ' autres diamètres saisis (tableau de l0)
    condTete = LireColonne(ws, TITRE_DEPL, "B") ' conditions en tête saisies (déplacements)
    ws.Cells.Clear
    Do While ws.ChartObjects.Count > 0       ' graphiques de la création précédente
        ws.ChartObjects(1).Delete
    Loop
    a = LIGNE_ACIER                           ' caractéristiques de l'acier (Données)
    lp = LIGNE_PIEU                           ' caractéristiques de la fondation (Données)
    lm = LIGNE_COUCHE_MOLLE                   ' frottement négligé et couche molle (Données)

    With ws.Range("B1:I1")
        .Merge
        .Value = "VÉRIFICATION STRUCTURELLE DES MICROPIEUX"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
        .HorizontalAlignment = xlCenter
    End With

    ' --- Paramètres lus dans les autres feuilles (formules directes) ---
    ' Armature (« Données du projet », caractéristiques de l'acier) : type, d, di, E
    tube = DONNEES & "$D$" & a & "=""Tube N80"""
    dA = "N(" & DONNEES & "$D$" & a + 2 & ")"
    dI = "N(" & DONNEES & "$D$" & a + 4 & ")"
    eA = "N(" & DONNEES & "$D$" & a + 6 & ")"
    ' Méthode de calcul, type de sollicitation, B0, couche molle : EM ou qc, alpha, beta
    ' (CPT : sols argileux 12, limons 7,5, sols sableux, craies, marnes, rochers 4,5)
    pen = "(" & DONNEES & CELLULE_METHODE & "=""Méthode pénétrométrique"")"
    sol = DONNEES & "$D$" & lp + 21
    b0 = "N(" & DONNEES & "$D$" & lp + 22 & ")"
    emq = "N(" & DONNEES & "$D$" & lm + 1 & ")"
    alp = "N(" & DONNEES & "$D$" & lm + 6 & ")"
    nat = DONNEES & "$D$" & lm + 5
    bet = "IF(OR(" & nat & "=""Tourbe""," & nat & "=""Argile""),12,IF(" & nat & "=""Limon"",7.5,4.5))"
    ' Poussée des terres au repos sur la fondation existante : h = hauteur du poteau du
    ' soubassement + hauteur de la semelle (descente de charges, colonne F), gamma et phi'
    ' de la couche 1, K0 = 1 - sin phi', H = K0 gamma h² / 2 (kN/ml), e = h / 3
    If Not wd Is Nothing Then
        ' Lignes du soubassement dans la descente de charges (mur 1) : poteau et semelle
        For r = 5 To 79
            If CStr(wd.Range("B" & r).Value) = "Soubassement" Then
                For i = r + 1 To r + 6
                    If Left(CStr(wd.Range("B" & i).Value), 6) = "Poteau" And rP = 0 Then rP = i
                    If Left(CStr(wd.Range("B" & i).Value), 7) = "Semelle" And rS = 0 Then rS = i
                Next i
                Exit For
            End If
        Next r
        dn = "'" & wd.Name & "'!"
    End If
    hT = "(" & IIf(rP > 0, "N(" & dn & "$F$" & rP & ")", "0") & "+" & _
        IIf(rS > 0, "N(" & dn & "$F$" & rS & ")", "0") & ")"
    k0 = "(1-SIN(RADIANS(N(" & DONNEES & "$L$" & LIGNE_COUCHES & "))))"
    hH = "(" & k0 & "*N(" & DONNEES & "$J$" & LIGNE_COUCHES & ")*" & hT & "^2/2)"
    eE = "(" & hT & "/3)"

    etape = "micropieux de la vérification de la profondeur"
    ' Vérification de la profondeur (« Calcul et vérification ») : en-tête « N° MICROPIEU »
    ' en colonne A, combinaisons (Vmax) dans la colonne « COMBINAISONS ... », coefficients
    ' des combinaisons sous le tableau
    ' (position variable : sous le tableau de la capacité portante par pas de calcul)
    On Error Resume Next
    r0 = Application.Match("N° MICROPIEU", wc.Range("A1:A5000"), 0)
    On Error GoTo Erreur
    If r0 > 0 Then
        For i = 1 To 60
            If CStr(wc.Cells(r0, i).Value) = "COMBINAISONS DE CHARGES VERTICALES (kN)" Then cC = i
        Next i
        Do While Left(CStr(wc.Range("A" & r0 + 3 + m).Value), 2) = "MP"
            m = m + 1
        Loop
        For r = r0 + 3 + m To r0 + 3 + m + 10
            If Left(CStr(wc.Range("A" & r).Value), 14) = "COEFFICIENTS D" Then tC = r: Exit For
        Next r
    End If
    If Not wd Is Nothing Then
        Do While Left(CStr(wd.Range("B" & 86 + n).Value), 4) = "Mur "
            n = n + 1
        Loop
        rd = 88 + n + 3                       ' MP1 dans la matrice des micropieux
    End If
    If m = 0 Or cC = 0 Or wd Is Nothing Then
        ws.Range("A3").Value = "Créez la matrice des micropieux (descente de charges) " & _
            "et « Calcul et vérification », puis relancez CreerVerificationStructurelle."
        GoTo MiseEnForme
    End If
    rL = 3                                    ' titre du tableau de la longueur de transfert

    ' --- Longueur de transfert l0 de chaque micropieu ---
    ' Diamètre du pieu foré B = Phi retenu (vérification de la profondeur) ou autre valeur
    ' saisie ; section du pieu foré A = pi B² / 4 ; enrobage de l'armature par le coulis
    ' c = (B - d) / 2 ; protection par le coulis (acier sollicité en compression) : efficace
    ' si enrobage >= 5 cm, dosage > 500 kg/m3 et E/C < 0,5 (« Données du projet ») ; rayon
    ' de corrosion rc (« Données du projet » : durée du projet x nature du terrain), 0 si
    ' la protection par le coulis est efficace ; armature : diamètre après corrosion
    ' d - 2 rc (tube : face extérieure, barre : tout le pourtour), section brute,
    ' corrodée et après corrosion ; inertie de la section après corrosion et module de
    ' flexion EI = E x I ; module de réaction Kf selon la méthode de calcul :
    '   pressiométrique : EM et coefficient rhéologique alpha, rho = max(B / B0 ; 1),
    '                     Kf = EM x 18 rho / [4 (2,65 rho)^alpha + 3 alpha rho] ;
    '   pénétrométrique : qc et beta, Kf = beta qc / 2 ;
    ' Es selon le type de sollicitation ; l0 = (4 EI / Es)^(1/4) ; D = profondeur réelle.
    etape = "longueur de transfert"
    For i = 1 To 80
        If CStr(wc.Cells(r0 + 1, i).Value) = ChrW(934) & " RETENU" Then cPhi = i
        If CStr(wc.Cells(r0, i).Value) = "PROFONDEUR RÉELLE (m)" Then cDr = i
        If cDr = 0 And CStr(wc.Cells(r0, i).Value) = "PROFONDEUR RETENUE (m)" Then cDr = i
    Next i
    rL = rv + 2 + m + 2
    With ws.Range("A" & rL)
        .Value = "LONGUEUR DE TRANSFERT l0 DES MICROPIEUX (APRÈS CORROSION)"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
    entete = Array("N° MICROPIEU", ChrW(934) & " retenu (m)", "Autre diamètre (m)", _
        "Diamètre du pieu foré B (m)", "Section du pieu foré A = " & ChrW(960) & ".B²/4 (m²)", _
        "Enrobage c = (B - d) / 2 (mm)", _
        "Protection par le coulis (c >= 50 mm, C > 500 kg/m3, E/C < 0,5, compression)", _
        "Rayon de corrosion (Données) (mm)", "Rayon de corrosion retenu rc (mm)", _
        "Diamètre de l'armature d (mm)", "Diamètre après corrosion d - 2 rc (mm)", _
        "Section brute de l'armature (mm²)", "Section corrodée (mm²)", "Section après corrosion As (mm²)", _
        "Inertie de la section brute (mm4)", "Inertie de la section après corrosion I (mm4)", _
        "Module de flexion EI = E x I (kN.m²)", _
        "=IF(" & pen & ",""Résistance de pointe qc (MPa)""," & _
            """Module pressiométrique EM (MPa)"")", _
        "=IF(" & pen & ",""" & ChrW(946) & " (CPT)""," & _
            """Coefficient rhéologique " & ChrW(945) & """)", _
        ChrW(961) & " = max(B/B0 ; 1)", "=""Module de réaction Kf (MPa) - ""&LOWER(" & DONNEES & CELLULE_METHODE & ")", _
        "Es (MPa)", "l0 = (4 EI / Es)^(1/4) (m)", "Profondeur D (m)", "Comportement")
    For j = 0 To UBound(entete)
        If Left(entete(j), 1) = "=" Then
            ws.Cells(rL + 2, 1 + j).Formula = entete(j)
        Else
            ws.Cells(rL + 2, 1 + j).Value = entete(j)
        End If
    Next j
    Commentaire ws.Range("C" & rL + 2), "Autre diamètre : à saisir pour remplacer " & ChrW(934) & _
        " retenu (vide : " & ChrW(934) & " retenu de la vérification de la profondeur)"
    Commentaire ws.Range("F" & rL + 2), "Enrobage : épaisseur de coulis entre l'armature et la " & _
        "paroi du forage, c = (B - d) / 2"
    Commentaire ws.Range("G" & rL + 2), "Aciers sollicités en compression : un enrobage de mortier ou " & _
        "de coulis de ciment approprié constitue une protection efficace si l'enrobage minimal est de " & _
        "5 cm et si le coulis est dosé à plus de 500 kg de ciment par m3 avec un rapport E/C inférieur " & _
        "à 0,5 (dosage et E/C : « Données du projet », caractéristiques du coulis) ; la corrosion " & _
        "n'est alors pas déduite"
    Commentaire ws.Range("H" & rL + 2), "Perte d'épaisseur de « Données du projet » (caractéristiques " & _
        "du coulis), selon la durée d'utilisation du projet et la nature du terrain (NF EN 1993-5, " & _
        "tableaux 4.1 / 4.2)"
    Commentaire ws.Range("I" & rL + 2), "rc = 0 si la protection par le coulis est efficace, sinon " & _
        "rayon de corrosion de « Données du projet »"
    Commentaire ws.Range("M" & rL + 2), "Section corrodée = section brute - section après corrosion " & _
        "(tube : face extérieure seulement, barre : tout le pourtour)"
    Commentaire ws.Range("P" & rL + 2), "Inertie de la section après corrosion : tube " & ChrW(960) & _
        "/64 ((d - 2 rc)^4 - di^4), barre " & ChrW(960) & "/64 (d - 2 rc)^4"
    Commentaire ws.Range("R" & rL + 2), "EM (méthode pressiométrique) ou qc (méthode " & _
        "pénétrométrique) de la couche molle, selon la méthode de calcul de « Données du projet »"
    Commentaire ws.Range("S" & rL + 2), "Méthode pressiométrique : coefficient rhéologique " & ChrW(945) & _
        " (tableaux I.1.3.1 / I.1.3.2) ; méthode pénétrométrique : " & ChrW(946) & " (CPT)"
    Commentaire ws.Range("U" & rL + 2), "Pressiométrique : Kf = EM x 18 " & ChrW(961) & " / [4 (2,65 " & _
        ChrW(961) & ")^" & ChrW(945) & " + 3 " & ChrW(945) & " " & ChrW(961) & "] ; pénétrométrique : " & _
        "Kf = " & ChrW(946) & " qc / 2"
    Commentaire ws.Range("X" & rL + 2), "D : profondeur réelle du micropieu (D / cos " & ChrW(945) & _
        " s'il est incliné), vérification de la profondeur"
    For i = 1 To m
        r = rL + 2 + i
        x = CStr(r)
        ws.Range("A" & r).Value = "MP" & i
        If cPhi > 0 Then
            ws.Range("B" & r).Formula = "=IF(ISNUMBER(" & CALCUL & wc.Cells(r0 + 2 + i, cPhi).Address & ")," & _
                CALCUL & wc.Cells(r0 + 2 + i, cPhi).Address & ",N(" & CALCUL & "$D$19))"
        Else
            ws.Range("B" & r).Formula = "=N(" & CALCUL & "$D$19)"
        End If
        If IsArray(autre) Then
            If i <= UBound(autre) Then
                If IsNumeric(autre(i)) And Not IsEmpty(autre(i)) Then ws.Range("C" & r).Value = autre(i)
            End If
        End If
        ws.Range("D" & r).Formula = "=IF(N(C" & x & ")>0,C" & x & ",B" & x & ")"
        ws.Range("E" & r).Formula = "=PI()*D" & x & "^2/4"
        ws.Range("F" & r).Formula = "=MAX(0,(1000*D" & x & "-J" & x & ")/2)"
        ' Protection par le coulis : enrobage >= 50 mm, dosage > 500 kg/m3, E/C < 0,5,
        ' micropieu en compression
        ws.Range("G" & r).Formula = "=IF(AND(F" & x & ">=50,N(" & DONNEES & "D" & lp + 28 & ")>500," & _
            "N(" & DONNEES & "D" & lp + 27 & ")>0,N(" & DONNEES & "D" & lp + 27 & ")<0.5," & _
            DONNEES & "D" & lp + 12 & "<>""Traction""),""Efficace"",""Non efficace"")"
        ws.Range("H" & r).Formula = "=N(" & DONNEES & "D" & lp + 30 & ")"
        ws.Range("I" & r).Formula = "=IF(G" & x & "=""Efficace"",0,H" & x & ")"
        ws.Range("J" & r).Formula = "=" & dA & ""
        ws.Range("K" & r).Formula = "=MAX(0,J" & x & "-2*I" & x & ")"
        ws.Range("L" & r).Formula = "=IF(" & tube & ",PI()/4*(J" & x & "^2-" & dI & "^2),PI()/4*J" & x & "^2)"
        ws.Range("N" & r).Formula = "=MAX(0,IF(" & tube & ",PI()/4*(K" & x & "^2-" & dI & "^2),PI()/4*K" & x & "^2))"
        ws.Range("M" & r).Formula = "=L" & x & "-N" & x
        ws.Range("O" & r).Formula = "=IF(" & tube & ",PI()/64*(J" & x & "^4-" & dI & "^4),PI()/64*J" & x & "^4)"
        ws.Range("P" & r).Formula = "=MAX(0,IF(" & tube & ",PI()/64*(K" & x & "^4-" & dI & "^4),PI()/64*K" & x & "^4))"
        ws.Range("Q" & r).Formula = "=" & eA & "*P" & x & "*1E-9"
        ' EM ou qc et alpha ou beta selon la méthode de calcul (« Données du projet »)
        ws.Range("R" & r).Formula = "=" & emq & ""
        ws.Range("S" & r).Formula = "=IF(" & pen & "," & bet & "," & alp & ")"
        ws.Range("T" & r).Formula = "=IF(" & b0 & "<=0,"""",MAX(D" & x & "/" & b0 & ",1))"
        ws.Range("U" & r).Formula = "=IF(OR(N(R" & x & ")<=0,T" & x & "=""""),"""",IF(" & pen & "," & _
            "S" & x & "*R" & x & "/2,R" & x & "*18*T" & x & "/(4*(2.65*T" & x & ")^S" & x & "+3*S" & x & "*T" & x & ")))"
        ws.Range("V" & r).Formula = "=IF(U" & x & "="""","""",IF(OR(LEFT(" & sol & ",1)=""c"",LEFT(" & sol & ",1)=""d"")," & _
            "2,1)*U" & x & ")"
        ws.Range("W" & r).Formula = "=IF(OR(V" & x & "="""",N(V" & x & ")=0),"""",(4*Q" & x & "/(V" & x & _
            "*1000))^0.25)"
        If cDr > 0 Then
            ws.Range("X" & r).Formula = "=N(" & CALCUL & wc.Cells(r0 + 2 + i, cDr).Address & ")"
        Else
            ws.Range("X" & r).Formula = "=N(" & CALCUL & "$D$20)"
        End If
        ws.Range("Y" & r).Formula = "=IF(W" & x & "="""",""-"",IF(W" & x & ">=X" & x & _
            ",""Rigide (l0 >= D)"",IF(X" & x & ">=3*W" & x & ",""Souple (D >= 3 l0)"",""Intermédiaire"")))"
    Next i
    With ws.Range("A" & rL + 2 & ":Y" & rL + 2 + m)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A" & rL + 2 & ":Y" & rL + 2)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .WrapText = True
    End With
    ws.Range("A" & rL + 3 & ":A" & rL + 2 + m).Font.Bold = True
    ws.Range("B" & rL + 3 & ":Y" & rL + 2 + m).Interior.Color = JAUNE_CALCUL
    ws.Range("C" & rL + 3 & ":C" & rL + 2 + m).Interior.Color = VERT_SAISIE
    ws.Range("B" & rL + 3 & ":E" & rL + 2 + m).NumberFormat = "0.000"
    ws.Range("F" & rL + 3 & ":F" & rL + 2 + m & ",H" & rL + 3 & ":K" & rL + 2 + m).NumberFormat = "0.00"
    ws.Range("L" & rL + 3 & ":N" & rL + 2 + m).NumberFormat = "0.0"
    ws.Range("O" & rL + 3 & ":P" & rL + 2 + m).NumberFormat = "0.000E+00"
    ws.Range("Q" & rL + 3 & ":X" & rL + 2 + m).NumberFormat = "0.00"
    ws.Range("G" & rL + 3 & ":G" & rL + 2 + m & ",P" & rL + 3 & ":Q" & rL + 2 + m & ",W" & rL + 3 & ":W" & _
        rL + 2 + m & ",Y" & rL + 3 & ":Y" & rL + 2 + m).Font.Bold = True
    With ws.Range("G" & rL + 3 & ":G" & rL + 2 + m).FormatConditions
        .Delete
        .Add(xlCellValue, xlEqual, "=""Non efficace""").Font.Color = RGB(192, 0, 0)
    End With


    ' --- Déplacement en tête de chaque micropieu, pour chaque combinaison ---
    ' (entre le tableau de la longueur de transfert et celui des efforts)
    ' Poutre de longueur D sur appuis élastiques (module Es, longueur de transfert l0,
    ' lambda = 1 / l0, x = D / l0), extrémité inférieure libre. Effort transversal en
    ' tête H = V (poussée des terres, tableau des efforts) + T0 ; micropieu incliné
    ' (« Calcul et vérification », colonnes Incliné et Angle) : T0 = Nc x sin(angle),
    ' composante perpendiculaire à l'axe du micropieu de l'effort vertical de la descente
    ' de charges (Nc de la combinaison), 0 sinon ; moment M du tableau des efforts.
    ' Souplesses en tête (solution analytique de Hetényi, poutre de longueur finie ;
    ' module Es en kPa) :
    '   y = a H + b M, rotation = b H + c M, avec
    '   a = 2 lambda / Es x (sh ch - s c) / (sh² - s²)
    '   b = 2 lambda² / Es x (sh² + s²) / (sh² - s²)
    '   c = 4 lambda³ / Es x (sh ch + s c) / (sh² - s²)
    '   (sh, ch = sinh x, cosh x ; s, c = sin x, cos x). Limites : micropieu souple
    '   (x grand) y = 2 H / (Es l0) + 2 M / (Es l0²) ; rigide (x petit)
    '   y = 4 H / (Es D) + 6 M / (Es D²).
    ' Tête libre : y0 = a H + b M ; tête encastrée (rotation nulle) : y0 = (a - b² / c) H,
    ' moment d'encastrement M0 = b / c x H (souple : H l0 / 2 ; rigide : H D / 2).
    ' Comportement linéaire (partie élastique de la courbe de réaction, p <= Pf).
    etape = "déplacements en tête"
    rY = rL + 2 + m + 3
    ' Profils Y, T, M le long du micropieu (sous ce tableau) :
    ' profondeurs du tableau de la capacité portante par pas de calcul de « Calcul et
    ' vérification » (colonne C, sous l'en-tête « Profondeur z (m) »), nombre de pas
    ' écrit en ZA deux lignes au-dessus de l'en-tête
    rT = rY + 3 + m + 3                       ' titre des constantes de Y, T, M le long du micropieu
    hZ = Application.Match("Profondeur z (m)", wc.Range("C1:C200"), 0)
    nZ = 0
    If Not IsError(hZ) Then
        If IsNumeric(wc.Range("ZA" & CLng(hZ) - 2).Value) Then nZ = CLng(N0(wc.Range("ZA" & CLng(hZ) - 2).Value))
        If nZ <= 0 Then
            Do While IsNumeric(wc.Range("C" & CLng(hZ) + 1 + nZ).Value) And _
                Not IsEmpty(wc.Range("C" & CLng(hZ) + 1 + nZ).Value) And nZ < 5000
                nZ = nZ + 1
            Loop
        End If
    End If
    If nZ < 1 Then nZ = 1
    rP = rT + 3 + m + 3                       ' en-tête des profils le long du micropieu
    rv = rP + 4 + nZ + 5                      ' en-tête du tableau des efforts
    For i = 1 To 80                           ' inclinaison et angle de chaque micropieu
        If CStr(wc.Cells(r0 + 1, i).Value) = "Incliné" Then cIn = i
        If CStr(wc.Cells(r0 + 1, i).Value) = "Angle (°)" Then cAn = i
    Next i
    With ws.Range("A" & rY)
        .Value = TITRE_DEPL
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
    ws.Range("B" & rY + 1).Value = "Calcul linéaire (Es, l0 du tableau de la longueur de transfert), " & _
        "solution de la poutre de longueur D sur appuis élastiques (Hetényi) ; H = V + T0, M du " & _
        "tableau des efforts ; y en mm"
    r = rY + 2
    ws.Range("A" & r & ":H" & r).Value = Array("N° MICROPIEU", "Condition en tête", "Incliné", _
        "Angle (°)", "x = D / l0", "a (m/kN)", "b (1/kN)", "c (1/(kN.m))")
    ws.Range("I" & r).Value = "EFFORT HORIZONTAL T0 = Nc x sin " & ChrW(945) & " (kN)"
    ws.Range("M" & r).Value = "EFFORT TRANSVERSAL EN TÊTE H = V + T0 (kN)"
    ws.Range("Q" & r).Value = "DÉPLACEMENT EN TÊTE y0 (mm)"
    ws.Range("U" & r).Value = "MOMENT D'ENCASTREMENT EN TÊTE M0 (kN.m)"
    ws.Range("Y" & r).Value = "ROTATION EN TÊTE " & ChrW(952) & "0 (rad)"
    ws.Range("AC" & r).Value = "MOMENT EN TÊTE DU MICROPIEU Mh (kN.m)"
    Commentaire ws.Range("Y" & r), ChrW(952) & "0 = -(b H + c Mh) en tête libre, 0 en tête encastrée"
    Commentaire ws.Range("AC" & r), "Mh = M (tableau des efforts) en tête libre ; Mh = -M0 = -b/c x H en " & _
        "tête encastrée"
    For k = 0 To 7
        ws.Range(ws.Cells(r, 1 + k), ws.Cells(r + 1, 1 + k)).Merge
    Next k
    For j = 0 To 5
        ws.Range(ws.Cells(r, 9 + 4 * j), ws.Cells(r, 12 + 4 * j)).Merge
        ws.Range(ws.Cells(r + 1, 9 + 4 * j), ws.Cells(r + 1, 12 + 4 * j)).Value = _
            Array("ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
    Next j
    Commentaire ws.Range("B" & r), "Libre en tête : rotation libre, y0 = a H + b M ; encastré en " & _
        "tête : rotation bloquée par la semelle, y0 = (a - b²/c) H (valeur par défaut : « Données du " & _
        "projet », conditions en tête)"
    Commentaire ws.Range("C" & r), "Inclinaison et angle du micropieu : vérification de la " & _
        "profondeur (« Calcul et vérification »)"
    Commentaire ws.Range("E" & r), "x = D / l0 : micropieu souple si x >= 3, rigide si x <= 1"
    Commentaire ws.Range("F" & r), "Souplesses en tête (Hetényi, poutre de longueur finie, extrémité " & _
        "libre) : a = 2" & ChrW(955) & "/Es (sh ch - s c)/(sh² - s²), b = 2" & ChrW(955) & "²/Es (sh² + " & _
        "s²)/(sh² - s²), c = 4" & ChrW(955) & "³/Es (sh ch + s c)/(sh² - s²), " & ChrW(955) & " = 1/l0, " & _
        "sh = sinh x, s = sin x"
    Commentaire ws.Range("I" & r), "T0 : effort horizontal dû à l'inclinaison du micropieu, " & _
        "composante perpendiculaire à son axe de l'effort vertical de la descente de charges " & _
        "(Nc de la combinaison, tableau des efforts) : T0 = Nc x sin " & ChrW(945) & " ; 0 si le " & _
        "micropieu n'est pas incliné"
    Commentaire ws.Range("M" & r), "H = V + T0 : effort tranchant de la poussée des terres (tableau " & _
        "des efforts) augmenté de T0 (même sens, hypothèse défavorable)"
    Commentaire ws.Range("Q" & r), "y0 : déplacement horizontal en tête, comportement linéaire (à " & _
        "l'ELU, valeur approchée si la réaction dépasse le palier de fluage Pf)"
    Commentaire ws.Range("U" & r), "M0 = b/c x H : moment repris par l'encastrement en tête (tête " & _
        "encastrée seulement)"
    For i = 1 To m
        r = rY + 3 + i
        x = CStr(r)
        rl0 = CStr(rL + 2 + i)                ' ligne du micropieu dans le tableau de l0
        ry0 = CStr(rv + 1 + i)                ' ligne du micropieu dans le tableau des efforts
        ws.Range("A" & r).Value = "MP" & i
        With ws.Range("B" & r)
            .Validation.Delete
            .Validation.Add xlValidateList, xlValidAlertStop, , "Libre en tête,Encastré en tête"
            .Value = IIf(CStr(ActiveWorkbook.Worksheets("Données du projet").Range("D" & lp + 18).Value) = _
                "Encastré en tête", "Encastré en tête", "Libre en tête")
            If IsArray(condTete) Then
                If i <= UBound(condTete) Then
                    If CStr(condTete(i)) = "Libre en tête" Or CStr(condTete(i)) = "Encastré en tête" Then _
                        .Value = condTete(i)
                End If
            End If
        End With
        If cIn > 0 Then
            ws.Range("C" & r).Formula = "=IF(" & CALCUL & wc.Cells(r0 + 2 + i, cIn).Address & "=""Oui"",""Oui"",""Non"")"
        Else
            ws.Range("C" & r).Value = "Non"
        End If
        If cAn > 0 Then
            ws.Range("D" & r).Formula = "=IF(C" & x & "=""Oui"",N(" & CALCUL & wc.Cells(r0 + 2 + i, cAn).Address & "),0)"
        Else
            ws.Range("D" & r).Value = 0
        End If
        ws.Range("E" & r).Formula = "=IF(OR(N(W" & rl0 & ")<=0,N(X" & rl0 & ")<=0),"""",X" & rl0 & "/W" & rl0 & ")"
        xs = "E" & x
        ' (sh ch - s c), (sh² + s²), (sh ch + s c) divisés par (sh² - s²) ; Es en kPa
        ws.Range("F" & r).Formula = "=IF(" & xs & "="""","""",2/W" & rl0 & "/(V" & rl0 & "*1000)*" & _
            "(SINH(" & xs & ")*COSH(" & xs & ")-SIN(" & xs & ")*COS(" & xs & "))/(SINH(" & xs & ")^2-SIN(" & xs & ")^2))"
        ws.Range("G" & r).Formula = "=IF(" & xs & "="""","""",2/W" & rl0 & "^2/(V" & rl0 & "*1000)*" & _
            "(SINH(" & xs & ")^2+SIN(" & xs & ")^2)/(SINH(" & xs & ")^2-SIN(" & xs & ")^2))"
        ws.Range("H" & r).Formula = "=IF(" & xs & "="""","""",4/W" & rl0 & "^3/(V" & rl0 & "*1000)*" & _
            "(SINH(" & xs & ")*COSH(" & xs & ")+SIN(" & xs & ")*COS(" & xs & "))/(SINH(" & xs & ")^2-SIN(" & xs & ")^2))"
        For k = 0 To 3
            ' T0 = Nc x sin(angle) (Nc : colonne 3 + k du tableau des efforts), H = V + T0
            ' (V : colonne 11 + k), M (colonne 15 + k)
            ws.Cells(r, 9 + k).Formula = "=IF(C" & x & "=""Oui"",ABS(N(" & _
                ws.Cells(CLng(ry0), 3 + k).Address(False, False) & "))*SIN(RADIANS(D" & x & ")),0)"
            ws.Cells(r, 13 + k).Formula = "=N(" & ws.Cells(CLng(ry0), 11 + k).Address(False, False) & ")+" & _
                ws.Cells(r, 9 + k).Address(False, False)
            ws.Cells(r, 17 + k).Formula = "=IF(F" & x & "="""","""",1000*IF(B" & x & "=""Encastré en tête""," & _
                "(F" & x & "-G" & x & "^2/H" & x & ")*" & ws.Cells(r, 13 + k).Address(False, False) & "," & _
                "F" & x & "*" & ws.Cells(r, 13 + k).Address(False, False) & "+G" & x & "*N(" & _
                ws.Cells(CLng(ry0), 15 + k).Address(False, False) & ")))"
            ' Mh (moment en tête du micropieu) et rotation en tête theta0
            ws.Cells(r, 29 + k).Formula = "=IF(F" & x & "="""","""",IF(B" & x & "=""Encastré en tête"",-G" & x & _
                "/H" & x & "*" & ws.Cells(r, 13 + k).Address(False, False) & ",N(" & _
                ws.Cells(CLng(ry0), 15 + k).Address(False, False) & ")))"
            ws.Cells(r, 25 + k).Formula = "=IF(F" & x & "="""","""",IF(B" & x & "=""Encastré en tête"",0,-(G" & x & _
                "*" & ws.Cells(r, 13 + k).Address(False, False) & "+H" & x & "*" & _
                ws.Cells(r, 29 + k).Address(False, False) & ")))"
            ws.Cells(r, 21 + k).Formula = "=IF(OR(F" & x & "="""",B" & x & "<>""Encastré en tête""),""-""," & _
                "G" & x & "/H" & x & "*" & ws.Cells(r, 13 + k).Address(False, False) & ")"
        Next k
    Next i
    With ws.Range("A" & rY + 2 & ":AF" & rY + 3 + m)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A" & rY + 2 & ":AF" & rY + 3)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .WrapText = True
    End With
    ws.Range("A" & rY + 4 & ":A" & rY + 3 + m).Font.Bold = True
    ws.Range("C" & rY + 4 & ":AF" & rY + 3 + m).Interior.Color = JAUNE_CALCUL
    ws.Range("Y" & rY + 4 & ":AB" & rY + 3 + m).NumberFormat = "0.000E+00"
    ws.Range("AC" & rY + 4 & ":AF" & rY + 3 + m).NumberFormat = "0.0"
    ws.Range("B" & rY + 4 & ":B" & rY + 3 + m).Interior.Color = VERT_SAISIE
    ws.Range("D" & rY + 4 & ":D" & rY + 3 + m).NumberFormat = "0.0"
    ws.Range("E" & rY + 4 & ":E" & rY + 3 + m).NumberFormat = "0.00"
    ws.Range("F" & rY + 4 & ":H" & rY + 3 + m).NumberFormat = "0.000E+00"
    ws.Range("I" & rY + 4 & ":P" & rY + 3 + m & ",U" & rY + 4 & ":X" & rY + 3 + m).NumberFormat = "0.0"
    ws.Range("Q" & rY + 4 & ":T" & rY + 3 + m).NumberFormat = "0.00"
    ws.Range("Q" & rY + 4 & ":T" & rY + 3 + m).Font.Bold = True

    ' --- Y, T et M le long du micropieu (z de 0 à la profondeur réelle D) ---
    ' Équation EI y'''' + Es y = 0, lambda = 1 / l0. Solution du pieu long à pointe
    ' encastrée : onde issue de la tête et onde issue de la pointe (u = D - z) :
    '   y(z) = e^(-lambda z) [A cos lambda z + B sin lambda z]
    '        + e^(-lambda u) [C cos lambda u + E sin lambda u]
    ' Tête (H = V + T0 et M du tableau des déplacements, Es en kPa) :
    '   tête libre    : A = 2 lambda H / Es + 2 lambda² M / Es, B = -2 lambda² M / Es ;
    '   tête encastrée : A = B = lambda H / Es ;
    ' pointe encastrée (y(D) = 0, y'(D) = 0) :
    '   C = -e^(-lambda D) [A cos lambda D + B sin lambda D]
    '   E = e^(-lambda D) [(B - 2A) cos lambda D - (A + 2B) sin lambda D] ;
    ' moment M(z) = EI y'' et effort tranchant T(z) = EI y''' :
    '   M = 2 EI lambda² { e^(-lambda z) [A sin - B cos](lambda z) + e^(-lambda u) [C sin - E cos](lambda u) }
    '   T = 2 EI lambda³ { e^(-lambda z) [(A + B) cos + (B - A) sin](lambda z)
    '                    - e^(-lambda u) [(C + E) cos + (E - C) sin](lambda u) }
    ' Valable pour un micropieu long (D >= 3 l0) ; D = profondeur réelle (« Calcul et
    ' vérification »). Profondeurs : 0, pas de calcul du tableau de la capacité portante
    ' (« Calcul et vérification ») jusqu'à D, puis D. Graphiques à la fin de la feuille.
    etape = "Y, T et M le long du micropieu"
    With ws.Range("A" & rT)
        .Value = "DÉPLACEMENT Y, EFFORT TRANCHANT T ET MOMENT M LE LONG DU MICROPIEU (PIEU LONG, " & _
            "POINTE ENCASTRÉE)"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
    ws.Range("A" & rT + 1).Value = "Constantes de la solution y(z) = e^(-" & ChrW(955) & "z)[A cos " & _
        ChrW(955) & "z + B sin " & ChrW(955) & "z] + e^(-" & ChrW(955) & "u)[C cos " & ChrW(955) & "u + E sin " & _
        ChrW(955) & "u], u = D - z (A, B, C, E en m)"
    ' Tableau des constantes : A micropieu, B lambda, C D réelle, D x = D / l0, E validité,
    ' puis A, B, C, E de chaque combinaison (colonnes 6 à 21)
    r = rT + 2
    ws.Range("A" & r & ":E" & r).Value = Array("N° MICROPIEU", ChrW(955) & " = 1 / l0 (1/m)", _
        "Profondeur réelle D (m)", "x = D / l0", "Validité (pieu long D >= 3 l0)")
    For k = 0 To 4
        ws.Range(ws.Cells(r, 1 + k), ws.Cells(r + 1, 1 + k)).Merge
    Next k
    For k = 0 To 3
        ws.Cells(r, 6 + 4 * k).Value = Choose(k + 1, "ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
        ws.Range(ws.Cells(r, 6 + 4 * k), ws.Cells(r, 9 + 4 * k)).Merge
        ws.Range(ws.Cells(r + 1, 6 + 4 * k), ws.Cells(r + 1, 9 + 4 * k)).Value = Array("A", "B", "C", "E")
    Next k
    Commentaire ws.Range("E" & r), "Solution du pieu long : écart inférieur à e^(-2D/l0) (0,25 % pour " & _
        "D = 3 l0) ; pour un micropieu plus court, valeurs approchées (voir le tableau des déplacements " & _
        "en tête, solution de longueur finie)"
    For i = 1 To m
        r = rT + 3 + i
        x = CStr(rY + 3 + i)                  ' ligne du micropieu dans le tableau des déplacements
        rl0 = CStr(rL + 2 + i)                ' ligne du micropieu dans le tableau de l0
        rd0 = CStr(rv + 1 + i)                ' ligne du micropieu dans le tableau des efforts
        ws.Range("A" & r).Value = "MP" & i
        ws.Range("B" & r).Formula = "=IF(N(W" & rl0 & ")<=0,"""",1/W" & rl0 & ")"
        ws.Range("C" & r).Formula = "=N(X" & rl0 & ")"
        ws.Range("D" & r).Formula = "=IF(B" & r & "="""","""",B" & r & "*C" & r & ")"
        ws.Range("E" & r).Formula = "=IF(D" & r & "="""",""-"",IF(D" & r & ">=3,""Pieu long"",""D < 3 l0 : approché""))"
        For k = 0 To 3
            hs = ws.Cells(CLng(x), 13 + k).Address(False, False)               ' H = V + T0
            mhs = "N(" & ws.Cells(CLng(rd0), 15 + k).Address(False, False) & ")"  ' M
            cond0 = "=IF(OR($B" & r & "="""",N(V" & rl0 & ")<=0),"""","
            kA = ws.Cells(r, 6 + 4 * k).Address(False, False)
            kB = ws.Cells(r, 7 + 4 * k).Address(False, False)
            ' A, B : conditions en tête (libre ou encastrée, tableau des déplacements)
            ws.Cells(r, 6 + 4 * k).Formula = cond0 & "IF($B$" & x & "=""Encastré en tête"",$B" & r & "*" & hs & _
                "/(V" & rl0 & "*1000),2*$B" & r & "*" & hs & "/(V" & rl0 & "*1000)+2*$B" & r & "^2*" & mhs & _
                "/(V" & rl0 & "*1000)))"
            ws.Cells(r, 7 + 4 * k).Formula = cond0 & "IF($B$" & x & "=""Encastré en tête""," & kA & _
                ",-2*$B" & r & "^2*" & mhs & "/(V" & rl0 & "*1000)))"
            ' C, E : pointe encastrée
            ws.Cells(r, 8 + 4 * k).Formula = "=IF(" & kA & "="""","""",-EXP(-$D" & r & ")*(" & kA & "*COS($D" & r & _
                ")+" & kB & "*SIN($D" & r & ")))"
            ws.Cells(r, 9 + 4 * k).Formula = "=IF(" & kA & "="""","""",EXP(-$D" & r & ")*((" & kB & "-2*" & kA & _
                ")*COS($D" & r & ")-(" & kA & "+2*" & kB & ")*SIN($D" & r & ")))"
        Next k
    Next i
    With ws.Range("A" & rT + 2 & ":U" & rT + 3 + m)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A" & rT + 2 & ":U" & rT + 3)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .WrapText = True
    End With
    ws.Range("A" & rT + 4 & ":A" & rT + 3 + m).Font.Bold = True
    ws.Range("B" & rT + 4 & ":U" & rT + 3 + m).Interior.Color = JAUNE_CALCUL
    ws.Range("B" & rT + 4 & ":D" & rT + 3 + m).NumberFormat = "0.000"
    ws.Range("F" & rT + 4 & ":U" & rT + 3 + m).NumberFormat = "0.000E+00"
    With ws.Range("E" & rT + 4 & ":E" & rT + 3 + m).FormatConditions
        .Delete
        .Add(xlCellValue, xlEqual, "=""D < 3 l0 : approché""").Font.Color = RGB(192, 0, 0)
    End With

    ' Profils : un tableau par micropieu (côte à côte), z, lambda z, lambda u, puis Y (mm),
    ' T (kN), M (kN.m) de chaque combinaison
    ws.Range("A" & rP - 1).Value = "Profils le long du micropieu : z = 0, pas de calcul du tableau de la " & _
        "capacité portante (« Calcul et vérification ») jusqu'à la profondeur réelle D, puis D ; " & _
        "Y en mm, T en kN, M en kN.m"
    For i = 1 To m
        cb = 1 + LARG_PROFIL * (i - 1)        ' 1re colonne du tableau du micropieu
        rl0 = CStr(rL + 2 + i)
        pr = CStr(rT + 3 + i)                 ' ligne des constantes du micropieu
        ws.Cells(rP, cb).Value = "MP" & i
        ws.Range(ws.Cells(rP, cb), ws.Cells(rP, cb + 14)).Merge
        ws.Range(ws.Cells(rP + 1, cb), ws.Cells(rP + 1, cb + 2)).Value = Array("z (m)", _
            ChrW(955) & "z", ChrW(955) & "u = " & ChrW(955) & "(D - z)")
        For k = 0 To 2
            ws.Range(ws.Cells(rP + 1, cb + k), ws.Cells(rP + 2, cb + k)).Merge
        Next k
        For k = 0 To 3
            ws.Cells(rP + 1, cb + 3 + 3 * k).Value = Choose(k + 1, "ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
            ws.Range(ws.Cells(rP + 1, cb + 3 + 3 * k), ws.Cells(rP + 1, cb + 5 + 3 * k)).Merge
            ws.Range(ws.Cells(rP + 2, cb + 3 + 3 * k), ws.Cells(rP + 2, cb + 5 + 3 * k)).Value = _
                Array("Y (mm)", "T (kN)", "M (kN.m)")
        Next k
        For j = 0 To nZ + 1
            r = rP + 3 + j
            xs = ws.Cells(r, cb).Address(False, False)
            ' z : 0, puis profondeurs du pas de calcul inférieures à D, puis D
            If j = 0 Then
                ws.Cells(r, cb).Formula = "=IF(N($C$" & pr & ")>0,0,"""")"
            Else
                hs = ws.Cells(r - 1, cb).Address(False, False)
                If j <= nZ And Not IsError(hZ) Then
                    zk = "N(" & CALCUL & "$C$" & CLng(hZ) + j & ")"
                Else
                    zk = "0"
                End If
                ws.Cells(r, cb).Formula = "=IF(" & hs & "="""","""",IF(" & hs & ">=$C$" & pr & "-1E-9,""""," & _
                    "IF(AND(" & zk & ">" & hs & "," & zk & "<$C$" & pr & "-1E-9)," & zk & ",$C$" & pr & ")))"
            End If
            ws.Cells(r, cb + 1).Formula = "=IF(" & xs & "="""",""""," & "$B$" & pr & "*" & xs & ")"
            ws.Cells(r, cb + 2).Formula = "=IF(" & xs & "="""",""""," & "$B$" & pr & "*($C$" & pr & "-" & xs & "))"
            kA = ws.Cells(r, cb + 1).Address(False, False)   ' lambda z
            kB = ws.Cells(r, cb + 2).Address(False, False)   ' lambda u
            For k = 0 To 3
                y0s = ws.Cells(CLng(pr), 6 + 4 * k).Address   ' A
                hs = ws.Cells(CLng(pr), 7 + 4 * k).Address    ' B
                ths = ws.Cells(CLng(pr), 8 + 4 * k).Address   ' C
                mhs = ws.Cells(CLng(pr), 9 + 4 * k).Address   ' E
                cond0 = "=IF(OR(" & xs & "=""""," & y0s & "=""""),"""","
                ws.Cells(r, cb + 3 + 3 * k).Formula = cond0 & "1000*(EXP(-" & kA & ")*(" & y0s & "*COS(" & kA & _
                    ")+" & hs & "*SIN(" & kA & "))+EXP(-" & kB & ")*(" & ths & "*COS(" & kB & ")+" & mhs & _
                    "*SIN(" & kB & "))))"
                ws.Cells(r, cb + 4 + 3 * k).Formula = cond0 & "2*$Q$" & rl0 & "*$B$" & pr & "^3*(EXP(-" & kA & _
                    ")*((" & y0s & "+" & hs & ")*COS(" & kA & ")+(" & hs & "-" & y0s & ")*SIN(" & kA & "))-EXP(-" & _
                    kB & ")*((" & ths & "+" & mhs & ")*COS(" & kB & ")+(" & mhs & "-" & ths & ")*SIN(" & kB & "))))"
                ws.Cells(r, cb + 5 + 3 * k).Formula = cond0 & "2*$Q$" & rl0 & "*$B$" & pr & "^2*(EXP(-" & kA & _
                    ")*(" & y0s & "*SIN(" & kA & ")-" & hs & "*COS(" & kA & "))+EXP(-" & kB & ")*(" & ths & _
                    "*SIN(" & kB & ")-" & mhs & "*COS(" & kB & "))))"
            Next k
        Next j
        With ws.Range(ws.Cells(rP, cb), ws.Cells(rP + 4 + nZ, cb + 14))
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        With ws.Range(ws.Cells(rP, cb), ws.Cells(rP + 2, cb + 14))
            .Font.Bold = True
            .Interior.Color = VERT_ENTETE
            .WrapText = True
        End With
        ws.Range(ws.Cells(rP + 3, cb), ws.Cells(rP + 4 + nZ, cb + 14)).Interior.Color = JAUNE_CALCUL
        ws.Range(ws.Cells(rP + 3, cb), ws.Cells(rP + 4 + nZ, cb + 2)).NumberFormat = "0.000"
        ws.Range(ws.Cells(rP + 3, cb + 3), ws.Cells(rP + 4 + nZ, cb + 14)).NumberFormat = "0.00"
        ws.Range(ws.Cells(rP + 3, cb), ws.Cells(rP + 4 + nZ, cb)).Font.Bold = True
    Next i

    ' --- Efforts dans les micropieux (sous le tableau des déplacements) ---
    etape = "efforts dans les micropieux"
    With ws.Range("A" & rv - 2)
        .Value = "EFFORTS DANS LES MICROPIEUX"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
    ' Poussée des terres au repos (rappel) : h = hp + hs, K0, H et e
    ws.Range("B" & rv - 1).Formula = "=""Poussée des terres au repos : h = hp + hs = ""&TEXT(" & hT & _
        ",""0.00"")&"" m ; K0 = 1 - sin " & ChrW(966) & "' = ""&TEXT(" & k0 & ",""0.000"")&"" ; H = K0 " & _
        ChrW(947) & " h² / 2 = ""&TEXT(" & hH & ",""0.00"")&"" kN/ml ; e = h / 3 = ""&TEXT(" & eE & _
        ",""0.00"")&"" m"""
    ws.Range("A" & rv).Value = "N° MICROPIEU"
    ws.Range("B" & rv).Value = "Longueur de mur reprise L (m)"
    ws.Range("C" & rv).Value = "EFFORT DE COMPRESSION Nc (kN)"
    ws.Range("G" & rv).Value = "EFFORT DE TRACTION Nt (kN)"
    ws.Range("K" & rv).Value = "EFFORT TRANCHANT EN TÊTE V (kN)"
    ws.Range("O" & rv).Value = "MOMENT EN TÊTE M = V x e (kN.m)"
    ws.Range("A" & rv & ":A" & rv + 1).Merge
    ws.Range("B" & rv & ":B" & rv + 1).Merge
    For k = 0 To 3
        ws.Range(ws.Cells(rv, 3 + 4 * k), ws.Cells(rv, 6 + 4 * k)).Merge
        ws.Range(ws.Cells(rv + 1, 3 + 4 * k), ws.Cells(rv + 1, 6 + 4 * k)).Value = _
            Array("ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
    Next k
    Commentaire ws.Range("C" & rv), "Nc : Vmax de chaque combinaison (vérification de la profondeur, " & _
        "« Calcul et vérification »)"
    Commentaire ws.Range("G" & rv), "Nt : effort de traction, opposé de l'effort de compression"
    Commentaire ws.Range("K" & rv), "V = " & ChrW(947) & "G,sup x H x L : poussée des terres au repos " & _
        "sur la fondation existante (H en kN/ml) multipliée par la longueur de mur reprise L et par le " & _
        "coefficient " & ChrW(947) & "G,sup de la combinaison"
    Commentaire ws.Range("O" & rv), "M = V x e : e = h / 3, distance entre le point d'application de " & _
        "la poussée et la fibre inférieure de la semelle existante (tête du micropieu)"
    entete = Array("ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
    For i = 1 To m
        r = rv + 1 + i
        ws.Range("A" & r).Value = "MP" & i
        ' Longueur de mur reprise : somme de la ligne du micropieu dans la matrice
        ws.Range("B" & r).Formula = "=SUM(" & dn & wd.Range(wd.Cells(rd + i - 1, 3), _
            wd.Cells(rd + i - 1, 2 + n)).Address & ")"
        For k = 0 To 3
            ws.Cells(r, 3 + k).Formula = "=N(" & CALCUL & wc.Cells(r0 + 2 + i, cC + 2 * k).Address & ")"
            ws.Cells(r, 7 + k).Formula = "=-" & ws.Cells(r, 3 + k).Address(False, False)
            ' gamma G,sup de la combinaison (tableau des coefficients, colonne D) ; 1 sinon
            If tC > 0 Then
                ws.Cells(r, 11 + k).Formula = "=N(" & CALCUL & "$D$" & tC + 2 + k & ")*" & hH & "*B" & r
            Else
                ws.Cells(r, 11 + k).Formula = "=" & Choose(k + 1, "1", "1", "1.35", "1") & "*" & hH & "*B" & r
            End If
            ws.Cells(r, 15 + k).Formula = "=" & ws.Cells(r, 11 + k).Address(False, False) & "*" & eE
        Next k
    Next i
    With ws.Range("A" & rv & ":R" & rv + 1 + m)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A" & rv & ":R" & rv + 1)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .WrapText = True
    End With
    ws.Range("A" & rv + 2 & ":A" & rv + 1 + m).Font.Bold = True
    ws.Range("B" & rv + 2 & ":R" & rv + 1 + m).Interior.Color = JAUNE_CALCUL
    ws.Range("B" & rv + 2 & ":R" & rv + 1 + m).NumberFormat = "0.0"

MiseEnForme:
    etape = "mise en forme"
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:Y").ColumnWidth = 12
    ' Graphiques Y, T et M le long de chaque micropieu, à la fin de la feuille (sous le
    ' tableau des efforts)
    If rP > 0 And m > 0 Then
        etape = "graphiques Y, T et M"
        ws.Calculate
        rG = rv + 1 + m + 3
        With ws.Range("A" & rG)
            .Value = "GRAPHIQUES - DÉPLACEMENT Y, EFFORT TRANCHANT T ET MOMENT M LE LONG DES MICROPIEUX"
            .Font.Bold = True
            .Font.Size = 14
            .Font.Color = VERT_FONCE
        End With
        For i = 1 To m
            GraphiquesProfil ws, i, 1 + LARG_PROFIL * (i - 1), rP, nZ, rG + 2 + (i - 1) * HAUT_GRAPHE
        Next i
    End If
    ws.Activate
Fin:
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    MsgBox "Erreur pendant la création de la feuille « " & NOM_FEUILLE & " », partie « " & etape & _
        " » : " & Err.Description & " (n° " & Err.Number & ")", vbExclamation
    Resume Fin
End Sub

' Une ligne « intitulé (B:C fusionnées) | valeur ou formule (D) | unité (E) » ; un
' intitulé commençant par « = » est une formule
Private Sub Ligne(ws As Worksheet, r As Long, intitule As String, formule As String, unite As String)
    If Left(intitule, 1) = "=" Then
        ws.Range("B" & r).Formula = intitule
    Else
        ws.Range("B" & r).Value = intitule
    End If
    ws.Range("B" & r & ":C" & r).Merge
    If Left(formule, 1) = "=" Then
        ws.Range("D" & r).Formula = formule
    Else
        ws.Range("D" & r).Value = formule
    End If
    ws.Range("E" & r).Value = unite
End Sub

' « Autre diamètre » saisis dans le tableau de la longueur de transfert (colonne C,
' lignes MP1...), conservés quand la feuille est recréée ; vide si absent
Private Function LireAutreDiametre(ws As Worksheet) As Variant
    Dim v As Variant, r As Long, i As Long, t() As Variant
    v = Application.Match("LONGUEUR DE TRANSFERT l0 DES MICROPIEUX (APRÈS CORROSION)", ws.Range("A1:A5000"), 0)
    If IsError(v) Then Exit Function
    r = CLng(v) + 3
    Do While Left(CStr(ws.Range("A" & r + i).Value), 2) = "MP"
        i = i + 1
    Loop
    If i = 0 Then Exit Function
    ReDim t(1 To i)
    For i = 1 To UBound(t)
        t(i) = ws.Range("C" & r + i - 1).Value
    Next i
    LireAutreDiametre = t
End Function

' Valeurs d'une colonne (col) d'un tableau de cette feuille (titre en colonne A, lignes
' MP1... à partir de la 4e ligne sous le titre), conservées quand la feuille est
' recréée ; vide si absent
Private Function LireColonne(ws As Worksheet, titre As String, col As String) As Variant
    Dim v As Variant, r As Long, i As Long, t() As Variant
    v = Application.Match(titre, ws.Range("A1:A5000"), 0)
    If IsError(v) Then Exit Function
    r = CLng(v) + 4
    Do While Left(CStr(ws.Range("A" & r + i).Value), 2) = "MP"
        i = i + 1
    Loop
    If i = 0 Then Exit Function
    ReDim t(1 To i)
    For i = 1 To UBound(t)
        t(i) = ws.Range(col & r + i - 1).Value
    Next i
    LireColonne = t
End Function

' Graphiques Y (mm), T (kN) et M (kN.m) le long du micropieu n° i en fonction de z
' (croissante vers le bas), à partir de la ligne rH : une courbe par combinaison, données
' du profil du micropieu (1re colonne cb, en-tête en ligne rP), lignes renseignées
Private Sub GraphiquesProfil(ws As Worksheet, i As Long, cb As Long, rP As Long, nZ As Long, rH As Long)
    Dim co As ChartObject, g As Long, k As Long, n As Long, r1 As Long, lg As Double, rz As Range
    Dim noms As Variant, titres As Variant
    r1 = rP + 3
    Do While n <= nZ + 1                      ' lignes renseignées (z numérique)
        If IsEmpty(ws.Cells(r1 + n, cb).Value) Or Not IsNumeric(ws.Cells(r1 + n, cb).Value) Then Exit Do
        If CStr(ws.Cells(r1 + n, cb).Value) = "" Then Exit Do
        n = n + 1
    Loop
    If n < 2 Then Exit Sub
    Set rz = ws.Range(ws.Cells(r1, cb), ws.Cells(r1 + n - 1, cb))
    noms = Array("ELS CARA", "ELS QP", "ELU FOND", "ELU ACC")
    titres = Array("Déplacement Y (mm)", "Effort tranchant T (kN)", "Moment M (kN.m)")
    lg = 380
    For g = 0 To 2
        With ws.Cells(rH, 1)
            Set co = ws.ChartObjects.Add(.Left + g * (lg + 10), .Top, lg, 360)
        End With
        co.Placement = xlFreeFloating
        With co.Chart
            .ChartType = xlXYScatterSmoothNoMarkers
            Do While .SeriesCollection.Count > 0
                .SeriesCollection(1).Delete
            Loop
            For k = 0 To 3
                With .SeriesCollection.NewSeries
                    .Name = noms(k)
                    .XValues = ws.Range(ws.Cells(r1, cb + 3 + 3 * k + g), ws.Cells(r1 + n - 1, cb + 3 + 3 * k + g))
                    .Values = rz
                    .ChartType = xlXYScatterSmoothNoMarkers
                    .Format.Line.Weight = 2
                End With
            Next k
            .HasTitle = True
            .ChartTitle.Text = "MP" & i & " - " & titres(g) & " le long du micropieu"
            .HasLegend = True
            .Legend.Position = xlLegendPositionBottom
            With .Axes(xlValue)                ' profondeur, croissante vers le bas
                .ReversePlotOrder = True
                .MinimumScale = 0
                .HasTitle = True
                .AxisTitle.Text = "Profondeur z (m)"
            End With
            With .Axes(xlCategory)
                .HasTitle = True
                .AxisTitle.Text = titres(g)
            End With
        End With
    Next g
End Sub

' Valeur numérique (0 si vide ou texte)
Private Function N0(v As Variant) As Double
    If Not IsEmpty(v) And IsNumeric(v) Then N0 = CDbl(v)
End Function

' Sous-titre en gras, souligné
Private Sub SousTitre(ws As Worksheet, r As Long, texte As String)
    With ws.Range("B" & r)
        .Value = texte
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
End Sub

' Commentaire (signification) sur une cellule d'en-tête
Private Sub Commentaire(c As Range, texte As String)
    c.ClearComments
    c.AddComment texte
    c.Comment.Shape.TextFrame.AutoSize = True
End Sub
