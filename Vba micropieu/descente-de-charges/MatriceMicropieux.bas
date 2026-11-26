Attribute VB_Name = "MatriceMicropieux"
Option Explicit
' Couleurs de la charte graphique GEOSKILLS (bleu nuit, bleu océan, vert émeraude)
Private Const COUL_ENTETE As Long = 14337945       ' RGB(153, 199, 218) : Bleu Océan moyen (en-têtes)
Private Const COUL_SAISIE As Long = 15201485       ' RGB(205, 244, 231) : Vert Émeraude éclairci
Private Const COUL_CALCUL As Long = 15590349       ' RGB(205, 227, 237) : Bleu Océan clair (valeurs calculées)

' Module « Matrice des micropieux » : sous les tableaux de descente de charges des murs
' (à partir de la ligne LIGNE_MATRICE), répartition des charges des murs sur les
' micropieux.
'   L+1 à L+3 : tableau Élément / Poids volumique (kN/m3) / Base b (m) / Hauteur h (m)
'               avec une ligne Massif (carré b x b, hauteur h) et une ligne Longrine
'               (section b x h)
'   L+5       : charges des murs, matrice n x 3 (n = nombre de murs, G1) : Mur 01,
'               Mur 02, ... avec G et Q (kN/ml) et G poteaux (kN) du récapitulatif de
'               chaque mur
'   plus bas  : matrice m x n des longueurs de murs reprises par chaque micropieu
'               (m = nombre de micropieux, K1 ; lignes MP1, MP2, ..., colonnes Mur 01,
'               Mur 02, ...) : longueur Li du mur sur la diagonale, 0 ailleurs, les
'               valeurs peuvent être modifiées ; à côté, matrice IDENTIFICATION POTEAU
'               (0 / 1 par micropieu et par mur : 1 = charge des poteaux du mur ajoutée) ;
'               puis, à droite :
'               CHARGES DES MURS G (kN) = longueurs x G (kN/ml) + identification x G
'               poteaux (kN) ; Q (kN) = longueurs x Q (kN/ml) ;
'               STRUCTURE GC : M (massif) ou L (longrine) ;
'               POIDS MASSIF (kN) = poids volumique x b² x h (si M, sinon 0) ;
'               POIDS LONGRINE (kN) = poids volumique x section x longueur, longueur =
'               somme des longueurs de murs de la ligne du micropieu / 2 (si L, sinon 0) ;
'               POIDS TOTAL ÉLÉMENTAIRE DDC : G = G des murs + massif + longrine, Q ;
'               CHARGES DE L'EAU EB (basses eaux), EH (hautes eaux), EE (eaux
'               exceptionnelles) en kN, à saisir : colonnes créées seulement s'il y a
'               une nappe phréatique (« Données du projet », D34) ;
'               CHARGE DU TALUS en fin de tableau : H0 (kN) et M0 (kN.m) en tête de
'               chaque micropieu, à saisir (repris dans « Vérification structurelle »).
' Les saisies (dimensions, longueurs Li, longueurs modifiées de la matrice, choix M / L,
' charges de l'eau, charge du talus)
' sont conservées quand le tableau est recréé (nombre de murs ou de micropieux changé).
' Appelé par DupliquerMurs (nombre de murs, G1) et quand K1 change (CodeFeuille.txt) ;
' met ensuite à jour la vérification de la profondeur des micropieux de « Calcul et
' vérification » (ActualiserVerificationProfondeur, module CalculFondation).

Public Const LIGNE_MATRICE As Long = 80             ' titre du tableau
Public Const CELLULE_NB_MICROPIEUX As String = "K1"

' Charte GEOSKILLS : saisies en vert émeraude clair, valeurs calculées en
' bleu très clair, en-têtes en bleu océan clair, titres en bleu nuit

Public Sub ConstruireMatrice(Optional ws As Worksheet)
    Dim n As Long, m As Long, i As Long, j As Long, L As Long, r0 As Long, rG As Long
    Dim nOld As Long, mOld As Long, r0Old As Long, x0 As Long, r As Long, derniere As Long
    Dim li() As Variant, mat() As Variant, gc() As Variant, dims As Variant, dimsN() As Variant, eau() As Variant
    Dim ligne As String, cG As String, cQ As String, c As Range, evenements As Boolean
    Dim nappe As Boolean, f As Worksheet, xf As Long, xt As Long, talus() As Variant, cT As Long
    Dim pot() As Variant, x0Old As Long, cP As String, ligneP As String, cLo As Long, cIo As Long
    Dim typeF As String, isolee As Boolean, nom As String, nL As Long, nP As Long, fG As String, fQ As String

    If ws Is Nothing Then Set ws = ActiveSheet
    L = LIGNE_MATRICE
    n = Val(ws.Range("G1").Value)
    If n < 1 Then n = 1
    m = Val(ws.Range(CELLULE_NB_MICROPIEUX).Value)
    If m < 1 Then m = 1

    evenements = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Fin

    ' --- Saisies existantes, conservées ---
    ReDim li(1 To n)
    ReDim mat(1 To m, 1 To n)
    ReDim gc(1 To m)
    ReDim eau(1 To m, 1 To 3)
    ReDim talus(1 To m, 1 To 2)
    ReDim pot(1 To m, 1 To n)
    dims = ws.Range("C" & L + 1 & ":E" & L + 3).Value
    If Left(CStr(ws.Range("B" & L + 1).Value), 7) <> "Élément" Then
        ' Ancienne disposition (gamma ; massif l, b, h ; longrine b, h) : conversion
        ReDim dimsN(1 To 3, 1 To 3)
        dimsN(2, 1) = dims(1, 1): dimsN(2, 2) = dims(2, 1): dimsN(2, 3) = dims(2, 3)
        dimsN(3, 1) = dims(1, 1): dimsN(3, 2) = dims(3, 1): dimsN(3, 3) = dims(3, 2)
        dims = dimsN
    End If
    Do While Left(CStr(ws.Range("B" & L + 6 + nOld).Value), 4) = "Mur " Or _
            Left(CStr(ws.Range("B" & L + 6 + nOld).Value), 4) = "Pot "
        nOld = nOld + 1
    Loop
    r0Old = L + 8 + nOld
    If CStr(ws.Range("B" & r0Old).Value) = "N° MICROPIEU" Then
        Do While Left(CStr(ws.Range("B" & r0Old + 3 + mOld).Value), 2) = "MP"
            mOld = mOld + 1
        Loop
        ' Blocs de l'ancien tableau repérés par leur en-tête : longueurs des murs,
        ' identification des poteaux, charges (1re colonne des résultats)
        cLo = 0: cIo = 0: x0Old = 0
        For j = 3 To 3 + 2 * nOld
            If Left(CStr(ws.Cells(r0Old, j).Value), 9) = "LONGUEURS" Then cLo = j
            If Left(CStr(ws.Cells(r0Old, j).Value), 14) = "IDENTIFICATION" Then cIo = j
            If Left(CStr(ws.Cells(r0Old, j).Value), 11) = "CHARGES DES" Then x0Old = j: Exit For
        Next j
        If x0Old = 0 Then x0Old = 3 + nOld
        For j = 1 To Application.Min(n, nOld)
            If cLo > 0 Then li(j) = ws.Cells(r0Old + 2, cLo - 1 + j).Value
        Next j
        For i = 1 To Application.Min(m, mOld)
            For j = 1 To Application.Min(n, nOld)
                If cLo > 0 Then
                    Set c = ws.Cells(r0Old + 2 + i, cLo - 1 + j)
                    If Not c.HasFormula Then mat(i, j) = c.Value
                End If
                If cIo > 0 Then pot(i, j) = ws.Cells(r0Old + 2 + i, cIo - 1 + j).Value
            Next j
            gc(i) = ws.Cells(r0Old + 2 + i, x0Old + 2).Value
            For j = 1 To 3                   ' eau EB, EH, EE
                eau(i, j) = ws.Cells(r0Old + 2 + i, x0Old + 6 + j).Value
            Next j
        Next i
        ' Charge du talus H0, M0 saisies (colonnes repérées par leur en-tête)
        For j = 3 To 250
            If CStr(ws.Cells(r0Old, j).Value) = "CHARGE DU TALUS" Then cT = j: Exit For
        Next j
        If cT > 0 Then
            For i = 1 To Application.Min(m, mOld)
                talus(i, 1) = ws.Cells(r0Old + 2 + i, cT).Value
                talus(i, 2) = ws.Cells(r0Old + 2 + i, cT + 1).Value
            Next i
        End If
    End If

    ' --- Ancien tableau effacé (lignes à partir de L) ---
    derniere = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If derniere < L + 200 Then derniere = L + 200
    ws.Range(ws.Rows(L), ws.Rows(derniere)).Clear

    ' --- Présence de nappe phréatique (« Données du projet », D34) : sinon eau = 0 ---
    For Each f In ws.Parent.Worksheets
        If f.Name = "Données du projet" Then nappe = (f.Range("D34").Value = "Oui")
    Next f

    ' --- Ligne du récapitulatif G / Q des murs (sous le tableau de chaque mur) ---
    For r = 5 To L - 1
        If CStr(ws.Range("B" & r).Value) = "G (kN/ml)" Or CStr(ws.Range("B" & r).Value) = "G (kN)" Then rG = r
    Next r

    ' --- Titre et données d'entrée ---
    With ws.Range("B" & L)
        .Value = "RÉPARTITION DES CHARGES DES MURS SUR LES MICROPIEUX"
        .Font.Bold = True
        .Font.Size = 12
    End With
    ' Tableau : une ligne par élément (massif carré b x b x h, longrine b x h)
    ws.Range("B" & L + 1 & ":E" & L + 1).Value = _
        Array("Élément", "Poids volumique (kN/m3)", "Base b (m)", "Hauteur h (m)")
    With ws.Range("B" & L + 1 & ":E" & L + 1)
        .Font.Bold = True
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = COUL_ENTETE
    End With
    ws.Range("B" & L + 2).Value = "Massif"
    ws.Range("B" & L + 3).Value = "Longrine"
    For i = 2 To 3
        ws.Range("C" & L + i & ":E" & L + i).Value = _
            Array(IIf(IsEmpty(dims(i, 1)), 25, dims(i, 1)), dims(i, 2), dims(i, 3))
    Next i
    ws.Range("C" & L + 2 & ":E" & L + 3).Interior.Color = COUL_SAISIE
    ws.Range("B" & L + 1 & ":E" & L + 3).Borders.LineStyle = xlContinuous

    ' --- Matrice n x 3 : G et Q (kN/ml) et charge des poteaux G (kN) de chaque mur
    ' (récapitulatif du mur k, colonne C du tableau d'origine + (k - 1) x PAS_MUR) ---
    ' Type de fondation (M1) : semelle isolée -> descente de charges sur les poteaux
    ' (Pot 01, ...), charges en kN, matrice d'identification seule ; semelle filante ->
    ' matrice des longueurs seule ; mixte -> les deux
    typeF = CStr(ws.Range("M1").Value)
    isolee = (typeF = "Semelle isolée")
    nom = IIf(isolee, "Pot ", "Mur ")
    nL = IIf(isolee, 0, n)                    ' colonnes de la matrice des longueurs
    nP = IIf(typeF = "Semelle filante", 0, n) ' colonnes de la matrice d'identification
    If isolee Then
        ws.Range("B" & L + 5 & ":E" & L + 5).Value = Array("POTEAU", "G (kN)", "Q (kN)", "G poteaux (kN)")
    Else
        ws.Range("B" & L + 5 & ":E" & L + 5).Value = Array("MUR", "G (kN/ml)", "Q (kN/ml)", "G poteaux (kN)")
    End If
    For j = 1 To n
        r = L + 5 + j
        ws.Range("B" & r).Value = nom & Format(j, "00")
        If rG > 0 Then
            ws.Range("C" & r).Formula = "=N(" & ws.Cells(rG, 3 + (j - 1) * PAS_MUR).Address & ")"
            ws.Range("D" & r).Formula = "=N(" & ws.Cells(rG + 1, 3 + (j - 1) * PAS_MUR).Address & ")"
            ws.Range("E" & r).Formula = "=N(" & ws.Cells(rG + 2, 3 + (j - 1) * PAS_MUR).Address & ")"
        End If
    Next j
    ws.Range("C" & L + 6 & ":E" & L + 5 + n).Interior.Color = COUL_CALCUL
    cG = "$C$" & L + 6 & ":$C$" & L + 5 + n
    cQ = "$D$" & L + 6 & ":$D$" & L + 5 + n
    cP = "$E$" & L + 6 & ":$E$" & L + 5 + n

    ' --- Matrice m x n des longueurs et résultats ---
    r0 = L + 8 + n                            ' 1re ligne d'en-tête
    x0 = 3 + nL + nP                          ' 1re colonne après les matrices
    ws.Cells(r0, 2).Value = "N° MICROPIEU"
    ws.Range(ws.Cells(r0, 2), ws.Cells(r0 + 1, 2)).Merge
    If nL > 0 Then
        ws.Cells(r0, 3).Value = "LONGUEURS DES MURS REPRISES PAR LE MICROPIEU (m)"
        If n > 1 Then ws.Range(ws.Cells(r0, 3), ws.Cells(r0, 2 + n)).Merge
    End If
    ' Matrice IDENTIFICATION POTEAU (0 / 1) à côté : 1 = charge des poteaux (Pot n en
    ' semelle isolée, poteaux du mur en mixte) ajoutée à la charge du micropieu, 0 = non
    If nP > 0 Then
        ws.Cells(r0, 3 + nL).Value = "IDENTIFICATION POTEAU (0 / 1)"
        If n > 1 Then ws.Range(ws.Cells(r0, 3 + nL), ws.Cells(r0, 2 + nL + nP)).Merge
    End If
    ws.Cells(r0, x0).Value = IIf(isolee, "CHARGES DES POTEAUX", "CHARGES DES MURS")
    ws.Range(ws.Cells(r0, x0), ws.Cells(r0, x0 + 1)).Merge
    ws.Cells(r0, x0 + 2).Value = "STRUCTURE GC"
    ws.Cells(r0, x0 + 3).Value = "POIDS MASSIF (kN)"
    ws.Cells(r0, x0 + 4).Value = "POIDS LONGRINE (kN)"
    For i = 2 To 4
        ws.Range(ws.Cells(r0, x0 + i), ws.Cells(r0 + 1, x0 + i)).Merge
    Next i
    ws.Cells(r0, x0 + 5).Value = "POIDS TOTAL ÉLÉMENTAIRE DDC"
    ws.Range(ws.Cells(r0, x0 + 5), ws.Cells(r0, x0 + 6)).Merge
    ' Charges de l'eau : colonnes créées seulement s'il y a une nappe phréatique
    xf = x0 + IIf(nappe, 9, 6)               ' dernière colonne avant la charge du talus
    ' Charge du talus (à saisir) en fin de tableau : effort horizontal H0 et moment M0 en
    ' tête de chaque micropieu, repris dans « Vérification structurelle »
    xt = xf + 1
    ws.Cells(r0, xt).Value = "CHARGE DU TALUS"
    ws.Range(ws.Cells(r0, xt), ws.Cells(r0, xt + 1)).Merge
    ws.Range(ws.Cells(r0 + 1, xt), ws.Cells(r0 + 1, xt + 1)).Value = Array("H0 (kN)", "M0 (kN.m)")
    If nappe Then
        ws.Cells(r0, x0 + 7).Value = "CHARGES DE L'EAU (kN)"
        ws.Range(ws.Cells(r0, x0 + 7), ws.Cells(r0, x0 + 9)).Merge
        ws.Range(ws.Cells(r0 + 1, x0 + 7), ws.Cells(r0 + 1, x0 + 9)).Value = Array("EAU EB", "EAU EH", "EAU EE")
    End If
    For j = 1 To n
        If nL > 0 Then ws.Cells(r0 + 1, 2 + j).Value = "Mur " & Format(j, "00")
        If nP > 0 Then ws.Cells(r0 + 1, 2 + nL + j).Value = nom & Format(j, "00")
    Next j
    ws.Cells(r0 + 1, x0).Value = "G (kN)"
    ws.Cells(r0 + 1, x0 + 1).Value = "Q (kN)"
    ws.Cells(r0 + 1, x0 + 5).Value = "G (kN)"
    ws.Cells(r0 + 1, x0 + 6).Value = "Q (kN)"

    ' Longueur Li de chaque mur (saisie), reprise sur la diagonale de la matrice
    If nL > 0 Then
        ws.Cells(r0 + 2, 2).Value = "Longueur du mur Li (m)"
        For j = 1 To n
            ws.Cells(r0 + 2, 2 + j).Value = li(j)
        Next j
    End If

    For i = 1 To m
        r = r0 + 2 + i
        ws.Cells(r, 2).Value = "MP" & i
        For j = 1 To nL
            If Not IsEmpty(mat(i, j)) Then
                ws.Cells(r, 2 + j).Value = mat(i, j)
            ElseIf i = j Then
                ws.Cells(r, 2 + j).Formula = "=N(" & ws.Cells(r0 + 2, 2 + j).Address & ")"
            Else
                ws.Cells(r, 2 + j).Value = 0
            End If
        Next j
        ' Identification des poteaux (0 / 1, saisie conservée ; par défaut 0, ou 1 sur la
        ' diagonale en semelle isolée : micropieu i sous le poteau Pot i)
        For j = 1 To nP
            With ws.Cells(r, 2 + nL + j)
                If Not IsEmpty(pot(i, j)) Then
                    .Value = pot(i, j)
                Else
                    .Value = IIf(isolee And i = j, 1, 0)
                End If
                .Validation.Delete
                .Validation.Add xlValidateList, xlValidAlertStop, , "0,1"
            End With
        Next j
        If nL > 0 Then ligne = ws.Range(ws.Cells(r, 3), ws.Cells(r, 2 + n)).Address(False, False)
        If nP > 0 Then ligneP = ws.Range(ws.Cells(r, 3 + nL), ws.Cells(r, 2 + nL + n)).Address(False, False)
        ' Charges du micropieu :
        '   semelle filante : longueurs x (G, Q des murs, kN/ml) ;
        '   mixte           : idem + identification x G poteaux des murs (kN) ;
        '   semelle isolée  : identification x (G, Q des poteaux Pot n, kN)
        If isolee Then
            fG = "MMULT(" & ligneP & "," & cG & ")": fQ = "MMULT(" & ligneP & "," & cQ & ")"
        Else
            fG = "MMULT(" & ligne & "," & cG & ")": fQ = "MMULT(" & ligne & "," & cQ & ")"
            If nP > 0 Then fG = fG & "+MMULT(" & ligneP & "," & cP & ")"
        End If
        ws.Cells(r, x0).Formula = "=IFERROR(" & fG & ","""")"
        ws.Cells(r, x0 + 1).Formula = "=IFERROR(" & fQ & ","""")"
        ' Structure GC : M (massif) ou L (longrine), jamais les deux
        With ws.Cells(r, x0 + 2)
            .Validation.Add xlValidateList, xlValidAlertStop, , "M,L"
            .Value = IIf(IsEmpty(gc(i)), "M", gc(i))
        End With
        ws.Cells(r, x0 + 3).Formula = "=IF(" & ws.Cells(r, x0 + 2).Address(False, False) & "=""M""," & _
            "$C$" & L + 2 & "*$D$" & L + 2 & "^2*$E$" & L + 2 & ",0)"
        ws.Cells(r, x0 + 4).Formula = "=IF(" & ws.Cells(r, x0 + 2).Address(False, False) & "=""L""," & _
            "$C$" & L + 3 & "*$D$" & L + 3 & "*$E$" & L + 3 & "*" & IIf(nL > 0, "SUM(" & ligne & ")", "0") & "/2,0)"
        ws.Cells(r, x0 + 5).Formula = "=N(" & ws.Cells(r, x0).Address(False, False) & ")+" & _
            ws.Cells(r, x0 + 3).Address(False, False) & "+" & ws.Cells(r, x0 + 4).Address(False, False)
        ws.Cells(r, x0 + 6).Formula = "=N(" & ws.Cells(r, x0 + 1).Address(False, False) & ")"
        ' Eau (basses eaux EB, hautes eaux EH, eaux exceptionnelles EE) : 0 sans nappe
        ' phréatique (« Données du projet », D34) ; à saisir sinon
        If nappe Then
            For j = 7 To 9
                ws.Cells(r, x0 + j).Value = IIf(IsEmpty(eau(i, j - 6)), 0, eau(i, j - 6))
            Next j
        End If
        ' Charge du talus H0, M0 (saisies conservées, 0 par défaut)
        ws.Cells(r, xt).Value = IIf(IsEmpty(talus(i, 1)), 0, talus(i, 1))
        ws.Cells(r, xt + 1).Value = IIf(IsEmpty(talus(i, 2)), 0, talus(i, 2))
    Next i

    ' Ligne TOTAL
    r = r0 + 3 + m
    ws.Cells(r, 2).Value = "TOTAL"
    For i = 0 To xf - x0
        If i <> 2 Then ws.Cells(r, x0 + i).Formula = "=SUM(" & _
            ws.Range(ws.Cells(r0 + 3, x0 + i), ws.Cells(r - 1, x0 + i)).Address(False, False) & ")"
    Next i
    For i = 0 To 1                           ' total de la charge du talus
        ws.Cells(r, xt + i).Formula = "=SUM(" & _
            ws.Range(ws.Cells(r0 + 3, xt + i), ws.Cells(r - 1, xt + i)).Address(False, False) & ")"
    Next i
    ws.Range(ws.Cells(r, 2), ws.Cells(r, xt + 1)).Font.Bold = True

    ' --- Mise en forme ---
    With ws.Range(ws.Cells(L + 5, 2), ws.Cells(L + 5 + n, 5))
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(r0, 2), ws.Cells(r, xt + 1))
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With Union(ws.Range("B" & L + 5 & ":E" & L + 5), _
            ws.Range(ws.Cells(r0, 2), ws.Cells(r0 + 1, xt + 1)))
        .Font.Bold = True
        .Interior.Color = COUL_ENTETE
        .WrapText = True
    End With
    If nL > 0 Then ws.Range(ws.Cells(r0 + 2, 3), ws.Cells(r - 1, 2 + nL)).Interior.Color = COUL_SAISIE
    If nP > 0 Then ws.Range(ws.Cells(r0 + 3, 3 + nL), ws.Cells(r - 1, 2 + nL + nP)).Interior.Color = COUL_SAISIE
    ws.Range(ws.Cells(r0 + 3, x0 + 2), ws.Cells(r - 1, x0 + 2)).Interior.Color = COUL_SAISIE
    ws.Range(ws.Cells(r0 + 3, x0), ws.Cells(r, x0 + 1)).Interior.Color = COUL_CALCUL
    ws.Range(ws.Cells(r0 + 3, x0 + 3), ws.Cells(r, x0 + 6)).Interior.Color = COUL_CALCUL
    If nappe Then                            ' eau à saisir (vert émeraude), total en bleu clair
        ws.Range(ws.Cells(r0 + 3, x0 + 7), ws.Cells(r - 1, x0 + 9)).Interior.Color = COUL_SAISIE
        ws.Range(ws.Cells(r, x0 + 7), ws.Cells(r, x0 + 9)).Interior.Color = COUL_CALCUL
    End If
    ws.Range(ws.Cells(r0 + 2, 3), ws.Cells(r, xt + 1)).NumberFormat = "0.00"
    If nP > 0 Then ws.Range(ws.Cells(r0 + 3, 3 + nL), ws.Cells(r - 1, 2 + nL + nP)).NumberFormat = "0"   ' identification
    ws.Range(ws.Cells(r0 + 3, xt), ws.Cells(r - 1, xt + 1)).Interior.Color = COUL_SAISIE   ' talus
    ws.Range(ws.Cells(r, xt), ws.Cells(r, xt + 1)).Interior.Color = COUL_CALCUL
    ws.Range("C" & L + 2 & ":E" & L + 3 & ",C" & L + 6 & ":E" & L + 5 + n).NumberFormat = "0.00"
    ws.Range("C" & L + 2 & ":E" & L + 3).HorizontalAlignment = xlCenter
    ws.Rows(r0).RowHeight = 45

    ' Tableau de vérification de la profondeur des micropieux de « Calcul et
    ' vérification » mis à jour (module CalculFondation, s'il est présent)
    On Error Resume Next
    Application.Run "ActualiserVerificationProfondeur"
    ' puis « Vérification structurelle » recréée avec le nouveau nombre de micropieux
    ' (module VerificationStructurelle ; pas pendant une création de « Calcul Portance »)
    If Not CBool(Application.Run("ReconstructionMatriceEnCours")) Then Application.Run "ActualiserVerificationStructurelle"

Fin:
    Application.EnableEvents = evenements
    Application.ScreenUpdating = True
End Sub
