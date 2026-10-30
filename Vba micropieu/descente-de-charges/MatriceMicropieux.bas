Attribute VB_Name = "MatriceMicropieux"
Option Explicit

' Module « Matrice des micropieux » : sous les tableaux de descente de charges des murs
' (à partir de la ligne LIGNE_MATRICE), répartition des charges des murs sur les
' micropieux.
'   L+1 à L+3 : poids volumique du béton armé, dimensions des massifs (longueur,
'               largeur, hauteur) et section des longrines (b, h), en m
'   L+5       : charges linéiques des murs, matrice n x 2 (n = nombre de murs, G1) :
'               Mur 01, Mur 02, ... avec G et Q (kN/ml) du récapitulatif de chaque mur
'   plus bas  : matrice m x n des longueurs de murs reprises par chaque micropieu
'               (m = nombre de micropieux, K1 ; lignes MP1, MP2, ..., colonnes Mur 01,
'               Mur 02, ...) : longueur Li du mur sur la diagonale, 0 ailleurs, les
'               valeurs peuvent être modifiées ; puis, à droite :
'               CHARGES DES MURS G, Q (kN) = matrice des longueurs x matrice n x 2 ;
'               STRUCTURE GC : M (massif) ou L (longrine) ;
'               POIDS MASSIF (kN) = poids volumique x volume du massif (si M, sinon 0) ;
'               POIDS LONGRINE (kN) = poids volumique x section x longueur, longueur =
'               somme des longueurs de murs de la ligne du micropieu / 2 (si L, sinon 0) ;
'               POIDS TOTAL ÉLÉMENTAIRE DDC : G = G des murs + massif + longrine, Q.
' Les saisies (dimensions, longueurs Li, longueurs modifiées de la matrice, choix M / L)
' sont conservées quand le tableau est recréé (nombre de murs ou de micropieux changé).
' Appelé par DupliquerMurs (nombre de murs, G1) et quand K1 change (CodeFeuille.txt).

Public Const LIGNE_MATRICE As Long = 80             ' titre du tableau
Public Const CELLULE_NB_MICROPIEUX As String = "K1"

' Thème : saisies en vert, valeurs calculées en jaune, en-têtes en vert clair
Private Const VERT_SAISIE As Long = 13561798   ' RGB(198, 239, 206)
Private Const JAUNE_CALCUL As Long = 13434879  ' RGB(255, 255, 204)
Private Const VERT_ENTETE As Long = 14348258   ' RGB(226, 239, 218)

Public Sub ConstruireMatrice(Optional ws As Worksheet)
    Dim n As Long, m As Long, i As Long, j As Long, L As Long, r0 As Long, rG As Long
    Dim nOld As Long, mOld As Long, r0Old As Long, x0 As Long, r As Long, derniere As Long
    Dim li() As Variant, mat() As Variant, gc() As Variant, dims As Variant
    Dim ligne As String, cG As String, cQ As String, c As Range, evenements As Boolean

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
    dims = ws.Range("C" & L + 1 & ":E" & L + 3).Value
    Do While Left(CStr(ws.Range("B" & L + 6 + nOld).Value), 4) = "Mur "
        nOld = nOld + 1
    Loop
    r0Old = L + 8 + nOld
    If CStr(ws.Range("B" & r0Old).Value) = "N° MICROPIEU" Then
        Do While Left(CStr(ws.Range("B" & r0Old + 3 + mOld).Value), 2) = "MP"
            mOld = mOld + 1
        Loop
        For j = 1 To Application.Min(n, nOld)
            li(j) = ws.Cells(r0Old + 2, 2 + j).Value
        Next j
        For i = 1 To Application.Min(m, mOld)
            For j = 1 To Application.Min(n, nOld)
                Set c = ws.Cells(r0Old + 2 + i, 2 + j)
                If Not c.HasFormula Then mat(i, j) = c.Value
            Next j
            gc(i) = ws.Cells(r0Old + 2 + i, 5 + nOld).Value
        Next i
    End If

    ' --- Ancien tableau effacé (lignes à partir de L) ---
    derniere = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If derniere < L + 200 Then derniere = L + 200
    ws.Range(ws.Rows(L), ws.Rows(derniere)).Clear

    ' --- Ligne du récapitulatif G / Q des murs (sous le tableau de chaque mur) ---
    For r = 5 To L - 1
        If CStr(ws.Range("B" & r).Value) = "G (kN/ml)" Then rG = r
    Next r

    ' --- Titre et données d'entrée ---
    With ws.Range("B" & L)
        .Value = "RÉPARTITION DES CHARGES DES MURS SUR LES MICROPIEUX"
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.Range("B" & L + 1).Value = "Poids volumique du béton armé (kN/m3)"
    ws.Range("B" & L + 2).Value = "Massif : longueur, largeur, hauteur (m)"
    ws.Range("B" & L + 3).Value = "Longrine : largeur b, hauteur h (m)"
    ws.Range("C" & L + 1).Value = IIf(IsEmpty(dims(1, 1)), 25, dims(1, 1))
    ws.Range("C" & L + 2 & ":E" & L + 2).Value = Array(dims(2, 1), dims(2, 2), dims(2, 3))
    ws.Range("C" & L + 3 & ":D" & L + 3).Value = Array(dims(3, 1), dims(3, 2))
    ws.Range("C" & L + 1 & ",C" & L + 2 & ":E" & L + 2 & ",C" & L + 3 & ":D" & L + 3).Interior.Color = _
        VERT_SAISIE
    ws.Range("B" & L + 1 & ":E" & L + 3).Borders.LineStyle = xlContinuous

    ' --- Matrice n x 2 : G et Q (kN/ml) de chaque mur (récapitulatif du mur k, colonne
    ' C du tableau d'origine + (k - 1) x PAS_MUR) ---
    ws.Range("B" & L + 5 & ":D" & L + 5).Value = Array("MUR", "G (kN/ml)", "Q (kN/ml)")
    For j = 1 To n
        r = L + 5 + j
        ws.Range("B" & r).Value = "Mur " & Format(j, "00")
        If rG > 0 Then
            ws.Range("C" & r).Formula = "=N(" & ws.Cells(rG, 3 + (j - 1) * PAS_MUR).Address & ")"
            ws.Range("D" & r).Formula = "=N(" & ws.Cells(rG + 1, 3 + (j - 1) * PAS_MUR).Address & ")"
        End If
    Next j
    ws.Range("C" & L + 6 & ":D" & L + 5 + n).Interior.Color = JAUNE_CALCUL
    cG = "$C$" & L + 6 & ":$C$" & L + 5 + n
    cQ = "$D$" & L + 6 & ":$D$" & L + 5 + n

    ' --- Matrice m x n des longueurs et résultats ---
    r0 = L + 8 + n                            ' 1re ligne d'en-tête
    x0 = 3 + n                                ' 1re colonne après la matrice
    ws.Cells(r0, 2).Value = "N° MICROPIEU"
    ws.Range(ws.Cells(r0, 2), ws.Cells(r0 + 1, 2)).Merge
    ws.Cells(r0, 3).Value = "LONGUEURS DES MURS REPRISES PAR LE MICROPIEU (m)"
    If n > 1 Then ws.Range(ws.Cells(r0, 3), ws.Cells(r0, 2 + n)).Merge
    ws.Cells(r0, x0).Value = "CHARGES DES MURS"
    ws.Range(ws.Cells(r0, x0), ws.Cells(r0, x0 + 1)).Merge
    ws.Cells(r0, x0 + 2).Value = "STRUCTURE GC"
    ws.Cells(r0, x0 + 3).Value = "POIDS MASSIF (kN)"
    ws.Cells(r0, x0 + 4).Value = "POIDS LONGRINE (kN)"
    For i = 2 To 4
        ws.Range(ws.Cells(r0, x0 + i), ws.Cells(r0 + 1, x0 + i)).Merge
    Next i
    ws.Cells(r0, x0 + 5).Value = "POIDS TOTAL ÉLÉMENTAIRE DDC"
    ws.Range(ws.Cells(r0, x0 + 5), ws.Cells(r0, x0 + 6)).Merge
    For j = 1 To n
        ws.Cells(r0 + 1, 2 + j).Value = "Mur " & Format(j, "00")
    Next j
    ws.Cells(r0 + 1, x0).Value = "G (kN)"
    ws.Cells(r0 + 1, x0 + 1).Value = "Q (kN)"
    ws.Cells(r0 + 1, x0 + 5).Value = "G (kN)"
    ws.Cells(r0 + 1, x0 + 6).Value = "Q (kN)"

    ' Longueur Li de chaque mur (saisie), reprise sur la diagonale de la matrice
    ws.Cells(r0 + 2, 2).Value = "Longueur du mur Li (m)"
    For j = 1 To n
        ws.Cells(r0 + 2, 2 + j).Value = li(j)
    Next j

    For i = 1 To m
        r = r0 + 2 + i
        ws.Cells(r, 2).Value = "MP" & i
        For j = 1 To n
            If Not IsEmpty(mat(i, j)) Then
                ws.Cells(r, 2 + j).Value = mat(i, j)
            ElseIf i = j Then
                ws.Cells(r, 2 + j).Formula = "=N(" & ws.Cells(r0 + 2, 2 + j).Address & ")"
            Else
                ws.Cells(r, 2 + j).Value = 0
            End If
        Next j
        ligne = ws.Range(ws.Cells(r, 3), ws.Cells(r, 2 + n)).Address(False, False)
        ' Charges des murs = (ligne de la matrice des longueurs) x (matrice n x 2)
        ws.Cells(r, x0).Formula = "=IFERROR(MMULT(" & ligne & "," & cG & "),"""")"
        ws.Cells(r, x0 + 1).Formula = "=IFERROR(MMULT(" & ligne & "," & cQ & "),"""")"
        ' Structure GC : M (massif) ou L (longrine), jamais les deux
        With ws.Cells(r, x0 + 2)
            .Validation.Add xlValidateList, xlValidAlertStop, , "M,L"
            .Value = IIf(IsEmpty(gc(i)), "M", gc(i))
        End With
        ws.Cells(r, x0 + 3).Formula = "=IF(" & ws.Cells(r, x0 + 2).Address(False, False) & "=""M""," & _
            "$C$" & L + 1 & "*$C$" & L + 2 & "*$D$" & L + 2 & "*$E$" & L + 2 & ",0)"
        ws.Cells(r, x0 + 4).Formula = "=IF(" & ws.Cells(r, x0 + 2).Address(False, False) & "=""L""," & _
            "$C$" & L + 1 & "*$C$" & L + 3 & "*$D$" & L + 3 & "*SUM(" & ligne & ")/2,0)"
        ws.Cells(r, x0 + 5).Formula = "=N(" & ws.Cells(r, x0).Address(False, False) & ")+" & _
            ws.Cells(r, x0 + 3).Address(False, False) & "+" & ws.Cells(r, x0 + 4).Address(False, False)
        ws.Cells(r, x0 + 6).Formula = "=N(" & ws.Cells(r, x0 + 1).Address(False, False) & ")"
    Next i

    ' Ligne TOTAL
    r = r0 + 3 + m
    ws.Cells(r, 2).Value = "TOTAL"
    For i = 0 To 6
        If i <> 2 Then ws.Cells(r, x0 + i).Formula = "=SUM(" & _
            ws.Range(ws.Cells(r0 + 3, x0 + i), ws.Cells(r - 1, x0 + i)).Address(False, False) & ")"
    Next i
    ws.Range(ws.Cells(r, 2), ws.Cells(r, x0 + 6)).Font.Bold = True

    ' --- Mise en forme ---
    With ws.Range(ws.Cells(L + 5, 2), ws.Cells(L + 5 + n, 4))
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(r0, 2), ws.Cells(r, x0 + 6))
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With Union(ws.Range("B" & L + 5 & ":D" & L + 5), _
            ws.Range(ws.Cells(r0, 2), ws.Cells(r0 + 1, x0 + 6)))
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .WrapText = True
    End With
    ws.Range(ws.Cells(r0 + 2, 3), ws.Cells(r - 1, 2 + n)).Interior.Color = VERT_SAISIE
    ws.Range(ws.Cells(r0 + 3, x0 + 2), ws.Cells(r - 1, x0 + 2)).Interior.Color = VERT_SAISIE
    ws.Range(ws.Cells(r0 + 3, x0), ws.Cells(r, x0 + 1)).Interior.Color = JAUNE_CALCUL
    ws.Range(ws.Cells(r0 + 3, x0 + 3), ws.Cells(r, x0 + 6)).Interior.Color = JAUNE_CALCUL
    ws.Range(ws.Cells(r0 + 2, 3), ws.Cells(r, x0 + 6)).NumberFormat = "0.00"
    ws.Range("C" & L + 1 & ":E" & L + 3 & ",C" & L + 6 & ":D" & L + 5 + n).NumberFormat = "0.00"
    ws.Range("C" & L + 1 & ":E" & L + 3).HorizontalAlignment = xlCenter
    ws.Rows(r0).RowHeight = 45

Fin:
    Application.EnableEvents = evenements
    Application.ScreenUpdating = True
End Sub
