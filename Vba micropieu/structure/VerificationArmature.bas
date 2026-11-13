Attribute VB_Name = "VerificationArmature"
Option Explicit

' Module « Vérification de l'armature » : ajoute, sous les graphiques Y, T, M de la
' feuille « Vérification structurelle » (module VerificationStructurelle), le tableau
' de l'armature du micropieu le plus long. Toutes les valeurs sont des formules.
'   Caractéristiques du tube étudié : type et désignation (« Données du projet »), d, e,
'     di, rayon de corrosion retenu rc, de = d - 2 rc, t = (de - di) / 2, As, I,
'     Wel = 2 I / de, Wpl = (de³ - di³) / 6, E, fy ;
'   sollicitations agissant sur le tube, pour chaque combinaison (ELS CARA, ELS QP,
'     ELU FOND, ELU ACC) et maximum des combinaisons : effort normal Nc, effort
'     horizontal en tête T0 = H, moment en tête M0 (tableau des efforts en tête, issus
'     de la descente de charges).
' Appelée à la fin de CreerVerificationStructurelle ; peut être relancée seule.

Private Const NOM_FEUILLE As String = "Vérification structurelle"
Private Const DONNEES As String = "'Données du projet'!"
Private Const TITRE As String = "VÉRIFICATION DE L'ARMATURE - CARACTÉRISTIQUES DU TUBE ET SOLLICITATIONS"

Private Const JAUNE_CALCUL As Long = 13434879  ' RGB(255, 255, 204)
Private Const VERT_ENTETE As Long = 14348258   ' RGB(226, 239, 218)
Private Const VERT_FONCE As Long = 2315831     ' RGB(55, 86, 35)

Public Sub VerificationArmature()
    Dim ws As Worksheet, co As ChartObject, v As Variant, r As Long, r0 As Long, k As Long
    Dim bas As Double, hP As Long, rN As Long, nP As Long, cl As String, c As String
    Dim rde As String, rdi As String, rAs As String, rI As String, fy As String, rrc As String
    Dim lig As Object, etape As String, zone As String, j As Long

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo Erreur
    If ws Is Nothing Then
        MsgBox "Créez d'abord la feuille « " & NOM_FEUILLE & " » (CreerVerificationStructurelle).", vbExclamation
        Exit Sub
    End If

    ' --- Cellules de la feuille « Vérification structurelle » (repérées par intitulé) ---
    etape = "lecture de la feuille"
    rde = Valeur(ws, "Diamètre après corrosion d - 2 rc")
    rdi = Valeur(ws, "Diamètre intérieur di (tube)")
    rAs = Valeur(ws, "Section après corrosion As")
    rI = Valeur(ws, "Inertie de la section après corrosion I")
    rrc = Valeur(ws, "Rayon de corrosion retenu rc (0 si protection efficace)")
    v = Application.Match("Nc", ws.Range("C1:C500"), 0)       ' efforts en tête
    If rde = "" Or rAs = "" Or rI = "" Or IsError(v) Then
        MsgBox "Tableaux de la feuille « " & NOM_FEUILLE & " » introuvables : relancez " & _
            "CreerVerificationStructurelle.", vbExclamation
        Exit Sub
    End If
    rN = CLng(v)
    v = Application.Match("z (m)", ws.Range("A1:A1000"), 0)    ' profil Y, T, M
    If IsError(v) Then
        MsgBox "Profil Y, T, M introuvable : relancez CreerVerificationStructurelle.", vbExclamation
        Exit Sub
    End If
    hP = CLng(v)
    Do While Not IsEmpty(ws.Cells(hP + 2 + nP, 1).Value) Or ws.Cells(hP + 2 + nP, 1).HasFormula
        nP = nP + 1
        If nP > 5000 Then Exit Do
    Loop
    If nP < 1 Then nP = 1
    fy = "N(" & DONNEES & "$D$" & LIGNE_ACIER + 7 & ")"

    ' --- Position : sous les graphiques (ou sous le profil) ; ancien tableau effacé ---
    etape = "position du tableau"
    v = Application.Match(TITRE, ws.Range("A1:A5000"), 0)
    If Not IsError(v) Then ws.Range(ws.Rows(CLng(v)), ws.Rows(CLng(v) + 60)).Clear
    bas = ws.Cells(hP + 2 + nP, 1).Top
    For Each co In ws.ChartObjects
        If co.Top + co.Height > bas Then bas = co.Top + co.Height
    Next co
    r0 = hP + 2 + nP
    Do While ws.Cells(r0, 1).Top < bas + 10
        r0 = r0 + 1
    Loop
    r0 = r0 + 1

    With ws.Range("A" & r0)
        .Value = TITRE
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = VERT_FONCE
    End With
    r = r0 + 1
    ws.Range("B" & r & ":I" & r).Value = Array("Grandeur", "", "Symbole / formule", "ELS CARA", "ELS QP", _
        "ELU FOND", "ELU ACC", "Unité")
    ws.Range("B" & r & ":C" & r).Merge
    Set lig = CreateObject("Scripting.Dictionary")

    ' --- Caractéristiques du tube étudié (valeur en E, commune aux combinaisons) ---
    etape = "caractéristiques du tube"
    Section ws, r, lig, "typ", "Type d'armature", "« Données du projet »", "=" & DONNEES & "$D$" & LIGNE_ACIER & "&""""", ""
    Section ws, r, lig, "des", "Désignation de l'armature", "« Données du projet »", "=" & DONNEES & "$D$" & _
        LIGNE_ACIER + 1 & "&""""", ""
    Section ws, r, lig, "d", "Diamètre extérieur nominal", "d", "=N(" & DONNEES & "$D$" & LIGNE_ACIER + 2 & ")", "mm"
    Section ws, r, lig, "e", "Épaisseur nominale", "e", "=N(" & DONNEES & "$D$" & LIGNE_ACIER + 3 & ")", "mm"
    Section ws, r, lig, "di", "Diamètre intérieur", "di = d - 2 e (0 pour une barre)", "=" & IIf(rdi = "", "0", rdi), "mm"
    Section ws, r, lig, "rc", "Rayon de corrosion retenu", "rc", "=" & IIf(rrc = "", "0", rrc), "mm"
    Section ws, r, lig, "de", "Diamètre extérieur après corrosion", "de = d - 2 rc", "=" & rde, "mm"
    Section ws, r, lig, "t", "Épaisseur après corrosion", "t = (de - di) / 2", "=(" & lig("de") & "-" & _
        lig("di") & ")/2", "mm"
    Section ws, r, lig, "As", "Section d'acier après corrosion", "As = " & ChrW(960) & " (de² - di²) / 4", "=" & rAs, "mm²"
    Section ws, r, lig, "I", "Inertie après corrosion", "I = " & ChrW(960) & " (de^4 - di^4) / 64", "=" & rI, "mm4"
    Section ws, r, lig, "Wel", "Module de flexion élastique", "Wel = 2 I / de", "=IF(N(" & lig("de") & _
        ")>0,2*" & lig("I") & "/" & lig("de") & ","""")", "mm³"
    Section ws, r, lig, "Wpl", "Module de flexion plastique", "Wpl = (de³ - di³) / 6", "=(" & lig("de") & _
        "^3-" & lig("di") & "^3)/6", "mm³"
    Section ws, r, lig, "E", "Module de Young", "E (« Données du projet »)", "=N(" & DONNEES & "$D$" & _
        LIGNE_ACIER + 6 & ")", "MPa"
    Section ws, r, lig, "fy", "Limite d'élasticité de l'acier", "fy = fyk (« Données du projet »)", "=" & fy, "MPa"

    ' --- Sollicitations agissant sur le tube : efforts en tête de la descente de charges
    ' pour chaque combinaison (tableau des efforts en tête) et valeur maximale ---
    etape = "sollicitations"
    Ligne4 ws, r, lig, "N", "Effort normal de compression (descente de charges)", "Nc = Vmax", "kN", _
        Array("=ABS(N(E" & rN & "))", "=ABS(N(F" & rN & "))", "=ABS(N(G" & rN & "))", "=ABS(N(H" & rN & "))")
    Ligne4 ws, r, lig, "T0", "Effort horizontal en tête (inclinaison)", "T0 = H = Nc x sin " & ChrW(945), "kN", _
        Array("=ABS(N(E" & rN + 2 & "))", "=ABS(N(F" & rN + 2 & "))", "=ABS(N(G" & rN + 2 & "))", _
        "=ABS(N(H" & rN + 2 & "))")
    Ligne4 ws, r, lig, "M0", "Moment en tête (selon la condition en tête)", "M0", "kN.m", _
        Array("=ABS(N(E" & rN + 3 & "))", "=ABS(N(F" & rN + 3 & "))", "=ABS(N(G" & rN + 3 & "))", _
        "=ABS(N(H" & rN + 3 & "))")
    ' Valeur maximale des quatre combinaisons (colonne J)
    For j = 0 To 2
        ws.Range("J" & r - 2 + j).Formula = "=MAX(E" & r - 2 + j & ":H" & r - 2 + j & ")"
    Next j
    ws.Range("J" & r0 + 1).Value = "Maximum"
    ws.Range("J" & r - 2 & ":J" & r).NumberFormat = "0.00"
    ws.Range("J" & r - 2 & ":J" & r).Font.Bold = True

    ' --- Mise en forme ---
    etape = "mise en forme"
    With ws.Range("B" & r0 + 1 & ":J" & r)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("E" & r0 + 1 & ":J" & r).HorizontalAlignment = xlCenter
    With ws.Range("B" & r0 + 1 & ":J" & r0 + 1)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("E" & r0 + 2 & ":H" & r & ",J" & r - 2 & ":J" & r).Interior.Color = JAUNE_CALCUL
    Exit Sub
Erreur:
    MsgBox "Erreur pendant la vérification de l'armature, partie « " & etape & " » : " & Err.Description & _
        " (n° " & Err.Number & ")", vbExclamation
End Sub

' Adresse ($D$n) de la valeur de la ligne d'intitulé lib (colonne B) ; vide si absente
Private Function Valeur(ws As Worksheet, lib As String) As String
    Dim v As Variant
    v = Application.Match(lib, ws.Range("B1:B500"), 0)
    If Not IsError(v) Then Valeur = "$D$" & CLng(v)
End Function

' Ligne de caractéristique de la section : intitulé (B:C), symbole / formule (D),
' valeur (E, fusionnée sur E:H), unité (I) ; adresse de la valeur mémorisée sous cle
Private Sub Section(ws As Worksheet, r As Long, lig As Object, cle As String, intitule As String, _
        symbole As String, formule As String, unite As String)
    r = r + 1
    ws.Range("B" & r).Value = intitule
    ws.Range("B" & r & ":C" & r).Merge
    ws.Range("D" & r).Value = symbole
    ws.Range("E" & r).Formula = formule
    ws.Range("E" & r & ":H" & r).Merge
    ws.Range("I" & r).Value = unite
    If cle = "I" Or cle = "Wel" Or cle = "Wpl" Or cle = "W" Then
        ws.Range("E" & r).NumberFormat = "0.000E+00"
    ElseIf cle = "cl" Then
        ws.Range("E" & r).NumberFormat = "0"
    Else
        ws.Range("E" & r).NumberFormat = "0.00"
    End If
    lig(cle) = "$E$" & r
End Sub

' Ligne par combinaison : formules des colonnes E à H (tableau de 4 formules) ; la 1re
' colonne de la ligne est mémorisée sous cle (référence relative de colonne « E$n »)
Private Sub Ligne4(ws As Worksheet, r As Long, lig As Object, cle As String, intitule As String, _
        symbole As String, unite As String, f As Variant)
    Dim k As Long
    r = r + 1
    ws.Range("B" & r).Value = intitule
    ws.Range("B" & r & ":C" & r).Merge
    ws.Range("D" & r).Value = symbole
    ws.Range("I" & r).Value = unite
    lig(cle) = "E" & r
    For k = 0 To 3
        ws.Cells(r, 5 + k).Formula = Remplacer(CStr(f(k)), lig, r, k)
    Next k
    ws.Range("E" & r & ":H" & r).NumberFormat = "0.00"
End Sub

' Quatre fois la même formule (une par combinaison) : #cle désigne la cellule de la même
' combinaison de la ligne cle (remplacée dans Ligne4)
Private Function Formules(f As String) As Variant
    Formules = Array(f, f, f, f)
End Function

' Remplace les #cle d'une formule par la cellule de la combinaison k (colonne E + k)
' de la ligne mémorisée sous cle
Private Function Remplacer(f As String, lig As Object, r As Long, k As Long) As String
    Dim cles As Variant, i As Long, cle As String, cl As String
    cl = Chr(Asc("E") + k)
    cles = lig.Keys
    ' clés les plus longues d'abord (évite de remplacer #s dans #sN, #svm...)
    Dim n As Long, j As Long, tmp As String
    n = UBound(cles)
    For i = 0 To n - 1
        For j = i + 1 To n
            If Len(cles(j)) > Len(cles(i)) Then tmp = cles(i): cles(i) = cles(j): cles(j) = tmp
        Next j
    Next i
    For i = 0 To n
        cle = cles(i)
        If Left(lig(cle), 1) = "E" Then        ' ligne par combinaison : même colonne
            f = Replace(f, "#" & cle, cl & "$" & Mid(lig(cle), 2))
        End If
    Next i
    Remplacer = f
End Function

' Formule « max |valeur| » d'une colonne du profil (lignes r1 à r1 + n - 1)
Private Function MaxAbs(ws As Worksheet, r1 As Long, n As Long, col As Long) As String
    Dim z As String
    z = ws.Range(ws.Cells(r1, col), ws.Cells(r1 + n - 1, col)).Address
    MaxAbs = "=IF(COUNT(" & z & ")=0,"""",MAX(MAX(" & z & "),-MIN(" & z & ")))"
End Function
