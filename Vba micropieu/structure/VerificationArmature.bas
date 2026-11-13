Attribute VB_Name = "VerificationArmature"
Option Explicit

' Module « Vérification de l'armature » : ajoute, sous les graphiques Y, T, M de la
' feuille « Vérification structurelle » (module VerificationStructurelle), le tableau
' de la vérification métallique du tube (ou de la barre) du micropieu le plus long,
' après corrosion. Toutes les valeurs sont des formules : elles suivent la feuille.
'   Section (NF EN 1993-1-1) : diamètre extérieur après corrosion de, diamètre intérieur
'     di, épaisseur t = (de - di) / 2, As, I, Wel = 2 I / de, Wpl = (de³ - di³) / 6,
'     epsilon = racine(235 / fy), d / t et classe du tube (tableau 5.2 : classe 1 si
'     d/t <= 50 eps², 2 si <= 70 eps², 3 si <= 90 eps², 4 sinon), module retenu
'     (Wpl en classe 1 ou 2, Wel en classe 3), aire de cisaillement Av = 2 As / pi ;
'   par combinaison (ELS CARA, ELS QP, ELU FOND, ELU ACC) : coefficient partiel
'     (1 aux ELS ; gamma s ELU fondamental et accidentel de « Données du projet »),
'     fyd = fy / gamma ; efforts N = Nc (efforts en tête), Mmax = max |M(z)| et
'     Tmax = max |T(z)| (profil le long du micropieu) ;
'     contraintes : sigma N = N / As, sigma M = Mmax / Wel, sigma = sigma N + sigma M,
'     tau = 2 Tmax / As (tube mince), sigma VM = racine(sigma² + 3 tau²), taux
'     sigma VM / fyd ;
'     résistances : Npl,Rd = As fyd, Mpl,Rd = W fyd, Vpl,Rd = Av fyd / racine(3),
'     n = N / Npl,Rd, MN,Rd = Mpl,Rd (1 - n^1,7) (tube circulaire, NF EN 1993-1-1,
'     6.2.9, à vérifier selon l'annexe nationale), vérification N <= Npl,Rd,
'     Mmax <= MN,Rd et Tmax <= Vpl,Rd.
' Appelée à la fin de CreerVerificationStructurelle ; peut être relancée seule.

Private Const NOM_FEUILLE As String = "Vérification structurelle"
Private Const DONNEES As String = "'Données du projet'!"
Private Const TITRE As String = "VÉRIFICATION DE L'ARMATURE (VÉRIFICATION MÉTALLIQUE DU TUBE APRÈS CORROSION)"

Private Const JAUNE_CALCUL As Long = 13434879  ' RGB(255, 255, 204)
Private Const VERT_ENTETE As Long = 14348258   ' RGB(226, 239, 218)
Private Const VERT_FONCE As Long = 2315831     ' RGB(55, 86, 35)

Public Sub VerificationArmature()
    Dim ws As Worksheet, co As ChartObject, v As Variant, r As Long, r0 As Long, k As Long
    Dim bas As Double, hP As Long, rN As Long, nP As Long, cl As String, c As String
    Dim rde As String, rdi As String, rAs As String, rI As String, fy As String
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

    ' --- Caractéristiques de la section (valeur en E, commune aux combinaisons) ---
    etape = "caractéristiques de la section"
    Section ws, r, lig, "fy", "Limite d'élasticité de l'acier", "fy = fyk (« Données du projet »)", "=" & fy, "MPa"
    Section ws, r, lig, "de", "Diamètre extérieur après corrosion", "de = d - 2 rc", "=" & rde, "mm"
    Section ws, r, lig, "di", "Diamètre intérieur", "di (0 pour une barre)", "=" & IIf(rdi = "", "0", rdi), "mm"
    Section ws, r, lig, "t", "Épaisseur après corrosion", "t = (de - di) / 2", "=(" & lig("de") & "-" & _
        lig("di") & ")/2", "mm"
    Section ws, r, lig, "As", "Section après corrosion", "As", "=" & rAs, "mm²"
    Section ws, r, lig, "I", "Inertie après corrosion", "I", "=" & rI, "mm4"
    Section ws, r, lig, "Wel", "Module de flexion élastique", "Wel = 2 I / de", "=IF(N(" & lig("de") & _
        ")>0,2*" & lig("I") & "/" & lig("de") & ","""")", "mm³"
    Section ws, r, lig, "Wpl", "Module de flexion plastique", "Wpl = (de³ - di³) / 6", "=(" & lig("de") & _
        "^3-" & lig("di") & "^3)/6", "mm³"
    Section ws, r, lig, "eps", "Coefficient " & ChrW(949), ChrW(949) & " = racine(235 / fy)", "=IF(N(" & _
        lig("fy") & ")>0,SQRT(235/" & lig("fy") & "),"""")", ""
    Section ws, r, lig, "dt", "Élancement de la paroi", "d / t = de / t", "=IF(N(" & lig("t") & ")>0," & _
        lig("de") & "/" & lig("t") & ","""")", ""
    Section ws, r, lig, "cl", "Classe de la section (NF EN 1993-1-1, tableau 5.2, tube)", _
        "1 : d/t <= 50 " & ChrW(949) & "² ; 2 : <= 70 " & ChrW(949) & "² ; 3 : <= 90 " & ChrW(949) & "² ; 4 sinon", _
        "=IF(OR(" & lig("dt") & "=""""," & lig("eps") & "=""""),""-"",IF(" & lig("dt") & "<=50*" & lig("eps") & _
        "^2,1,IF(" & lig("dt") & "<=70*" & lig("eps") & "^2,2,IF(" & lig("dt") & "<=90*" & lig("eps") & "^2,3,4))))", ""
    Section ws, r, lig, "W", "Module de flexion retenu", "W = Wpl (classe 1 ou 2), Wel (classe 3)", _
        "=IF(N(" & lig("cl") & ")<=2," & lig("Wpl") & "," & lig("Wel") & ")", "mm³"
    Section ws, r, lig, "Av", "Aire de cisaillement", "Av = 2 As / " & ChrW(960), "=2*" & lig("As") & "/PI()", "mm²"

    ' --- Par combinaison : efforts, contraintes, résistances, vérifications ---
    etape = "contraintes et résistances"
    Ligne4 ws, r, lig, "g", "Coefficient partiel de l'acier", ChrW(947) & " (ELS : 1 ; ELU : " & ChrW(947) & _
        "s de « Données du projet »)", "", Array("=1", "=1", "=N(" & DONNEES & "$D$" & LIGNE_ACIER + 8 & ")", _
        "=N(" & DONNEES & "$D$" & LIGNE_ACIER + 9 & ")")
    Ligne4 ws, r, lig, "fyd", "Limite d'élasticité de calcul", "fyd = fy / " & ChrW(947), "MPa", _
        Formules("=IF(N(#g)>0," & lig("fy") & "/#g,"""")")
    Ligne4 ws, r, lig, "N", "Effort normal (descente de charges)", "N = Nc (efforts en tête)", "kN", _
        Array("=ABS(N(E" & rN & "))", "=ABS(N(F" & rN & "))", "=ABS(N(G" & rN & "))", "=ABS(N(H" & rN & "))")
    ' Mmax et Tmax du profil : colonnes M (7 + 3 k) et T (6 + 3 k) de la combinaison k
    Ligne4 ws, r, lig, "M", "Moment maximal le long du micropieu", "Mmax = max |M(z)| (profil)", "kN.m", _
        Array(MaxAbs(ws, hP + 2, nP, 7), MaxAbs(ws, hP + 2, nP, 10), MaxAbs(ws, hP + 2, nP, 13), _
        MaxAbs(ws, hP + 2, nP, 16))
    Ligne4 ws, r, lig, "T", "Effort tranchant maximal le long du micropieu", "Tmax = max |T(z)| (profil)", "kN", _
        Array(MaxAbs(ws, hP + 2, nP, 6), MaxAbs(ws, hP + 2, nP, 9), MaxAbs(ws, hP + 2, nP, 12), _
        MaxAbs(ws, hP + 2, nP, 15))
    Ligne4 ws, r, lig, "sN", "Contrainte normale due à N", ChrW(963) & "N = N / As", "MPa", _
        Formules("=IF(N(" & lig("As") & ")>0,1000*#N/" & lig("As") & ","""")")
    Ligne4 ws, r, lig, "sM", "Contrainte normale due à M", ChrW(963) & "M = Mmax / Wel", "MPa", _
        Formules("=IF(N(" & lig("Wel") & ")>0,1000000*#M/" & lig("Wel") & ","""")")
    Ligne4 ws, r, lig, "s", "Contrainte normale maximale", ChrW(963) & " = " & ChrW(963) & "N + " & ChrW(963) & "M", _
        "MPa", Formules("=IF(OR(#sN="""",#sM=""""),"""",#sN+#sM)")
    Ligne4 ws, r, lig, "tau", "Contrainte de cisaillement maximale (tube mince)", ChrW(964) & " = 2 Tmax / As", _
        "MPa", Formules("=IF(N(" & lig("As") & ")>0,2000*#T/" & lig("As") & ","""")")
    Ligne4 ws, r, lig, "svm", "Contrainte équivalente de von Mises", ChrW(963) & "VM = racine(" & ChrW(963) & _
        "² + 3 " & ChrW(964) & "²)", "MPa", Formules("=IF(OR(#s="""",#tau=""""),"""",SQRT(#s^2+3*#tau^2))")
    Ligne4 ws, r, lig, "tx", "Taux de travail en contraintes", ChrW(963) & "VM / fyd", "", _
        Formules("=IF(OR(#svm="""",N(#fyd)=0),"""",#svm/#fyd)")
    Ligne4 ws, r, lig, "vs", "Vérification des contraintes", ChrW(963) & "VM <= fyd", "", _
        Formules("=IF(#tx="""",""-"",IF(#tx<=1,""Vérifiée"",""Non vérifiée""))")
    Ligne4 ws, r, lig, "Npl", "Résistance plastique à l'effort normal", "Npl,Rd = As fyd", "kN", _
        Formules("=IF(#fyd="""","""",#fyd*" & lig("As") & "/1000)")
    Ligne4 ws, r, lig, "Mpl", "Moment résistant", "Mpl,Rd = W fyd", "kN.m", _
        Formules("=IF(#fyd="""","""",#fyd*" & lig("W") & "/1000000)")
    Ligne4 ws, r, lig, "Vpl", "Résistance plastique au cisaillement", "Vpl,Rd = Av fyd / racine(3)", "kN", _
        Formules("=IF(#fyd="""","""",#fyd*" & lig("Av") & "/SQRT(3)/1000)")
    Ligne4 ws, r, lig, "n", "Taux d'effort normal", "n = N / Npl,Rd", "", _
        Formules("=IF(N(#Npl)>0,#N/#Npl,"""")")
    Ligne4 ws, r, lig, "MN", "Moment résistant réduit par l'effort normal (tube)", _
        "MN,Rd = Mpl,Rd (1 - n^1,7) (6.2.9)", "kN.m", Formules("=IF(#n="""","""",IF(#n>=1,0,#Mpl*(1-#n^1.7)))")
    Ligne4 ws, r, lig, "tM", "Taux de flexion composée", "Mmax / MN,Rd", "", _
        Formules("=IF(N(#MN)>0,#M/#MN,IF(#n="""","""",99))")
    Ligne4 ws, r, lig, "tV", "Taux de cisaillement", "Tmax / Vpl,Rd (> 0,5 : réduction de fy)", "", _
        Formules("=IF(N(#Vpl)>0,#T/#Vpl,"""")")
    Ligne4 ws, r, lig, "vr", "Vérification des résistances", "n <= 1, Mmax <= MN,Rd, Tmax <= Vpl,Rd", "", _
        Formules("=IF(OR(#n="""",#tM="""",#tV=""""),""-"",IF(AND(#n<=1,#tM<=1,#tV<=1),""Vérifiée"",""Non vérifiée""))")

    ' --- Mise en forme ---
    etape = "mise en forme"
    With ws.Range("B" & r0 + 1 & ":I" & r)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("E" & r0 + 1 & ":I" & r).HorizontalAlignment = xlCenter
    With ws.Range("B" & r0 + 1 & ":I" & r0 + 1)
        .Font.Bold = True
        .Interior.Color = VERT_ENTETE
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("E" & r0 + 2 & ":H" & r).Interior.Color = JAUNE_CALCUL
    zone = lig("vs") & "," & lig("vr")
    For j = 0 To 1
        With ws.Range(Split(zone, ",")(j)).Resize(1, 4)
            .Font.Bold = True
            .FormatConditions.Delete
            .FormatConditions.Add(xlCellValue, xlEqual, "=""Non vérifiée""").Font.Color = RGB(192, 0, 0)
        End With
    Next j
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
