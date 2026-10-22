Attribute VB_Name = "TableauxNFP94262"
Option Explicit

' Module « Tableaux NF P94-262 » : crée (ou recrée) la feuille « Tableaux NF P94-262 »
' avec les tableaux de la norme utilisés pour le calcul des micropieux (méthode
' pressiométrique) :
'   - Tableau A.1   : classes et catégories de pieux (techniques de mise en oeuvre)
'   - Tableau F.4.2.1 : facteur de portance pressiométrique kp max (par classe)
'   - Tableau F.5.2.1 : coefficient alpha pieu-sol (par catégorie)
'   - Tableau F.5.2.2 : frottement axial unitaire limite qs max en kPa (par catégorie)
'   - Tableau F.5.2.3 : paramètres a, b, c de la courbe fsol
'                       fsol = (a × pl* + b) × (1 - exp(-c × pl*)), pl* en MPa
'   - Facteurs partiels sur les résistances (approche 2) et facteurs de modèle
' Colonnes de sols : catégories de la NF P94-262 (argiles et limons, sols
' intermédiaires, sables et graves, craies, marnes et calcaires marneux, roches
' altérées ou fragmentées). « - » : valeur non fournie par la norme.
' ATTENTION : valeurs saisies de mémoire, à vérifier sur la norme avant usage.

Private Const NOM_FEUILLE As String = "Tableaux NF P94-262"
Private Const SOLS As String = "Argiles / limons|Sols intermédiaires|Sables / graves|" & _
    "Craies|Marnes / calcaires marneux|Roches altérées ou fragmentées"

Public Sub CreerTableauxNFP94262()
    Dim ws As Worksheet, r As Long, s As String

    Application.ScreenUpdating = False
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(NOM_FEUILLE)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = NOM_FEUILLE
    End If
    ws.Cells.Clear

    With ws.Range("B1:I1")
        .Merge
        .Value = "TABLEAUX DE LA NORME NF P94-262 (juillet 2012) - MICROPIEUX"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("B2:I2")
        .Merge
        .Value = "Valeurs à vérifier sur la norme avant utilisation."
        .Font.Italic = True
        .Font.Color = RGB(192, 0, 0)
        .HorizontalAlignment = xlCenter
    End With
    r = 4

    ' --- Tableau A.1 : classes et catégories ---
    s = "1|1|Foré simple (pieux et barrettes)|FS;1|2|Foré boue (pieux et barrettes)|FB;" & _
        "1|3|Foré tubé (virole perdue)|FTP;1|4|Foré tubé (virole récupérée)|FTR;" & _
        "1|5|Foré simple ou boue avec rainurage ou puits|FSR, FBR, PU;" & _
        "2|6|Foré tarière creuse simple ou double rotation|FTC, FTCD;" & _
        "3|7|Vissé moulé|VM;3|8|Vissé tubé|VT;" & _
        "4|9|Battu béton préfabriqué ou précontraint|BPF, BPR;" & _
        "4|10|Battu enrobé (béton, mortier, coulis)|BE;4|11|Battu moulé|BM;" & _
        "4|12|Battu acier fermé|BAF;5|13|Battu acier ouvert|BAO;" & _
        "6|14|Profilé H battu|HB;6|15|Profilé H battu injecté|HBi;" & _
        "7|16|Palplanches battues|PP;1bis|17|Micropieu type I|M1;" & _
        "1bis|18|Micropieu type II|M2;" & _
        "8|19|Pieu ou micropieu injecté mode IGU (type III)|PIGU, MIGU;" & _
        "8|20|Pieu ou micropieu injecté mode IRS (type IV)|PIRS, MIRS"
    r = EcrireTableau(ws, r, "Tableau A.1 - Classes et catégories de pieux", _
        "Classe|Catégorie|Technique de mise en oeuvre|Abréviation", s)

    ' --- Tableau F.4.2.1 : kp max par classe ---
    s = "1|1.15|1.1|1|1.45|1.45|1.45;2|1.3|1.65|1.6|2|2|2;3|1.55|1.95|3.5|3.4|3.4|3.4;" & _
        "4|1.35|1.65|3.1|2.35|2.35|2.35;5|1|1|1|1|1|1;6|1.2|1.1|1|1.45|1.45|1.45;" & _
        "7|1|1|1|1|1|1;8|1.15|1.1|1|1.45|1.45|1.45"
    r = EcrireTableau(ws, r, "Tableau F.4.2.1 - Facteur de portance pressiométrique kp max " & _
        "(résistance de pointe ; non retenue pour les micropieux)", "Classe|" & SOLS, s)

    ' --- Tableau F.5.2.1 : alpha pieu-sol par catégorie ---
    s = "1 - FS|1.1|1.25|1.4|1.8|1.5|1.6;2 - FB|1.25|1.4|1.4|1.8|1.5|1.6;" & _
        "3 - FTP|0.7|0.6|0.5|0.9|-|-;4 - FTR|1.25|1.4|1.7|1.4|-|-;5 - FSR, FBR, PU|1.3|-|-|-|-|-;" & _
        "6 - FTC, FTCD|1.5|1.8|2.1|1.6|1.6|2;7 - VM|1.9|2.1|1.7|1.7|-|-;8 - VT|0.6|0.6|1|0.7|-|-;" & _
        "9 - BPF, BPR|1.1|1.4|1|0.9|-|-;10 - BE|2|2.1|1.9|1.6|-|-;11 - BM|1.2|1.4|2.1|1|-|-;" & _
        "12 - BAF|0.8|1.2|0.4|0.9|-|-;13 - BAO|1.2|0.7|0.5|1|1|1;14 - HB|1.1|1|0.4|1|0.9|-;" & _
        "15 - HBi|2.7|2.9|2.4|2.4|2.4|2.4;16 - PP|0.9|0.8|0.4|1.2|1.2|-;17 - M1|-|-|-|-|-|-;" & _
        "18 - M2|-|-|-|-|-|-;19 - PIGU, MIGU|2.7|2.9|2.4|2.4|2.4|2.4;" & _
        "20 - PIRS, MIRS|3.4|3.8|3.8|3.1|3.1|3.1"
    r = EcrireTableau(ws, r, "Tableau F.5.2.1 - Coefficient " & ChrW(945) & _
        " pieu-sol (méthode pressiométrique)", "Catégorie|" & SOLS, s)

    ' --- Tableau F.5.2.2 : qs max par catégorie (kPa) ---
    s = "1 - FS|90|90|90|200|170|200;2 - FB|90|90|90|200|170|200;3 - FTP|50|50|50|50|90|-;" & _
        "4 - FTR|90|90|90|170|170|-;5 - FSR, FBR, PU|90|-|-|-|-|-;6 - FTC, FTCD|90|90|170|200|200|200;" & _
        "7 - VM|130|130|200|170|170|-;8 - VT|50|50|90|90|90|-;9 - BPF, BPR|130|130|130|90|90|-;" & _
        "10 - BE|170|170|260|200|200|-;11 - BM|90|90|130|260|200|-;12 - BAF|90|90|50|90|90|90;" & _
        "13 - BAO|90|90|50|90|90|90;14 - HB|90|90|130|90|90|-;15 - HBi|200|200|380|320|320|320;" & _
        "16 - PP|90|90|50|90|90|90;17 - M1|-|-|-|-|-|-;18 - M2|-|-|-|-|-|-;" & _
        "19 - PIGU, MIGU|200|200|380|320|320|320;20 - PIRS, MIRS|200|200|440|440|440|500"
    r = EcrireTableau(ws, r, "Tableau F.5.2.2 - Frottement axial unitaire limite qs max (kPa)", _
        "Catégorie|" & SOLS, s)

    ' --- Tableau F.5.2.3 : paramètres de fsol ---
    s = "a|0.003|0.01|0.007|0.008|0.01|0.01;b|0.04|0.06|0.07|0.08|0.08|0.08;" & _
        "c|3.5|1.2|1.3|3|3|3"
    r = EcrireTableau(ws, r, "Tableau F.5.2.3 - Paramètres de la courbe fsol = (a.pl* + b)" & _
        "(1 - exp(-c.pl*)), pl* en MPa, fsol en MPa", "Paramètre|" & SOLS, s)

    ' --- Facteurs partiels et de modèle ---
    s = ChrW(947) & "b (pointe) - ELU fondamental|1.1|-;" & _
        ChrW(947) & "s (frottement, compression) - ELU fondamental|1.1|-;" & _
        ChrW(947) & "t (total, compression) - ELU fondamental|1.1|-;" & _
        ChrW(947) & "s;t (frottement, traction) - ELU fondamental|1.15|-;" & _
        ChrW(947) & "b, " & ChrW(947) & "s, " & ChrW(947) & "t - ELU accidentel|1|-;" & _
        ChrW(947) & "s;t - ELU accidentel|1.05|-;" & _
        ChrW(947) & "cr (compression) - ELS caractéristique|0.9|-;" & _
        ChrW(947) & "cr (compression) - ELS quasi-permanent|1.1|-;" & _
        ChrW(947) & "s;cr (traction) - ELS caractéristique|1.1|-;" & _
        ChrW(947) & "s;cr (traction) - ELS quasi-permanent|1.5|-;" & _
        ChrW(947) & "R;d1 - méthode pressiométrique, compression|1.15|-;" & _
        ChrW(947) & "R;d1 - méthode pressiométrique, traction|1.4|-;" & _
        ChrW(947) & "R;d2 - méthode pressiométrique|1.1|-"
    r = EcrireTableau(ws, r, "Facteurs partiels sur les résistances (approche 2) et " & _
        "facteurs de modèle", "Facteur|Valeur|Remarque", s)

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 45
    ws.Columns("C:I").ColumnWidth = 16
    ws.Activate
    Application.ScreenUpdating = True
End Sub

' Écrit un tableau à partir de la ligne r : titre (gras, souligné), ligne d'en-tête
' (colonnes séparées par « | ») puis lignes de données (lignes séparées par « ; »,
' colonnes par « | », nombres avec un point décimal). Renvoie la ligne libre
' suivante (une ligne vide après le tableau).
Private Function EcrireTableau(ws As Worksheet, ByVal r As Long, titre As String, _
        entetes As String, donnees As String) As Long
    Dim e As Variant, lignes As Variant, cols As Variant, i As Long, j As Long, debut As Long

    With ws.Range("B" & r)
        .Value = titre
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
    End With
    r = r + 1
    debut = r
    e = Split(entetes, "|")
    For j = 0 To UBound(e)
        ws.Cells(r, 2 + j).Value = e(j)
    Next j
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 2 + UBound(e))).Font.Bold = True
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 2 + UBound(e))).Interior.Color = RGB(217, 225, 242)
    lignes = Split(donnees, ";")
    For i = 0 To UBound(lignes)
        r = r + 1
        cols = Split(lignes(i), "|")
        For j = 0 To UBound(cols)
            If IsNumeric(Replace(cols(j), ".", Application.International(xlDecimalSeparator))) _
                    And j > 0 Then
                ws.Cells(r, 2 + j).Value = Val(cols(j))      ' Val lit le point décimal
            Else
                ws.Cells(r, 2 + j).NumberFormat = "@"         ' texte (« 1bis », « - »)
                ws.Cells(r, 2 + j).Value = cols(j)
            End If
        Next j
    Next i
    With ws.Range(ws.Cells(debut, 2), ws.Cells(r, 2 + UBound(e)))
        .Borders.LineStyle = xlContinuous
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(debut, 3), ws.Cells(r, 2 + UBound(e))).HorizontalAlignment = xlCenter
    EcrireTableau = r + 2
End Function
