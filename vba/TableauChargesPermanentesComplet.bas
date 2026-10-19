Attribute VB_Name = "TableauChargesPermanentesComplet"
Option Explicit

' Crée l'en-tête du tableau de descente de charges (colonnes B à H) :
'   B : Désignation de l'ouvrage
'   C : Poids volumique (kN/m3)
'   D : Poids Surfacique (kN/m²)
'   E:F : Dimensions (fusionnées) -> Largeur b (m) / Hauteur h (m)
'   G : Charges permanentes G (kN/ml)
'   H : Charge d'exploitation Q (kN/ml), calculée sur les lignes de sous-total
'   I : ΣQ (kN/ml) (Σ écrit ChrW(931)), cumul des Q de la toiture jusqu'au niveau
'   J : αn (α écrit ChrW(945)), coefficient de réduction selon le nombre d'étages
'   K : αn × ΣQ (kN/ml), charge d'exploitation cumulée réduite
' Descente de charges en kN/ml pour un mur : les lignes de saisie calculent
' G = γ × b × h + g × b, b étant la largeur reprise (m).
' En ligne 1 : listes déroulantes du niveau d'étage (RDC à 5) et de la catégorie du bâtiment (A à G),
' nombre de murs (G1) et pourcentage des ouvertures des murs (I1).

Private Const LIGNE_NIVEAU As Long = 1
Private Const LIGNE_TITRE As Long = 3
Private Const LIGNE_SOUS_TITRE As Long = 4
Private Const PREMIERE_LIGNE As Long = 5
Private Const NB_LIGNES As Long = 57

Public Sub CreerTableauChargesPermanentesComplet()
    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim r As Long, c As Long
    Dim Entete(1 To 2, 1 To 10) As String
    Dim titres As Variant, sousTitres As Variant

    Set ws = ActiveSheet
    derniereLigne = PREMIERE_LIGNE + NB_LIGNES - 1

    Application.ScreenUpdating = False

    ' --- Effacement complet de la feuille (contenu, formats, fusions, listes) ---
    ws.Cells.Clear

    ' --- Choix du niveau d'étage (liste déroulante RDC à 5) ---
    ws.Range("B" & LIGNE_NIVEAU).Value = "Niveau d'étage"
    ws.Range("B" & LIGNE_NIVEAU).Font.Bold = True
    With ws.Range("C" & LIGNE_NIVEAU)
        .Value = "RDC"
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="RDC,1,2,3,4,5"  ' en VBA, toujours la virgule
        .Validation.InCellDropdown = True
        .Validation.ErrorMessage = "Choisir un niveau d'étage : RDC, 1, 2, 3, 4 ou 5."
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With

    ' --- Choix de la catégorie du bâtiment (liste déroulante A à G) ---
    ws.Range("D" & LIGNE_NIVEAU).Value = "Catégorie du bâtiment"
    ws.Range("D" & LIGNE_NIVEAU).Font.Bold = True
    With ws.Range("E" & LIGNE_NIVEAU)
        .Value = "A"
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="A,B,C,D,E,F,G"
        .Validation.InCellDropdown = True
        .Validation.ErrorMessage = "Choisir une catégorie de bâtiment : A, B, C, D, E, F ou G."
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With

    ' --- Nombre de murs (G1) : le tableau est dupliqué vers la droite (DuplicationMurs) ---
    ws.Range("F" & LIGNE_NIVEAU).Value = "Entrez le nombre de murs"
    ws.Range("F" & LIGNE_NIVEAU).Font.Bold = True
    With ws.Range("G" & LIGNE_NIVEAU)
        .Value = 1
        .Validation.Delete
        .Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="1", Formula2:="50"
        .Validation.ErrorMessage = "Entrer un nombre entier de murs (1 à 50)."
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With

    ' --- Pourcentage des ouvertures des murs (I1) : G du mur × (100 - I1) / 100 ---
    ws.Range("H" & LIGNE_NIVEAU).Value = "Pourcentage des ouvertures des murs (%)"
    ws.Range("H" & LIGNE_NIVEAU).Font.Bold = True
    With ws.Range("I" & LIGNE_NIVEAU)
        .Value = 0
        .Validation.Delete
        .Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="0", Formula2:="100"
        .Validation.ErrorMessage = "Entrer un pourcentage d'ouvertures entre 0 et 100."
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With

    ' --- Titres (ligne 1) et sous-titres (ligne 2) remplis par boucle ---
    titres = Array("Désignation de l'ouvrage", "Poids volumique (kN/m3)", _
        "Poids Surfacique (kN/m²)", "Dimensions", "", "Charges permanentes G (kN/ml)", _
        "Charge d'exploitation Q (kN/ml)", ChrW(931) & "Q (kN/ml)", ChrW(945) & "n", ChrW(945) & "n × " & ChrW(931) & "Q (kN/ml)")
    sousTitres = Array("", "", "", "Largeur b (m)", "Hauteur h (m)", "", "", "", "", "")
    For r = 1 To 2
        For c = 1 To 10
            Entete(r, c) = IIf(r = 1, titres(c - 1), sousTitres(c - 1))
        Next c
    Next r
    ws.Range("B" & LIGNE_TITRE).Resize(2, 10).Value = Entete

    ' --- Fusions : verticale si pas de sous-titre, sinon « Dimensions » sur E:F ---
    For c = 1 To 10
        If Entete(2, c) = "" Then ws.Cells(LIGNE_TITRE, c + 1).Resize(2, 1).Merge
    Next c
    ws.Range("E" & LIGNE_TITRE & ":F" & LIGNE_TITRE).Merge

    ' --- Mise en forme de l'en-tête ---
    With ws.Range("B" & LIGNE_TITRE & ":K" & LIGNE_SOUS_TITRE)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' --- Lignes de saisie et formule de G ---
    For r = PREMIERE_LIGNE To derniereLigne
        ws.Range("G" & r).Formula = "=IF(B" & r & "="""",""""," & _
            "(C" & r & "*E" & r & "*F" & r & "+D" & r & "*E" & r & "))"
    Next r

    ' --- Ligne de total ---
    ws.Range("B" & derniereLigne + 1).Value = "TOTAL G"
    ws.Range("G" & derniereLigne + 1).Formula = _
        "=SUBTOTAL(109,G" & PREMIERE_LIGNE & ":G" & derniereLigne & ")"  ' ignore les lignes masquées
    With ws.Range("B" & derniereLigne + 1 & ":K" & derniereLigne + 1)
        .Font.Bold = True
    End With

    ' --- Bordures ---
    With ws.Range("B" & LIGNE_TITRE & ":K" & derniereLigne + 1).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
    ws.Range("B" & LIGNE_TITRE & ":K" & derniereLigne + 1).BorderAround _
        LineStyle:=xlContinuous, Weight:=xlMedium

    ' --- Renvoi à la ligne automatique (hauteur des lignes ajustée au texte) ---
    With ws.Range("B" & LIGNE_TITRE & ":K" & derniereLigne + 1)
        .WrapText = True
        .Rows.AutoFit
    End With
    ' Colonne G (charges G) ajustée à son contenu, puis hauteurs réajustées
    ws.Range("G" & PREMIERE_LIGNE & ":G" & derniereLigne + 1).WrapText = False
    ws.Columns("G").AutoFit
    If ws.Columns("G").ColumnWidth < 12 Then ws.Columns("G").ColumnWidth = 12   ' place pour le titre
    ws.Range("B" & LIGNE_TITRE & ":K" & derniereLigne + 1).Rows.AutoFit
    ' Ligne 1 : niveau d'étage, catégorie du bâtiment, nombre de murs et % d'ouvertures
    With ws.Range("B" & LIGNE_NIVEAU & ":I" & LIGNE_NIVEAU)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Rows.AutoFit
    End With

    Application.ScreenUpdating = True
    MsgBox "Tableau des charges permanentes créé.", vbInformation
End Sub

