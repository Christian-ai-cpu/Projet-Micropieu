Attribute VB_Name = "TableauChargesPermanentesComplet"
Option Explicit

' Crée l'en-tête du tableau de descente de charges (colonnes B à H) :
'   B : Désignation de l'ouvrage
'   C : Poids volumique (kN/m3)
'   D : Poids Surfacique (kN/m²)
'   E:G : Dimensions (fusionnées) -> Largeur b (m) / Épaisseur h (m) / Nombre n
'   H : Charges permanentes G (kN/ml)
' Les lignes de saisie calculent G = n × (γ × b × h + g × b).
' En ligne 1 : liste déroulante pour choisir le niveau d'étage (1 à 5).

Private Const LIGNE_NIVEAU As Long = 1
Private Const LIGNE_TITRE As Long = 2
Private Const LIGNE_SOUS_TITRE As Long = 3
Private Const PREMIERE_LIGNE As Long = 4
Private Const NB_LIGNES As Long = 15

Public Sub CreerTableauChargesPermanentesComplet()
    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim r As Long

    Set ws = ActiveSheet
    derniereLigne = PREMIERE_LIGNE + NB_LIGNES - 1

    Application.ScreenUpdating = False

    ' --- Choix du niveau d'étage (liste déroulante 1 à 5) ---
    ws.Range("B" & LIGNE_NIVEAU).Value = "Niveau d'étage"
    ws.Range("B" & LIGNE_NIVEAU).Font.Bold = True
    With ws.Range("C" & LIGNE_NIVEAU)
        .Value = 1
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="1,2,3,4,5"  ' en VBA, toujours la virgule
        .Validation.InCellDropdown = True
        .Validation.ErrorMessage = "Choisir un niveau d'étage entre 1 et 5."
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 242, 204)
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With

    ' --- Titres principaux ---
    ws.Range("B" & LIGNE_TITRE).Value = "Désignation de l'ouvrage"
    ws.Range("C" & LIGNE_TITRE).Value = "Poids volumique (kN/m3)"
    ws.Range("D" & LIGNE_TITRE).Value = "Poids Surfacique (kN/m²)"
    ws.Range("E" & LIGNE_TITRE).Value = "Dimensions"
    ws.Range("H" & LIGNE_TITRE).Value = "Charges permanentes G (kN/ml)"

    ' --- Sous-titres de la colonne Dimensions ---
    ws.Range("E" & LIGNE_SOUS_TITRE).Value = "Largeur b (m)"
    ws.Range("F" & LIGNE_SOUS_TITRE).Value = "Épaisseur h (m)"
    ws.Range("G" & LIGNE_SOUS_TITRE).Value = "Nombre n"

    ' --- Fusions ---
    ws.Range("B" & LIGNE_TITRE & ":B" & LIGNE_SOUS_TITRE).Merge
    ws.Range("C" & LIGNE_TITRE & ":C" & LIGNE_SOUS_TITRE).Merge
    ws.Range("D" & LIGNE_TITRE & ":D" & LIGNE_SOUS_TITRE).Merge
    ws.Range("E" & LIGNE_TITRE & ":G" & LIGNE_TITRE).Merge
    ws.Range("H" & LIGNE_TITRE & ":H" & LIGNE_SOUS_TITRE).Merge

    ' --- Mise en forme de l'en-tête ---
    With ws.Range("B" & LIGNE_TITRE & ":H" & LIGNE_SOUS_TITRE)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' --- Lignes de saisie et formule de G ---
    For r = PREMIERE_LIGNE To derniereLigne
        ws.Range("G" & r).Value = 1
        ws.Range("H" & r).Formula = "=IF(B" & r & "="""","""",G" & r & _
            "*(C" & r & "*E" & r & "*F" & r & "+D" & r & "*E" & r & "))"
    Next r

    ' --- Ligne de total ---
    ws.Range("B" & derniereLigne + 1).Value = "TOTAL G"
    ws.Range("H" & derniereLigne + 1).Formula = _
        "=SUM(H" & PREMIERE_LIGNE & ":H" & derniereLigne & ")"
    With ws.Range("B" & derniereLigne + 1 & ":H" & derniereLigne + 1)
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With

    ' --- Formats numériques ---
    ws.Range("C" & PREMIERE_LIGNE & ":F" & derniereLigne).NumberFormat = "0.00"
    ws.Range("G" & PREMIERE_LIGNE & ":G" & derniereLigne).NumberFormat = "0"
    ws.Range("H" & PREMIERE_LIGNE & ":H" & derniereLigne + 1).NumberFormat = "0.00"
    ws.Range("C" & PREMIERE_LIGNE & ":H" & derniereLigne + 1).HorizontalAlignment = xlCenter

    ' --- Bordures ---
    With ws.Range("B" & LIGNE_TITRE & ":H" & derniereLigne + 1).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
    ws.Range("B" & LIGNE_TITRE & ":H" & derniereLigne + 1).BorderAround _
        LineStyle:=xlContinuous, Weight:=xlMedium

    ' --- Largeurs de colonnes et hauteurs d'en-tête ---
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 35
    ws.Columns("C:D").ColumnWidth = 16
    ws.Columns("E:G").ColumnWidth = 13
    ws.Columns("H").ColumnWidth = 20
    ws.Rows(LIGNE_TITRE).RowHeight = 30
    ws.Rows(LIGNE_SOUS_TITRE).RowHeight = 30

    Application.ScreenUpdating = True
    MsgBox "Tableau des charges permanentes créé.", vbInformation
End Sub
