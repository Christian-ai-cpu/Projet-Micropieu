Attribute VB_Name = "TableauChargesPermanentesComplet"
Option Explicit

' Crée l'en-tête du tableau de descente de charges (colonnes B à H) :
'   B : Désignation de l'ouvrage
'   C : Poids volumique (kN/m3)
'   D : Poids Surfacique (kN/m²)
'   E:G : Dimensions (fusionnées) -> Largeur b (m) / Épaisseur h (m) / Nombre n
'   H : Linéaire L (m)
'   I : Charges permanentes G (kN)
'   J : Charge d'exploitation Q (kN), calculée sur les lignes de sous-total
'   K : ΣQ (kN) (Σ écrit ChrW(931)), cumul des Q de la toiture jusqu'au niveau
'   L : αn (α écrit ChrW(945)), coefficient de réduction selon le nombre d'étages
'   M : αn × ΣQ (kN), charge d'exploitation cumulée réduite
' Les lignes de saisie calculent G = n × (L / 2) × (γ × b × h + g × b),
' sauf pour les poteaux où le linéaire L est pris entier.
' En ligne 1 : listes déroulantes du niveau d'étage (RDC à 5) et de la catégorie du bâtiment (A à G).

Private Const LIGNE_NIVEAU As Long = 1
Private Const LIGNE_TITRE As Long = 3
Private Const LIGNE_SOUS_TITRE As Long = 4
Private Const PREMIERE_LIGNE As Long = 5
Private Const NB_LIGNES As Long = 49

Public Sub CreerTableauChargesPermanentesComplet()
    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim r As Long, c As Long
    Dim Entete(1 To 2, 1 To 12) As String
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

    ' --- Titres (ligne 1) et sous-titres (ligne 2) remplis par boucle ---
    titres = Array("Désignation de l'ouvrage", "Poids volumique (kN/m3)", _
        "Poids Surfacique (kN/m²)", "Dimensions", "", "", "Linéaire L (m)", "Charges permanentes G (kN)", _
        "Charge d'exploitation Q (kN)", ChrW(931) & "Q (kN)", ChrW(945) & "n", ChrW(945) & "n × " & ChrW(931) & "Q (kN)")
    sousTitres = Array("", "", "", "Largeur b (m)", "Épaisseur h (m)", "Nombre n", "", "", "", "", "", "")
    For r = 1 To 2
        For c = 1 To 12
            Entete(r, c) = IIf(r = 1, titres(c - 1), sousTitres(c - 1))
        Next c
    Next r
    ws.Range("B" & LIGNE_TITRE).Resize(2, 12).Value = Entete

    ' --- Fusions : verticale si pas de sous-titre, sinon « Dimensions » sur E:G ---
    For c = 1 To 12
        If Entete(2, c) = "" Then ws.Cells(LIGNE_TITRE, c + 1).Resize(2, 1).Merge
    Next c
    ws.Range("E" & LIGNE_TITRE & ":G" & LIGNE_TITRE).Merge

    ' --- Mise en forme de l'en-tête ---
    With ws.Range("B" & LIGNE_TITRE & ":M" & LIGNE_SOUS_TITRE)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' --- Lignes de saisie et formule de G ---
    For r = PREMIERE_LIGNE To derniereLigne
        ws.Range("G" & r).Value = 1
        ws.Range("I" & r).Formula = "=IF(B" & r & "="""","""",G" & r & "*H" & r & "/IF(LEFT(B" & r & ",6)=""Poteau"",1,2)" & _
            "*(C" & r & "*E" & r & "*F" & r & "+D" & r & "*E" & r & "))"
    Next r

    ' --- Ligne de total ---
    ws.Range("B" & derniereLigne + 1).Value = "TOTAL G"
    ws.Range("I" & derniereLigne + 1).Formula = _
        "=SUBTOTAL(109,I" & PREMIERE_LIGNE & ":I" & derniereLigne & ")"  ' ignore les lignes masquées
    With ws.Range("B" & derniereLigne + 1 & ":M" & derniereLigne + 1)
        .Font.Bold = True
    End With

    ' --- Bordures ---
    With ws.Range("B" & LIGNE_TITRE & ":M" & derniereLigne + 1).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
    ws.Range("B" & LIGNE_TITRE & ":M" & derniereLigne + 1).BorderAround _
        LineStyle:=xlContinuous, Weight:=xlMedium

    ' --- Renvoi à la ligne automatique (hauteur des lignes ajustée au texte) ---
    With ws.Range("B" & LIGNE_TITRE & ":M" & derniereLigne + 1)
        .WrapText = True
        .Rows.AutoFit
    End With
    ' Colonne I (charges G) ajustée à son contenu, puis hauteurs réajustées
    ws.Range("I" & PREMIERE_LIGNE & ":I" & derniereLigne + 1).WrapText = False
    ws.Columns("I").AutoFit
    If ws.Columns("I").ColumnWidth < 12 Then ws.Columns("I").ColumnWidth = 12   ' place pour le titre
    ws.Range("B" & LIGNE_TITRE & ":M" & derniereLigne + 1).Rows.AutoFit
    ' Ligne 1 : « Niveau d'étage » et « Catégorie du bâtiment » avec leurs listes
    With ws.Range("B" & LIGNE_NIVEAU & ":E" & LIGNE_NIVEAU)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Rows.AutoFit
    End With

    Application.ScreenUpdating = True
    MsgBox "Tableau des charges permanentes créé.", vbInformation
End Sub

