Attribute VB_Name = "RemplissageDesignation"
Option Explicit

' Remplit la colonne « Désignation de l'ouvrage » (B) du tableau des charges
' permanentes avec des menus et leurs sous-éléments, à partir de la ligne 4.
' Les lignes des sous-éléments sont groupées : le bouton +/- à gauche de la
' feuille permet de déplier ou replier chaque menu.
' À lancer après CreerTableauChargesPermanentes (ou sa version complète).

Private Const PREMIERE_LIGNE As Long = 4

Public Sub RemplirDesignationOuvrage()
    Dim menus As Variant, sousElements As Variant
    Dim ws As Worksheet
    Dim r As Long, i As Long, j As Long

    menus = Array("Toiture terrasse", "Plancher R+5")
    sousElements = Array( _
        Array("Poutre BA", "Poteau BA"), _
        Array("Plancher", "Poutre BA", "Poteau BA"))

    Set ws = ActiveSheet
    ws.Cells.ClearOutline              ' évite d'empiler les groupes si on relance
    ws.Outline.SummaryRow = xlAbove
    r = PREMIERE_LIGNE

    For i = 0 To UBound(menus)
        ' Ligne du menu : titre en gras, sans calcul de G
        ws.Range("B" & r).Value = menus(i)
        ws.Range("B" & r).Font.Bold = True
        ws.Range("G" & r & ":H" & r).ClearContents
        r = r + 1

        ' Sous-éléments : décalés d'un retrait et groupés sous le menu
        For j = 0 To UBound(sousElements(i))
            ws.Range("B" & r).Value = sousElements(i)(j)
            ws.Range("B" & r).IndentLevel = 1
            r = r + 1
        Next j
        ws.Rows(r - UBound(sousElements(i)) - 1 & ":" & r - 1).Group
    Next i
End Sub
