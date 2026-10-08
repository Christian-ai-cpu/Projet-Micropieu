Attribute VB_Name = "RemplissageDesignation"
Option Explicit

' Remplit la colonne « Désignation de l'ouvrage » (B) du tableau des charges
' permanentes avec des menus et leurs sous-éléments, à partir de la ligne 4.
' Un élément écrit "choix1,choix2,..." devient une liste déroulante
' (choix1 affiché par défaut) ; un élément sans virgule est un texte simple.
' Les lignes des sous-éléments sont groupées : le bouton +/- à gauche de la
' feuille permet de déplier ou replier chaque menu.
' À lancer après CreerTableauChargesPermanentes (ou sa version complète).

Private Const PREMIERE_LIGNE As Long = 4

Public Sub RemplirDesignationOuvrage()
    Dim menus As Variant, sousElements As Variant
    Dim ws As Worksheet
    Dim r As Long, i As Long, j As Long

    menus = Array("Toiture,Toiture terrasse", "Plancher R+5")
    sousElements = Array( _
        Array("Poutre BA,Poutre bois,Poutre métallique", "Poteau BA,Poteau bois,Poteau métallique"), _
        Array("Plancher", "Poutre BA", "Poteau BA"))

    Set ws = ActiveSheet
    ws.Cells.ClearOutline              ' évite d'empiler les groupes si on relance
    ws.Outline.SummaryRow = xlAbove
    r = PREMIERE_LIGNE

    For i = 0 To UBound(menus)
        ' Ligne du menu : titre en gras, sans calcul de G
        EcrireChoix ws.Range("B" & r), menus(i)
        ws.Range("B" & r).Font.Bold = True
        ws.Range("G" & r & ":H" & r).ClearContents
        r = r + 1

        ' Sous-éléments : décalés d'un retrait et groupés sous le menu
        For j = 0 To UBound(sousElements(i))
            EcrireChoix ws.Range("B" & r), sousElements(i)(j)
            ws.Range("B" & r).IndentLevel = 1
            r = r + 1
        Next j
        ws.Rows(r - UBound(sousElements(i)) - 1 & ":" & r - 1).Group
    Next i
End Sub

' Écrit le premier choix dans la cellule ; s'il y en a plusieurs,
' ajoute une liste déroulante avec tous les choix.
Private Sub EcrireChoix(cellule As Range, choix As Variant)
    cellule.Validation.Delete
    cellule.Value = Split(choix, ",")(0)
    If InStr(choix, ",") > 0 Then
        cellule.Validation.Add xlValidateList, xlValidAlertStop, , choix
    End If
End Sub
