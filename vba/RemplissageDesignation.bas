Attribute VB_Name = "RemplissageDesignation"
Option Explicit

' Listes déroulantes sous « Désignation de l'ouvrage » :
'   B4 : Toiture / Toiture terrasse
'   B5 : Poutre BA / Poutre bois / Poutre métallique
'   B6 : Poteau BA / Poteau bois / Poteau métallique
Public Sub RemplirDesignationOuvrage()
    Dim listes As Variant, i As Long

    listes = Array("Toiture,Toiture terrasse", _
        "Poutre BA,Poutre bois,Poutre métallique", _
        "Poteau BA,Poteau bois,Poteau métallique")

    For i = 0 To UBound(listes)
        With ActiveSheet.Range("B" & 4 + i)
            .Validation.Delete
            .Validation.Add xlValidateList, xlValidAlertStop, , listes(i)
            .Value = Split(listes(i), ",")(0)
        End With
    Next i
End Sub
