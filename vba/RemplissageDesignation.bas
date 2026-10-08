Attribute VB_Name = "RemplissageDesignation"
Option Explicit

' Listes déroulantes sous « Désignation de l'ouvrage » :
'   B5 : Toiture / Toiture terrasse
'   B6 : Poutre BA / Poutre bois / Poutre métallique
'   B7 : Poteau BA / Poteau bois / Poteau métallique
Public Sub RemplirDesignationOuvrage()

        With ActiveSheet.Range("B" & 5)
            .Validation.Delete
            .Validation.Add xlValidateList, xlValidAlertStop, , "Toiture,Toiture terrasse"
            .Value = "Toiture"
        End With
        With ActiveSheet.Range("B" & 6)
            .Validation.Delete
            .Validation.Add xlValidateList, xlValidAlertStop, , "Poutre BA,Poutre bois,Poutre métallique"
            .Value = "Poutre BA"
        End With
        With ActiveSheet.Range("B" & 7)
            .Validation.Delete
            .Validation.Add xlValidateList, xlValidAlertStop, , "Poteau BA,Poteau bois,Poteau métallique"
            .Value = "Poteau BA"
        End With
End Sub
