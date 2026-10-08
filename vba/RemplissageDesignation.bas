Attribute VB_Name = "RemplissageDesignation"
Option Explicit

' Listes déroulantes sous « Désignation de l'ouvrage », à partir de B5 :
'   B5 à B7   : Toiture / Toiture terrasse, puis Poutre et Poteau
'   B8 à B10  : Étage 1, puis Poutre et Poteau
'   ...
'   B20 à B22 : Étage 5, puis Poutre et Poteau
Public Sub RemplirDesignationOuvrage()
    Const POUTRES As String = "Poutre BA,Poutre bois,Poutre métallique"
    Const POTEAUX As String = "Poteau BA,Poteau bois,Poteau métallique"
    Dim listes As Variant, i As Long, e As Long, r As Long

    r = 5
    For e = 0 To 5
        If e = 0 Then
            listes = Array("Toiture,Toiture terrasse", POUTRES, POTEAUX)
        Else
            listes = Array("Étage " & e, POUTRES, POTEAUX)
        End If

        For i = 0 To UBound(listes)
            With ActiveSheet.Range("B" & r)
                .Validation.Delete
                ' Liste déroulante seulement s'il y a plusieurs choix
                If InStr(listes(i), ",") > 0 Then
                    .Validation.Add xlValidateList, xlValidAlertStop, , listes(i)
                End If
                .Value = Split(listes(i), ",")(0)
            End With
            r = r + 1
        Next i
    Next e
End Sub
