Attribute VB_Name = "RemplissageDesignation"
Option Explicit

' Listes déroulantes sous « Désignation de l'ouvrage », à partir de B5 :
'   B5 à B7   : Toiture / Toiture terrasse, puis Poutre et Poteau
'   B8 à B10  : Étage 5, puis Poutre et Poteau
'   ...
'   B20 à B22 : Étage 1, puis Poutre et Poteau
' Seuls les étages jusqu'au niveau choisi en C1 sont affichés (voir AfficherEtages).
Public Sub RemplirDesignationOuvrage()
    Const POUTRES As String = "Poutre BA,Poutre bois,Poutre métallique"
    Const POTEAUX As String = "Poteau BA,Poteau bois,Poteau métallique"
    Dim listes As Variant, i As Long, e As Long, r As Long

    r = 5
    For e = 6 To 1 Step -1             ' 6 = toiture, puis étages 5 à 1
        If e = 6 Then
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
                .Font.Bold = (i = 0)          ' Toiture et Étage n en gras
            End With
            r = r + 1
        Next i
    Next e

    AfficherEtages ActiveSheet
End Sub

' Affiche la toiture et les étages 1 à niveau (C1), masque les autres.
' RDC -> toiture seule ; 1 -> toiture + étage 1 ; ... ; 5 -> tout.
' Appelée aussi automatiquement quand C1 change (voir CodeFeuille.txt).
Public Sub AfficherEtages(ws As Worksheet)
    Dim niveau As Long, e As Long

    niveau = Val(ws.Range("C1").Value)     ' "RDC" donne 0
    For e = 1 To 5
        ' Bloc de l'étage e : 3 lignes à partir de la ligne 5 + 3 × (6 - e)
        ws.Rows(5 + 3 * (6 - e)).Resize(3).Hidden = (e > niveau)
    Next e
End Sub
