Attribute VB_Name = "RemplissageDesignation"
Option Explicit

' Listes déroulantes sous « Désignation de l'ouvrage », à partir de B5 :
'   B5 à B7   : Toiture / Toiture terrasse, puis Poutre et Poteau
'   B8 à B12  : Étage 5, puis Plancher, Poutre, Poteau et Mur
'   ...
'   B28 à B32 : Étage 1, puis Plancher, Poutre, Poteau et Mur
' Seuls les étages jusqu'au niveau choisi en C1 sont affichés (voir AfficherEtages).
' Une liste déroulante Excel est limitée à 255 caractères : la liste des murs
' utilise donc des libellés courts (Banché = béton banché, Agglo = agglos béton).
Public Sub RemplirDesignationOuvrage()
    Const PLANCHERS As String = "Plancher dalle pleine,Plancher à poutrelles,Plancher bois traditionnel"
    Const POUTRES As String = "Poutre BA,Poutre bois,Poutre métallique"
    Const POTEAUX As String = "Poteau BA,Poteau bois,Poteau métallique"
    Const MURS As String = "Agglo creux 15cm,Agglo creux 20cm,Agglo plein 15cm,Agglo plein 20cm," & _
        "Banché 20cm,Banché 25cm,Banché 30cm," & _
        "Brique creuse 15cm,Brique creuse 20cm,Brique creuse 25cm," & _
        "Monomur 30cm,Monomur 37cm," & _
        "Brique pleine 10cm,Brique pleine 21.5cm,Brique pleine 25cm"
    Dim listes As Variant, i As Long, e As Long, r As Long

    r = 5
    For e = 6 To 1 Step -1             ' 6 = toiture, puis étages 5 à 1
        If e = 6 Then
            listes = Array("Toiture,Toiture terrasse", POUTRES, POTEAUX)
        Else
            listes = Array("Étage " & e, PLANCHERS, POUTRES, POTEAUX, MURS)
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
        ' Bloc de l'étage e : 5 lignes à partir de la ligne 8 + 5 × (5 - e)
        ws.Rows(8 + 5 * (5 - e)).Resize(5).Hidden = (e > niveau)
    Next e
End Sub
