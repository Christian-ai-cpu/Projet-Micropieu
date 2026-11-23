Attribute VB_Name = "Charte"
Option Explicit

' Module « Charte graphique GEOSKILLS ENGINEERING » : couleurs et polices communes à
' toutes les feuilles du projet (charte graphique officielle) :
'   Bleu Nuit      #13005A  RGB(19, 0, 90)     couleur primaire : titres, textes forts
'   Bleu Foncé     #00337C  RGB(0, 51, 124)    couleur secondaire : éléments graphiques
'   Bleu Océan     #1C82AD  RGB(28, 130, 173)  couleur tertiaire : arrière-plans
'   Vert Émeraude  #03C988  RGB(3, 201, 136)   couleur d'accentuation : éléments clés
'   Blanc          #FFFFFF                     fond principal
' Fonds de cellules (teintes éclaircies de la palette, texte en Bleu Nuit lisible) :
'   en-têtes de tableaux : Bleu Océan éclairci ; valeurs saisies : Vert Émeraude
'   éclairci ; valeurs calculées : Bleu Océan très clair ; mini-titres : Bleu Océan clair.
' Polices : titres et sous-titres en CheGuevara Text Sans, corps de texte en Inter (à
' installer sur le poste ; sinon Excel les remplace par une police voisine).
'   AppliquerCharte ws : police Inter sur toute la feuille, CheGuevara Text Sans sur
'     les titres (texte Bleu Nuit en taille 13 ou plus, colonnes A et B) ;
'   AppliquerCharteClasseur : même chose pour toutes les feuilles du classeur ;
'   CouleurSerie k : couleur de la k-ième courbe d'un graphique (palette de la charte).
' Couleurs écrites en nombre : rouge + 256 x vert + 65536 x bleu.

Public Const BLEU_NUIT As Long = 5898259       ' RGB(19, 0, 90)
Public Const BLEU_FONCE As Long = 8139520      ' RGB(0, 51, 124)
Public Const BLEU_OCEAN As Long = 11371036     ' RGB(28, 130, 173)
Public Const VERT_EMERAUDE As Long = 8964355   ' RGB(3, 201, 136)
Public Const BLANC As Long = 16777215          ' RGB(255, 255, 255)

Private Const COUL_TITRE As Long = 5898259      ' Bleu Nuit : titres et sous-titres
Private Const COUL_ENTETE As Long = 15458502    ' RGB(198, 224, 235) : Bleu Océan éclairci
Private Const COUL_SAISIE As Long = 15201485    ' RGB(205, 244, 231) : Vert Émeraude éclairci
Private Const COUL_CALCUL As Long = 16315885    ' RGB(237, 245, 248) : Bleu Océan très clair
Private Const COUL_SOUS_TITRE As Long = 16117988 ' RGB(228, 240, 245) : Bleu Océan clair
Private Const COUL_ALERTE As Long = 192         ' RGB(192, 0, 0) : messages « non vérifié »

Public Const POLICE_TITRE As String = "CheGuevara Text Sans"
Public Const POLICE_TEXTE As String = "Inter"

' Polices de la charte sur une feuille : Inter partout, CheGuevara Text Sans sur les
' titres (colonnes A et B, texte Bleu Nuit en taille 13 ou plus)
Public Sub AppliquerCharte(ws As Worksheet)
    Dim c As Range, zone As Range, ecran As Boolean
    ecran = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error Resume Next
    ws.Cells.Font.Name = POLICE_TEXTE
    Set zone = Intersect(ws.UsedRange, ws.Range("A:B"))
    If Not zone Is Nothing Then
        For Each c In zone.Cells
            If Not IsEmpty(c.Value) Then
                If c.Font.Color = COUL_TITRE And c.Font.Size >= 13 Then c.Font.Name = POLICE_TITRE
            End If
        Next c
    End If
    On Error GoTo 0
    Application.ScreenUpdating = ecran
End Sub

' Polices de la charte sur toutes les feuilles du classeur actif
Public Sub AppliquerCharteClasseur()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        AppliquerCharte ws
    Next ws
End Sub

' Couleur de la k-ième courbe d'un graphique (k = 0, 1, 2, ...)
Public Function CouleurSerie(k As Long) As Long
    Select Case k Mod 4
        Case 0: CouleurSerie = BLEU_NUIT
        Case 1: CouleurSerie = BLEU_OCEAN
        Case 2: CouleurSerie = VERT_EMERAUDE
        Case Else: CouleurSerie = BLEU_FONCE
    End Select
End Function
