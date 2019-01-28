# VBA macros

A repository of spreadsheets with VBA macros

#HSLIDE

### Première macro

- pour pratiquer l'enregistrement de macros
- découvrir le modèle objet d'Excel
- pour faire un peu de programmation
- et ajouter quelques contrôles (bouton)

#HSLIDE
<<<<<<< HEAD
=======

### Le pitch

- on reçoit dans la feuille "AVANT" des données brutes
- dont on souhaite modifier la mise en page dans une feuille "APRES"
- et associer un bouton à la macro ainsi créée

#HSLIDE

### Le code - fonction auxillaire

```vbscript
Option Explicit
Function sheetExists(sheetToFind As String) As Boolean
    Dim Sheet As Object
        
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function
```
@[1](On se force à déclarer toutes les variables - c'est une bonne pratique)
@[2](On décrit les arguments en entrée (sheetToFind) et le type du résultat)
@[5](On suppose que la feuille cherchée n'existe pas)
@[6](On parcourt toutes les feuilles de la collection (Worksheets))
@[7-10](Si on a un match, alors la feuille existe et on renverra "True")

#HSLIDE

### Le code - macro principale

```vbscript
Sub MacroButeurs()

    Dim ws As Object
    Dim Color As Integer
    Dim loopRow As Object
    
    ' On détruit la feuille "APRES"
    '    et on en recrée une neuve.
    If sheetExists("APRES") Then
        ' Pour éviter la demande de confirmation
        Application.DisplayAlerts = False
        Sheets("APRES").Delete
        ' Pour restaurer la demande de confirmation
        Application.DisplayAlerts = True
    End If
    Sheets("AVANT").Select
    Set ws = Sheets.Add(After:=ActiveSheet)
    ws.Name = "APRES"
```
@[1](On donne le nom MacroButeurs à notre macro)
@[7-18](On détruit la feuille APRES et en recrée une toute neuve)

#HSLIDE

### Le code - macro principale - suite

```vbscript
    ' Copie des colonnes et ré-arrangement
    Sheets("AVANT").Select
    Columns("D:D").Select
    Selection.Copy
    Sheets("APRES").Select
    Columns("A:A").Select
    ActiveSheet.Paste
    
    ...

    Sheets("AVANT").Select
    Range("A1").Select
```
@[1-7](Copie de la colonne D de AVANT vers la colonne A de APRES)
@[9](On répète pour les 3 autres colonnes)
@[11-12](On "clique" dans la cellule A1)

#HSLIDE

### Le code - macro principale - suite

```vbscript
    ' En-têtes en gras et centré
    Sheets("APRES").Select
    Range("A1:D1").Select
    With Selection
        .Font.Bold = True
    End With
    
    ' Dimensionnement automatique des colonnes
    ActiveSheet.Range("A:D").EntireColumn.AutoFit
```
@[1-6](Sélection de l'entête, centré, gras)
@[8-9](Ajustement automatique de la largeur des colonnes)

#HSLIDE

### Le code - macro principale - suite

```vbscript
    ' Coloriage des cellules en fonction du pays
    Sheets("APRES").Select
    For Each loopRow In ActiveSheet.UsedRange.Rows
        'Debug.Print loopRow.Row
        Select Case loopRow.Cells(1, 1).Value
        Case "Allemagne"
            loopRow.Interior.Color = Sheets("Description").Range("Q17").Interior.Color
        Case "Italie"
            loopRow.Interior.Color = Sheets("Description").Range("Q18").Interior.Color
        Case "France"
            loopRow.Interior.Color = Sheets("Description").Range("Q19").Interior.Color
        Case "Espagne"
            loopRow.Interior.Color = Sheets("Description").Range("Q20").Interior.Color
        Case "Angleterre"
            loopRow.Interior.Color = Sheets("Description").Range("Q21").Interior.Color
        End Select
    Next
```
@[3](On parcourt les lignes utiles de la feuille - UsedRange.Rows)
@[5](Aiguillage en fonction de la valeur de la première cellule de la la ligne)
@[7,9,11,13,15](Coloriage avec le code couleur de la page "Description")
    
#HSLIDE

### Le code - macro principale - suite

```vbscript
    ' Calcul du nombre total de buts marqués
    Range("D46").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-44]C:R[-1]C)"
    
    ' Plus de sélection pour le copy-paste
    Sheets("APRES").Select
    Range("A1").Select
    Application.CutCopyMode = False
    
End Sub
```
@[3](On rentre la fonction à appliquer dans la cellule : on somme cwles données entre les lignes [ici-44] et [ici-1])
@[6-7](On clique sur la cellule "A1")
>>>>>>> f10a57c5274c38d307bc68ef6054f0ab45a118a6
