# VBA macros

A repository of spreadsheets with VBA macros

#HSLIDE

### Première macro

- pour pratiquer l'enregistrement de macros
- découvrir une partie du modèle object d'Excel
- pour faire un peu de programmation
- et ajouter quelques contrôles (bouton)

#HSLIDE

### Le pitch

- on reçoit dans la feuille "AVANT" des données brutes
- dont on souhaite modifier la mise en page dans une feuille "APRES"
- et associer un bouton à la macro ainsi créée

#HSLIDE

### Le code

```vbscript
Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
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
Sub MacroButeurs()
'
' MacroButeurs Macro
'
'
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
    
    ' Copie des colonnes et ré-arrangement
    Sheets("AVANT").Select
    Columns("D:D").Select
    Selection.Copy
    Sheets("APRES").Select
    Columns("A:A").Select
    ActiveSheet.Paste
    
    Sheets("AVANT").Select
    Columns("A:A").Select
    Selection.Copy
    Sheets("APRES").Select
    Columns("B:B").Select
    ActiveSheet.Paste

    Sheets("AVANT").Select
    Columns("B:B").Select
    Selection.Copy
    Sheets("APRES").Select
    Columns("C:C").Select
    ActiveSheet.Paste
    
    Sheets("AVANT").Select
    Columns("C:C").Select
    Selection.Copy
    Sheets("APRES").Select
    Columns("D:D").Select
    ActiveSheet.Paste
    
    Sheets("AVANT").Select
    Range("A1").Select
 
    ' En-têtes en gras et centré
    Sheets("APRES").Select
    Range("A1:D1").Select
    With Selection
        .Font.Bold = True
    End With
    
    ' Dimensionnement automatique des colonnes
    ActiveSheet.Range("A:D").EntireColumn.AutoFit
    
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
    
    ' Calcul du nombre total de buts marqués
    Range("D46").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-44]C:R[-1]C)"
    
    ' Plus de sélection pour le copy-paste
    Sheets("APRES").Select
    Range("A1").Select
    Application.CutCopyMode = False
    
    ' MsgBox (Application.Version)
End Sub
```
@[1](On importe les fonctions mathématiques - pour la racine carrée plus bas)
@[3-4](Définition de la fonction f)
@[6-7](Définition de la dérivée fp)
@[11](Estimation initiale de la solution)
@[12](On va faire 10 itérations)
@[13](On met à jour l'estimation de la solution)
@[15](Affichage de la solution obtenue par la méthode de Newton-Raphson)
@[16](Affichade la la solution "exacte" $sqrt(3)$)

