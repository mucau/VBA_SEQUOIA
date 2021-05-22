Attribute VB_Name = "Module13"
Sub test()

'Sélection des données
ActiveCell.CurrentRegion.Select
Selection.Style = "Normal"
ROW = Selection.Rows.Count
column = Selection.Columns.Count
Value = InputBox("Entrez le nombre de colonne fixe", "test", 1)

'Déclaration des feuilles
ActiveSheet.Name = "SheetA"
Sheets.Add After:=Sheets(Sheets.Count)
ActiveSheet.Name = "SheetB"

Dim wb As Workbook: Set wb = ActiveWorkbook
Dim wsa As Worksheet: Set wsa = wb.Worksheets("SheetA")
Dim wsb As Worksheet: Set wsb = wb.Worksheets("SheetB")
  
'Inscription des titres fixe
For A = 1 To Value
    wsb.Cells(1, A) = wsa.Cells(1, A)
Next
wsb.Cells(1, Value + 1) = "Colonne1"
wsb.Cells(1, Value + 2) = "Colonne2"


'Mise en BDD
R = 2
For B = 2 To ROW
    For C = Value + 1 To column
        For A = 1 To Value
        wsb.Cells(R, A) = wsa.Cells(B, A)
        Next
    wsb.Cells(R, Value + 1) = wsa.Cells(1, C)
    wsb.Cells(R, Value + 2) = wsa.Cells(B, C)
    R = R + 1
    Next
Next

'Mise en forme de tableau
wsb.Cells(1, 1).Select
ActiveCell.CurrentRegion.Select

Nom_tableau = "Table991"
ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = Nom_tableau
ActiveSheet.ListObjects(Nom_tableau).TableStyle = "TableStyleLight9"
ActiveSheet.ListObjects(Nom_tableau).ShowTableStyleRowStripes = False
ActiveSheet.ListObjects(Nom_tableau).ShowTotals = True
        
End Sub
