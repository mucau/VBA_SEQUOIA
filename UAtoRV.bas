Attribute VB_Name = "Module7"
Sub UA_RV()
' Début
Dim ROW, COLUMN As Variant

Cells.Select
Selection.Style = "Normal"

Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select

ROW = Selection.Rows.Count
COLUMN = Selection.Columns.Count

'Remplacement des points
COL = 1
For i = 1 To COLUMN
    If Left(Cells(1, COL), 5) = "SURF_" Then
        Columns(COL).Select
        Selection.Replace What:=".", Replacement:=".", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Selection.NumberFormat = "0.0000"
        COL = COL + 1
    Else
        COL = COL + 1
    End If
Next

'Mise en forme tableau
Randomize
Nom_tableau = "Tableau" & Str(Int(50 * Rnd) + 1)
ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(ROW, COLUMN)), , xlYes).Name = Nom_tableau
ActiveSheet.ListObjects(Nom_tableau).TableStyle = "TableStyleLight9"
ActiveSheet.ListObjects(Nom_tableau).ShowTableStyleRowStripes = False
ActiveSheet.ListObjects(Nom_tableau).ShowTotals = True


End Sub

