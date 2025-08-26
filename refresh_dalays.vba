Sub atualizar()

Dim x As Long, y As Long, col As Integer

tempoinicial = Timer
Application.ScreenUpdating = False
ActiveWorkbook.RefreshAll

Sheets("Subgr fatu").Select
x = Range("A" & Rows.Count).End(xlUp).Row
y = Range("C" & Rows.Count).End(xlUp).Row

If x > y Then
    Range("C2:H2").Copy
    Range("C4").Select
    ActiveSheet.Paste
    Range("C4:H" & x).Select
    Selection.FillDown
    Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End If
    
Sheets("Base Embalagens").Select
x = Range("A" & Rows.Count).End(xlUp).Row
y = Range("V" & Rows.Count).End(xlUp).Row
Range("AE1").FormulaLocal = "=SE(AB1=1;1/CONT.SES($B$1:$B$" & x & ";B1);0)"
Range("AF1").FormulaLocal = "=SE(E(AB1=1;AC1>0);1/CONT.SES($B$1:$B$" & x & ";B1);0)"

If x > y Then
    For col = 22 To 35
        Cells(1, col).Copy
        Cells(3, col).Select
        ActiveSheet.Paste
        Range(Cells(3, col), Cells(x, col)).Select
        Selection.FillDown
        Calculate
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
           :=False, Transpose:=False
     Next col
End If

Sheets("Análises").Select
ActiveSheet.PivotTables("Tabela dinâmica5").PivotCache.Refresh
x = Range("W" & Rows.Count).End(xlUp).Row
y = Range("X" & Rows.Count).End(xlUp).Row

For col = 25 To 27
    Cells(1, col).Copy
    Cells(4, col).Select
    ActiveSheet.Paste
    Range(Cells(4, col), Cells(x, col)).Select
    Selection.FillDown
    Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Next col

ActiveWorkbook.Worksheets("Análises").AutoFilter.Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Análises").AutoFilter.Sort.SortFields.Add2 Key:= _
    Range("Z3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets("Análises").AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

Range(Cells(1, 24), Cells(2, 24)).Copy
Range("X4").Select
ActiveSheet.Paste
Range(Cells(5, 24), Cells(x, 24)).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Range("AB1:AM1").Copy
Range("AB4:AM4").Select
ActiveSheet.Paste
Range("AB5:AM5").Select
ActiveSheet.Paste

Range("AD5").FormulaR1C1 = "=R[-1]C+RC[-1]"
Range("AG5").FormulaR1C1 = "=R[-1]C+RC[-1]"
Range("AJ5").FormulaR1C1 = "=R[-1]C+RC[-1]"
Range("AM5").FormulaR1C1 = "=R[-1]C+RC[-1]"
    
For col = 28 To 39
    Range(Cells(5, col), Cells(x, col)).Select
    Selection.FillDown
    Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Next col

Application.CutCopyMode = False
ActiveWorkbook.RefreshAll
Sheets("Dashboard").Select
Range("A1").Select
tempofinal = (Timer - tempoinicial) / 60
MsgBox ("Atualização finalizada com sucesso!" & Chr(13) & "Tempo: " & Round(tempofinal, 2) & " minutos")

End Sub
