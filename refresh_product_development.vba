Sub Atualizar_1()

Dim x, y As Long

tempoinicial = Timer

Application.EnableEvents = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False
ActiveWorkbook.RefreshAll

'Arrasta formulas Subgr fatu

Sheets("Subgr fatu").Select
ActiveSheet.PivotTables("Tabela dinâmica2").PivotCache.Refresh
y = Range("A3").End(xlDown).Row
Range("C2:L2").Copy
Range("C4:L4").PasteSpecial
Range("C4:L" & y).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'Arrasta formulas Subgr prod

Sheets("Subgr prod").Select
ActiveSheet.PivotTables("Tabela dinâmica3").PivotCache.Refresh
y = Range("A3").End(xlDown).Row
Range("C2:J2").Copy
Range("C4:J4").PasteSpecial
Range("C4:J" & y).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

'Arrasta formula subgrupo na base de faturamento

Sheets("Base faturamento").Select
x = Range("A" & Rows.Count).End(xlUp).Row
y = Range("AY" & Rows.Count).End(xlUp).Row
Range("AY1").Copy
Range("AY3").PasteSpecial
Range("AY3:AY" & x).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'Arrasta formula subgrupo na base de produtos

Sheets("Base produtos").Select
x = Range("A" & Rows.Count).End(xlUp).Row
y = Range("K" & Rows.Count).End(xlUp).Row
Range("K1").Copy
Range("K3").PasteSpecial
Range("K3:K" & x).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'Copia e desduplica nomes dos produtos da base de faturamento e base de produtos

ActiveWorkbook.Sheets.Add
ActiveSheet.Name = "Corrigir"
Range("A1").Value = "PRODUTOS BASE"
Sheets("Base faturamento").Select
Range("AZ3:AZ" & Range("AZ3").End(xlDown).Row).Select
Selection.Copy
Sheets("Corrigir").Select
Range("A2").PasteSpecial
Sheets("Base produtos").Select
Range("L3:L" & Range("L3").End(xlDown).Row).Select
Selection.Copy
Sheets("Corrigir").Select
Range("A1").End(xlDown).Offset(1, 0).Select
Selection.PasteSpecial
ActiveSheet.Columns("$A:$A").RemoveDuplicates Columns:=1, Header:= _
        xlYes
Range("B1").Value = "PROCV"
Range("B2").FormulaLocal = "=PROCV(A2;Análises!B:B;1;0)"
y = Range("A" & Rows.Count).End(xlUp).Row
Range("B2:B" & y).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Range("B1").Select
Selection.AutoFilter
Columns.AutoFit
tempofinal = (Timer - tempoinicial) / 60
MsgBox ("Parte 1 finalizada! Verificar nomes com #N/D." & Chr(13) & "Tempo: " & Round(tempofinal, 2) & " minutos")

Application.CutCopyMode = False

End Sub

Sub Atualizar_2()

Dim x, y As Long, k As Byte

tempoinicial = Timer

On Error GoTo erro

Application.EnableEvents = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False

k = InputBox("Houve correções de nomes de produtos? ( 1-Sim / 0-Não )")
If k = 1 Then

'Arrasta formulas Subgr fatu

Sheets("Subgr fatu").Select
y = Range("A3").End(xlDown).Row
Range("J2").Copy
Range("J4").PasteSpecial
Range("J4:J" & y).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'Arrasta formulas Subgr prod

Sheets("Subgr prod").Select
y = Range("A3").End(xlDown).Row
Range("J2").Copy
Range("J4").PasteSpecial
Range("J4:J" & y).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

'Arrasta formula subgrupo na base de faturamento

Sheets("Base faturamento").Select
x = Range("A" & Rows.Count).End(xlUp).Row
y = Range("AX" & Rows.Count).End(xlUp).Row
Range("AY1").Copy
Range("AY3").PasteSpecial
Range("AY3:AY" & x).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'Arrasta formula subgrupo na base de produtos

Sheets("Base produtos").Select
x = Range("A" & Rows.Count).End(xlUp).Row
y = Range("J" & Rows.Count).End(xlUp).Row
Range("K1").Copy
Range("K3").PasteSpecial
Range("K3:K" & x).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Else
End If
    
'Atualiza tabela dinamica Grupo fatu e arrasta formulas
    
Sheets("Grupo fatu").Select
With ActiveSheet.PivotTables("Grupo fatu")
        .PivotCache.Refresh
        .PivotFields("NOME_SUBGRUPOPRODUTO").ClearAllFilters
        .PivotFields("NOME_SUBGRUPOPRODUTO").PivotItems("E/").Visible = False
End With
x = Range("A" & Rows.Count).End(xlUp).Row
y = Range("Z" & Rows.Count).End(xlUp).Row
Range("Z5:Z" & y).Clear
Range("Z3").Copy
Range("Z5").PasteSpecial
Range("Z5:Z" & x).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

'Atualiza tabela dinamica Grupo prod e arrasta formulas
    
Sheets("Grupo prod").Select
With ActiveSheet.PivotTables("Grupo prod")
        .PivotCache.Refresh
        .PivotFields("NOME_SUBGRUPOPRODUTO").ClearAllFilters
        .PivotFields("NOME_SUBGRUPOPRODUTO").PivotItems("E/").Visible = False
End With
x = Range("A" & Rows.Count).End(xlUp).Row
y = Range("Z" & Rows.Count).End(xlUp).Row
Range("Z5:Z" & y).Clear
Range("Z3").Copy
Range("Z5").PasteSpecial
Range("Z5:Z" & x).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

'Arrasta formula grupo na base de faturamento

Sheets("Base faturamento").Select
x = Range("A" & Rows.Count).End(xlUp).Row
y = Range("AX" & Rows.Count).End(xlUp).Row
Range("AX1").Copy
Range("AX3").PasteSpecial
Range("AX3:AX" & x).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

'Arrasta formula produto,BDBE,check,E/,nome ajustado na base de faturamento

Range("AZ1:BD1").Copy
Range("AZ3:BD3").PasteSpecial
Range("AZ3:BD" & x).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'Arrasta formula cont.se,udm na coluna inteira na base de faturamento

y = Range("A" & Rows.Count).End(xlUp).Row
Range("BE1").FormulaLocal = "=1/CONT.SES($BD$1:$BD$" & y & ";BD1;$G$1:$G$" & y & ";G1)"
Range("BE1:BF1").Copy
Range("BE3:BF3").PasteSpecial
Range("BE3:BF" & x).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'Arrasta formula grupo na base de produtos

Sheets("Base produtos").Select
x = Range("A" & Rows.Count).End(xlUp).Row
y = Range("J" & Rows.Count).End(xlUp).Row
Range("J1").Copy
Range("J3").PasteSpecial
Range("J3:J" & x).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

'Arrasta formula produto,E/,mês,ano,UDM ajustado na base de produtos

Range("L1:P1").Copy
Range("L3:P3").PasteSpecial
Range("L3:P" & x).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'Atualiza tabela dinamica Tab. dinâmica e arrasta formulas
    
Sheets("Tab. dinâmica").Select
ActiveSheet.PivotTables("Tab. dinâmica").PivotCache.Refresh
Rows(3).Clear
x = Range("A6").End(xlToRight).Column
For y = x To x - 11 Step -1
    Cells(3, y) = "UDM"
Next y
x = Range("A" & Rows.Count).End(xlUp).Row
y = Range("CA" & Rows.Count).End(xlUp).Row
Range("CA7:CH" & y).Clear
Range("CA5:CH5").Copy
Range("CA7:CH7").PasteSpecial
Range("CA7:CH" & x).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

'Limpa "Análises", cola nomes novos e deleta "corrigir"

Sheets("Análises").Select
y = Range("A" & Rows.Count).End(xlUp).Row
Range("A5:T" & y).Clear
Sheets("Corrigir").Select
ActiveSheet.ShowAllData
y = Range("A" & Rows.Count).End(xlUp).Row
Range("A2:B" & y).Clear
Sheets("Base faturamento").Select
Range("AZ3:AZ" & Range("AZ3").End(xlDown).Row).Select
Selection.Copy
Sheets("Corrigir").Select
Range("A2").PasteSpecial
Sheets("Base produtos").Select
Range("L3:L" & Range("L3").End(xlDown).Row).Select
Selection.Copy
Sheets("Corrigir").Select
Range("A1").End(xlDown).Offset(1, 0).Select
Selection.PasteSpecial
ActiveSheet.Columns("$A:$A").RemoveDuplicates Columns:=1, Header:= _
        xlYes
ActiveSheet.ShowAllData
y = Range("A" & Rows.Count).End(xlUp).Row
Range("B1").Value = "Outros"
Range("B2").FormulaLocal = "=SE($B$1=A2;LIN(A2);0)"
Range("B2:B" & y).Select
Selection.FillDown
Calculate
L = WorksheetFunction.Sum(Range("B2:B" & y))
Rows(L).Select
Selection.Delete
ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
     "Corrigir!R1C1:R" & y & "C2", Version:=xlPivotTableVersion14).CreatePivotTable _
     TableDestination:="Corrigir!R1C5", TableName:="Tabela dinâmica1", _
     DefaultVersion:=xlPivotTableVersion14
With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("PRODUTOS BASE")
     .Orientation = xlRowField
     .Position = 1
End With
y = Range("E" & Rows.Count).End(xlUp).Row - 2
Range("E2:E" & y).Copy
Sheets("Análises").Select
Range("B5").PasteSpecial
Sheets("Corrigir").Delete

Sheets("Análises").Select
y = Range("B" & Rows.Count).End(xlUp).Row
Range("A3").Copy
Range("A5").Select
ActiveSheet.Paste
Range("A5:A" & y).Select
Selection.FillDown
Calculate
Range("C3:T3").Copy
Range("C5:T5").Select
ActiveSheet.Paste
Range("C5:T" & y).Select
Selection.FillDown
Calculate
Range("A5:T" & y).Select
Range("A5:T" & y).Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Range("H2").FormulaLocal = "=SOMASES('Base faturamento'!$X:$X;'Base faturamento'!$BB:$BB;1;'Base faturamento'!$G:$G;H$4)-SOMA(H5:H" & y & ")"
Range("H2:L2").Select
Selection.FillRight
Range("M2").FormulaLocal = "=SOMASES('Base faturamento'!$X:$X;'Base faturamento'!$BB:$BB;1;'Base faturamento'!$BF:$BF;M$4)-SOMA(M5:M" & y & ")"
Calculate
ActiveWorkbook.RefreshAll

Application.CutCopyMode = False
Sheets("Dashboard").Select
Range("C82:E82").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Análises!$B$5:$B$" & y
    End With
Range("A1").Select
tempofinal = (Timer - tempoinicial) / 60
MsgBox ("Atualização finalizada com sucesso!" & Chr(13) & "Tempo: " & Round(tempofinal, 2) & " minutos")

Exit Sub
erro: Resume Next

End Sub
