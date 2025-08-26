Sub Atualizar_1()

Dim t As Byte, k As Integer, y As Long

tempoinicial = Timer

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.StatusBar = False

On Error GoTo erro
Path = ThisWorkbook.Path


Sheets("Setores").Select
Range("A3:AK" & Rows.Count).Clear


'Tapeçaria

c = 1
For k = 1 To 23
    If k = 2 Or k = 3 Or k = 4 Or k = 5 Or k = 6 Or k = 9 _
    Or k = 12 Or k = 15 Or k = 18 Or k = 21 Then
        Workbooks.Open (Path & "\V6 Tapecaria.xlsb")
        Sheets("Base Células").Select
        y = Range("D" & Rows.Count).End(xlUp).Row
        Range(Cells(5, k), Cells(y, k)).Copy
        Workbooks.Open (Path & "\IVA e Eficiência V7.xlsb")
        Sheets("Setores").Select
        y = Range("A" & Rows.Count).End(xlUp).Row + 1
        Cells(y, c + 1).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        c = c + 1
    End If
Next k
c = 31
For k = 31 To 36
    If k = 31 Or k = 34 Or k = 35 Then
        Workbooks.Open (Path & "\V6 Tapecaria.xlsb")
        Sheets("Base Células").Select
        y = Range("D" & Rows.Count).End(xlUp).Row
        Range(Cells(5, k), Cells(y, k)).Copy
        Workbooks.Open (Path & "\IVA e Eficiência V7.xlsb")
        Sheets("Setores").Select
        y = Range("A" & Rows.Count).End(xlUp).Row + 1
        Cells(y, c).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        c = c + 1
    End If
Next k
y = Range("D" & Rows.Count).End(xlUp).Row
x = Range("A" & Rows.Count).End(xlUp).Row + 1
Range(Cells(x, 1), Cells(y, 1)).Value = "Tapeçaria"
Workbooks.Open (Path & "\V6 Tapecaria.xlsb")
ActiveWorkbook.Close

'Laminação

c = 1
For k = 1 To 23
    If k = 2 Or k = 3 Or k = 4 Or k = 5 Or k = 6 Or k = 9 _
    Or k = 12 Or k = 15 Or k = 18 Or k = 21 Then
        Workbooks.Open (Path & "\V6 Laminacao.xlsb")
        Sheets("Base Células").Select
        y = Range("D" & Rows.Count).End(xlUp).Row
        Range(Cells(5, k), Cells(y, k)).Copy
        Workbooks.Open (Path & "\IVA e Eficiência V7.xlsb")
        Sheets("Setores").Select
        y = Range("A" & Rows.Count).End(xlUp).Row + 1
        Cells(y, c + 1).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        c = c + 1
    End If
Next k
c = 31
For k = 31 To 36
    If k = 31 Or k = 34 Or k = 35 Then
        Workbooks.Open (Path & "\V6 Laminacao.xlsb")
        Sheets("Base Células").Select
        y = Range("D" & Rows.Count).End(xlUp).Row
        Range(Cells(5, k), Cells(y, k)).Copy
        Workbooks.Open (Path & "\IVA e Eficiência V7.xlsb")
        Sheets("Setores").Select
        y = Range("A" & Rows.Count).End(xlUp).Row + 1
        Cells(y, c).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        c = c + 1
    End If
Next k
y = Range("D" & Rows.Count).End(xlUp).Row
x = Range("A" & Rows.Count).End(xlUp).Row + 1
Range(Cells(x, 1), Cells(y, 1)).Value = "Laminação"
Workbooks.Open (Path & "\V6 Laminacao.xlsb")
ActiveWorkbook.Close

'Embalagem

c = 1
For k = 1 To 23
    If k = 2 Or k = 3 Or k = 4 Or k = 5 Or k = 6 Or k = 9 _
    Or k = 12 Or k = 15 Or k = 18 Or k = 21 Then
        Workbooks.Open (Path & "\V6 Embalagem.xlsb")
        Sheets("Base Celula").Select
        y = Range("D" & Rows.Count).End(xlUp).Row
        Range(Cells(5, k), Cells(y, k)).Copy
        Workbooks.Open (Path & "\IVA e Eficiência V7.xlsb")
        Sheets("Setores").Select
        y = Range("A" & Rows.Count).End(xlUp).Row + 1
        Cells(y, c + 1).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        c = c + 1
    End If
Next k
c = 31
For k = 31 To 36
    If k = 31 Or k = 34 Or k = 35 Then
        Workbooks.Open (Path & "\V6 Embalagem.xlsb")
        Sheets("Base Celula").Select
        y = Range("D" & Rows.Count).End(xlUp).Row
        Range(Cells(5, k), Cells(y, k)).Copy
        Workbooks.Open (Path & "\IVA e Eficiência V7.xlsb")
        Sheets("Setores").Select
        y = Range("A" & Rows.Count).End(xlUp).Row + 1
        Cells(y, c).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        c = c + 1
    End If
Next k
y = Range("D" & Rows.Count).End(xlUp).Row
x = Range("A" & Rows.Count).End(xlUp).Row + 1
Range(Cells(x, 1), Cells(y, 1)).Value = "Embalagem"
Workbooks.Open (Path & "\V6 Embalagem.xlsb")
ActiveWorkbook.Close

'Montagem

c = 1
For k = 1 To 23
    If k = 2 Or k = 3 Or k = 4 Or k = 5 Or k = 6 Or k = 9 _
    Or k = 12 Or k = 15 Or k = 18 Or k = 21 Then
        Workbooks.Open (Path & "\V6 Montagem.xlsb")
        Sheets("Base Células").Select
        y = Range("D" & Rows.Count).End(xlUp).Row
        Range(Cells(5, k), Cells(y, k)).Copy
        Workbooks.Open (Path & "\IVA e Eficiência V7.xlsb")
        Sheets("Setores").Select
        y = Range("A" & Rows.Count).End(xlUp).Row + 1
        Cells(y, c + 1).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        c = c + 1
    End If
Next k
c = 31
For k = 31 To 36
    If k = 31 Or k = 34 Or k = 35 Then
        Workbooks.Open (Path & "\V6 Montagem.xlsb")
        Sheets("Base Células").Select
        y = Range("D" & Rows.Count).End(xlUp).Row
        Range(Cells(5, k), Cells(y, k)).Copy
        Workbooks.Open (Path & "\IVA e Eficiência V7.xlsb")
        Sheets("Setores").Select
        y = Range("A" & Rows.Count).End(xlUp).Row + 1
        Cells(y, c).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        c = c + 1
    End If
Next k
y = Range("D" & Rows.Count).End(xlUp).Row
x = Range("A" & Rows.Count).End(xlUp).Row + 1
Range(Cells(x, 1), Cells(y, 1)).Value = "Montagem"
Workbooks.Open (Path & "\V6 Montagem.xlsb")
ActiveWorkbook.Close


'Espumação

c = 1
For k = 1 To 23
    If k = 2 Or k = 3 Or k = 4 Or k = 5 Or k = 6 Or k = 9 _
    Or k = 12 Or k = 15 Or k = 18 Or k = 21 Then
        Workbooks.Open (Path & "\V6 Espumacao.xlsb")
        Sheets("Base Células").Select
        y = Range("D" & Rows.Count).End(xlUp).Row
        Range(Cells(5, k), Cells(y, k)).Copy
        Workbooks.Open (Path & "\IVA e Eficiência V7.xlsb")
        Sheets("Setores").Select
        y = Range("A" & Rows.Count).End(xlUp).Row + 1
        Cells(y, c + 1).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        c = c + 1
    End If
Next k
c = 31
For k = 31 To 36
    If k = 31 Or k = 34 Or k = 35 Then
        Workbooks.Open (Path & "\V6 Espumacao.xlsb")
        Sheets("Base Células").Select
        y = Range("D" & Rows.Count).End(xlUp).Row
        Range(Cells(5, k), Cells(y, k)).Copy
        Workbooks.Open (Path & "\IVA e Eficiência V7.xlsb")
        Sheets("Setores").Select
        y = Range("A" & Rows.Count).End(xlUp).Row + 1
        Cells(y, c).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        c = c + 1
    End If
Next k
y = Range("D" & Rows.Count).End(xlUp).Row
x = Range("A" & Rows.Count).End(xlUp).Row + 1
Range(Cells(x, 1), Cells(y, 1)).Value = "Espumação"
Workbooks.Open (Path & "\V6 Espumacao.xlsb")
ActiveWorkbook.Close

'Costura

c = 1
For k = 1 To 48
    If k = 2 Or k = 3 Or k = 4 Or k = 5 Or k = 6 Or k = 9 _
    Or k = 12 Or k = 15 Or k = 18 Or k = 21 Or k = 24 _
    Or k = 27 Or k = 30 Or k = 33 Or k = 36 Or k = 39 _
    Or k = 42 Or k = 45 Or k = 48 Then
        Workbooks.Open (Path & "\V6 Costura.xlsb")
        Sheets("Base Células").Select
        y = Range("D" & Rows.Count).End(xlUp).Row
        Range(Cells(5, k), Cells(y, k)).Copy
        Workbooks.Open (Path & "\IVA e Eficiência V7.xlsb")
        Sheets("Setores").Select
        y = Range("A" & Rows.Count).End(xlUp).Row + 1
        Cells(y, c + 1).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        c = c + 1
    End If
Next k
c = 31
For k = 67 To 72
    If k = 67 Or k = 70 Or k = 71 Then
        Workbooks.Open (Path & "\V6 Costura.xlsb")
        Sheets("Base Células").Select
        y = Range("D" & Rows.Count).End(xlUp).Row
        Range(Cells(5, k), Cells(y, k)).Copy
        Workbooks.Open (Path & "\IVA e Eficiência V7.xlsb")
        Sheets("Setores").Select
        y = Range("A" & Rows.Count).End(xlUp).Row + 1
        Cells(y, c).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        c = c + 1
    End If
Next k
y = Range("D" & Rows.Count).End(xlUp).Row
x = Range("A" & Rows.Count).End(xlUp).Row + 1
Range(Cells(x, 1), Cells(y, 1)).Value = "Costura"
Workbooks.Open (Path & "\V6 Costura.xlsb")
ActiveWorkbook.Close



Sheets("Setores").Select
Columns("D:D").Select
Selection.NumberFormat = "m/d/yyyy"
y = Range("A" & Rows.Count).End(xlUp).Row
Range("AH1:AK1").Copy
Range("AH3:AK3").Select
ActiveSheet.Paste
Range("AH3:AK" & y).Select
Selection.FillDown
Calculate
Range("AH3:AK" & y).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Tempos PUXAR TEMPOS IVA

Workbooks.Open (Path & "\Tempos IVA.xlsb")
Sheets("Tempos IVA").Select
x = Range("A" & Rows.Count).End(xlUp).Row - 1
Range("A2:G" & x).Copy
Workbooks.Open (Path & "\IVA e Eficiência V7.xlsb")
Sheets("Tab. Dinâmica").Select
Range("B3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Workbooks.Open (Path & "\Tempos IVA.xlsb")
ActiveWorkbook.Close
Workbooks.Open (Path & "\IVA e Eficiência V7.xlsb")

'Arrasta formulas na base de faturamento

Sheets("Base faturamento").Select
Range("AX3:BJ" & Rows.Count).Clear
x = Range("A" & Rows.Count).End(xlUp).Row
y = Range("AX" & Rows.Count).End(xlUp).Row
Range("AX1:BJ1").Copy
Range("AX" & y + 1).PasteSpecial
Range("AX" & y + 1 & ":BJ" & x).Select
Selection.FillDown
Calculate
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

Application.CutCopyMode = False
'ActiveWorkbook.RefreshAll
Sheets("Dashboard").Select
Range("A1").Select
tempofinal = (Timer - tempoinicial) / 60
MsgBox ("Atualização finalizada com sucesso!" & Chr(13) & "Tempo: " & Round(tempofinal, 2) & " minutos")

Exit Sub
erro: Resume Next
End Sub
