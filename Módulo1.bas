Attribute VB_Name = "Módulo1"
Sub actionMacro()
Dim x As Integer
x = 13 'O
    While x < 500
        Columns(x).Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ActiveSheet.Range(Cells(6, x - 1), Cells(6, x)).Select
        Selection.Merge
        Cells(7, x).Select
        ActiveCell.FormulaR1C1 = "VALOR CORRETO"
        Cells(8, x).Select
        ActiveCell.FormulaR1C1 = _
            "=IFERROR(VLOOKUP(VLOOKUP(R6C" + CStr(x - 1) + ",Planilha2!R2C1:R173C2,2,FALSE),Planilha3!R39C3:R51C7,2,FALSE),"""")"
        Selection.AutoFill Destination:=ActiveSheet.Range(Cells(8, x), Cells(11, x)), Type:=xlFillDefault
        Cells(12, x).Select
        ActiveCell.FormulaR1C1 = _
            "=IFERROR(VLOOKUP(VLOOKUP(R6C" + CStr(x - 1) + ",Planilha2!R2C1:R173C2,2,FALSE),Planilha3!R39C3:R51C7,3,FALSE),"""")"
        Selection.AutoFill Destination:=ActiveSheet.Range(Cells(12, x), Cells(23, x)), Type:=xlFillDefault
        Cells(24, x).Select
        ActiveCell.FormulaR1C1 = _
            "=IFERROR(VLOOKUP(VLOOKUP(R6C" + CStr(x - 1) + ",Planilha2!R2C1:R173C2,2,FALSE),Planilha3!R39C3:R51C7,4,FALSE),"""")"
        Selection.AutoFill Destination:=ActiveSheet.Range(Cells(24, x), Cells(35, x)), Type:=xlFillDefault
        Cells(36, x).Select
        ActiveCell.FormulaR1C1 = _
            "=IFERROR(VLOOKUP(VLOOKUP(R6C" + CStr(x - 1) + ",Planilha2!R2C1:R173C2,2,FALSE),Planilha3!R39C3:R51C7,5,FALSE),"""")"
        Selection.AutoFill Destination:=ActiveSheet.Range(Cells(36, x), Cells(38, x)), Type:=xlFillDefault
        Cells(39, x - 1).Select
        ActiveCell.FormulaR1C1 = "=SUM(R[-31]C:R[-1]C)"
        Cells(39, x - 1).Select
        Selection.AutoFill Destination:=ActiveSheet.Range(Cells(39, x - 1), Cells(39, x)), Type:=xlFillDefault
        Cells(40, x - 1).Select
        ActiveCell.FormulaR1C1 = "Diferença"
        Cells(40, x).Select
        ActiveCell.FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
        x = x + 2
    Wend
End Sub
