Attribute VB_Name = "Módulo1"
Sub Macro2()
Dim x As Integer
x = 15
    While x < 200
        Columns(x).Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        x = x + 2
        MsgBox (x)
    Wend
End Sub

Sub Macro3()
Dim x As Integer
    Worksheets("Planilha1").Activate
    x = 15
    Columns(x).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveSheet.Range(Cells(6, x - 1), Cells(6, x)).Select
    Selection.Merge
    Cells(7, x).Select
    ActiveCell.FormulaR1C1 = "VALOR CORRETO"
    Cells(8, x).Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(VLOOKUP(R6C14,Planilha2!R2C1:R173C2,2,FALSE),Planilha3!R39C3:R51C7,2,FALSE),"""")"
    Cells(8, x).Select
    Worksheets("Planilha1").Activate
    Selection.AutoFill Destination:=ActiveSheet.Range(Cells(8, x), Cells(38, x)), Type:=xlFillDefault
    Cells(39, x - 1).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-31]C:R[-1]C)"
    Cells(39, x - 1).Select
    Selection.AutoFill Destination:=ActiveSheet.Range(Cells(39, x - 1), Cells(39, x)), Type:=xlFillDefault
    Cells(40, x - 1).Select
    ActiveCell.FormulaR1C1 = "Diferença"
    Cells(40, x).Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
End Sub

