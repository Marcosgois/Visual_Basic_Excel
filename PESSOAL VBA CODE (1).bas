Attribute VB_Name = "Módulo1"
Public Sub changeColor()
    Dim linha, coluna, limite, i, j As Integer
    Dim MyRange As Range
    Dim restante As Double
    Dim Biggers(1 To 10), number As Long
    Dim planilha
    Set planilha = Worksheets(ActiveSheet.Name)
    linha = 36 '36
    coluna = 30 '30
    limite = 57 '57
    'Set MyRange = Range("D3", "H3")
    
    Do While linha < 199 '199
        j = 1
        For contador = coluna To limite Step 3
            Biggers(j) = Abs(Cells(linha, contador).Value)
            'MsgBox ("Cont: " & contador & " Limite: " & limite & " Biggers: " & Biggers(j) & " J: " & j)
            j = j + 1
        Next
        Call QuickSort(Biggers, 1, UBound(Biggers))
        'MsgBox (Biggers(1) & " " & Biggers(2) & " " & Biggers(3) & " " & Biggers(4) & " " & Biggers(5) & " " & Biggers(6) & " " & Biggers(7) & " " & Biggers(8) & " " & Biggers(9) & " " & Biggers(10))
        i = 10
        Value = Abs(planilha.Range("G" & linha))
        testedValue = Abs(planilha.Range("Y" & linha))
        restante = ((0.8 * Value) - testedValue)
        'maior = Application.WorksheetFunction.Large(Range(Cells(linha, coluna), Cells(linha, coluna + 4)), i)
        'MsgBox ("Restante - Maior: " & restante & "-" & Biggers(i))
        Do While (restante > 0)
            contador = 0
            'MsgBox (i)
            If Biggers(i) <> 0 Then
                For contador = coluna To limite Step 3
                    'number = Cells(linha, contador).Value
                    'MsgBox (Biggers(i) & " é = ? " & Abs(Cells(linha, contador).Value))
                    If Abs(Cells(linha, contador).Value) = Biggers(i) Then
                        Cells(linha, contador).Interior.ColorIndex = 3
                    End If
                Next
                restante = restante - Biggers(i)
            End If
            If restante < 0 Then
                Exit Do
            End If
            i = i - 1
            If i < 1 Then
                Exit Do
            End If
            'maior = Application.WorksheetFunction.Large(Range(Cells(linha, coluna), Cells(linha, coluna + 4)), i)
        Loop
            linha = linha + 1
    Loop
End Sub

Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)

  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub

