Attribute VB_Name = "mainCode"
Option Explicit
Public Sub Normal(Optional CMD As Variant)
    Dim numeroRegistros, linha As Integer
    Dim nomeArquivo, nome, dia, mes, ano, siape, valor, contrato As String
    Dim comando As Integer
    Dim arquivo, planilha
    Set planilha = Worksheets(ActiveSheet.Name)
    
    dia = planilha.Range("K12")
    mes = planilha.Range("K12")
    ano = planilha.Range("L12")
    
    nomeArquivo = ActiveWorkbook.Path & "\CONSIGSIAPE" & Format(dia, "00") & Format(mes, "00") & Format(ano, "0000") & "100908.txt"
    If (MsgBox("Confirma a geração do arquivo " & nomeArquivo & " ?", vbYesNo, "") <> vbYes) Then
        Exit Sub
    End If
    
    Set arquivo = CreateObject("Scripting.FileSystemObject").createTextFile(nomeArquivo, True)
    
    ' HEADER           Tipo   Órgão    Constante_____    Mês_____________    Ano_______________   Constante   Constante_____     CNPJ___________    Constante____  Filler_________
    arquivo.writeline ("0" & "45206" & String(16, "0") & Format(mes, "00") & Format(ano, "0000") & "AFIPEA" & String(15, " ") & "01264183000115" & "CONSIGSIAPE" & String(53, " "))
        
    
    numeroRegistros = 0
    linha = 6
    nome = planilha.Range("B" & linha)
    
    
    While (Trim(nome) <> "")
        numeroRegistros = numeroRegistros + 1
        If CMD = 4 Or CMD = 3 Then
        contrato = String(16, " ")
        valor = 0
        siape = planilha.Range("A" & linha)
        Else
        contrato = planilha.Range("D" & linha)
        siape = planilha.Range("A" & linha)
        valor = planilha.Range("C" & linha)
        End If
        
        If CMD = 4 Then
        ' CONSIGNAÇÃO           Tipo   Órgão    Matrícula______________________   DV    CMD   IND   Rúbrica   SEQ   Valor_______________________________________    Prazo   Num. origem__    Mês     Ano      Data Antg.   hhmmss    Filler_______    MF    Núm Contrato
        arquivo.writeline (Left("1" & "45206" & Left(siape & String(7, "0"), 7) & "0" & CMD & "2" & "34642" & "1" & String(28, "0"), 28) & Format(valor, "00000") & "000" & String(8, "0") & "00" & "0000" & "00000000" & "000000" & String(5, " ") & "8" & Format((linha - 5), String(20, "0")) & String(37, "0"))
        Else
        ' CONSIGNAÇÃO      Tipo   Órgão    Matrícula______________    DV    CMD   IND   Rúbrica   SEQ   Valor____________________________   Prazo   Num. origem__    Mês     Ano      Data Antg.   hhmmss    Filler_______    MF    Número Contrato___________________
        arquivo.writeline ("1" & "45206" & Format(siape, "0000000") & "0" & CMD & "2" & "34642" & "1" & "000000" & Format(valor, "00000") & "000" & String(8, "0") & "00" & "0000" & "00000000" & "000000" & String(5, " ") & "8" & Left(contrato & String(9, " "), 16) & String(4, " ") & String(37, "0"))
        End If
        linha = linha + 1
        nome = planilha.Range("B" & linha)
    Wend
    
    ' TRAILLER         TIPO   SIAPE    CONSTANTE_____    REGISTROS________________________      FILLER
    arquivo.writeline ("9" & "45206" & String(16, "9") & Format(numeroRegistros, "0000000") & String(98, " "))
    arquivo.Close
    
    Dim msg

    msg = msg & "Arquivo " & nomeArquivo & " gerado com sucesso! "
    msg = msg & vbCrLf & vbCrLf & _
        "Qtde de registros: " & numeroRegistros
    MsgBox (msg)
Exit_inclusao:
    Exit Sub
    Resume
    
End Sub


Sub Pensionista(Optional CMD As Variant)
    Dim numeroRegistros, linha As Integer
    Dim nomeArquivo, nome, dia, mes, ano, siape, MAT_INST, MAT_BENEF, valor, contrato As String
    Dim arquivo, planilha
    
    Set planilha = Worksheets(ActiveSheet.Name)
    
    dia = planilha.Range("H12")
    mes = planilha.Range("I12")
    ano = planilha.Range("J12")
    
    On Error GoTo Err_inclusao
    
    nomeArquivo = ActiveWorkbook.Path & "\CONSIGSIAPE_PENS" & Format(dia, "00") & Format(mes, "00") & Format(ano, "0000") & "101840.txt"
    If (MsgBox("Confirma a geração do arquivo " & nomeArquivo & " ?", vbYesNo, "") <> vbYes) Then
        Exit Sub
    End If
    
    Set arquivo = CreateObject("Scripting.FileSystemObject").createTextFile(nomeArquivo, True)
    ' HEADER           Tipo   Órgão    Constante_____    Mês_____________    Ano_______________   Constante   Constante____    CNPJ___________    Constante____   Filler_________
    arquivo.writeline ("0" & "45206" & String(23, "0") & Format(mes, "00") & Format(ano, "0000") & "AFIPEA" & String(8, " ") & "01264183000115" & "CONSIG-PENS" & String(53, " "))
    
    numeroRegistros = 0
    linha = 6
    nome = planilha.Range("B" & linha)
    
    While (Trim(nome) <> "")
        numeroRegistros = numeroRegistros + 1
        
        MAT_INST = planilha.Range("A" & linha)
        MAT_BENEF = planilha.Range("B" & linha)
        valor = planilha.Range("D" & linha)
        contrato = planilha.Range("E" & linha)
        
        If CMD = 4 Then
        ' CONSIGNAÇÃO      Tipo   Órgão    Matrícula Instituidor_____    Matrícula Beneficiário______    CMD   IND   Rúbrica   SEQ   Valor___________________________    Prazo   Mês     Ano     Data Antg.    hhmmss    Filler_______    MF    Filler________
        arquivo.writeline ("1" & "45206" & Format(MAT_INST, "0000000") & Format(MAT_BENEF, "00000000") & CMD & "2" & "30893" & "1" & "000000" & Format(valor, "00000") & "000" & "00" & "0000" & "00000000" & "000000" & String(5, " ") & "8" & Format((linha - 5), String(20, "0")) & String(37, "0"))
        Else
        ' CONSIGNAÇÃO      Tipo   Órgão    Matrícula Instituidor_____    Matrícula Beneficiário______    CMD   IND   Rúbrica   SEQ   Valor___________________________    Prazo   Mês     Ano     Data Antg.    hhmmss    Filler_______    MF                                          Filler________
        arquivo.writeline ("1" & "45206" & Format(MAT_INST, "0000000") & Format(MAT_BENEF, "00000000") & CMD & "2" & "30893" & "1" & "000000" & Format(valor, "00000") & "000" & "00" & "0000" & "00000000" & "000000" & String(5, " ") & "8" & Left(contrato & String(13, " "), 20) & String(38, "0"))
        End If
        
        linha = linha + 1
        nome = planilha.Range("C" & linha)
    Wend
    
    ' TRAILLER         TIPO   SIAPE    CONSTANTE_____    REGISTROS________________________      FILLER
    arquivo.writeline ("9" & "45206" & String(23, "9") & Format(numeroRegistros, "0000000") & String(91, " "))
    arquivo.Close
    
    Dim msg

    msg = msg & "Arquivo " & nomeArquivo & " gerado com sucesso! "
    msg = msg & vbCrLf & vbCrLf & _
        "Qtde de registros: " & numeroRegistros
    MsgBox (msg)
    
Exit_inclusao:
    Exit Sub
    Resume
    
Err_inclusao:
    Err.Raise vbObjectError + 513, ActiveSheet.Name, Err.Description
End Sub

Public Sub txtToExcel()
    Dim Ficheiro As String
    Ficheiro = ActiveWorkbook.Path + "C:\Users\goisi\Documents\AFIPEA\txt to excel\relanaliticoD8012017.txt"
    
    Dim rg As Range
    Set rg = Range("A1")
    
    Open Ficheiro For Input As #1
    
    Dim S As String, N As Integer, C As Integer, X As Variant
    Do Until EOF(1)
        Line Input #1, S
        C = 0
        X = Split(S, " ")
        For N = 0 To UBound(X)
            If X(N) <> "" Then
                rg.Offset(0, C) = X(N)
                C = C + 1
            End If
        Next N
        Set rg = rg.Offset(1, 0)
    Loop
    
    Close #1
End Sub

Public Sub ImportarTexto()
    Dim Ficheiro, fileName As String
    Dim planilha
    
    Set planilha = Worksheets(ActiveSheet.Name)
    fileName = planilha.Range("F14")
    
    Ficheiro = ActiveWorkbook.Path + "\" + fileName + ".txt" 'relanaliticoD8012017
    MsgBox (Ficheiro)
    Dim rg As Range
    Set rg = Range("A3")
    
    Open Ficheiro For Input As #1
    
    Dim S As String
    Do Until EOF(1)
        Line Input #1, S
        'rg = Left(S, 5)                            'Código do Órgão
        rg.Offset(0, 0) = Mid(S, 6, 7)              'Matrícula do Servidor
        'rg.Offset(0, 2) = Mid(S, 13, 9)
        'rg.Offset(0, 3) = Mid(S, 22, 2)
        rg.Offset(0, 1) = Mid(S, 24, 50)            'Nome do Servidor
        'rg.Offset(0, 5) = Mid(S, 74, 11)
        'rg.Offset(0, 6) = Mid(S, 85, 5)
        'rg.Offset(0, 7) = Mid(S, 90, 1)
        rg.Offset(0, 2) = CDbl(Mid(S, 91, 11) / 100) 'Valor da Rúbrica
        'rg.Offset(0, 9) = Mid(S, 102, 3)
        'rg.Offset(0, 10) = Mid(S, 105, 6)
        'rg.Offset(0, 11) = Mid(S, 111, 12)
        rg.Offset(0, 3) = Format(Mid(S, 123, 20), "0###############-###")            'Número do contrato
        
        'rg.Offset(0, 13) = Mid(S, 143, 2)
        'rg.Offset(0, 14) = Mid(S, 145, 5)
        
        Set rg = rg.Offset(1, 0)
    Loop
    
    Close #1
    MsgBox ("Importação do arquivo """ + fileName + """ com sucesso!")
End Sub

Public Sub ImportarTextoRejeitado(Optional TIPO As Variant)
    Dim Ficheiro, fileName As String
    Dim planilha
    
    Set planilha = Worksheets(ActiveSheet.Name)
    fileName = planilha.Range("F15")
    
    Ficheiro = ActiveWorkbook.Path + "\" + fileName + ".txt"
    MsgBox (Ficheiro)
    Dim rg As Range
    Set rg = Range("A3")
    
    Open Ficheiro For Input As #1
    
    Dim S As String
    Do Until EOF(1)
        Line Input #1, S
        If TIPO = 1 Then    'PENSIONISTAS
        'rg = Left(S, 6)
        'rg.Offset(0, 0) = Mid(S, 7, 14)
        rg.Offset(0, 0) = Mid(S, 21, 5) 'Orgão
        rg.Offset(0, 1) = Mid(S, 26, 7) 'MatrículaIns
        rg.Offset(0, 2) = Mid(S, 33, 8) 'MatrículaBene
        'rg.Offset(0, 3) = Mid(S, 41, 1)
        'rg.Offset(0, 1) = Mid(S, 42, 1)
        rg.Offset(0, 3) = CDbl(Mid(S, 43, 11) / 100) 'Valor
        'rg.Offset(0, 6) = Mid(S, 54, 3)
        'rg.Offset(0, 7) = Mid(S, 57, 5)
        'rg.Offset(0, 2) = Mid(S, 62, 1)
        rg.Offset(0, 4) = Mid(S, 63, 60)   'Mensagem
        'rg.Offset(0, 10) = Mid(S, 123, 6)
        'rg.Offset(0, 11) = Mid(S, 129, 9)
        'rg.Offset(0, 3) = Mid(S, 138, 2)
        'rg.Offset(0, 13) = Mid(S, 140, 20)
        'rg.Offset(0, 14) = Mid(S, 160, 2)
        'rg.Offset(0, 14) = Mid(S, 162, 5)
        Else    'SERVIDORES
        'rg = Left(S, 6)
        'rg.Offset(0, 0) = Mid(S, 7, 14)
        rg.Offset(0, 0) = Mid(S, 21, 13) 'Matrícula
        'rg.Offset(0, 3) = Mid(S, 34, 1)
        'rg.Offset(0, 1) = Mid(S, 35, 1)
        rg.Offset(0, 1) = CDbl(Mid(S, 36, 11) / 100) 'Valor
        'rg.Offset(0, 6) = Mid(S, 47, 3)
        'rg.Offset(0, 7) = Mid(S, 50, 5)
        'rg.Offset(0, 2) = Mid(S, 55, 1)
        rg.Offset(0, 2) = Mid(S, 56, 60)   'Mensagem
        'rg.Offset(0, 10) = Mid(S, 116, 6)
        'rg.Offset(0, 11) = Mid(S, 122, 9)
        'rg.Offset(0, 3) = Mid(S, 131, 2)
        'rg.Offset(0, 13) = Mid(S, 193, 20)
        'rg.Offset(0, 14) = Mid(S, 153, 2)
        'rg.Offset(0, 14) = Mid(S, 155, 5)
        End If
        Set rg = rg.Offset(1, 0)
    Loop
    
    Close #1
    MsgBox ("Importação do arquivo """ + fileName + """ com sucesso!")
End Sub

Public Sub ImportarTextoAceito(Optional TIPO As Variant)
    Dim Ficheiro, fileName As String
    Dim planilha
    
    Set planilha = Worksheets(ActiveSheet.Name)
    fileName = planilha.Range("F15")
    
    Ficheiro = ActiveWorkbook.Path + "\" + fileName + ".txt"
    MsgBox (Ficheiro)
    Dim rg As Range
    Set rg = Range("A3")
    
    Open Ficheiro For Input As #1
    
    Dim S As String
    Do Until EOF(1)
        Line Input #1, S
        If TIPO = 1 Then
        'rg = Left(S, 6)
        'rg.Offset(0, 0) = Mid(S, 7, 14)
        rg.Offset(0, 0) = Mid(S, 21, 5) 'Orgão
        rg.Offset(0, 2) = Mid(S, 26, 7) 'MI
        rg.Offset(0, 3) = Mid(S, 33, 8) 'MB
        'rg.Offset(0, 6) = Mid(S, 41, 45)
        'rg.Offset(0, 7) = Mid(S, 86, 1)
        'rg.Offset(0, 2) = Mid(S, 87, 1)
        rg.Offset(0, 4) = CDbl(Mid(S, 88, 11) / 100) 'Valor
        'rg.Offset(0, 10) = Mid(S, 99, 3)
        'rg.Offset(0, 11) = Mid(S, 102, 5)
        'rg.Offset(0, 3) = Mid(S, 107, 1)
        rg.Offset(0, 5) = Mid(S, 108, 60)   'Mensagem
        'rg.Offset(0, 13) = Mid(S, 168, 6)
        'rg.Offset(0, 14) = Mid(S, 174, 9)
        'rg.Offset(0, 14) = Mid(S, 183, 2)
        'rg.Offset(0, 14) = Mid(S, 185, 20)
        'rg.Offset(0, 14) = Mid(S, 205, 2)
        'rg.Offset(0, 14) = Mid(S, 207, 5)
        Else
        'rg = Left(S, 6)
        'rg.Offset(0, 0) = Mid(S, 7, 14)
        rg.Offset(0, 0) = Mid(S, 21, 13) 'Matrícula
        'rg.Offset(0, 3) = Mid(S, 34, 60)
        'rg.Offset(0, 1) = Mid(S, 94, 1)
        'rg.Offset(0, 6) = Mid(S, 95, 1)
        rg.Offset(0, 1) = CDbl(Mid(S, 96, 11) / 100) 'Valor
        'rg.Offset(0, 10) = Mid(S, 107, 3)
        'rg.Offset(0, 11) = Mid(S, 110, 5)
        'rg.Offset(0, 3) = Mid(S, 115, 1)
        rg.Offset(0, 2) = Mid(S, 116, 60)   'Mensagem
        'rg.Offset(0, 13) = Mid(S, 176, 6)
        'rg.Offset(0, 14) = Mid(S, 182, 9)
        'rg.Offset(0, 14) = Mid(S, 191, 2)
        'rg.Offset(0, 14) = Mid(S, 193, 20)
        'rg.Offset(0, 14) = Mid(S, 213, 2)
        'rg.Offset(0, 14) = Mid(S, 215, 5)
        End If
        Set rg = rg.Offset(1, 0)
    Loop
    
    Close #1
    MsgBox ("Importação do arquivo """ + fileName + """ com sucesso!")
End Sub
