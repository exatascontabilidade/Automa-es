Attribute VB_Name = "clsUtilitarios"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Option Base 1

Public Sub DesabilitarControles()
    
    With Application
        .ScreenUpdating = False     'Desabilita a atualização de tela
        .DisplayAlerts = False      'Desabilita as mensagens de alerta
        .EnableEvents = False       'Desativa os eventos especiais
        .Calculation = xlManual     'Habilita o cálculo manual na planilha
    End With
    
End Sub

Public Sub HabilitarControles()
    
    With Application
        .EnableEvents = True        'Reativa os eventos especiais
        .DisplayAlerts = True       'Reabilita as mensagens de alerta
        .ScreenUpdating = True      'Reabilita a atualização de tela
        .Calculation = xlAutomatic  'Habilita o cálculo automático na planilha
        .ActiveSheet.Calculate      'Calcula a planilha ativa
    End With
    
End Sub

Public Sub LimparFiltros(ByVal Planilha As Worksheet)
    
    With Planilha
        
        If .[A1] = "FILTROS" Then .Range("B1:CZ1").ClearContents
        If .AutoFilterMode Then .AutoFilter.ShowAllData
        
    End With
    
End Sub

Public Function FormatarCNPJ(ByVal CNPJ As String)
    
    FormatarCNPJ = VBA.Format(VBA.Trim(Replace(Replace(Replace(CNPJ, ".", ""), "/", ""), "-", "")), VBA.String(14, "0"))
    
End Function

Public Sub LimparDados(ByVal Planilha As Worksheet, ByVal LinhaInicial As Long, ByVal Confirmacao As Boolean)

Dim Resposta As VbMsgBoxResult
    
    With Planilha
        
        If .FilterMode Then .AutoFilter.ShowAllData
        
        Select Case True
            
            Case Confirmacao = True
                Resposta = MsgBox("Tem certeza que deseja apagar TODOS os dados da planilha?" & vbCrLf & _
                                  "Essa operação NÃO pode ser desfeita.", vbExclamation + vbYesNo, "Deletar Dados")
                
                If Resposta = vbYes Then .Rows(LinhaInicial & ":" & Rows.Count).ClearContents
                
            Case Confirmacao = False
                .Rows(LinhaInicial & ":" & Rows.Count).ClearContents
                
        End Select
        
    End With
    
End Sub

Public Sub OtimizarCelulas(ByVal Planilha As Worksheet, ByVal LinhaTitulos As Long)

Dim UltCol As Long, LinhaInicial As Long
    
    LinhaInicial = LinhaTitulos + 1
    
    With Planilha
        
        On Error Resume Next
        If .FilterMode Then .AutoFilter.ShowAllData
        
        UltCol = .Cells(LinhaTitulos, .Columns.Count).END(xlToLeft).Column
        If UltCol < .Columns.Count Then .Range(.Cells(1, UltCol + 1), .Cells(1, .Columns.Count)).EntireColumn.Delete
        
        .Rows(LinhaInicial & ":" & Rows.Count).Delete
        
    End With
    
End Sub

Public Sub DeletarDados(ByVal Planilha As Worksheet, ByVal LinhaInicial As Long, ByVal Confirmacao As Boolean)

Dim Resposta As VbMsgBoxResult
    
    With Planilha
        
        If .FilterMode Then .AutoFilter.ShowAllData
        
        Select Case True
                    
            Case Confirmacao = True
                Resposta = MsgBox("Tem certeza que deseja apagar TODOS os dados da planilha?" & vbCrLf & _
                                  "Essa operação NÃO pode ser desfeita.", vbCritical + vbYesNo, "Deletar Dados")
                
                If Resposta = vbYes Then .Rows(LinhaInicial & ":" & Rows.Count).ClearContents
            
            Case Confirmacao = False
                .Rows(LinhaInicial & ":" & Rows.Count).ClearContents
                .Rows(LinhaInicial & ":" & Rows.Count).ClearFormats
                
        End Select
    
    End With
    
End Sub

Public Function NomeCol(ByVal Coluna As String)
    
Dim i As Byte
    
    For i = 1 To Len(Coluna)
        If VBA.Mid(Coluna, i, 1) Like "[A-Za-z]" Then NomeCol = NomeCol & VBA.Mid(Coluna, i, 1)
    Next i

End Function

Public Sub ApagarDicionarios(ParamArray ArrayDicionarios() As Variant)
    
Dim Dicionario
    
    For Each Dicionario In ArrayDicionarios
        Dicionario.RemoveAll
    Next Dicionario
    
End Sub

Public Sub ExportarDadosArrayList(ByVal Planilha As Worksheet, arrDados As ArrayList, Optional CelulaInicial As String, Optional qtdCols As Integer)

Dim arrAuxiliar As New ArrayList
Dim Endereco As Range
Dim item As Variant
Dim Regs As Long
    
    If CelulaInicial = "" Then CelulaInicial = "A" & Planilha.Range("A" & Rows.Count).END(xlUp).Row + 1
    If arrDados.Count = 0 Then Exit Sub
    
    If LBound(arrDados.item(0)) = 0 Then qtdCols = UBound(arrDados.item(0)) + 1 Else qtdCols = UBound(arrDados.item(0))
    
    If arrDados.Count > 50000 Then
        
        Do While arrDados.Count > 0
            
            If arrDados.Count < 50000 Then Regs = arrDados.Count Else Regs = 50000
            arrAuxiliar.addRange arrDados.getRange(0, Regs)
            
            Call ExportarDadosArrayList(Planilha, arrAuxiliar)
            If Regs < 50000 Then Call arrDados.Clear Else Call arrDados.removeRange(0, Regs)
            Call arrAuxiliar.Clear
            
        Loop
        
        Exit Sub
        
    End If
    
    If arrDados.Count > 0 Then
        
        If Planilha.AutoFilterMode Then Planilha.AutoFilter.ShowAllData
        Set Endereco = Planilha.Range(CelulaInicial).Resize(arrDados.Count, qtdCols)
        With Application
            Endereco.value = .Transpose(.Transpose(arrDados.toArray))
        End With
        
        Call fnExcel.FormatarIntervalo(Endereco, Planilha)
        
        If arrDados.Count <> 50000 Then
            Call FuncoesFormatacao.AplicarFormatacao(Planilha)
        End If
        
    End If
    
End Sub

Public Sub ExportarListaJson(ByVal Planilha As Worksheet, arrDados As ArrayList, Optional CelulaInicial As String, Optional qtdCols As Integer)

Dim arrLista As New ArrayList
Dim Endereco As Range
Dim item As Variant
    
    If CelulaInicial = "" Then CelulaInicial = "A" & Planilha.Range("A" & Rows.Count).END(xlUp).Row + 1
    
    If qtdCols = 0 Then
        For Each item In arrDados
            item = item.toArray()
            If LBound(item) = 0 Then qtdCols = UBound(item) + 1 Else qtdCols = UBound(item)
            Exit For
        Next item
    End If
    
    If arrDados.Count > 50000 Then
    
        For Each item In arrDados
            
            Call Util.AntiTravamento(a, 50)
            
            arrLista.Add item.toArray()
            arrDados.Remove item
            
            If arrLista.Count = 50000 Then
                Call ExportarDadosArrayList(Planilha, arrLista)
                arrLista.Clear
            End If
            
        Next item
        
        Call ExportarDadosArrayList(Planilha, arrLista)
        Exit Sub
        
    End If
    
    If arrDados.Count > 0 Then
    
        If Planilha.AutoFilterMode Then Planilha.AutoFilter.ShowAllData
        Set Endereco = Planilha.Range(CelulaInicial).Resize(arrDados.Count, qtdCols)
        With Application
            
            If qtdCols > 1 Then
                Endereco.value = .Transpose(.Transpose(arrDados.toArray))
            Else
                Endereco.value = .Transpose(arrDados.toArray)
            End If
            
        End With
        
    End If
    
    Call FuncoesFormatacao.AplicarFormatacao(Planilha)
    
End Sub

Public Sub ExportarDadosDicionario(ByVal Planilha As Worksheet, ByRef dicDados As Dictionary, Optional CelulaInicial As String, Optional qtdCols As Integer)

Dim teste As Variant
Dim Endereco As Range
Dim Chave As Variant
Dim dicAuxiliar As New Dictionary
    
    If CelulaInicial = "" Then CelulaInicial = "A" & Planilha.Cells(Rows.Count, 1).END(xlUp).Row + 1
    If dicDados.Count = 0 Then Exit Sub
    
    If LBound(dicDados.Items(0)) = 0 Then qtdCols = UBound(dicDados.Items(0)) + 1 Else qtdCols = UBound(dicDados.Items(0))
    
    If dicDados.Count > 50000 Then
        
        For Each Chave In dicDados.Keys
            
            Call Util.AntiTravamento(a, 50)
            
            dicAuxiliar(Chave) = dicDados(Chave)
            Call dicDados.Remove(Chave)
            
            If dicAuxiliar.Count = 50000 Then
                Call ExportarDadosDicionario(Planilha, dicAuxiliar)
                Call dicAuxiliar.RemoveAll
            End If
            
        Next Chave
        
        Call ExportarDadosDicionario(Planilha, dicAuxiliar)
        Call dicDados.RemoveAll
        Call dicAuxiliar.RemoveAll
        
        Exit Sub
        
    End If
    
    If Planilha.AutoFilterMode Then Planilha.AutoFilter.ShowAllData
    Set Endereco = Planilha.Range(CelulaInicial).Resize(dicDados.Count, qtdCols)
    
    If dicDados.Count = 1 Then
        
        Endereco.value = dicDados.Items(0)
        
    ElseIf dicDados.Count > 0 Then
        
        With Application
            Endereco.value = .Transpose(.Transpose(dicDados.Items))
        End With
        
    End If
    
    If dicDados.Count > 0 Then
        
        Call fnExcel.FormatarIntervalo(Endereco, Planilha)
        Call FuncoesFormatacao.AplicarFormatacao(Planilha)
        
    End If
    
End Sub

Public Sub AdicionarDadosDicionario(ByVal Planilha As Worksheet, Dicionario As Dictionary, Optional CelulaInicial As String, Optional qtdCols As Integer)

Dim Endereco As Range
Dim Chave As Variant

    If CelulaInicial = "" Then CelulaInicial = "A" & Planilha.Range("A" & Rows.Count).END(xlUp).Row + 1
    
    If qtdCols = 0 Then
        For Each Chave In Dicionario.Keys
            If LBound(Dicionario(Chave)) = 0 Then qtdCols = UBound(Dicionario(Chave)) + 1 Else qtdCols = UBound(Dicionario(Chave))
            Exit For
        Next Chave
    End If
    
    Set Endereco = Planilha.Range(CelulaInicial).Resize(Dicionario.Count, qtdCols)
    
    With Application
        If Dicionario.Count > 0 Then Endereco.value = .Transpose(.Transpose(Dicionario.Items))
    End With
    
End Sub

Public Function ValidarNumero(ByVal Valor As String)
    
Dim i As Long
    
    For i = 1 To Len(Valor)
        If VBA.Mid(Valor, i, 1) Like "[0-9,]" Then ValidarNumero = ValidarNumero & VBA.Mid(Valor, i, 1)
    Next i
    
End Function

Public Function FormatarData(ByVal Data As String)
    If Data = "CANCELAMENTO ESCRITURADO" Or Data = "DENEGADA" Or Data = "INUTILIZADA" Then FormatarData = Data: Exit Function
    If Not IsDate(Data) Then FormatarData = VBA.Format(VBA.Format(Data, "00/00/0000"), "yyyy-mm-dd") Else FormatarData = VBA.Format(Data, "yyyy-mm-dd")
End Function

Public Function ValidarValores(ByVal Valor As Variant, Optional Arredondar As Boolean, Optional CasasDecimais As Byte) As Double
    
    If Valor = "" Or Valor = "-" Then Valor = 0
    If Arredondar Then Valor = VBA.Round(Valor, CasasDecimais)
    ValidarValores = Valor
    
End Function

Public Function ValidarOperacao(ByVal Operacao As Variant) As String
    
    Select Case Operacao
        Case "0"
            ValidarOperacao = "Entrada"
        
        Case "1"
            ValidarOperacao = "Saida"
            
    End Select
    
End Function

Public Function ValidarEmissao(ByVal Emissao As Variant) As String
    
    Select Case Emissao
        Case "0"
            ValidarEmissao = "Própria"
        
        Case "1"
            ValidarEmissao = "Terceiros"
            
    End Select
    
End Function

Public Sub MsgInformativa(ByVal Mensagem As String, ByVal Titulo As String, ByVal TempoInicial As Date)
    MsgBox Mensagem & vbCrLf & vbCrLf & "Tempo decorrido: " & VBA.Format(Now() - TempoInicial, "ttttt"), vbInformation, Titulo
End Sub

Public Sub FormatarColunaNumero(ByVal Planilha As Worksheet, ParamArray colunas() As Variant)

Dim Coluna

    For Each Coluna In colunas
        Planilha.Range(Coluna & "2:" & Coluna & Rows.Count).TextToColumns
    Next Coluna

End Sub

Public Function ValidarCPF(ByVal CPF As String) As Boolean

Dim Caracter As Byte
Dim DV1 As Integer
Dim DV2 As Integer
    
    CPF = VBA.Format(CPF, VBA.String(11, "0"))
   
   For Caracter = 1 To 9
        DV1 = Val(VBA.Mid(CPF, Caracter, 1)) * Caracter + DV1
        If Caracter > 1 Then DV2 = Val(VBA.Mid(CPF, Caracter, 1)) * (Caracter - 1) + DV2
   Next
   
   DV1 = VBA.Right(DV1 Mod 11, 1)
   DV2 = VBA.Right((DV2 + (DV1 * 9)) Mod 11, 1)
   
   If VBA.Mid(CPF, 10, 1) = DV1 And VBA.Mid(CPF, 11, 1) = DV2 Then ValidarCPF = True
   
End Function

Public Function ValidarCNPJ(CNPJ As String) As Boolean

Dim Dgv1 As Integer
Dim Dgv2 As Integer
Dim i As Byte, Peso As Byte
Dim CNPJinv As String
    
    On Error GoTo Sair
    Dim VarUltDig As Integer

    If CNPJ <> "" Then
        
        CNPJ = VBA.Format(CNPJ, VBA.String(14, "0"))
        
        For i = 1 To 12
            CNPJinv = VBA.Mid(CNPJ, i, 1) & CNPJinv
        Next i
        
        For i = 1 To 12
            
            If i < 9 Then Peso = i + 1
            If i = 9 Then Peso = 2
            If i > 9 Then Peso = i - 7
            
            Dgv1 = VBA.Mid(CNPJinv, i, 1) * Peso + Dgv1
            If Peso = 9 Then Peso = 1
            
            Dgv2 = VBA.Mid(CNPJinv, i, 1) * (Peso + 1) + Dgv2
        
        Next
            
        Dgv1 = Dgv1 Mod 11
        If Dgv1 < 2 Then Dgv1 = 0 Else Dgv1 = 11 - Dgv1
                        
        Dgv2 = (2 * Dgv1 + Dgv2) Mod 11
        If Dgv2 < 2 Then Dgv2 = 0 Else Dgv2 = 11 - Dgv2
        
        If VBA.Mid(CNPJ, 13, 1) = Dgv1 And VBA.Mid(CNPJ, 14, 1) = Dgv2 Then ValidarCNPJ = True
        
    End If
    
Sair:
    
End Function

Public Function ValidarCPFCNPJ(ByVal CPFCNPJ As String) As String
    
    Select Case Len(CPFCNPJ)
        
        Case Is < 12
            
            If ValidarCPF(CPFCNPJ) Then
                ValidarCPFCNPJ = VBA.Format(CPFCNPJ, VBA.String(0, 11))
                Exit Function
            Else
                GoTo CNPJ:
            End If
            
        Case Else
CNPJ:
            If ValidarCNPJ(CPFCNPJ) Then
                ValidarCPFCNPJ = VBA.Format(CPFCNPJ, WorksheetFunction.Rept(0, 14))
                Exit Function
            End If
            
    End Select
    
End Function

Public Sub ClassificarDadosAcendentes(ByVal Planilha As Worksheet, LinhaInicial As Long, Coluna As String)

Dim rng As Range
    
    Planilha.Activate
    Set rng = DefinirIntervalo(Planilha, 3, 3)
    
    With Planilha.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=rng, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
End Sub

Public Function ExtrairNumeros(ByVal Texto As String) As String
   
Dim i As Byte
    
    For i = 1 To Len(Texto)
        If VBA.Mid(Texto, i, 1) Like "#" Then ExtrairNumeros = ExtrairNumeros & VBA.Mid(Texto, i, 1)
    Next i
    
End Function

Public Function GerarCarimboData()
    GerarCarimboData = Replace(Replace(Replace(Now, " ", ""), ":", ""), "/", "")
End Function

Public Function TratarAcentuacao(ByVal Texto As String)
    
    'MAIÚSCULAS
    Texto = Replace(Texto, "ï¿½", "A")
    Texto = Replace(Texto, "ï¿½", "E")
    Texto = Replace(Texto, "ï¿½", "I")
    Texto = Replace(Texto, "ï¿½", "O")
    Texto = Replace(Texto, "ï¿½", "U")
    
    Texto = Replace(Texto, "ï¿½", "A")
    Texto = Replace(Texto, "ï¿½", "E")
    Texto = Replace(Texto, "ï¿½", "O")
    
    Texto = Replace(Texto, "ï¿½", "A")
    Texto = Replace(Texto, "ï¿½", "O")
    
    Texto = Replace(Texto, "ï¿½", "C")

    'minúsculas
    Texto = Replace(Texto, "ï¿½", "a")
    Texto = Replace(Texto, "ï¿½", "e")
    Texto = Replace(Texto, "ï¿½", "i")
    Texto = Replace(Texto, "ï¿½", "o")
    Texto = Replace(Texto, "ï¿½", "u")
    
    Texto = Replace(Texto, "ï¿½", "a")
    Texto = Replace(Texto, "ï¿½", "e")
    Texto = Replace(Texto, "ï¿½", "o")
    
    Texto = Replace(Texto, "ï¿½", "a")
    Texto = Replace(Texto, "ï¿½", "o")
    
    Texto = Replace(Texto, "ï¿½", "c")
    
    TratarAcentuacao = Texto
    
End Function

Public Function ImportarJson(ByVal Arq As String)

    ' Leia o conteúdo do arquivo usando a codificação UTF-8
    With CreateObject("ADODB.Stream")
    
        .Open
        .Charset = "UTF-8"
        .LoadFromFile Arq
        ImportarJson = Split(.ReadText, vbLf)
        .Close
        
    End With

End Function

Public Function ImportarTxt(ByVal Arq As String, Optional Integro As Boolean)

Dim fso As FileSystemObject
Dim ARQUIVO As TextStream
Dim Conteudo As String
    
    Set fso = New FileSystemObject
    Set ARQUIVO = fso.OpenTextFile(Arq, ForReading, False)
    Conteudo = ARQUIVO.ReadAll
    
    ARQUIVO.Close
    ImportarTxt = Conteudo
    If Integro Then Exit Function
    
    ImportarTxt = VBA.Split(Conteudo, vbCrLf)
    If UBound(ImportarTxt) <= 1 Then ImportarTxt = VBA.Split(Conteudo, vbLf)
    
End Function

Public Function ImportarConteudo(ByVal Arq As String)

    ' Leia o conteúdo do arquivo usando a codificação UTF-8
    With CreateObject("ADODB.Stream")
    
        .Open
        .Charset = "ISO-8859-1"
        .LoadFromFile Arq
        ImportarConteudo = .ReadText
        .Close
        
    End With

End Function

Public Sub ExportarTxt(ByVal Arq As String, TXT As Variant)
    
Dim NUM As Integer
    
    NUM = FreeFile
    Open Arq For Output As #NUM
        Print #NUM, TXT
    Close #NUM
    
End Sub

Public Sub AdicionarRegistros(ByVal Arq As String, TXT As Variant)
    
Dim NUM As Integer
    
    NUM = FreeFile
    Open Arq For Append As #NUM
        Print #NUM, TXT
    Close #NUM
    
End Sub

Private Function ExcluirAssinaturaEFD(ByVal Arq As String)

Dim EFD As New ArrayList
Dim NUM As Integer
Dim i As Long
Dim Registro

    NUM = FreeFile
    Open Arq For Input As #NUM
        
        i = 0
        Do While Not EOF(NUM)
            
            i = i + 1
            If i Mod 300 = 0 Then DoEvents
            Line Input #NUM, Registro
            
            Select Case VBA.Mid(Registro, 2, 4)
            
                Case "BRCA"
                    Exit Do
                
                Case Else
                    EFD.Add Registro

            End Select
        
        Loop
        
        Close #NUM
        Select Case EFD.Count
        
            Case Is <= 1
                ExcluirAssinaturaEFD = Split(Join(EFD.toArray, ""), vbLf)
                If UBound(ExcluirAssinaturaEFD) = 0 Then ExcluirAssinaturaEFD = Split(Join(EFD.toArray, ""), vbCrLf)
                If UBound(ExcluirAssinaturaEFD) = 0 Then ExcluirAssinaturaEFD = Split(Join(EFD.toArray, ""), vbCr)
                    
            Case Else
                ExcluirAssinaturaEFD = EFD.toArray
                
        End Select
        
End Function

Public Sub RemoverAssinaturaEFD(ByVal Arqs As Variant)

Dim EFD As New ArrayList
Dim NUM As Integer
Dim Registro, Arq
Dim i As Long

    NUM = FreeFile
    For Each Arq In Arqs
        
        i = i + 1
        Application.StatusBar = "Removendo assinatura do arquivo " & i & " de " & UBound(Arqs)
        
        Open Arq For Input As #NUM
            
            Do While Not EOF(NUM)
                
                If EFD.Count Mod 300 = 0 Then DoEvents
                Line Input #NUM, Registro
                
                Select Case VBA.Mid(Registro, 2, 4)
                    
                    Case "BRCA"
                        Exit Do
                        
                    Case Else
                        EFD.Add Registro
                        
                End Select
                
            Loop
            
            Close #NUM
            Call ExportarTxt(Arq, Join(EFD.toArray, vbCrLf))
            EFD.RemoveAll
            
    Next Arq
    
End Sub

Public Function GerarSequencial(ByVal nRegistro As String, nCaracteres As Integer)
    GerarSequencial = VBA.Format(nRegistro, VBA.String(nCaracteres, "0"))
End Function

Public Function GerarBrancos(ByVal QTD As String)
    GerarBrancos = VBA.String(QTD, " ")
End Function

Public Function FormatarTexto(ByVal Texto As String)
    If Texto <> "" Then
        Texto = VBA.Replace(Texto, "'", "")
        FormatarTexto = "'" & VBA.Trim(Texto)
    End If
End Function

Public Function FormatarValores(ByVal Valor As String) As Double
    If IsNumeric(Valor) = False Then Valor = 0
    FormatarValores = Replace(Valor, ".", ",")
End Function

Public Function ArredondarValores(ByVal Valor As String, Optional ByVal nDecimais As Byte) As Double
    
    Select Case True
        
        Case (IsNumeric(Valor)) And (nDecimais > 0)
            ArredondarValores = VBA.Round(VBA.Replace(Valor, ".", ","), nDecimais)
            
        Case (IsNumeric(Valor)) And (nDecimais = 0)
            ArredondarValores = VBA.Round(VBA.Replace(Valor, ".", ","), 2)
            
        Case (Valor = "") And (nDecimais > 0)
            ArredondarValores = VBA.Round(VBA.Replace(Valor, ".", ","), nDecimais)
            
        Case (Valor = "") And (nDecimais = 0)
            ArredondarValores = VBA.Round(VBA.Replace(Valor, ".", ","), 2)
            
    End Select
    
End Function

Public Function ConverterIBGE_UF(ByVal cIBGE As String) As String

    Select Case VBA.Left(cIBGE, 2)
    
        Case "11"
            ConverterIBGE_UF = "RO"
            
        Case "12"
            ConverterIBGE_UF = "AC"
            
        Case "13"
            ConverterIBGE_UF = "AM"
            
        Case "14"
            ConverterIBGE_UF = "RR"
            
        Case "15"
            ConverterIBGE_UF = "PA"
            
        Case "16"
            ConverterIBGE_UF = "AP"
            
        Case "17"
            ConverterIBGE_UF = "TO"
            
        Case "21"
            ConverterIBGE_UF = "MA"
            
        Case "22"
            ConverterIBGE_UF = "PI"
            
        Case "23"
            ConverterIBGE_UF = "CE"
            
        Case "24"
            ConverterIBGE_UF = "RN"
            
        Case "25"
            ConverterIBGE_UF = "PB"
            
        Case "26"
            ConverterIBGE_UF = "PE"
            
        Case "27"
            ConverterIBGE_UF = "AL"
            
        Case "28"
            ConverterIBGE_UF = "SE"
            
        Case "29"
            ConverterIBGE_UF = "BA"
            
        Case "31"
            ConverterIBGE_UF = "MG"
            
        Case "32"
            ConverterIBGE_UF = "ES"
            
        Case "33"
            ConverterIBGE_UF = "RJ"
            
        Case "35"
            ConverterIBGE_UF = "SP"
            
        Case "41"
            ConverterIBGE_UF = "PR"
            
        Case "42"
            ConverterIBGE_UF = "SC"
            
        Case "43"
            ConverterIBGE_UF = "RS"
            
        Case "50"
            ConverterIBGE_UF = "MS"
            
        Case "51"
            ConverterIBGE_UF = "MT"
            
        Case "52"
            ConverterIBGE_UF = "GO"
            
        Case "53"
            ConverterIBGE_UF = "DF"
        
        Case "", "99"
            ConverterIBGE_UF = "EX"
            
    End Select

End Function

Public Function ClassificarNotaFiscal(ByVal CNPJEmitente As String, ByVal tpNF As String, ByVal Modelo As String, ByVal Registro As Variant, _
                                      ByRef arrChaves As ArrayList, ByRef dicEntradasNFe As Dictionary, ByRef dicSaidasNFe As Dictionary, _
                                      Optional ByRef dicSaidasCTes As Dictionary, Optional ByRef dicEntradasCTes As Dictionary, _
                                      Optional ByRef dicSaidasNFCe As Dictionary, Optional ByRef dicSaidasCFe As Dictionary)
                     
Dim chNFe As String

On Error GoTo Tratar:
    
    chNFe = VBA.Replace(Registro(5), "'", "")
    Select Case CNPJContribuinte = CNPJEmitente
        
        Case True
            
            Select Case tpNF
                
                Case "Entrada"
                    If Modelo = "55" And Not arrChaves.contains(chNFe) Then
                        dicEntradasNFe(chNFe) = Registro
                        arrChaves.Add chNFe
                    End If
                    
                    If (Modelo = "57" Or Modelo = "67") And Not arrChaves.contains(chNFe) Then
                        dicEntradasCTes(chNFe) = Registro
                        arrChaves.Add chNFe
                    End If
                    
                Case "Saida"
                    If Modelo = "55" And Not arrChaves.contains(chNFe) Then
                        dicSaidasNFe(chNFe) = Registro
                        arrChaves.Add chNFe
                    End If
                    
                    If (Modelo = "57" Or Modelo = "67") And Not arrChaves.contains(chNFe) Then
                        dicSaidasCTes(chNFe) = Registro
                        arrChaves.Add chNFe
                    End If
                    
                    If Modelo = "59" And Not arrChaves.contains(chNFe) Then
                        dicSaidasCFe(chNFe) = Registro
                        arrChaves.Add chNFe
                    End If
                    
                    If Modelo = "65" And Not arrChaves.contains(chNFe) Then
                        dicSaidasNFCe(chNFe) = Registro
                        arrChaves.Add chNFe
                    End If
                    
            End Select
            
        Case False
            If Modelo = "55" And Not arrChaves.contains(chNFe) Then
                dicEntradasNFe(chNFe) = Registro
                arrChaves.Add chNFe
            End If
            
            If (Modelo = "57" Or Modelo = "67") And Not arrChaves.contains(chNFe) Then
                dicEntradasCTes(chNFe) = Registro
                arrChaves.Add chNFe
            End If
                    
    End Select

Exit Function
Tratar:

    Resume
    If Err.Number <> 0 Then Call TratarErros(Err, "Utilitarios.ClassificarNotaFiscal")
    
End Function

Public Function GerarObservacoes(ByVal StatusNF As String, ByVal CNPJEmit As String, ByVal UF As String, ByVal tpNF As String)

On Error GoTo Tratar:

    With DadosDoce
    
        Select Case StatusNF
            
            Case "Cancelada"
            
                If CNPJContribuinte <> CNPJEmit Then
                    
                    .vNF = 0
                    .StatusSPED = "NÃO LANÇAR"
                    .OBSERVACOES = "NF CANCELADA"
                    
                ElseIf CNPJContribuinte = CNPJEmit Then
                    
                    .vNF = 0
                    .OBSERVACOES = "NF CANCELADA"
                    If tpNF = "Entrada" Then
                        .CNPJPart = CNPJContribuinte
                        .RazaoPart = RazaoContribuinte
                        .UF = UF
                        .OBSERVACOES = "CANCELADA CONTRIBUINTE"
                    End If
                    
                End If
                            
            Case "Denegada"
            
                If CNPJContribuinte <> CNPJEmit Then
                
                    .vNF = 0
                    .StatusSPED = "NÃO LANÇAR"
                    .OBSERVACOES = "NF DENEGADA"
                    
                ElseIf CNPJContribuinte = CNPJEmit Then
                
                    .vNF = 0
                    .OBSERVACOES = "NF DENEGADA"
                    
                End If
            
            Case "Autorizada"
                
                If CNPJContribuinte <> CNPJEmit Then
                
                    If tpNF = "Entrada" Then
                        .StatusSPED = "MANIFESTAR NF"
                        .OBSERVACOES = "DEVOLUÇÃO DE FORNECEDOR"
                    End If
                
                ElseIf CNPJContribuinte = CNPJEmit Then
                    
                    If tpNF = "Entrada" Then
'                        .CNPJPart = CNPJContribuinte
'                        .RazaoPart = RazaoContribuinte
                        .UF = UF
                        .OBSERVACOES = "ENTRADA PRÓPRIA"
                        
                    End If
                
                End If
                
        End Select
    
    End With
    
Exit Function
Tratar:
    
    If Err.Number <> 0 Then Call TratarErros(Err, "Utilitarios.GerarObservacoes")
    
End Function

Public Function FormatarDadosContribuinte(ByVal Dado As String)
    
    On Error GoTo Tratar:
    
    FormatarDadosContribuinte = Trim(Replace(Replace(Dado, "=", ""), """", ""))
    
Exit Function
Tratar:
    
    If Err.Number <> 0 Then Call Erro.TratarErros(Err, "Utilitarios.FormatarDadosContribuinte")
            
End Function

Public Function FormatarEXTIPI(ByVal Texto As String) As String

    On Error GoTo Tratar:

        Texto = Trim(Texto)
        Texto = Replace(Texto, "Ex ", "")
        FormatarEXTIPI = VBA.Format(Texto, "000")
    
Exit Function
Tratar:
    
    If Err.Number <> 0 Then Call Erro.TratarErros(Err, "Utilitarios.FormatarEXTIPI")
        
End Function

Public Function ListarArquivos(ByRef arrXMLs As ArrayList, ByVal Caminho As String, Optional Comeco As Double)

Dim fso As New FileSystemObject
Dim ARQUIVO As file
Dim pasta As Folder
Dim subpasta As Folder
                
    Set pasta = fso.GetFolder(Caminho)
    For Each ARQUIVO In pasta.Files
        Call Util.AntiTravamento(a, 100, "Listando XMLS para importação, por favor aguarde", arrXMLs.Count, Comeco)
        If VBA.LCase(ARQUIVO.Path) Like "*.xml" Then arrXMLs.Add ARQUIVO.Path
    Next ARQUIVO
    
    For Each subpasta In pasta.SubFolders
        Call ListarArquivos(arrXMLs, subpasta.Path, Comeco)
    Next subpasta
    
End Function

Public Function MontarCSTICMS(ByRef Produto As IXMLDOMNode)

Dim orig As String, CST As String

    orig = Produto.SelectSingleNode("imposto/ICMS//orig").text
    If Not Produto.SelectSingleNode("imposto/ICMS//CST") Is Nothing Then CST = Produto.SelectSingleNode("imposto/ICMS//CST").text
    If Not Produto.SelectSingleNode("imposto/ICMS//CSOSN") Is Nothing Then CST = Produto.SelectSingleNode("imposto/ICMS//CSOSN").text
    
    MontarCSTICMS = orig & CST
    
End Function

Public Function MsgAlerta(ByVal Msg As String, Optional ByVal Titulo As String) As String
    MsgBox Msg, vbExclamation, Titulo
End Function

Public Function MsgCritica(ByVal Msg As String, Optional ByVal Titulo As String) As String
    MsgBox Msg, vbCritical, Titulo
End Function

Public Function MsgDecisao(ByVal Msg As String, Optional ByVal Titulo As String) As VbMsgBoxResult
    MsgDecisao = MsgBox(Msg, vbExclamation + vbYesNo, Titulo)
End Function

Public Function MsgInformativaDecisao(ByVal Msg As String, Optional ByVal Titulo As String) As VbMsgBoxResult
    MsgInformativaDecisao = MsgBox(Msg, vbInformation + vbYesNo, Titulo)
End Function

Public Function MsgAviso(ByVal Msg As String, Optional ByVal Titulo As String) As String
    MsgBox Msg, vbInformation, Titulo
End Function

Public Function SelecionarArquivo(ByVal Extensao As String, Optional ByVal Titulo As String)
    
    On Error GoTo Tratar:
        
        SelecionarArquivo = Application.GetOpenFilename("Arquivos " & Extensao & "(*." & Extensao & "), *." & Extensao, , Titulo, , False)
        
Exit Function
Tratar:
    
    If Err.Number <> 0 Then Call Erro.TratarErros(Err, "Utilitarios.SelecionarArquivo")
    
End Function

Public Function SelecionarArquivos(ByVal Extensao As String, Optional ByVal Titulo As String)
    
    On Error GoTo Tratar:
        
        SelecionarArquivos = Application.GetOpenFilename("Arquivos " & Extensao & "(*." & Extensao & "), *." & Extensao, , Titulo, , True)
        
Exit Function
Tratar:
    
    If Err.Number <> 0 Then Call Erro.TratarErros(Err, "Utilitarios.SelecionarArquivos")
    
End Function

'Public Function SelecionarPasta(Optional ByVal Titulo As String) As String
'
'Dim fd As FileDialog
'Dim fso As New FileSystemObject
'Dim objFolder As Object
'Dim strFullPathFromDialog As String
'
'    On Error GoTo ErroSelecaoPasta
'
'    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
'
'    If Len(Titulo) > 0 Then
'        fd.Title = Titulo
'    Else
'        fd.Title = "Selecione a pasta onde os arquivos XML estão armazenados"
'    End If
'
'    If fd.Show = -1 Then
'        strFullPathFromDialog = fd.SelectedItems(1)
'
'        Set fso = CreateObject("Scripting.FileSystemObject")
'        If fso.FolderExists(strFullPathFromDialog) Then
'            Set objFolder = fso.GetFolder(strFullPathFromDialog)
'            SelecionarPasta = objFolder.path
'
'        Else
'
'            SelecionarPasta = strFullPathFromDialog
'
'        End If
'    Else
'
'        SelecionarPasta = ""
'    End If
'
'Limpeza:
'    Set objFolder = Nothing
'    Set fso = Nothing
'    Set fd = Nothing
'    Exit Function
'
'ErroSelecaoPasta:
'    SelecionarPasta = ""
'    MsgBox "Erro ao selecionar a pasta ou obter o caminho: " & Err.Description, vbExclamation, "Erro"
'    Resume Limpeza
'
'End Function

Public Function SelecionarPasta(Optional ByVal Titulo As String) As String
    
    On Error GoTo Tratar:
        
        With Application.FileDialog(msoFileDialogFolderPicker)
        
            If Titulo <> "" Then .Title = Titulo
            .AllowMultiSelect = False
            .Show
            
            If .SelectedItems.Count > 0 Then SelecionarPasta = .SelectedItems(1)
        
        End With

Exit Function
Tratar:
    
    If Err.Number <> 0 Then Call Erro.TratarErros(Err, "Utilitarios.SelecionarPasta")
    
End Function

Public Function TransporTabela(ByVal Dados As Variant)
    
Dim i As Long, j As Long
ReDim Dados(LBound(Dados, 2) To UBound(Dados, 2), LBound(Dados, 1) To UBound(Dados, 1))
    
    On Error GoTo Tratar:
        
        For i = LBound(Dados, 2) To UBound(Dados, 2)
            For j = LBound(Dados, 1) To UBound(Dados, 1)
                Dados(i, j) = Dados(j, i)
            Next
        Next
        TransporTabela = Dados
    
Exit Function
Tratar:
    
    If Err.Number <> 0 Then Call Erro.TratarErros(Err, "Utilitarios.TransporTabela")
        
End Function

Public Function TratarDataHora(ByVal DataHora As String)
    TratarDataHora = Replace(VBA.Left(DataHora, 19), "T", " ")
End Function

Public Function TratarNumero(ByVal Numero As String)
    
    On Error GoTo Tratar:
    
        TratarNumero = "'" & Replace(Numero, ",", ".") & "'"
            
Exit Function
Tratar:
    
    If Err.Number <> 0 Then Call Erro.TratarErros(Err, "Utilitarios.TratarNumero")
                  
End Function

Public Function TratarTexto(ByVal Texto As String)
      
    On Error GoTo Tratar:
        
        TratarTexto = Texto
        If InStr(1, Texto, "|") > 0 Then TratarTexto = VBA.Replace(TratarTexto, "|", " ")
    
Exit Function
Tratar:
    
    If Err.Number <> 0 Then Call Erro.TratarErros(Err, "Utilitarios.TratarTexto")
                    
End Function

Public Function RemoverPipes(ByVal Texto As String)
    
    RemoverPipes = VBA.Replace(Texto, "|", "/")
    
End Function

Public Function TratarTextoNCM(ByVal Texto As String) As String
    
    On Error GoTo Tratar:
        
        Texto = Trim(Texto)
        Texto = Replace(Texto, "'", "")
        Texto = Replace(Texto, "-- ", "")
        Texto = Replace(Texto, "- ", "")
        TratarTextoNCM = Texto
        
Exit Function
Tratar:
    
    If Err.Number <> 0 Then Call Erro.TratarErros(Err, "Utilitarios.TratarTexto")
    
End Function

Public Function UltimaLinha(ByRef Plan As Worksheet, ByVal Coluna As String)

    On Error GoTo Tratar:
        UltimaLinha = Plan.Range(Coluna & Plan.Rows.Count).END(xlUp).Row
    
Exit Function
Tratar:
    
    If Err.Number <> 0 Then Call Erro.TratarErros(Err, "Utilitarios.UltimaLinha")
    
End Function

Public Function ValidarCRT(ByVal CRT As String) As String
    If CRT = "1" Then ValidarCRT = "SIM" Else ValidarCRT = "NÃO"
End Function

Public Function ValidarPercentual(ByRef NFe As IXMLDOMNode, ByVal Tag As String)
    If Not NFe.SelectSingleNode(Tag) Is Nothing Then ValidarPercentual = NFe.SelectSingleNode(Tag).text / 100
End Function

Public Function ValidarXML(ByRef NFe As IXMLDOMNode) As Boolean
    If Not NFe.SelectSingleNode("nfeProc") Is Nothing Then ValidarXML = True
End Function

Public Function ValidarTag(ByRef NFe As IXMLDOMNode, ByVal Tag As String) As String
    If Not NFe.SelectSingleNode(Tag) Is Nothing Then ValidarTag = NFe.SelectSingleNode(Tag).text
End Function

Public Function VerificarDicionario(ByRef Dicionario As Dictionary, ByVal Chave As Variant, ByVal Posicao As Byte) As Variant
    If Dicionario.Exists(Chave) Then VerificarDicionario = Dicionario(Chave)(Posicao) Else VerificarDicionario = "Chave do dicionário Inexistente"
End Function

Public Function GuardarEnderecosArrayList(ByVal Extensao As String, ByRef arrXMLs As ArrayList) As Boolean

Dim XMLS As Variant, XML
    
    XMLS = Util.SelecionarArquivos(Extensao)
    
    Inicio = Now()
    If VarType(XMLS) <> 11 Then
        
        For Each XML In XMLS
            
            Call Util.AntiTravamento(a, 5)
            arrXMLs.Add XML
            
        Next XML
        
        GuardarEnderecosArrayList = True
        
    End If
    
End Function

Function ConverterNumeroColuna(nCol As Long) As String

   Dim a As Long
   Dim b As Long
   a = nCol
   ConverterNumeroColuna = ""
   Do While nCol > 0
      a = Int((nCol - 1) / 26)
      b = (nCol - 1) Mod 26
      ConverterNumeroColuna = Chr(b + 65) & ConverterNumeroColuna
      nCol = a
   Loop

End Function

Public Function IndexarCampos(ByRef Campos As Variant, dicTitulos As Dictionary)

Dim Campo As Variant
Dim i As Integer

    For Each Campo In Campos
                
        dicTitulos(Campo) = i
        i = i + 1
    
    Next Campo

End Function

Public Function AntiTravamento(ByRef a As Long, Optional ByVal Intervalo As Long, Optional ByVal Msg As String, _
                               Optional ByVal totArqs As Long, Optional Inicio As Double)

Dim Decorrido As Double
Dim Estimativa As Double

    a = a + 1
    If Intervalo = 0 Then Intervalo = 100
    
    If Inicio > 0 Then Decorrido = Timer - Inicio
    If Decorrido > 0 Then
        Estimativa = (Decorrido / a) * (totArqs - a)
        Msg = Msg & " - Tempo estimado para conclusão: " & VBA.Format(Estimativa / 86400, "hh:mm:ss")
    End If
    
    If a Mod Intervalo = 0 Then
        If Msg <> "" Then Application.StatusBar = Msg
        'DoEvents
    End If
    
    If a > 2000000 Then a = 0

End Function

Public Function FimMes(ByVal Data As String)
    FimMes = DateSerial(Year(Data), Month(Data) + 1, 1) - 1
End Function

Function ConverterPeriodoData(ByVal Periodo As String) As Date

  ' Converte o período "MM/AAAA" para uma data.
  Dim DataInicial As Date
  DataInicial = DateSerial(Right(Periodo, 4), Left(Periodo, 2), 1)

  ' Calcula o último dia do mês.
  ConverterPeriodoData = DateSerial(Year(DataInicial), Month(DataInicial) + 1, 0)

End Function

Function ExtrairPeriodo(ByVal Data As String) As String

Dim DataInicial As String
    
    DataInicial = fnSPED.FormatarDataSPED(Data)
    ExtrairPeriodo = VBA.Format(DataInicial, "mm/yyyy")
    
End Function

Public Function DefinirTitulos(ByRef Plan As Worksheet, ByVal nLin As String) As Variant
    
Dim UltCol As Long
    
    UltCol = Plan.Cells(nLin, Columns.Count).END(xlToLeft).Column
    DefinirTitulos = Plan.Cells(nLin, 1).Resize(, UltCol)

End Function

Public Function IndexarDados(ByVal Dados As Variant) As Dictionary
    
Dim i As Long
Dim Dado As Variant
Dim dicDados As New Dictionary
    
    i = 1
    For Each Dado In Dados
                
        dicDados(Dado) = i
        i = i + 1
        
    Next Dado
    
    Set IndexarDados = dicDados
    
End Function

Public Function DefinirDados(ByRef Plan As Worksheet, ByVal LinhaInicial As String, ByVal LinhaTitulo As String) As Variant
Dim UltLin As Long
Dim rng As Range

    UltLin = Util.UltimaLinha(Plan, "A")
    Set rng = DefinirIntervalo(Plan, LinhaInicial, LinhaTitulo)
    If rng Is Nothing Then DefinirDados = Empty Else DefinirDados = rng.Value2
    
End Function

Public Function DefinirIntervalo(ByRef Plan As Worksheet, ByVal Lini As String, ByVal LinT As String) As Range

Dim UltLin As Long
Dim UltCol As Long
    
    UltLin = Util.UltimaLinha(Plan, "A")
    UltCol = Plan.Cells(LinT, 1).END(xlToRight).Column
    If UltLin >= Lini Then Set DefinirIntervalo = Plan.Cells(Lini, 1).Resize(UltLin - Lini + 1, UltCol)
    
End Function

Public Function DefinirCodificacao(ByVal Arq As String, ByVal Encoding As String) As String

    ' Leia o conteúdo do arquivo usando a codificação UTF-8
    With CreateObject("ADODB.Stream")
    
        .Open
        .Charset = Encoding
        .LoadFromFile Arq
        DefinirCodificacao = .ReadText
        .Close
        
    End With
    
End Function

Public Function DefinirIntervaloFiltrado(ByRef Plan As Worksheet, ByVal Lini As Long, ByVal LinT As Long, Optional arrList As Boolean) As Variant

Dim UltLin As Long
Dim UltCol As Long
Dim Intervalo As Range, Area As Range
Dim Linha As Range
Dim i As Long
Dim Campos As Variant, Dados
Dim arrDados As New ArrayList
    
    'Definindo o intervalo com células visíveis
    UltLin = Util.UltimaLinha(Plan, "A")
    UltCol = Plan.Cells(LinT, 1).END(xlToRight).Column
    
    If UltLin >= Lini Then
        
        On Error GoTo Tratar:
            Set Intervalo = Plan.Cells(Lini, 1).Resize(UltLin - Lini + 1, UltCol).SpecialCells(xlCellTypeVisible)
            
        On Error GoTo 0
        
        'Coletando todas as informações das linhas visíveis
        For Each Area In Intervalo.Areas
            For Each Linha In Area.Rows
                Campos = Linha
                arrDados.Add Campos
            Next Linha
        Next Area
        
    End If
    
    If arrList = True Then
    
        Set DefinirIntervaloFiltrado = arrDados
        Exit Function
        
    End If
    
    If arrDados.Count = 0 Then Exit Function
    
    If arrDados.Count = 1 Then
        
        Dados = Campos
        DefinirIntervaloFiltrado = Dados
        Exit Function
        
    End If
    
    DefinirIntervaloFiltrado = Application.Transpose(Application.Transpose(arrDados.toArray))
    
Exit Function
Tratar:
    
End Function

'Public Function BidimensionarArray(ByRef Campos As Variant) As Variant
'
'Dim i As Long
'ReDim Dados(1 To 1, 1 To UBound(Campos))
'
'    For i = 1 To UBound(Campos)
'        Dados(1, i) = Campos(i)
'    Next i
'
'    BidimensionarArray = Dados
'
'End Function

Public Function ClassificarDicionario(ByRef dicDados As Dictionary, Optional ByVal OrdemDecrescente As Boolean)

Dim arrDados As New ArrayList
Dim dicAuxiliar As New Dictionary
Dim Chave
    
    For Each Chave In dicDados.Keys()
        arrDados.Add Chave
    Next Chave
    
    arrDados.Sort
    If OrdemDecrescente = True Then arrDados.Reverse
    
    For Each Chave In arrDados
        dicAuxiliar(Chave) = dicDados(Chave)
    Next Chave
    
    dicDados.RemoveAll
    Set dicDados = dicAuxiliar
    
End Function

Public Function ExtrairDadoDicionario(ByRef dicDados As Dictionary, ByRef dicTitulos As Dictionary, ByVal Chave As String, ByVal nCampo As String) As Variant
    
Dim i As Byte
Dim dicCampos As Variant
    
    If dicDados.Exists(Chave) Then
        
        dicCampos = dicDados(Chave)
        If LBound(dicCampos) = 0 Then i = 1
        
        ExtrairDadoDicionario = dicCampos(dicTitulos(nCampo) - i)
        
    End If
    
End Function

Function VerificarTipoSPED(Registro As String) As String

Dim Campos As Variant
    
    Campos = VBA.Split(Registro, "|")
    If UBound(Campos) < 7 Then
        VerificarTipoSPED = "Desconhecido"
        Exit Function
    End If
    
    If IsDate(Util.FormatarData(Campos(4))) And IsDate(Util.FormatarData(Campos(5))) Then
        VerificarTipoSPED = "Fiscal"
        VersaoFiscal = VBA.Format(Campos(2), "000")
        
    ElseIf IsDate(Util.FormatarData(Campos(6))) And IsDate(Util.FormatarData(Campos(7))) Then
        VerificarTipoSPED = "Contribuições"
        VersaoContribuicoes = VBA.Format(Campos(2), "000")
        
    End If
    
End Function

Public Function ValidarUF(UF As String) As Boolean
    
    Select Case UF
        
        Case "AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MG", "MS", "MT", "PA", "PB", "PE", "PI", "PR", "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO"
            ValidarUF = True
            
    End Select
    
End Function

Public Function CriarDicionarioRegistro(ByVal Plan As Worksheet, ParamArray CamposChave() As Variant) As Dictionary

Dim Titulos As Variant, Registro, Campos, CampoChave, CamposTexto, Campo
Dim arrCamposChave As New ArrayList
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim dicDados As New Dictionary
Dim Chave As String
Dim b As Long
    
    If UBound(CamposChave) = -1 Then CamposChave = Array("CHV_REG")
    If Plan.AutoFilterMode Then Plan.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(Plan, 4, 3)
    If Dados Is Nothing Then
        
        Set CriarDicionarioRegistro = New Dictionary
        Exit Function
        
    End If
    
    b = 0
    Comeco = Timer
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 10, "Carregando dados do " & Plan.name, Dados.Rows.Count, Comeco)
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
                        
            For Each CampoChave In CamposChave
                arrCamposChave.Add Campos(dicTitulos(CampoChave))
            Next CampoChave
            
            Chave = VBA.Replace(VBA.Join(arrCamposChave.toArray()), " ", "")
            
            dicDados(Chave) = Campos
            arrCamposChave.Clear
            
         End If
         
    Next Linha
    
    Set CriarDicionarioRegistro = dicDados
    
End Function

Public Function CriarArrayListRegistro(ByVal Plan As Worksheet, Optional LinIni As Integer = 4, Optional LinTit As Integer = 3) As ArrayList

Dim Dados As Range, Linha As Range
Dim arrRegistros As New ArrayList
Dim dicTitulos As New Dictionary
Dim dicDados As New Dictionary
Dim Campos As Variant
Dim Chave As String
Dim b As Long
    
    If Plan.AutoFilterMode Then Plan.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(Plan, LinIni, LinTit)
    If Dados Is Nothing Then Exit Function
    
    b = 0
    Comeco = Timer
    Set dicTitulos = Util.MapearTitulos(Plan, LinTit)
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 10, "Carregando dados do " & Plan.name, Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then arrRegistros.Add Campos
        
    Next Linha
    
    Set CriarArrayListRegistro = arrRegistros
    
End Function

Public Function CarregarChavesEmitentes(ByVal Plan As Worksheet, ByVal CNPJEmit As String, ParamArray CamposSelecionados() As Variant) As ArrayList

Dim Titulos As Variant, Registro, Campos, Campo, CamposTexto
Dim arrCamposSelecionados As New ArrayList
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim arrRegistros As New ArrayList
Dim Chave As String, chvDoc$, CNPJ$
Dim b As Long
    
    If UBound(CamposSelecionados) = -1 Then Exit Function
    
    'Limpa os Filtros
    If Plan.AutoFilterMode Then Plan.AutoFilter.ShowAllData
    
    'Definie o intervalo de dados
    Set Dados = Util.DefinirIntervalo(Plan, 4, 3)
    
    'Sai da função caso não existam dados a processar
    If Dados Is Nothing Then GoTo Sair:
    
    b = 0
    Comeco = Timer
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100)
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            'Seleciona os campos escolhidos pelo usuário
            For Each Campo In CamposSelecionados
            
                chvDoc = Campos(dicTitulos(Campo))
                CNPJ = VBA.Mid(chvDoc, 7, 14)
                If chvDoc Like "*" & CNPJEmit & "*" Then arrCamposSelecionados.Add chvDoc
                
            Next Campo
            
            'Forma um registro com os campos selecionados
            arrRegistros.addRange arrCamposSelecionados
            arrCamposSelecionados.Clear
            
         End If
         
    Next Linha
    
Sair:
    'Retorna os dados filtraos para o usuário
    Set CarregarChavesEmitentes = arrRegistros
    
End Function

Public Function CriarDicionarioCorrelacoes(ByVal Plan As Worksheet) As Dictionary

Dim Chave As String, CNPJForn$, cProdForn$, unForn$
Dim arrCamposChave As New ArrayList
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim dicDados As New Dictionary
Dim Campos As Variant
    
    If Plan.AutoFilterMode Then Plan.AutoFilter.ShowAllData
    Set Dados = Util.DefinirIntervalo(Plan, 4, 3)
    
    If Dados Is Nothing Then Exit Function
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            CNPJForn = VBA.Format(Campos(dicTitulos("CNPJ_FORNECEDOR")), String(14, "0"))
            cProdForn = Campos(dicTitulos("COD_PROD_FORNEC"))
            unForn = Campos(dicTitulos("UND_FORNEC"))
            
            Chave = CNPJForn & cProdForn & unForn
            Campos(dicTitulos("CNPJ_FORNECEDOR")) = Util.FormatarTexto(CNPJForn)
            Campos(dicTitulos("COD_PROD_FORNEC")) = Util.FormatarTexto(cProdForn)
            
            dicDados(Chave) = Campos
            
         End If
         
    Next Linha
    
    Set CriarDicionarioCorrelacoes = dicDados
    
End Function

Public Function CriarDicionarioTributacao(ByVal Plan As Worksheet, Optional ByVal tpTributacao As String) As Dictionary

Dim Chave As String, COD_ITEM$, CFOP$, UF$, COD_BARRA$, COD_NCM$
Dim arrCamposChave As New ArrayList
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim dicDados As New Dictionary
Dim Campos As Variant
    
    If Plan.AutoFilterMode Then Plan.AutoFilter.ShowAllData
    Set Dados = Util.DefinirIntervalo(Plan, 4, 3)
    
    If Dados Is Nothing Then Exit Function
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            'Carrega os campos chave do registro de tributação do item
            COD_ITEM = Campos(dicTitulos("COD_ITEM"))
            COD_BARRA = Campos(dicTitulos("COD_BARRA"))
            COD_NCM = Campos(dicTitulos("COD_NCM"))
            CFOP = Campos(dicTitulos("CFOP"))
            UF = Campos(dicTitulos("UF"))
            
            'Define chave de tributação da operação
            If tpTributacao = "COD_NCM" Then
                
                'Cria a chave de tributação por NCM
                Chave = COD_NCM & UF & CFOP
                
            ElseIf tpTributacao = "COD_BARRA" Then
                
                'Cria a chave de tributação por COD_BARRA
                Chave = COD_BARRA & UF & CFOP
                
            Else
                
                'Cria a chave de tributação geral
                Chave = COD_ITEM & UF & CFOP
                
            End If
            
            'Armazena a tributação da operação
            dicDados(Chave) = Campos
            
         End If
         
    Next Linha
    
    Set CriarDicionarioTributacao = dicDados
    
End Function

Public Function CriarDicionarioValorOperacoesC190(ByVal Plan As Worksheet) As Dictionary

Dim Dados As Range, Linha As Range
Dim Titulos As Variant, Registro, Campos, CampoChave
Dim dicDados As New Dictionary
Dim dicTitulos As New Dictionary
Dim arrCamposChave As New ArrayList
Dim Chave As String
Dim Valor As Double
    
    If Plan.AutoFilterMode Then Plan.AutoFilter.ShowAllData
    Set Dados = Util.DefinirIntervalo(Plan, 4, 3)
    
    If Dados Is Nothing Then Exit Function
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            Valor = Util.ValidarValores(Campos(dicTitulos("VL_OPR")))
            
            Chave = Campos(dicTitulos("CHV_PAI_FISCAL"))
            If dicDados.Exists(Chave) Then Valor = Valor + CDbl(dicDados(Chave))
            
            dicDados(Chave) = Valor
            
         End If
         
    Next Linha
    
    Set CriarDicionarioValorOperacoesC190 = dicDados
    
End Function

Public Function CarregarDadosRegistro(ByVal Plan As Worksheet, Optional ByVal FormatarSPED As Boolean, Optional SPEDContr As Boolean) As Dictionary

Dim Dados As Range, Linha As Range
Dim Titulos As Variant, Registro, Campos
Dim dicDados As New Dictionary
Dim dicTitulos As New Dictionary
Dim Chave As String, chReg$
Dim Comeco As Double
Dim i As Long, b&
    
    If Plan.AutoFilterMode Then Plan.AutoFilter.ShowAllData
    Set Dados = Util.DefinirIntervalo(Plan, 4, 3)
    If Not Dados Is Nothing Then
        
        b = 0
        Comeco = Timer
        
        Set dicTitulos = Util.MapearTitulos(Plan, 3)
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                Call Util.AntiTravamento(b, 100, "Carregando dados do registro " & VBA.Left(Plan.name, 4), Dados.Rows.Count, Comeco)
                If FormatarSPED Then Call FormatarCamposSPED(Campos, dicTitulos)
                If FormatarSPED Then Call ValidarCampos(Plan, Campos, dicTitulos, SPEDContr)
                
                Registro = Plan.CodeName
                Select Case True
                    
                    Case Registro Like "*0000*"
                        Chave = Campos(dicTitulos("CHV_REG"))
                        
                    Case SPEDContr
                        Chave = Campos(dicTitulos("CHV_PAI_CONTRIBUICOES"))
                        
                    Case Else
                        Chave = Campos(dicTitulos("CHV_PAI_FISCAL"))
                        
                End Select
                
                If Not dicDados.Exists(Chave) Then Set dicDados(Chave) = New Dictionary
                dicDados(Chave)(Campos(dicTitulos("CHV_REG"))) = Campos
                
            End If
            
        Next Linha
        
    End If
    
    Set CarregarDadosRegistro = dicDados
    
End Function

Public Function FormatarCamposSPED(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim Titulos As Variant
Dim i As Integer
    
    Titulos = dicTitulos.Keys()
    For i = LBound(Campos) To UBound(Campos)
        
        'Identificando e tratando os dados de data
        If VBA.InStr(1, Titulos(i - 1), "DT_") > 0 Then
            If Campos(i) <> "" Then Campos(i) = VBA.Format(fnExcel.FormatarData(Campos(i)), "ddmmyyyy")
        End If
        
        'Identificando e tratando os dados de alíquota
        If VBA.InStr(1, Titulos(i - 1), "ALIQ_") > 0 Then
            If Campos(i) <> "" Then Campos(i) = fnSPED.FormatarPercentuais(Campos(i))
        End If
        
        'Identificando e tratando os dados de data
        If VBA.InStr(1, Titulos(i - 1), "DESCR_") > 0 Then
            If Campos(i) <> "" Then Campos(i) = VBA.Trim(Campos(i))
        End If
        
        'Identificando e tratando os dados com casas decimais
        If VBA.InStr(1, Titulos(i - 1), "VL_") > 0 Then
            If Campos(i) <> "" Then Campos(i) = VBA.Round(Campos(i), 2)
        End If
        'Or Campos(i) Like "* – *"
        'Verifica se exitem enumerações e as retira
        If (VBA.InStr(1, Campos(i), " - ") > 0) And VBA.InStr(1, Titulos(i - 1), "TXT") = 0 And Titulos(i - 1) <> "CHV_REG" And Titulos(i - 1) <> "CHV_PAI_FISCAL" And Titulos(i - 1) <> "CHV_PAI_CONTRIBUICOES" Then
            Campos(i) = ValidacoesSPED.Fiscal.Enumeracoes.RemoverEnumeracoes(Campos(i), Titulos(i - 1))
        End If
    
    Next i
    
End Function

Public Function ValidarCampos(ByRef Plan As Worksheet, ByRef Campos As Variant, _
    ByRef dicTitulos As Dictionary, Optional SPEDContr As Boolean)

Dim Titulos As Variant
Dim i As Integer
    
    Titulos = dicTitulos.Keys()
    Select Case Plan.name
        
        Case "C100"
            If Campos(EncontrarColuna("COD_SIT", Titulos)) = "02" Or Campos(EncontrarColuna("COD_SIT", Titulos)) = "03" _
               Or Campos(EncontrarColuna("COD_SIT", Titulos)) = "04" Or Campos(EncontrarColuna("COD_SIT", Titulos)) = "05" Then
                
                For i = LBound(Campos) To UBound(Campos)
                    
                    If VBA.InStr(1, Titulos(i - 1), "DT_") > 0 Or VBA.InStr(1, Titulos(i - 1), "VL_") > 0 Or VBA.InStr(1, Titulos(i - 1), "IND_") > 0 _
                       And VBA.InStr(1, Titulos(i - 1), "OPER") = 0 And VBA.InStr(1, Titulos(i - 1), "EMIT") = 0 Then
                        Campos(i) = ""
                    End If
                    
                Next i
                
            ElseIf Campos(EncontrarColuna("COD_MOD", Titulos)) = "65" And Not SPEDContr Then
                
                For i = LBound(Campos) To UBound(Campos)
                    
                    Select Case True
                        
                        Case Titulos(i - 1) Like "*_ST", Titulos(i - 1) Like "*_IPI", Titulos(i - 1) Like "COD_PART", _
                             Titulos(i - 1) Like "*_PIS", Titulos(i - 1) Like "*_COFINS"
                            Campos(i) = ""
                            
                    End Select
                    
                Next i
                
            End If
            
    End Select
    
End Function

Public Function ValidarData(ByVal Data As String) As String

Dim NovaData As String
    
    If Data <> "" Then
        
        NovaData = VBA.Month(Data) & "/" & VBA.Day(Data) & "/" & VBA.Year(Data)
        If IsDate(NovaData) Then
            ValidarData = VBA.Format(CDate(NovaData), "ddmmyyyy")
        Else
            ValidarData = VBA.Format(CDate(Data), "ddmmyyyy")
        End If
        
    End If
    
End Function

Public Function CarregarRegistrosBloco(ByVal Bloco As String) As Dictionary

Dim Plan As Worksheet
Dim UltLin As Long
    
    'Call dicTitulos.RemoveAll
    
    For Each Plan In ThisWorkbook.Worksheets
            
        If VBA.Left(Plan.name, 1) = Bloco Then
            
            If Plan.AutoFilterMode Then Plan.AutoFilter.ShowAllData
            UltLin = Util.UltimaLinha(Plan, "A")
            If UltLin > 3 Then
                
                Select Case Plan.name
                                
                    Case "C140", "C141"
                        If ExportarC140Filhos Then Set dicRegistros(Plan.name) = Util.CarregarDadosRegistro(Plan, True)
                        
                    Case "C175_Contrib"
                        If ExportarC175Contruicoes Then Set dicRegistros(Plan.name) = Util.CarregarDadosRegistro(Plan, True)
                        
                    Case Else
                        
                        If VBA.Right(Plan.name, 3) > "001" And VBA.Right(Plan.name, 3) < "990" Then
                            Set dicRegistros(Plan.name) = Util.CarregarDadosRegistro(Plan, True)
                        End If
                
                End Select
                
            End If
            
        End If
        
    Next Plan
    
    Set CarregarRegistrosBloco = dicRegistros
    
End Function

Public Function CarregarRegistrosSPEDFiscal() As Dictionary

Dim Plan As Worksheet
Dim UltLin As Long
Dim NOME As String
    
    Call dicRegistros.RemoveAll
    For Each Plan In ThisWorkbook.Worksheets
        
        UltLin = Util.UltimaLinha(Plan, "A")
        If UltLin > 3 Then
            
            Select Case True
                
                Case Plan.name = "C140", Plan.name = "C141"
                    If ExportarC140Filhos Then Set dicRegistros(Plan.name) = Util.CarregarDadosRegistro(Plan, True)
                    
                Case Plan.name = "C175_Contrib"
                    If ExportarC175Contruicoes Then Set dicRegistros(Plan.name) = Util.CarregarDadosRegistro(Plan, True)
                    
                Case VBA.Left(Plan.CodeName, 3) = "reg" And VBA.Right(Plan.name, 3) <> "001" And VBA.Right(Plan.name, 3) <> "990"
                    Set dicRegistros(Plan.name) = Util.CarregarDadosRegistro(Plan, True)
                    
            End Select
            
        End If
            
    Next Plan
    
    Set CarregarRegistrosSPEDFiscal = dicRegistros
    
End Function

Public Function CarregarRegistrosSPED(ByRef dicEstruturaSPED As Dictionary, Optional ByVal SPEDContr As Boolean) As Dictionary

Dim REG As String, ARQUIVO$, CHV_REG$, CHV_PAI$
Dim dicRegistros As New Dictionary
Dim dicArquivo As New Dictionary
Dim Registro As Variant, Campos
Dim dicDados As New Dictionary
Dim Plan As Worksheet
Dim Comeco As Double
Dim b As Long
    
    b = 0
    Comeco = Timer
    For Each Registro In dicEstruturaSPED.Keys()
        
        Call Util.AntiTravamento(b, 1, "Verificando dados do registro " & Registro, dicEstruturaSPED.Count, Comeco)
        If Registro Like "C14*" And Not ExportarC140Filhos Then GoTo Prx:
        
        On Error GoTo Tratar:
            Set Plan = ThisWorkbook.Worksheets(Registro)
        On Error GoTo 0
        
        Set dicDados = Util.CarregarDadosRegistro(Plan, True, SPEDContr)
        Select Case True
            
            Case Registro Like "*001" Or Registro Like "*990"
                
                If dicDados.Count = 0 Then
                    
                    Set dicArquivo = Util.CriarDicionarioRegistro(Worksheets(dicEstruturaSPED.Keys(0)))
                    For Each Campos In dicArquivo.Items
                        
                        REG = VBA.Left(Registro, 4)
                        ARQUIVO = VBA.Replace(Campos(2), "'", "")
                        CHV_PAI = VBA.Replace(Campos(3), "'", "")
                        CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, CStr(REG))
                        
                        If Not dicDados.Exists(CHV_PAI) Then Set dicDados(CHV_PAI) = New Dictionary
                        
                        Select Case True
                            
                            Case SPEDContr And (Registro Like "0001" Or Registro Like "9001")
                                dicDados(CHV_PAI)(CHV_REG) = Array(REG, ARQUIVO, CHV_REG, "", CHV_PAI, "0")
                                
                            Case SPEDContr
                                dicDados(CHV_PAI)(CHV_REG) = Array(REG, ARQUIVO, CHV_REG, "", CHV_PAI, "1")
                                
                            Case Registro Like "0001" Or Registro Like "E001" Or Registro Like "9001"
                                dicDados(CHV_PAI)(CHV_REG) = Array(REG, ARQUIVO, CHV_REG, CHV_PAI, "", "0")
                                
                            Case Else
                                dicDados(CHV_PAI)(CHV_REG) = Array(REG, ARQUIVO, CHV_REG, CHV_PAI, "", "1")
                                
                        End Select
                        
                    Next Campos
                    
                End If
                
        End Select
        
        If dicDados.Count > 0 Then Set dicRegistros(Registro) = dicDados
Prx:
    Next Registro
    
    Set CarregarRegistrosSPED = dicRegistros
    
Exit Function
Tratar:
    Resume Next
    
End Function

Public Function IndexarCamposRegistros()

Dim Campo As Variant, Campos
Dim nCampos As Variant
Dim arrCampos As New ArrayList
    
    For Each nCampos In dicNomes.Keys()
        
        Campos = dicNomes(nCampos)
        Set dicNomes(nCampos) = CreateObject("System.Collections.ArrayList")
        For Each Campo In Campos
            dicNomes(nCampos).Add Campo
        Next Campo
        
    Next nCampos
    
End Function

Public Function MapearTitulos(ByVal Plan As Worksheet, ByVal LinhaTitulo As Long) As Dictionary

Dim dicTitulos As New Dictionary
Dim Titulos As Variant
Dim UltCol As Long
Dim i As Integer
    
    UltCol = Plan.Cells(LinhaTitulo, Plan.Columns.Count).END(xlToLeft).Column
    
    Titulos = Plan.Cells(LinhaTitulo, 1).Resize(, UltCol)
    Titulos = Application.index(Titulos, 0, 0)
    
    For i = LBound(Titulos) To UBound(Titulos)
        dicTitulos(Titulos(i)) = i
    Next i
    
    Set MapearTitulos = dicTitulos

End Function

Public Function MapearArray(ByVal Titulos As Variant) As Dictionary

Dim dicTitulos As New Dictionary
Dim i As Integer
    
    For i = LBound(Titulos) To UBound(Titulos)
        dicTitulos(Titulos(i)) = i
    Next i
    
    Set MapearArray = dicTitulos
    
Exit Function
Tratar:

End Function

Public Function FiltrarRegistros(ByRef PlanOrig As Worksheet, ByRef PlanDest As Worksheet, ByVal CamposOrig As Variant, ByVal CamposDest As Variant)

Dim dicTitulos As New Dictionary
Dim Titulos As Variant, Campos
Dim arrDados As New ArrayList
Dim Intervalo As Range
Dim i As Integer
    
    If VarType(CamposOrig) = vbString Then CamposOrig = Array(CamposOrig)
    If VarType(CamposDest) = vbString Then CamposDest = Array(CamposDest)
    
    If PlanDest.AutoFilterMode Then PlanDest.AutoFilter.ShowAllData
    
    Set Intervalo = Util.DefinirIntervalo(PlanDest, 3, 3)
    Set dicTitulos = Util.MapearTitulos(PlanDest, 3)
    
    For i = LBound(CamposDest) To UBound(CamposDest)
        
        Set arrDados = Util.ObterCampoEspecificoFiltrado(PlanOrig, 4, 3, CamposOrig(i))
        
        If arrDados.Count = 1 Then
            
            Intervalo.AutoFilter Field:=dicTitulos(CamposDest(i)), Criteria1:=Array(CStr(arrDados.item(0))), Operator:=xlFilterValues
            
        ElseIf arrDados.Count > 1 Then
            
            Campos = arrDados.toArray()
            Intervalo.AutoFilter Field:=dicTitulos(CamposDest(i)), Criteria1:=Campos, Operator:=xlFilterValues
            
        End If
        
        arrDados.Clear
        
    Next i
    
    If Util.UltimaLinha(PlanDest, "A") > 3 Then
        Call Application.GoTo(PlanDest.[C3], True)
    Else
        Call Util.MsgAlerta("Não existem dados filtrados na planilha de destino.", "Filtragem sem Dados")
    End If
    
End Function

Public Function ObterCampoEspecificoFiltrado(ByRef Plan As Worksheet, ByVal Lini As Long, ByVal LinT As Long, ByVal Campo As String) As ArrayList
    
Dim Intervalo As Range, Area As Range, Cel As Range
Dim dicTitulos As New Dictionary
Dim arrCampos As New ArrayList
Dim UltLin As Long, col&
Dim Valor As Variant
    
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    UltLin = Util.UltimaLinha(Plan, "A")
    col = dicTitulos(Campo)
    
    If UltLin > 3 Then
        
        Set Intervalo = Plan.Cells(Lini, 1).Resize(UltLin - Lini + 1, Plan.Cells(LinT, 1).END(xlToRight).Column - 1).SpecialCells(xlCellTypeVisible)
        
        If Intervalo.Areas.Count > 0 Then
            
            For Each Area In Intervalo.Areas
                
                For Each Cel In Area.Columns(col).Cells
                    Valor = fnExcel.FormatarCampoFiltrado(Campo, Cel.value)
                    If Not arrCampos.contains(Valor) Then arrCampos.Add CStr(Valor)
                Next Cel
                
            Next Area
            
        End If
        
    End If
    
    Set ObterCampoEspecificoFiltrado = arrCampos
    
End Function

Function TransformarDicionarioArrayList(ByRef Dicionario As Dictionary) As ArrayList
    
Dim arrDados As New ArrayList
Dim Chave As Variant
    
    For Each Chave In Dicionario.Keys()
        arrDados.Add Chave
    Next Chave
    
    Set TransformarDicionarioArrayList = arrDados

End Function

Function TransformarArrayListDicionario(ByRef ArrayList As ArrayList) As Dictionary
    
Dim dicDados As New Dictionary
Dim Titulos As Variant, Titulo
Dim i As Long

    'Titulos = ArrayList.toArray()
    For i = 1 To ArrayList.Count
        dicDados(ArrayList(i - 1)) = i
    Next i
    
    Set TransformarArrayListDicionario = dicDados

End Function

Public Function TratarParticularidadesRegistros(ByRef dicRegistros As Dictionary)

Dim Registro As Variant, Campos, Chave
Dim dicC100 As New Dictionary
Dim dicC170 As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicChaves As New Dictionary
Dim arrDocs65 As New ArrayList
Dim arrDocsProprios As New ArrayList
    
    For Each Registro In dicRegistros.Keys()
        
        Select Case True
            Case Registro = "C170" And ExportarC170Proprios = False
                Set dicC100 = dicRegistros("C100")
                Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
                
                If dicC100.Count > 0 Then
                    For Each Chave In dicC100.Keys()
                        Set dicChaves = dicC100(Chave)
                        For Each Campos In dicChaves.Items
                            If Campos(dicTitulosC100("IND_EMIT")) = "0" Then arrDocsProprios.Add Campos(dicTitulosC100("CHV_REG"))
                        Next Campos
                    Next Chave
                End If
                
                Set dicC170 = dicRegistros(Registro)
                Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
                If dicC170.Count > 0 Then
                    
                    For Each Chave In dicC170.Keys()
                        If arrDocsProprios.contains(Chave) Then Call dicC170.Remove(Chave)
                    Next Chave
                    
                    Set dicRegistros(Registro) = dicC170
                    
                End If
                
        End Select
        
    Next Registro
    
End Function

Public Function TratarParticularidadesEFDContribuicoes(ByRef dicRegistros As Dictionary)

Dim Registro As Variant, Campos, Chave, chReg
Dim dicRegSel As New Dictionary
Dim dicC100 As New Dictionary
Dim dicC170 As New Dictionary
Dim dicTitulosRegSel As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicChaves As New Dictionary
Dim arrDocs65 As New ArrayList
Dim totRegistros As Long
Dim arrDocsProprios As New ArrayList
Dim COD_SIT As String, IND_OPER$, IND_MOV$
    
    For Each Registro In dicRegistros.Keys()
        
        Select Case True
            
            Case Registro = "C100"
                Set dicRegSel = dicRegistros(Registro)
                Set dicTitulosRegSel = Util.MapearTitulos(regC100, 3)
                If dicRegSel.Count > 0 Then
                    
                    For Each Chave In dicRegSel.Keys()
                        
                        Set dicChaves = dicRegSel(Chave)
                        For Each chReg In dicChaves.Keys
                            
                            Campos = dicChaves(chReg)
                            If Util.ChecarCamposPreenchidos(Campos) Then
                            
                                COD_SIT = Campos(dicTitulosRegSel("COD_SIT"))
                                Select Case VBA.Left(COD_SIT, 2)
                                    
                                    Case "02", "03", "04", "05"
                                        Call dicChaves.Remove(chReg)
                                        
                                End Select
                                
                            End If
                            
                        Next chReg
                        
                    Next Chave
                    
                    Set dicRegistros(Registro) = dicRegSel
                    
                End If
                
            Case Registro = "C170"
                Set dicC100 = dicRegistros("C100")
                Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
                If dicC100.Count > 0 Then
                    
                    For Each Chave In dicC100.Keys()
                        
                        Set dicChaves = dicC100(Chave)
                        For Each Campos In dicChaves.Items
                            If Campos(dicTitulosC100("COD_MOD")) = "65" Then arrDocs65.Add Campos(dicTitulosC100("CHV_REG"))
                        Next Campos
                        
                    Next Chave
                    
                End If
                
                Set dicC170 = dicRegistros(Registro)
                Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
                If dicC170.Count > 0 Then
                    
                    For Each Chave In dicC170.Keys()
                        If arrDocs65.contains(Chave) Then Call dicC170.Remove(Chave)
                    Next Chave
                    
                    Set dicRegistros(Registro) = dicC170
                    
                End If
            
            Case Registro = "D100"
                Set dicRegSel = dicRegistros(Registro)
                Set dicTitulosRegSel = Util.MapearTitulos(regD100, 3)
                If dicRegSel.Count > 0 Then
                    
                    For Each Chave In dicRegSel.Keys()
                        
                        Set dicChaves = dicRegSel(Chave)
                        For Each chReg In dicChaves.Keys
                            
                            Campos = dicChaves(chReg)
                            IND_OPER = Campos(dicTitulosRegSel("IND_OPER"))
                            If IND_OPER Like "1*" Then Call dicChaves.Remove(chReg)
                            
                        Next chReg
                        
                    Next Chave
                    
                    Set dicRegistros(Registro) = dicRegSel
                    
                End If
                
            Case Registro = "1001"
                Set dicRegSel = dicRegistros(Registro)
                Set dicTitulosRegSel = Util.MapearTitulos(reg1001, 3)
                
                For Each Chave In dicRegistros.Keys()
                    
                    If Chave Like "1*" Then
                        
                        totRegistros = totRegistros + dicRegistros(Chave).Count
                        If totRegistros > 2 Then Exit For
                        
                    End If
                    
                Next Chave
                
                If dicRegSel.Count > 0 Then
                    
                    For Each Chave In dicRegSel.Keys()
                        
                        Set dicChaves = dicRegSel(Chave)
                        For Each chReg In dicChaves.Keys
                            
                            Campos = dicChaves(chReg)
                                
                                IND_MOV = "1"
                                Campos(dicTitulosRegSel("IND_MOV")) = IND_MOV
                                
                            dicChaves(chReg) = Campos
                            
                        Next chReg
                        
                    Next Chave
                    
                    Set dicRegistros(Registro) = dicRegSel
                    
                End If
                
        End Select
        
    Next Registro
    
End Function

Public Function ChecarAusenciaDados(ByRef Plan As Worksheet, Optional OmitirMsg As Boolean = True, Optional Msg As String) As Boolean
    
Dim UltLin As Long
    
    UltLin = UltimaLinha(Plan, "A")
    If UltLin <= 4 Then
    
        ChecarAusenciaDados = True
        If Not OmitirMsg Then Call MsgAlerta_AusenciaDados(Msg)
    
    End If
    
End Function

Public Function ChecarCamposPreenchidos(ByVal Campos As Variant) As Boolean
    If VBA.Join(Campos, "") <> "" Then ChecarCamposPreenchidos = True
End Function

Public Function CarregarChavesAcessoDoces(ByRef arrChaves As ArrayList)
    
Dim Dados As Range, Linha As Range
Dim arrAuxiliar As New ArrayList
Dim Planilha As Worksheet
Dim Planilhas As New Collection
Dim Campos As Variant
Dim Chave As String
    
    Planilhas.Add EntNFe
    Planilhas.Add EntCTe
    Planilhas.Add SaiNFe
    Planilhas.Add SaiNFCe
    Planilhas.Add SaiCTe
    Planilhas.Add SaiCFe
    
    For Each Planilha In Planilhas
        
        With Planilha
            
            If .AutoFilterMode Then .AutoFilter.ShowAllData
            
            Set arrAuxiliar = Util.ListarValoresUnicos(Planilha, 4, 3, "Chave de Acesso")
            arrChaves.addRange arrAuxiliar
            
        End With
        
        arrAuxiliar.Clear
        
    Next Planilha
    
End Function

Public Function AtualizarSituacaoDoce(ByRef Plan As Worksheet, ByRef dicDados As Dictionary, _
    Optional ByRef arrCanceladas As ArrayList, Optional ByRef dicChavesReferenciadas As Dictionary)
    
Dim ChavesReferenciadas As Variant
Dim dicTitulos As New Dictionary
Dim Campos As Variant, Chave
    
    'Carrega os títulos da planilha selecionada
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    
    'Percorre os registros e faz verificações para atualizar
    For Each Chave In dicDados.Keys()
        
        'Carrega os campos do registro atual
        Campos = dicDados(Chave)
        
        'Verifica se a nota está cancelada
        If Not arrCanceladas Is Nothing Then
            If arrCanceladas.contains(Campos(dicTitulos("Chave de Acesso"))) Then
                Campos(dicTitulos("Valor")) = 0
                Campos(dicTitulos("Situacao")) = "Cancelada"
            End If
        End If
        
        'Verifica se a nota foi referenciada
        If Not dicChavesReferenciadas Is Nothing Then
            If dicChavesReferenciadas.Exists(Chave) Then
                ChavesReferenciadas = dicChavesReferenciadas(Chave)
                
                Select Case True
                    
                    Case UBound(ChavesReferenciadas) > 0
                        Campos(dicTitulos("Observações")) = _
                            "CHAVES REFERENCIADAS: " & VBA.Join(ChavesReferenciadas, ", ")
                    
                    Case UBound(ChavesReferenciadas) = 0
                        Campos(dicTitulos("Observações")) = _
                            "CHAVE REFERENCIADA: " & VBA.Join(ChavesReferenciadas)
                            
                End Select
                
            End If
            
        End If
        
        dicDados(Chave) = Campos
        
    Next Chave
    
End Function

Public Function FormatarNomePersonalizado(ByVal NOME As String)
    On Error Resume Next
        FormatarNomePersonalizado = VBA.Replace(VBA.Replace(Names(NOME), "=", ""), """", "")
    On Error GoTo 0
End Function

Public Function AtualizarBarraStatus(ByVal Msg As Variant)
    Application.StatusBar = Msg
    DoEvents
End Function

Public Function PararProcesso() As Boolean
    If InterromperProcesso Then
        PararProcesso = True
        InterromperProcesso = False
    End If
End Function

Public Function ApenasNumeros(ByVal Valor As String) As String

Dim Caract As String, NUM$
Dim i As Long
    
    For i = 1 To VBA.Len(Valor)
        Caract = VBA.Mid(Valor, i, 1)
        If IsNumeric(Caract) Then NUM = NUM & Caract
    Next i
        
    ApenasNumeros = NUM
    
End Function

Function AlterarChaveAcesso(ByVal Chave As String, ByVal CNPJEmit As String, ByVal CNPJDeclarante As String) As String

Dim soma As Integer
Dim resto As Integer
Dim Digito As Integer
Dim i As Integer
Dim Peso As Integer
Dim CNPJ As String
    
    soma = 0
    Peso = 2
    CNPJ = VBA.Mid(Chave, 7, 14)
    
    If CNPJ = CNPJDeclarante Then
    
        Chave = VBA.Replace(Chave, CNPJ, CNPJEmit)
        Chave = VBA.Left(Chave, 43)
        
        'Loop reverso na chave, começando do último dígito
        For i = Len(Chave) To 1 Step -1
            soma = soma + (Mid(Chave, i, 1) * Peso)
            Peso = Peso + 1
            If Peso > 9 Then Peso = 2
        Next i
        
        resto = soma Mod 11
        
        If resto < 2 Then
            Digito = 0
        Else
            Digito = 11 - resto
        End If
        
        Chave = Chave & Digito
    
    End If
    
    AlterarChaveAcesso = Chave
    
End Function

Public Function GerarJson(ByVal Arq As String) As String

Dim Registros As Variant, Registro
Dim Conteudo As String, json$
Dim fso As FileSystemObject
Dim ARQUIVO As TextStream
    
    Set fso = New FileSystemObject
    Set ARQUIVO = fso.OpenTextFile(Arq, ForReading, False)
    Conteudo = ARQUIVO.ReadAll
    ARQUIVO.Close
    
    Registros = VBA.Split(Conteudo, vbCrLf)
                
    json = "["""
        
        For Each Registro In Registros
        
            If Registro <> "" Then
                json = json & VBA.Replace(Registro, """", "\""") & """, """
            End If
            
            If VBA.Left(Registro, 5) = "|9999" Then Exit For
            
        Next Registro
        
    GerarJson = VBA.Left(json, VBA.Len(json) - 4) & """]"
    
End Function

Public Function GravarSugestao(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal INCONSISTENCIA As String, _
    ByVal SUGESTAO As String, Optional ByRef dicInconsistenciasIgnoradas As Dictionary)

Dim i As Integer
Dim InconsistenciaIgnorada As Variant
Dim CHV_REG As String
    
    If Not dicInconsistenciasIgnoradas Is Nothing Then
        
        CHV_REG = Campos(dicTitulos("CHV_REG"))
        If dicInconsistenciasIgnoradas.Exists(CHV_REG) Then
            
            For Each InconsistenciaIgnorada In dicInconsistenciasIgnoradas(CHV_REG)
                
                If InconsistenciaIgnorada = INCONSISTENCIA Or InconsistenciaIgnorada = "O XML dessa operação não foi importado" Then
                    Exit Function
                End If
                
            Next InconsistenciaIgnorada
            
        End If
        
    End If
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    Campos(dicTitulos("INCONSISTENCIA") - i) = INCONSISTENCIA
    Campos(dicTitulos("SUGESTAO") - i) = SUGESTAO
    
End Function

Public Function SelecionarChaveSPED(ByVal CHV_BASE As String, ByVal CNPJBase As String, _
    ByVal CNPJEmit As String, ByVal CNPJDest As String, ByVal SPEDContrib As Boolean) As String
    
    If SPEDContrib Then
        
        Select Case True
            
            Case CNPJEmit Like CNPJBase & "*"
                CHV_BASE = fnSPED.GerarChaveRegistro(CHV_BASE, CNPJEmit)
                
            Case CNPJDest Like CNPJBase & "*"
                CHV_BASE = fnSPED.GerarChaveRegistro(CHV_BASE, CNPJDest)
                
        End Select
        
    End If
    
    SelecionarChaveSPED = CHV_BASE
    
End Function

Function ValidarIEBahia(IE As String) As String

Dim Peso() As Variant
Dim soma As Integer
Dim i As Integer
Dim resto As Integer
Dim Digito As Integer
Dim Tamanho As Integer
Dim Base As Integer
Dim PosicaoDigito As Integer

    ' Remove caracteres não numéricos
    IE = ApenasNumeros(IE)

    ' Remove zeros à esquerda se necessário
    IE = VBA.LTrim(Str(Val(IE)))

    'Define tamanho e base dependendo do comprimento da inscrição estadual
    Tamanho = VBA.Len(IE)
    If Tamanho = 8 Then
        Peso = Array(7, 6, 5, 4, 3, 2)
    ElseIf Tamanho = 9 Then
        Peso = Array(8, 7, 6, 5, 4, 3, 2)
    Else
        ValidarIEBahia = ""
        Exit Function
    End If
    
    Base = UBound(Peso)
    
    ' Calcula o dígito verificador
    soma = 0
    For i = 1 To Base
        soma = soma + CInt(VBA.Mid(IE, i, 1)) * Peso(i)
    Next i

    resto = soma Mod 10
    If resto = 0 Then
        Digito = 0
    Else
        Digito = 10 - resto
    End If

    ' Compara o dígito verificador calculado com o dígito informado na inscrição
    PosicaoDigito = Base
    If Digito = CInt(VBA.Mid(IE, PosicaoDigito, 1)) Then
        ValidarIEBahia = IE
    Else
        If VBA.Left(IE, 1) = "0" Then
            ' Tenta novamente sem o zero à esquerda se a primeira tentativa falhar
            ValidarIEBahia = ValidarIEBahia(Mid(IE, 2))
        Else
            ValidarIEBahia = ""
        End If
    End If
    
End Function

Public Sub CarregarTabelaCEST(ByRef dicDadosCEST As Dictionary)

Dim TabelaCEST As String, CEST$, DESC$, DT_INI$, DT_FIM$
Dim Registros As Variant, Registro, Campos
Dim CustomPart As New clsCustomPartXML
    
    TabelaCEST = CustomPart.ExtrairTXTPartXML("TabelaCEST")
    
    Registros = VBA.Split(TabelaCEST, vbLf)
    For Each Registro In Registros
        
        If Not Registro Like "*COD_CEST*" Then
            
            Campos = VBA.Split(Registro, "|")
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                CEST = Campos(0)
                DESC = Campos(1)
                DT_INI = Campos(2)
                DT_FIM = Campos(3)
                
                If DT_INI <> "" Then DT_INI = Util.FormatarData(DT_INI)
                If DT_FIM <> "" Then DT_FIM = Util.FormatarData(DT_FIM)
                
                If CEST <> "CEST" Then dicDadosCEST(CEST) = Array(CEST, DESC, DT_INI, DT_FIM)
                
            End If
        
        End If
        
    Next Registro
    
End Sub

Public Function CarregarTabelaNCM(ByRef dicTabelaNCM As Dictionary)

Dim CustomPart As New clsCustomPartXML
    
    Call dicTabelaNCM.RemoveAll
    Set dicTabelaNCM = JsonConverter.ParseJson(CustomPart.ExtrairJsonXmlPart("TabelaNCM"))
    
End Function

Public Function CriarLista(ByRef Campos As Variant) As ArrayList
    
Dim arrLista As New ArrayList
Dim Campo As Variant
    
    For Each Campo In Campos
        arrLista.Add Campo
    Next Campo
    
    Set CriarLista = arrLista
    
End Function

Public Function RemoverAspaSimples(ByVal Texto As String, Optional ByVal Caracteres As Byte)
    
    Texto = VBA.Replace(Texto, "'", "")
    Texto = VBA.Trim(Texto)
    
    If Caracteres > 0 Then Texto = VBA.Left(Texto, Caracteres)
    
    RemoverAspaSimples = Texto
    
End Function

Public Function LimparTexto(Texto As String, Optional LimiteCaracteres As Integer) As String

Dim i As Integer
Dim Caractere As String
Dim TextoLimpo As String
Dim arrPermitidos As New ArrayList
    
    'Carrega
    Call CarregarCaracteresPermitidos(arrPermitidos)
    
    'Inicialize a string limpa
    TextoLimpo = ""
    
    'Percorre cada caractere do texto de entrada
    For i = 1 To Len(Texto)
        
        Caractere = Mid(Texto, i, 1)
        
        ' Verifica se o caractere está na lista de caracteres permitidos
        If arrPermitidos.contains(Caractere) Then
            TextoLimpo = TextoLimpo & Caractere
        End If
        
    Next i
    
    'Remover espaços duplicados
    TextoLimpo = RemoverQuebrasLinhaEspacosDuplicados(TextoLimpo)
    
    If LimiteCaracteres > 0 Then TextoLimpo = VBA.Left(TextoLimpo, LimiteCaracteres)
        
    'Retorna o texto limpo
    LimparTexto = TextoLimpo
    
End Function

Private Function CarregarCaracteresPermitidos(ByRef arrPermitidos As ArrayList)

Dim CaracteresPermitidos As String, Caractere$
Dim i As Integer

    CaracteresPermitidos = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789" & _
                           "áéíóúãõâêîôûàèìòùäëïöüçÁÉÍÓÚÃÕÂÊÎÔÛÀÈÌÒÙÄËÏÖÜÇ " & _
                           "!@#$%¨&*()_+=-[]{};:'<>,.?/`~^"
        
    For i = 1 To VBA.Len(CaracteresPermitidos)
    
        Caractere = VBA.Mid(CaracteresPermitidos, i, 1)
        If Not arrPermitidos.contains(Caractere) Then arrPermitidos.Add Caractere
        
    Next i
    
End Function

Private Function RemoverQuebrasLinhaEspacosDuplicados(Texto As String) As String

Dim TextoLimpo As String
    
    TextoLimpo = Texto
    
    ' Remove quebras de linha
    TextoLimpo = Replace(TextoLimpo, Chr(10), " ")
    TextoLimpo = Replace(TextoLimpo, Chr(13), " ")
    
    ' Substitui múltiplos espaços por um único espaço
    Do While InStr(TextoLimpo, "  ") > 0
        TextoLimpo = Replace(TextoLimpo, "  ", " ")
    Loop
    
    RemoverQuebrasLinhaEspacosDuplicados = TextoLimpo
    
End Function

Public Function GerarModeloRelatorio(ByRef Plan As Worksheet, ByVal NomePlan As String)

Dim Campos As Variant, TitulosParaRemover, Titulo
Dim dicTitulos As New Dictionary
    
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    
    'Lista campos a serem removidos
    TitulosParaRemover = Array("INCONSISTENCIA", "SUGESTAO", "REG", "ARQUIVO", "CHV_REG", "CHV_PAI_FISCAL")
    For Each Titulo In TitulosParaRemover
        
        'Verifica existência do campo e o remove e o remove caso exista
        If dicTitulos.Exists(Titulo) Then Call dicTitulos.Remove(Titulo)
        
    Next Titulo
    
    Campos = dicTitulos.Keys()
    
    Workbooks.Add
    ActiveSheet.name = NomePlan
    With ActiveSheet.Range("A1").Resize(, UBound(Campos) + 1)
        
        .value = Campos
        .Font.Bold = True
        .Columns.AutoFit
        
    End With
    
End Function

Public Function ExportarDadosRelatorio(ByRef PlanOrig As Worksheet, ByVal NomePlan As String) As Boolean

Dim Campos As Variant
Dim Plan As Worksheet
Dim dicTitulos As New Dictionary
    
    On Error GoTo Tratar:
    
    With PlanOrig
        
        If .AutoFilter.FilterMode Then .ShowAllData
        Set dicTitulos = Util.MapearTitulos(PlanOrig, 3)
        
    End With
    
    'Adiciona uma noa pasta de trabalho
    Workbooks.Add
    
    'Seta planilha da nova pasta de trabalho
    Set Plan = ActiveSheet
    
    'Define nome da nova planilha
    Plan.name = NomePlan
    
    'Copiar dados da planilha atual para a nova pasta de trabalho
    PlanOrig.Range("A3").CurrentRegion.Copy Plan.Range("A1")
    
    'Elimina as linhas 1 e 2 para remover filtros e subtotais
    Plan.Rows("1:2").Delete
    
    On Error Resume Next
    'Remove as colunas desnecessárias
    Plan.Columns(dicTitulos("SUGESTAO")).Delete
    Plan.Columns(dicTitulos("INCONSISTENCIA")).Delete
    Plan.Columns(dicTitulos("CHV_PAI_FISCAL")).Delete
    Plan.Columns(dicTitulos("CHV_REG")).Delete
    Plan.Columns(dicTitulos("ARQUIVO")).Delete
    Plan.Columns(dicTitulos("ITEM")).Delete
    
    'Redimensiona colunas
    Plan.Cells.Columns.AutoFit
        
    ExportarDadosRelatorio = True

Exit Function
Tratar:

End Function

Function GerarCNPJCompleto(CNPJBase As String, estabelecimento As String) As String

Dim cnpjCompleto As String
Dim pesos1 As Variant
Dim pesos2 As Variant
Dim soma As Long
Dim resto As Long
Dim digito1 As String
Dim digito2 As String
Dim i As Integer
    
    'Pesos para o cálculo dos dígitos
    pesos1 = Array(5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    pesos2 = Array(6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    
    'Remover caracteres não numéricos do CNPJ base e do estabelecimento
    CNPJBase = Replace(CNPJBase, ".", "")
    CNPJBase = Replace(CNPJBase, "-", "")
    CNPJBase = Replace(CNPJBase, "/", "")
    
    estabelecimento = Replace(estabelecimento, ".", "")
    estabelecimento = Replace(estabelecimento, "-", "")
    estabelecimento = Replace(estabelecimento, "/", "")
    
    'Verificar se o CNPJ base tem 8 dígitos e o estabelecimento tem 4 dígitos
    If Len(CNPJBase) <> 8 Or Len(estabelecimento) <> 4 Then
        GerarCNPJCompleto = "CNPJ Base ou Estabelecimento inválido"
        Exit Function
    End If
    
    'Concatenar CNPJ base com o estabelecimento
    cnpjCompleto = CNPJBase & estabelecimento
    
    'Cálculo do primeiro dígito verificador
    soma = 0
    For i = 1 To 12
        soma = soma + CInt(Mid(cnpjCompleto, i, 1)) * pesos1(i)
    Next i
    
    resto = soma Mod 11
    
    If resto < 2 Then
        digito1 = "0"
    Else
        digito1 = CStr(11 - resto)
    End If
    
    'Cálculo do segundo dígito verificador
    soma = 0
    cnpjCompleto = cnpjCompleto & digito1
    For i = 1 To 13
        soma = soma + CInt(Mid(cnpjCompleto, i, 1)) * pesos2(i)
    Next i
    resto = soma Mod 11
    If resto < 2 Then
        digito2 = "0"
    Else
        digito2 = CStr(11 - resto)
    End If
    
    'Retornar o CNPJ completo com os dígitos verificadores
    GerarCNPJCompleto = Left(cnpjCompleto, 12) & digito1 & digito2
    
End Function

Public Function ExtrairRegistroDicionario(ByRef dicDados As Dictionary, ParamArray CamposChave() As Variant)
    
Dim CHV_REG As String
    
    CHV_REG = VBA.Join(CamposChave)
    If dicDados.Exists(CHV_REG) Then ExtrairRegistroDicionario = dicDados(CHV_REG)
    
End Function

Public Function AtualizarDicionario(ByRef dicDados As Dictionary, ByRef dicAuxiliar As Dictionary)

Dim Chave As Variant
    
    For Each Chave In dicAuxiliar: dicDados(Chave) = dicAuxiliar(Chave): Next Chave
    
End Function

Public Function ListarValoresOcorrenciaUnica(ByRef Plan As Worksheet, ByVal Lini As Long, ByVal LinT As Long, ByVal Campo As String) As ArrayList

Dim UltLin As Long, col&
Dim dicTitulos As New Dictionary
Dim Intervalo As Range, Cel As Range
Dim dicOcorrencias As New Dictionary
Dim arrOcorrencia As New ArrayList
Dim Campos As Variant, Valor, Chave, Ocorrencia
    
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    UltLin = Util.UltimaLinha(Plan, "A")
    
    If UltLin > 3 Then
        
        col = dicTitulos(Campo)
        Set Intervalo = Plan.Cells(Lini, col).Resize(UltLin - Lini + 1)
        Campos = Application.Transpose(Intervalo)
        
        If VarType(Campos) = vbString Then Campos = Array(Campos)
        For Each Valor In Campos
            If Not dicOcorrencias.Exists(CStr(Valor)) Then _
                dicOcorrencias.Add CStr(Valor), 1 Else _
                    dicOcorrencias(CStr(Valor)) = dicOcorrencias(CStr(Valor)) + 1
        Next Valor
        
    End If
    
    For Each Chave In dicOcorrencias.Keys()
        
        If dicOcorrencias(Chave) = 1 Then arrOcorrencia.Add Chave
        
    Next Chave
    
    Set ListarValoresOcorrenciaUnica = arrOcorrencia
    
End Function

Public Function ListarValoresUnicos(ByRef Plan As Worksheet, ByVal Lini As Long, ByVal LinT As Long, ByVal Campo As String) As ArrayList

Dim UltLin As Long, col&
Dim dicTitulos As New Dictionary
Dim Intervalo As Range, Cel As Range
Dim arrCampos As New ArrayList
Dim Campos As Variant, Valor
    
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    UltLin = Util.UltimaLinha(Plan, "A")
    
    If UltLin > 3 Then
        
        col = dicTitulos(Campo)
        Set Intervalo = Plan.Cells(Lini, col).Resize(UltLin - Lini + 1)
        Campos = Application.Transpose(Intervalo)
        
        If VarType(Campos) = vbString Then Campos = Array(Campos)
        For Each Valor In Campos
            If Not arrCampos.contains(CStr(Valor)) Then arrCampos.Add CStr(Valor)
        Next Valor
        
    End If
    
    arrCampos.Sort
    
    Set ListarValoresUnicos = arrCampos
    
End Function

Public Function ListarValoresUnicosDicionario(ByRef dicDados As Dictionary, ByVal dicTitulos As Dictionary, ByVal Campo As String) As ArrayList

Dim arrCampos As New ArrayList
Dim Campos As Variant, Valor
Dim i As Byte
    
    For Each Campos In dicDados.Items()
        
        If LBound(Campos) = 0 Then i = 1
        Valor = Campos(dicTitulos("CFOP") - i)
        
        If Not arrCampos.contains(Valor) Then arrCampos.Add Valor
        
    Next Campos
    
    arrCampos.Sort
    Set ListarValoresUnicosDicionario = arrCampos
    
End Function

Public Function SomarValoresDicionario(ByRef dicDados As Dictionary, ByRef Campos As Variant, ByVal CHV_REG As String)
    
Dim i As Byte
Dim dicCampos As Variant
    
    If dicDados.Exists(CHV_REG) Then
        
        dicCampos = dicDados(CHV_REG)
        For i = LBound(Campos) To UBound(Campos)
            
            If VarType(Campos(i)) = vbDouble Then Campos(i) = CDbl(Campos(i)) + CDbl(dicCampos(i))
            
        Next i
        
    End If
    
    dicDados(CHV_REG) = Campos
    
End Function

Public Function AlterarCampoDicionario(ByRef dicDados As Dictionary, ByRef dicTitulos As Dictionary, _
    ByVal CHV_REG As String, ByVal nCampo As String, ByVal Valor As Variant)
    
Dim dicCampos As Variant
Dim i As Byte

    If dicDados.Exists(CHV_REG) Then
        
        dicCampos = dicDados(CHV_REG)
        If LBound(dicCampos) = 0 Then i = 1
        
        dicCampos(dicTitulos(nCampo) - i) = Valor
        
        dicDados(CHV_REG) = dicCampos
        
    End If
    
End Function

Public Function SegregarEntradasSaidas(ByRef dicDados As Dictionary, ByRef dicTitulos As Dictionary, _
    ByRef dicEntradas As Dictionary, ByRef dicSaidas As Dictionary)
    
Dim Campos As Variant
Dim CHV_REG As String, CFOP$
    
    For Each Campos In dicDados.Items()
        
        CHV_REG = Campos(dicTitulos("CHV_REG"))
        CFOP = Campos(dicTitulos("CFOP"))
        
        If CFOP < 4000 Then dicEntradas(CHV_REG) = Campos Else dicSaidas(CHV_REG) = Campos
        
    Next Campos
    
End Function

Public Function ConfirmarOperacao(Msg As String, Titulo As String) As VbMsgBoxResult

Dim vbResult As VbMsgBoxResult

    vbResult = MsgBox(Msg, vbYesNo + vbExclamation, Titulo)
    
End Function

Public Function AgruparRegistros(ByRef Plan As Worksheet)

Dim dicAuxiliar As New Dictionary
Dim dicTitulos As New Dictionary
Dim Dados As Range, Linha As Range
Dim Campos As Variant
Dim CHV_REG As String
    
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    
    Set Dados = Util.DefinirIntervalo(Plan, 4, 3)
    If Dados Is Nothing Then Exit Function
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Agrupando dados do registro " & Plan.name & ", por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            CHV_REG = Campos(dicTitulos("CHV_REG"))
            Call Util.SomarValoresDicionario(dicAuxiliar, Campos, CHV_REG)
            
        End If
        
    Next Linha
    
    Call Util.LimparDados(Plan, 4, False)
    Call Util.ExportarDadosDicionario(Plan, dicAuxiliar)
    
End Function

Public Function RemoverNamespaces(ByVal Tag As String) As String

Dim i As Byte
    
    For i = 1 To 6: Tag = VBA.Replace(Tag, "ns" & i & ":", ""): Next i
    RemoverNamespaces = Tag
    
End Function

Public Function SomarValoresCamposSelecionados(ByRef dicDados As Dictionary, ByRef dicTitulos As Dictionary, _
    ByRef Titulos As Variant, ByRef Campos As Variant, ByVal CHV_REG As String)
    
Dim i As Byte, Posicao As Byte
Dim dicCampos As Variant, Titulo
        
    If dicDados.Exists(CHV_REG) Then
        
        dicCampos = dicDados(CHV_REG)
        For Each Titulo In Titulos
            
            Posicao = dicTitulos(Titulo)
            Campos(Posicao) = CDbl(Campos(Posicao)) + CDbl(dicCampos(Posicao))
            
        Next Titulo
        
    End If
    
    dicDados(CHV_REG) = Campos
    
End Function

Function RemoverCaracteresNaoImprimiveis(ByVal Texto As String) As String
    
Dim i As Long
Dim resultado As String

    For i = 1 To VBA.Len(Texto)
        If VBA.Asc(VBA.Mid(Texto, i, 1)) > 31 And VBA.Asc(VBA.Mid(Texto, i, 1)) <> 127 Then
            resultado = resultado & VBA.Mid(Texto, i, 1)
        End If
    Next i
    
    RemoverCaracteresNaoImprimiveis = resultado
    
End Function

Public Function PossuiDadosInformados(ByRef Plan As Worksheet) As Boolean

Dim UltLin As Long
    
    UltLin = UltimaLinha(Plan, "A")
    If UltLin > 3 Then PossuiDadosInformados = True
    
End Function

Public Function UnirCampos(ParamArray Campos() As Variant)
    
    UnirCampos = VBA.Replace(VBA.Join(Campos), " ", "")
    
End Function

Public Function ValidarDicionario(ByRef Dicionario As Dictionary) As Boolean
    
    If Not Dicionario Is Nothing Then ValidarDicionario = True
    
End Function

Public Function VerificarStringVazia(ByVal Valor As String) As Boolean
    
    VerificarStringVazia = IsEmpty(Valor) Or Len(Valor) = 0
    
End Function

Public Function VerificarPosicaoInicialArray(ByRef Campos As Variant) As Byte
    
    If LBound(Campos) = 0 Then VerificarPosicaoInicialArray = 1 Else VerificarPosicaoInicialArray = 0
    
End Function

Public Sub MsgAlerta_AusenciaDados(Optional Msg As String = "")
    
    If Msg = "" Then Msg = "Para utilizar essa função é necessário haver dados na planilha."
    Call MsgAlerta(Msg, "Registro sem dados informados")
    
End Sub

Public Function ConverterArrayListEmArray(ByRef arrDados As Variant) As Variant
    
    If TypeName(arrDados) = "ArrayList" Then
        
        ConverterArrayListEmArray = arrDados.toArray
        
    Else
        
        ConverterArrayListEmArray = arrDados
        
    End If
    
End Function

Public Function ExtrairArrayArquivoTXT(ByVal CaminhoArquivo As String) As Variant

Dim objFSO As New FileSystemObject
Dim arrConteudo() As String
Dim objStream As Object
Dim Conteudo As String

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(CaminhoArquivo) Then

        Set objStream = CreateObject("ADODB.Stream")

        With objStream

            .Type = 2
            .Charset = "Windows-1252"
            .Open
            .LoadFromFile CaminhoArquivo
            Conteudo = .ReadText(-1)

            .Close

        End With

        Set objStream = Nothing

        If Len(Conteudo) > 0 Then
            
            Conteudo = VBA.Replace(Conteudo, vbCrLf, vbLf)
            
            'Remover um vbLf final, se houver, para evitar um elemento vazio no final do array
            If Right(Conteudo, Len(vbLf)) = vbLf Then
                Conteudo = Left(Conteudo, Len(Conteudo) - Len(vbLf))
            End If
            
            'Remover um vbLf inicial, se houver (caso o arquivo comece com uma linha em branco)
            If Left(Conteudo, Len(vbLf)) = vbLf Then
                Conteudo = Mid(Conteudo, Len(vbLf) + 1)
            End If
            
            arrConteudo = VBA.Split(Conteudo, vbLf)
        End If

        ExtrairArrayArquivoTXT = arrConteudo

    Else

        ExtrairArrayArquivoTXT = Array()

    End If

    Set objFSO = Nothing

End Function

Public Function MesclarDicionarios(ByRef dicAtualizacoes As Dictionary, ByRef Plan As Worksheet) As Dictionary

Dim dicDadosOrigem As Object
Dim Chave As Variant

    Set dicDadosOrigem = Util.CriarDicionarioRegistro(Plan)
    
    For Each Chave In dicDadosOrigem.Keys()
        
        If Not dicAtualizacoes.Exists(Chave) Then
            dicAtualizacoes(Chave) = dicDadosOrigem(Chave)
        End If
    
    Next
    
    Set MesclarDicionarios = dicAtualizacoes

End Function

Public Function ListarTitulosPlanilha(ByVal Plan As Worksheet, ByVal LinhaTitulo As Long) As ArrayList

Dim arrTitulos As New ArrayList
Dim Titulos As Variant
Dim UltCol As Long
Dim i As Integer
    
    On Error GoTo Tratar:
    UltCol = Plan.Cells(LinhaTitulo, Plan.Columns.Count).END(xlToLeft).Column
    
    Titulos = Plan.Cells(LinhaTitulo, 1).Resize(, UltCol)
    Titulos = Application.index(Titulos, 0, 0)
    
    For i = LBound(Titulos) To UBound(Titulos)
        arrTitulos.Add Titulos(i)
    Next i
    
    Set ListarTitulosPlanilha = arrTitulos
    
Exit Function
Tratar:
    
    Set ListarTitulosPlanilha = Nothing
    
End Function

Public Function PastaTrabalhoSalva() As Boolean
    
Dim Msg As String
    
    If ThisWorkbook.Path = "" Then
        
        Msg = "Esta pasta de trabalho nunca foi salva." & vbCrLf & _
        Msg = Msg & "Por favor, salve-a manualmente com um nome e local antes de continuar."
        
        Call MsgAviso(Msg, "Ação Necessária")
                
    Else
        
        If Not ThisWorkbook.Saved Then
            
            On Error Resume Next
            
            ThisWorkbook.Save
            If Err.Number <> 0 Then
                MsgBox "Não foi possível salvar automaticamente a pasta de trabalho no local: " & ThisWorkbook.FullName & vbCrLf & _
                       "Verifique permissões ou se o arquivo está em uso." & vbCrLf & vbCrLf & _
                       "Erro VBA: " & Err.Description, vbCritical, "Falha ao Salvar Automaticamente"
                Err.Clear
                
            Else
                
                PastaTrabalhoSalva = True
                
            End If
            
            On Error GoTo 0
            
        Else
            
            PastaTrabalhoSalva = True
            
        End If
        
    End If
    
End Function
