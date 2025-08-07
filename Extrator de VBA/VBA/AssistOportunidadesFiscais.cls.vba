Attribute VB_Name = "AssistOportunidadesFiscais"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Public ARQUIVO As String
Public SPEDOriginal As Boolean
Public CamposFiscais As Variant
Public CamposContribuicoes As Variant
Public dicTitulosIPI As New Dictionary
Public dicTitulosICMS As New Dictionary
Public dicRelatorioIPI As New Dictionary
Public dicRelatorioICMS As New Dictionary
Public dicTitulosContribuicoes As New Dictionary
Public dicRelatorioContribuicoes As New Dictionary

Public Function GerarRelatorioOportunidades(ByVal SPEDOriginal As Boolean, ByVal Imposto As String)

Dim Caminho As String
Dim arrFiscal As New ArrayList
Dim arrContribuicoes As New ArrayList
    
    Me.SPEDOriginal = SPEDOriginal
    Caminho = Util.SelecionarPasta("Selecione a pasta com os SPEDs a serem processados")
    
    If Caminho = "" Then Exit Function
    
    Inicio = Now()
    
    Call fnSPED.ListarSPEDs(Caminho, arrFiscal, arrContribuicoes)
    
    If arrFiscal.Count > 0 Then
        
        If Imposto = "ICMS" Then
            
            Set dicRelatorioICMS = Util.CriarDicionarioRegistro(assOportunidadesICMS, "ARQUIVO")
            Set dicTitulosICMS = Util.MapearTitulos(assOportunidadesICMS, 3)
            Call ProcessarSPEDs(arrFiscal, Imposto)
            
            Call Util.LimparDados(assOportunidadesICMS, 4, False)
            Call Util.ExportarDadosDicionario(assOportunidadesICMS, dicRelatorioICMS)
            
        ElseIf Imposto = "IPI" Then
            
            Set dicRelatorioIPI = Util.CriarDicionarioRegistro(assOportunidadesIPI, "ARQUIVO")
            Set dicTitulosIPI = Util.MapearTitulos(assOportunidadesIPI, 3)
            Call ProcessarSPEDs(arrFiscal, Imposto)
            
            Call Util.LimparDados(assOportunidadesIPI, 4, False)
            Call Util.ExportarDadosDicionario(assOportunidadesIPI, dicRelatorioIPI)
            
        End If
        
    End If
    
    If arrContribuicoes.Count > 0 And Imposto = "PISCOFINS" Then
        
        Set dicRelatorioContribuicoes = Util.CriarDicionarioRegistro(assOportunidadesPIS_COFINS, "ARQUIVO")
        Set dicTitulosContribuicoes = Util.MapearTitulos(assOportunidadesPIS_COFINS, 3)
        Call ProcessarSPEDs(arrContribuicoes, Imposto)
        
        Call Util.LimparDados(assOportunidadesPIS_COFINS, 4, False)
        Call Util.ExportarDadosDicionario(assOportunidadesPIS_COFINS, dicRelatorioContribuicoes)
        
    End If
    
    Call dicRelatorioICMS.RemoveAll
    Call dicRelatorioContribuicoes.RemoveAll
    
End Function

Private Function ProcessarSPEDs(ByRef arrArqs As ArrayList, ByVal Imposto As String)
    
Dim Arq As Variant
    
    For Each Arq In arrArqs
        
        Call ProcessarSPED(Arq, Imposto)
        
    Next Arq
    
End Function

Private Function ProcessarSPED(ByVal Arq As String, ByVal Imposto As String)

Dim fso As New FileSystemObject
Dim Stream As TextStream
Dim Registro As String
Dim SairArquivo As Boolean
    
    'Verifica se o arquivo existe
    If Not fso.FileExists(Arq) Then Exit Function
    
    'Abre o arquivo para leitura
    Set Stream = fso.OpenTextFile(Arq, ForReading)
    Do While Not Stream.AtEndOfStream
        
        Registro = Stream.ReadLine
        Select Case Imposto
            
            Case "IPI"
                SairArquivo = ProcessarRegistroIPI(Registro)
                
            Case "ICMS"
                SairArquivo = ProcessarRegistroICMS(Registro)
                            
            Case "PISCOFINS"
                SairArquivo = ProcessarRegistroSPEDContribuicoes(Registro)
                
        End Select
        
        If SairArquivo Then Exit Function
        
    Loop
    
End Function

Private Function ProcessarRegistroICMS(ByVal Registro As String) As Boolean

Dim nReg As String, CNPJ$, DT_INI$, Chave$
Dim Campos As Variant
    
    If IsEmpty(Registro) Then Exit Function
    
    Campos = VBA.Split(Registro, "|")
    nReg = Campos(1)
    
    Select Case True
        
        Case nReg = "0000"
            CNPJ = Campos(7)
            DT_INI = Campos(4)
            Chave = MontarChaveArquivo(CNPJ, DT_INI)
            Me.ARQUIVO = Chave
            AtribuirValor "ARQUIVO", ARQUIVO, Chave, "ICMS"
            
        Case nReg = "E110"
            Call ExtrairSaldoICMS(Campos)
            
        Case nReg > "E110"
            Call CalcularDiferenca("ICMS")
            Call GerarRecomendacao("ICMS")
            ProcessarRegistroICMS = True
            
    End Select
    
End Function

Private Function ProcessarRegistroIPI(ByVal Registro As String) As Boolean

Dim nReg As String, CNPJ$, DT_INI$, Chave$
Dim Campos As Variant
    
    If IsEmpty(Registro) Then Exit Function
    
    Campos = VBA.Split(Registro, "|")
    nReg = Campos(1)
    
    Select Case True
        
        Case nReg = "0000"
            CNPJ = Campos(7)
            DT_INI = Campos(4)
            Chave = MontarChaveArquivo(CNPJ, DT_INI)
            Me.ARQUIVO = Chave
            AtribuirValor "ARQUIVO", ARQUIVO, Chave, "IPI"
            
        Case nReg = "E520"
            Call ExtrairSaldoIPI(Campos)
            
        Case nReg > "E520"
            Call CalcularDiferenca("IPI")
            Call GerarRecomendacao("IPI")
            ProcessarRegistroIPI = True
            
    End Select
    
End Function

Private Function ProcessarRegistroSPEDContribuicoes(ByVal Registro As String) As Boolean

Dim nReg As String, CNPJ$, DT_INI$, Chave$
Dim Campos As Variant
    
    If IsEmpty(Registro) Then Exit Function
    
    Campos = VBA.Split(Registro, "|")
    nReg = Campos(1)
    
    Select Case True
        
        Case nReg = "0000"
            CNPJ = Campos(9)
            DT_INI = Campos(6)
            Chave = MontarChaveArquivo(CNPJ, DT_INI)
            Me.ARQUIVO = Chave
            AtribuirValor "ARQUIVO", ARQUIVO, Chave, "PIS"
            
        Case nReg = "M100"
            Call ExtrairSaldoCredorPIS(Campos)
            
        Case nReg = "M200"
            Call ExtrairSaldoDevedorPIS(Campos)
            
        Case nReg = "M500"
            Call ExtrairSaldoCredorCOFINS(Campos)
            
        Case nReg = "M600"
            Call ExtrairSaldoDevedorCOFINS(Campos)
            
        Case nReg > "M600"
            Call CalcularDiferenca("PIS")
            Call CalcularDiferenca("COFINS")
            Call GerarRecomendacao("PIS")
            Call GerarRecomendacao("COFINS")
            ProcessarRegistroSPEDContribuicoes = True
            
    End Select
    
End Function

Private Function MontarChaveArquivo(ByVal CNPJ As String, ByVal DT_INI As String)

Dim Periodo As String
    
    Periodo = Util.ExtrairPeriodo(DT_INI)
    MontarChaveArquivo = Periodo & "-" & CNPJ
    
End Function

Private Function ExtrairSaldoICMS(ByRef Campos As Variant)

Dim vDevedorICMS As Double, vCredorICMS As Double, vSaldoICMS As Double
    
    vDevedorICMS = fnExcel.ConverterValores(Campos(13), True, 2)
    vCredorICMS = fnExcel.ConverterValores(Campos(14), True, 2)
    
    vSaldoICMS = ExtrairSaldoCredorDevedor(vCredorICMS, vDevedorICMS, "ICMS", "Fiscal")
    
    If SPEDOriginal Then AtribuirValor "ICMS_ORIGINAL", ARQUIVO, vSaldoICMS, "ICMS" Else AtribuirValor "ICMS_CORRIGIDO", ARQUIVO, vSaldoICMS, "ICMS"
    
End Function

Private Function ExtrairSaldoIPI(ByRef Campos As Variant)

Dim vDevedorIPI As Double, vCredorIPI As Double, vSaldoIPI As Double
    
    vCredorIPI = fnExcel.ConverterValores(Campos(7), True, 2)
    vDevedorIPI = fnExcel.ConverterValores(Campos(8), True, 2)
    
    vSaldoIPI = ExtrairSaldoCredorDevedor(vCredorIPI, vDevedorIPI, "IPI", "Fiscal")
    
    If SPEDOriginal Then AtribuirValor "IPI_ORIGINAL", ARQUIVO, vSaldoIPI, "IPI" Else AtribuirValor "IPI_CORRIGIDO", ARQUIVO, vSaldoIPI, "IPI"
    
End Function

Private Function ExtrairSaldoCredorPIS(ByRef Campos As Variant)

Dim SaldoCredorPIS As Double
    
    If Not VerificarUtilizacaoTotalCredito(Campos) Then
        
        SaldoCredorPIS = fnExcel.ConverterValores(Campos(15), True, 2)
        If SaldoCredorPIS > 0 Then AtribuirSaldoImposto "PIS", -SaldoCredorPIS
        
    End If
    
End Function

Private Function ExtrairSaldoCredorCOFINS(ByRef Campos As Variant)

Dim SaldoCredorCOFINS As Double
    
    If Not VerificarUtilizacaoTotalCredito(Campos) Then
                
        SaldoCredorCOFINS = fnExcel.ConverterValores(Campos(15), True, 2)
        If SaldoCredorCOFINS > 0 Then AtribuirSaldoImposto "COFINS", -SaldoCredorCOFINS
    
    End If
    
End Function

Private Function ExtrairSaldoDevedorPIS(ByRef Campos As Variant)

Dim SaldoDevedorPIS As Double
    
    SaldoDevedorPIS = fnExcel.ConverterValores(Campos(13), True, 2)
    If SaldoDevedorPIS > 0 Then AtribuirSaldoImposto "PIS", SaldoDevedorPIS
        
End Function

Private Function ExtrairSaldoDevedorCOFINS(ByRef Campos As Variant)

Dim SaldoDevedorCOFINS As Double
    
    SaldoDevedorCOFINS = fnExcel.ConverterValores(Campos(13), True, 2)
    If SaldoDevedorCOFINS > 0 Then AtribuirSaldoImposto "COFINS", SaldoDevedorCOFINS
    
End Function

Private Function VerificarUtilizacaoTotalCredito(ByRef Campos As Variant) As Boolean
    
    Select Case Campos(13)
        
        Case "0"
            VerificarUtilizacaoTotalCredito = True
            
    End Select
    
End Function

Private Sub AtribuirSaldoImposto(ByVal Imposto As String, ByVal Saldo As Double)
    
    Select Case SPEDOriginal
        
        Case True
            AtribuirValor Imposto & "_ORIGINAL", ARQUIVO, Saldo, Imposto
            
        Case False
            AtribuirValor Imposto & "_CORRIGIDO", ARQUIVO, Saldo, Imposto
            
    End Select
    
End Sub

Public Function AtribuirValor(ByVal Titulo As String, ByVal Chave As String, ByVal Valor As Variant, ByVal Imposto As String)
    
    Select Case Imposto
        
        Case "IPI"
            Call AtribuirValorIPI(Titulo, Chave, Valor)
            
        Case "ICMS"
            Call AtribuirValorICMS(Titulo, Chave, Valor)
        
        Case "PIS", "COFINS"
            Call AtribuirValorContribuicoes(Titulo, Chave, Valor)
            
    End Select
    
End Function

Private Function AtribuirValorICMS(ByVal Titulo As String, ByVal Chave As String, ByVal Valor As Variant)

Dim i As Byte

    If dicRelatorioICMS.Exists(Chave) Then
        
        CamposFiscais = dicRelatorioICMS(Chave)
        
        If LBound(CamposFiscais) = 0 Then i = 1 Else i = 0
        CamposFiscais(dicTitulosICMS(Titulo) - i) = fnExcel.FormatarTipoDado(Titulo, Valor)
        
    Else
        
        Call RedimensionarArrayFiscal(dicTitulosICMS.Count)
        CamposFiscais(dicTitulosICMS(Titulo)) = fnExcel.FormatarTipoDado(Titulo, Valor)
        
    End If
        
    dicRelatorioICMS(Chave) = CamposFiscais

End Function

Private Function AtribuirValorIPI(ByVal Titulo As String, ByVal Chave As String, ByVal Valor As Variant)

Dim i As Byte

    If dicRelatorioIPI.Exists(Chave) Then
        
        CamposFiscais = dicRelatorioIPI(Chave)
        
        If LBound(CamposFiscais) = 0 Then i = 1 Else i = 0
        CamposFiscais(dicTitulosIPI(Titulo) - i) = fnExcel.FormatarTipoDado(Titulo, Valor)
        
    Else
        
        Call RedimensionarArrayFiscal(dicTitulosIPI.Count)
        CamposFiscais(dicTitulosIPI(Titulo)) = fnExcel.FormatarTipoDado(Titulo, Valor)
        
    End If
        
    dicRelatorioIPI(Chave) = CamposFiscais

End Function

Private Function AtribuirValorContribuicoes(ByVal Titulo As String, ByVal Chave As String, ByVal Valor As Variant)

Dim i As Byte
    
    If dicRelatorioContribuicoes.Exists(Chave) Then
        
        CamposContribuicoes = dicRelatorioContribuicoes(Chave)
        
        If LBound(CamposContribuicoes) = 0 Then i = 1 Else i = 0
        CamposContribuicoes(dicTitulosContribuicoes(Titulo)) = fnExcel.FormatarTipoDado(Titulo, Valor)
        
    Else
        
        Call RedimensionarArrayContribuicoes(dicTitulosContribuicoes.Count)
        CamposContribuicoes(dicTitulosContribuicoes(Titulo)) = fnExcel.FormatarTipoDado(Titulo, Valor)
        
    End If
    
    dicRelatorioContribuicoes(Chave) = CamposContribuicoes
    
End Function

Public Function RedimensionarArrayFiscal(ByVal NumCampos As Long)
    
    ReDim CamposFiscais(1 To NumCampos) As Variant
    
End Function

Public Function RedimensionarArrayContribuicoes(ByVal NumCampos As Long)
    
    ReDim CamposContribuicoes(1 To NumCampos) As Variant
    
End Function

Private Function ExtrairSaldoCredorDevedor(ByVal SaldoCredor As Double, _
    ByVal SaldoDevedor As Double, ByVal Imposto As String, ByVal TipoSPED As String) As Double
    
    Select Case True
        
        Case SaldoCredor > 0 And SaldoDevedor = 0
            ExtrairSaldoCredorDevedor = -SaldoCredor
            
        Case SaldoDevedor > 0 And SaldoCredor = 0
            ExtrairSaldoCredorDevedor = SaldoDevedor
            
        Case SaldoDevedor = 0 And SaldoCredor = 0
            ExtrairSaldoCredorDevedor = 0
            
        Case Else
            AtribuirValor "DIFERENCA_" & Imposto, ARQUIVO, -999999.99, TipoSPED
            ExtrairSaldoCredorDevedor = -999999.99
            
    End Select
    
End Function

Private Function GerarRecomendacao(ByVal Imposto As String)

Dim DIFERENCA As Variant
Dim RECOMENDACAO As String
    
    DIFERENCA = ExtrairDiferenca(Imposto)
    
    Select Case DIFERENCA
        
        Case Is = -999999.99
            RECOMENDACAO = "O período com saldo devedor e credor ao mesmo tempo, provável erro de apuração."
            
        Case Is = 0
            RECOMENDACAO = "Saldos ORIGINAL e CORRIGIDO idênticos, provável erro na importação dos dados."
        
        Case ""
            RECOMENDACAO = "Importe as apurações com valores corrigidos para gerar a análise."
            
        Case Is > 0
            RECOMENDACAO = "Oferecer Retificação do SPED com denúncia espontânea para eliminar a multa."
            
        Case Is < 0
            RECOMENDACAO = "Oferecer recuperação tributária para reaver o imposto pago a maior."
        
    End Select
    
    AtribuirValor "RECOMENDACAO", ARQUIVO, RECOMENDACAO, Imposto
    
End Function

Private Function ExtrairDiferenca(ByVal Imposto As String) As Variant
    
    Select Case Imposto
        
        Case "IPI"
            ExtrairDiferenca = CamposFiscais(dicTitulosIPI("DIFERENCA_" & Imposto))
            
        Case "ICMS"
            ExtrairDiferenca = CamposFiscais(dicTitulosICMS("DIFERENCA_" & Imposto))
            
        Case "PIS", "COFINS"
            ExtrairDiferenca = CamposContribuicoes(dicTitulosContribuicoes("DIFERENCA_" & Imposto))
            
    End Select
    
End Function

Private Function CalcularDiferenca(ByVal Imposto As String) As Double

Dim DIFERENCA As Variant
    
    Select Case Imposto
        
        Case "IPI"
            DIFERENCA = CalcularDiferencaIPI()
            
        Case "ICMS"
            DIFERENCA = CalcularDiferencaICMS()
            
        Case "PIS", "COFINS"
            DIFERENCA = CalcularDiferencaContribuicoes(Imposto)
            
    End Select
    
    AtribuirValor "DIFERENCA_" & Imposto, ARQUIVO, DIFERENCA, Imposto
    
End Function

Private Function CalcularDiferencaContribuicoes(ByVal Imposto As String) As Variant

Dim VL_ORIGINAL As Double, VL_CORRIGIDO#
    
    VL_ORIGINAL = fnExcel.ConverterValores(CamposContribuicoes(dicTitulosContribuicoes(Imposto & "_ORIGINAL")), True, 2)
    VL_CORRIGIDO = fnExcel.ConverterValores(CamposContribuicoes(dicTitulosContribuicoes(Imposto & "_CORRIGIDO")), True, 2)
    
    CalcularDiferencaContribuicoes = ValidarCalculoDiferenca(VL_ORIGINAL, VL_CORRIGIDO)
    
End Function

Private Function CalcularDiferencaICMS() As Variant

Dim VL_ORIGINAL As Double, VL_CORRIGIDO#
    
    VL_ORIGINAL = fnExcel.ConverterValores(CamposFiscais(dicTitulosICMS("ICMS_ORIGINAL")), True, 2)
    VL_CORRIGIDO = fnExcel.ConverterValores(CamposFiscais(dicTitulosICMS("ICMS_CORRIGIDO")), True, 2)
    
    CalcularDiferencaICMS = ValidarCalculoDiferenca(VL_ORIGINAL, VL_CORRIGIDO)
    
End Function

Private Function CalcularDiferencaIPI() As Variant

Dim VL_ORIGINAL As Double, VL_CORRIGIDO#
    
    VL_ORIGINAL = fnExcel.ConverterValores(CamposFiscais(dicTitulosIPI("IPI_ORIGINAL")), True, 2)
    VL_CORRIGIDO = fnExcel.ConverterValores(CamposFiscais(dicTitulosIPI("IPI_CORRIGIDO")), True, 2)
    
    CalcularDiferencaIPI = ValidarCalculoDiferenca(VL_ORIGINAL, VL_CORRIGIDO)
    
End Function

Private Function ValidarCalculoDiferenca(ByVal VL_ORIGINAL As Double, ByVal VL_CORRIGIDO As Double) As Variant
    
    Select Case True
        
        Case VL_ORIGINAL = -999999.99 Or VL_CORRIGIDO = -999999.99
            ValidarCalculoDiferenca = -999999.99
            
        Case VL_ORIGINAL <> 0 And VL_CORRIGIDO <> 0
            ValidarCalculoDiferenca = VBA.Round(VL_CORRIGIDO - VL_ORIGINAL, 2)
            
        Case Else
            ValidarCalculoDiferenca = ""
            
    End Select
    
End Function
