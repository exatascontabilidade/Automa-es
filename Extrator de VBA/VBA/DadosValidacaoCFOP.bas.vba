Attribute VB_Name = "DadosValidacaoCFOP"
Option Explicit

Private ValidacoesCFOP As New clsRegrasFiscaisCFOP
Public Campos As Variant
Public TipoRelatorio As String
Public TabelaCFOP As New Dictionary
Public CamposCFOP As CamposValidacaoCFOP
Public dicTitulosRelatorio As New Dictionary
Public CompraRevendaSemST As Boolean
Public CompraRevendaComST As Boolean
Public CompraUsoConsumoSemST As Boolean
Public CompraUsoConsumoComST As Boolean
Public CompraImobilizadoSemST As Boolean
Public CompraImobilizadoComST As Boolean
Public CompraPrestacaoServico As Boolean
Public CompraIndustrializacao As Boolean
Public Industrializacao As Boolean
Public Comercializacao As Boolean
Public CompraCombustivelConsumo As Boolean
Public CompraCombustivelRevenda As Boolean
Public CompraIndustrializacaoSemST As Boolean
Public CompraIndustrializacaoComST As Boolean
Public EntradaInterna As Boolean
Public EntradaInterestadual As Boolean
Public SaidaInterestadual As Boolean
Public TamanhoCFOP As Boolean
Public ExisteCFOP As Boolean
Public SaidaInterna As Boolean
Public Importacao As Boolean
Public Exportacao As Boolean
Public VendaSemST As Boolean
Public VendaComST As Boolean
Public VendaCombustivel As Boolean
Public EntradaSemST As Boolean
Public EntradaComST As Boolean
Public CST_IPITributado As Boolean
Public CST_IPIAliqZero As Boolean
Public CST_IPIIsento As Boolean
Public CST_IPINaoTributado As Boolean
Public CST_IPIImune As Boolean
Public CST_IPISuspensao As Boolean
Public CST_IPIOutrasSaidas As Boolean
Public CST_IPIEntradaTributada As Boolean
Public CST_IPIEntradaZero As Boolean
Public CST_IPIEntradaIsenta As Boolean
Public CST_IPIEntradaNaoTributada As Boolean
Public CST_IPIEntradaImune As Boolean
Public CST_IPIEntradaSuspensa As Boolean
Public CST_IPIOutrasEntradas As Boolean
Public CST_ICMSTributado As Boolean
Public CST_ICMSReducao As Boolean
Public CST_ICMSComST As Boolean
Public CST_ICMSIsento As Boolean
Public CST_ICMSNaoTributado As Boolean
Public CST_ICMSSuspensao As Boolean
Public CST_ICMSDiferimento As Boolean
Public CST_ICMSOutras As Boolean
Public CST_PISAliqBasica As Boolean
Public CST_PISAliqDiferenciada As Boolean
Public CST_PISQuantidade As Boolean
Public CST_PISMonofasico As Boolean
Public CST_PISSubstituicao As Boolean
Public CST_PISAliqZero As Boolean
Public CST_PISIsento As Boolean
Public CST_PISSemIncidencia As Boolean
Public CST_PISSuspensao As Boolean
Public CST_PISOutrasSaidas As Boolean
Public CST_PISComCredito As Boolean
Public CST_PISPresumido As Boolean
Public CST_PISSemCredito As Boolean
Public CST_PISAquisicaoIsenta As Boolean
Public CST_PISAquisicaoSuspensao As Boolean
Public CST_PISAquisicaoZero As Boolean
Public CST_PISAquisicaoSemIncidencia As Boolean
Public CST_PISAquisicaoSubstituicao As Boolean
Public CST_PISOutrasEntradas As Boolean
Public CST_PISOutrasOperacoes As Boolean
Public CST_COFINSAliqBasica As Boolean
Public CST_COFINSAliqDiferenciada As Boolean
Public CST_COFINSQuantidade As Boolean
Public CST_COFINSMonofasico As Boolean
Public CST_COFINSSubstituicao As Boolean
Public CST_COFINSAliqZero As Boolean
Public CST_COFINSIsento As Boolean
Public CST_COFINSSemIncidencia As Boolean
Public CST_COFINSSuspensao As Boolean
Public CST_COFINSOutrasSaidas As Boolean
Public CST_COFINSComCredito As Boolean
Public CST_COFINSPresumido As Boolean
Public CST_COFINSSemCredito As Boolean
Public CST_COFINSAquisicaoIsenta As Boolean
Public CST_COFINSAquisicaoSuspensao As Boolean
Public CST_COFINSAquisicaoZero As Boolean
Public CST_COFINSAquisicaoSemIncidencia As Boolean
Public CST_COFINSAquisicaoSubstituicao As Boolean
Public CST_COFINSOutrasEntradas As Boolean
Public CST_COFINSOutrasOperacoes As Boolean

Public Type CamposValidacaoCFOP
    
    ARQUIVO As String
    COD_CFOP As String
    CFOP_SPED As String
    CFOP_NF As String
    DT_DOC As String
    DT_REF As String
    IND_OPER As String
    DT_ENT_SAI As String
    DESCRICAO As String
    UF_CONTRIB As String
    UF_PART As String
    CST_IPI As String
    CST_ICMS As String
    CST_PIS As String
    CST_COFINS As String
    CST_ICMS_NF As String
    CST_ICMS_SPED As String
    VL_IPI As Double
    VL_ICMS As Double
    VL_PIS As Double
    VL_COFINS As Double
    VIGENCIA_FINAL As String
    VIGENCIA_INICIAL As String
    INCONSISTENCIA As String
    SUGESTAO As String
    
End Type

Public Function CarregarTitulosRelatorio(ByRef Plan As Worksheet)
    
    Set dicTitulosRelatorio = Util.MapearTitulos(Plan, 3)
    
End Function

Public Function CarregarCamposCFOP(ByVal CamposRel As Variant, ByVal Imposto As String)
    
    Campos = CamposRel
    TipoRelatorio = Imposto
    If dicTitulosRelatorio.Count = 0 Then Call CarregarTitulosRelatorio(ActiveSheet)
    
    Call DirecionarCarregamentoCamposCFOP
    
End Function

Private Function DirecionarCarregamentoCamposCFOP()
    
    If TipoRelatorio Like "*DIVERGENCIAS*" Then _
        Call CarregarDadosDivergenciasCFOP Else _
            Call CarregarDadosApuracaoCFOP
    
End Function

Public Function CarregarDadosDivergenciasCFOP()

Dim i As Long
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    CamposCFOP.ARQUIVO = Util.RemoverAspaSimples(Campos(dicTitulosRelatorio("ARQUIVO") - i))
    CamposCFOP.CFOP_NF = fnExcel.ConverterValores(Campos(dicTitulosRelatorio("CFOP_NF") - i), True, 0)
    CamposCFOP.CFOP_SPED = Util.RemoverAspaSimples(Campos(dicTitulosRelatorio("CFOP_SPED") - i))
    CamposCFOP.COD_CFOP = CamposCFOP.CFOP_SPED
    CamposCFOP.CST_ICMS_NF = Util.RemoverAspaSimples(Campos(dicTitulosRelatorio("CST_ICMS_NF") - i))
    CamposCFOP.CST_ICMS_SPED = Util.RemoverAspaSimples(Campos(dicTitulosRelatorio("CST_ICMS_SPED") - i))
    'CamposCFOP.UF_PART = fnExcel.FormatarData(Campos(dicTitulosRelatorio("UF_PART") - i))
    'CamposCFOP.UF_CONTRIB = fnExcel.FormatarData(Campos(dicTitulosRelatorio("UF_CONTRIB") - i))
    
    Call DirecionarCarregamentoVerificacoesCFOP
    
End Function

Public Function CarregarDadosApuracaoCFOP()

Dim i As Long
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    CamposCFOP.ARQUIVO = Util.RemoverAspaSimples(Campos(dicTitulosRelatorio("ARQUIVO") - i))
    CamposCFOP.COD_CFOP = Util.ApenasNumeros(Campos(dicTitulosRelatorio("CFOP") - i))
    CamposCFOP.IND_OPER = Util.RemoverAspaSimples(Campos(dicTitulosRelatorio("IND_OPER") - i))
    CamposCFOP.DT_DOC = fnExcel.FormatarData(Campos(dicTitulosRelatorio("DT_DOC") - i))
    CamposCFOP.DT_ENT_SAI = fnExcel.FormatarData(Campos(dicTitulosRelatorio("DT_ENT_SAI") - i))
    CamposCFOP.UF_PART = Campos(dicTitulosRelatorio("UF_PART") - i)
    CamposCFOP.UF_CONTRIB = Campos(dicTitulosRelatorio("UF_CONTRIB") - i)
    
    Call DirecionarCarregamentoCamposCFOPImpostos
    
End Function

Private Function DirecionarCarregamentoCamposCFOPImpostos()
    
    Select Case True
        
        Case TipoRelatorio Like "*IPI*"
            Call CarregarDadosApuracaoCFOP_IPI
            
        Case TipoRelatorio Like "*ICMS*"
            Call CarregarDadosApuracaoCFOP_ICMS
            
        Case TipoRelatorio Like "*PISCOFINS*"
            Call CarregarDadosApuracaoCFOP_PISCOFINS
            
    End Select
    
    Call CarregarVerificacoesCFOP
    
End Function

Public Function CarregarDadosApuracaoCFOP_IPI()

Dim i As Long
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    CamposCFOP.CST_IPI = Util.ApenasNumeros(Campos(dicTitulosRelatorio("CST_IPI") - i))
    CamposCFOP.VL_IPI = fnExcel.ConverterValores(Campos(dicTitulosRelatorio("VL_IPI") - i))
    
End Function

Public Function CarregarDadosApuracaoCFOP_ICMS()

Dim i As Long
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    CamposCFOP.CST_ICMS = Util.ApenasNumeros(Campos(dicTitulosRelatorio("CST_ICMS") - i))
    CamposCFOP.VL_ICMS = fnExcel.ConverterValores(Campos(dicTitulosRelatorio("VL_ICMS") - i))
    
End Function

Public Function CarregarDadosApuracaoCFOP_PISCOFINS()

Dim i As Long
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    CamposCFOP.CST_PIS = Util.ApenasNumeros(Campos(dicTitulosRelatorio("CST_PIS") - i))
    CamposCFOP.CST_COFINS = Util.ApenasNumeros(Campos(dicTitulosRelatorio("CST_COFINS") - i))
    CamposCFOP.VL_PIS = fnExcel.ConverterValores(Campos(dicTitulosRelatorio("VL_PIS") - i))
    CamposCFOP.VL_COFINS = fnExcel.ConverterValores(Campos(dicTitulosRelatorio("VL_COFINS") - i))
    
End Function

Private Function DirecionarCarregamentoVerificacoesCFOP()
    
    Select Case True
        
        Case TipoRelatorio Like "*IPI*"
            Call CarregarVerificacoesCFOP_IPI
            
        Case TipoRelatorio Like "*ICMS*"
            Call CarregarVerificacoesCFOP_ICMS
            
        Case TipoRelatorio Like "*PISCOFINS*"
            Call CarregarVerificacoesCFOP_PIS
            Call CarregarVerificacoesCFOP_COFINS
        
        Case TipoRelatorio Like "*DIVERGENCIAS*"
            Call CarregarVerificacoesCFOP_Divergencias
            
    End Select
        
End Function

Public Function CarregarDadosTabelaCFOP(ByVal Campos As Variant)
    
    CamposCFOP.COD_CFOP = Campos(0)
    CamposCFOP.DESCRICAO = Campos(1)
    CamposCFOP.VIGENCIA_INICIAL = fnExcel.FormatarData(Campos(2))
    CamposCFOP.VIGENCIA_FINAL = fnExcel.FormatarData(Campos(3))
    
End Function

Private Function CarregarVerificacoesCFOP()
    
    EntradaInterna = ValidacoesCFOP.ValidarCFOPEntradaInterna(CamposCFOP.COD_CFOP)
    EntradaInterestadual = ValidacoesCFOP.ValidarCFOPEntradaInterestadual(CamposCFOP.COD_CFOP)
    Importacao = ValidacoesCFOP.ValidarCFOPImportacao(CamposCFOP.COD_CFOP)
    
    CompraUsoConsumoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CamposCFOP.COD_CFOP)
    CompraUsoConsumoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CamposCFOP.COD_CFOP)
    
    CompraImobilizadoSemST = ValidacoesCFOP.ValidarCFOPCompraAtivoImobilizadoSemST(CamposCFOP.COD_CFOP)
    CompraImobilizadoComST = ValidacoesCFOP.ValidarCFOPCompraAtivoImobilizadoComST(CamposCFOP.COD_CFOP)
    
    CompraRevendaSemST = ValidacoesCFOP.ValidarCFOPCompraRevendaSemST(CamposCFOP.COD_CFOP)
    CompraRevendaComST = ValidacoesCFOP.ValidarCFOPCompraRevendaComST(CamposCFOP.COD_CFOP)
    CompraIndustrializacao = ValidacoesCFOP.ValidarCFOPCompraIndustrializacao(CamposCFOP.COD_CFOP)
    CompraIndustrializacaoSemST = ValidacoesCFOP.ValidarCFOPCompraIndustrializacaoSemST(CamposCFOP.COD_CFOP)
    CompraIndustrializacaoComST = ValidacoesCFOP.ValidarCFOPCompraIndustrializacaoComST(CamposCFOP.COD_CFOP)
    CompraCombustivelRevenda = ValidacoesCFOP.ValidarCFOPCompraCombustiveisRevenda(CamposCFOP.COD_CFOP)
    CompraCombustivelConsumo = ValidacoesCFOP.ValidarCFOPCompraCombustiveisConsumo(CamposCFOP.COD_CFOP)
    
    SaidaInterna = ValidacoesCFOP.ValidarCFOPSaidaInterna(CamposCFOP.COD_CFOP)
    SaidaInterestadual = ValidacoesCFOP.ValidarCFOPSaidaInterestadual(CamposCFOP.COD_CFOP)
    Exportacao = ValidacoesCFOP.ValidarCFOPExportacao(CamposCFOP.COD_CFOP)
    VendaComST = ValidacoesCFOP.ValidarCFOPVendaComST(CamposCFOP.COD_CFOP)
    VendaSemST = ValidacoesCFOP.ValidarCFOPVendaSemST(CamposCFOP.COD_CFOP)
    VendaCombustivel = ValidacoesCFOP.ValidarCFOPVendaCombustiveis(CamposCFOP.COD_CFOP)
    
    EntradaSemST = CamposCFOP.COD_CFOP < 4000 And (Not CamposCFOP.COD_CFOP Like "#4##" And Not CamposCFOP.COD_CFOP Like "#55#" And Not CamposCFOP.COD_CFOP Like "#65#" And Not CamposCFOP.COD_CFOP Like "#9##")
    EntradaComST = CamposCFOP.COD_CFOP < 4000 And Not CamposCFOP.COD_CFOP Like "#9##" And (CamposCFOP.COD_CFOP Like "#4##" Or CamposCFOP.COD_CFOP Like "#65#")
       
    Call DirecionarCarregamentoVerificacoesCFOP
    
End Function

Private Function CarregarVerificacoesCFOP_IPI()
    
    'Saídas
    CST_IPITributado = CamposCFOP.CST_IPI Like "*50"
    CST_IPIAliqZero = CamposCFOP.CST_IPI Like "*51"
    CST_IPIIsento = CamposCFOP.CST_IPI Like "*52"
    CST_IPINaoTributado = CamposCFOP.CST_IPI Like "*53"
    CST_IPIImune = CamposCFOP.CST_IPI Like "*54"
    CST_IPISuspensao = CamposCFOP.CST_IPI Like "*55"
    CST_IPIOutrasSaidas = CamposCFOP.CST_IPI Like "*99"
    
    'Entradas
    CST_IPIEntradaTributada = CamposCFOP.CST_IPI Like "*00"
    CST_IPIEntradaZero = CamposCFOP.CST_IPI Like "*01"
    CST_IPIEntradaIsenta = CamposCFOP.CST_IPI Like "*02"
    CST_IPIEntradaNaoTributada = CamposCFOP.CST_IPI Like "*03"
    CST_IPIEntradaImune = CamposCFOP.CST_IPI Like "*04"
    CST_IPIEntradaSuspensa = CamposCFOP.CST_IPI Like "*05"
    CST_IPIOutrasEntradas = CamposCFOP.CST_IPI Like "*49"

End Function

Private Function CarregarVerificacoesCFOP_ICMS()
    
    CST_ICMSTributado = CamposCFOP.CST_ICMS Like "*00"
    CST_ICMSReducao = CamposCFOP.CST_ICMS Like "*20"
    CST_ICMSComST = CamposCFOP.CST_ICMS Like "*10" Or CamposCFOP.CST_ICMS Like "*30" Or CamposCFOP.CST_ICMS Like "*60" Or CamposCFOP.CST_ICMS Like "*61" Or CamposCFOP.CST_ICMS Like "*70"
    CST_ICMSIsento = CamposCFOP.CST_ICMS Like "*40"
    CST_ICMSNaoTributado = CamposCFOP.CST_ICMS Like "*41"
    CST_ICMSSuspensao = CamposCFOP.CST_ICMS Like "*50"
    CST_ICMSDiferimento = CamposCFOP.CST_ICMS Like "*51"
    CST_ICMSOutras = CamposCFOP.CST_ICMS Like "*50"
    
End Function

Private Function CarregarVerificacoesCFOP_PIS()
    
    'Operações de Saída
    CST_PISAliqBasica = CamposCFOP.CST_PIS Like "*1"
    CST_PISAliqDiferenciada = CamposCFOP.CST_PIS Like "*2"
    CST_PISQuantidade = CamposCFOP.CST_PIS Like "*3"
    CST_PISMonofasico = CamposCFOP.CST_PIS Like "*4"
    CST_PISSubstituicao = CamposCFOP.CST_PIS Like "*5"
    CST_PISAliqZero = CamposCFOP.CST_PIS Like "*6"
    CST_PISIsento = CamposCFOP.CST_PIS Like "*7"
    CST_PISSemIncidencia = CamposCFOP.CST_PIS Like "*8"
    CST_PISSuspensao = CamposCFOP.CST_PIS Like "*9"
    CST_PISOutrasSaidas = CamposCFOP.CST_PIS Like "49"
    
    'Operações de Entrada
    CST_PISComCredito = CamposCFOP.CST_PIS Like "5*"
    CST_PISPresumido = CamposCFOP.CST_PIS Like "6*"
    CST_PISSemCredito = CamposCFOP.CST_PIS Like "*70"
    CST_PISAquisicaoIsenta = CamposCFOP.CST_PIS Like "*71"
    CST_PISAquisicaoSuspensao = CamposCFOP.CST_PIS Like "*72"
    CST_PISAquisicaoZero = CamposCFOP.CST_PIS Like "*73"
    CST_PISAquisicaoSemIncidencia = CamposCFOP.CST_PIS Like "*74"
    CST_PISAquisicaoSubstituicao = CamposCFOP.CST_PIS Like "*75"
    CST_PISOutrasEntradas = CamposCFOP.CST_PIS Like "*98"
    
    CST_PISOutrasOperacoes = CamposCFOP.CST_PIS Like "*99"
    
End Function

Private Function CarregarVerificacoesCFOP_COFINS()
    
    'Operações de Saída
    CST_COFINSAliqBasica = CamposCFOP.CST_COFINS Like "*1"
    CST_COFINSAliqDiferenciada = CamposCFOP.CST_COFINS Like "*2"
    CST_COFINSQuantidade = CamposCFOP.CST_COFINS Like "*3"
    CST_COFINSMonofasico = CamposCFOP.CST_COFINS Like "*4"
    CST_COFINSSubstituicao = CamposCFOP.CST_COFINS Like "*5"
    CST_COFINSAliqZero = CamposCFOP.CST_COFINS Like "*6"
    CST_COFINSIsento = CamposCFOP.CST_COFINS Like "*7"
    CST_COFINSSemIncidencia = CamposCFOP.CST_COFINS Like "*8"
    CST_COFINSSuspensao = CamposCFOP.CST_COFINS Like "*9"
    CST_COFINSOutrasSaidas = CamposCFOP.CST_COFINS Like "49"
    
    'Operações de Entrada
    CST_COFINSComCredito = CamposCFOP.CST_COFINS Like "5*"
    CST_COFINSPresumido = CamposCFOP.CST_COFINS Like "6*"
    CST_COFINSSemCredito = CamposCFOP.CST_COFINS Like "*70"
    CST_COFINSAquisicaoIsenta = CamposCFOP.CST_COFINS Like "*71"
    CST_COFINSAquisicaoSuspensao = CamposCFOP.CST_COFINS Like "*72"
    CST_COFINSAquisicaoZero = CamposCFOP.CST_COFINS Like "*73"
    CST_COFINSAquisicaoSemIncidencia = CamposCFOP.CST_COFINS Like "*74"
    CST_COFINSAquisicaoSubstituicao = CamposCFOP.CST_COFINS Like "*75"
    CST_COFINSOutrasEntradas = CamposCFOP.CST_COFINS Like "*98"
    
    CST_COFINSOutrasOperacoes = CamposCFOP.CST_COFINS Like "*99"
    
End Function

Private Sub CarregarVerificacoesCFOP_Divergencias()
    
    EntradaInterna = ValidacoesCFOP.ValidarCFOPEntradaInterna(CamposCFOP.CFOP_SPED)
    EntradaInterestadual = ValidacoesCFOP.ValidarCFOPEntradaInterestadual(CamposCFOP.CFOP_SPED)
    Importacao = ValidacoesCFOP.ValidarCFOPImportacao(CamposCFOP.CFOP_SPED)
    
    CompraUsoConsumoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CamposCFOP.CFOP_SPED)
    CompraUsoConsumoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CamposCFOP.CFOP_SPED)
    
    CompraImobilizadoSemST = ValidacoesCFOP.ValidarCFOPCompraAtivoImobilizadoSemST(CamposCFOP.CFOP_SPED)
    CompraImobilizadoComST = ValidacoesCFOP.ValidarCFOPCompraAtivoImobilizadoComST(CamposCFOP.CFOP_SPED)
    
    CompraRevendaSemST = ValidacoesCFOP.ValidarCFOPCompraRevendaSemST(CamposCFOP.CFOP_SPED)
    CompraRevendaComST = ValidacoesCFOP.ValidarCFOPCompraRevendaComST(CamposCFOP.CFOP_SPED)
    CompraIndustrializacao = ValidacoesCFOP.ValidarCFOPCompraIndustrializacao(CamposCFOP.CFOP_SPED)
    CompraIndustrializacaoSemST = ValidacoesCFOP.ValidarCFOPCompraIndustrializacaoSemST(CamposCFOP.CFOP_SPED)
    CompraIndustrializacaoComST = ValidacoesCFOP.ValidarCFOPCompraIndustrializacaoComST(CamposCFOP.CFOP_SPED)
    CompraCombustivelRevenda = ValidacoesCFOP.ValidarCFOPCompraCombustiveisRevenda(CamposCFOP.CFOP_SPED)
    CompraCombustivelConsumo = ValidacoesCFOP.ValidarCFOPCompraCombustiveisConsumo(CamposCFOP.CFOP_SPED)
    
    SaidaInterna = ValidacoesCFOP.ValidarCFOPSaidaInterna(CamposCFOP.CFOP_NF)
    SaidaInterestadual = ValidacoesCFOP.ValidarCFOPSaidaInterestadual(CamposCFOP.CFOP_NF)
    Exportacao = ValidacoesCFOP.ValidarCFOPExportacao(CamposCFOP.CFOP_NF)
    VendaComST = ValidacoesCFOP.ValidarCFOPVendaComST(CamposCFOP.CFOP_NF)
    VendaSemST = ValidacoesCFOP.ValidarCFOPVendaSemST(CamposCFOP.CFOP_NF)
    VendaCombustivel = ValidacoesCFOP.ValidarCFOPVendaCombustiveis(CamposCFOP.CFOP_NF)
    
    Call VerificarAtividadeContribuinte
    
End Sub

Public Function VerificarAtividadeContribuinte() As Boolean

Dim IND_ATIV As String
    
    Industrializacao = False
    Comercializacao = False
    
    If SPEDFiscal.dicDados0000 Is Nothing Then Call DadosSPEDFiscal.CarregarDadosRegistro0000
    
    With CamposCFOP
        
        If SPEDFiscal.dicDados0000.Exists(.ARQUIVO) Then
            
            IND_ATIV = Util.RemoverAspaSimples(SPEDFiscal.dicDados0000(.ARQUIVO)(SPEDFiscal.dicTitulos0000("IND_ATIV")))
            If IND_ATIV = "0" Then Industrializacao = True Else Comercializacao = True
            
        End If
        
    End With
    
End Function

Public Function ResetarCamposCFOP()
    
    Dim CamposVazios As CamposValidacaoCFOP
    LSet CamposCFOP = CamposVazios
    
End Function
