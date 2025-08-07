Attribute VB_Name = "clsRegrasDivergencias"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private ValidacoesCFOP As New clsRegrasFiscaisCFOP

Public Function VerificarCampoCFOP(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, _
    ByRef dicTitulos0000 As Dictionary, ByRef dicDados0000 As Dictionary)

Dim i As Byte
Dim CFOP_SPED As Integer, CFOP_NF As Integer
Dim INCONSISTENCIA As String, SUGESTAO$, ARQUIVO$, IND_ATIV$
Dim VL_OPER_NF As Double, VL_OPER_SPED#, VL_DESC_SPED#, VL_IPI_SPED#, VL_ITEM_SPED#, VL_ICMS_ST_SPED#
Dim CompraRevSemST As Boolean, CompraRevComST As Boolean, CompIndustrializacao As Boolean, CompPrestServico As Boolean, _
    EntradaInterna As Boolean, EntradaInterestadual  As Boolean, SaidaInterna As Boolean, SaidaInterestadual  As Boolean, _
    Importacao As Boolean, ImportacaoNF  As Boolean, VendaComST As Boolean, VendaSemST As Boolean, CompraIndSemST As Boolean, _
    UsoConsumoSemST As Boolean, UsoConsumoComST As Boolean, AtivoSemST As Boolean, AtivoComST As Boolean, CompraComb As Boolean, _
    CompraIndComST As Boolean, vIndustria As Boolean, vComercio As Boolean, CompraIndustria As Boolean, CompraCombUso As Boolean, _
    VendaComb As Boolean
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega dados do arquivo
    ARQUIVO = Campos(dicTitulos("ARQUIVO") - i)
    If dicDados0000.Exists(ARQUIVO) Then
        IND_ATIV = Util.ApenasNumeros(dicDados0000(ARQUIVO)(dicTitulos0000("IND_ATIV")))
        If IND_ATIV = "0" Then vIndustria = True Else vComercio = True
    End If
    
    'Carrega Informações da NF
    CFOP_NF = fnExcel.ConverterValores(Campos(dicTitulos("CFOP_NF") - i))
    
    'Carrega informações do SPED
    CFOP_SPED = fnExcel.ConverterValores(Campos(dicTitulos("CFOP_SPED") - i))
    
    'Faz verificações do CFOP de entrada
    EntradaInterna = ValidacoesCFOP.ValidarCFOPEntradaInterna(CFOP_SPED)
    EntradaInterestadual = ValidacoesCFOP.ValidarCFOPEntradaInterestadual(CFOP_SPED)
    Importacao = ValidacoesCFOP.ValidarCFOPImportacao(CFOP_SPED)
    
    UsoConsumoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP_SPED)
    UsoConsumoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP_SPED)
    
    AtivoSemST = ValidacoesCFOP.ValidarCFOPCompraAtivoImobilizadoSemST(CFOP_SPED)
    AtivoComST = ValidacoesCFOP.ValidarCFOPCompraAtivoImobilizadoComST(CFOP_SPED)
    
    CompraRevSemST = ValidacoesCFOP.ValidarCFOPCompraRevendaSemST(CFOP_SPED)
    CompraRevComST = ValidacoesCFOP.ValidarCFOPCompraRevendaComST(CFOP_SPED)
    CompraIndustria = ValidacoesCFOP.ValidarCFOPCompraIndustrializacao(CFOP_SPED)
    CompraIndSemST = ValidacoesCFOP.ValidarCFOPCompraIndustrializacaoSemST(CFOP_SPED)
    CompraIndComST = ValidacoesCFOP.ValidarCFOPCompraIndustrializacaoComST(CFOP_SPED)
    CompraComb = ValidacoesCFOP.ValidarCFOPCompraCombustiveisRevenda(CFOP_SPED)
    CompraCombUso = ValidacoesCFOP.ValidarCFOPCompraCombustiveisConsumo(CFOP_SPED)
    
    'Faz verificações do CFOP da NF
    SaidaInterna = ValidacoesCFOP.ValidarCFOPSaidaInterna(CFOP_NF)
    SaidaInterestadual = ValidacoesCFOP.ValidarCFOPSaidaInterestadual(CFOP_NF)
    ImportacaoNF = ValidacoesCFOP.ValidarCFOPExportacao(CFOP_NF)
    VendaComST = ValidacoesCFOP.ValidarCFOPVendaComST(CFOP_NF)
    VendaSemST = ValidacoesCFOP.ValidarCFOPVendaSemST(CFOP_NF)
    VendaComb = ValidacoesCFOP.ValidarCFOPVendaCombustiveis(CFOP_NF)
    
    'Verifica se o campo CFOP_SPED foi informado
    If CFOP_SPED = 0 Then
        INCONSISTENCIA = "O CFOP_SPED não foi informado"
        SUGESTAO = "Informe um CFOP válido no campo CFOP_SPED"
        Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)
        Exit Function
        
    End If
    
    If vComercio And CompraIndustria Then
        
        INCONSISTENCIA = "CFOP_SPED (" & CFOP_SPED & ") indicando compra para industrialização, mas contribuinte não é indústria (Campo IND_ATIV do registro 0000 = 1 - Outros)"
        SUGESTAO = "Informe um CFOP de compra para revenda"

    'Verifica como as operações de saídas internas foram escrituradas pelo contribuinte
    ElseIf SaidaInterna Then
        
        Select Case True
            
            'Identifica operações internas escrituradas como interestaduais
            Case EntradaInterestadual
                INCONSISTENCIA = "CFOP_NF (" & CFOP_NF & ") indicando saída interna com CFOP_SPED (" & CFOP_SPED & ") indicando entrada interestadual"
                SUGESTAO = "Informe um CFOP de entrada interna no campo CFOP_SPED"
                
            'Identifica operações internas escrituradas como importação
            Case Importacao
                INCONSISTENCIA = "CFOP_NF (" & CFOP_NF & ") indicando saída interna com CFOP_SPED (" & CFOP_SPED & ") indicando importação"
                SUGESTAO = "Informe um CFOP de entrada interna no campo CFOP_SPED"
                
            'Identifica operações de venda internas sem ST escrituradas sem CFOP de compra sem ST
            Case VendaSemST And (Not CompraRevSemST And Not CompraIndSemST And Not UsoConsumoSemST And Not AtivoSemST) And Not CFOP_SPED Like "#9##"
                INCONSISTENCIA = "CFOP_NF (" & CFOP_NF & ") indicando venda interna sem ST com CFOP_SPED (" & CFOP_SPED & ") divergente de operação de compra sem ST"
                SUGESTAO = "Informe um CFOP de compra interna sem ST"
                
            'Identifica operações de venda interna com ST escrituradas como compra sem ST
            Case VendaComST And Not CompraRevComST And Not UsoConsumoComST And Not AtivoComST And Not CFOP_SPED Like "#9##"
                INCONSISTENCIA = "CFOP_NF (" & CFOP_NF & ") indicando venda interna com ST com CFOP_SPED (" & CFOP_SPED & ") divergente de operação de compra com ST"
                SUGESTAO = "Informe um CFOP de compra interna com ST"
            
            'Identifica operações de venda interna de combustíveis e lubrificantes não escrituradas como compra de combustíveis e lubrificantes
            Case VendaComb And Not CompraComb And Not CompraCombUso
                INCONSISTENCIA = "CFOP_NF (" & CFOP_NF & ") indicando venda de combustíveis e lubrificantes com CFOP_SPED (" & CFOP_SPED & ") divergente de operação de compra de combustíveis e lubrificantes"
                SUGESTAO = "Informe um CFOP de compra interna de combustíveis e lubrificantes"
                
            
        End Select
        
    'Verifica como as operações de saídas insterestaduais foram escrituradas pelo contribuinte
    ElseIf SaidaInterestadual Then
        
        Select Case True
            
            'Identifica operações interestaduais escrituradas como internas
            Case EntradaInterna
                INCONSISTENCIA = "CFOP_NF (" & CFOP_NF & ") indicando saída interestadual com CFOP_SPED (" & CFOP_SPED & ") indicando entrada interna"
                SUGESTAO = "Informe um CFOP de entrada interestadual no campo CFOP_SPED"
                
            'Identifica operações interestaduais escrituradas como importação
            Case Importacao
                INCONSISTENCIA = "CFOP_NF (" & CFOP_NF & ") indicando saída interestadual com CFOP_SPED (" & CFOP_SPED & ") indicando importação"
                SUGESTAO = "Informe um CFOP de entrada interestadual no campo CFOP_SPED"
                
            'Identifica operações de venda interestadual com ST escrituradas como compra sem ST
            Case VendaComST And Not CompraRevComST And Not UsoConsumoComST And Not AtivoComST And Not CompraIndComST And Not CFOP_SPED Like "#9##"
                INCONSISTENCIA = "CFOP_NF (" & CFOP_NF & ") indicando venda interestadual com ST com CFOP_SPED (" & CFOP_SPED & ") divergente de operação de compra com ST"
                SUGESTAO = "Informe um CFOP de compra interestadual com ST"
                
            'Identifica operações de venda interestadual de combustíveis e lubrificantes não escrituradas como compra de combustíveis e lubrificantes
            Case VendaComb And Not CompraComb And Not CompraCombUso
                INCONSISTENCIA = "CFOP_NF (" & CFOP_NF & ") indicando venda de combustíveis e lubrificantes com CFOP_SPED (" & CFOP_SPED & ") divergente de operação de compra de combustíveis e lubrificantes"
                SUGESTAO = "Informe um CFOP de compra interestadual de combustíveis e lubrificantes"
                
        End Select
        
    'Verifica como as operações de importação foram escrituradas pelo contribuinte
    ElseIf ImportacaoNF Then
        
        Select Case True
            
            'Identifica operações de importação escrituradas como internas
            Case EntradaInterna
                INCONSISTENCIA = "CFOP_NF (" & CFOP_NF & ") indicando importação com CFOP_SPED (" & CFOP_SPED & ") indicando entrada interna"
                SUGESTAO = "Informe um CFOP de importação no campo CFOP_SPED"
                
            'Identifica operações de importação escrituradas como interestaduais
            Case EntradaInterestadual
                INCONSISTENCIA = "CFOP_NF (" & CFOP_NF & ") indicando importação com CFOP_SPED (" & CFOP_SPED & ") indicando entrada interestadual"
                SUGESTAO = "Informe um CFOP de importação no campo CFOP_SPED"
                
        End Select
        
    End If
    
    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)
    
End Function

Public Function VerificarCampoVL_ITEM(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim VL_ITEM_NF As Double, VL_ITEM_SPED#
Dim INCONSISTENCIA As String, SUGESTAO$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    VL_ITEM_NF = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM_NF") - i), True, 2)
    
    'Carrega informações do SPED
    VL_ITEM_SPED = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM_SPED") - i), True, 2)
            
    If VL_ITEM_NF <> VL_ITEM_SPED Then
        INCONSISTENCIA = "Os valores dos campos VL_ITEM_NF e VL_ITEM_SPED estão divergentes"
        SUGESTAO = "Informar o mesmo valor do campo VL_ITEM_NF para o campo VL_ITEM_SPED"
    End If

    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)

End Function

Public Function VerificarCampoDESCR_ITEM(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim DESCR_ITEM_NF As String, DESCR_ITEM_SPED$, COD_ITEM_SPED$, INCONSISTENCIA$, SUGESTAO$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    DESCR_ITEM_NF = Campos(dicTitulos("DESCR_ITEM_NF") - i)
    
    'Carrega informações do SPED
    DESCR_ITEM_SPED = Campos(dicTitulos("DESCR_ITEM_SPED") - i)
    COD_ITEM_SPED = Campos(dicTitulos("COD_ITEM_SPED") - i)
    
    If DESCR_ITEM_SPED = "ITEM NÃO IDENTIFICADO" Then
        INCONSISTENCIA = "O código " & COD_ITEM_SPED & " informado no campo COD_ITEM_SPED não possui cadastro no registro 0200"
        SUGESTAO = "Cadastrar item no registro 0200"
    End If
    
    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)
    
End Function

Public Function VerificarCampoVL_BC_IPI(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim INCONSISTENCIA As String, SUGESTAO$
Dim VL_BC_IPI_NF As Double, VL_BC_IPI_SPED#
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    VL_BC_IPI_NF = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_IPI_NF") - i), True, 2)
    
    'Carrega informações do SPED
    VL_BC_IPI_SPED = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_IPI_SPED") - i), True, 2)
    
    If VL_BC_IPI_NF <> VL_BC_IPI_SPED And (VL_BC_IPI_NF + VL_BC_IPI_SPED > 0) Then
        INCONSISTENCIA = "O valor do campo VL_BC_IPI_SPED deve igual ao valor do campo VL_BC_IPI_NF"
        SUGESTAO = "Informar o mesmo valor do campo VL_BC_IPI_NF para o campo VL_BC_IPI_SPED"
    
    End If
    
    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)
    
End Function

Public Function VerificarCampoALIQ_IPI(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim INCONSISTENCIA As String, SUGESTAO$
Dim ALIQ_IPI_NF As Double, ALIQ_IPI_SPED#
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    ALIQ_IPI_NF = fnExcel.ConverterValores(Campos(dicTitulos("ALIQ_IPI_NF") - i), True, 2)
    
    'Carrega informações do SPED
    ALIQ_IPI_SPED = fnExcel.ConverterValores(Campos(dicTitulos("ALIQ_IPI_SPED") - i), True, 2)
    
    If ALIQ_IPI_NF <> ALIQ_IPI_SPED And (ALIQ_IPI_NF + ALIQ_IPI_SPED > 0) Then
        INCONSISTENCIA = "O valor do campo ALIQ_IPI_SPED deve igual ao valor do campo ALIQ_IPI_NF"
        SUGESTAO = "Informar o mesmo valor do campo ALIQ_IPI_NF para o campo ALIQ_IPI_SPED"
    
    End If
    
    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)
    
End Function

Public Function VerificarCampoVL_IPI(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim INCONSISTENCIA As String, SUGESTAO$
Dim VL_IPI_NF As Double, VL_IPI_SPED#
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    VL_IPI_NF = fnExcel.ConverterValores(Campos(dicTitulos("VL_IPI_NF") - i), True, 2)
    
    'Carrega informações do SPED
    VL_IPI_SPED = fnExcel.ConverterValores(Campos(dicTitulos("VL_IPI_SPED") - i), True, 2)
    
    If VL_IPI_NF <> VL_IPI_SPED And (VL_IPI_NF + VL_IPI_SPED > 0) Then
        INCONSISTENCIA = "O valor do campo VL_IPI_SPED deve igual ao valor do campo VL_IPI_NF"
        SUGESTAO = "Informar o mesmo valor do campo VL_IPI_NF para o campo VL_IPI_SPED"
    
    End If
    
    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)
    
End Function

Public Function VerificarCampoVL_DESC(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim INCONSISTENCIA As String, SUGESTAO$
Dim VL_DESC_NF As Double, VL_DESC_SPED#
Dim i As Byte

    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    VL_DESC_NF = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESC_NF") - i), True, 2)
    
    'Carrega informações do SPED
    VL_DESC_SPED = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESC_SPED") - i), True, 2)
        
    If VL_DESC_NF <> VL_DESC_SPED Then
        INCONSISTENCIA = "O valor do campo VL_DESC_SPED deve igual ao valor do campo VL_DESC_NF"
        SUGESTAO = "Informar o mesmo valor do campo VL_DESC_NF para o campo VL_DESC_SPED"

    End If

    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)

End Function

Public Function VerificarCampoVL_BC_ICMS_ST(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Byte
Dim INCONSISTENCIA As String, SUGESTAO$
Dim VL_BC_ICMS_ST_NF As Double, VL_BC_ICMS_ST_SPED#
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    VL_BC_ICMS_ST_NF = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_ICMS_ST_NF") - i), True, 2)
    
    'Carrega informações do SPED
    VL_BC_ICMS_ST_SPED = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_ICMS_ST_SPED") - i), True, 2)
    
    If VL_BC_ICMS_ST_NF <> VL_BC_ICMS_ST_SPED And (VL_BC_ICMS_ST_NF + VL_BC_ICMS_ST_SPED) > 0 Then
        INCONSISTENCIA = "O valor do campo VL_BC_ICMS_ST_SPED deve igual ao valor do campo VL_BC_ICMS_ST_NF"
        SUGESTAO = "Informar o mesmo valor do campo VL_BC_ICMS_ST_NF para o campo VL_BC_ICMS_ST_SPED"

    End If

    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)

End Function

Public Function VerificarCampoVL_ICMS_ST(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Byte
Dim INCONSISTENCIA As String, SUGESTAO$
Dim VL_ICMS_ST_NF As Double, VL_ICMS_ST_SPED#
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    VL_ICMS_ST_NF = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS_ST_NF") - i), True, 2)
    
    'Carrega informações do SPED
    VL_ICMS_ST_SPED = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS_ST_SPED") - i), True, 2)
    
    If VL_ICMS_ST_NF <> VL_ICMS_ST_SPED And (VL_ICMS_ST_NF + VL_ICMS_ST_SPED) > 0 Then
        INCONSISTENCIA = "O valor do campo VL_ICMS_ST_SPED deve igual ao valor do campo VL_ICMS_ST_NF"
        SUGESTAO = "Informar o mesmo valor do campo VL_ICMS_ST_NF para o campo VL_ICMS_ST_SPED"

    End If

    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)

End Function

Public Function VerificarCampoALIQ_ST(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Byte
Dim INCONSISTENCIA As String, SUGESTAO$
Dim ALIQ_ST_NF As Double, ALIQ_ST_SPED#
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    ALIQ_ST_NF = fnExcel.ConverterValores(Campos(dicTitulos("ALIQ_ST_NF") - i), True, 2)
    
    'Carrega informações do SPED
    ALIQ_ST_SPED = fnExcel.ConverterValores(Campos(dicTitulos("ALIQ_ST_SPED") - i), True, 2)
    
    If ALIQ_ST_NF <> ALIQ_ST_SPED And (ALIQ_ST_NF + ALIQ_ST_SPED) > 0 Then
        INCONSISTENCIA = "O valor do campo ALIQ_ST_SPED deve igual ao valor do campo ALIQ_ST_NF"
        SUGESTAO = "Informar o mesmo valor do campo ALIQ_ST_NF para o campo ALIQ_ST_SPED"
        
    End If
    
    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)
    
End Function

Public Function VerificarCampoVL_OPER(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Byte
Dim INCONSISTENCIA As String, SUGESTAO$
Dim VL_OPER_NF As Double, VL_OPER_SPED#, VL_DESC_SPED#, VL_IPI_SPED#, VL_ITEM_SPED#, VL_ICMS_ST_SPED#
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    VL_OPER_NF = fnExcel.ConverterValores(Campos(dicTitulos("VL_OPER_NF") - i), True, 2)
    
    'Carrega informações do SPED
    VL_IPI_SPED = fnExcel.ConverterValores(Campos(dicTitulos("VL_IPI_SPED") - i), True, 2)
    VL_DESC_SPED = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESC_SPED") - i), True, 2)
    VL_ITEM_SPED = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM_SPED") - i), True, 2)
    VL_ICMS_ST_SPED = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS_ST_SPED") - i), True, 2)
    
    VL_OPER_SPED = fnExcel.ConverterValores(VL_ITEM_SPED - VL_DESC_SPED + VL_IPI_SPED + VL_ICMS_ST_SPED, True, 2)
    Campos(dicTitulos("VL_OPER_SPED") - i) = VL_OPER_SPED
    
    If VL_OPER_NF <> VL_OPER_SPED Then
        INCONSISTENCIA = "O valor da operação (VL_ITEM + VL_ICMS_ST + VL_IPI - VL_DESC) está divergente entre o XML e o SPED (campos VL_OPER_NF e VL_OPER_SPED)"
        SUGESTAO = "Identificar causa da divergencia e corrigir"

    End If
    
    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)
    
End Function

Public Function VerificarCampoVL_ICMS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim VL_ICMS_NF As Double, VL_ICMS_SPED#
Dim CFOP_NF As Integer, CFOP_SPED%
Dim CST_CSOSN_NF As String
Dim i As Byte
Dim CompRevendaSemST As Boolean, CompRevendaComST As Boolean, CompIndustrializacao As Boolean, _
    CompPrestServico As Boolean, FatSemST As Boolean, FatComST As Boolean, CSOSN_NF As Boolean, _
    CompraComb As Boolean, CompraCombUso As Boolean, FatComb As Boolean, CompraUso As Boolean, CompraUsoST As Boolean
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    CST_CSOSN_NF = Util.ApenasNumeros(Campos(dicTitulos("CST_ICMS_NF") - i))
    If VBA.Len(CST_CSOSN_NF) = 4 Then CSOSN_NF = True Else CSOSN_NF = False
    
    'Carrega Informações da NF
    CFOP_NF = fnExcel.ConverterValores(Campos(dicTitulos("CFOP_NF") - i))
    VL_ICMS_NF = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS_NF") - i))
    
    'Carrega informações do SPED
    CFOP_SPED = fnExcel.ConverterValores(Campos(dicTitulos("CFOP_SPED") - i))
    VL_ICMS_SPED = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS_SPED") - i))
    
    'Faz verificações do CFOP de entrada
    CompRevendaSemST = ValidacoesCFOP.ValidarCFOPCompraRevendaSemST(CFOP_SPED)
    CompRevendaComST = ValidacoesCFOP.ValidarCFOPCompraRevendaComST(CFOP_SPED)
    CompIndustrializacao = ValidacoesCFOP.ValidarCFOPCompraIndustrializacao(CFOP_SPED)
    CompPrestServico = ValidacoesCFOP.ValidarCFOPCompraPrestacaoServico(CFOP_SPED)
    CompraComb = ValidacoesCFOP.ValidarCFOPCompraCombustiveisRevenda(CFOP_SPED)
    CompraCombUso = ValidacoesCFOP.ValidarCFOPCompraCombustiveisConsumo(CFOP_SPED)
    CompraUso = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP_SPED)
    CompraUsoST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP_SPED)
    
    'Faz verificações do CFOP de saída
    FatSemST = ValidacoesCFOP.ValidarCFOPVendaSemST(CFOP_NF)
    FatComST = ValidacoesCFOP.ValidarCFOPVendaComST(CFOP_NF)
    FatComb = ValidacoesCFOP.ValidarCFOPVendaCombustiveis(CFOP_NF)
    
    If ApropriarCreditosICMS And Not CompraCombUso And Not CompraUso And Not CompraUsoST And VL_ICMS_NF <> VL_ICMS_SPED Then
        
        Call Util.GravarSugestao(Campos, dicTitulos, _
            INCONSISTENCIA:="Campo VL_ICMS_NF divergente do campo VL_ICMS_SPED", _
            SUGESTAO:="Apropiar crédito do ICMS", _
            dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
            
        Exit Function
        
    End If
    
    If FatSemST And CFOP_NF Like "5*" And (CompRevendaSemST Or CompIndustrializacao) And VL_ICMS_NF > 0 And VL_ICMS_SPED <= 0 And Not CSOSN_NF Then
        
        Call Util.GravarSugestao(Campos, dicTitulos, _
            INCONSISTENCIA:="Venda sem ST (CFOP_NF) e operação de entrada com direito a crédito (CFOP_SPED) sem aproveitamento do ICMS", _
            SUGESTAO:="Apropiar crédito do ICMS", _
            dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
            
    ElseIf FatComST And CFOP_NF Like "5*" And CompRevendaComST And VL_ICMS_SPED > 0 Then
        
        Call Util.GravarSugestao(Campos, dicTitulos, _
            INCONSISTENCIA:="Venda com ST (" & CFOP_NF & ") e operação de entrada sem direito a crédito (" & CFOP_SPED & ") com aproveitamento do ICMS", _
            SUGESTAO:="Zerar campos do ICMS", _
            dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
            
    ElseIf FatComb And CFOP_NF Like "5*" And (CompraComb Or CompraCombUso) And VL_ICMS_SPED > 0 Then
        
        Call Util.GravarSugestao(Campos, dicTitulos, _
            INCONSISTENCIA:="Venda de combustíveis e lubrificantes (" & CFOP_NF & ") e operação de entrada sem direito a crédito (" & CFOP_SPED & ") com aproveitamento do ICMS", _
            SUGESTAO:="Zerar campos do ICMS", _
            dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
            
    End If
    
End Function

Public Function VerificarCampoCEST(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim vCEST_NF As Boolean, vCEST_SPED As Boolean
Dim INCONSISTENCIA As String, SUGESTAO$
Dim CEST_NF As String, CEST_SPED$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    CEST_NF = Util.ApenasNumeros(Campos(dicTitulos("CEST_NF") - i))
    vCEST_NF = RegrasFiscais.Geral.ValidarCEST(CEST_NF)
    
    'Carrega informações do SPED
    CEST_SPED = Util.ApenasNumeros(Campos(dicTitulos("CEST_SPED") - i))
    vCEST_SPED = RegrasFiscais.Geral.ValidarCEST(CEST_SPED)
    
    'Verifica se o CEST informado na NF é válido
'    If Not vCEST_NF And CEST_NF <> "" Then
'        INCONSISTENCIA = "O valor informado no campo CEST_NF está inválido"
'        SUGESTAO = "Apagar valor do CEST informado no campo CEST_NF"

    'Verifica se o CEST informado no SPED é válido
    If Not vCEST_SPED And CEST_SPED <> "" Then
        INCONSISTENCIA = "O valor informado no campo CEST_SPED está inválido"
        SUGESTAO = "Apagar CEST informado no SPED"

    'Verifica se há divergências entre os campos CEST_NF e CEST_SPED
    ElseIf CEST_NF <> CEST_SPED And vCEST_NF Then
        INCONSISTENCIA = "Os campos CEST_NF e CEST_SPED estão divergentes"
        SUGESTAO = "Informar o mesmo código CEST do XML para o SPED"

    End If

    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)

End Function

Public Function VerificarCampoCST_ICMS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim ORIG_NF As String, ORIG_SPED As String
Dim CST_ICMS_NF As String, CST_ICMS_SPED$, INCONSISTENCIA$, SUGESTAO$
Dim vCSOSN_NF As Boolean
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    CST_ICMS_NF = Util.ApenasNumeros(Campos(dicTitulos("CST_ICMS_NF") - i))
    ORIG_NF = VBA.Left(CST_ICMS_NF, 1)
    
    'Carrega informações do SPED
    CST_ICMS_SPED = Util.ApenasNumeros(Campos(dicTitulos("CST_ICMS_SPED") - i))
    ORIG_SPED = VBA.Left(CST_ICMS_SPED, 1)
    
    'Verificações
    vCSOSN_NF = VBA.Len(CST_ICMS_NF) = 4
    
    If ORIG_NF = "1" And ORIG_SPED <> "2" Then
        INCONSISTENCIA = "O dígito de origem do CST_ICMS_SPED deve ser igual a 2"
        SUGESTAO = "Mudar o dígito de origem do CST_ICMS_SPED para 2"
        
    ElseIf ORIG_NF = "6" And ORIG_SPED <> "7" Then
        INCONSISTENCIA = "O dígito de origem do CST_ICMS_SPED deve ser igual a 7"
        SUGESTAO = "Mudar o dígito de origem do CST_ICMS_SPED para 7"
    
    'Verifica as operações com CSOSN
    ElseIf vCSOSN_NF Then
        
        Select Case True
                
            'Identifica CSOSN de operações com permissão de crédito
            Case CST_ICMS_NF Like "#101" And Not CST_ICMS_SPED Like "*90"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica operação com permissão de crédito do ICMS"
                SUGESTAO = "Informar o CST 90 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações sem permissão de crédito
            Case CST_ICMS_NF Like "#102" And Not CST_ICMS_SPED Like "*90"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica operação sem permissão de crédito do ICMS"
                SUGESTAO = "Informar o CST 90 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações com isenção do ICMS
            Case CST_ICMS_NF Like "#103" And Not CST_ICMS_SPED Like "*40"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica operação isenta"
                SUGESTAO = "Informar o CST 40 da tabela B para o campo CST_ICMS_SPED"
            
            'Identifica CSOSN de operações sujeitas a cobrança da Substituição tributária do ICMS
            Case CST_ICMS_NF Like "#20#" And Not CST_ICMS_SPED Like "*60"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica operação com cobrança da Substituição Tributária do ICMS"
                SUGESTAO = "Informar o CST 60 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações com imunidade
            Case CST_ICMS_NF Like "#300" And Not CST_ICMS_SPED Like "*41"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica operação com imune"
                SUGESTAO = "Informar o CST 41 da tabela B para o campo CST_ICMS_SPED"
            
            'Identifica CSOSN de operações não-tributadas
            Case CST_ICMS_NF Like "#400" And Not CST_ICMS_SPED Like "*41"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica operação não-tributada"
                SUGESTAO = "Informar o CST 41 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações sujeitas a Substituição tributária do ICMS
            Case CST_ICMS_NF Like "#500" And Not CST_ICMS_SPED Like "*60"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica operação com ICMS cobrado anteriormente por substituição"
                SUGESTAO = "Informar o CST 60 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações com tributação do ICMS
            Case CST_ICMS_NF Like "#900" And Not CST_ICMS_SPED Like "*90"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica outras operações"
                SUGESTAO = "Informar o CST 90 da tabela B para o campo CST_ICMS_SPED"

        End Select
'
'    ElseIf (CST_ICMS_NF Like "#500" Or CST_ICMS_NF Like "#20#") And vCSOSN_NF And Not CST_ICMS_SPED Like "*60" Then
'
'        Call Util.GravarSugestao(Campos, dicTitulos, _
'            Inconsistencia:="Operação sujeita a ST de fornecedor optante pelo SMIPLES NACIONAL", _
'            Sugestao:="Informar o CST 60 da tabela B para o campo CST_ICMS_SPED", _
'            dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
'
'    ElseIf CST_ICMS_NF Like "#103" And vCSOSN_NF And Not CST_ICMS_SPED Like "*40" Then
'
'        Call Util.GravarSugestao(Campos, dicTitulos, _
'            Inconsistencia:="Operação Isenta de fornecedor optante pelo SMIPLES NACIONAL", _
'            Sugestao:="Informar o CST 40 da tabela B para o campo CST_ICMS_SPED", _
'            dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
'
'    ElseIf (CST_ICMS_NF Like "#300" Or CST_ICMS_NF Like "#400" Or CST_ICMS_NF Like "#103") And vCSOSN_NF And Not CST_ICMS_SPED Like "*41" Then
'
'        Call Util.GravarSugestao(Campos, dicTitulos, _
'            Inconsistencia:="Operação Imune ou Não Tributada de fornecedor optante pelo SMIPLES NACIONAL", _
'            Sugestao:="Informar o CST 41 da tabela B para o campo CST_ICMS_SPED", _
'            dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
'
'    ElseIf (CST_ICMS_NF Like "#900" Or CST_ICMS_NF Like "#10#") And Not CST_ICMS_NF Like "#103" And vCSOSN_NF And Not CST_ICMS_SPED Like "*90" Then
'
'        Call Util.GravarSugestao(Campos, dicTitulos, _
'            Inconsistencia:="Operação sem ST de fornecedor optante pelo SMIPLES NACIONAL", _
'            Sugestao:="Informar o CST 90 da tabela B para o campo CST_ICMS_SPED", _
'            dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
                                
    End If

    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)

End Function

Public Function VerificarCampoCOD_BARRA(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Byte
Dim INCONSISTENCIA As String, SUGESTAO$
Dim COD_BARRA_NF As String, COD_BARRA_SPED$
Dim vCOD_BARRA_NF As Boolean, vCOD_BARRA_SPED As Boolean
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    COD_BARRA_NF = Util.ApenasNumeros(Campos(dicTitulos("COD_BARRA_NF") - i))
    vCOD_BARRA_NF = RegrasFiscais.Geral.CodigoBarras.ValidarCodigoBarras(COD_BARRA_NF)
    
    'Carrega informações do SPED
    COD_BARRA_SPED = Util.ApenasNumeros(Campos(dicTitulos("COD_BARRA_SPED") - i))
    vCOD_BARRA_SPED = RegrasFiscais.Geral.CodigoBarras.ValidarCodigoBarras(COD_BARRA_SPED)
    
    If Not vCOD_BARRA_SPED And COD_BARRA_SPED <> "" Then
        INCONSISTENCIA = "O valor informado no campo COD_BARRA_SPED está inválido"
        SUGESTAO = "Apagar código de barras informado no SPED"
        
    ElseIf COD_BARRA_NF <> COD_BARRA_SPED And vCOD_BARRA_NF Then
        INCONSISTENCIA = "Os campos COD_BARRA_NF e COD_BARRA_SPED estão divergentes"
        SUGESTAO = "Informar o mesmo código de barras do XML para o SPED"
        
    End If
    
    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)
    
End Function

Public Function VerificarCampoCOD_PART(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim COD_PART_NF As String, COD_PART_SPED$
Dim INCONSISTENCIA As String, SUGESTAO$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
        
    'Carrega Informações da NF
    COD_PART_NF = Util.ApenasNumeros(Campos(dicTitulos("COD_PART_NF") - i))
    
    'Carrega informações do SPED
    COD_PART_SPED = Util.ApenasNumeros(Campos(dicTitulos("COD_PART_SPED") - i))
    
    If COD_PART_NF <> COD_PART_SPED Then
        INCONSISTENCIA = "Os participantes informados nos campos COD_PART_NF e COD_PART_SPED estão divergentes"
        SUGESTAO = "Informar o mesmo participante do XML para o SPED"

    End If
    
    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)
        
End Function

Public Function VerificarCampoDT_DOC(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim INCONSISTENCIA As String, SUGESTAO$
Dim DT_DOC_NF As String, DT_DOC_SPED$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    DT_DOC_NF = fnExcel.FormatarData(Campos(dicTitulos("DT_DOC_NF") - i))
    
    'Carrega informações do SPED
    DT_DOC_SPED = fnExcel.FormatarData(Campos(dicTitulos("DT_DOC_SPED") - i))
    
    If DT_DOC_NF <> DT_DOC_SPED Then
        INCONSISTENCIA = "Os campos DT_DOC_NF e DT_DOC_SPED estão divergentes"
        SUGESTAO = "Informar a mesma data do XML no SPED"

    End If
    
    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)

End Function

Public Function VerificarCampoDT_E_S(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim INCONSISTENCIA As String, SUGESTAO$
Dim DT_E_S_NF As String, DT_E_S_SPED$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    DT_E_S_NF = fnExcel.FormatarData(Campos(dicTitulos("DT_E_S_NF") - i))
    
    'Carrega informações do SPED
    DT_E_S_SPED = fnExcel.FormatarData(Campos(dicTitulos("DT_E_S_SPED") - i))
    
    If DT_E_S_NF > DT_E_S_SPED Then
        INCONSISTENCIA = "A data de entrada informada no SPED (DT_E_S_SPED) é anterior a data de emissão do documento fiscal (DT_DOC_NF)"
        SUGESTAO = "Informar uma data igual ou posterior a data de emissão do documnto fiscal"

    End If
    
    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)

End Function

Public Function VerificarCampoCOD_SIT(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim INCONSISTENCIA As String, SUGESTAO$
Dim COD_SIT_NF As String, COD_SIT_SPED$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
        
    'Carrega Informações da NF
    COD_SIT_NF = Util.ApenasNumeros(Campos(dicTitulos("COD_SIT_NF") - i))
    
    'Carrega informações do SPED
    COD_SIT_SPED = Util.ApenasNumeros(Campos(dicTitulos("COD_SIT_SPED") - i))
    
    If COD_SIT_NF <> COD_SIT_SPED Then
        INCONSISTENCIA = "A situação do documento fiscal no XML (COD_SIT_NF) e no SPED (COD_SIT_SPED) estão divergentes"
        SUGESTAO = "Informar a mesma situação do XML para o SPED"

    End If
    
    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)
    
End Function

Public Function VerificarCampoCOD_ITEM(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim INCONSISTENCIA As String, SUGESTAO$
Dim COD_ITEM_NF As String, COD_ITEM_SPED$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
        
    'Carrega Informações da NF
    COD_ITEM_NF = Util.ApenasNumeros(Campos(dicTitulos("COD_ITEM_NF") - i))
    
    'Carrega informações do SPED
    COD_ITEM_SPED = Util.ApenasNumeros(Campos(dicTitulos("COD_ITEM_SPED") - i))
    
    If COD_ITEM_NF = COD_ITEM_SPED Then
        INCONSISTENCIA = "Os campos COD_ITEM_NF (" & COD_ITEM_NF & ") e COD_ITEM_SPED (" & COD_ITEM_SPED & ") estão iguais"
        SUGESTAO = "O lançamento de itens no SPED deve conter o COD_ITEM do contribuinte do arquivo não o do fornecedor"

    End If
    
    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)
    
End Function

Public Function VerificarCampoQTD(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim QTD_NF As Double, QTD_SPED#
Dim INCONSISTENCIA As String, SUGESTAO$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    QTD_NF = fnExcel.ConverterValores(Campos(dicTitulos("QTD_NF") - i), True, 2)
    
    'Carrega informações do SPED
    QTD_SPED = fnExcel.ConverterValores(Campos(dicTitulos("QTD_SPED") - i), True, 2)
            
    If QTD_NF <> QTD_SPED Then
        INCONSISTENCIA = "Os valores dos campos QTD_NF e QTD_SPED estão divergentes"
        SUGESTAO = "Informar o mesmo valor do campo QTD_NF para o campo QTD_SPED"
    End If

    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)

End Function

Public Function VerificarCampoUNID(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim UNID_NF As String, UNID_SPED$, INCONSISTENCIA$, SUGESTAO$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    UNID_NF = fnExcel.FormatarTexto(Campos(dicTitulos("UNID_NF") - i))
    
    'Carrega informações do SPED
    UNID_SPED = fnExcel.FormatarTexto(Campos(dicTitulos("UNID_SPED") - i))
            
    If UNID_NF <> UNID_SPED Then
        INCONSISTENCIA = "Os valores dos campos UNID_NF e UNID_SPED estão divergentes"
        SUGESTAO = "Informar o mesmo valor do campo UNID_NF para o campo UNID_SPED"
    End If

    'Registra a inconsistência caso exista
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA, SUGESTAO, dicInconsistenciasIgnoradas)

End Function
