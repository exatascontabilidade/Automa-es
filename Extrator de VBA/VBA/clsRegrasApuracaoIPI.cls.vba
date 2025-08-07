Attribute VB_Name = "clsRegrasApuracaoIPI"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
            
Private ValidacoesCFOP As New clsRegrasFiscaisCFOP

Public Function VerificarCampoCFOP(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Integer
Dim VL_IPI As Double
Dim CFOP As String, CST_IPI$, INCONSISTENCIA$, SUGESTAO$
Dim vCFOP As Boolean, tCFOP As Boolean, vEntradaSemST As Boolean, vEntradaComST As Boolean, vCST_IPITributado As Boolean, _
    vCST_IPIIsento As Boolean, vCST_IPIComST As Boolean
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega informações do relatório
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_IPI = Util.ApenasNumeros(Campos(dicTitulos("CST_IPI") - i))
    VL_IPI = fnExcel.ConverterValores(Campos(dicTitulos("VL_IPI") - i))
    
    'Gera verificações sobre os dados do relatório
'    tCFOP = VBA.Len(CFOP) > 0 And VBA.Len(CFOP) < 4
'    vCFOP = ValidacoesCFOP.ValidarCFOP(CFOP)
    vEntradaSemST = CFOP < 4000 And (Not CFOP Like "#4##" And Not CFOP Like "#65#" And Not CFOP Like "#9##")
    vEntradaComST = CFOP < 4000 And Not CFOP Like "#9##" And (CFOP Like "#4##" Or CFOP Like "#65#")
    vCST_IPIComST = CST_IPI Like "*10" Or CST_IPI Like "*60" Or CST_IPI Like "*61" Or CST_IPI Like "*70"
    vCST_IPIIsento = CST_IPI Like "*30" Or CST_IPI Like "*40" Or CST_IPI Like "*41"
    vCST_IPITributado = CST_IPI Like "*00" Or CST_IPI Like "*20"
    
    'Operação sujeita a ST com CST_IPI incorreto
    If CFOP = "" Then
        INCONSISTENCIA = "O campo CFOP não foi informado"
        SUGESTAO = "Informe um CFOP válido para a operação"
    
    'Verifica a quantidade de dígitos do CFOP
    ElseIf tCFOP Then
        INCONSISTENCIA = "O campo CFOP deve possuir exatamente 4 dígitos"
        SUGESTAO = "Informe um CFOP válido para a operação"
    
    'Verifica se o CFOP informado é válido
    ElseIf Not vCFOP And CFOP <> "" Then
        INCONSISTENCIA = "O CFOP informado não existe na tabela CFOP"
        SUGESTAO = "Informe um CFOP válido para a operação"
        
    'Verifica operações com CFOP indicando entrada sem ST
    ElseIf vEntradaSemST Then
        
        Select Case True
            
            'Identifica operações de entrada sem ST com CST_IPI e VL_IPI indicando operação com ST
            Case vCST_IPIComST And VL_IPI = 0
                INCONSISTENCIA = "O CFOP (" & CFOP & ") indica entrada sem ST com o campo CST_IPI (" & CST_IPI & ") indicando operacao com ST sem aproveitamento de crédito de IPI (VL_IPI = R$ " & VBA.Format(VL_IPI, "#0.00") & ")"
                SUGESTAO = "Informe um CFOP válido para a operação"
        
        End Select
    
    'Verifica operações com CFOP indicando entrada com ST
    ElseIf vEntradaComST Then
        
        Select Case True
            
            'Identifica operações com CFOP incando operação de entrada com ST com CST_IPI e VL_IPI indicando operação tributada
            Case vCST_IPITributado And VL_IPI > 0
                INCONSISTENCIA = "O CFOP (" & CFOP & ") indica entrada com ST com o campo CST_IPI (" & CST_IPI & ") indicando operacao tributada com aproveitamento de crédito de IPI (VL_IPI = R$ " & VBA.Format(VL_IPI, "#0.00") & ")"
                SUGESTAO = "Informe um CFOP válido para a operação"
        
        End Select
        
        
    End If

    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)

End Function

Public Function ValidarCampo_VL_IPI(Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Integer
Dim VL_IPI As Double, VL_BC_IPI#, ALIQ_IPI#, VL_IPI_CALC#, DIFERENCA#
Dim CFOP As String, CST_IPI$, INCONSISTENCIA, SUGESTAO$
Dim vCompraUsoConsumoSemST As Boolean, vCompraUsoConsumoComST As Boolean, vCompraAtivoImobilizadoSemST As Boolean, _
    vCompraAtivoImobilizadoComST As Boolean, vCompraCombustivelRevenda As Boolean, vCompraCombustivelConsumo As Boolean, _
    vCompraRevendaComST As Boolean
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega informações do relatório
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_IPI = Util.ApenasNumeros(Campos(dicTitulos("CST_IPI") - i))
    VL_BC_IPI = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_IPI") - i))
    ALIQ_IPI = fnExcel.ConverterValores(Campos(dicTitulos("ALIQ_IPI") - i))
    VL_IPI = fnExcel.ConverterValores(Campos(dicTitulos("VL_IPI") - i))
    VL_IPI_CALC = fnExcel.ConverterValores(VL_BC_IPI * ALIQ_IPI, True, 2)
    
    DIFERENCA = VBA.Round(VBA.Abs(VL_IPI_CALC - VL_IPI), 2)
    
    'Gera verificações sobre os dados do relatório
    vCompraUsoConsumoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP)
    vCompraUsoConsumoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP)
    vCompraAtivoImobilizadoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP)
    vCompraAtivoImobilizadoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP)
    vCompraCombustivelRevenda = ValidacoesCFOP.ValidarCFOPCompraCombustiveisRevenda(CFOP)
    vCompraCombustivelConsumo = ValidacoesCFOP.ValidarCFOPCompraCombustiveisConsumo(CFOP)
    vCompraRevendaComST = ValidacoesCFOP.ValidarCFOPCompraRevendaComST(CFOP)
    
    'Verifica operações com apropriação indevida do valor do IPI
    If VL_IPI > 0 Then
                
        Select Case True
        
            'Identifica apropriação indevida do IPI em operação de aquisição para uso e consumo
            Case vCompraUsoConsumoSemST Or vCompraUsoConsumoComST
                INCONSISTENCIA = "Apropriação indevida do IPI em operação de aquisição para uso e consumo"
                SUGESTAO = "Zerar campos do IPI"
            
            'Identifica apropriação indevida do IPI em operação de aquisição para o ativo imobilizado
            Case vCompraAtivoImobilizadoSemST Or vCompraAtivoImobilizadoComST
                INCONSISTENCIA = "Apropriação indevida do IPI em operação de aquisição para o ativo imobilizado"
                SUGESTAO = "Zerar campos do IPI"
                
            'Identifica apropriação indevida do IPI em operação de aquisição de combustíveis e lubrificantes
            Case vCompraCombustivelRevenda Or vCompraCombustivelConsumo
                INCONSISTENCIA = "Apropriação indevida do IPI em operação de aquisição de combustíveis e lubrificantes"
                SUGESTAO = "Zerar campos do IPI"
                
            'Identifica apropriação indevida do IPI em operação de compra para revenda com ST
            Case vCompraRevendaComST
                INCONSISTENCIA = "Apropriação indevida do IPI em operação de compra para revenda com ST"
                SUGESTAO = "Zerar campos do IPI"
            
            Case VL_IPI_CALC > VL_IPI
                INCONSISTENCIA = "Cálculo do campo VL_IPI maior que o destacado [Diferença: R$ " & DIFERENCA & "]"
                SUGESTAO = "Recalcular o campos VL_IPI"
            
            Case VL_IPI_CALC < VL_IPI
                INCONSISTENCIA = "Cálculo do campo VL_IPI menor que o destacado [Diferença: R$ " & DIFERENCA & "]"
                SUGESTAO = "Recalcular o campos VL_IPI"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_ALIQ_IPI(Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Integer
Dim ALIQ_IPI As Double, VL_IPI#
Dim CFOP As String, CST_IPI$, INCONSISTENCIA, SUGESTAO$, ALIQ_IPI_FORM$, VL_IPI_FORM$
Dim vCompraUsoConsumoSemST As Boolean, vCompraUsoConsumoComST As Boolean, vCompraAtivoImobilizadoSemST As Boolean, _
    vCompraAtivoImobilizadoComST As Boolean, vCompraCombustivelRevenda As Boolean, vCompraCombustivelConsumo As Boolean

    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega informações do relatório
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_IPI = Util.ApenasNumeros(Campos(dicTitulos("CST_IPI") - i))
    ALIQ_IPI = fnExcel.FormatarValores(Campos(dicTitulos("ALIQ_IPI") - i))
    VL_IPI = fnExcel.FormatarValores(Campos(dicTitulos("VL_IPI") - i))
    ALIQ_IPI_FORM = VBA.Format(ALIQ_IPI, "#0.00%")
    VL_IPI_FORM = VBA.Format(VL_IPI, "R$ #,#0.00")
    
    'Gera verificações sobre os dados do relatório
    vCompraUsoConsumoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP)
    vCompraUsoConsumoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP)
    vCompraAtivoImobilizadoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP)
    vCompraAtivoImobilizadoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP)
    vCompraCombustivelRevenda = ValidacoesCFOP.ValidarCFOPCompraCombustiveisRevenda(CFOP)
    vCompraCombustivelConsumo = ValidacoesCFOP.ValidarCFOPCompraCombustiveisConsumo(CFOP)
    
    'Verifica operações com Alíquota 0
    If ALIQ_IPI = 0 Then
                
        Select Case True
        
            'Identifica apropriação indevida do IPI em operação de aquisição para uso e consumo
            Case VL_IPI > 0
                INCONSISTENCIA = "Alíquota zerada (ALIQ_IPI = 0) com valor destacado de IPI (VL_IPI = " & VL_IPI_FORM & ")"
                SUGESTAO = "Informe uma alíquota de IPI compatível com a operação no campo ALIQ_IPI"

        End Select
        
    ElseIf ALIQ_IPI > 0 Then
    
        Select Case True
        
            'Identifica apropriação indevida do IPI em operação de aquisição para uso e consumo
            Case vCompraUsoConsumoComST Or vCompraUsoConsumoSemST
                INCONSISTENCIA = "Alíquota maior que zero (ALIQ_IPI = " & ALIQ_IPI_FORM & ") em operação de compra para uso e consumo (CFOP = " & CFOP & ")"
                SUGESTAO = "Zerar Alíquota do IPI"

        End Select
    
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_CST_IPI(Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim i As Integer
Dim VL_IPI As Double, VL_BC_IPI#, VL_ITEM#, VL_DESC#
Dim CST_IPI As String, CFOP$, INCONSISTENCIA$, SUGESTAO$
Dim CSTEnt As Boolean, CSTSai As Boolean, vCFOPEntrada As Boolean, vCFOPSaida As Boolean, _
    AquisicaoUsoConsumo As Boolean, AquisicaoAtivoImobilizado As Boolean, AquisicaoCombustivelConsumo As Boolean
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega informações do relatório
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_IPI = Util.ApenasNumeros(Campos(dicTitulos("CST_IPI") - i))
    VL_IPI = fnExcel.ConverterValores(Campos(dicTitulos("VL_IPI") - i))
    VL_BC_IPI = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_IPI") - i))
    VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM") - i))
    VL_DESC = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESC") - i))
    
    AquisicaoUsoConsumo = ValidacoesCFOP.ValidarAquisicaoUsoConsumo(CFOP)
    AquisicaoAtivoImobilizado = ValidacoesCFOP.ValidarAquisicaoAtivoImobilizado(CFOP)
    AquisicaoCombustivelConsumo = ValidacoesCFOP.ValidarCFOPCompraCombustiveisConsumo(CFOP)
    
    'Gera verificações sobre os dados do relatório
    CSTEnt = fnExcel.ConverterValores(CST_IPI) < 50
    CSTSai = fnExcel.ConverterValores(CST_IPI) > 49
    vCFOPEntrada = CFOP < 4000
    vCFOPSaida = CFOP > 4000
    
    'Identifica operações sem informação do CST/IPI
    If CST_IPI = "" Then
        INCONSISTENCIA = "CST_IPI não foi informado"
        SUGESTAO = "informar um valor válido para o campo CST_IPI"
        
    ElseIf AquisicaoCombustivelConsumo Then
        
        Select Case True
            
            'Identifica operações de entrada com CST_IPI de saída
            Case Not CST_IPI Like "49*"
                INCONSISTENCIA = "CST_IPI (" & CST_IPI & ") incompatível com operação de aquisição de combustível para uso e consumo"
                SUGESTAO = "Informar CST_IPI 49 - Outras Entradas"
                
        End Select
        
    ElseIf AquisicaoUsoConsumo Then
        
        Select Case True
            
            'Identifica operações de entrada com CST_IPI de saída
            Case Not CST_IPI Like "49*"
                INCONSISTENCIA = "CST_IPI (" & CST_IPI & ") incompatível com operação de aquisição para uso e consumo"
                SUGESTAO = "Informar CST_IPI 49 - Outras Entradas"
                
        End Select
        
    ElseIf AquisicaoAtivoImobilizado Then
        
        Select Case True
            
            'Identifica operações de entrada com CST_IPI de saída
            Case Not CST_IPI Like "49*"
                INCONSISTENCIA = "CST_IPI (" & CST_IPI & ") incompatível com operação de aquisição para ativo imobilizado"
                SUGESTAO = "Informar CST_IPI 49 - Outras Entradas"
                
        End Select
        
    'Verifica operações de entrada
    ElseIf vCFOPEntrada Then
        
        Select Case True
            
            'Identifica operações de entrada com CST_IPI de saída
            Case CSTSai
                INCONSISTENCIA = "Informado CST_IPI (" & CST_IPI & ") de saída em operação com CFOP (" & CFOP & ") de entrada "
                SUGESTAO = "Informar um CST_IPI de entrada para a operação"
                
            Case CST_IPI Like "00*" And VL_IPI = 0
                INCONSISTENCIA = "Informado CST_IPI (" & CST_IPI & ") de entrada com recuperação de crédito com VL_IPI (R$ " & VL_IPI & ") igual a zero"
                SUGESTAO = "Informar CST_IPI sem aproveitamento de crédito"
                
            Case CST_IPI Like "01*" And VL_IPI > 0
                INCONSISTENCIA = "Informado CST_IPI (" & CST_IPI & ") de entrada a alíquota zero com VL_IPI (R$ " & VL_IPI & ") maior que zero"
                SUGESTAO = "Informar CST_IPI 00 - Entrada com recuperação de crédito"
                
            Case CST_IPI Like "02*" And VL_IPI > 0
                INCONSISTENCIA = "Informado CST_IPI (" & CST_IPI & ") de entrada isenta com VL_IPI (R$ " & VL_IPI & ") maior que zero"
                SUGESTAO = "Informar CST_IPI 00 - Entrada com recuperação de crédito"
                
            Case CST_IPI Like "03*" And VL_IPI > 0
                INCONSISTENCIA = "Informado CST_IPI (" & CST_IPI & ") de entrada não-tributada com VL_IPI (R$ " & VL_IPI & ") maior que zero"
                SUGESTAO = "Informar CST_IPI 00 - Entrada com recuperação de crédito"
                
            Case CST_IPI Like "04*" And VL_IPI > 0
                INCONSISTENCIA = "Informado CST_IPI (" & CST_IPI & ") de entrada imune com VL_IPI (R$ " & VL_IPI & ") maior que zero"
                SUGESTAO = "Informar CST_IPI 00 - Entrada com recuperação de crédito"
            
            Case CST_IPI Like "05*" And VL_IPI > 0
                INCONSISTENCIA = "Informado CST_IPI (" & CST_IPI & ") de entrada com suspensão com VL_IPI (R$ " & VL_IPI & ") maior que zero"
                SUGESTAO = "Informar CST_IPI 00 - Entrada com recuperação de crédito"
                
        End Select
    
    'Verifica operações de saída
    ElseIf vCFOPSaida Then

        Select Case True
            
            'Identifica operações de saída com CST_IPI de entrada
            Case CSTEnt
                INCONSISTENCIA = "Informado CST_IPI (" & CST_IPI & ") de entrada em operação com CFOP (" & CFOP & ") de saída "
                SUGESTAO = "Informar um CST_IPI de saída para a operação"
            
            Case CST_IPI Like "50*" And VL_IPI = 0
                INCONSISTENCIA = "Informado CST_IPI (" & CST_IPI & ") de saída com recuperação de crédito com VL_IPI (R$ " & VL_IPI & ") igual a zero"
                SUGESTAO = "Informar CST_IPI sem aproveitamento de crédito"
                
            Case CST_IPI Like "51*" And VL_IPI > 0
                INCONSISTENCIA = "Informado CST_IPI (" & CST_IPI & ") de saída a alíquota zero com VL_IPI (R$ " & VL_IPI & ") maior que zero"
                SUGESTAO = "Informar CST_IPI 00 - Entrada com recuperação de crédito"
                
            Case CST_IPI Like "52*" And VL_IPI > 0
                INCONSISTENCIA = "Informado CST_IPI (" & CST_IPI & ") de saída isenta com VL_IPI (R$ " & VL_IPI & ") maior que zero"
                SUGESTAO = "Informar CST_IPI 00 - Entrada com recuperação de crédito"
                        
            Case CST_IPI Like "53*" And VL_IPI > 0
                INCONSISTENCIA = "Informado CST_IPI (" & CST_IPI & ") de saída não-tributada com VL_IPI (R$ " & VL_IPI & ") maior que zero"
                SUGESTAO = "Informar CST_IPI 00 - Entrada com recuperação de crédito"
                
            Case CST_IPI Like "54*" And VL_IPI > 0
                INCONSISTENCIA = "Informado CST_IPI (" & CST_IPI & ") de saída imune com VL_IPI (R$ " & VL_IPI & ") maior que zero"
                SUGESTAO = "Informar CST_IPI 00 - Entrada com recuperação de crédito"
            
            Case CST_IPI Like "55*" And VL_IPI > 0
                INCONSISTENCIA = "Informado CST_IPI (" & CST_IPI & ") de saída com suspensão com VL_IPI (R$ " & VL_IPI & ") maior que zero"
                SUGESTAO = "Informar CST_IPI 00 - Entrada com recuperação de crédito"
                
        End Select
    
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_DT_ENT_SAI(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Byte
Dim DT_ENT_SAI As String, CFOP$, INCONSISTENCIA$, SUGESTAO$
    
    If LBound(Campos) = 0 Then i = 1
    
    DT_ENT_SAI = Campos(dicTitulos("DT_ENT_SAI") - i)
    CFOP = Campos(dicTitulos("CFOP") - i)
    
    Select Case True
        
        Case CFOP < 4000 And DT_ENT_SAI = ""
            INCONSISTENCIA = "O campo DT_ENT_SAI não foi informado"
            SUGESTAO = "Informar uma data de entrada para o documento fiscal"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_IND_APUR(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Byte
Dim IND_APUR As String, CFOP$, INCONSISTENCIA$, SUGESTAO$
    
    If LBound(Campos) = 0 Then i = 1
    
    IND_APUR = Campos(dicTitulos("IND_APUR") - i)
    
    Select Case True
        
        Case IND_APUR = ""
            INCONSISTENCIA = "O campo IND_APUR não foi informado"
            SUGESTAO = "Informe um indicador para o período de apuração do IPI"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

