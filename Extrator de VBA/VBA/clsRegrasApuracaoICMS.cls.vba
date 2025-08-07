Attribute VB_Name = "clsRegrasApuracaoICMS"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
            
Private ValidacoesCFOP As New clsRegrasFiscaisCFOP

Public Function VerificarCampoCEST(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim CEST As String, INCONSISTENCIA$, SUGESTAO$
Dim vCEST As Boolean, tCEST As Boolean
Dim i As Integer
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega informações do relatório
    CEST = Util.ApenasNumeros(Campos(dicTitulos("CEST") - i))
    vCEST = RegrasFiscais.Geral.ValidarCEST(CEST)
    tCEST = VBA.Len(CEST) > 0 And VBA.Len(CEST) < 7
    
    'Verifica validade do CEST
    If Not vCEST And CEST <> "" Then
        INCONSISTENCIA = "O valor informado no campo CEST está inválido"
        SUGESTAO = "Apagar valor do CEST informado no campo CEST"

    'Verifica se o tamanho do CEST está correto
    ElseIf tCEST Then
        INCONSISTENCIA = "O CEST precisa ter 7 dígitos"
        SUGESTAO = "Adicionar zeros a esquerda do CEST"
            
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function VerificarCampoCONTRIBUINTE(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Integer
Dim VL_ICMS As Double
Dim vCONTRIBUINTE As Boolean, tCONTRIBUINTE As Boolean
Dim CONTRIBUINTE As String, TIPO_PART$, COD_PART$, CFOP$, INCONSISTENCIA$, SUGESTAO$, CST_ICMS$
Dim vCompraUsoConsumoSemST As Boolean, vCompraUsoConsumoComST As Boolean, vCompraAtivoImobilizadoSemST As Boolean, _
    vCompraAtivoImobilizadoComST As Boolean, vCompraCombustivelRevenda As Boolean, vCompraCombustivelConsumo As Boolean, _
    vCompraIndustrializacao As Boolean, vCompraRevendaComST As Boolean, vCompraRevenda As Boolean
    
    'Carrega informações do relatório
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    CST_ICMS = Util.RemoverAspaSimples(Campos(dicTitulos("CST_ICMS") - i))
    VL_ICMS = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS") - i), True, 2)
    CONTRIBUINTE = Campos(dicTitulos("CONTRIBUINTE") - i)
    TIPO_PART = Campos(dicTitulos("TIPO_PART") - i)
    COD_PART = Campos(dicTitulos("COD_PART") - i)
    CFOP = Campos(dicTitulos("CFOP") - i)
    
    'Gera verificações sobre os dados do relatório
    vCompraIndustrializacao = ValidacoesCFOP.ValidarCFOPCompraIndustrializacao(CFOP)
    vCompraUsoConsumoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP)
    vCompraUsoConsumoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP)
    vCompraAtivoImobilizadoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP)
    vCompraAtivoImobilizadoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP)
    vCompraCombustivelRevenda = ValidacoesCFOP.ValidarCFOPCompraCombustiveisRevenda(CFOP)
    vCompraCombustivelConsumo = ValidacoesCFOP.ValidarCFOPCompraCombustiveisConsumo(CFOP)
    vCompraRevendaComST = ValidacoesCFOP.ValidarCFOPCompraRevendaComST(CFOP)
    vCompraRevenda = ValidacoesCFOP.ValidarCFOPCompraRevendaSemST(CFOP)
    
    If CFOP < 4000 And TIPO_PART = "PJ" And CONTRIBUINTE = "NÃO" Then
        
        Select Case True
            
            Case vCompraRevenda And VL_ICMS > 0
                INCONSISTENCIA = "Fornecedor (" & TIPO_PART & ") com status de NÃO CONTRIBUINTE em operação de compra para revenda (CFOP: " & CFOP & "), com aproveitamento de crédito do ICMS"
                SUGESTAO = "Informar ""SIM"" no campo CONTRIBUINTE"
                
            Case vCompraIndustrializacao And VL_ICMS > 0
                INCONSISTENCIA = "Fornecedor (" & TIPO_PART & ") com status de NÃO CONTRIBUINTE em operação de compra para revenda (CFOP: " & CFOP & "), com aproveitamento de crédito do ICMS"
                SUGESTAO = "Informar ""SIM"" no campo CONTRIBUINTE"
                
            Case vCompraUsoConsumoSemST, vCompraUsoConsumoComST, vCompraAtivoImobilizadoSemST, vCompraAtivoImobilizadoComST, _
                vCompraCombustivelRevenda, vCompraCombustivelConsumo, vCompraRevendaComST
                INCONSISTENCIA = "Fornecedor (" & TIPO_PART & ") com status de NÃO Contribuinte em operação de compra (" & CFOP & ")"
                SUGESTAO = "Informar ""SIM"" no campo CONTRIBUINTE"
                
        End Select
        
    ElseIf CFOP < 4000 And TIPO_PART = "PJ" And CONTRIBUINTE = "SIM" Then
        
        Select Case True
            
            Case vCompraRevenda And CST_ICMS Like "#00" And VL_ICMS = 0
                INCONSISTENCIA = "Fornecedor (" & TIPO_PART & ") com status de CONTRIBUINTE em operação de compra para revenda (CFOP: " & CFOP & "), sem aproveitamento de crédito do ICMS e CST_ICMS (" & CST_ICMS & ")"
                SUGESTAO = "Informar ""NÃO"" no campo CONTRIBUINTE"
                
            Case vCompraIndustrializacao And Not CFOP Like "#4##" And CST_ICMS Like "#00" And VL_ICMS = 0
                INCONSISTENCIA = "Fornecedor (" & TIPO_PART & ") com status de CONTRIBUINTE em operação de compra para industrialização (CFOP: " & CFOP & "), sem aproveitamento de crédito do ICMS e CST_ICMS (" & CST_ICMS & ")"
                SUGESTAO = "Informar ""NÃO"" no campo CONTRIBUINTE"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_VL_ICMS(Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Integer
Dim VL_ICMS As Double, VL_BC_ICMS#, ALIQ_ICMS#, VL_ICMS_CALC#, DIFERENCA#
Dim CFOP As String, CST_ICMS$, INCONSISTENCIA, SUGESTAO$
Dim vCompraUsoConsumoSemST As Boolean, vCompraUsoConsumoComST As Boolean, vCompraAtivoImobilizadoSemST As Boolean, _
    vCompraAtivoImobilizadoComST As Boolean, vCompraCombustivelRevenda As Boolean, vCompraCombustivelConsumo As Boolean, _
    vCompraRevendaComST As Boolean, vCompraIndustrializacao As Boolean
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega informações do relatório
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_ICMS = Util.ApenasNumeros(Campos(dicTitulos("CST_ICMS") - i))
    VL_BC_ICMS = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_ICMS") - i))
    ALIQ_ICMS = fnExcel.ConverterValores(Campos(dicTitulos("ALIQ_ICMS") - i))
    VL_ICMS = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS") - i))
    VL_ICMS_CALC = fnExcel.ConverterValores(VL_BC_ICMS * ALIQ_ICMS, True, 2)
    
    DIFERENCA = VBA.Round(VBA.Abs(VL_ICMS_CALC - VL_ICMS), 2)
    
    'Gera verificações sobre os dados do relatório
    vCompraIndustrializacao = ValidacoesCFOP.ValidarCFOPCompraIndustrializacao(CFOP)
    vCompraUsoConsumoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP)
    vCompraUsoConsumoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP)
    vCompraAtivoImobilizadoSemST = ValidacoesCFOP.ValidarCFOPCompraAtivoImobilizadoSemST(CFOP)
    vCompraAtivoImobilizadoComST = ValidacoesCFOP.ValidarCFOPCompraAtivoImobilizadoComST(CFOP)
    vCompraCombustivelRevenda = ValidacoesCFOP.ValidarCFOPCompraCombustiveisRevenda(CFOP)
    vCompraCombustivelConsumo = ValidacoesCFOP.ValidarCFOPCompraCombustiveisConsumo(CFOP)
    vCompraRevendaComST = ValidacoesCFOP.ValidarCFOPCompraRevendaComST(CFOP)
    
    'Verifica operações com apropriação indevida do valor do ICMS
    If VL_ICMS > 0 And CFOP < 4000 Then
        
        Select Case True
            
            'Identifica apropriação indevida do ICMS em operação de aquisição para uso e consumo
            Case vCompraUsoConsumoSemST Or vCompraUsoConsumoComST
                INCONSISTENCIA = "Apropriação indevida do ICMS em operação de aquisição para uso e consumo"
                SUGESTAO = "Zerar campos do ICMS"
                
            'Identifica apropriação indevida do ICMS em operação de aquisição para o ativo imobilizado
            Case vCompraAtivoImobilizadoSemST Or vCompraAtivoImobilizadoComST
                INCONSISTENCIA = "Apropriação indevida do ICMS em operação de aquisição para o ativo imobilizado"
                SUGESTAO = "Zerar campos do ICMS"
                
            'Identifica apropriação indevida do ICMS em operação de aquisição de combustíveis e lubrificantes
            Case vCompraCombustivelRevenda Or vCompraCombustivelConsumo
                INCONSISTENCIA = "Apropriação indevida do ICMS em operação de aquisição de combustíveis e lubrificantes"
                SUGESTAO = "Zerar campos do ICMS"
                
            'Identifica apropriação indevida do ICMS em operação de compra para revenda com ST
            Case vCompraRevendaComST
                INCONSISTENCIA = "Apropriação indevida do ICMS em operação de compra para revenda com ST"
                SUGESTAO = "Zerar campos do ICMS"
                
            'Identifica apropriação indevida do ICMS em operação de entrada com ST
            Case CST_ICMS Like "*60"
                INCONSISTENCIA = "Apropriação indevida do ICMS em operação de entrada com ST"
                SUGESTAO = "Zerar campos do ICMS"
                
            'Identifica apropriação indevida do ICMS em operação de entrada com ST
            Case CST_ICMS Like "*10" And Not vCompraIndustrializacao
                INCONSISTENCIA = "Apropriação indevida do ICMS em operação de entrada com ST"
                SUGESTAO = "Zerar campos do ICMS"
                
            'Valor do ICMS calculado é maior que o ICMS destacado
            Case VL_ICMS_CALC > VL_ICMS
                INCONSISTENCIA = "Cálculo do campo VL_ICMS maior que o destacado [Diferença: R$ " & DIFERENCA & "]"
                SUGESTAO = "Recalcular o campo VL_ICMS"
                
            'Valor do ICMS calculado é menor que o ICMS destacado
            Case VL_ICMS_CALC < VL_ICMS
                INCONSISTENCIA = "Cálculo do campo VL_ICMS menor que o destacado [Diferença: R$ " & DIFERENCA & "]"
                SUGESTAO = "Recalcular o campo VL_ICMS"
                
        End Select
        
    ElseIf VL_ICMS > 0 And CFOP > 4000 Then
        
        Select Case True
            
            'Identifica apropriação indevida do ICMS em operação de saída com ST
            Case CST_ICMS Like "*60"
                INCONSISTENCIA = "Destaque indevido do ICMS em operação de saída com ST"
                SUGESTAO = "Zerar campos do ICMS"
                
            'Valor do ICMS calculado é maior que o ICMS destacado
            Case VL_ICMS_CALC > VL_ICMS
                INCONSISTENCIA = "Cálculo do campo VL_ICMS maior que o destacado [Diferença: R$ " & DIFERENCA & "]" ' (Calculado: R$ " & VL_ICMS_CALC & ") > Destacado: R$ " & VL_ICMS & ")"
                SUGESTAO = "Recalcular o campo VL_ICMS"
                
            'Valor do ICMS calculado é menor que o ICMS destacado
            Case VL_ICMS_CALC < VL_ICMS
                INCONSISTENCIA = "Cálculo do campo VL_ICMS menor que o destacado [Diferença: R$ " & DIFERENCA & "]" ' (Calculado: R$ " & VL_ICMS_CALC & ") < Destacado: R$ " & VL_ICMS & ")" '"O valor calculado do campo VL_ICMS (" & VL_ICMS_CALC & ") está menor que o campo VL_ICMS (R$ " & VL_ICMS & ")"
                SUGESTAO = "Recalcular o campo VL_ICMS"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_ALIQ_ICMS(Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Integer
Dim ALIQ_ICMS As Double, VL_ICMS#
Dim CFOP As String, CST_ICMS$, INCONSISTENCIA, SUGESTAO$, ALIQ_ICMS_FORM$, VL_ICMS_FORM$
Dim vCompraUsoConsumoSemST As Boolean, vCompraUsoConsumoComST As Boolean, vCompraAtivoImobilizadoSemST As Boolean, _
    vCompraAtivoImobilizadoComST As Boolean, vCompraCombustivelRevenda As Boolean, vCompraCombustivelConsumo As Boolean

    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega informações do relatório
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_ICMS = Util.ApenasNumeros(Campos(dicTitulos("CST_ICMS") - i))
    ALIQ_ICMS = fnExcel.FormatarValores(Campos(dicTitulos("ALIQ_ICMS") - i))
    VL_ICMS = fnExcel.FormatarValores(Campos(dicTitulos("VL_ICMS") - i))
    ALIQ_ICMS_FORM = VBA.Format(ALIQ_ICMS, "#0.00%")
    VL_ICMS_FORM = VBA.Format(VL_ICMS, "R$ #,#0.00")
    
    'Gera verificações sobre os dados do relatório
    vCompraUsoConsumoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP)
    vCompraUsoConsumoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP)
    vCompraAtivoImobilizadoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP)
    vCompraAtivoImobilizadoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP)
    vCompraCombustivelRevenda = ValidacoesCFOP.ValidarCFOPCompraCombustiveisRevenda(CFOP)
    vCompraCombustivelConsumo = ValidacoesCFOP.ValidarCFOPCompraCombustiveisConsumo(CFOP)
    
    'Verifica operações com Alíquota 0
    If ALIQ_ICMS = 0 Then
                
        Select Case True
        
            'Identifica apropriação indevida do ICMS em operação de aquisição para uso e consumo
            Case VL_ICMS > 0
                INCONSISTENCIA = "Alíquota zerada (ALIQ_ICMS = 0) com valor destacado de ICMS (VL_ICMS = " & VL_ICMS_FORM & ")"
                SUGESTAO = "Informe uma alíquota de ICMS compatível com a operação no campo ALIQ_ICMS"

        End Select
        
    ElseIf ALIQ_ICMS > 0 Then
    
        Select Case True
        
            'Identifica apropriação indevida do ICMS em operação de aquisição para uso e consumo
            Case vCompraUsoConsumoComST Or vCompraUsoConsumoSemST
                INCONSISTENCIA = "Alíquota maior que zero (ALIQ_ICMS = " & ALIQ_ICMS_FORM & ") em operação de compra para uso e consumo (CFOP = " & CFOP & ")"
                SUGESTAO = "Zerar Alíquota do ICMS"

        End Select
    
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_VL_ICMS_ST(Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Integer
Dim CFOP As String, INCONSISTENCIA$, SUGESTAO$, CST_ICMS$
Dim VL_ICMS_ST As Double
Dim vCompraUsoConsumoSemST As Boolean, vCompraUsoConsumoComST As Boolean, vCompraAtivoImobilizadoSemST As Boolean, _
    vCompraAtivoImobilizadoComST As Boolean, vCompraRevendaSemST As Boolean, vCompraRevendaComST As Boolean, _
    vCompraCombustiveisRevenda As Boolean, vCompraCombustiveisConsumo As Boolean

    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega informações do relatório
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_ICMS = Util.ApenasNumeros(Campos(dicTitulos("CST_ICMS") - i))
    VL_ICMS_ST = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS_ST") - i))

    'Gera verificações sobre os dados do relatório
    vCompraRevendaSemST = ValidacoesCFOP.ValidarCFOPCompraRevendaSemST(CFOP)
    vCompraRevendaComST = ValidacoesCFOP.ValidarCFOPCompraRevendaComST(CFOP)
    vCompraUsoConsumoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP)
    vCompraUsoConsumoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP)
    vCompraAtivoImobilizadoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP)
    vCompraAtivoImobilizadoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP)
    vCompraCombustiveisRevenda = ValidacoesCFOP.ValidarCFOPCompraCombustiveisRevenda(CFOP)
    vCompraCombustiveisConsumo = ValidacoesCFOP.ValidarCFOPCompraCombustiveisConsumo(CFOP)
    
    'Verifica operações em que o valor do ICMS_ST é maior que zero
    If VL_ICMS_ST > 0 Then
    
        Select Case True
            
            Case CFOP Like "*910"
                INCONSISTENCIA = "O valor do campo VL_ICMS_ST deve ser somado ao campo VL_ITEM para operações de entrada em bonificação"
                SUGESTAO = "Somar valor do campo VL_ICMS_ST ao campo VL_ITEM"
                
            'Identifica lançamento do ICMS_ST em campo próprio em operações de compra para revenda
            Case vCompraRevendaSemST Or vCompraRevendaComST
                INCONSISTENCIA = "O valor do campo VL_ICMS_ST deve ser somado ao campo VL_ITEM para operações de compra para revenda"
                SUGESTAO = "Somar valor do campo VL_ICMS_ST ao campo VL_ITEM"
            
            'Identifica aproveitamento de crédito do ICMS em operações de compra de combustíveis e lubrificantes
            Case vCompraCombustiveisRevenda Or vCompraCombustiveisConsumo
                INCONSISTENCIA = "O valor do campo VL_ICMS_ST deve ser somado ao campo VL_ITEM em operações de compra de combustíveis e lubrificantes"
                SUGESTAO = "Somar valor do campo VL_ICMS_ST ao campo VL_ITEM"
            
            'Identifica aproveitamento de créditos do ICMS em operações de aquisição para uso e consumo
            Case vCompraUsoConsumoSemST Or vCompraUsoConsumoComST
                INCONSISTENCIA = "O valor do campo VL_ICMS_ST deve ser somado ao campo VL_ITEM para aquisições de uso e consumo"
                SUGESTAO = "Somar valor do campo VL_ICMS_ST ao campo VL_ITEM"
            
            'Identifica aproveitamento de créditos do ICMS em operações de aquisição para ativo imobilizado
            Case vCompraAtivoImobilizadoSemST Or vCompraAtivoImobilizadoComST
                INCONSISTENCIA = "O valor do campo VL_ICMS_ST deve ser somado ao campo VL_ITEM para aquisições de ativo imobilizado"
                SUGESTAO = "Somar valor do campo VL_ICMS_ST ao campo VL_ITEM"
            
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_CST_ICMS(Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Integer
Dim VL_ICMS As Double, VL_BC_ICMS#, VL_ITEM#, VL_DESC#, VL_ICMS_ST#
Dim CST_ICMS As String, CFOP$, CST_tbA$, CST_tbB$, INCONSISTENCIA$, SUGESTAO$
Dim vCompraUsoConsumoSemST As Boolean, vCompraUsoConsumoComST As Boolean, vCompraAtivoImobilizadoSemST As Boolean, _
    vCompraAtivoImobilizadoComST As Boolean, vCSOSN As Boolean, vCompraCombustivelIndustrializacao As Boolean, _
    vCompraCombustivelRevenda As Boolean, vCompraCombustivelConsumo As Boolean, vCompraRevendaSemST As Boolean, _
    vCompraRevendaComST As Boolean, vCFOPEntrada As Boolean, vCFOPSaida As Boolean, vEntradaBonificacao As Boolean, _
    vImportacao As Boolean
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega informações do relatório
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_ICMS = Util.ApenasNumeros(Campos(dicTitulos("CST_ICMS") - i))
    VL_ICMS = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS") - i))
    VL_ICMS_ST = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS_ST") - i))
    VL_BC_ICMS = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_ICMS") - i))
    VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM") - i))
    VL_DESC = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESC") - i))
    
    'Gera informações derivadas dos dados principais
    CST_tbA = VBA.Left(CST_ICMS, 1)
    CST_tbB = VBA.Right(CST_ICMS, 2)
    
    'Gera verificações sobre os dados do relatório
    vCompraRevendaSemST = ValidacoesCFOP.ValidarCFOPCompraRevendaSemST(CFOP)
    vCompraRevendaComST = ValidacoesCFOP.ValidarCFOPCompraRevendaComST(CFOP)
    vCompraUsoConsumoSemST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoSemST(CFOP)
    vCompraUsoConsumoComST = ValidacoesCFOP.ValidarCFOPCompraUsoConsumoComST(CFOP)
    vCompraAtivoImobilizadoSemST = ValidacoesCFOP.ValidarCFOPCompraAtivoImobilizadoSemST(CFOP)
    vCompraAtivoImobilizadoComST = ValidacoesCFOP.ValidarCFOPCompraAtivoImobilizadoComST(CFOP)
    vCompraCombustivelIndustrializacao = ValidacoesCFOP.ValidarCFOPCompraCombustiveisIndustrializacao(CFOP)
    vCompraCombustivelRevenda = ValidacoesCFOP.ValidarCFOPCompraCombustiveisRevenda(CFOP)
    vCompraCombustivelConsumo = ValidacoesCFOP.ValidarCFOPCompraCombustiveisConsumo(CFOP)
    vEntradaBonificacao = ValidacoesCFOP.ValidarCFOPEntradaBonificacao(CFOP)
    vCSOSN = VBA.Len(CST_ICMS) = 4
    vImportacao = ValidacoesCFOP.ValidarCFOPImportacao(CFOP)
    
    'Identifica operações sem informação do CST/ICMS
    If CST_ICMS = "" Then
        INCONSISTENCIA = "CST_ICMS não foi informado"
        SUGESTAO = "informar um valor válido para o campo CST_ICMS"
    
    'Identifica operações informadas com CSOSN ao invés do CST/ICMS
    ElseIf vCSOSN Then
        
        'TODO: Criar regra para identificar CST/CSOSN com menos de 3 dígitos
        Select Case True
                
            'Identifica CSOSN de operações com permissão de crédito
            Case CST_ICMS Like "#101"
                INCONSISTENCIA = "CSOSN indica operação com permissão de crédito do ICMS"
                SUGESTAO = "Informar o CST 90 da tabela B para o campo CST_ICMS"
                
            'Identifica CSOSN de operações sem permissão de crédito
            Case CST_ICMS Like "#102"
                INCONSISTENCIA = "CSOSN indica operação sem permissão de crédito do ICMS"
                SUGESTAO = "Informar o CST 90 da tabela B para o campo CST_ICMS"
                
            'Identifica CSOSN de operações com isenção do ICMS
            Case CST_ICMS Like "#103"
                INCONSISTENCIA = "CSOSN indica operação isenta"
                SUGESTAO = "Informar o CST 40 da tabela B para o campo CST_ICMS"
            
            'Identifica CSOSN de operações sujeitas a cobrança da Substituição tributária do ICMS
            Case CST_ICMS Like "#20#"
                INCONSISTENCIA = "CSOSN indica operação com cobrança da Substituição Tributária do ICMS"
                SUGESTAO = "Informar o CST 60 da tabela B para o campo CST_ICMS"
                
            'Identifica CSOSN de operações com imunidade
            Case CST_ICMS Like "#300"
                INCONSISTENCIA = "CSOSN indica operação com imune"
                SUGESTAO = "Informar o CST 41 da tabela B para o campo CST_ICMS"
            
            'Identifica CSOSN de operações não-tributadas
            Case CST_ICMS Like "#400"
                INCONSISTENCIA = "CSOSN indica operação não-tributada"
                SUGESTAO = "Informar o CST 41 da tabela B para o campo CST_ICMS"
                
            'Identifica CSOSN de operações sujeitas a Substituição tributária do ICMS
            Case CST_ICMS Like "#500"
                INCONSISTENCIA = "CSOSN indica operação com ICMS cobrado anteriormente por substituição"
                SUGESTAO = "Informar o CST 60 da tabela B para o campo CST_ICMS"
                
            'Identifica CSOSN de operações com tributação do ICMS
            Case CST_ICMS Like "#900"
                INCONSISTENCIA = "CSOSN indica outras operações"
                SUGESTAO = "Informar o CST 90 da tabela B para o campo CST_ICMS"
                
            'Em qualquer outra situação aplica a regra abaixo
            Case Else
                INCONSISTENCIA = "CSOSN informado indevidamente no campo CST_ICMS"
                SUGESTAO = "Informe um CST/ICMS válido para o campo CST_ICMS"
                
        End Select
    
    'Verifica operações com mudança no dígito de origem (Tabela A) igual a 1
    ElseIf CST_tbA = "1" Then
        
        Select Case True
            
            'Identifica operações com dígito de origem igual a 1 que não sejam de importação
            Case Not vImportacao
                INCONSISTENCIA = "O dígito de origem do campo CST_ICMS deve ser igual a 2"
                SUGESTAO = "Mudar o dígito de origem do campo CST_ICMS para 2"
            
        End Select
        
    'Verifica operações com mudança no dígito de origem (Tabela A) igual a 6
    ElseIf CST_tbA = "6" Then
        
        Select Case True
            
            'Identifica operações com dígito de origem igual a 6 que não sejam de importação
            Case Not vImportacao
                INCONSISTENCIA = "O dígito de origem do campo CST_ICMS deve ser igual a 7"
                SUGESTAO = "Mudar o dígito de origem do campo CST_ICMS para 7"
        
        End Select
        
    'Identifica operações com CST/ICMS 20 que não possuem redução de base cálculo do ICMS
    ElseIf CST_tbB = "00" Then
        
        Select Case True
            
            Case VL_BC_ICMS = 0 And VL_ICMS = 0
                INCONSISTENCIA = "CST/ICMS inconsistente (CST_ICMS = *" & CST_tbB & ") com operação tributada integralmente"
                SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 90"
                
        End Select
        
    'Identifica operações com CST/ICMS 20 que não possuem redução de base cálculo do ICMS
    ElseIf CST_tbB = "20" Then
        
        Select Case True
            
            Case vCompraRevendaComST
                INCONSISTENCIA = "CST/ICMS inconsistente (CST_ICMS = *" & CST_tbB & ") com operação sujeita a ST"
                SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 60"
                
            'Identifica operações com CST/ICMS 20, que são tributadas integralmente
            Case VL_BC_ICMS >= VL_ITEM - VL_DESC
                INCONSISTENCIA = "CST/ICMS inconsistente (CST_ICMS = *" & CST_tbB & ") com operação tributada integralmente"
                SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 00"
                
            Case VL_ITEM > 0 And VL_BC_ICMS = 0
                INCONSISTENCIA = "CST/ICMS inconsistente (CST_ICMS = *" & CST_tbB & ") com operação sem tributação do ICMS"
                SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 90"
                
        End Select
        
    'Identifica operações com CST/ICMS 40 que não possuem isenção
    ElseIf CST_tbB = "40" Then
        
        Select Case True
            
            'Identifica operações com CST/ICMS 40, que são tributadas integralmente
            Case VL_ICMS > 0
                If VL_BC_ICMS >= VL_ITEM - VL_DESC Then
                    INCONSISTENCIA = "CST/ICMS inconsistente (CST_ICMS = *" & CST_tbB & ") com operação isenta"
                    SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 00"
                    
                ElseIf VL_BC_ICMS < VL_ITEM - VL_DESC Then
                    INCONSISTENCIA = "CST/ICMS inconsistente (CST_ICMS = *" & CST_tbB & ") com operação isenta"
                    SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 20"
                    
                End If
            
        End Select
        
    'Identifica operações com CST/ICMS 40 que não possuem não tributação
    ElseIf CST_tbB = "41" Then
        
        Select Case True
            
            'Identifica operações com CST/ICMS 41, que são tributadas integralmente
            Case VL_BC_ICMS >= VL_ITEM - VL_DESC
                INCONSISTENCIA = "CST/ICMS inconsistente (CST_ICMS = *" & CST_tbB & ") com operação não tributada"
                SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 00"

        End Select
        
    'Identifica operações de compra para revenda sem ST
    ElseIf vCompraRevendaSemST Then
        
        'Verifica operações sem aproveitamento do crédito do ICMS
        If VL_ICMS = 0 Then
            
            Select Case True
                
                'Identifica operações com CST/ICMS 00 sem aproveitamento do crédito de ICMS
                Case CST_tbB = "00"
                    INCONSISTENCIA = "Operação tributada integralmente (CST_ICMS = *" & CST_tbB & ") sem aproveitamento de crédito (VL_ICMS = 0)"
                    SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 90"
                
                'Identifica operações com CST/ICMS 20 sem aproveitamento do crédito de ICMS
                Case CST_tbB = "20"
                    INCONSISTENCIA = "Operação com redução de base (CST_ICMS = *" & CST_tbB & ") sem aproveitamento de crédito (VL_ICMS = 0)"
                    SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 90"
                    
            End Select

        End If
                
    'Identifica operações de compra para revenda com ST
    ElseIf vCompraRevendaComST Then
    
        Select Case True
            
            'Identifica CST/ICMS inconsistente em operações compra para revenda sujeita a ST
            Case CST_tbB <> "60"
                INCONSISTENCIA = "CST/ICMS inconsistente (CST_ICMS = *" & CST_tbB & ") com operação de compra para revenda sujeita a ST"
                SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 60"
                                                                                     
        End Select
            
    'Identifica CST/ICMS inconsistente em operações compra de combustíveis e lubrificantes
    ElseIf (vCompraCombustivelRevenda Or vCompraCombustivelConsumo) Then
        
        Select Case True
                        
            Case CST_tbB <> "60" And VL_ICMS = 0
                INCONSISTENCIA = "CST/ICMS inconsistente (CST_ICMS = *" & CST_tbB & ") com operação de compra de combustíveis e lubrificantes"
                SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 60"
                
        End Select
    
    'Identifica CST/ICMS inconsistente em operações compra para uso e consumo sem ST
    ElseIf vCompraUsoConsumoSemST Then
        
        Select Case True
                        
            Case CST_tbB <> "90" And VL_ICMS = 0
                INCONSISTENCIA = "CST/ICMS inconsistente (CST_ICMS = *" & CST_tbB & ") com operação de compra para uso e consumo sem ST"
                SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 90"
                
        End Select
        
    'Identifica CST/ICMS inconsistente em operações de compra para uso e consumo com ST
    ElseIf vCompraUsoConsumoComST Then
        
        Select Case True
            
            Case CST_tbB <> "60" And VL_ICMS = 0
                INCONSISTENCIA = "CST/ICMS inconsistente (CST_ICMS = *" & CST_tbB & ") com operação de compra para uso e consumo com ST"
                SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 60"
                
        End Select

    'Identifica CST/ICMS inconsistente em operações de compra para o ativo imobilizado sem ST
    ElseIf vCompraAtivoImobilizadoSemST Then
        
        Select Case True
            
            Case CST_tbB <> "90" And VL_ICMS = 0
                INCONSISTENCIA = "CST/ICMS inconsistente (CST_ICMS = *" & CST_tbB & ") com operação de compra para o ativo imobilizado sem ST"
                SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 90"
                
        End Select

    
    'Identifica CST/ICMS inconsistente em operações de compra para o ativo imobilizado com ST
    ElseIf vCompraAtivoImobilizadoComST Then
                
        Select Case True
                        
            Case CST_tbB <> "60" And VL_ICMS = 0
                INCONSISTENCIA = "CST/ICMS inconsistente (CST_ICMS = *" & CST_tbB & ") com operação de compra para o ativo imobilizado com ST"
                SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 60"
                
        End Select
    
    ElseIf VL_ICMS_ST = 0 And CFOP < 4000 Then
    
        Select Case True
            
            Case CST_ICMS Like "*10"
                INCONSISTENCIA = "CST_ICMS indica operação de entrada com cobrança de ST, porém não há destaque do imposto"
                SUGESTAO = "Alterar dígitos da Tabela B do CST/ICMS para 60"
                
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

