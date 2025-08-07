Attribute VB_Name = "clsRegrasApuracaoPISCOFINS"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private ValidacoesCFOP As New clsRegrasFiscaisCFOP
Private Const RegistrosPIS As String = "C181,D201"
Private Const RegistrosCOFINS As String = "C185,D205"
Private ListaCFOPsCreditaveis As New ArrayList
Private CamposPISCOFINS As CamposApuracaoPISCOFINS
Private dicTitulos As New Dictionary

Private Type CamposApuracaoPISCOFINS
    
    REG As String
    CFOP As String
    CST_PIS As String
    CST_COFINS As String
    COD_INC_TRIB As String
    TIPO_PART As String
    INCONSISTENCIA As String
    SUGESTAO As String
    
End Type

Private Function CarregarCamposRegistroApuracaoPISCOFINS(ByRef Campos As Variant, ByVal COD_INC_TRIB As String)

Dim i As Byte
    
    If dicTitulos.Count = 0 Then Set dicTitulos = Util.MapearTitulos(assApuracaoPISCOFINS, 3)
    If LBound(Campos) = 0 Then i = 1
    
    With CamposPISCOFINS
        
        .REG = Util.RemoverAspaSimples(Campos(dicTitulos("REG") - i))
        .CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
        .CST_PIS = Util.ApenasNumeros(Campos(dicTitulos("CST_PIS") - i))
        .CST_COFINS = Util.ApenasNumeros(Campos(dicTitulos("CST_COFINS") - i))
        .TIPO_PART = Util.RemoverAspaSimples(Campos(dicTitulos("TIPO_PART") - i))
        .INCONSISTENCIA = Campos(dicTitulos("INCONSISTENCIA") - i)
        .COD_INC_TRIB = COD_INC_TRIB
        
    End With
    
End Function

Private Function ResetarCamposRegistroApuracaoPISCOFINS()
    
    Dim CamposVazios As CamposApuracaoPISCOFINS
    LSet CamposPISCOFINS = CamposVazios
    
End Function

Public Function ValidacoesCampo_CST_PIS_COFINS(ByRef Campos As Variant, ByVal COD_INC_TRIB As String)
    
    Call CarregarCamposRegistroApuracaoPISCOFINS(Campos, COD_INC_TRIB)
    
    With CamposPISCOFINS
        
        Select Case True
            
            Case RegistrosPIS Like "*" & .REG & "*" And .INCONSISTENCIA = ""
                Call ValidacoesCampo_CST_PIS
                
            Case RegistrosCOFINS Like "*" & .REG & "*" And .INCONSISTENCIA = ""
                Call ValidacoesCampo_CST_COFINS
                
            Case Else
                Call ValidacoesCampo_CST_PIS
                If .INCONSISTENCIA = "" Then Call ValidacoesCampo_CST_COFINS
                
        End Select
        
        If .INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
            INCONSISTENCIA:=.INCONSISTENCIA, SUGESTAO:=.SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
            
    End With
    
    Call ResetarCamposRegistroApuracaoPISCOFINS
    
End Function

Private Sub ValidacoesCampo_CST_PIS()
    
    With CamposPISCOFINS
        
        Call VerificarCST_PIS
        If .INCONSISTENCIA <> "" Or .CFOP = "" Then Exit Sub
        
        Select Case True
            
            Case .CFOP < 4000
                Call ValidarCST_PISemOperacaoDeEntrada
                
            Case .CFOP > 4000
                Call ValidarCST_PISemOperacaoDeSaida
                
        End Select
        
    End With
    
End Sub

Private Function VerificarCST_PIS()
    
    With CamposPISCOFINS
        
        Select Case True
            
            Case .CST_PIS = ""
                .INCONSISTENCIA = "CST_PIS não foi informado"
                .SUGESTAO = "Informar um valor válido para o CST_PIS"
                
            Case .CST_PIS Like "00*"
                .INCONSISTENCIA = "O CST_PIS informado está inválido"
                .SUGESTAO = "Informar um CST_PIS válido"
                
            Case .CST_PIS <> .CST_COFINS And Not RegistrosPIS Like "*" & .REG & "*"
                .INCONSISTENCIA = "Os campos CST_PIS e CST_COFINS estão divergentes"
                .SUGESTAO = "Informar o mesmo CST para o PIS e a COFINS"
                
        End Select
        
    End With
    
End Function

Private Function ValidarCST_PISemOperacaoDeEntrada()
    
    If ListaCFOPsCreditaveis.Count = 0 Then Call CarregarCFOPSGeradoresCredito(ListaCFOPsCreditaveis)
        
    With CamposPISCOFINS
        
        Select Case True
            
            Case Not ListaCFOPsCreditaveis.contains(.CFOP) And .CST_PIS Like "5*" Or .CST_PIS Like "6*" And .CFOP < 4000 And (.CFOP Like "#556" Or .CFOP Like "#551")
                .INCONSISTENCIA = "CFOP (" & .CFOP & ") não consta na tabela de operações geradoras de crédito da EFD Contribuições"
                .SUGESTAO = "Informar CST_PIS 98 - Outras Operações de Entrada"
            
            Case Not ListaCFOPsCreditaveis.contains(.CFOP) And .CST_PIS Like "5*" Or .CST_PIS Like "6*"
                .INCONSISTENCIA = "CFOP (" & .CFOP & ") não consta na tabela de operações geradoras de crédito da EFD Contribuições"
                .SUGESTAO = "Informar CST_PIS igual a 70 - Operação de Aquisição sem Direito a Crédito"
                
            Case .CST_PIS < 50
                .INCONSISTENCIA = "CST_PIS de saída informado para CFOP de entrada"
                .SUGESTAO = "Informar um CST_PIS para operação de entrada"
                
            Case .CFOP Like "#910" And Not .CST_PIS Like "98*"
                .INCONSISTENCIA = "CST_PIS (" & .CST_PIS & ") incorreto para operação de entrada em bonificação"
                .SUGESTAO = "Informar CST_PIS 98 - Outras Operações de Entrada"
                
            Case (.CFOP Like "#407" Or .CFOP Like "#556") And Not .CST_PIS Like "98*"
                .INCONSISTENCIA = "CST_PIS (" & .CST_PIS & ") incorreto para operação de aquisição para uso e consumo"
                .SUGESTAO = "Informar CST_PIS 98 - Outras Operações de Entrada"
                
            Case .CST_PIS Like "5*" And .TIPO_PART Like "*PF*"
                .INCONSISTENCIA = "Não deve ser informado CST referente a Operações com Direito a Crédito(50 a 56) para operações cujo participante é pessoa física (TIPO_PART = PF)"
                .SUGESTAO = "Informar CST_PIS igual a 70 - Operação de Aquisição sem Direito a Crédito"
                                
            Case (Not .CST_PIS Like "9*" And Not .CST_PIS Like "7*")
                If .COD_INC_TRIB = "2" Then
                    .SUGESTAO = "Informar CST_PIS igual a 70 - Operação de Aquisição sem Direito a Crédito"
                    .INCONSISTENCIA = "CST_PIS inconsistente com apuração no Regime Cumulativo (Lucro Presumido)"
                End If
                
        End Select
        
    End With
    
End Function

Private Function ValidarCST_PISemOperacaoDeSaida()

Dim fatCFOP As Boolean, devCompCFOP As Boolean
    
    With CamposPISCOFINS
        
        fatCFOP = ValidacoesCFOP.ValidarCFOPFaturamento(.CFOP)
        devCompCFOP = ValidacoesCFOP.ValidarCFOPDevolucaoCompra(.CFOP)
        
        Select Case True
            
            Case .CST_PIS > 49 And .CST_PIS < 99
                .INCONSISTENCIA = "CST_PIS de entrada informado para CFOP de saída"
                .SUGESTAO = "Informar um CST_PIS para operação de saída"
                
            Case Not fatCFOP And .CST_PIS < 7
                .INCONSISTENCIA = "CST_PIS tributável com CFOP de operação não tributável"
                .SUGESTAO = "Alterar CST_PIS"
                
            Case devCompCFOP And .CST_PIS <> 49
                .INCONSISTENCIA = "CST_PIS diferente de 49 informado em operação de devolução"
                .SUGESTAO = "Alterar CST_PIS para 49"
                
            Case fatCFOP And .CST_PIS Like "*7"
                .INCONSISTENCIA = "Operação de Venda com CST_PIS indicando operação isenta"
                .SUGESTAO = "Verificar se o CST_PIS está correto"
                
            Case fatCFOP And .CST_PIS Like "*8"
                .INCONSISTENCIA = "Operação de Venda com CST_PIS indicando operação sem incidência"
                .SUGESTAO = "Verificar se o CST_PIS está correto"
                
            Case fatCFOP And .CST_PIS > 9
                .INCONSISTENCIA = "CST_PIS incorreto para operação de Venda"
                .SUGESTAO = "Informar CST_PIS correto"
                
            Case .CFOP Like "#910" And Not .CST_PIS Like "49*"
                .INCONSISTENCIA = "CST_PIS (" & .CST_PIS & ") incorreto para operação de saída em bonificação"
                .SUGESTAO = "Informar CST_PIS 49 - Outras Operações de Saída"
                
        End Select
        
    End With
    
End Function

Private Sub ValidacoesCampo_CST_COFINS()
    
    With CamposPISCOFINS
        
        Call VerificarCST_COFINS
        If .INCONSISTENCIA <> "" Or .CFOP = "" Then Exit Sub
        
        Select Case True
            
            Case .CFOP < 4000
                Call ValidarCST_COFINSemOperacaoDeEntrada
                
            Case .CFOP > 4000
                Call ValidarCST_COFINSemOperacaoDeSaida
                
        End Select
        
    End With
    
End Sub

Private Function VerificarCST_COFINS()
    
    With CamposPISCOFINS
        
        Select Case True
            
            Case .CST_COFINS = ""
                .INCONSISTENCIA = "CST_COFINS não foi informado"
                .SUGESTAO = "Informar um valor válido para o CST_COFINS"
                
            Case .CST_COFINS Like "00*"
                .INCONSISTENCIA = "O CST_COFINS informado está inválido"
                .SUGESTAO = "Informar um CST_COFINS válido"
                
            Case .CST_PIS <> .CST_COFINS And Not RegistrosCOFINS Like "*" & .REG & "*"
                .INCONSISTENCIA = "Os campos CST_PIS e CST_COFINS estão divergentes"
                .SUGESTAO = "Informar o mesmo CST para o PIS e a COFINS"
                
        End Select
        
    End With
    
End Function

Private Function ValidarCST_COFINSemOperacaoDeEntrada()
    
    If ListaCFOPsCreditaveis.Count = 0 Then Call CarregarCFOPSGeradoresCredito(ListaCFOPsCreditaveis)
    
    With CamposPISCOFINS
        
        Select Case True
            
            Case Not ListaCFOPsCreditaveis.contains(.CFOP) And .CST_COFINS Like "5*" Or .CST_COFINS Like "6*" And .CFOP < 4000 And (.CFOP Like "#556" Or .CFOP Like "#551")
                .INCONSISTENCIA = "CFOP (" & .CFOP & ") não consta na tabela de operações geradoras de crédito da EFD Contribuições"
                .SUGESTAO = "Informar CST_COFINS 98 - Outras Operações de Entrada"
            
            Case Not ListaCFOPsCreditaveis.contains(.CFOP) And .CST_COFINS Like "5*" Or .CST_COFINS Like "6*"
                .INCONSISTENCIA = "CFOP (" & .CFOP & ") não consta na tabela de operações geradoras de crédito da EFD Contribuições"
                .SUGESTAO = "Informar CST_COFINS igual a 70 - Operação de Aquisição sem Direito a Crédito"
                
            Case .CST_COFINS < 50
                .INCONSISTENCIA = "CST_COFINS de saída informado para CFOP de entrada"
                .SUGESTAO = "Informar um CST_COFINS para operação de entrada"
                
            Case .CFOP Like "#910" And Not .CST_COFINS Like "98*"
                .INCONSISTENCIA = "CST_COFINS (" & .CST_COFINS & ") incorreto para operação de entrada em bonificação"
                .SUGESTAO = "Informar CST_COFINS 98 - Outras Operações de Entrada"
                
            Case (.CFOP Like "#407" Or .CFOP Like "#556") And Not .CST_COFINS Like "98*"
                .INCONSISTENCIA = "CST_COFINS (" & .CST_COFINS & ") incorreto para operação de aquisição para uso e consumo"
                .SUGESTAO = "Informar CST_COFINS 98 - Outras Operações de Entrada"
                
            Case .CST_COFINS Like "5*" And .TIPO_PART Like "*PF*"
                .INCONSISTENCIA = "Não deve ser informado CST referente a Operações com Direito a Crédito(50 a 56) para operações cujo participante é pessoa física (TIPO_PART = PF)"
                .SUGESTAO = "Informar CST_COFINS igual a 70 - Operação de Aquisição sem Direito a Crédito"
                
            Case (Not .CST_COFINS Like "9*" And Not .CST_COFINS Like "7*")
                If .COD_INC_TRIB = "2" Then
                    .SUGESTAO = "Informar CST_COFINS igual a 70 - Operação de Aquisição sem Direito a Crédito"
                    .INCONSISTENCIA = "CST_COFINS inconsistente com apuração no Regime Cumulativo (Lucro Presumido)"
                End If
                                
        End Select
        
    End With
    
End Function

Private Function ValidarCST_COFINSemOperacaoDeSaida()

Dim fatCFOP As Boolean, devCompCFOP As Boolean
    
    With CamposPISCOFINS
        
        fatCFOP = ValidacoesCFOP.ValidarCFOPFaturamento(.CFOP)
        devCompCFOP = ValidacoesCFOP.ValidarCFOPDevolucaoCompra(.CFOP)
        
        Select Case True
            
            Case .CST_COFINS > 49 And .CST_COFINS < 99
                .INCONSISTENCIA = "CST_COFINS de entrada informado para CFOP de saída"
                .SUGESTAO = "Informar um CST_COFINS para operação de saída"
                
            Case Not fatCFOP And .CST_COFINS < 7
                .INCONSISTENCIA = "CST_COFINS tributável com CFOP de operação não tributável"
                .SUGESTAO = "Alterar CST_COFINS"
                
            Case devCompCFOP And .CST_COFINS <> 49
                .INCONSISTENCIA = "CST_COFINS diferente de 49 informado em operação de devolução"
                .SUGESTAO = "Alterar CST_COFINS para 49"
                
            Case fatCFOP And .CST_COFINS Like "*7"
                .INCONSISTENCIA = "Operação de Venda com CST_COFINS indicando operação isenta"
                .SUGESTAO = "Verificar se o CST_COFINS está correto"
                
            Case fatCFOP And .CST_COFINS Like "*8"
                .INCONSISTENCIA = "Operação de Venda com CST_COFINS indicando operação sem incidência"
                .SUGESTAO = "Verificar se o CST_COFINS está correto"
                
            Case fatCFOP And .CST_COFINS > 9
                .INCONSISTENCIA = "CST_COFINS incorreto para operação de Venda"
                .SUGESTAO = "Informar CST_COFINS correto"
                
            Case .CFOP Like "#910" And Not .CST_COFINS Like "49*"
                .INCONSISTENCIA = "CST_COFINS (" & .CST_COFINS & ") incorreto para operação de saída em bonificação"
                .SUGESTAO = "Informar CST_COFINS 49 - Outras Operações de Saída"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_PIS_COFINS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
'TODO: Refatorar código para a nova estrutura usando a estrutura CamposPISCOFINS As CamposApuracaoPISCOFINS
Dim i As Byte
Dim REG As String, INCONSISTENCIA$
    
    If LBound(Campos) = 0 Then i = 1
    
    'Identificação do registro
    REG = Campos(dicTitulos("REG") - i)
    INCONSISTENCIA = Campos(dicTitulos("INCONSISTENCIA") - i)
    
    Select Case True
        
        Case RegistrosPIS Like "*" & REG & "*" And INCONSISTENCIA = ""
            Call ValidarCampo_VL_PIS(Campos, dicTitulos)
            
        Case RegistrosCOFINS Like "*" & REG & "*" And INCONSISTENCIA = ""
            Call ValidarCampo_VL_COFINS(Campos, dicTitulos)
            
        Case Else
            Call ValidarCampo_VL_PIS(Campos, dicTitulos)
            If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidarCampo_VL_COFINS(Campos, dicTitulos)
            
    End Select
    
End Function

Private Function ValidarCampo_VL_PIS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim ALIQ_PIS As Double, VL_PIS#, VL_BC_PIS_CALC#, VL_PIS_CALC#, DIFERENCA#
Dim CST_PIS$, CFOP$, INCONSISTENCIA$, DT_DOC$, DT_ENT_SAI$, SUGESTAO$
Dim fatCFOP As Boolean
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do relatório inteligente de PIS/COFINS
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_PIS = Util.ApenasNumeros(Campos(dicTitulos("CST_PIS") - i))
    ALIQ_PIS = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_PIS") - i))
    VL_PIS = fnExcel.ConverterValores(Campos(dicTitulos("VL_PIS") - i), True, 2)
    
    'Campos calculados
    VL_BC_PIS_CALC = CalcularBasePISCOFINS(dicTitulos, Campos)
    VL_PIS_CALC = VBA.Round(VL_BC_PIS_CALC * ALIQ_PIS, 2)
    
    DIFERENCA = VBA.Round(VBA.Abs(VL_PIS_CALC - VL_PIS), 2)
    
    'Verificações de dados com CFOP
    If CFOP <> "" Then
        
        fatCFOP = ValidacoesCFOP.ValidarCFOPFaturamento(CFOP)
        
    End If
    
    If CFOP <> "" Then
        
        Select Case True
            
            Case CFOP > 4000 And Not fatCFOP And VL_PIS > 0
                INCONSISTENCIA = "CFOP informado não indica receita operacional com VL_PIS maior que zero"
                SUGESTAO = "Zerar valores do PIS"
                
            Case CFOP < 4000 And (CFOP Like "#406" Or CFOP Like "#551") And VL_PIS > 0
                INCONSISTENCIA = "CFOP (" & CFOP & ") indicando operação de aquisção para ativo imobilizado com VL_PIS (R$ " & VL_PIS & ") maior que zero"
                SUGESTAO = "Zerar valores do PIS"
                
            Case CFOP < 4000 And (CFOP Like "#407" Or CFOP Like "#556") And VL_PIS > 0
                INCONSISTENCIA = "CFOP (" & CFOP & ") indicando operação de aquisção para uso e consumo com VL_PIS (R$ " & VL_PIS & ") maior que zero"
                SUGESTAO = "Zerar valores do PIS"
                
            Case CFOP < 4000 And CFOP Like "#910" And VL_PIS > 0
                INCONSISTENCIA = "CFOP (" & CFOP & ") indicando operação de entrada em bonificação com VL_PIS (R$ " & VL_PIS & ") maior que zero"
                SUGESTAO = "Zerar valores do PIS"
                                
        End Select
        
    End If
    
    If INCONSISTENCIA = "" Then
        
        Select Case True
            
            Case VL_PIS_CALC > VL_PIS And VL_PIS_CALC > 0
                INCONSISTENCIA = "Cálculo do campo VL_PIS maior que o destacado [Diferença: R$ " & DIFERENCA & "]"
                SUGESTAO = "Recalcular valor do PIS"
                
            Case VL_PIS_CALC < VL_PIS And VL_PIS_CALC > 0
                INCONSISTENCIA = "Cálculo do campo VL_PIS menor que o destacado [Diferença: R$ " & DIFERENCA & "]"
                SUGESTAO = "Recalcular valor do PIS"
                
            Case CST_PIS Like "7*" And VL_PIS > 0
                INCONSISTENCIA = "CST_PIS indica operação sem direito a crédito com campo VL_PIS maior que zero"
                SUGESTAO = "Zerar valores do PIS"
                
            Case (CST_PIS < 4 Or CST_PIS Like "5*") And Not CST_PIS Like "00*" And VL_PIS = 0 And VL_PIS_CALC > 0
                INCONSISTENCIA = "CST_PIS indicando operação tributada com campo VL_PIS igual a zero"
                SUGESTAO = "Recalcular valor do PIS"
                
            Case CST_PIS > 3 And CST_PIS < 10 And CST_PIS <> 5 And VL_PIS > 0
                INCONSISTENCIA = "CST_PIS indicando operação não tributada com campo VL_PIS maior que zero"
                SUGESTAO = "Zerar valores do PIS"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Private Function ValidarCampo_VL_COFINS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim VL_ITEM As Double, VL_DESP#, VL_DESC#, VL_ICMS#, VL_BC_COFINS#, VL_BC_COFINS_CALC#, VL_COFINS_CALC#, VL_ICMS_ST#, VL_COFINS#, ALIQ_COFINS#, DIFERENCA#
Dim CST_COFINS$, CFOP$, INCONSISTENCIA$, SUGESTAO$
Dim fatCFOP As Boolean
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do relatório inteligente de COFINS/COFINS
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_COFINS = Util.ApenasNumeros(Campos(dicTitulos("CST_COFINS") - i))
    ALIQ_COFINS = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_COFINS") - i))
    VL_COFINS = fnExcel.ConverterValores(Campos(dicTitulos("VL_COFINS") - i), True, 2)
    
    'Campos calculados
    VL_BC_COFINS_CALC = CalcularBasePISCOFINS(dicTitulos, Campos)
    VL_COFINS_CALC = VBA.Round(VL_BC_COFINS_CALC * ALIQ_COFINS, 2)
    
    DIFERENCA = VBA.Round(VBA.Abs(VL_COFINS_CALC - VL_COFINS), 2)
    
    'Verificações de dados com CFOP
    If CFOP <> "" Then
        
        fatCFOP = ValidacoesCFOP.ValidarCFOPFaturamento(CFOP)
        
    End If
    
    If CFOP <> "" Then
        
        Select Case True
            
            Case CFOP > 4000 And Not fatCFOP And VL_COFINS > 0
                INCONSISTENCIA = "CFOP informado não indica receita operacional com VL_COFINS maior que zero"
                SUGESTAO = "Zerar valores da COFINS"
                
            Case CFOP < 4000 And (CFOP Like "#406" Or CFOP Like "#551") And VL_COFINS > 0
                INCONSISTENCIA = "CFOP (" & CFOP & ") indicando operação de aquisção para ativo imobilizado com VL_COFINS (R$ " & VL_COFINS & ") maior que zero"
                SUGESTAO = "Zerar valores da COFINS"
                
            Case CFOP < 4000 And (CFOP Like "#407" Or CFOP Like "#556") And VL_COFINS > 0
                INCONSISTENCIA = "CFOP (" & CFOP & ") indicando operação de aquisção para uso e consumo com VL_COFINS (R$ " & VL_COFINS & ") maior que zero"
                SUGESTAO = "Zerar valores da COFINS"
                
            Case CFOP < 4000 And CFOP Like "#910" And VL_COFINS > 0
                INCONSISTENCIA = "CFOP (" & CFOP & ") indicando operação de entrada em bonificação com VL_COFINS (R$ " & VL_COFINS & ") maior que zero"
                SUGESTAO = "Zerar valores da COFINS"
                
        End Select
        
    End If
    
    If INCONSISTENCIA = "" Then
        
        Select Case True
            
            Case VL_COFINS_CALC > VL_COFINS And VL_COFINS_CALC > 0
                INCONSISTENCIA = "Cálculo do campo VL_COFINS maior que o destacado [Diferença: R$ " & DIFERENCA & "]"
                SUGESTAO = "Recalcular valor da COFINS"
                
            Case VL_COFINS_CALC < VL_COFINS And VL_COFINS_CALC > 0
                INCONSISTENCIA = "Cálculo do campo VL_COFINS menor que o destacado [Diferença: R$ " & DIFERENCA & "]"
                SUGESTAO = "Recalcular valor da COFINS"
                
            Case CST_COFINS Like "7*" And VL_COFINS > 0
                INCONSISTENCIA = "CST_COFINS indica operação sem direito a crédito com campo VL_COFINS maior que zero"
                SUGESTAO = "Zerar valores da COFINS"
                
            Case (CST_COFINS < 4 Or CST_COFINS Like "5*") And Not CST_COFINS Like "00*" And VL_COFINS = 0 And VL_COFINS_CALC > 0
                INCONSISTENCIA = "CST_COFINS indicando operação tributada com campo VL_COFINS igual a zero"
                SUGESTAO = "Recalcular valor da COFINS"
                
            Case CST_COFINS > 3 And CST_COFINS < 10 And CST_COFINS <> 5 And VL_COFINS > 0
                INCONSISTENCIA = "CST_COFINS indicando operação não tributada com campo VL_COFINS maior que zero"
                SUGESTAO = "Zerar valores da COFINS"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_ALIQ_PIS_COFINS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal COD_INC_TRIB As String)
'TODO: Refatorar código para a nova estrutura usando a estrutura CamposPISCOFINS As CamposApuracaoPISCOFINS
Dim i As Byte
Dim REG As String, INCONSISTENCIA$
    
    If LBound(Campos) = 0 Then i = 1
    
    'Identificação do registro
    REG = Campos(dicTitulos("REG") - i)
    INCONSISTENCIA = Campos(dicTitulos("INCONSISTENCIA") - i)
    
    Select Case True
        
        Case RegistrosPIS Like "*" & REG & "*" And INCONSISTENCIA = ""
            Call ValidarCampo_ALIQ_PIS(Campos, dicTitulos, COD_INC_TRIB)
            
        Case RegistrosCOFINS Like "*" & REG & "*" And INCONSISTENCIA = ""
            Call ValidarCampo_ALIQ_COFINS(Campos, dicTitulos, COD_INC_TRIB)
            
        Case Else
            Call ValidarCampo_ALIQ_PIS(Campos, dicTitulos, COD_INC_TRIB)
            If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidarCampo_ALIQ_COFINS(Campos, dicTitulos, COD_INC_TRIB)
            
    End Select
    
End Function

Private Function ValidarCampo_ALIQ_PIS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal COD_INC_TRIB As String)
    
Dim VL_BC_PIS As Double, VL_BC_COFINS#, VL_BC_PIS_CALC#, VL_PIS_CALC#, VL_PIS#, ALIQ_PIS#
Dim CST_PIS$, CFOP$, INCONSISTENCIA$, SUGESTAO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do relatório inteligente de PIS/COFINS
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_PIS = Util.ApenasNumeros(Campos(dicTitulos("CST_PIS") - i))
    ALIQ_PIS = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_PIS") - i))
    
    'Campos calculados
    VL_BC_PIS_CALC = CalcularBasePISCOFINS(dicTitulos, Campos)
    VL_PIS_CALC = VBA.Round(VL_BC_PIS_CALC * ALIQ_PIS, 2)
    
    'Verificações de dados com CFOP
    If CFOP <> "" Then
        
        fatCFOP = ValidacoesCFOP.ValidarCFOPFaturamento(CFOP)
        devCompCFOP = ValidacoesCFOP.ValidarCFOPDevolucaoCompra(CFOP)
        
    End If
    
    If CFOP <> "" Then
        
        Select Case True
            
            Case CFOP > 4000 And Not fatCFOP And ALIQ_PIS > 0
                INCONSISTENCIA = "CFOP informado não indica receita operacional com ALIQ_PIS maior que zero"
                SUGESTAO = "Zerar valores do PIS"
                
        End Select
        
    End If
    
    If INCONSISTENCIA = "" Then
        
        Select Case True
            
            Case CST_PIS Like "*05*" And ALIQ_PIS = 0
                INCONSISTENCIA = "Para operações com o CST_PIS = 05 informar o campo ALIQ_PIS maior que zero"
                SUGESTAO = "Informar uma alíquota maior que 0 para o campo ALIQ_PIS"
                
            Case COD_INC_TRIB = "1" And ALIQ_PIS > 0 And ALIQ_PIS <> 0.0165
                INCONSISTENCIA = "Empresa no regime não-cumulativo com ALIQ_PIS diferente de 1,65%"
                SUGESTAO = "Informar alíquota de 1,65% para o PIS"
                
            Case COD_INC_TRIB = "2" And ALIQ_PIS > 0 And ALIQ_PIS <> 0.0065
                INCONSISTENCIA = "Empresa no regime cumulativo com ALIQ_PIS diferente de 0,65%"
                SUGESTAO = "Informar alíquota de 0,65% para o PIS"
                
            Case CST_PIS Like "7*" And ALIQ_PIS > 0
                INCONSISTENCIA = "CST_PIS indicando operação sem crédito do imposto com campo ALIQ_PIS maior que zero"
                SUGESTAO = "Zerar valores do PIS"
                
            Case CST_PIS Like "5*" And ALIQ_PIS = 0
                INCONSISTENCIA = "CST_PIS indicando operação com crédito do imposto com campo ALIQ_PIS igual a zero"
                If COD_INC_TRIB = "1" Then SUGESTAO = "Informar alíquota de 1,65% para o PIS"
                If COD_INC_TRIB = "2" Then SUGESTAO = "Informar alíquota de 0,65% para o PIS"
                If COD_INC_TRIB = "3" Or COD_INC_TRIB = "" Then SUGESTAO = "Informar alíquota do PIS"
                SUGESTAO = SUGESTAO
                
            Case (CST_PIS > 0 And CST_PIS < 4 Or CST_PIS Like "5*") And ALIQ_PIS = 0
                If COD_INC_TRIB = "1" Then SUGESTAO = "Informar alíquota de 1,65% para o PIS"
                If COD_INC_TRIB = "2" Then SUGESTAO = "Informar alíquota de 0,65% para o PIS"
                If COD_INC_TRIB = "3" Or COD_INC_TRIB = "" Then SUGESTAO = "Informar alíquota do PIS"
                INCONSISTENCIA = "CST_PIS indicando operação tributada com campo ALIQ_PIS igual a zero"
                SUGESTAO = SUGESTAO
                
            Case CST_PIS > 3 And CST_PIS < 10 And CST_PIS <> 5 And ALIQ_PIS <> 0
                INCONSISTENCIA = "CST_PIS indicando operação não tributada com campo ALIQ_PIS diferente de zero"
                SUGESTAO = "Zerar valores do PIS"
                
            Case VL_BC_PIS_CALC < 0 And VL_BC_PIS <> 0 And ALIQ_PIS <> 0
                INCONSISTENCIA = "O valor calculado do campo VL_BC_PIS está negativo"
                SUGESTAO = "Zerar valores do PIS"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Private Function ValidarCampo_ALIQ_COFINS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal COD_INC_TRIB As String)
    
Dim VL_BC_COFINS As Double, VL_BC_COFINS_CALC#, VL_COFINS_CALC#, VL_COFINS#, ALIQ_COFINS#
Dim CST_COFINS$, CFOP$, INCONSISTENCIA$, SUGESTAO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1

    'Carregar campos do relatório inteligente de COFINS/COFINS
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_COFINS = Util.ApenasNumeros(Campos(dicTitulos("CST_COFINS") - i))
    ALIQ_COFINS = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_COFINS") - i))

    'Campos calculados
    VL_BC_COFINS_CALC = CalcularBasePISCOFINS(dicTitulos, Campos)
    VL_COFINS_CALC = VBA.Round(VL_BC_COFINS_CALC * ALIQ_COFINS, 2)


    'Verificações de dados com CFOP
    If CFOP <> "" Then
    
        fatCFOP = ValidacoesCFOP.ValidarCFOPFaturamento(CFOP)
        devCompCFOP = ValidacoesCFOP.ValidarCFOPDevolucaoCompra(CFOP)
    
    End If
    
    If CFOP <> "" Then
        
        Select Case True
            
            Case CFOP > 4000 And Not fatCFOP And ALIQ_COFINS > 0
                INCONSISTENCIA = "CFOP informado não indica receita operacional com ALIQ_COFINS maior que zero"
                SUGESTAO = "Zerar valores da COFINS"
        
        End Select
        
    End If
    
    If INCONSISTENCIA = "" Then
        
        Select Case True
        
            Case CST_COFINS Like "*05*" And ALIQ_COFINS = 0
                INCONSISTENCIA = "Para operações com o CST_COFINS = 05 informar o campo ALIQ_COFINS maior que zero"
                SUGESTAO = "Informar uma alíquota maior que 0 para o campo ALIQ_COFINS"
                
            Case COD_INC_TRIB = "1" And ALIQ_COFINS > 0 And ALIQ_COFINS <> 0.076
                INCONSISTENCIA = "Empresa no regime não-cumulativo com ALIQ_COFINS diferente de 7,6%"
                SUGESTAO = "Informar alíquota de 7,60% para a COFINS"
                
            Case COD_INC_TRIB = "2" And ALIQ_COFINS > 0 And ALIQ_COFINS <> 0.03
                INCONSISTENCIA = "Empresa no regime cumulativo com ALIQ_COFINS diferente de 3,00%"
                SUGESTAO = "Informar alíquota de 3,00% para a COFINS"
                
            Case CST_COFINS Like "7*" And ALIQ_COFINS > 0
                INCONSISTENCIA = "CST_COFINS indicando operação sem crédito do imposto com campo ALIQ_COFINS maior que zero"
                SUGESTAO = "Zerar valores da COFINS"
                
            Case CST_COFINS Like "5*" And ALIQ_COFINS = 0
                INCONSISTENCIA = "CST_COFINS indicando operação com crédito do imposto com campo ALIQ_COFINS igual a zero"
                If COD_INC_TRIB = "1" Then SUGESTAO = "Informar alíquota de 7,60% para a COFINS"
                If COD_INC_TRIB = "2" Then SUGESTAO = "Informar alíquota de 3,00% para a COFINS"
                If COD_INC_TRIB = "3" Or COD_INC_TRIB = "" Then SUGESTAO = "Informar alíquota da COFINS"
                SUGESTAO = SUGESTAO
                
            Case (CST_COFINS > 0 And CST_COFINS < 4 Or CST_COFINS Like "5*") And ALIQ_COFINS = 0
                If COD_INC_TRIB = "1" Then SUGESTAO = "Informar alíquota de 7,60% para a COFINS"
                If COD_INC_TRIB = "2" Then SUGESTAO = "Informar alíquota de 3,00% para a COFINS"
                If COD_INC_TRIB = "3" Or COD_INC_TRIB = "" Then SUGESTAO = "Informar alíquota da COFINS"
                INCONSISTENCIA = "CST_COFINS indicando operação tributada com campo ALIQ_COFINS igual a zero"
                SUGESTAO = SUGESTAO
                
            Case CST_COFINS > 3 And CST_COFINS < 10 And CST_COFINS <> 5 And ALIQ_COFINS <> 0
                INCONSISTENCIA = "CST_COFINS indicando operação não tributada com campo ALIQ_COFINS diferente de zero"
                SUGESTAO = "Zerar valores da COFINS"
                
            Case VL_BC_COFINS_CALC < 0 And VL_BC_COFINS <> 0 And ALIQ_COFINS <> 0
                INCONSISTENCIA = "O valor calculado do campo VL_BC_COFINS está negativo"
                SUGESTAO = "Zerar valores da COFINS"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_VL_BC_PIS_COFINS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
'TODO: Refatorar código para a nova estrutura usando a estrutura CamposPISCOFINS As CamposApuracaoPISCOFINS
Dim i As Byte
Dim REG As String, INCONSISTENCIA$
    
    If LBound(Campos) = 0 Then i = 1
    
    'Identificação do registro
    REG = Campos(dicTitulos("REG") - i)
    INCONSISTENCIA = Campos(dicTitulos("INCONSISTENCIA") - i)
    
    Select Case True
        
        Case RegistrosPIS Like "*" & REG & "*" And INCONSISTENCIA = ""
            Call ValidarCampo_VL_BC_PIS(Campos, dicTitulos)
            
        Case RegistrosCOFINS Like "*" & REG & "*" And INCONSISTENCIA = ""
            Call ValidarCampo_VL_BC_COFINS(Campos, dicTitulos)
            
        Case Else
            Call ValidarCampo_VL_BC_PIS(Campos, dicTitulos)
            If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidarCampo_VL_BC_COFINS(Campos, dicTitulos)
            
    End Select
    
End Function

Private Function ValidarCampo_VL_BC_PIS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim VL_BC_PIS As Double, VL_BC_COFINS#, VL_BC_PIS_CALC#, VL_PIS_CALC#, VL_PIS#, ALIQ_PIS#
Dim CST_PIS$, CFOP$, REG$, INCONSISTENCIA$, SUGESTAO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1

    'Carregar campos do relatório inteligente de PIS/COFINS
    REG = Campos(dicTitulos("REG") - i)
    CST_PIS = Util.ApenasNumeros(Campos(dicTitulos("CST_PIS") - i))
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    VL_PIS = fnExcel.ConverterValores(Campos(dicTitulos("VL_PIS") - i), True, 2)
    VL_BC_PIS = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_PIS") - i), True, 2)
    VL_BC_COFINS = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_COFINS") - i), True, 2)
    
    'Campos calculados
    VL_BC_PIS_CALC = CalcularBasePISCOFINS(dicTitulos, Campos)
    VL_PIS_CALC = VBA.Round(VL_BC_PIS_CALC * ALIQ_PIS, 2)

    'Verificações de dados com CFOP
    If CFOP <> "" Then
    
        fatCFOP = ValidacoesCFOP.ValidarCFOPFaturamento(CFOP)
    
    End If
    
    If CFOP <> "" Then
        
        Select Case True

            Case CFOP > 4000 And Not fatCFOP And VBA.Round(VL_BC_PIS, 2) > 0
                INCONSISTENCIA = "CFOP informado não indica receita operacional com VL_BC_PIS maior que zero"
                SUGESTAO = "Zerar valores do PIS"
        
        End Select
    
    End If
    
    If INCONSISTENCIA = "" Then
        
        Select Case True
            
            Case VL_BC_PIS_CALC < 0
                INCONSISTENCIA = "O cálculo do campo VL_BC_PIS está negativo (R$" & VL_BC_PIS_CALC & ")"
                SUGESTAO = "Verificar o cálculo da base de cálculo para encontrar o motivo da inconsistência"
                
            Case CST_PIS Like "7*" And VL_BC_PIS > 0
                INCONSISTENCIA = "CST_PIS indicando operação sem crédito do imposto com campo VL_BC_PIS maior que zero"
                SUGESTAO = "Zerar valores do PIS"
            
            Case CST_PIS > 0 And CST_PIS < 6 And CST_PIS <> 4 And VL_BC_PIS = 0 And VL_BC_PIS_CALC > 0
                INCONSISTENCIA = "CST_PIS indicando operação tributada com campo VL_BC_PIS igual a zero"
                SUGESTAO = "Gerar base de cálculo do PIS"
            
            Case (CST_PIS < 4 Or CST_PIS Like "5*" Or CST_PIS Like "6*") And Not CST_PIS Like "00*" And VL_BC_PIS = 0 And VL_BC_PIS_CALC > 0
                INCONSISTENCIA = "CST_PIS indicando operação tributada com campo VL_BC_PIS igual a zero"
                SUGESTAO = "Gerar base de cálculo do PIS"
                
            Case VL_BC_PIS_CALC > VL_BC_PIS And VL_BC_PIS_CALC <> 0 And (CST_PIS < 4 Or CST_PIS Like "5*")
                INCONSISTENCIA = "Base de cálculo do PIS (VL_BC_PIS) está informada a menor"
                SUGESTAO = "Gerar base de cálculo do PIS"
                
            Case VL_BC_PIS_CALC < VL_BC_PIS And VL_BC_PIS_CALC <> 0 And (CST_PIS < 4 Or CST_PIS Like "5*")
                INCONSISTENCIA = "Base de cálculo do PIS (VL_BC_PIS) está informada a maior"
                SUGESTAO = "Gerar base de cálculo do PIS"
                
            Case VL_BC_PIS_CALC > VL_BC_PIS And VL_BC_PIS_CALC <> 0 And VL_PIS > 0
                INCONSISTENCIA = "Base de cálculo do PIS (VL_BC_PIS) está informada a menor"
                SUGESTAO = "Gerar base de cálculo do PIS"
                
            Case VL_BC_PIS_CALC < VL_BC_PIS And VL_BC_PIS_CALC <> 0 And VL_PIS > 0
                INCONSISTENCIA = "Base de cálculo do PIS (VL_BC_PIS) está informada a maior"
                SUGESTAO = "Gerar base de cálculo do PIS"
                
            Case CST_PIS > 3 And CST_PIS < 10 And CST_PIS <> 5 And VL_BC_PIS > 0
                INCONSISTENCIA = "CST_PIS indicando operação não tributada com campo VL_BC_PIS maior que zero"
                SUGESTAO = "Zerar valores do PIS"
                
            Case VL_BC_PIS <> VL_BC_COFINS And Not RegistrosPIS Like "*" & REG & "*"
                INCONSISTENCIA = "Os campos 'VL_BC_PIS' e 'VL_BC_COFINS' estão com valores diferentes"
                SUGESTAO = "Recalcular bases de PIS e COFINS"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Private Function ValidarCampo_VL_BC_COFINS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim VL_BC_COFINS As Double, VL_BC_COFINS_CALC#, VL_COFINS_CALC#, VL_BC_PIS#, VL_COFINS#, ALIQ_COFINS#
Dim CST_COFINS$, CFOP$, REG$, INCONSISTENCIA$, SUGESTAO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1

    'Carregar campos do relatório inteligente de COFINS/COFINS
    REG = Campos(dicTitulos("REG") - i)
    CST_COFINS = Util.ApenasNumeros(Campos(dicTitulos("CST_COFINS") - i))
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    VL_BC_PIS = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_PIS") - i), True, 2)
    VL_COFINS = fnExcel.ConverterValores(Campos(dicTitulos("VL_COFINS") - i), True, 2)
    VL_BC_COFINS = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_COFINS") - i), True, 2)

    'Campos calculados
    VL_BC_COFINS_CALC = CalcularBasePISCOFINS(dicTitulos, Campos)
    VL_COFINS_CALC = VBA.Round(VL_BC_COFINS_CALC * ALIQ_COFINS, 2)

    'Verificações de dados com CFOP
    If CFOP <> "" Then
    
        fatCFOP = ValidacoesCFOP.ValidarCFOPFaturamento(CFOP)
    
    End If
    
    If CFOP <> "" Then
        
        Select Case True

            Case CFOP > 4000 And Not fatCFOP And VBA.Round(VL_BC_COFINS, 2) > 0
                INCONSISTENCIA = "CFOP informado não indica receita operacional com VL_BC_COFINS maior que zero"
                SUGESTAO = "Zerar valores da COFINS"
        
        End Select
    
    End If
    
    If INCONSISTENCIA = "" Then
        
        Select Case True
            
            Case VL_BC_COFINS_CALC < 0
                INCONSISTENCIA = "O cálculo do campo VL_BC_COFINS está negativo (R$" & VL_BC_COFINS_CALC & ")"
                SUGESTAO = "Verificar o cálculo da base de cálculo para encontrar o motivo da inconsistência"
                
            Case CST_COFINS Like "7*" And VL_BC_COFINS > 0
                INCONSISTENCIA = "CST_COFINS indicando operação sem crédito do imposto com campo VL_BC_COFINS maior que zero"
                SUGESTAO = "Zerar valores da COFINS"
                
            Case CST_COFINS > 0 And CST_COFINS < 6 And CST_COFINS <> 4 And VL_BC_COFINS = 0 And VL_BC_COFINS_CALC > 0
                INCONSISTENCIA = "CST_COFINS indicando operação tributada com campo VL_BC_COFINS igual a zero"
                SUGESTAO = "Gerar base de cálculo da COFINS"
                
            Case (CST_COFINS < 4 Or CST_COFINS Like "5*" Or CST_COFINS Like "6*") And Not CST_COFINS Like "00*" And VL_BC_COFINS = 0 And VL_BC_COFINS_CALC > 0
                INCONSISTENCIA = "CST_COFINS indicando operação tributada com campo VL_BC_COFINS igual a zero"
                SUGESTAO = "Gerar base de cálculo da COFINS"
                
            Case VL_BC_COFINS_CALC > VL_BC_COFINS And VL_BC_COFINS_CALC <> 0 And (CST_COFINS < 4 Or CST_COFINS Like "5*")
                INCONSISTENCIA = "Base de cálculo da COFINS (VL_BC_COFINS) está informada a menor"
                SUGESTAO = "Gerar base de cálculo da COFINS"
                
            Case VL_BC_COFINS_CALC < VL_BC_COFINS And VL_BC_COFINS_CALC <> 0 And (CST_COFINS < 4 Or CST_COFINS Like "5*")
                INCONSISTENCIA = "Base de cálculo da COFINS (VL_BC_COFINS) está informada a maior"
                SUGESTAO = "Gerar base de cálculo da COFINS"
                
            Case VL_BC_COFINS_CALC > VL_BC_COFINS And VL_BC_COFINS_CALC <> 0 And VL_COFINS > 0
                INCONSISTENCIA = "Base de cálculo da COFINS (VL_BC_COFINS) está informada a menor"
                SUGESTAO = "Gerar base de cálculo da COFINS"
                
            Case VL_BC_COFINS_CALC < VL_BC_COFINS And VL_BC_COFINS_CALC <> 0 And VL_COFINS > 0
                INCONSISTENCIA = "Base de cálculo da COFINS (VL_BC_COFINS) está informada a maior"
                SUGESTAO = "Gerar base de cálculo da COFINS"
                
            Case CST_COFINS > 3 And CST_COFINS < 10 And CST_COFINS <> 5 And VL_BC_COFINS > 0
                INCONSISTENCIA = "CST_COFINS indicando operação não tributada com campo VL_BC_COFINS maior que zero"
                SUGESTAO = "Zerar valores da COFINS"
                
            Case VL_BC_PIS <> VL_BC_COFINS And Not RegistrosCOFINS Like "*" & REG & "*"
                INCONSISTENCIA = "Os campos 'VL_BC_PIS' e 'VL_BC_COFINS' estão com valores diferentes"
                SUGESTAO = "Recalcular bases de PIS e COFINS"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_COD_NAT_PIS_COFINS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim CST_PIS$, CST_COFINS$, CFOP$, COD_NAT_PIS_COFINS$, INCONSISTENCIA$, SUGESTAO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim VL_PIS As Double, VL_COFINS#
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    COD_NAT_PIS_COFINS = Util.ApenasNumeros(Campos(dicTitulos("COD_NAT_PIS_COFINS") - i))
    CST_PIS = Util.ApenasNumeros(Campos(dicTitulos("CST_PIS") - i))
    CST_COFINS = Util.ApenasNumeros(Campos(dicTitulos("CST_COFINS") - i))
    VL_PIS = fnExcel.ConverterValores(Campos(dicTitulos("VL_PIS") - i))
    VL_COFINS = fnExcel.ConverterValores(Campos(dicTitulos("VL_COFINS") - i))
    
    Select Case True
        
        Case CST_PIS > "03" And CST_PIS < "10" And VL_PIS = 0 And COD_NAT_PIS_COFINS = ""
            INCONSISTENCIA = "O campo COD_NAT_PIS_COFINS deve ser informado para operações entre o CST 04 e 09"
            SUGESTAO = "Informar um valor válido para o campo COD_NAT_PIS_COFINS"
            
        Case CST_COFINS > "03" And CST_COFINS < "10" And VL_COFINS = 0 And COD_NAT_PIS_COFINS = ""
            INCONSISTENCIA = "O campo COD_NAT_COFINS_COFINS deve ser informado para operações entre o CST 04 e 09"
            SUGESTAO = "Informar um valor válido para o campo COD_NAT_COFINS_COFINS"
                        
    End Select
    
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

Public Function ValidarCampo_REGIME_TRIBUTARIO(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim REGIME_TRIBUTARIO As String, INCONSISTENCIA$, SUGESTAO$, CST_PIS$, CST_COFINS$
Dim ALIQ_PIS As Double, ALIQ_COFINS#
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    REGIME_TRIBUTARIO = Campos(dicTitulos("REGIME_TRIBUTARIO") - i)
    ALIQ_PIS = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_PIS") - i))
    ALIQ_COFINS = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_COFINS") - i))
    CST_PIS = Util.ApenasNumeros(Campos(dicTitulos("CST_PIS") - i))
    CST_COFINS = Util.ApenasNumeros(Campos(dicTitulos("CST_COFINS") - i))
    
    Select Case True
        
        Case REGIME_TRIBUTARIO = "", REGIME_TRIBUTARIO Like "*Código Inválido*", REGIME_TRIBUTARIO Like "*DEFINA UM REGIME PARA A OPERAÇÃO*"
            INCONSISTENCIA = "Regime Tributário não definido para a Operação"
            SUGESTAO = "Informe um dos seguintes valores no campo REGIME_TRIBUTARIO (1 para: Não-Cumultivo ou 2 para: Cumulativo)"
            
        Case Not REGIME_TRIBUTARIO Like "2*" And ALIQ_PIS = 0.0065
            INCONSISTENCIA = "Regime Tributário indicado pelo campo ALIQ_PIS (2 - Cumulativo) divergente do campo REGIME_TRIBUTARIO"
            SUGESTAO = "Alterar campo REGIME_TRIBUTARIO para 2 - Cumulativo"
            
        Case Not REGIME_TRIBUTARIO Like "2*" And ALIQ_COFINS = 0.03
            INCONSISTENCIA = "Regime Tributário indicado pelo campo ALIQ_COFINS (2 - Cumulativo) divergente do campo REGIME_TRIBUTARIO"
            SUGESTAO = "Alterar campo REGIME_TRIBUTARIO para 2 - Cumulativo"
            
        Case Not REGIME_TRIBUTARIO Like "1*" And ALIQ_PIS = 0.0165
            INCONSISTENCIA = "Regime Tributário indicado pelo campo ALIQ_PIS (1 - Não-Cumulativo) divergente do campo REGIME_TRIBUTARIO"
            SUGESTAO = "Alterar campo REGIME_TRIBUTARIO para 1 - Não-Cumulativo"

        Case Not REGIME_TRIBUTARIO Like "1*" And ALIQ_COFINS = 0.076
            INCONSISTENCIA = "Regime Tributário indicado pelo campo ALIQ_COFINS (1 - Não-Cumulativo) divergente do campo REGIME_TRIBUTARIO"
            SUGESTAO = "Alterar campo REGIME_TRIBUTARIO para 1 - Não-Cumulativo"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Private Function CarregarCFOPSGeradoresCredito(ByRef ListaCFOPsCreditaveis As ArrayList)

Dim CFOPS As Variant, CFOP

    CFOPS = Array("1102", "1113", "1117", "1118", "1121", "1251", "1403", "1652", "2102", "2113", "2117", "2118", "2121", "2251", "2403", "2652", "3102", "3251", "3652", "1101", "1111", "1116", "1120", "1122", "1126", "1128", "1401", "1407", "1556", "1651", "1653", "2101", "2111", "2116", "2120", "2122", "2126", "2128", "2401", "2407", "2556", "2651", "2653", "3101", "3126", "3128", "3556", "3651", "3653", "1124", "1125", "1933", "2124", "2125", "2933", "1201", "1202", "1203", "1204", "1410", "1411", "1660", "1661", "1662", "2201", "2202", "2410", "2411", "2660", "2661", "2662", "1922", "2922", "1206", "2206", "1207", "2207", "1135", "2135", "1132", "2132", "1215", "1216", "2215", "2216", "1159", "2159", "1456", "2456")
    
    For Each CFOP In CFOPS
        
        If Not ListaCFOPsCreditaveis.contains(CFOP) Then ListaCFOPsCreditaveis.Add CFOP
        
    Next CFOP

End Function

Public Function CalcularBasePISCOFINS(ByRef dicTitulos As Dictionary, ByRef Campos As Variant) As Double

Dim i As Byte
Dim CFOP As String, DT_DOC$, DT_ENT_SAI$, Periodo$
Dim VL_ITEM As Double, VL_DESP#, VL_DESC#, VL_ICMS#, VL_BC_PIS_COFINS#
    
    CFOP = fnExcel.FormatarValores(Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i)), , 0)
    VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM") - i), True, 2)
    VL_DESP = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESP") - i), True, 2)
    VL_DESC = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESC") - i), True, 2)
    VL_ICMS = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS") - i), True, 2)
    DT_DOC = fnExcel.FormatarData(Campos(dicTitulos("DT_DOC") - i))
    DT_ENT_SAI = fnExcel.FormatarData(Campos(dicTitulos("DT_ENT_SAI") - i))
    Periodo = VBA.Split(Campos(dicTitulos("ARQUIVO") - i), "-")(0)
    
    If DT_ENT_SAI = "" Then DT_ENT_SAI = Util.ConverterPeriodoData(Periodo)
    If DT_DOC = "" Then DT_DOC = Util.ConverterPeriodoData(Periodo)
    
    Select Case True
        
        Case CFOP < 4000 And CDate(DT_ENT_SAI) > CDate("2023-05-01")
            VL_BC_PIS_COFINS = VBA.Round(VL_ITEM + VL_DESP - VL_DESC - VL_ICMS, 2)
            
        Case CFOP > 4000 And CDate(DT_DOC) > CDate("2017-03-15")
            VL_BC_PIS_COFINS = VBA.Round(VL_ITEM + VL_DESP - VL_DESC - VL_ICMS, 2)
                    
        Case Else
            VL_BC_PIS_COFINS = VBA.Round(VL_ITEM + VL_DESP - VL_DESC, 2)
            
    End Select
    
    CalcularBasePISCOFINS = VL_BC_PIS_COFINS
    
End Function
