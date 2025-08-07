Attribute VB_Name = "clsRegrasCadastraisPISCOFINS"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private ValidacoesCFOP As New clsRegrasFiscaisCFOP

Public Function ValidarCampo_CST_PIS_COFINS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1

    Call ValidarCampo_CST_PIS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidarCampo_CST_COFINS(Campos, dicTitulos)
        
End Function

Private Function ValidarCampo_CST_PIS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim CST_PIS$, CFOP$, INCONSISTENCIA$, SUGESTAO$, REGIME_TRIBUTARIO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1

    'Carregar campos do relatório de tributação PIS/COFINS
    REGIME_TRIBUTARIO = Util.ApenasNumeros(Campos(dicTitulos("REGIME_TRIBUTARIO") - i))
    CST_PIS = Util.ApenasNumeros(Campos(dicTitulos("CST_PIS") - i))
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    
    'Verificações de dados com CFOP
    If CFOP <> "" Then
    
        fatCFOP = ValidacoesCFOP.ValidarCFOPFaturamento(CFOP)
        devCompCFOP = ValidacoesCFOP.ValidarCFOPDevolucaoCompra(CFOP)
    
    End If
    
    Select Case True
        
        Case CST_PIS = ""
            INCONSISTENCIA = "CST_PIS não foi informado"
            SUGESTAO = "Informar um valor válido para o CST_PIS"
            
        Case CST_PIS Like "00*"
            INCONSISTENCIA = "O CST_PIS informado está inválido"
            SUGESTAO = "Informar um CST_PIS válido"
            
    End Select
                    
    If CFOP <> "" And INCONSISTENCIA = "" Then
        
        Select Case True
            
            Case CFOP > 4000 And fatCFOP And CST_PIS Like "*49*"
                INCONSISTENCIA = "CFOP informado indica receita operacional com CST_PIS de outras saídas"
                SUGESTAO = "Informar CST_PIS correto"
                
            Case CFOP > 4000 And CST_PIS > 49 And CST_PIS < 99
                INCONSISTENCIA = "CST_PIS de entrada informado para CFOP de saída"
                SUGESTAO = "Informar um CST_PIS para operação de saída"
                
            Case CFOP < 4000 And CST_PIS < 50
                INCONSISTENCIA = "CST_PIS de saída informado para CFOP de entrada"
                SUGESTAO = "Informar um CST_PIS para operação de entrada"
                
            Case Not fatCFOP And CST_PIS < 7
                INCONSISTENCIA = "CST_PIS tributável com CFOP de operação não tributável"
                SUGESTAO = "Alterar CST_PIS"
                
            Case devCompCFOP And CST_PIS <> 49
                INCONSISTENCIA = "CST_PIS diferente de 49 informado em operação de devolução"
                SUGESTAO = "Alterar CST_PIS para 49"
                
            Case fatCFOP And CST_PIS > 9
                INCONSISTENCIA = "CST_PIS incorreto para operação de Venda"
                SUGESTAO = "Informar CST_PIS correto"
            
            Case (CFOP < 4000 And CFOP Like "#910") And Not CST_PIS Like "98*"
                INCONSISTENCIA = "CST_PIS (" & CST_PIS & ") incorreto para operação de entrada em bonificação"
                SUGESTAO = "Informar CST_PIS 98 - Outras Operações de Entrada"
            
            Case CFOP < 4000 And (CFOP Like "#407" Or CFOP Like "#556") And Not CST_PIS Like "98*"
                INCONSISTENCIA = "CST_PIS (" & CST_PIS & ") incorreto para operação de aquisição para uso e consumo"
                SUGESTAO = "Informar CST_PIS 98 - Outras Operações de Entrada"
                
            Case (CFOP > 4000 And CFOP Like "#910") And Not CST_PIS Like "49*"
                INCONSISTENCIA = "CST_PIS (" & CST_PIS & ") incorreto para operação de saída em bonificação"
                SUGESTAO = "Informar CST_PIS 49 - Outras Operações de Saída"

            Case (CFOP < 4000 And Not CST_PIS Like "9*" And Not CST_PIS Like "7*")
                If REGIME_TRIBUTARIO = "2" Then
                    SUGESTAO = "Informar CST_PIS igual a 70 - Operação de Aquisição sem Direito a Crédito"
                    INCONSISTENCIA = "CST_PIS inconsistente com apuração no Regime Cumulativo (Lucro Presumido)"
                    SUGESTAO = SUGESTAO
                End If
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
        
End Function

Private Function ValidarCampo_CST_COFINS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim CST_COFINS$, CFOP$, INCONSISTENCIA$, SUGESTAO$, REGIME_TRIBUTARIO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do relatório de tributação PIS/COFINS
    REGIME_TRIBUTARIO = Util.ApenasNumeros(Campos(dicTitulos("REGIME_TRIBUTARIO") - i))
    CST_COFINS = Util.ApenasNumeros(Campos(dicTitulos("CST_COFINS") - i))
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    
    'Verificações de dados com CFOP
    If CFOP <> "" Then
    
        fatCFOP = ValidacoesCFOP.ValidarCFOPFaturamento(CFOP)
        devCompCFOP = ValidacoesCFOP.ValidarCFOPDevolucaoCompra(CFOP)
    
    End If
    
    Select Case True
    
        Case CST_COFINS = ""
            INCONSISTENCIA = "CST_COFINS não foi informado"
            SUGESTAO = "Informar um valor válido para o CST_COFINS"
        
        Case CST_COFINS Like "00*"
            INCONSISTENCIA = "O CST_COFINS informado está inválido"
            SUGESTAO = "Informar um CST_COFINS válido"
            
    End Select
                    
    If CFOP <> "" And INCONSISTENCIA = "" Then
        
        Select Case True
            
            Case CFOP > 4000 And fatCFOP And CST_COFINS Like "*49*"
                INCONSISTENCIA = "CFOP informado indica receita operacional com CST_COFINS de outras saídas"
                SUGESTAO = "Informar CST_COFINS correto"
                
            Case CFOP > 4000 And CST_COFINS > 49 And CST_COFINS < 99
                INCONSISTENCIA = "CST_COFINS de entrada informado para CFOP de saída"
                SUGESTAO = "Informar um CST_COFINS para operação de saída"
                
            Case CFOP < 4000 And CST_COFINS < 50
                INCONSISTENCIA = "CST_COFINS de saída informado para CFOP de entrada"
                SUGESTAO = "Informar um CST_COFINS para operação de entrada"
                
            Case Not fatCFOP And CST_COFINS < 7
                INCONSISTENCIA = "CST_COFINS tributável com CFOP de operação não tributável"
                SUGESTAO = "Alterar CST_COFINS"
                
            Case devCompCFOP And CST_COFINS <> 49
                INCONSISTENCIA = "CST_COFINS diferente de 49 informado em operação de devolução"
                SUGESTAO = "Alterar CST_COFINS para 49"
                
            Case fatCFOP And CST_COFINS > 9
                INCONSISTENCIA = "CST_COFINS incorreto para operação de Venda"
                SUGESTAO = "Informar CST_COFINS correto"

            Case (CFOP < 4000 And CFOP Like "#910") And Not CST_COFINS Like "98*"
                INCONSISTENCIA = "CST_COFINS (" & CST_COFINS & ") incorreto para operação de entrada em bonificação"
                SUGESTAO = "Informar CST_COFINS 98 - Outras Operações de Entrada"
            
            Case CFOP < 4000 And (CFOP Like "#407" Or CFOP Like "#556") And Not CST_COFINS Like "98*"
                INCONSISTENCIA = "CST_COFINS (" & CST_COFINS & ") incorreto para operação de aquisição para uso e consumo"
                SUGESTAO = "Informar CST_COFINS 98 - Outras Operações de Entrada"
                
            Case (CFOP > 4000 And CFOP Like "#910") And Not CST_COFINS Like "49*"
                INCONSISTENCIA = "CST_COFINS (" & CST_COFINS & ") incorreto para operação de saída em bonificação"
                SUGESTAO = "Informar CST_COFINS 49 - Outras Operações de Saída"
                
            Case (CFOP < 4000 And Not CST_COFINS Like "9*" And Not CST_COFINS Like "7*")
                If REGIME_TRIBUTARIO = "2" Then
                    SUGESTAO = "Informar CST_COFINS igual a 70 - Operação de Aquisição sem Direito a Crédito"
                    INCONSISTENCIA = "CST_COFINS inconsistente com apuração no Regime Cumulativo (Lucro Presumido)"
                    SUGESTAO = SUGESTAO
                End If
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
        
End Function

Public Function ValidarCampo_ALIQ_PIS_COFINS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim i As Byte
Dim REG As String
    
    If LBound(Campos) = 0 Then i = 1
    
    Call ValidarCampo_ALIQ_PIS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidarCampo_ALIQ_COFINS(Campos, dicTitulos)
        
End Function

Private Function ValidarCampo_ALIQ_PIS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim CST_PIS$, CFOP$, INCONSISTENCIA$, SUGESTAO$, REGIME_TRIBUTARIO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim ALIQ_PIS As Double
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do relatório inteligente de COFINS/COFINS
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_PIS = Util.ApenasNumeros(Campos(dicTitulos("CST_PIS") - i))
    REGIME_TRIBUTARIO = Util.ApenasNumeros(Campos(dicTitulos("REGIME_TRIBUTARIO") - i))
    ALIQ_PIS = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_PIS") - i))

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
                
            Case CST_PIS Like "7*" And ALIQ_PIS > 0
                INCONSISTENCIA = "CST_PIS indicando operação sem crédito do imposto com campo ALIQ_PIS maior que zero"
                SUGESTAO = "Zerar alíquota do PIS"
                
            Case REGIME_TRIBUTARIO = "1" And ALIQ_PIS > 0 And ALIQ_PIS <> 0.0165
                INCONSISTENCIA = "Empresa no regime não-cumulativo com ALIQ_PIS diferente de 1,65%"
                SUGESTAO = "Informar alíquota de 1,65% para o PIS"
                
            Case REGIME_TRIBUTARIO = "2" And ALIQ_PIS > 0 And ALIQ_PIS <> 0.0065
                INCONSISTENCIA = "Empresa no regime cumulativo com ALIQ_PIS diferente de 0,65%"
                SUGESTAO = "Informar alíquota de 0,65% para o PIS"
                
            Case (CST_PIS > 0 And CST_PIS < 4 Or CST_PIS Like "5*") And ALIQ_PIS = 0
                If REGIME_TRIBUTARIO = "1" Then SUGESTAO = "Informar alíquota de 1,65% para o PIS"
                If REGIME_TRIBUTARIO = "2" Then SUGESTAO = "Informar alíquota de 0,65% para o PIS"
                INCONSISTENCIA = "CST_PIS indicando operação tributada com campo ALIQ_PIS igual a zero"
                SUGESTAO = SUGESTAO
                
            Case CST_PIS > 3 And CST_PIS < 10 And CST_PIS <> 5 And ALIQ_PIS <> 0
                INCONSISTENCIA = "CST_PIS indicando operação não tributada com campo ALIQ_PIS diferente de zero"
                SUGESTAO = "Zerar alíquota do PIS"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
        
End Function

Private Function ValidarCampo_ALIQ_COFINS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim CST_COFINS$, CFOP$, INCONSISTENCIA$, SUGESTAO$, REGIME_TRIBUTARIO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim ALIQ_COFINS As Double
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do relatório inteligente de COFINS/COFINS
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_COFINS = Util.ApenasNumeros(Campos(dicTitulos("CST_COFINS") - i))
    REGIME_TRIBUTARIO = Util.ApenasNumeros(Campos(dicTitulos("REGIME_TRIBUTARIO") - i))
    ALIQ_COFINS = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_COFINS") - i))

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
                
            Case CST_COFINS Like "7*" And ALIQ_COFINS > 0
                INCONSISTENCIA = "CST_COFINS indicando operação sem crédito do imposto com campo ALIQ_COFINS maior que zero"
                SUGESTAO = "Zerar alíquota da COFINS"
                
            Case REGIME_TRIBUTARIO = "1" And ALIQ_COFINS > 0 And ALIQ_COFINS <> 0.076
                INCONSISTENCIA = "Empresa no regime não-cumulativo com ALIQ_COFINS diferente de 7,6%"
                SUGESTAO = "Informar alíquota de 7,60% para a COFINS"
                
            Case REGIME_TRIBUTARIO = "2" And ALIQ_COFINS > 0 And ALIQ_COFINS <> 0.03
                INCONSISTENCIA = "Empresa no regime cumulativo com ALIQ_COFINS diferente de 3,00%"
                SUGESTAO = "Informar alíquota de 3,00% para a COFINS"
                
            Case (CST_COFINS > 0 And CST_COFINS < 4 Or CST_COFINS Like "5*") And ALIQ_COFINS = 0
                If REGIME_TRIBUTARIO = "1" Then SUGESTAO = "Informar alíquota de 7,60% para a COFINS"
                If REGIME_TRIBUTARIO = "2" Then SUGESTAO = "Informar alíquota de 3,00% para a COFINS"
                INCONSISTENCIA = "CST_COFINS indicando operação tributada com campo ALIQ_COFINS igual a zero"
                SUGESTAO = SUGESTAO
                
            Case CST_COFINS > 3 And CST_COFINS < 10 And CST_COFINS <> 5 And ALIQ_COFINS <> 0
                INCONSISTENCIA = "CST_COFINS indicando operação não tributada com campo ALIQ_COFINS diferente de zero"
                SUGESTAO = "Zerar alíquota da COFINS"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
        
End Function

Public Function ValidarCampo_ALIQ_PIS_COFINS_QUANT(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim i As Byte
Dim REG As String
    
    If LBound(Campos) = 0 Then i = 1
    
    Call ValidarCampo_ALIQ_PIS_QUANT(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidarCampo_ALIQ_COFINS_QUANT(Campos, dicTitulos)
        
End Function

Private Function ValidarCampo_ALIQ_PIS_QUANT(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim CST_PIS$, INCONSISTENCIA$, SUGESTAO$
Dim ALIQ_PIS_QUANT As Double
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do relatório inteligente de COFINS/COFINS
    CST_PIS = Util.ApenasNumeros(Campos(dicTitulos("CST_PIS") - i))
    ALIQ_PIS_QUANT = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_PIS_QUANT") - i))

    If INCONSISTENCIA = "" Then
        
        Select Case True
        
            Case CST_PIS Like "*03*" And ALIQ_PIS_QUANT = 0
                INCONSISTENCIA = "Para operações com o CST_PIS = 03 informar o campo ALIQ_PIS_QUANT maior que zero"
                SUGESTAO = "Informar um valor maior que 0 para o campo ALIQ_PIS_QUANT"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
        
End Function

Private Function ValidarCampo_ALIQ_COFINS_QUANT(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim CST_COFINS$, INCONSISTENCIA$, SUGESTAO$
Dim ALIQ_COFINS_QUANT As Double
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do relatório inteligente de COFINS/COFINS
    CST_COFINS = Util.ApenasNumeros(Campos(dicTitulos("CST_COFINS") - i))
    ALIQ_COFINS_QUANT = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_COFINS_QUANT") - i))

    If INCONSISTENCIA = "" Then
        
        Select Case True
        
            Case CST_COFINS Like "*03*" And ALIQ_COFINS_QUANT = 0
                INCONSISTENCIA = "Para operações com o CST_COFINS = 03 informar o campo ALIQ_COFINS_QUANT maior que zero"
                SUGESTAO = "Informar um valor maior que 0 para o campo ALIQ_COFINS_QUANT"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
        
End Function

Public Function ValidarEnumeracao_REGIME_TRIBUTARIO(ByVal REGIME_TRIBUTARIO As String)
    
    Select Case VBA.Val(REGIME_TRIBUTARIO)
            
        Case "1"
            ValidarEnumeracao_REGIME_TRIBUTARIO = "1 - Não-Cumulativo"
            
        Case "2"
            ValidarEnumeracao_REGIME_TRIBUTARIO = "2 - Cumulativo"

        Case Else
            ValidarEnumeracao_REGIME_TRIBUTARIO = REGIME_TRIBUTARIO & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarCampo_COD_NAT_PIS_COFINS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim CST_PIS$, CST_COFINS$, CFOP$, COD_NAT_PIS_COFINS$, INCONSISTENCIA$, SUGESTAO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    COD_NAT_PIS_COFINS = Util.ApenasNumeros(Campos(dicTitulos("COD_NAT_PIS_COFINS") - i))
    CST_PIS = Util.ApenasNumeros(Campos(dicTitulos("CST_PIS") - i))
    CST_COFINS = Util.ApenasNumeros(Campos(dicTitulos("CST_COFINS") - i))
    
    Select Case True
        
        Case CST_PIS > "03" And CST_PIS < "10" And COD_NAT_PIS_COFINS = ""
            INCONSISTENCIA = "O campo COD_NAT_PIS_COFINS deve ser informado para operações entre o CST 04 e 09"
            SUGESTAO = "Informar um valor válido para o campo COD_NAT_PIS_COFINS"
            
        Case CST_COFINS > "03" And CST_COFINS < "10" And COD_NAT_PIS_COFINS = ""
            INCONSISTENCIA = "O campo COD_NAT_COFINS_COFINS deve ser informado para operações entre o CST 04 e 09"
            SUGESTAO = "Informar um valor válido para o campo COD_NAT_COFINS_COFINS"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
        
End Function
