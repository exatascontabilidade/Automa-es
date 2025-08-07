Attribute VB_Name = "clsRegrasCadastraisICMS"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private ValidacoesCFOP As New clsRegrasFiscaisCFOP

Public Function ValidarCampo_CST_ICMS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim CST_ICMS$, CFOP$, INCONSISTENCIA$, SUGESTAO$
Dim ALIQ_ICMS As Double
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do relatório de tributação ICMS
    CST_ICMS = Util.ApenasNumeros(Campos(dicTitulos("CST_ICMS") - i))
    ALIQ_ICMS = Util.ApenasNumeros(Campos(dicTitulos("ALIQ_ICMS") - i))
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    
    If INCONSISTENCIA = "" Then
        
        Select Case True
            
            Case CFOP Like "#551" And Not CST_ICMS Like "#90"
                INCONSISTENCIA = "CST_ICMS incompatível com operação de aquisição de ativo imobilizado"
                SUGESTAO = "Informar CST_ICMS " & VBA.Left(CST_ICMS, 1) & 90 & "para a operação"
                
            Case CFOP Like "#556" And Not CST_ICMS Like "#90"
                INCONSISTENCIA = "CST_ICMS incompatível com operação de aquisição para uso e consumo"
                SUGESTAO = "Informar CST_ICMS " & VBA.Left(CST_ICMS, 1) & 90 & "para a operação"
                
            Case CFOP Like "#406" And Not CST_ICMS Like "#60"
                INCONSISTENCIA = "CST_ICMS incompatível com operação de aquisição de ativo imobilizado com ST"
                SUGESTAO = "Informar CST_ICMS " & VBA.Left(CST_ICMS, 1) & 60 & "para a operação"
                
            Case CFOP Like "#407" And Not CST_ICMS Like "#60"
                INCONSISTENCIA = "CST_ICMS incompatível com operação de aquisição para uso e consumo com ST"
                SUGESTAO = "Informar CST_ICMS " & VBA.Left(CST_ICMS, 1) & 60 & "para a operação"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
    
End Function

Public Function ValidarCampo_CFOP(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim UF_CONTRIB As String, UF_PART$, CFOP$, INCONSISTENCIA$, SUGESTAO$
Dim ALIQ_ICMS As Double
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do relatório de tributação ICMS
    UF_CONTRIB = VBA.Left(Campos(dicTitulos("UF_CONTRIB") - i), 2)
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    UF_PART = VBA.Left(Campos(dicTitulos("UF_PART") - i), 2)
    
    If INCONSISTENCIA = "" Then
        
        Select Case True
            
            Case (CFOP > 2000 And CFOP < 3000) And UF_CONTRIB = UF_PART
                INCONSISTENCIA = "CFOP (" & CFOP & ") incompatível com a operação (UF do Participante [" & UF_PART & "] igual da UF do Contribuinte [" & UF_CONTRIB & "])"
                SUGESTAO = "Informe um CFOP começando com o dígito 1"
                
            Case CFOP < 2000 And UF_CONTRIB <> UF_PART
                INCONSISTENCIA = "CFOP (" & CFOP & ") incompatível com a operação (UF do Participante [" & UF_PART & "] diferente a UF do Contribuinte [" & UF_CONTRIB & "])"
                SUGESTAO = "Informe um CFOP começando com o dígito 2"
                
            Case (CFOP > 4000 And CFOP < 6000) And UF_CONTRIB <> UF_PART
                INCONSISTENCIA = "CFOP (" & CFOP & ") incompatível com a operação (UF do Participante [" & UF_PART & "] diferente a UF do Contribuinte [" & UF_CONTRIB & "])"
                SUGESTAO = "Informe um CFOP começando com o dígito 5"
                
            Case (CFOP > 6000 And CFOP < 7000) And UF_CONTRIB = UF_PART
                INCONSISTENCIA = "CFOP (" & CFOP & ") incompatível com a operação (UF do Participante [" & UF_PART & "] igual da UF do Contribuinte [" & UF_CONTRIB & "])"
                SUGESTAO = "Informe um CFOP começando com o dígito 6"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
    
End Function

Public Function ValidarCampo_ALIQ_ICMS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim CST_ICMS$, CFOP$, INCONSISTENCIA$, SUGESTAO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim ALIQ_ICMS As Double
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do relatório inteligente de COFINS/COFINS
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_ICMS = Util.ApenasNumeros(Campos(dicTitulos("CST_ICMS") - i))
    ALIQ_ICMS = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_ICMS") - i))
    
    'Verificações de dados com CFOP
    If CFOP <> "" Then
        
        fatCFOP = ValidacoesCFOP.ValidarCFOPFaturamento(CFOP)
        devCompCFOP = ValidacoesCFOP.ValidarCFOPDevolucaoCompra(CFOP)
        
    End If
    
    If CFOP <> "" Then
        
        Select Case True
            
            Case CFOP > 4000 And Not fatCFOP And ALIQ_ICMS > 0
                INCONSISTENCIA = "CFOP informado não indica receita operacional com ALIQ_ICMS maior que zero"
                SUGESTAO = "Zerar valores do PIS"
                
        End Select
        
    End If
    
    If INCONSISTENCIA = "" Then
        
        Select Case True
            
            Case CST_ICMS Like "*00" And ALIQ_ICMS = 0
                INCONSISTENCIA = "CST_ICMS indicando operação tributada integralmente com campo ALIQ_ICMS igual a zero"
                SUGESTAO = "Informar uma alíquota maior que 0 para o campo ALIQ_ICMS"
                
            Case (CST_ICMS Like "*10" Or CST_ICMS Like "*60" Or CST_ICMS Like "*70") And ALIQ_ICMS > 0
                INCONSISTENCIA = "CST_ICMS indicando operação sem crédito do imposto com campo ALIQ_ICMS maior que zero"
                SUGESTAO = "Zerar alíquota do ICMS"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
        
End Function

Public Function ValidarCampo_ALIQ_ST(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim CST_ICMS$, CFOP$, INCONSISTENCIA$, SUGESTAO$
Dim fatCFOP As Boolean, devCompCFOP As Boolean
Dim ALIQ_ST As Double
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do relatório inteligente de COFINS/COFINS
    CFOP = Util.ApenasNumeros(Campos(dicTitulos("CFOP") - i))
    CST_ICMS = Util.ApenasNumeros(Campos(dicTitulos("CST_ICMS") - i))
    ALIQ_ST = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_ST") - i))

    'Verificações de dados com CFOP
    If CFOP <> "" Then
    
        fatCFOP = ValidacoesCFOP.ValidarCFOPFaturamento(CFOP)
        devCompCFOP = ValidacoesCFOP.ValidarCFOPDevolucaoCompra(CFOP)
    
    End If
    
    If CFOP <> "" Then
        
        Select Case True
            
            Case CFOP > 4000 And Not fatCFOP And ALIQ_ST > 0
                INCONSISTENCIA = "CFOP informado não indica receita operacional com ALIQ_ST maior que zero"
                SUGESTAO = "Zerar valores da COFINS"
        
        End Select
        
    End If
    
    If INCONSISTENCIA = "" Then
        
        Select Case True
        
            Case CST_ICMS Like "*05*" And ALIQ_ST = 0
                INCONSISTENCIA = "Para operações com o CST_ICMS = 05 informar o campo ALIQ_ST maior que zero"
                SUGESTAO = "Informar uma alíquota maior que 0 para o campo ALIQ_ST"
                
            Case CST_ICMS Like "7*" And ALIQ_ST > 0
                INCONSISTENCIA = "CST_ICMS indicando operação sem crédito do imposto com campo ALIQ_ST maior que zero"
                SUGESTAO = "Zerar alíquota da COFINS"
                
            Case CST_ICMS > 3 And CST_ICMS < 10 And CST_ICMS <> 5 And ALIQ_ST <> 0
                INCONSISTENCIA = "CST_ICMS indicando operação não tributada com campo ALIQ_ST diferente de zero"
                SUGESTAO = "Zerar alíquota da COFINS"
                
        End Select
        
    End If
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO)
        
End Function
