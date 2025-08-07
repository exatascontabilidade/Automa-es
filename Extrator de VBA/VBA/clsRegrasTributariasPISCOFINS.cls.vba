Attribute VB_Name = "clsRegrasTributariasPISCOFINS"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function ValidarCampo_ALIQ_PIS_COFINS(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim i As Byte
Dim REG As String
    
    If LBound(Campos) = 0 Then i = 1
    
    'Identificação do registro
    REG = Campos(dicTitulosApuracao("REG") - i)
    
    If REG <> "D205" Then Call ValidarCampo_ALIQ_PIS(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
    If REG <> "D201" And Campos(dicTitulosApuracao("INCONSISTENCIA") - i) = "" Then Call ValidarCampo_ALIQ_COFINS(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
    
End Function

Private Function ValidarCampo_ALIQ_PIS(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim ALIQ_PIS As Double, ALIQ_PIS_TRIB#
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração PIS/COFINS
    ALIQ_PIS = fnExcel.FormatarPercentuais(Campos(dicTitulosApuracao("ALIQ_PIS") - i))
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)
    
    If LBound(Campos) = 0 Then i = 1
    'Carregar campos do assistente de Tributação PIS/COFINS
    
    ALIQ_PIS_TRIB = fnExcel.FormatarPercentuais(CamposTrib(dicTitulosTributacao("ALIQ_PIS") - i))
    
    Select Case True
        
        Case ALIQ_PIS <> ALIQ_PIS_TRIB
            INCONSISTENCIA = "ALIQ_PIS divergente: " & fnExcel.FormatarPercentuais(ALIQ_PIS) & " (informado) vs " & fnExcel.FormatarPercentuais(ALIQ_PIS_TRIB) & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar a alíquota do PIS cadastrada na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Private Function ValidarCampo_ALIQ_COFINS(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim ALIQ_COFINS As Double, ALIQ_COFINS_TRIB#
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração COFINS/COFINS
    ALIQ_COFINS = fnExcel.FormatarPercentuais(Campos(dicTitulosApuracao("ALIQ_COFINS") - i))
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)
    
    If LBound(Campos) = 0 Then i = 1
    'Carregar campos do assistente de Tributação COFINS/COFINS
    
    ALIQ_COFINS_TRIB = fnExcel.FormatarPercentuais(CamposTrib(dicTitulosTributacao("ALIQ_COFINS") - i))
    
    Select Case True
        
        Case ALIQ_COFINS <> ALIQ_COFINS_TRIB
            INCONSISTENCIA = "ALIQ_COFINS divergente: " & fnExcel.FormatarPercentuais(ALIQ_COFINS) & " (informado) vs " & fnExcel.FormatarPercentuais(ALIQ_COFINS_TRIB) & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar a alíquota da COFINS cadastrada na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_COD_NAT_PIS_COFINS(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim COD_NAT_PIS_COFINS As String, COD_NAT_PIS_COFINS_TRIB$
Dim INCONSISTENCIA As String, SUGESTAO$, DESCR_ITEM$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração PIS/COFINS
    COD_NAT_PIS_COFINS = Campos(dicTitulosApuracao("COD_NAT_PIS_COFINS") - i)
    DESCR_ITEM = Campos(dicTitulosApuracao("DESCR_ITEM") - i)
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Tributação PIS/COFINS
    COD_NAT_PIS_COFINS_TRIB = CamposTrib(dicTitulosTributacao("COD_NAT_PIS_COFINS") - i)
    
    Select Case True
        
        Case COD_NAT_PIS_COFINS <> COD_NAT_PIS_COFINS_TRIB
            INCONSISTENCIA = "COD_NAT_PIS_COFINS divergente: " & COD_NAT_PIS_COFINS & " (informado) vs " & COD_NAT_PIS_COFINS_TRIB & " (cadastrado) para o item: " & DESCR_ITEM
            SUGESTAO = "Aplicar Natureza cadastrada na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_COD_CTA(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim COD_CTA As String, COD_CTA_TRIB$
Dim INCONSISTENCIA As String, SUGESTAO$, DESCR_ITEM$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração PIS/COFINS
    COD_CTA = Campos(dicTitulosApuracao("COD_CTA") - i)
    DESCR_ITEM = Campos(dicTitulosApuracao("DESCR_ITEM") - i)
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Tributação PIS/COFINS
    COD_CTA_TRIB = CamposTrib(dicTitulosTributacao("COD_CTA") - i)
    
    Select Case True
        
        Case COD_CTA <> COD_CTA_TRIB And COD_CTA_TRIB <> ""
            INCONSISTENCIA = "COD_CTA divergente: " & COD_CTA & " (informado) vs " & COD_CTA_TRIB & " (cadastrado) para o item: " & DESCR_ITEM
            SUGESTAO = "Aplicar o código da conta analítica cadastrada na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_CST_PIS_COFINS(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim i As Byte
Dim REG As String
    
    If LBound(Campos) = 0 Then i = 1
    
    'Identificação do registro
    REG = Campos(dicTitulosApuracao("REG") - i)
    
    If REG <> "D205" Then Call ValidarCampo_CST_PIS(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
    If REG <> "D201" And Campos(dicTitulosApuracao("INCONSISTENCIA") - i) = "" Then Call ValidarCampo_CST_COFINS(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
    
End Function

Private Function ValidarCampo_CST_PIS(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim CST_PIS As String, CST_PIS_TRIB$
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração PIS/COFINS
    CST_PIS = Campos(dicTitulosApuracao("CST_PIS") - i)
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)
    
    If LBound(Campos) = 0 Then i = 1
    'Carregar campos do assistente de Tributação PIS/COFINS
    
    CST_PIS_TRIB = CamposTrib(dicTitulosTributacao("CST_PIS") - i)
    
    Select Case True
        
        Case CST_PIS <> CST_PIS_TRIB
            INCONSISTENCIA = "CST_PIS divergente: " & CST_PIS & " (informado) vs " & CST_PIS_TRIB & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar o CST do PIS cadastrada na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Private Function ValidarCampo_CST_COFINS(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim CST_COFINS As String, CST_COFINS_TRIB$
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração PIS/COFINS
    CST_COFINS = Campos(dicTitulosApuracao("CST_COFINS") - i)
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Tributação PIS/COFINS
    CST_COFINS_TRIB = CamposTrib(dicTitulosTributacao("CST_COFINS") - i)
    
    Select Case True
        
        Case CST_COFINS <> CST_COFINS_TRIB
            INCONSISTENCIA = "CST_COFINS divergente: " & CST_COFINS & " (informado) vs " & CST_COFINS_TRIB & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar o CST da COFINS cadastrada na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_ALIQ_PIS_COFINS_QUANT(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim i As Byte
Dim REG As String
    
    If LBound(Campos) = 0 Then i = 1
    
    'Identificação do registro
    REG = Campos(dicTitulosApuracao("REG") - i)
    
    If REG <> "D205" Then Call ValidarCampo_ALIQ_PIS_QUANT(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
    If REG <> "D201" And Campos(dicTitulosApuracao("INCONSISTENCIA") - i) = "" Then Call ValidarCampo_ALIQ_COFINS_QUANT(Campos, dicTitulosApuracao, CamposTrib, dicTitulosTributacao)
    
End Function

Private Function ValidarCampo_ALIQ_PIS_QUANT(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim ALIQ_PIS_QUANT As Double, ALIQ_PIS_QUANT_TRIB#
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração PIS/COFINS
    ALIQ_PIS_QUANT = fnExcel.ConverterValores(Campos(dicTitulosApuracao("ALIQ_PIS_QUANT") - i))
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)

    If LBound(Campos) = 0 Then i = 1
    'Carregar campos do assistente de Tributação PIS/COFINS
    
    ALIQ_PIS_QUANT_TRIB = fnExcel.ConverterValores(CamposTrib(dicTitulosTributacao("ALIQ_PIS_QUANT") - i))
    
    Select Case True
        
        Case ALIQ_PIS_QUANT <> ALIQ_PIS_QUANT_TRIB
            INCONSISTENCIA = "ALIQ_PIS_QUANT divergente: " & fnExcel.FormatarValores(ALIQ_PIS_QUANT) & " (informado) vs " & fnExcel.FormatarValores(ALIQ_PIS_QUANT_TRIB) & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar a alíquota por quantidade do PIS cadastrada na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Private Function ValidarCampo_ALIQ_COFINS_QUANT(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim ALIQ_COFINS_QUANT As Double, ALIQ_COFINS_QUANT_TRIB#
Dim INCONSISTENCIA As String, SUGESTAO$, CFOP$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração PIS/COFINS
    ALIQ_COFINS_QUANT = fnExcel.ConverterValores(Campos(dicTitulosApuracao("ALIQ_COFINS_QUANT") - i))
    CFOP = Campos(dicTitulosApuracao("CFOP") - i)
    
    If LBound(Campos) = 0 Then i = 1
    'Carregar campos do assistente de Tributação PIS/COFINS
    
    ALIQ_COFINS_QUANT_TRIB = fnExcel.ConverterValores(CamposTrib(dicTitulosTributacao("ALIQ_COFINS_QUANT") - i))
    
    Select Case True
        
        Case ALIQ_COFINS_QUANT <> ALIQ_COFINS_QUANT_TRIB
            INCONSISTENCIA = "ALIQ_COFINS divergente: " & fnExcel.FormatarValores(ALIQ_COFINS_QUANT) & " (informado) vs " & fnExcel.FormatarValores(ALIQ_COFINS_QUANT_TRIB) & " (cadastrado) para a operação com CFOP " & CFOP
            SUGESTAO = "Aplicar a alíquota por quantidade da COFINS cadastrada na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

