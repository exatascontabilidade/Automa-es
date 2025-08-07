Attribute VB_Name = "DadosApuracaoPISCOFINS"
Option Explicit

Public dicTitulosApuracaoPISCOFINS As New Dictionary
Public CamposPISCOFINS As CamposApuracaoPISCOFINS

Public Type CamposApuracaoPISCOFINS
    
    REG As String
    ARQUIVO As String
    CHV_PAI As String
    CHV_REG As String
    REGIME_TRIBUTARIO As String
    CNPJ_ESTABELECIMENTO As String
    COD_MOD As String
    CHV_NFE As String
    NUM_DOC As String
    SER As String
    IND_OPER As String
    DT_DOC As String
    DT_ENT_SAI As String
    COD_PART As String
    TIPO_PART As String
    NOME_RAZAO As String
    COD_ITEM As String
    DESCR_ITEM As String
    COD_BARRA As String
    COD_NCM As String
    EX_IPI As String
    TIPO_ITEM As String
    IND_MOV As String
    UF_PART As String
    CFOP As Integer
    VL_ITEM As Double
    VL_DESP As Double
    VL_DESC As Double
    VL_ICMS As Double
    CST_PIS As String
    CST_COFINS As String
    VL_BC_PIS As Double
    ALIQ_PIS As Double
    QUANT_BC_PIS As Double
    ALIQ_PIS_QUANT As Double
    VL_PIS As Double
    VL_BC_COFINS As Double
    ALIQ_COFINS As Double
    QUANT_BC_COFINS As Double
    ALIQ_COFINS_QUANT As Double
    VL_COFINS As Double
    COD_CTA As String
    COD_NAT_PIS_COFINS As String
    INCONSISTENCIA As String
    SUGESTAO As String
    
End Type

Private Function CarregarTitulosApuracaoPISCOFINS()
    
    Set dicTitulosApuracaoPISCOFINS = Util.MapearTitulos(assApuracaoPISCOFINS, 3)
    
End Function

Public Function CarregarRegistroApuracaoPISCOFINS(ByVal Campos As Variant)

Dim i As Long
    
    If dicTitulosApuracaoPISCOFINS.Count = 0 Then Call CarregarTitulosApuracaoPISCOFINS
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposPISCOFINS
        
        .REG = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("REG") - i))
        .ARQUIVO = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("ARQUIVO") - i))
        .CHV_PAI = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("CHV_PAI_FISCAL") - i))
        .CHV_REG = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("CHV_REG") - i))
        .REGIME_TRIBUTARIO = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("REGIME_TRIBUTARIO") - i))
        .CNPJ_ESTABELECIMENTO = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("CNPJ_ESTABELECIMENTO") - i))
        .COD_MOD = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("COD_MOD") - i))
        .CHV_NFE = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("CHV_NFE") - i))
        .NUM_DOC = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("NUM_DOC") - i))
        .SER = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("SER") - i))
        .IND_OPER = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("IND_OPER") - i))
        .DT_DOC = fnExcel.FormatarData(Campos(dicTitulosApuracaoPISCOFINS("DT_DOC") - i))
        .DT_ENT_SAI = fnExcel.FormatarData(Campos(dicTitulosApuracaoPISCOFINS("DT_ENT_SAI") - i))
        .COD_PART = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("COD_PART") - i))
        .TIPO_PART = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("TIPO_PART") - i))
        .NOME_RAZAO = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("NOME_RAZAO") - i))
        .COD_ITEM = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("COD_ITEM") - i))
        .DESCR_ITEM = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("DESCR_ITEM") - i))
        .COD_BARRA = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("COD_BARRA") - i))
        .COD_NCM = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("COD_NCM") - i))
        .EX_IPI = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("EX_IPI") - i))
        .TIPO_ITEM = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("TIPO_ITEM") - i))
        .IND_MOV = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("IND_MOV") - i))
        .UF_PART = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("UF_PART") - i))
        .CFOP = Campos(dicTitulosApuracaoPISCOFINS("CFOP") - i)
        .VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulosApuracaoPISCOFINS("VL_ITEM") - i), True, 2)
        .VL_DESP = fnExcel.ConverterValores(Campos(dicTitulosApuracaoPISCOFINS("VL_DESP") - i), True, 2)
        .VL_DESC = fnExcel.ConverterValores(Campos(dicTitulosApuracaoPISCOFINS("VL_DESC") - i), True, 2)
        .VL_ICMS = fnExcel.ConverterValores(Campos(dicTitulosApuracaoPISCOFINS("VL_ICMS") - i), True, 2)
        .CST_PIS = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("CST_PIS") - i))
        .CST_COFINS = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("CST_COFINS") - i))
        .VL_BC_PIS = fnExcel.ConverterValores(Campos(dicTitulosApuracaoPISCOFINS("VL_BC_PIS") - i), True, 2)
        .ALIQ_PIS = fnExcel.ConverterValores(Campos(dicTitulosApuracaoPISCOFINS("ALIQ_PIS") - i))
        .QUANT_BC_PIS = fnExcel.ConverterValores(Campos(dicTitulosApuracaoPISCOFINS("QUANT_BC_PIS") - i))
        .ALIQ_PIS_QUANT = fnExcel.ConverterValores(Campos(dicTitulosApuracaoPISCOFINS("ALIQ_PIS_QUANT") - i))
        .VL_PIS = fnExcel.ConverterValores(Campos(dicTitulosApuracaoPISCOFINS("VL_PIS") - i), True, 2)
        .VL_BC_COFINS = fnExcel.ConverterValores(Campos(dicTitulosApuracaoPISCOFINS("VL_BC_COFINS") - i), True, 2)
        .ALIQ_COFINS = fnExcel.ConverterValores(Campos(dicTitulosApuracaoPISCOFINS("ALIQ_COFINS") - i))
        .QUANT_BC_COFINS = fnExcel.ConverterValores(Campos(dicTitulosApuracaoPISCOFINS("QUANT_BC_COFINS") - i))
        .ALIQ_COFINS_QUANT = fnExcel.ConverterValores(Campos(dicTitulosApuracaoPISCOFINS("ALIQ_COFINS_QUANT") - i))
        .VL_COFINS = fnExcel.ConverterValores(Campos(dicTitulosApuracaoPISCOFINS("VL_COFINS") - i), True, 2)
        .COD_CTA = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("COD_CTA") - i))
        .COD_NAT_PIS_COFINS = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("COD_NAT_PIS_COFINS") - i))
        .INCONSISTENCIA = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("INCONSISTENCIA") - i))
        .SUGESTAO = Util.RemoverAspaSimples(Campos(dicTitulosApuracaoPISCOFINS("SUGESTAO") - i))
        
    End With
    
End Function

Public Function AtribuirCamposPISCOFINS(ByRef Campos As Variant) As Variant

Dim i As Long
    
    If dicTitulosApuracaoPISCOFINS.Count = 0 Then Call CarregarTitulosApuracaoPISCOFINS
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposPISCOFINS
        
        Campos(dicTitulosApuracaoPISCOFINS("REG") - i) = .REG
        Campos(dicTitulosApuracaoPISCOFINS("ARQUIVO") - i) = .ARQUIVO
        Campos(dicTitulosApuracaoPISCOFINS("CHV_PAI_FISCAL") - i) = .CHV_PAI
        Campos(dicTitulosApuracaoPISCOFINS("CHV_REG") - i) = .CHV_REG
        Campos(dicTitulosApuracaoPISCOFINS("REGIME_TRIBUTARIO") - i) = .REGIME_TRIBUTARIO
        Campos(dicTitulosApuracaoPISCOFINS("CNPJ_ESTABELECIMENTO") - i) = .CNPJ_ESTABELECIMENTO
        Campos(dicTitulosApuracaoPISCOFINS("COD_MOD") - i) = .COD_MOD
        Campos(dicTitulosApuracaoPISCOFINS("CHV_NFE") - i) = .CHV_NFE
        Campos(dicTitulosApuracaoPISCOFINS("NUM_DOC") - i) = .NUM_DOC
        Campos(dicTitulosApuracaoPISCOFINS("SER") - i) = .SER
        Campos(dicTitulosApuracaoPISCOFINS("IND_OPER") - i) = .IND_OPER
        Campos(dicTitulosApuracaoPISCOFINS("DT_DOC") - i) = .DT_DOC
        Campos(dicTitulosApuracaoPISCOFINS("DT_ENT_SAI") - i) = .DT_ENT_SAI
        Campos(dicTitulosApuracaoPISCOFINS("COD_PART") - i) = .COD_PART
        Campos(dicTitulosApuracaoPISCOFINS("TIPO_PART") - i) = .TIPO_PART
        Campos(dicTitulosApuracaoPISCOFINS("NOME_RAZAO") - i) = .NOME_RAZAO
        Campos(dicTitulosApuracaoPISCOFINS("COD_ITEM") - i) = .COD_ITEM
        Campos(dicTitulosApuracaoPISCOFINS("DESCR_ITEM") - i) = .DESCR_ITEM
        Campos(dicTitulosApuracaoPISCOFINS("COD_BARRA") - i) = .COD_BARRA
        Campos(dicTitulosApuracaoPISCOFINS("COD_NCM") - i) = .COD_NCM
        Campos(dicTitulosApuracaoPISCOFINS("EX_IPI") - i) = .EX_IPI
        Campos(dicTitulosApuracaoPISCOFINS("TIPO_ITEM") - i) = .TIPO_ITEM
        Campos(dicTitulosApuracaoPISCOFINS("IND_MOV") - i) = .IND_MOV
        Campos(dicTitulosApuracaoPISCOFINS("UF_PART") - i) = .UF_PART
        Campos(dicTitulosApuracaoPISCOFINS("CFOP") - i) = .CFOP
        Campos(dicTitulosApuracaoPISCOFINS("VL_ITEM") - i) = .VL_ITEM
        Campos(dicTitulosApuracaoPISCOFINS("VL_DESP") - i) = .VL_DESP
        Campos(dicTitulosApuracaoPISCOFINS("VL_DESC") - i) = .VL_DESC
        Campos(dicTitulosApuracaoPISCOFINS("VL_ICMS") - i) = .VL_ICMS
        Campos(dicTitulosApuracaoPISCOFINS("CST_PIS") - i) = .CST_PIS
        Campos(dicTitulosApuracaoPISCOFINS("CST_COFINS") - i) = .CST_COFINS
        Campos(dicTitulosApuracaoPISCOFINS("VL_BC_PIS") - i) = .VL_BC_PIS
        Campos(dicTitulosApuracaoPISCOFINS("ALIQ_PIS") - i) = .ALIQ_PIS
        Campos(dicTitulosApuracaoPISCOFINS("QUANT_BC_PIS") - i) = .QUANT_BC_PIS
        Campos(dicTitulosApuracaoPISCOFINS("ALIQ_PIS_QUANT") - i) = .ALIQ_PIS_QUANT
        Campos(dicTitulosApuracaoPISCOFINS("VL_PIS") - i) = .VL_PIS
        Campos(dicTitulosApuracaoPISCOFINS("VL_BC_COFINS") - i) = .VL_BC_COFINS
        Campos(dicTitulosApuracaoPISCOFINS("ALIQ_COFINS") - i) = .ALIQ_COFINS
        Campos(dicTitulosApuracaoPISCOFINS("QUANT_BC_COFINS") - i) = .QUANT_BC_COFINS
        Campos(dicTitulosApuracaoPISCOFINS("ALIQ_COFINS_QUANT") - i) = .ALIQ_COFINS_QUANT
        Campos(dicTitulosApuracaoPISCOFINS("VL_COFINS") - i) = .VL_COFINS
        Campos(dicTitulosApuracaoPISCOFINS("COD_CTA") - i) = .COD_CTA
        Campos(dicTitulosApuracaoPISCOFINS("COD_NAT_PIS_COFINS") - i) = .COD_NAT_PIS_COFINS
        Campos(dicTitulosApuracaoPISCOFINS("INCONSISTENCIA") - i) = .INCONSISTENCIA
        Campos(dicTitulosApuracaoPISCOFINS("SUGESTAO") - i) = .SUGESTAO
        
    End With
    
End Function

Public Function ResetarCamposPISCOFINS()
    
    Dim CamposVazios As CamposApuracaoPISCOFINS
    LSet CamposPISCOFINS = CamposVazios
    
End Function

