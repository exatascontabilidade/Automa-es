Attribute VB_Name = "DadosDivergenciasProdutos"
Option Explicit

Public CamposProduto As CamposDivergenciaProdutos
Public dicTitulosProdutos As New Dictionary

Public Type CamposDivergenciaProdutos
    
    REG As String
    ARQUIVO As String
    CHV_PAI As String
    CHV_REG As String
    CHV_NFE As String
    NUM_DOC As String
    SER As String
    NUM_ITEM_NF As Integer
    NUM_ITEM_SPED As Integer
    COD_ITEM_NF As String
    COD_ITEM_SPED As String
    DESCR_ITEM_NF As String
    DESCR_ITEM_SPED As String
    COD_BARRA_NF As String
    COD_BARRA_SPED As String
    COD_NCM_NF As String
    COD_NCM_SPED As String
    EX_IPI_NF As String
    EX_IPI_SPED As String
    CEST_NF As String
    CEST_SPED As String
    QTD_NF As Double
    QTD_SPED As Double
    UNID_NF As String
    UNID_SPED As String
    CFOP_NF As String
    CFOP_SPED As String
    CST_ICMS_NF As String
    CST_ICMS_SPED As String
    VL_ITEM_NF As Double
    VL_ITEM_SPED As Double
    VL_DESC_NF As Double
    VL_DESC_SPED As Double
    VL_BC_ICMS_NF As Double
    VL_BC_ICMS_SPED As Double
    ALIQ_ICMS_NF As Double
    ALIQ_ICMS_SPED As Double
    VL_ICMS_NF As Double
    VL_ICMS_SPED As Double
    VL_BC_ICMS_ST_NF As Double
    VL_BC_ICMS_ST_SPED As Double
    ALIQ_ST_NF As Double
    ALIQ_ST_SPED As Double
    VL_ICMS_ST_NF As Double
    VL_ICMS_ST_SPED As Double
    CST_IPI_NF As String
    CST_IPI_SPED As String
    VL_BC_IPI_NF As Double
    VL_BC_IPI_SPED As Double
    ALIQ_IPI_NF As Double
    ALIQ_IPI_SPED As Double
    VL_IPI_NF As Double
    VL_IPI_SPED As Double
    CST_PIS_NF As String
    CST_PIS_SPED As String
    VL_BC_PIS_NF As Double
    VL_BC_PIS_SPED As Double
    ALIQ_PIS_NF As Double
    ALIQ_PIS_SPED As Double
    QUANT_BC_PIS_NF As Double
    QUANT_BC_PIS_SPED As Double
    ALIQ_PIS_QUANT_NF As Double
    ALIQ_PIS_QUANT_SPED As Double
    VL_PIS_NF As Double
    VL_PIS_SPED As Double
    CST_COFINS_NF As String
    CST_COFINS_SPED As String
    VL_BC_COFINS_NF As Double
    VL_BC_COFINS_SPED As Double
    ALIQ_COFINS_NF As Double
    ALIQ_COFINS_SPED As Double
    QUANT_BC_COFINS_NF As Double
    QUANT_BC_COFINS_SPED As Double
    ALIQ_COFINS_QUANT_NF As Double
    ALIQ_COFINS_QUANT_SPED As Double
    VL_COFINS_NF As Double
    VL_COFINS_SPED As Double
    VL_OPER_NF As Double
    VL_OPER_SPED As Double
    INCONSISTENCIA As String
    SUGESTAO As String
    
End Type

Public Type VerificacoesCamposProdutos
    
    QTD As Boolean
    CEST As Boolean
    UNID As Boolean
    EX_IPI As Boolean
    COD_NCM As Boolean
    NUM_ITEM As Boolean
    COD_BARRA As Boolean
    DESCR_ITEM As Boolean
    
    VL_IPI As Boolean
    VL_ICMS As Boolean
    VL_ITEM As Boolean
    VL_OPER As Boolean
    VL_DESC As Boolean
    VL_ICMS_ST As Boolean
    
    VL_BC_IPI As Boolean
    VL_BC_ICMS As Boolean
    VL_BC_ICMS_ST As Boolean
    
End Type

Public Type DivergenciasQuantidadeItens
    
    QtdDivergente As Boolean
    QtdItensXML As Integer
    QtdItensSPED As Integer
    
End Type

Public Function CarregarDadosRegistroDivergenciaProdutos(ByVal Campos As Variant)

Dim i As Long
    
    If dicTitulosProdutos.Count = 0 Then Call CarregarTitulosDivergenciaProdutos
    Call CarregarTitulosDivergenciaProdutos
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposProduto

        .REG = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("REG") - i))
        .ARQUIVO = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("ARQUIVO") - i))
        .CHV_PAI = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("CHV_PAI_FISCAL") - i))
        .CHV_REG = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("CHV_REG") - i))
        .CHV_NFE = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("CHV_NFE") - i))
        .NUM_DOC = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("NUM_DOC") - i))
        .SER = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("SER") - i))
        .NUM_ITEM_NF = CInt(fnExcel.ConverterValores(Campos(dicTitulosProdutos("NUM_ITEM_NF") - i)))
        .NUM_ITEM_SPED = CInt(fnExcel.ConverterValores(Campos(dicTitulosProdutos("NUM_ITEM_SPED") - i)))
        .COD_ITEM_NF = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("COD_ITEM_NF") - i))
        .COD_ITEM_SPED = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("COD_ITEM_SPED") - i))
        .DESCR_ITEM_NF = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("DESCR_ITEM_NF") - i))
        .DESCR_ITEM_SPED = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("DESCR_ITEM_SPED") - i))
        .COD_BARRA_NF = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("COD_BARRA_NF") - i))
        .COD_BARRA_SPED = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("COD_BARRA_SPED") - i))
        .COD_NCM_NF = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("COD_NCM_NF") - i))
        .COD_NCM_SPED = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("COD_NCM_SPED") - i))
        .EX_IPI_NF = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("EX_IPI_NF") - i))
        .EX_IPI_SPED = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("EX_IPI_SPED") - i))
        .CEST_NF = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("CEST_NF") - i))
        .CEST_SPED = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("CEST_SPED") - i))
        .QTD_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("QTD_NF") - i))
        .QTD_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("QTD_SPED") - i))
        .UNID_NF = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("UNID_NF") - i))
        .UNID_SPED = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("UNID_SPED") - i))
        .CFOP_NF = Campos(dicTitulosProdutos("CFOP_NF") - i)
        .CFOP_SPED = Campos(dicTitulosProdutos("CFOP_SPED") - i)
        .CST_ICMS_NF = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("CST_ICMS_NF") - i))
        .CST_ICMS_SPED = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("CST_ICMS_SPED") - i))
        .VL_ITEM_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_ITEM_NF")), True, 2)
        .VL_ITEM_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_ITEM_SPED")), True, 2)
        .VL_DESC_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_DESC_NF")), True, 2)
        .VL_DESC_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_DESC_SPED")), True, 2)
        .VL_BC_ICMS_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_BC_ICMS_NF")), True, 2)
        .VL_BC_ICMS_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_BC_ICMS_SPED")), True, 2)
        .ALIQ_ICMS_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("ALIQ_ICMS_NF") - i))
        .ALIQ_ICMS_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("ALIQ_ICMS_SPED") - i))
        .VL_ICMS_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_ICMS_NF")), True, 2)
        .VL_ICMS_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_ICMS_SPED")), True, 2)
        .VL_BC_ICMS_ST_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_BC_ICMS_ST_NF")), True, 2)
        .VL_BC_ICMS_ST_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_BC_ICMS_ST_SPED")), True, 2)
        .ALIQ_ST_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("ALIQ_ST_NF") - i))
        .ALIQ_ST_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("ALIQ_ST_SPED") - i))
        .VL_ICMS_ST_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_ICMS_ST_NF")), True, 2)
        .VL_ICMS_ST_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_ICMS_ST_SPED")), True, 2)
        .CST_IPI_NF = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("CST_IPI_NF") - i))
        .CST_IPI_SPED = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("CST_IPI_SPED") - i))
        .VL_BC_IPI_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_BC_IPI_NF")), True, 2)
        .VL_BC_IPI_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_BC_IPI_SPED")), True, 2)
        .ALIQ_IPI_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("ALIQ_IPI_NF") - i))
        .ALIQ_IPI_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("ALIQ_IPI_SPED") - i))
        .VL_IPI_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_IPI_NF")), True, 2)
        .VL_IPI_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_IPI_SPED")), True, 2)
        .CST_PIS_NF = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("CST_PIS_NF") - i))
        .CST_PIS_SPED = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("CST_PIS_SPED") - i))
        .VL_BC_PIS_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_BC_PIS_NF")), True, 2)
        .VL_BC_PIS_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_BC_PIS_SPED")), True, 2)
        .ALIQ_PIS_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("ALIQ_PIS_NF") - i))
        .ALIQ_PIS_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("ALIQ_PIS_SPED") - i))
        .QUANT_BC_PIS_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("QUANT_BC_PIS_NF") - i))
        .QUANT_BC_PIS_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("QUANT_BC_PIS_SPED") - i))
        .ALIQ_PIS_QUANT_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("ALIQ_PIS_QUANT_NF") - i))
        .ALIQ_PIS_QUANT_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("ALIQ_PIS_QUANT_SPED") - i))
        .VL_PIS_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_PIS_NF")), True, 2)
        .VL_PIS_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_PIS_SPED")), True, 2)
        .CST_COFINS_NF = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("CST_COFINS_NF") - i))
        .CST_COFINS_SPED = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("CST_COFINS_SPED") - i))
        .VL_BC_COFINS_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_BC_COFINS_NF")), True, 2)
        .VL_BC_COFINS_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_BC_COFINS_SPED")), True, 2)
        .ALIQ_COFINS_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("ALIQ_COFINS_NF") - i))
        .ALIQ_COFINS_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("ALIQ_COFINS_SPED") - i))
        .QUANT_BC_COFINS_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("QUANT_BC_COFINS_NF") - i))
        .QUANT_BC_COFINS_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("QUANT_BC_COFINS_SPED") - i))
        .ALIQ_COFINS_QUANT_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("ALIQ_COFINS_QUANT_NF") - i))
        .ALIQ_COFINS_QUANT_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("ALIQ_COFINS_QUANT_SPED") - i))
        .VL_COFINS_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_COFINS_NF")), True, 2)
        .VL_COFINS_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_COFINS_SPED")), True, 2)
        .VL_OPER_NF = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_OPER_NF")), True, 2)
        .VL_OPER_SPED = fnExcel.ConverterValores(Campos(dicTitulosProdutos("VL_OPER_SPED")), True, 2)
        .INCONSISTENCIA = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("INCONSISTENCIA") - i))
        .SUGESTAO = Util.RemoverAspaSimples(Campos(dicTitulosProdutos("SUGESTAO") - i))

    End With

End Function

Public Function AtribuirCamposProduto(ByRef Campos As Variant) As Variant

Dim i As Long
    
    If dicTitulosProdutos.Count = 0 Then Call CarregarTitulosDivergenciaProdutos
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposProduto
                
        Campos(dicTitulosProdutos("REG") - i) = .REG
        Campos(dicTitulosProdutos("ARQUIVO") - i) = .ARQUIVO
        Campos(dicTitulosProdutos("CHV_PAI_FISCAL") - i) = .CHV_PAI
        Campos(dicTitulosProdutos("CHV_REG") - i) = .CHV_REG
        Campos(dicTitulosProdutos("CHV_NFE") - i) = .CHV_NFE
        Campos(dicTitulosProdutos("NUM_DOC") - i) = .NUM_DOC
        Campos(dicTitulosProdutos("SER") - i) = .SER
        Campos(dicTitulosProdutos("NUM_ITEM_NF") - i) = .NUM_ITEM_NF
        Campos(dicTitulosProdutos("NUM_ITEM_SPED") - i) = .NUM_ITEM_SPED
        Campos(dicTitulosProdutos("COD_ITEM_NF") - i) = .COD_ITEM_NF
        Campos(dicTitulosProdutos("COD_ITEM_SPED") - i) = .COD_ITEM_SPED
        Campos(dicTitulosProdutos("DESCR_ITEM_NF") - i) = .DESCR_ITEM_NF
        Campos(dicTitulosProdutos("DESCR_ITEM_SPED") - i) = .DESCR_ITEM_SPED
        Campos(dicTitulosProdutos("COD_BARRA_NF") - i) = .COD_BARRA_NF
        Campos(dicTitulosProdutos("COD_BARRA_SPED") - i) = .COD_BARRA_SPED
        Campos(dicTitulosProdutos("COD_NCM_NF") - i) = .COD_NCM_NF
        Campos(dicTitulosProdutos("COD_NCM_SPED") - i) = .COD_NCM_SPED
        Campos(dicTitulosProdutos("EX_IPI_NF") - i) = .EX_IPI_NF
        Campos(dicTitulosProdutos("EX_IPI_SPED") - i) = .EX_IPI_SPED
        Campos(dicTitulosProdutos("CEST_NF") - i) = .CEST_NF
        Campos(dicTitulosProdutos("CEST_SPED") - i) = .CEST_SPED
        Campos(dicTitulosProdutos("QTD_NF") - i) = .QTD_NF
        Campos(dicTitulosProdutos("QTD_SPED") - i) = .QTD_SPED
        Campos(dicTitulosProdutos("UNID_NF") - i) = .UNID_NF
        Campos(dicTitulosProdutos("UNID_SPED") - i) = .UNID_SPED
        Campos(dicTitulosProdutos("CFOP_NF") - i) = .CFOP_NF
        Campos(dicTitulosProdutos("CFOP_SPED") - i) = .CFOP_SPED
        Campos(dicTitulosProdutos("CST_ICMS_NF") - i) = .CST_ICMS_NF
        Campos(dicTitulosProdutos("CST_ICMS_SPED") - i) = .CST_ICMS_SPED
        Campos(dicTitulosProdutos("VL_ITEM_NF") - i) = .VL_ITEM_NF
        Campos(dicTitulosProdutos("VL_ITEM_SPED") - i) = .VL_ITEM_SPED
        Campos(dicTitulosProdutos("VL_DESC_NF") - i) = .VL_DESC_NF
        Campos(dicTitulosProdutos("VL_DESC_SPED") - i) = .VL_DESC_SPED
        Campos(dicTitulosProdutos("VL_BC_ICMS_NF") - i) = .VL_BC_ICMS_NF
        Campos(dicTitulosProdutos("VL_BC_ICMS_SPED") - i) = .VL_BC_ICMS_SPED
        Campos(dicTitulosProdutos("ALIQ_ICMS_NF") - i) = .ALIQ_ICMS_NF
        Campos(dicTitulosProdutos("ALIQ_ICMS_SPED") - i) = .ALIQ_ICMS_SPED
        Campos(dicTitulosProdutos("VL_ICMS_NF") - i) = .VL_ICMS_NF
        Campos(dicTitulosProdutos("VL_ICMS_SPED") - i) = .VL_ICMS_SPED
        Campos(dicTitulosProdutos("VL_BC_ICMS_ST_NF") - i) = .VL_BC_ICMS_ST_NF
        Campos(dicTitulosProdutos("VL_BC_ICMS_ST_SPED") - i) = .VL_BC_ICMS_ST_SPED
        Campos(dicTitulosProdutos("ALIQ_ST_NF") - i) = .ALIQ_ST_NF
        Campos(dicTitulosProdutos("ALIQ_ST_SPED") - i) = .ALIQ_ST_SPED
        Campos(dicTitulosProdutos("VL_ICMS_ST_NF") - i) = .VL_ICMS_ST_NF
        Campos(dicTitulosProdutos("VL_ICMS_ST_SPED") - i) = .VL_ICMS_ST_SPED
        Campos(dicTitulosProdutos("CST_IPI_NF") - i) = .CST_IPI_NF
        Campos(dicTitulosProdutos("CST_IPI_SPED") - i) = .CST_IPI_SPED
        Campos(dicTitulosProdutos("VL_BC_IPI_NF") - i) = .VL_BC_IPI_NF
        Campos(dicTitulosProdutos("VL_BC_IPI_SPED") - i) = .VL_BC_IPI_SPED
        Campos(dicTitulosProdutos("ALIQ_IPI_NF") - i) = .ALIQ_IPI_NF
        Campos(dicTitulosProdutos("ALIQ_IPI_SPED") - i) = .ALIQ_IPI_SPED
        Campos(dicTitulosProdutos("VL_IPI_NF") - i) = .VL_IPI_NF
        Campos(dicTitulosProdutos("VL_IPI_SPED") - i) = .VL_IPI_SPED
        Campos(dicTitulosProdutos("CST_PIS_NF") - i) = .CST_PIS_NF
        Campos(dicTitulosProdutos("CST_PIS_SPED") - i) = .CST_PIS_SPED
        Campos(dicTitulosProdutos("VL_BC_PIS_NF") - i) = .VL_BC_PIS_NF
        Campos(dicTitulosProdutos("VL_BC_PIS_SPED") - i) = .VL_BC_PIS_SPED
        Campos(dicTitulosProdutos("ALIQ_PIS_NF") - i) = .ALIQ_PIS_NF
        Campos(dicTitulosProdutos("ALIQ_PIS_SPED") - i) = .ALIQ_PIS_SPED
        Campos(dicTitulosProdutos("QUANT_BC_PIS_NF") - i) = .QUANT_BC_PIS_NF
        Campos(dicTitulosProdutos("QUANT_BC_PIS_SPED") - i) = .QUANT_BC_PIS_SPED
        Campos(dicTitulosProdutos("ALIQ_PIS_QUANT_NF") - i) = .ALIQ_PIS_QUANT_NF
        Campos(dicTitulosProdutos("ALIQ_PIS_QUANT_SPED") - i) = .ALIQ_PIS_QUANT_SPED
        Campos(dicTitulosProdutos("VL_PIS_NF") - i) = .VL_PIS_NF
        Campos(dicTitulosProdutos("VL_PIS_SPED") - i) = .VL_PIS_SPED
        Campos(dicTitulosProdutos("CST_COFINS_NF") - i) = .CST_COFINS_NF
        Campos(dicTitulosProdutos("CST_COFINS_SPED") - i) = .CST_COFINS_SPED
        Campos(dicTitulosProdutos("VL_BC_COFINS_NF") - i) = .VL_BC_COFINS_NF
        Campos(dicTitulosProdutos("VL_BC_COFINS_SPED") - i) = .VL_BC_COFINS_SPED
        Campos(dicTitulosProdutos("ALIQ_COFINS_NF") - i) = .ALIQ_COFINS_NF
        Campos(dicTitulosProdutos("ALIQ_COFINS_SPED") - i) = .ALIQ_COFINS_SPED
        Campos(dicTitulosProdutos("QUANT_BC_COFINS_NF") - i) = .QUANT_BC_COFINS_NF
        Campos(dicTitulosProdutos("QUANT_BC_COFINS_SPED") - i) = .QUANT_BC_COFINS_SPED
        Campos(dicTitulosProdutos("ALIQ_COFINS_QUANT_NF") - i) = .ALIQ_COFINS_QUANT_NF
        Campos(dicTitulosProdutos("ALIQ_COFINS_QUANT_SPED") - i) = .ALIQ_COFINS_QUANT_SPED
        Campos(dicTitulosProdutos("VL_COFINS_NF") - i) = .VL_COFINS_NF
        Campos(dicTitulosProdutos("VL_COFINS_SPED") - i) = .VL_COFINS_SPED
        Campos(dicTitulosProdutos("VL_OPER_NF") - i) = .VL_OPER_NF
        Campos(dicTitulosProdutos("VL_OPER_SPED") - i) = .VL_OPER_SPED
        Campos(dicTitulosProdutos("INCONSISTENCIA") - i) = .INCONSISTENCIA
        Campos(dicTitulosProdutos("SUGESTAO") - i) = .SUGESTAO
        
        AtribuirCamposProduto = Campos
        
    End With
    
End Function

Private Function CarregarTitulosDivergenciaProdutos()
    
    Set dicTitulosProdutos = Util.MapearTitulos(relDivergenciasProdutos, 3)
    
End Function

Public Function ResetarCamposProduto()
    
    Dim CamposVazios As CamposDivergenciaProdutos
    LSet CamposProduto = CamposVazios
    
End Function

