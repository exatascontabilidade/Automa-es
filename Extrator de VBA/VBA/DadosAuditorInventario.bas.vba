Attribute VB_Name = "DadosAuditorInventario"
Option Explicit

Public CamposInventario As CamposAuditoriaInventario
Public dicTitulosAuditoriaInventario As New Dictionary

Public Type CamposAuditoriaInventario
    
    REG As String
    ARQUIVO As String
    CHV_PAI As String
    CHV_REG As String
    CHV_NFE As String
    COD_MOD As String
    NUM_DOC As String
    SER As String
    IND_OPER As String
    IND_EMIT As String
    COD_SIT As String
    DT_DOC As String
    COD_ITEM As String
    CFOP As String
    TIPO_ITEM As String
    QTD_COM As Double
    QTD_INV As Double
    UNID_COM As String
    UNID_INV As String
    VL_ITEM As Double
    VL_FRT As Double
    VL_SEG As Double
    VL_OUT_DA As Double
    VL_BC_ICMS As Double
    VL_ICMS As Double
    VL_ICMS_ST As Double
    CST_ICMS As String
    VL_IPI As Double
    INCONSISTENCIA As String
    SUGESTAO As String
    
End Type

Public Function CarregarDadosRegistroAuditoriaInventario(ByVal Campos As Variant)

Dim i As Long
    
    If dicTitulosAuditoriaInventario.Count = 0 Then Call CarregarTitulosAuditoriaInventario
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposInventario
        
        .REG = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("REG") - i))
        .ARQUIVO = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("ARQUIVO") - i))
        .CHV_PAI = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("CHV_PAI_FISCAL") - i))
        .CHV_REG = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("CHV_REG") - i))
        .CHV_NFE = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("CHV_NFE") - i))
        .COD_MOD = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("COD_MOD") - i))
        .NUM_DOC = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("NUM_DOC") - i))
        .SER = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("SER") - i))
        .IND_OPER = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("IND_OPER") - i))
        .IND_EMIT = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("IND_EMIT") - i))
        .COD_SIT = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("COD_SIT") - i))
        .DT_DOC = fnExcel.FormatarData(Campos(dicTitulosAuditoriaInventario("DT_DOC") - i))
        .COD_ITEM = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("COD_ITEM") - i))
        .CFOP = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("CFOP") - i))
        .TIPO_ITEM = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("TIPO_ITEM") - i))
        .QTD_COM = fnExcel.ConverterValores(Campos(dicTitulosAuditoriaInventario("QTD_COM")), True, 2)
        .QTD_INV = fnExcel.ConverterValores(Campos(dicTitulosAuditoriaInventario("QTD_INV")), True, 2)
        .UNID_COM = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("UNID_COM") - i))
        .UNID_INV = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("UNID_INV") - i))
        .VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulosAuditoriaInventario("VL_ITEM")), True, 2)
        .VL_FRT = fnExcel.ConverterValores(Campos(dicTitulosAuditoriaInventario("VL_FRT")), True, 2)
        .VL_SEG = fnExcel.ConverterValores(Campos(dicTitulosAuditoriaInventario("VL_SEG")), True, 2)
        .VL_OUT_DA = fnExcel.ConverterValores(Campos(dicTitulosAuditoriaInventario("VL_OUT_DA")), True, 2)
        .VL_BC_ICMS = fnExcel.ConverterValores(Campos(dicTitulosAuditoriaInventario("VL_BC_ICMS")), True, 2)
        .VL_ICMS = fnExcel.ConverterValores(Campos(dicTitulosAuditoriaInventario("VL_ICMS")), True, 2)
        .VL_ICMS_ST = fnExcel.ConverterValores(Campos(dicTitulosAuditoriaInventario("VL_ICMS_ST")), True, 2)
        .CST_ICMS = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("CST_ICMS") - i))
        .VL_IPI = fnExcel.ConverterValores(Campos(dicTitulosAuditoriaInventario("VL_IPI")), True, 2)
        .INCONSISTENCIA = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("INCONSISTENCIA") - i))
        .SUGESTAO = Util.RemoverAspaSimples(Campos(dicTitulosAuditoriaInventario("SUGESTAO") - i))
        
    End With
    
End Function

Public Function AtribuirCamposAuditoriaInventario(ByRef Campos As Variant) As Variant

Dim i As Long
    
    If dicTitulosAuditoriaInventario.Count = 0 Then Call CarregarTitulosAuditoriaInventario
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposInventario
        
        Campos(dicTitulosAuditoriaInventario("REG") - i) = .REG
        Campos(dicTitulosAuditoriaInventario("ARQUIVO") - i) = .ARQUIVO
        Campos(dicTitulosAuditoriaInventario("CHV_PAI_FISCAL") - i) = .CHV_PAI
        Campos(dicTitulosAuditoriaInventario("CHV_REG") - i) = .CHV_REG
        Campos(dicTitulosAuditoriaInventario("CHV_NFE") - i) = .CHV_NFE
        Campos(dicTitulosAuditoriaInventario("COD_MOD") - i) = .COD_MOD
        Campos(dicTitulosAuditoriaInventario("NUM_DOC") - i) = .NUM_DOC
        Campos(dicTitulosAuditoriaInventario("SER") - i) = .SER
        Campos(dicTitulosAuditoriaInventario("IND_OPER") - i) = .IND_OPER
        Campos(dicTitulosAuditoriaInventario("IND_EMIT") - i) = .IND_EMIT
        Campos(dicTitulosAuditoriaInventario("COD_SIT") - i) = .COD_SIT
        Campos(dicTitulosAuditoriaInventario("DT_DOC") - i) = .DT_DOC
        Campos(dicTitulosAuditoriaInventario("COD_ITEM") - i) = .COD_ITEM
        Campos(dicTitulosAuditoriaInventario("CFOP") - i) = .CFOP
        Campos(dicTitulosAuditoriaInventario("TIPO_ITEM") - i) = .TIPO_ITEM
        Campos(dicTitulosAuditoriaInventario("QTD_COM") - i) = .QTD_COM
        Campos(dicTitulosAuditoriaInventario("QTD_INV") - i) = .QTD_INV
        Campos(dicTitulosAuditoriaInventario("UNID_COM") - i) = .UNID_COM
        Campos(dicTitulosAuditoriaInventario("UNID_INV") - i) = .UNID_INV
        Campos(dicTitulosAuditoriaInventario("VL_ITEM") - i) = .VL_ITEM
        Campos(dicTitulosAuditoriaInventario("VL_FRT") - i) = .VL_FRT
        Campos(dicTitulosAuditoriaInventario("VL_SEG") - i) = .VL_SEG
        Campos(dicTitulosAuditoriaInventario("VL_OUT_DA") - i) = .VL_OUT_DA
        Campos(dicTitulosAuditoriaInventario("VL_BC_ICMS") - i) = .VL_BC_ICMS
        Campos(dicTitulosAuditoriaInventario("VL_ICMS") - i) = .VL_ICMS
        Campos(dicTitulosAuditoriaInventario("VL_ICMS_ST") - i) = .VL_ICMS_ST
        Campos(dicTitulosAuditoriaInventario("CST_ICMS") - i) = .CST_ICMS
        Campos(dicTitulosAuditoriaInventario("VL_IPI") - i) = .VL_IPI
        Campos(dicTitulosAuditoriaInventario("INCONSISTENCIA") - i) = .INCONSISTENCIA
        Campos(dicTitulosAuditoriaInventario("SUGESTAO") - i) = .SUGESTAO
        
        AtribuirCamposAuditoriaInventario = Campos
        
    End With
    
End Function

Private Function CarregarTitulosAuditoriaInventario()
    
    Set dicTitulosAuditoriaInventario = Util.MapearTitulos(relAuditoriaInventario, 3)
    
End Function

Public Function ResetarCamposAuditoriaInventario()
    
    Dim CamposVazios As CamposAuditoriaInventario
    LSet CamposInventario = CamposVazios
    
End Function
