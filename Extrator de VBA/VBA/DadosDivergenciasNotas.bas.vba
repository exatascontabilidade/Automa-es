Attribute VB_Name = "DadosDivergenciasNotas"
Option Explicit

Public CamposNota As CamposDivergenciaNotas
Public dicTitulosNotas As New Dictionary

Public Type CamposDivergenciaNotas
    
    REG As String
    ARQUIVO As String
    CHV_PAI As String
    CHV_REG As String
    CHV_NFE As String
    COD_MOD_NF As String
    COD_MOD_SPED As String
    NUM_DOC_NF As String
    NUM_DOC_SPED As String
    SER_NF As String
    SER_SPED As String
    IND_OPER_NF As String
    IND_OPER_SPED As String
    IND_EMIT_NF As String
    IND_EMIT_SPED As String
    COD_SIT_NF As String
    COD_SIT_SPED As String
    COD_PART_NF As String
    COD_PART_SPED As String
    NOME_RAZAO_NF As String
    NOME_RAZAO_SPED As String
    INSC_EST_NF As String
    INSC_EST_SPED As String
    DT_DOC_NF As String
    DT_DOC_SPED As String
    DT_E_S_NF As String
    DT_E_S_SPED As String
    IND_PGTO_NF As String
    IND_PGTO_SPED As String
    VL_DOC_NF As Double
    VL_DOC_SPED As Double
    VL_DESC_NF As Double
    VL_DESC_SPED As Double
    VL_ABAT_NT_NF As Double
    VL_ABAT_NT_SPED As Double
    VL_MERC_NF As Double
    VL_MERC_SPED As Double
    IND_FRT_NF As String
    IND_FRT_SPED As String
    VL_FRT_NF As Double
    VL_FRT_SPED As Double
    VL_SEG_NF As Double
    VL_SEG_SPED As Double
    VL_OUT_DA_NF As Double
    VL_OUT_DA_SPED As Double
    VL_BC_ICMS_NF As Double
    VL_BC_ICMS_SPED As Double
    VL_ICMS_NF As Double
    VL_ICMS_SPED As Double
    VL_BC_ICMS_ST_NF As Double
    VL_BC_ICMS_ST_SPED As Double
    VL_ICMS_ST_NF As Double
    VL_ICMS_ST_SPED As Double
    VL_IPI_NF As Double
    VL_IPI_SPED As Double
    VL_PIS_NF As Double
    VL_PIS_SPED As Double
    VL_COFINS_NF As Double
    VL_COFINS_SPED As Double
    VL_PIS_ST_NF As Double
    VL_PIS_ST_SPED As Double
    VL_COFINS_ST_NF As Double
    VL_COFINS_ST_SPED As Double
    INCONSISTENCIA As String
    SUGESTAO As String
    
End Type

Public Type VerificacoesCamposNotas
    
    VL_IPI As Boolean
    VL_ICMS As Boolean
    VL_DOC As Boolean
    VL_MERC As Boolean
    VL_DESC As Boolean
    VL_ICMS_ST As Boolean
    VL_BC_IPI As Boolean
    VL_BC_ICMS As Boolean
    VL_BC_ICMS_ST As Boolean
    
End Type

Public Function CarregarDadosRegistroDivergenciaNotas(ByVal Campos As Variant)

Dim i As Long
    
    If dicTitulosNotas.Count = 0 Then Call CarregarTitulosDivergenciaNotas
    Call CarregarTitulosDivergenciaNotas
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposNota
        
        .REG = Util.RemoverAspaSimples(Campos(dicTitulosNotas("REG") - i))
        .ARQUIVO = Util.RemoverAspaSimples(Campos(dicTitulosNotas("ARQUIVO") - i))
        .CHV_PAI = Util.RemoverAspaSimples(Campos(dicTitulosNotas("CHV_PAI_FISCAL") - i))
        .CHV_REG = Util.RemoverAspaSimples(Campos(dicTitulosNotas("CHV_REG") - i))
        .CHV_NFE = Util.RemoverAspaSimples(Campos(dicTitulosNotas("CHV_NFE") - i))
        .COD_MOD_NF = Util.RemoverAspaSimples(Campos(dicTitulosNotas("COD_MOD_NF") - i))
        .COD_MOD_SPED = Util.RemoverAspaSimples(Campos(dicTitulosNotas("COD_MOD_SPED") - i))
        .NUM_DOC_NF = Util.RemoverAspaSimples(Campos(dicTitulosNotas("NUM_DOC_NF") - i))
        .NUM_DOC_SPED = Util.RemoverAspaSimples(Campos(dicTitulosNotas("NUM_DOC_SPED") - i))
        .SER_NF = VBA.Format(Util.RemoverAspaSimples(Campos(dicTitulosNotas("SER_NF") - i)), "000")
        .SER_SPED = VBA.Format(Util.RemoverAspaSimples(Campos(dicTitulosNotas("SER_SPED") - i)), "000")
        .IND_OPER_NF = Util.RemoverAspaSimples(Campos(dicTitulosNotas("IND_OPER_NF") - i))
        .IND_OPER_SPED = Util.RemoverAspaSimples(Campos(dicTitulosNotas("IND_OPER_SPED") - i))
        .IND_EMIT_NF = Util.RemoverAspaSimples(Campos(dicTitulosNotas("IND_EMIT_NF") - i))
        .IND_EMIT_SPED = Util.RemoverAspaSimples(Campos(dicTitulosNotas("IND_EMIT_SPED") - i))
        .COD_SIT_NF = Util.RemoverAspaSimples(Campos(dicTitulosNotas("COD_SIT_NF") - i))
        .COD_SIT_SPED = Util.RemoverAspaSimples(Campos(dicTitulosNotas("COD_SIT_SPED") - i))
        .COD_PART_NF = Util.RemoverAspaSimples(Campos(dicTitulosNotas("COD_PART_NF") - i))
        .COD_PART_SPED = Util.RemoverAspaSimples(Campos(dicTitulosNotas("COD_PART_SPED") - i))
        .NOME_RAZAO_NF = Util.RemoverAspaSimples(Campos(dicTitulosNotas("NOME_RAZAO_NF") - i))
        .NOME_RAZAO_SPED = Util.RemoverAspaSimples(Campos(dicTitulosNotas("NOME_RAZAO_SPED") - i))
        .INSC_EST_NF = Util.RemoverAspaSimples(Campos(dicTitulosNotas("INSC_EST_NF") - i))
        .INSC_EST_SPED = Util.RemoverAspaSimples(Campos(dicTitulosNotas("INSC_EST_SPED") - i))
        .DT_DOC_NF = fnExcel.FormatarData(Campos(dicTitulosNotas("DT_DOC_NF") - i))
        .DT_DOC_SPED = fnExcel.FormatarData(Campos(dicTitulosNotas("DT_DOC_SPED") - i))
        .DT_E_S_NF = fnExcel.FormatarData(Campos(dicTitulosNotas("DT_E_S_NF") - i))
        .DT_E_S_SPED = fnExcel.FormatarData(Campos(dicTitulosNotas("DT_E_S_SPED") - i))
        .IND_PGTO_NF = Util.RemoverAspaSimples(Campos(dicTitulosNotas("IND_PGTO_NF") - i))
        .IND_PGTO_SPED = Util.RemoverAspaSimples(Campos(dicTitulosNotas("IND_PGTO_SPED") - i))
        .VL_DOC_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_DOC_NF")), True, 2)
        .VL_DOC_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_DOC_SPED")), True, 2)
        .VL_DESC_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_DESC_NF")), True, 2)
        .VL_DESC_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_DESC_SPED")), True, 2)
        .VL_ABAT_NT_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_ABAT_NT_NF")), True, 2)
        .VL_ABAT_NT_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_ABAT_NT_SPED")), True, 2)
        .VL_MERC_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_MERC_NF")), True, 2)
        .VL_MERC_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_MERC_SPED")), True, 2)
        .IND_FRT_NF = Util.RemoverAspaSimples(Campos(dicTitulosNotas("IND_FRT_NF") - i))
        .IND_FRT_SPED = Util.RemoverAspaSimples(Campos(dicTitulosNotas("IND_FRT_SPED") - i))
        .VL_FRT_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_FRT_NF")), True, 2)
        .VL_FRT_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_FRT_SPED")), True, 2)
        .VL_SEG_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_SEG_NF")), True, 2)
        .VL_SEG_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_SEG_SPED")), True, 2)
        .VL_OUT_DA_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_OUT_DA_NF")), True, 2)
        .VL_OUT_DA_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_OUT_DA_SPED")), True, 2)
        .VL_BC_ICMS_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_BC_ICMS_NF")), True, 2)
        .VL_BC_ICMS_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_BC_ICMS_SPED")), True, 2)
        .VL_ICMS_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_ICMS_NF")), True, 2)
        .VL_ICMS_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_ICMS_SPED")), True, 2)
        .VL_BC_ICMS_ST_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_BC_ICMS_ST_NF")), True, 2)
        .VL_BC_ICMS_ST_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_BC_ICMS_ST_SPED")), True, 2)
        .VL_ICMS_ST_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_ICMS_ST_NF")), True, 2)
        .VL_ICMS_ST_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_ICMS_ST_SPED")), True, 2)
        .VL_IPI_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_IPI_NF")), True, 2)
        .VL_IPI_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_IPI_SPED")), True, 2)
        .VL_PIS_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_PIS_NF")), True, 2)
        .VL_PIS_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_PIS_SPED")), True, 2)
        .VL_COFINS_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_COFINS_NF")), True, 2)
        .VL_COFINS_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_COFINS_SPED")), True, 2)
        .VL_PIS_ST_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_PIS_ST_NF")), True, 2)
        .VL_PIS_ST_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_PIS_ST_SPED")), True, 2)
        .VL_COFINS_ST_NF = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_COFINS_ST_NF")), True, 2)
        .VL_COFINS_ST_SPED = fnExcel.ConverterValores(Campos(dicTitulosNotas("VL_COFINS_ST_SPED")), True, 2)
        .INCONSISTENCIA = Util.RemoverAspaSimples(Campos(dicTitulosNotas("INCONSISTENCIA") - i))
        .SUGESTAO = Util.RemoverAspaSimples(Campos(dicTitulosNotas("SUGESTAO") - i))
        
    End With
    
End Function

Public Function AtribuirCamposNota(ByRef Campos As Variant) As Variant

Dim i As Long
    
    If dicTitulosNotas.Count = 0 Then Call CarregarTitulosDivergenciaNotas
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposNota
        
        Campos(dicTitulosNotas("REG") - i) = .REG
        Campos(dicTitulosNotas("ARQUIVO") - i) = .ARQUIVO
        Campos(dicTitulosNotas("CHV_PAI_FISCAL") - i) = .CHV_PAI
        Campos(dicTitulosNotas("CHV_REG") - i) = .CHV_REG
        Campos(dicTitulosNotas("CHV_NFE") - i) = .CHV_NFE
        Campos(dicTitulosNotas("COD_MOD_NF") - i) = .COD_MOD_NF
        Campos(dicTitulosNotas("COD_MOD_SPED") - i) = .COD_MOD_SPED
        Campos(dicTitulosNotas("NUM_DOC_NF") - i) = .NUM_DOC_NF
        Campos(dicTitulosNotas("NUM_DOC_SPED") - i) = .NUM_DOC_SPED
        Campos(dicTitulosNotas("SER_NF") - i) = .SER_NF
        Campos(dicTitulosNotas("SER_SPED") - i) = .SER_SPED
        Campos(dicTitulosNotas("IND_OPER_NF") - i) = .IND_OPER_NF
        Campos(dicTitulosNotas("IND_OPER_SPED") - i) = .IND_OPER_SPED
        Campos(dicTitulosNotas("IND_EMIT_NF") - i) = .IND_EMIT_NF
        Campos(dicTitulosNotas("IND_EMIT_SPED") - i) = .IND_EMIT_SPED
        Campos(dicTitulosNotas("COD_SIT_NF") - i) = .COD_SIT_NF
        Campos(dicTitulosNotas("COD_SIT_SPED") - i) = .COD_SIT_SPED
        Campos(dicTitulosNotas("COD_PART_NF") - i) = .COD_PART_NF
        Campos(dicTitulosNotas("COD_PART_SPED") - i) = .COD_PART_SPED
        Campos(dicTitulosNotas("NOME_RAZAO_NF") - i) = .NOME_RAZAO_NF
        Campos(dicTitulosNotas("NOME_RAZAO_SPED") - i) = .NOME_RAZAO_SPED
        Campos(dicTitulosNotas("INSC_EST_NF") - i) = .INSC_EST_NF
        Campos(dicTitulosNotas("INSC_EST_SPED") - i) = .INSC_EST_SPED
        Campos(dicTitulosNotas("DT_DOC_NF") - i) = .DT_DOC_NF
        Campos(dicTitulosNotas("DT_DOC_SPED") - i) = .DT_DOC_SPED
        Campos(dicTitulosNotas("DT_E_S_NF") - i) = .DT_E_S_NF
        Campos(dicTitulosNotas("DT_E_S_SPED") - i) = .DT_E_S_SPED
        Campos(dicTitulosNotas("IND_PGTO_NF") - i) = .IND_PGTO_NF
        Campos(dicTitulosNotas("IND_PGTO_SPED") - i) = .IND_PGTO_SPED
        Campos(dicTitulosNotas("VL_DOC_NF") - i) = .VL_DOC_NF
        Campos(dicTitulosNotas("VL_DOC_SPED") - i) = .VL_DOC_SPED
        Campos(dicTitulosNotas("VL_DESC_NF") - i) = .VL_DESC_NF
        Campos(dicTitulosNotas("VL_DESC_SPED") - i) = .VL_DESC_SPED
        Campos(dicTitulosNotas("VL_ABAT_NT_NF") - i) = .VL_ABAT_NT_NF
        Campos(dicTitulosNotas("VL_ABAT_NT_SPED") - i) = .VL_ABAT_NT_SPED
        Campos(dicTitulosNotas("VL_MERC_NF") - i) = .VL_MERC_NF
        Campos(dicTitulosNotas("VL_MERC_SPED") - i) = .VL_MERC_SPED
        Campos(dicTitulosNotas("IND_FRT_NF") - i) = .IND_FRT_NF
        Campos(dicTitulosNotas("IND_FRT_SPED") - i) = .IND_FRT_SPED
        Campos(dicTitulosNotas("VL_FRT_NF") - i) = .VL_FRT_NF
        Campos(dicTitulosNotas("VL_FRT_SPED") - i) = .VL_FRT_SPED
        Campos(dicTitulosNotas("VL_SEG_NF") - i) = .VL_SEG_NF
        Campos(dicTitulosNotas("VL_SEG_SPED") - i) = .VL_SEG_SPED
        Campos(dicTitulosNotas("VL_OUT_DA_NF") - i) = .VL_OUT_DA_NF
        Campos(dicTitulosNotas("VL_OUT_DA_SPED") - i) = .VL_OUT_DA_SPED
        Campos(dicTitulosNotas("VL_BC_ICMS_NF") - i) = .VL_BC_ICMS_NF
        Campos(dicTitulosNotas("VL_BC_ICMS_SPED") - i) = .VL_BC_ICMS_SPED
        Campos(dicTitulosNotas("VL_ICMS_NF") - i) = .VL_ICMS_NF
        Campos(dicTitulosNotas("VL_ICMS_SPED") - i) = .VL_ICMS_SPED
        Campos(dicTitulosNotas("VL_BC_ICMS_ST_NF") - i) = .VL_BC_ICMS_ST_NF
        Campos(dicTitulosNotas("VL_BC_ICMS_ST_SPED") - i) = .VL_BC_ICMS_ST_SPED
        Campos(dicTitulosNotas("VL_ICMS_ST_NF") - i) = .VL_ICMS_ST_NF
        Campos(dicTitulosNotas("VL_ICMS_ST_SPED") - i) = .VL_ICMS_ST_SPED
        Campos(dicTitulosNotas("VL_IPI_NF") - i) = .VL_IPI_NF
        Campos(dicTitulosNotas("VL_IPI_SPED") - i) = .VL_IPI_SPED
        Campos(dicTitulosNotas("VL_PIS_NF") - i) = .VL_PIS_NF
        Campos(dicTitulosNotas("VL_PIS_SPED") - i) = .VL_PIS_SPED
        Campos(dicTitulosNotas("VL_COFINS_NF") - i) = .VL_COFINS_NF
        Campos(dicTitulosNotas("VL_COFINS_SPED") - i) = .VL_COFINS_SPED
        Campos(dicTitulosNotas("VL_PIS_ST_NF") - i) = .VL_PIS_ST_NF
        Campos(dicTitulosNotas("VL_PIS_ST_SPED") - i) = .VL_PIS_ST_SPED
        Campos(dicTitulosNotas("VL_COFINS_ST_NF") - i) = .VL_COFINS_ST_NF
        Campos(dicTitulosNotas("VL_COFINS_ST_SPED") - i) = .VL_COFINS_ST_SPED
        Campos(dicTitulosNotas("INCONSISTENCIA") - i) = .INCONSISTENCIA
        Campos(dicTitulosNotas("SUGESTAO") - i) = .SUGESTAO
        
        AtribuirCamposNota = Campos
        
    End With
    
End Function

Private Function CarregarTitulosDivergenciaNotas()
    
    Set dicTitulosNotas = Util.MapearTitulos(relDivergenciasNotas, 3)
    
End Function

Public Function ResetarCamposNota()
    
    Dim CamposVazios As CamposDivergenciaNotas
    LSet CamposNota = CamposVazios
    
End Function
