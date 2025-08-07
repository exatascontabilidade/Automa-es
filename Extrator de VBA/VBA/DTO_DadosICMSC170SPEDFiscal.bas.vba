Attribute VB_Name = "DTO_DadosICMSC170SPEDFiscal"
Option Explicit

Public infoICMSC170 As InformacoesICMSC170SPEDFiscal

Public Type InformacoesICMSC170SPEDFiscal
    
    CHV_NFE As String
    NUM_ITEM As Integer
    COD_ITEM As String
    CFOP As String
    CST_ICMS As String
    VL_BC_ICMS As Double
    ALIQ_ICMS As Double
    VL_ICMS As Double
    
End Type

Public Function ResetarDadosICMSC170()
    
    Dim CamposVazios As InformacoesICMSC170SPEDFiscal
    LSet infoICMSC170 = CamposVazios
    
End Function

Public Function MontarArrayInfoICMSC170()

Dim arrCampos As New ArrayList

    With infoICMSC170
        
        arrCampos.Add .CFOP
        arrCampos.Add .CST_ICMS
        arrCampos.Add .VL_BC_ICMS
        arrCampos.Add .ALIQ_ICMS
        arrCampos.Add .VL_ICMS
        
    End With
    
    MontarArrayInfoICMSC170 = arrCampos.toArray()
    Set arrCampos = Nothing
    
End Function
