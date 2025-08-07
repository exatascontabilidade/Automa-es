Attribute VB_Name = "DTO_TributacaoNCM_PISCOFINS"
Option Explicit

Public tribNCM_PIS_COFINS As TributacaoNCM_PIS_COFINS

Public Type TributacaoNCM_PIS_COFINS
    
    COD_NCM As String
    EX_IPI As String
    CST_PIS_COFINS_ENT As String
    CST_PIS_COFINS_SAI As String
    ALIQ_PIS As String
    ALIQ_COFINS As String
    COD_NAT_PIS_COFINS As String
    
End Type

Public Function ResetarTributarioNCM_PIS_COFINS()

Dim CamposVazios As TributacaoNCM_PIS_COFINS
    
    LSet tribNCM_PIS_COFINS = CamposVazios
    
End Function

Public Function ObterCampos() As Variant
    
    ObterCampos = Array( _
      "COD_NCM", _
      "EX_IPI", _
      "CST_PIS_COFINS_ENT", _
      "CST_PIS_COFINS_SAI", _
      "ALIQ_PIS", _
      "ALIQ_COFINS", _
      "COD_NAT_PIS_COFINS")
    
End Function

