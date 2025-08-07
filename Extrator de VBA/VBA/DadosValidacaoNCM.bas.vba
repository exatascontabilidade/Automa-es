Attribute VB_Name = "DadosValidacaoNCM"
Option Explicit

Public TabelaNCM As New Dictionary
Public CamposNCM As CamposValidacaoNCM
Public dicTitulosApuracao As New Dictionary
Private CodigosNCM As New clsRegrasFiscaisNCM
Public Type CamposValidacaoNCM
    
    COD_NCM As String
    DT_DOC As String
    DT_REF As String
    IND_OPER As String
    DT_ENT_SAI As String
    DESCRICAO As String
    VIGENCIA_FINAL As String
    VIGENCIA_INICIAL As String
    INCONSISTENCIA As String
    SUGESTAO As String
    
End Type

Private Function CarregarTitulosApuracao(ByRef Plan As Worksheet)
    
    Set dicTitulosApuracao = Util.MapearTitulos(Plan, 3)
    
End Function

Public Function CarregarDadosApuracaoNCM(ByVal Campos As Variant, ByRef Plan As Worksheet)

Dim i As Long
    
    If dicTitulosApuracao.Count = 0 Then Call CarregarTitulosApuracao(Plan)
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    With CamposNCM
        
        .COD_NCM = Util.ApenasNumeros(Campos(dicTitulosApuracao("COD_NCM") - i))
        .IND_OPER = Util.RemoverAspaSimples(Campos(dicTitulosApuracao("IND_OPER") - i))
        .DT_DOC = fnExcel.FormatarData(Campos(dicTitulosApuracao("DT_DOC") - i))
        .DT_ENT_SAI = fnExcel.FormatarData(Campos(dicTitulosApuracao("DT_ENT_SAI") - i))
        
    End With
    
End Function

Public Function CarregarDadosTabelaNCM(ByVal Campos As Variant)
    
    With CamposNCM
        
        .DESCRICAO = Campos(0)
        .VIGENCIA_INICIAL = fnExcel.FormatarData(Campos(1))
        .VIGENCIA_FINAL = fnExcel.FormatarData(Campos(2))
        
    End With
    
End Function

Public Function ResetarCamposNCM()
    
    Dim CamposVazios As CamposValidacaoNCM
    LSet CamposNCM = CamposVazios
    
End Function

Public Function AtualizarTabelaNCM()
    
    Call CodigosNCM.BaixarTabelaNCM
    
End Function
