Attribute VB_Name = "DTO_Correlacoes"
Option Explicit

Public Correlacionamento As RegistrosCorrelacoes

Type RegistrosCorrelacoes
    
    dicCorrelacoes As New Dictionary
    dicTitulosCorrelacoes As New Dictionary
    
End Type

Public Sub CarregarCorrelacionamentos()
    
    With Correlacionamento
        
        Call .dicCorrelacoes.RemoveAll
        Call .dicTitulosCorrelacoes.RemoveAll

        Set .dicCorrelacoes = Util.CriarDicionarioCorrelacoes(Correlacoes)
        Set .dicTitulosCorrelacoes = Util.MapearTitulos(Correlacoes, 3)
    
    End With
    
End Sub

Public Function ResetarDadosCorrelacionamento()
    
    Dim CamposVazios As RegistrosCorrelacoes
    LSet Correlacionamento = CamposVazios
    
End Function
