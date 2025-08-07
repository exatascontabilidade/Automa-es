Attribute VB_Name = "DTO_DadosXML"
Option Explicit

Public DadosXML As InformacoesXML

Public Type InformacoesXML
    
    Periodo As String
    ARQUIVO As String
    CNPJ_EMITENTE As String
    CNPJ_DESTINATARIO As String
    CNPJ_ESTABELECIMENTO As String
    TIPO_NF As String
    TIPO_EMISSAO As String
    
End Type

Public Function ResetarDadosXML()
    
    Dim CamposVazios As InformacoesXML
    LSet DadosXML = CamposVazios
    
End Function
