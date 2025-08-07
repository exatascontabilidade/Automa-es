Attribute VB_Name = "DTO_DocumentosListados"
Option Explicit

Public DocsFiscais As DocumentosFiscais

Type DocumentosFiscais
    
    'Declarações
    arrSPEDs As New ArrayList
    arrSPEDFiscal As New ArrayList
    arrSPEDsInvalidos As New ArrayList
    arrSPEDContribuicoes As New ArrayList
    
    'Documentos
    arrNFeNFCe As New ArrayList
    arrCTe As New ArrayList
    arrCFe As New ArrayList
    arrNFSe As New ArrayList
    arrTodos As New ArrayList
    
    'Protocolos
    arrProtocolos As New ArrayList
    arrCanceladas As New ArrayList
    
    'Outros
    arrDocsInvalidos As New ArrayList
    arrChavesCanceladas As New ArrayList
    
End Type
