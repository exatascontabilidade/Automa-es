Attribute VB_Name = "FuncoesJson"
Option Explicit

Public Function CarregarEstruturaSPEDFiscalJson()
    
Dim CustomPart As New clsCustomPartXML
    
    Call dicLayoutFiscal.RemoveAll
    
    Set dicLayoutFiscal = JsonConverter.ParseJson(CustomPart.ExtrairJsonXmlPart("EstruturaSPEDFiscal"))
    
End Function

Public Function CarregarLayoutSPEDFiscal(ByVal versao As String)

Dim CustomPart As New clsCustomPartXML
    
    Call dicLayoutFiscal.RemoveAll
    
    If versao <> "" Then
        
        versao = "_" & versao
        Set dicLayoutFiscal = JsonConverter.ParseJson(CustomPart.ExtrairJsonXmlPart("EstruturaSPEDFiscal" & versao))
        Exit Function
        
    End If
    
End Function

Public Function CarregarLayoutSPEDContribuicoes(ByVal versao As String)

Dim CustomPart As New clsCustomPartXML
    
    Call dicLayoutContribuicoes.RemoveAll
    
    versao = "_" & versao
    Set dicLayoutContribuicoes = JsonConverter.ParseJson(CustomPart.ExtrairJsonXmlPart("EstruturaSPEDContribuicoes" & versao))
    
End Function

Public Function CarregarEstruturaSPEDFiscal(ByRef dicRegistros As Dictionary, ByVal versao As String)
    
Dim CustomPart As New clsCustomPartXML

    Call dicRegistros.RemoveAll
    
    If versao <> "" Then versao = "_" & versao
    Set dicRegistros = JsonConverter.ParseJson(CustomPart.ExtrairJsonXmlPart("EstruturaSPEDFiscal" & versao))
    
End Function

Public Function CarregarEstruturaSPEDContribuicoes(ByRef dicRegistros As Dictionary, ByVal versao As String)
    
Dim CustomPart As New clsCustomPartXML

    Call dicRegistros.RemoveAll
    
    If versao <> "" Then versao = "_" & versao
    Set dicRegistros = JsonConverter.ParseJson(CustomPart.ExtrairJsonXmlPart("EstruturaSPEDContribuicoes" & versao))
    
End Function

Public Function CarregarHierarquiaSPEDFiscal()
    
Dim CustomPart As New clsCustomPartXML

    Call dicHierarquiaSPEDFiscal.RemoveAll
    Set dicHierarquiaSPEDFiscal = JsonConverter.ParseJson(CustomPart.ExtrairJsonXmlPart("HierarquiaRegistrosSPEDFiscal"))
    
End Function

Public Function CarregarMapaChavesSPEDFiscal()
    
Dim CustomPart As New clsCustomPartXML

    Call dicMapaChavesSPEDFiscal.RemoveAll
    Set dicMapaChavesSPEDFiscal = JsonConverter.ParseJson(CustomPart.ExtrairJsonXmlPart("MapaChavesSPEDFiscal"))
    
End Function

