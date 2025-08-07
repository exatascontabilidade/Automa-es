Attribute VB_Name = "FuncoesScripts"
Option Explicit

Public Sub ImportarJson(ByRef control As IRibbonControl)

Dim CustomPart As New clsCustomPartXML
Dim Dados As String, NomeObjeto$
Dim Arqs As Variant, Arq
    
    On Error Resume Next
    Arqs = Util.SelecionarArquivos("json", "Selecione os arquivos que deseja importar")
    If VarType(Arqs) = vbBoolean Then Exit Sub
    
    Inicio = Now()
    For Each Arq In Arqs
        
        NomeObjeto = VBA.Replace(VBA.Right(Arq, VBA.Len(Arq) - VBA.InStrRev(Arq, "\")), ".json", "")
        Dados = Util.DefinirCodificacao(Arq, "UTF-8")
        
        Call CustomPart.SalvarJson(NomeObjeto, Dados)
        
    Next Arq
    
    Call Util.MsgInformativa("Os arquivos foram importados com sucesso!", "Importação da dados Json", Inicio)
    
End Sub

Public Sub ImportarTabela(ByRef control As IRibbonControl)

Dim CustomPart As New clsCustomPartXML
Dim Dados As String, NomeObjeto$
Dim Arqs As Variant, Arq
    
    On Error Resume Next
    Arqs = Util.SelecionarArquivos("txt", "Selecione os arquivos que deseja importar")
    If VarType(Arqs) = vbBoolean Then Exit Sub
    
    Inicio = Now()
    For Each Arq In Arqs
        
        NomeObjeto = VBA.Replace(VBA.Right(Arq, VBA.Len(Arq) - VBA.InStrRev(Arq, "\")), ".txt", "")
        Dados = Util.DefinirCodificacao(Arq, "windows-1252")
        
        Call CustomPart.SalvarTabelaTXT(NomeObjeto, Dados)
        
    Next Arq
    
    Call Util.MsgInformativa("Os arquivos foram importados com sucesso!", "Importação da dados Json", Inicio)
    
End Sub

Public Function ExportarJson(ByRef control As IRibbonControl)

Dim XmlPart As CustomXMLPart
Dim CustomPart As New clsCustomPartXML
Dim Caminho As String, Destino$, Namespace$, Conteudo$
    
    Caminho = Util.SelecionarPasta("Selecione a pasta para salvar os arquivos") & "\"
    If Caminho = "\" Then Exit Function
    
    Inicio = Now()
    For Each XmlPart In ThisWorkbook.CustomXMLParts
        
        Namespace = XmlPart.NamespaceURI
        Conteudo = CustomPart.ExtrairJsonXmlPart(Namespace)
        Destino = Caminho & Namespace & ".json"
        
        If isJson(Conteudo) Then Call Util.ExportarTxt(Destino, Conteudo)
        
    Next XmlPart
    
    Call Util.MsgInformativa("Os arquivos foram exportados com sucesso!", "Exportação da dados Json", Inicio)
    
End Function

Public Function ExportarTabela(ByRef control As IRibbonControl)

Dim XmlPart As CustomXMLPart
Dim CustomPart As New clsCustomPartXML
Dim Caminho As String, Destino$, Namespace$, Conteudo$
    
    Caminho = Util.SelecionarPasta("Selecione a pasta para salvar os arquivos") & "\"
    If Caminho = "\" Then Exit Function
    
    Inicio = Now()
    For Each XmlPart In ThisWorkbook.CustomXMLParts
        
        Namespace = XmlPart.NamespaceURI
        Conteudo = CustomPart.ExtrairTXTPartXML(Namespace)
        Destino = Caminho & Namespace & ".txt"
        
        If isTXT(Conteudo) Then Call Util.ExportarTxt(Destino, Conteudo)
        
    Next XmlPart
    
    Call Util.MsgInformativa("Os arquivos foram exportados com sucesso!", "Exportação da dados Tabela", Inicio)
    
End Function

Public Function isTXT(ByVal Conteudo As String) As Boolean

Dim ConteudoTratado As String
    
    ConteudoTratado = Trim(Conteudo)
    isTXT = (Not ConteudoTratado Like "{*") And (Not ConteudoTratado Like "[[]*") And (ConteudoTratado Like "*|*") And (Not ConteudoTratado = "")
    
End Function

Public Function isJson(ByVal Conteudo As String) As Boolean

Dim ConteudoTratado As String
    
    ConteudoTratado = Trim(Conteudo)
    isJson = (ConteudoTratado Like "{*") Or (ConteudoTratado Like "[[]*") And (Not ConteudoTratado = "")

End Function
