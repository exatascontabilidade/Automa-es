Attribute VB_Name = "ProcessarDocumentos"
Option Explicit

Public ClassificadorSPED As New AssistenteClassificacaoSPEDs
Public ClassificadorXML As New AssistenteClassificacaoXMLs

Public Function LimparListaDocumentos()
    
    With DocsFiscais
        
        'Declarações
        Set .arrSPEDs = New ArrayList
        Set .arrSPEDFiscal = New ArrayList
        Set .arrSPEDsInvalidos = New ArrayList
        Set .arrSPEDContribuicoes = New ArrayList
        
        'Documentos
        Set .arrNFeNFCe = New ArrayList
        Set .arrCTe = New ArrayList
        Set .arrCFe = New ArrayList
        Set .arrNFSe = New ArrayList
        Set .arrTodos = New ArrayList
        
        'Protocolos
        Set .arrProtocolos = New ArrayList
        Set .arrCanceladas = New ArrayList
        
        'Outros
        Set .arrDocsInvalidos = New ArrayList
        Set .arrChavesCanceladas = New ArrayList
        
    End With
    
End Function

Public Sub ListarTodosDocumentos()
    
    With DocsFiscais
        
        .arrTodos.addRange .arrNFeNFCe
        .arrTodos.addRange .arrCTe
        .arrTodos.addRange .arrCFe
        .arrTodos.addRange .arrNFSe
        
    End With
    
End Sub

Public Sub ListarTodosSPEDs()
    
    With DocsFiscais
        
        .arrSPEDs.addRange .arrSPEDFiscal
        .arrSPEDs.addRange .arrSPEDContribuicoes
        
    End With
    
End Sub

Public Function ResetarDadosDocumentosFiscais()
    
    Dim CamposVazios As DocumentosFiscais
    LSet DocsFiscais = CamposVazios
    
End Function

Public Function CarregarXMLeSPED(ByVal TipoImportacao As String) As Boolean

Dim caminhoXml As String
Dim CaminhoSPED As String
Dim XMLSelecionados As Variant
Dim SPEDSelecionados As Variant

    Call LimparListaDocumentos
    
    With DocsFiscais
        
        Select Case TipoImportacao
            
            Case "Arquivo"
                SPEDSelecionados = Util.SelecionarArquivos("txt", "Selecione os SPEDs que deseja importar")
                XMLSelecionados = Util.SelecionarArquivos("xml", "Selecione os XMLS que deseja importar")
                Inicio = Now()
                
                Call ClassificadorSPED.ClassificarSPEDS(SPEDSelecionados)
                Call ClassificadorXML.ClassificarXMLs(XMLSelecionados)
                
            Case "Lote"
                caminhoXml = Util.SelecionarPasta("Selecione a pasta que contém os arquivos XML")
                If caminhoXml = "" Then Exit Function
                
                CaminhoSPED = Util.SelecionarPasta("Selecione a pasta que contém os arquivos SPED")
                If CaminhoSPED = "" Then Exit Function
                
                Inicio = Now()
                Call ClassificadorXML.ListarArquivosXML(caminhoXml)
                Call ClassificadorSPED.ListarArquivosSPED(CaminhoSPED)
                
        End Select
        
        Call fnXML.CarregarProtocolosCancelamento(.arrProtocolos, .arrChavesCanceladas, "Verificando protocolos de cancelamento - ")
        CarregarXMLeSPED = True
        
    End With
    
End Function

Public Function CarregarXMLS(ByVal TipoImportacao As String) As Boolean

Dim Caminho As String
Dim Result As Boolean
    
    Call Util.DesabilitarControles
    
    Call LimparListaDocumentos
    
    With DocsFiscais
        
        Select Case TipoImportacao
            
            Case "Arquivo"
                Result = Util.GuardarEnderecosArrayList("xml", .arrTodos)
                If Not Result Then GoTo Finalizar:
                
            Case "Lote"
                Caminho = Util.SelecionarPasta("Selecione a pasta que contém os arquivos XML")
                If Caminho = "" Then GoTo Finalizar:
                
                Inicio = Now()
                Call ClassificadorXML.ListarArquivosXML(Caminho)
                CarregarXMLS = True
                
        End Select
        
        Call fnXML.CarregarProtocolosCancelamento(.arrProtocolos, .arrChavesCanceladas, "Verificando protocolos de cancelamento - ")
        
    End With
    
Finalizar:
    Call Util.HabilitarControles
    
End Function

Public Sub CarregarSPEDs(ByVal TipoImportacao As String)

Dim Caminho As String
Dim ArquivosSelecionados As Variant
    
    Call LimparListaDocumentos
    
    With DocsFiscais
        
        Select Case TipoImportacao
            
            Case "Arquivo"
                ArquivosSelecionados = Util.SelecionarArquivos("txt", "Selecione os SPEDs que deseja importar")
                Inicio = Now()
                
                Call ClassificadorSPED.ClassificarSPEDS(ArquivosSelecionados)
                
            Case "Lote"
                Caminho = Util.SelecionarPasta("Selecione a pasta que contém os arquivos SPED")
                If Caminho = "" Then Exit Sub
                
                Inicio = Now()
                Call ClassificadorSPED.ListarArquivosSPED(Caminho)
                
        End Select
                
    End With
    
End Sub

Public Function PossuiSPEDFiscalListado() As Boolean

    If DocsFiscais.arrSPEDFiscal.Count > 0 Then
        
        PossuiSPEDFiscalListado = True
        
    Else
        
        Call MensagemAlertaSPEDFiscalNaoDetectado
        
    End If
    
End Function

Public Function PossuiSPEDContribuicoesListado() As Boolean
    
    If DocsFiscais.arrSPEDContribuicoes.Count > 0 Then
        
        PossuiSPEDContribuicoesListado = True
    
    Else
    
        Call MensagemAlertaSPEDContribuicoesNaoDetectado
    
    End If
    
End Function

Private Sub MensagemAlertaSPEDFiscalNaoDetectado()

Dim Msg As String

    Msg = "Nenhum SPED Fiscal válido detectado." & vbCrLf
    Msg = Msg & "Por favor verifique os arquivos e diretório selecionados e tente novamente."
    
    Call Util.MsgAlerta(Msg, "SPED Fiscal não detectado")
    
End Sub

Private Sub MensagemAlertaSPEDContribuicoesNaoDetectado()

Dim Msg As String

    Msg = "Nenhum SPED Contribuições válido detectado." & vbCrLf
    Msg = Msg & "Por favor verifique os arquivos e diretório selecionados e tente novamente."
    
    Call Util.MsgAlerta(Msg, "SPED Contribuições não detectado")
    
End Sub
