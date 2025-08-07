Attribute VB_Name = "clsCorrelacionamentoSPEDXML"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Sub GerarCorrelacoesSPEDXML()

Dim Seguir As Boolean
Dim arrRelatorio As New ArrayList
Dim dicOperacoesXML As New Dictionary
Dim dicProdutosXML As New Dictionary
Dim dicCorrelacoes As New Dictionary
Dim dicProdutosSPED As New Dictionary
Dim dicOperacoesSPED As New Dictionary
    
    Seguir = ProcessarDocumentos.CarregarXMLeSPED("Lote")
    If Not Seguir Then Exit Sub
    
    Call ProcessarDocumentos.ListarTodosSPEDs
    
    With DocsFiscais
        
        Call ImportarDadosXML(.arrNFeNFCe, dicOperacoesXML, dicProdutosXML, dicCorrelacoes)
        Call ImportarDadosSPED(.arrSPEDs, dicProdutosSPED, dicOperacoesSPED)
        Call DefinirCorrelacoes(dicOperacoesXML, dicOperacoesSPED, dicProdutosXML, dicProdutosSPED, dicCorrelacoes, arrRelatorio)
        
        Call Util.LimparDados(Correlacoes, 4, False)
        Call Util.ExportarDadosDicionario(Correlacoes, dicCorrelacoes)
        
        Call Util.LimparDados(relCorrelacoes, 4, False)
        Call Util.ExportarDadosArrayList(relCorrelacoes, arrRelatorio)
        
        Call FuncoesFormatacao.DestacarMelhorCorrelacao(relCorrelacoes)
        
        Call Util.MsgInformativa("Correlações efetuadas com sucesso!", "Correlacionamento de Itens XML x SPED", Inicio)
        
    End With
    
End Sub

Public Function ImportarDadosXML(ByVal XMLS As Variant, ByRef dicOperacoesXML As Dictionary, _
    ByRef dicProdutosXML As Dictionary, ByRef dicCorrelacoes As Dictionary)
    
Dim XML As Variant
Dim NITEM As Integer
Dim NFe As New DOMDocument60
Dim Produto As IXMLDOMNode
Dim Produtos As IXMLDOMNodeList
Dim vProd As Double, vDesc#, vIPI#, vICMSST#, QTD#, vOperacao#
Dim chNFe As String, cProd$, xProd$, uCom$, codBarra$, NCM$, exTIPI$, CEST$, Chave$, CNPJEmit$, CNPJDest$, Razao$
    
    For Each XML In XMLS
        
        Set NFe = fnXML.RemoverNamespaces(XML)
        
        CNPJBase = VBA.Left(CNPJContribuinte, 8)
        CNPJDest = fnXML.ExtrairCNPJDestinatario(NFe)
        If fnXML.ValidarXMLNFe(NFe) And CNPJDest Like "*" & CNPJBase & "*" Then
            
            'Dados do fonecedor
            CNPJEmit = fnXML.ExtrairCNPJEmitente(NFe)
            If CNPJEmit = CNPJContribuinte Then GoTo Prx:
            
            Razao = fnXML.ValidarTag(NFe, "//emit/xNome")
            
            chNFe = fnXML.ExtrairChaveAcessoNFe(NFe)
            If Not dicOperacoesXML.Exists(chNFe) Then Set dicOperacoesXML(chNFe) = New Dictionary
            
            Set Produtos = NFe.SelectNodes("//det")
            For Each Produto In Produtos
                
                'Dados do Produto
                NITEM = fnXML.ValidarValores(Produto, "@nItem")
                cProd = fnXML.ValidarTag(Produto, "prod/cProd")
                xProd = fnXML.ValidarTag(Produto, "prod/xProd")
                codBarra = fnXML.ExtrairCodigoBarrasProduto(Produto)
                uCom = VBA.UCase(ValidarTag(Produto, "prod/uCom"))
                NCM = ValidarTag(Produto, "prod/NCM")
                exTIPI = ValidarTag(Produto, "prod/EXTIPI")
                CEST = ValidarTag(Produto, "prod/CEST")
                
                'Dados da Operação
                QTD = fnXML.ValidarValores(Produto, "prod/qCom")
                vProd = fnXML.ValidarValores(Produto, "prod/vProd")
                vDesc = fnXML.ValidarValores(Produto, "prod/vDesc")
                vICMSST = fnXML.ValidarValores(Produto, "imposto/ICMS//vICMSST") + fnXML.ValidarValores(Produto, "imposto/ICMS//vFCPST")
                vIPI = fnXML.ValidarValores(Produto, "imposto/IPI//vIPI")
                vOperacao = vProd + vICMSST + vIPI - vDesc
                
                'Armazena dados do produto
                If Not dicProdutosXML.Exists(cProd) Then _
                    dicProdutosXML(cProd) = Array(xProd, codBarra, uCom, NCM, exTIPI, CEST)
                    
                If Not dicOperacoesXML(chNFe).Exists(NITEM) Then _
                    dicOperacoesXML(chNFe)(NITEM) = Array(CNPJEmit, NITEM, cProd, QTD, uCom, vProd, vDesc, vICMSST, vIPI, vOperacao)
                    
                'Armazena dados da correlação
                Chave = CNPJEmit & cProd & uCom
                If Not dicCorrelacoes.Exists(Chave) Then _
                    dicCorrelacoes(Chave) = Array("'" & CNPJEmit, Razao, "'" & cProd, xProd, "'" & uCom, "", "", "", "")
                    
            Next Produto
            
        End If
Prx:
    Next XML
    
End Function

Public Sub ImportarDadosSPED(ByVal SPEDS As Variant, ByRef dicProdutosSPED As Dictionary, ByRef dicOperacoesSPED As Dictionary)

Dim dicParticipantes As New Dictionary
Dim SPED As Variant, Registros, Registro, Campos, CamposDic
Dim vProd As Double, vDesc#, vIPI#, vICMSST#, QTD#, vOperacao#
Dim chNFe As String, cProd$, xProd$, uCom$, codBarra$, NCM$, exTIPI$, CEST$, Chave$, uInv$, indOper$, DocEmit$, codPart$
Dim NITEM As Integer
Dim tipo As String
Dim nReg As String
    
    On Error GoTo Tratar:
    
    a = 0
    For Each SPED In SPEDS
        
        Registros = Util.ImportarTxt(SPED)
        For Each Registro In Registros
            
            Call Util.AntiTravamento(a, 100, "Importando registros SPED, por favor aguarde...")
            
            nReg = Mid(Registro, 2, 4)
            Select Case True
                
                Case nReg = "0150"
                    Campos = Split(Registro, "|")
                    codPart = Campos(2)
                    DocEmit = Campos(5) & Campos(6)
                    dicParticipantes(codPart) = DocEmit
                    
                Case nReg = "0200"
                    Campos = Split(Registro, "|")
                    cProd = Campos(2)
                    xProd = Campos(3)
                    codBarra = Campos(4)
                    uInv = VBA.UCase(Campos(6))
                    NCM = Campos(8)
                    exTIPI = Campos(9)
                    CEST = Campos(13)
                    
                    'Armazena dados do produto
                    If Not dicProdutosSPED.Exists(cProd) Then _
                        dicProdutosSPED(cProd) = Array(xProd, codBarra, uInv, NCM, exTIPI, CEST)
                    
                Case nReg = "C100"
                    Campos = Split(Registro, "|")
                    indOper = Campos(2)
                    
                    If indOper = "0" Then
                    
                        chNFe = Campos(9)
                        DocEmit = dicParticipantes(Campos(4))
                        If Not dicOperacoesSPED.Exists(chNFe) And indOper = "0" Then Set dicOperacoesSPED(chNFe) = New Dictionary
                    
                    Else
                    
                        chNFe = ""
                    
                    End If
                    
                Case nReg = "C170"
                    If chNFe <> "" Then
                    
                        Campos = Split(Registro, "|")
                        NITEM = Campos(2)
                        
                        cProd = Campos(3)
                        QTD = fnExcel.FormatarValores(Campos(5))
                        uCom = Campos(6)
                        vProd = fnExcel.FormatarValores(Campos(7))
                        vDesc = fnExcel.FormatarValores(Campos(8))
                        vICMSST = fnExcel.FormatarValores(Campos(18))
                        vIPI = fnExcel.FormatarValores(Campos(24))
                        vOperacao = vProd + vICMSST + vIPI - vDesc
                        
                        If dicProdutosSPED.Exists(cProd) Then _
                            dicOperacoesSPED(chNFe)(NITEM) = Array(DocEmit, NITEM, cProd, QTD, uCom, vProd, vDesc, vICMSST, vIPI, vOperacao)
                        
                    End If
                
                Case nReg > "C197"
                    Exit For
                    
            End Select
            
        Next Registro
        
    Next SPED
    
Exit Sub
Tratar:
    Stop
    Resume
    
End Sub

Public Function DefinirCorrelacoes(ByRef dicOperacoesXML As Dictionary, ByRef dicOperacoesSPED As Dictionary, _
    ByRef dicProdutosXML As Dictionary, ByRef dicProdutosSPED As Dictionary, ByRef dicCorrelacoes As Dictionary, _
    ByRef arrRelatorio As ArrayList)
        
Dim Operacao As Variant, Titulos, CamposXML, CamposSPED, nItemXML, nItemSPED, Campos
Dim Pontuacao As Double, MaiorPontuacao As Double
Dim arrItensCorrelacionados As New ArrayList
Dim dicTitulosDivergencias As New Dictionary
Dim dicTitulosCorrelacoes As New Dictionary
Dim dicMelhoresCorrelacoes As New Dictionary
Dim dicTitulosOperacoes As New Dictionary
Dim dicTitulosProdutos As New Dictionary
Dim dicMaiorPontuacao As New Dictionary
Dim dicCorrelacoesXMLSPED As New Dictionary
Dim dicOperacoesAuxiliar As New Dictionary
Dim dicPontuacao As New Dictionary
Dim dicItensSPED As New Dictionary
Dim dicOperacoes As New Dictionary
Dim dicItensXML As New Dictionary
Dim MelhorCorrelacao As Integer, QTD&
    
    Titulos = Array("xProd", "codBarra", "uCom", "NCM", "exTIPI", "CEST")
    Set dicTitulosProdutos = Util.MapearArray(Titulos)
    
    Titulos = Array("CNPJEmit", "nItem", "cProd", "Qtd", "uCom", "vProd", "vDesc", "vICMSST", "vIPI", "vOperacao")
    Set dicTitulosOperacoes = Util.MapearArray(Titulos)
    
    Set dicTitulosCorrelacoes = Util.MapearTitulos(Correlacoes, 3)
    Set dicTitulosDivergencias = Util.MapearTitulos(relCorrelacoes, 3)
    
    'Processa as operações dos arquivos XML
    For Each Operacao In dicOperacoesXML.Keys()
        
        'Pula operação caso a chave esteja vazia
        If Operacao = "" Then GoTo Prx:
        
        'Apaga a lista de itens correlacionados para a nota fiscal atual
        arrItensCorrelacionados.Clear
        
        'Cria registro para as correlações XML / SPED
        If Not dicCorrelacoesXMLSPED.Exists(Operacao) Then Set dicCorrelacoesXMLSPED(Operacao) = New Dictionary
                
        'Verifica se a operação também existe no dicionário do SPED
        If dicOperacoesSPED.Exists(Operacao) Then
                        
            'Carrega os itens das operações no XML e SPED
            Set dicItensXML = dicOperacoesXML(Operacao)
            Set dicItensSPED = dicOperacoesSPED(Operacao)
            
            'Processa todos os itens da operação do XML
            For Each nItemXML In dicItensXML.Keys()
                
                'Cria dicionário para registrar as correlações do item do XML com os itens do SPED
                If Not dicCorrelacoesXMLSPED(Operacao).Exists(nItemXML) Then Set dicCorrelacoesXMLSPED(Operacao)(nItemXML) = New Dictionary
                
                'limpa as variáveis de pontuação e Correlação para o item da operação
                MelhorCorrelacao = 0
                MaiorPontuacao = 0
                
                'Processa todos os itens da operação do SPED
                For Each nItemSPED In dicItensSPED.Keys()
                    
                    'Verifica se o item do SPED já foi correlacionado com o XML
                    If Not arrItensCorrelacionados.contains(nItemSPED) Then
                        
                        'Coleta os dados do item do SPED e XML
                        CamposXML = dicItensXML(nItemXML)
                        CamposSPED = dicItensSPED(nItemSPED)
                        
                        'Calcula a pontuação para a correlação
                        Pontuacao = CruzarDadosOperacao(CamposXML, CamposSPED, dicTitulosOperacoes, dicProdutosXML, dicProdutosSPED, dicTitulosProdutos)
                        
                        'Registra a correlação feita entre o item do XML e o item do SPED atual
                        Campos = RegistrarPontuacao(dicProdutosXML, dicProdutosSPED, dicTitulosProdutos, CamposXML, CamposSPED, Operacao, Pontuacao)
                        
                        If VarType(Campos) = 8204 Then _
                            dicCorrelacoesXMLSPED(Operacao)(nItemXML)(nItemSPED) = Campos
                        
                        If Pontuacao > MaiorPontuacao Then
        
                            MaiorPontuacao = Pontuacao
                            MelhorCorrelacao = nItemSPED
        
                        End If
                        
                    End If
                    
                Next nItemSPED
                
                'Adiciona a melhor correlação em uma lista
                arrItensCorrelacionados.Add MelhorCorrelacao
                
                'Cria dicionário para registrar as melhores correlações para cada item do XML
                If Not dicMelhoresCorrelacoes.Exists(Operacao) Then Set dicMelhoresCorrelacoes(Operacao) = New Dictionary
                
                'Marca a correlação como a melhor encontrada no dicionário de correlações entre XML e SPED
                If VarType(dicCorrelacoesXMLSPED(Operacao)(nItemXML)(MelhorCorrelacao)) = 8204 Then
                    dicCorrelacoesXMLSPED(Operacao)(nItemXML)(MelhorCorrelacao) = IdentificarMelhorCorrelacao(dicCorrelacoesXMLSPED(Operacao)(nItemXML)(MelhorCorrelacao))
                End If
                
                'Cria um dicionário para armazenar as melhores correlações da operação
                If Not dicMelhoresCorrelacoes(Operacao).Exists(MelhorCorrelacao) Then
                    Set dicMelhoresCorrelacoes(Operacao)(MelhorCorrelacao) = New Dictionary
                End If
                    
                'Registra a melhor correlação encontrada entre os itens do SPED para o item atual do XML no dicionário de Melhor Correlação
                dicMelhoresCorrelacoes(Operacao)(MelhorCorrelacao)(nItemXML) = dicCorrelacoesXMLSPED(Operacao)(nItemXML)(MelhorCorrelacao)
                
            Next nItemXML
            
            Call ChecarCorrelacoes(dicCorrelacoesXMLSPED(Operacao), dicMelhoresCorrelacoes(Operacao), dicTitulosDivergencias, dicCorrelacoes, dicTitulosCorrelacoes, arrRelatorio)

        End If
        
Prx:
    Next Operacao
    
End Function

Public Function CruzarDadosOperacao(ByRef CamposXML As Variant, ByRef CamposSPED As Variant, ByRef dicTitulos As Dictionary, _
    ByRef dicProdutosXML As Dictionary, ByRef dicProdutosSPED As Dictionary, ByRef dicTitulosProdutos As Dictionary) As Double
    
Dim Pontuacao As Double, Posicao As Integer
Dim cProdXML As String, cProdSPED$
Dim Titulo As Variant
    
    Pontuacao = 0
    For Each Titulo In dicTitulos.Keys()
        
        Posicao = dicTitulos(Titulo)
        Select Case Titulo
            
            Case "vOperacao"
                If CamposXML(Posicao) = CamposSPED(Posicao) Then Pontuacao = Pontuacao + 5
                
            Case "Qtd", "uCom", "vProd", "vDesc"
                If CamposXML(Posicao) = CamposSPED(Posicao) Then Pontuacao = Pontuacao + 1
                
            Case "cProd"
                cProdXML = CamposXML(Posicao)
                cProdSPED = CamposSPED(Posicao)
                
                Pontuacao = Pontuacao + CruzarDadosProdutos(dicProdutosXML, dicProdutosSPED, cProdXML, cProdSPED, dicTitulosProdutos)
                
        End Select
        
    Next Titulo
    
    CruzarDadosOperacao = Pontuacao
    
End Function

Public Function CruzarDadosProdutos(ByRef dicProdutosXML As Dictionary, ByRef dicProdutosSPED As Dictionary, _
    ByRef cProdXML As String, ByRef cProdSPED As String, ByRef dicTitulosProdutos As Dictionary) As Double

Dim CamposXML As Variant, CamposSPED, Titulos, Titulo
Dim Pontuacao As Double, Posicao As Integer
Dim xProdXML As String, xProdSPED$

    CamposXML = dicProdutosXML(cProdXML)
    CamposSPED = dicProdutosSPED(cProdSPED)
        
    Pontuacao = 0
    For Each Titulo In dicTitulosProdutos.Keys()
        
        Posicao = dicTitulosProdutos(Titulo)
        Select Case Titulo
            
            Case "codBarra", "NCM", "exTIPI", "CEST"
                If CamposXML(Posicao) & CamposSPED(Posicao) <> "" Then _
                    If CamposXML(Posicao) = CamposSPED(Posicao) Then Pontuacao = Pontuacao + 1
                
            Case "xProd"
                xProdXML = VBA.LCase(CamposXML(Posicao))
                xProdSPED = VBA.LCase(CamposSPED(Posicao))
                Pontuacao = Pontuacao + CompararDescricoes(xProdXML, xProdSPED)
                
        End Select
                
    Next Titulo
    
    CruzarDadosProdutos = Pontuacao

End Function

Public Function RegistrarPontuacao(ByRef dicProdutosXML As Dictionary, ByRef dicProdutosSPED As Dictionary, _
    ByRef dicTitulosProdutos As Dictionary, ByRef CamposXML As Variant, ByRef CamposSPED As Variant, _
    ByVal Operacao As String, ByVal Pontuacao As Double)

Dim Campos As Variant, Titulo
Dim arrCampos As New ArrayList
Dim cProdXML As String, cProdSPED$, Chave$
Dim arrDivergencias As New ArrayList
Dim i As Byte, j As Byte
        
    'Carrega os códigos de produtos do XML e SPED
    cProdXML = CamposXML(2)
    cProdSPED = CamposSPED(2)
    
    arrCampos.Add "'C170"
    arrCampos.Add "'" & Operacao
    For i = 0 To UBound(CamposXML)
        
        Select Case i
        
            Case 0
                arrCampos.Add fnExcel.FormatarTexto(CamposXML(i))
                arrCampos.Add fnExcel.FormatarTexto(CamposSPED(i))
            
            Case Else
                arrCampos.Add CamposXML(i)
                arrCampos.Add CamposSPED(i)
        
        End Select
        
        'Carrega informações do cadastro de produtos
        If i = 2 Then
            For j = 0 To 5
                
                Select Case j
    
                    Case 1, 3, 4, 5
                        arrCampos.Add fnExcel.FormatarTexto(dicProdutosXML(cProdXML)(j))
                        arrCampos.Add fnExcel.FormatarTexto(dicProdutosSPED(cProdSPED)(j))
                    
                    Case 2
                        
                    Case Else
                        arrCampos.Add dicProdutosXML(cProdXML)(j)
                        arrCampos.Add dicProdutosSPED(cProdSPED)(j)
                    
                End Select
                         
            Next j
        End If
            
    Next i
    
    arrCampos.Add Pontuacao
    arrCampos.Add "NÃO"
    
    RegistrarPontuacao = arrCampos.toArray()

End Function

Public Function ChecarCorrelacoes(ByRef dicCorrelacoesXMLSPED As Dictionary, _
    ByRef dicMelhoresCorrelacoes As Dictionary, ByRef dicTitulosOperacoes As Dictionary, _
    ByRef dicCorrelacoes As Dictionary, ByRef dicTitulosCorrelacoes As Dictionary, ByRef arrRelatorio As ArrayList)
    
Dim CNPJEmit As String, cProdNF$, cProdSPED$, xProdSPED$, uCom$, uInv$, Chave$
Dim nItemSPED As Variant, nItemXML, Correlacoes, Correlacao, Campos
Dim i As Byte
    
    'Percorre todas as melhores correlações entre os itens do SPED e XML
    For Each nItemSPED In dicMelhoresCorrelacoes.Keys()
            
        'Verifica se o item do SPED se correlacionou com mais de um item do XML
        If dicMelhoresCorrelacoes(nItemSPED).Count > 1 Then
            
            'Percorre os itens do XML que estão correlacionados ao item do SPED atual
            For Each nItemXML In dicMelhoresCorrelacoes(nItemSPED).Keys()
                
                'Registra todas as correlações do item do XML com todos os itens do SPED para o usuário escolher
                For Each Correlacao In dicCorrelacoesXMLSPED(nItemXML).Items
                    
                    If VarType(Correlacao) = 8204 Then
                        arrRelatorio.Add Correlacao
                    End If
                    
                Next Correlacao
            
            Next nItemXML
        
        Else
        
            'Percorre os itens do XML que estão correlacionados ao item do SPED atual
            For Each Correlacao In dicMelhoresCorrelacoes(nItemSPED).Items()
                
                If IsEmpty(Correlacao) Then GoTo Prx:
                If LBound(Correlacao) = 0 Then i = 1 Else i = 1
                CNPJEmit = Util.RemoverAspaSimples(Correlacao(dicTitulosOperacoes("CNPJ_EMIT_NF") - i))
                cProdNF = Util.RemoverAspaSimples(Correlacao(dicTitulosOperacoes("COD_PROD_NF") - i))
                cProdSPED = Correlacao(dicTitulosOperacoes("COD_PROD_SPED") - i)
                xProdSPED = Correlacao(dicTitulosOperacoes("DESCRICAO_SPED") - i)
                uCom = Util.RemoverAspaSimples(Correlacao(dicTitulosOperacoes("UND_NF") - i))
                uInv = Correlacao(dicTitulosOperacoes("UND_SPED") - i)
                
                Chave = CNPJEmit & cProdNF & uCom
                If dicCorrelacoes.Exists(Chave) Then
                    
                    Campos = dicCorrelacoes(Chave)
                    If LBound(Campos) = 0 Then i = 1 Else i = 0
                    
                        Campos(dicTitulosCorrelacoes("COD_ITEM") - i) = cProdSPED
                        Campos(dicTitulosCorrelacoes("DESCR_ITEM") - i) = xProdSPED
                        Campos(dicTitulosCorrelacoes("UND_INV") - i) = uInv
            
                    dicCorrelacoes(Chave) = Campos
                    
                End If
Prx:
            Next Correlacao
            
        End If

    Next nItemSPED
    
End Function

Public Function IdentificarMelhorCorrelacao(ByRef Campos As Variant) As Variant

    Campos(UBound(Campos)) = "SIM"
    IdentificarMelhorCorrelacao = Campos
    
End Function

Private Function TokenizarPalavras(Texto As String) As Collection

Dim dicPalavras As New Collection
Dim Palavras() As String
Dim i As Integer
    
    'Divide as palavras usando o espaço como delimitador
    Palavras = Split(Texto, " ")
    
    'Amarazena as palavras no Colletction
    For i = LBound(Palavras) To UBound(Palavras)
        On Error Resume Next
            dicPalavras.Add Palavras(i), Palavras(i)
        On Error GoTo 0
    Next i
    
    'Devolve as palavras tokenizadas como resultado da função
    Set TokenizarPalavras = dicPalavras
    
End Function

Private Function CompararDescricoes(descXML As String, descSPED As String) As Double

Dim colXML As New Collection
Dim colSPED As New Collection
Dim colUniao As New Collection
Dim colIntersecao As New Collection
Dim Palavra As Variant
Dim Pontuacao As Double

    Set colXML = TokenizarPalavras(descXML)
    Set colSPED = TokenizarPalavras(descSPED)
    
    For Each Palavra In colXML
        colUniao.Add Palavra, Palavra
    Next Palavra
    
    For Each Palavra In colSPED
        On Error Resume Next
            colUniao.Add Palavra, Palavra
            If Not IsError(colXML(Palavra)) Then
                colIntersecao.Add Palavra, Palavra
            End If
        On Error GoTo 0
    Next Palavra
    
    Pontuacao = colIntersecao.Count / colUniao.Count
    
    CompararDescricoes = Pontuacao
    
End Function

'TODO: Criar rotina para identificar correlações que possuem unidades de medidas diferentes

