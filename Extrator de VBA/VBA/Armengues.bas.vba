Attribute VB_Name = "Armengues"
Option Explicit

Private rC181 As New clsC181
Private rC185 As New clsC185
Private ImportNFeNFCe As New AssistenteImportacaoNFeNFCe
Public ValidacoesCFOP As New clsRegrasFiscaisCFOP
Public ValidacoesNCM As New clsRegrasFiscaisNCM
Public dtIni As String, dtFin$
Public DadosNFe As DadosNotasFiscais2

Public Type DadosNotasFiscais2
    
    cBarra As String
    CEST As String
    CFOP As String
    Chave As String
    chNFe As String
    CPF As String
    CNPJ As String
    CNPJDest As String
    CNPJEmit As String
    CNPJPart As String
    CNPJFornec As String
    CNPJPrest As String
    cPart As String
    vSubstituto As String
    vSubstituido As String
    cMun As String
    Filho As String
    Transmissao As String
    Registro As String
    RazaoFornec As String
    RazaoPrest As String
    InscEstadual As String
    cProd As String
    cAjuste As String
    cResponsavel As String
    indEmit As String
    ARQUIVO As String
    Movimento As String
    tpGIA As String
    tpAjuste As String
    vISS As String
    vContabil As String
    vContabilContrib As String
    vContabilNaoContrib As String
    cTributaria As String
    cANP As String
    CNAE As String
    DESCRICAO As String
    Referencia As String
    ItemServico As String
    vBCICMSDet As String
    vTotal As String
    vISSRetido As String
    vPISRetido As String
    vAjuste As String
    vCOFINSRetido As String
    vIRPJRetido As String
    vCSLLRetido As String
    vINSSRetido As String
    vLiquido As String
    vICMSDeson As String
    CSTICMS As String
    CodSituacao As String
    CodFornec As String
    redBCICMS As String
    vBCICMS As String
    vBCICMSContrib As String
    vBCICMSNaoContrib As String
    vST As String
    vICMSDet As String
    vOutroDet As String
    pISS As String
    itemST As String
    vBCISS As String
    AnexoSN As String
    vServico As String
    vDeducoes As String
    exTIPI As String
    Competencia As String
    dhEmi As String
    dtEmi As String
    dtEnt As String
    dtLancamento As String
    tpItem As String
    FatConv As String
    indFCP As String
    idDest As String
    item As String
    ItemPai As String
    ItemR1000 As String
    ItemR1100 As String
    ItemR1300 As String
    ItemR1500 As String
    DivergNF As String
    Hash As String
    pMVA As String
    Modelo As String
    NCM As String
    NITEM As String
    nNF As String
    OBSERVACOES As String
    pCargaTrib As String
    pFCP As String
    pICMS As String
    pICMSST As String
    pRedBC As String
    pRedBCST As String
    qCom As String
    qInv As String
    vConfICMS As String
    vFECOEPRessarcir As String
    vSTRecRessarcir As String
    vResultRecRessarcir As String
    bcICMS As String
    QTD As String
    vIsentas As String
    vRedBCICMS As String
    vOutras As String
    vOutrasDet As String
    RazaoEmit As String
    RazaoDest As String
    RazaoPart As String
    tpNF As String
    tpOperacao As String
    vOperacao As String
    vTotProd As String
    vUnit As String
    vIPI As String
    pIPI As String
    vICMSEfetivo As String
    vTotICMSEfetivo As String
    vMinUnit As String
    tpEmissao As String
    uCom As String
    uInv As String
    UFDest As String
    UFEmit As String
    UF As String
    vICMS As String
    vICMSPetroleo As String
    vICMSOutro As String
    vICMSST As String
    vICMSSTDet As String
    vICMSTotal As String
    vFCP As String
    vFCPST As String
    vNFSPED As String
    vNF As String
    vBCST As String
    vProd As String
    vFrete As String
    vSeg As String
    vDesc As String
    vOutro As String
    vMedBCST As String
    vComplementoICMS As String
    vUnCom As String
    xProd As String
    Status As String
    SERIE As String
    StatusSPED As String
    vTotBCST As String
    vTotICMS As String
    vUnitMedICMS As String
    vMedICMS As String
    qTotEntradas As Double
    qTotSaidas As Double
    
End Type

Public Type DadosConhecimentos2
    
    nCTe As String
    dhEmi As String
    CNPJEmit As String
    RazaoEmit As String
    vCTe As String
    chCTe As String
    UF As String
    UFOrig As String
    Stituacao As String
    tpOperacao As String
    dtLancamento As String
    DivCTe As String
    OBSERVACOES As String
    
End Type

Function CalcularFuncao(expression As String) As Variant
    CalcularFuncao = Evaluate(expression)
End Function

Public Function CorParaRGB(ByVal Cor As Long)

    Dim R As Long
    Dim G As Long
    Dim b As Long

    ' Extrair os componentes RGB
    R = Cor Mod 256
    G = (Cor \ 256) Mod 256
    b = Cor \ 65536

    ' Exibir os componentes RGB
    CorParaRGB = "RGB(" & R & ", " & G & ", " & b & ")"

End Function

Public Sub ImportarDadosXML(ByRef dicDados As Dictionary)

Dim Produtos As IXMLDOMNodeList
Dim Produto As IXMLDOMNode
Dim NFe As New DOMDocument60
Dim Arqs As Variant, Arq
                          
    Arqs = Util.SelecionarArquivos("xml")
    For Each Arq In Arqs
        
        With CamposC170
            
            Set NFe = fnXML.RemoverNamespaces(Arq)
            
            .CHV_PAI = VBA.Right(fnXML.ValidarTag(NFe, "//@Id"), 44)
            
            Set Produtos = NFe.SelectNodes("//det")
            For Each Produto In Produtos
            
                .NUM_ITEM = fnXML.ValidarTag(Produto, "@nItem")
                .VL_ITEM = fnXML.ValidarValores(Produto, "prod/vProd")
                .VL_DESC = fnXML.ValidarValores(Produto, "prod/vDesc")
                .VL_BC_ICMS = fnXML.ValidarValores(Produto, "imposto/ICMS//vBC")
                .ALIQ_ICMS = fnXML.ValidarPercentual(Produto, "imposto/ICMS//pICMS")
                .VL_ICMS = fnXML.ValidarValores(Produto, "imposto/ICMS//vICMS")
                .VL_BC_ICMS_ST = fnXML.ValidarValores(Produto, "imposto/ICMS//vBCST")
                .ALIQ_ST = fnXML.ValidarPercentual(Produto, "imposto/ICMS//pICMSST")
                .VL_ICMS_ST = fnXML.ValidarValores(Produto, "imposto/ICMS//vICMSST")
                .VL_BC_IPI = fnXML.ValidarValores(Produto, "imposto/IPI//vBC")
                .ALIQ_IPI = fnXML.ValidarPercentual(Produto, "imposto/IPI//pIPI")
                .VL_IPI = fnXML.ValidarValores(Produto, "imposto/IPI//vIPI")
                '.CST_ICMS = fnXML.ExtrairCST_CSOSN_ICMS(Produto)
                        
                dicDados(.CHV_PAI & .NUM_ITEM) = Array(CDbl(.VL_ITEM), CDbl(.VL_DESC), CDbl(.VL_BC_ICMS), _
                                                       CDbl(.ALIQ_ICMS), CDbl(.VL_ICMS), CDbl(.VL_BC_ICMS_ST), CDbl(.ALIQ_ST), _
                                                       CDbl(.VL_ICMS_ST), CDbl(.VL_BC_IPI), CDbl(.ALIQ_IPI), CDbl(.VL_IPI))
                                       
            Next Produto
        
        End With
        
    Next Arq
    
    Application.StatusBar = False
    
End Sub

Public Function GerarRegistros(ParamArray Registros() As Variant)

    Dim Registro
    
    For Each Registro In Registros
        Debug.Print "<button id=""btnReg" & Registro & """ Label=""" & Registro & """ image=""" & Registro & """ tag=""" & Registro & """ onAction=""AcionarBotao"" />"
    Next Registro
    
End Function

Private Sub ExtrairDadosAnaliticosEFD(ByVal Registro As String, ByVal Modelo As String, ByRef Dicionario As Dictionary)
    
    Dim Campos

    Campos = Split(Registro, "|")
    
    With DadosNFe
    
        .CFOP = Campos(3)
        .CSTICMS = Campos(2)
        .pICMS = Util.ValidarValores(Campos(4)) / 100
        .vOperacao = Campos(5)
        .bcICMS = Campos(6)
        .vICMS = Campos(7)
        
        Select Case Campos(1)
        
        Case "C190", "C590"
            .vBCST = Campos(8)
            .vICMSST = Campos(9)
            .vRedBCICMS = Campos(10)
            .vIPI = Campos(11)
            
        Case "D190"
            .vRedBCICMS = Campos(8)
                
        End Select
        
        If Campos(1) = "C590" Then .vIPI = 0
        .Chave = Modelo & .CFOP & .CSTICMS & .pICMS
        
        If Dicionario.Exists(.Chave) Then
        
            .vOperacao = .vOperacao + Dicionario(.Chave)(5)
            .bcICMS = .bcICMS + Dicionario(.Chave)(6)
            .vICMS = .vICMS + Dicionario(.Chave)(7)
            .vBCST = .vBCST + Dicionario(.Chave)(8)
            .vICMSST = .vICMSST + Dicionario(.Chave)(9)
            .vRedBCICMS = .vRedBCICMS + Dicionario(.Chave)(10)
            .vIPI = .vIPI + Dicionario(.Chave)(11)
            
        End If
        
        Select Case Right(.CSTICMS, 2)
            
        Case "20", "30", "40", "41", "70"
            .vIsentas = Round(CDbl(.vOperacao) - CDbl(.bcICMS) - CDbl(.vICMSST) - CDbl(.vIPI), 2)
            .vOutras = 0
                
        Case Else
            .vOutras = Round(CDbl(.vOperacao) - CDbl(.bcICMS) - CDbl(.vICMSST) - CDbl(.vIPI), 2)
            .vIsentas = 0
                
        End Select
        
        Dicionario(Dicionario.Count + 1) = Array(Modelo, CInt(.CFOP), "'" & .CSTICMS, CDbl(.pICMS), CDbl(.vOperacao), CDbl(.bcICMS), CDbl(.vICMS), _
                                                 CDbl(.vBCST), CDbl(.vICMSST), CDbl(.vRedBCICMS), CDbl(.vIPI), CDbl(.vIsentas), CDbl(.vOutras))
        
    End With
    
End Sub

Public Function ExtrairModelo(ByVal Registro As String) As String

    Dim Campos

    Campos = Split(Registro, "|")
    
    Select Case Campos(1)
    
    Case "C100", "C500", "D100"
        ExtrairModelo = Campos(5)
        
    Case "C800"
        ExtrairModelo = Campos(2)
        
    End Select

End Function

Public Sub SelecionarRegistros()

    Dim Registro As Variant, Registros As Variant
    Dim EFD As New ArrayList
    Dim nReg As String
    Dim Arqs, Arq
    
    Arqs = Util.SelecionarArquivos("txt")
    If VarType(Arqs) <> 11 Then
        
        Inicio = Now()
        For Each Arq In Arqs
            
            Registros = Util.ImportarTxt(Arq)
            For Each Registro In Registros
                
                a = Util.AntiTravamento(a)
                nReg = VBA.Mid(Registro, 2, 4)
                Select Case True
                                        
                Case (nReg = "0000") Or (nReg = "0150") Or (nReg = "0300") Or (nReg = "0305") Or (nReg = "0500") Or (nReg = "0600")
                    EFD.Add Registro
                        
                End Select
                
            Next Registro
            
            Arq = Replace(VBA.UCase(Arq), ".TXT", " - SELECIONADO.txt")
            Call Util.ExportarTxt(Arq, fnSPED.TotalizarRegistrosSPED(EFD))
            
            EFD.Clear
            
        Next Arq
        
    End If
    
    regC190.Activate
    Call Util.MsgInformativa("SPED estruturado com sucesso!", "Estruturação do SPED Fiscal", Inicio)
    
End Sub

Public Sub ImportarRegistroAtual()

Dim Registro As Variant, Registros As Variant
Dim EFD As New ArrayList
Dim nReg As String
Dim Arqs, Arq
    
    Arqs = Util.SelecionarArquivos("txt")
    If VarType(Arqs) <> 11 Then
        
        Inicio = Now()
        For Each Arq In Arqs
                
            Registros = Util.ImportarTxt(Arq)
            For Each Registro In Registros
                
                a = Util.AntiTravamento(a)
                nReg = VBA.Mid(Registro, 2, 4)
                Select Case True
                    
                    Case nReg = "0000"
                        Call r0000.ImportarParaExcel(Registro, regEFD.dic0000)
                        
                End Select
                
            Next Registro
            
            Arq = Replace(VBA.UCase(Arq), ".TXT", " - SELECIONADO.txt")
            Call fnSPED.ExportarSPED(Arq, fnSPED.TotalizarRegistrosSPED(EFD))
            
        Next Arq
        
    End If
    
    Call Util.ExportarDadosDicionario(reg0000, regEFD.dic0000)
    Call Util.MsgInformativa("Registro importado com sucesso!", "Importação de registro do SPED Fiscal", Inicio)
    
End Sub

Public Function DefinirCaminhoNome() As String

    Dim Caminho As Variant
    Dim NomeArquivo As String
    Dim ArquivoTXT As String

    On Error Resume Next

    ' Abre a caixa de diálogo personalizada para o usuário
    With Application.FileDialog(msoFileDialogSaveAs)
        
        .Title = "Selecione a pasta de destino e digite o nome do arquivo"
        .InitialFileName = "ArquivoExportado.txt"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Arquivos de Texto", "*.txt"
        
        If .Show <> -1 Then Exit Function
        Caminho = .SelectedItems(1)
        
    End With

    ' Verifica se o usuário cancelou a operação
    If Caminho = "" Then Exit Function

    ' Cria o caminho completo para o arquivo TXT
    ArquivoTXT = Caminho


End Function

Public Function ExportarTxtSelecionarCaminho() As String
    Dim Caminho As Variant
    Dim NomeArquivo As String
    Dim ArquivoTXT As String

    On Error Resume Next

    ' Abre a caixa de diálogo personalizada para o usuário
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "Selecione a pasta de destino e digite o nome do arquivo"
        .InitialFileName = "ArquivoExportado"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Arquivos de Texto", "*.txt"

        If .Show <> -1 Then Exit Function

        Caminho = .SelectedItems(1)
    End With

    ' Verifica se o usuário cancelou a operação
    If Caminho = "" Then Exit Function

    ' Cria o caminho completo para o arquivo TXT sem a extensão
    ArquivoTXT = Caminho

    ' Retorna o caminho completo do arquivo TXT
    ExportarTxtSelecionarCaminho = ArquivoTXT
End Function

Public Function ExportarTxtSelecionarCaminho2() As String
    Dim Caminho As Variant
    Dim NomeArquivo As String
    Dim ArquivoTXT As String

    On Error Resume Next

    ' Abre a caixa de diálogo personalizada para o usuário
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "Selecione a pasta de destino e digite o nome do arquivo"
        .InitialFileName = "ArquivoExportado"
        .AllowMultiSelect = False
        
        .Filters.Clear
        .Filters.Add "Arquivos de Texto", "*.txt"
        .FilterIndex = 1
        
        ' ... Outras configurações da caixa de diálogo ...
        
        If .Show <> -1 Then Exit Function
        
        Caminho = .SelectedItems(1)
    End With


    ' Verifica se o usuário cancelou a operação
    If Caminho = "" Then Exit Function

    ' Cria o caminho completo para o arquivo TXT sem a extensão
    ArquivoTXT = Caminho

    ' Retorna o caminho completo do arquivo TXT
    ExportarTxtSelecionarCaminho2 = ArquivoTXT
    
End Function

Public Function RecriarFiltros()

    Dim Plan As Worksheet
    
    For Each Plan In ThisWorkbook.Worksheets
        
        Select Case VBA.Left(Plan.CodeName, 3)
            
        Case "reg", "rel"
            Call FuncoesFiltragem.CriarFiltro(Plan)
                
        End Select
        
    Next Plan
    
End Function

Public Sub CriarPlanilhasRegistros()

    Dim Registro As Variant, Titulos
    Dim PlanExists As Boolean
    Dim Plan As Worksheet

    For Each Registro In dicLayoutContribuicoes.Keys()
        
        DoEvents
        Titulos = dicLayoutContribuicoes(Registro)("NomeCampos").toArray()
        Set Plan = SetarPlanilha(Registro, PlanExists)
        If PlanExists Then GoTo Prx:
        With Plan
        
            'Definir nome do registro
            .name = Registro

            'Remove o Autofiltro caso exista
            If .AutoFilterMode Then .AutoFilterMode = False
            
            'Limpa dados e formatações
            .Cells.Interior.ColorIndex = xlNone
            .Cells.Borders.LineStyle = xlNone
            .Cells.ClearContents
            .Cells.HorizontalAlignment = xlCenter
            .Cells.VerticalAlignment = xlCenter
            
            With .Range("A1").Resize(, UBound(Titulos) + 1)
                
                .Interior.Color = RGB(255, 242, 204)
                
                With .Borders
                    
                    .LineStyle = xlContinuous
                    .Color = RGB(0, 0, 0)
                    .Weight = xlThin
                    
                End With
                
            End With
            
            .Range("A1").value = "FILTROS"
            .Range("A1").Font.Bold = True
            With .Range("A3").Resize(, UBound(Titulos) + 1)
                .value = Titulos
                .Font.Bold = True
            End With
            .Columns("B:B").Insert Shift:=xlToRight
            .Columns("B:B").Insert Shift:=xlToRight
            .Columns("B:B").Insert Shift:=xlToRight
            .Range("B3:D3").value = Array("ARQUIVO", "CHV_REG", "CHV_PAI_FISCAL")
            
            .Range("A3").Resize(Rows.Count - 3, UBound(Titulos) + 4).AutoFilter
            .Columns.AutoFit
            
            Call FormatarColunas(Plan)
            
            'Configura planilha para proteger os títulos e fórmulas
            .Unprotect Password:="C1664B12A18ACF241F5403555175CBC7"
            .Cells.Locked = False
            .Rows("2:3").Locked = True
            .Protect Password:="C1664B12A18ACF241F5403555175CBC7", _
                     AllowFormattingCells:=True, _
                     AllowFormattingColumns:=True, _
                     AllowFormattingRows:=True, _
                     AllowInsertingColumns:=True, _
                     AllowInsertingRows:=True, _
                     AllowDeletingColumns:=True, _
                     AllowDeletingRows:=True, _
                     AllowSorting:=True, _
                     AllowFiltering:=True, _
                     AllowUsingPivotTables:=True, _
                     AllowInsertingHyperlinks:=True
            
        End With
Prx:
    Next Registro
    
End Sub

Sub FormatarColunas(ByRef Plan As Worksheet)

    Dim Titulos As Variant
    Dim i As Byte, j As Byte
    
    With Plan
        
        Titulos = Application.index(.Range(.Cells(3, 1), .Cells(3, .Columns.Count).END(xlToLeft)).value, 0, 0)
        
        For i = LBound(Titulos) To UBound(Titulos)
            
            Select Case True
                
            Case Titulos(i) Like "DT_*"
                .Columns(i).NumberFormat = "m/d/yyyy"
                    
            Case (Titulos(i) Like "VL_*") Or (Titulos(i) Like "QTD_*") Or (Titulos(i) Like "FAT_*")
                .Columns(i).Style = "Comma"
                    
                If Not Titulos(i) Like "FAT_*" Then
                    With .Cells(2, i)
                        
                        .Formula2R1C1 = "=SUBTOTAL(9,INDIRECT(SUBSTITUTE(ADDRESS(1, COLUMN(), 4), ""1"", """") & ""4:"" & SUBSTITUTE(ADDRESS(1, COLUMN(), 4), ""1"", """") & ""1048576""))"
                        .Font.Bold = True
                        .Font.Color = -4165632
                            
                    End With
                End If
                    
            Case Titulos(i) Like "ALIQ_*"
                .Columns(i).NumberFormat = "0.00%"
                    
            Case Else
                .Columns(i).NumberFormat = "@"
                    
            End Select
            
        Next i
        
    End With
    
End Sub

Public Function SetarPlanilha(ByVal Registro As String, ByRef PlanExists As Boolean) As Worksheet

    Dim nReg As String
    Dim PlanReg As Worksheet
    Dim PlanRef As Worksheet
    Dim Plan As Worksheet

    On Error Resume Next
    Set Plan = ThisWorkbook.Sheets(Registro)
    On Error GoTo 0

    If Not Plan Is Nothing Then
        Set SetarPlanilha = Plan
        PlanExists = True
        Exit Function
    End If
    
    On Error GoTo Tratar:
    PlanExists = False
    nReg = VBA.Left(Registro, 4)
    Set PlanRef = ThisWorkbook.Sheets(nReg)
    Set SetarPlanilha = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(PlanRef.index))
    On Error GoTo 0
    
    Exit Function
Tratar:
    
    For Each PlanReg In ThisWorkbook.Worksheets
        
        If (PlanReg.name > nReg Or PlanReg.name = "1001") And (PlanReg.name <> "Autenticação" And PlanReg.name <> "Scripts") Then
            Set SetarPlanilha = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(PlanReg.index - 1))
            Exit For
        End If
        
    Next PlanReg
    
    On Error GoTo 0
    
End Function

Public Function AlterarCodeNamePlanilha()

    Dim wbk As Object, sheet As Object
    Dim i As Long

    For i = 3 To ThisWorkbook.Worksheets.Count
    
        If Worksheets(i).CodeName Like "Planilha*" Then
            Set sheet = ActiveWorkbook.VBProject.VBComponents(ActiveWorkbook.Sheets(i).CodeName)
            sheet.name = "reg" & Worksheets(i).name
        End If
        
    Next i
    
End Function

Public Function CongelarPaineis()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        
        If ws.CodeName = "relInteligenteDivergencias" Then
        
            With ws
            
                .Activate
                ActiveWindow.FreezePanes = False
                .Cells(4, 1).Select
                ActiveWindow.FreezePanes = True
                
            End With
        
        End If
        
    Next ws
    
End Function

Public Function desprotegerplanilhas()

    Dim Plan As Worksheet
    
    For Each Plan In ThisWorkbook.Worksheets
        
        'Call FuncoesPlanilha.DesprotegerPlanilha(Plan)
        
    Next Plan
    
End Function

Public Function Redimensionarcolunas()

    Dim Plan As Worksheet
    
    For Each Plan In ThisWorkbook.Worksheets
        
        If Plan.Range("A4").value <> "" Then Plan.Columns.AutoFit
        
    Next Plan
    
End Function

Public Function FormatarPlanilhas()
    
Dim Plan As Worksheet
Dim qtdCols As Byte

    For Each Plan In ThisWorkbook.Worksheets
        
        With Plan
            Select Case True
                
                Case Plan.CodeName Like "reg*"
                    If .AutoFilterMode Then .AutoFilterMode = False
            
                    .Cells.HorizontalAlignment = xlCenter
                    .Cells.VerticalAlignment = xlCenter
                    
                    qtdCols = .Cells(3, Columns.Count).END(xlToLeft).Column
                    With .Range("A1").Resize(, qtdCols)
                        
                        .Interior.Color = RGB(255, 242, 204)
                        With .Borders
                            
                            .LineStyle = xlContinuous
                            .Color = RGB(0, 0, 0)
                            .Weight = xlThin
                            
                        End With
                        
                    End With
                    
                    .Range("A1").value = "FILTROS"
                    .Range("A1").Font.Bold = True
                    .Range("A3").Resize(, qtdCols).Font.Bold = True
                    .Range("A3").Resize(Rows.Count - 3, qtdCols).AutoFilter
                    .Columns.AutoFit
                    
                    Call FormatarColunas(Plan)
            
            End Select
        
        End With
    
    Next Plan
    
End Function

Sub AlterarDadosChaveAcessoSPEDContribuicoes(control As IRibbonControl)

Dim EFD As New ArrayList
Dim Registros As Variant, Registro, Campos
Dim Arq As String, nReg$, Chave$, chvDoc$, CNPJDeclarante$, CNPJ$
    
    CNPJ = "23896999000102"
    ' Abra o arquivo de texto
    Arq = Util.SelecionarArquivo("txt")
    
    Inicio = Now()
    
    Registros = Util.ImportarTxt(Arq)
    For Each Registro In Registros
        
        Campos = VBA.Split(Registro, "|")
        nReg = Campos(1)
        Select Case nReg
            
            Case "0000"
                CNPJDeclarante = Campos(9)
                Campos(8) = "ESCOLA DA AUTOMACAO FISCAL LTDA"
                Campos(9) = CNPJ
            
            Case "A010", "C010", "D010", "F010"
                Campos(2) = CNPJ
                
            Case "0140"
                Campos(3) = "ESCOLA DA AUTOMACAO FISCAL LTDA"
                Campos(4) = CNPJ
                Campos(6) = "200009351"
                
'            Case "A100"
'                ChvDoc = Campos(9)
'                Campos(9) = Util.AlterarChaveAcesso(ChvDoc, CNPJ, CNPJDeclarante)
                
            Case "C100"
                chvDoc = Campos(9)
                Campos(9) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
            
            Case "C500"
                chvDoc = Campos(15)
                Campos(15) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
            
            Case "C800"
                chvDoc = Campos(11)
                Campos(11) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "D100"
                chvDoc = Campos(10)
                Campos(10) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
                chvDoc = Campos(14)
                Campos(14) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "1101"
                chvDoc = Campos(9)
                Campos(9) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "1501"
                chvDoc = Campos(9)
                Campos(9) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
            
        End Select
            
        If Registro <> "" Then
            Registro = VBA.Join(Campos, "|")
            EFD.Add Registro
        End If
        
        If nReg = "9999" Then Exit For
        
    Next Registro
    
    If VBA.LCase(Arq) Like "*.txt" Then
        Arq = VBA.Replace(Arq, ".txt", "") & " - ALTERADO.txt"
    End If
    
    Call Util.ExportarTxt(Arq, VBA.Join(EFD.toArray(), vbCrLf))
    Call Util.MsgInformativa("Acabou", "Terminou", Inicio)
    
End Sub

Sub AlterarDadosChaveAcessoSPEDFiscal(control As IRibbonControl)

Dim EFD As New ArrayList
Dim Registros As Variant, Registro, Campos
Dim Arq As String, nReg$, CNPJDeclarante$, Chave$, chvDoc$, CNPJ$
    
    CNPJ = "23896999000102"
    ' Abra o arquivo de texto
    Arq = Util.SelecionarArquivo("txt")
    
    Inicio = Now()
    
    Registros = Util.ImportarTxt(Arq)
    For Each Registro In Registros
        
        Campos = VBA.Split(Registro, "|")
        nReg = Campos(1)
        Select Case nReg
            
            Case "0000"
                CNPJDeclarante = Campos(7)
                Campos(6) = "ESCOLA DA AUTOMACAO FISCAL LTDA"
                Campos(7) = CNPJ
                Campos(10) = "200009351"
            
            Case "0005"
                Campos(2) = "ESCOLA DA AUTOMACAO FISCAL"
                
            Case "B020"
                chvDoc = Campos(9)
                Campos(9) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "C100"
                chvDoc = Campos(9)
                Campos(9) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
            
            Case "C113"
                chvDoc = Campos(10)
                Campos(10) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "C116"
                chvDoc = Campos(4)
                Campos(10) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "C181"
                chvDoc = Campos(9)
                Campos(9) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "C186"
                chvDoc = Campos(12)
                Campos(12) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
            
            Case "C465"
                chvDoc = Campos(2)
                Campos(2) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "C500"
                chvDoc = Campos(28)
                Campos(28) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
            
                chvDoc = Campos(30)
                Campos(30) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "C800"
                chvDoc = Campos(11)
                Campos(11) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "D100"
                chvDoc = Campos(10)
                Campos(10) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
                chvDoc = Campos(14)
                Campos(14) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "D700"
                chvDoc = Campos(22)
                Campos(22) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                                                            
                chvDoc = Campos(26)
                Campos(26) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                                           
            Case "E113"
                chvDoc = Campos(10)
                Campos(10) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "E240"
                chvDoc = Campos(10)
                Campos(10) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "E313"
                chvDoc = Campos(7)
                Campos(7) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "E531"
                chvDoc = Campos(10)
                Campos(10) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "G130"
                chvDoc = Campos(7)
                Campos(7) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "1105"
                chvDoc = Campos(5)
                Campos(5) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "1110"
                chvDoc = Campos(7)
                Campos(7) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "1210"
                chvDoc = Campos(5)
                Campos(5) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
            Case "1923"
                chvDoc = Campos(10)
                Campos(10) = Util.AlterarChaveAcesso(chvDoc, CNPJ, CNPJDeclarante)
                
        End Select
            
        If Registro <> "" Then
            Registro = VBA.Join(Campos, "|")
            EFD.Add Registro
        End If
        
        If nReg = "9999" Then Exit For
        
    Next Registro
    
    If VBA.LCase(Arq) Like "*.txt" Then
        Arq = VBA.Replace(Arq, ".txt", "") & " - ALTERADO.txt"
    End If
    
    Call Util.ExportarTxt(Arq, VBA.Join(EFD.toArray(), vbCrLf))
    Call Util.MsgInformativa("Acabou", "Terminou", Inicio)
    
End Sub

Public Function AlterarNFeEmitidos()

Dim Arq As Variant
Dim chNFe As IXMLDOMNode
Dim Razao As IXMLDOMNode
Dim refNFe As IXMLDOMNode
Dim ChvNFe As IXMLDOMNode
Dim NFe As New DOMDocument60
Dim FANTASIA As IXMLDOMNode
Dim CNPJEmit As IXMLDOMNode
Dim arrArqs As New ArrayList
Dim NotasRef As IXMLDOMNodeList
Dim Chave As String, Origem$, Destino$, Caminho$, CNPJ$, CNPJRef$
    
    CNPJRef = "41670046000103"
    CNPJ = "23896999000102"
    
    Origem = Util.SelecionarPasta("Selecione a pasta raiz para buscar os XMLS")
    Destino = Util.SelecionarPasta("Selecione a pasta onde deseja salvar os XMLS alterados")
    
    If Origem = "" Or Destino = "" Then Exit Function
    Inicio = Now()
    Call Util.ListarArquivos(arrArqs, Origem)
    
    If arrArqs.Count > 0 Then
        
        For Each Arq In arrArqs
            
            With CamposC170
                
                Set NFe = fnXML.RemoverNamespaces(Arq)
                
                Set CNPJEmit = fnXML.SetarTag(NFe, "//emit/CNPJ")
                Set chNFe = fnXML.SetarTag(NFe, "//@Id")
                Chave = VBA.Right(chNFe.text, 44)
                
                If Not fnXML.ValidarNFe(NFe) Then GoTo Prx:
                If CNPJEmit Is Nothing Then GoTo Prx:
                
                If VBA.InStr(1, Chave, CNPJRef) > 0 Then
                    
                    Set Razao = fnXML.SetarTag(NFe, "//emit/xNome")
                    Set FANTASIA = fnXML.SetarTag(NFe, "//emit/xFant")
                    
                    chNFe.text = "NFe" & Util.AlterarChaveAcesso(VBA.Right(chNFe.text, 44), CNPJ, CNPJRef)
                    CNPJEmit.text = CNPJ
                    
                    If Not Razao Is Nothing Then Razao.text = "ESCOLA DA AUTOMACAO FISCAL LTDA"
                    If Not FANTASIA Is Nothing Then FANTASIA.text = "ESCOLA DA AUTOMACAO FISCAL"
                    
                    Set chNFe = fnXML.SetarTag(NFe, "//chNFe")
                    chNFe.text = Util.AlterarChaveAcesso(chNFe.text, CNPJ, CNPJRef)
                    
                    Set NotasRef = NFe.SelectNodes("//NFref")
                    For Each refNFe In NotasRef
                        
                        If VBA.InStr(1, refNFe.text, CNPJRef) > 0 Then refNFe.text = Util.AlterarChaveAcesso(refNFe.text, CNPJ, CNPJRef)
                        
                    Next refNFe
                    
                    If VBA.LCase(Arq) Like "*.xml" Then
                        Caminho = VBA.Left(Arq, VBA.InStrRev(Arq, "\") - 1)
                        Arq = VBA.Replace(VBA.Replace(Arq, Caminho, Destino), CNPJRef, CNPJ)
                        Arq = VBA.Replace(Arq, ".xml", "") & " - ALTERADO.xml"
                    End If
                    
                    Call NFe.Save(Arq)
                    
                End If
                
            End With

Prx:
        Next Arq
        
        Call Util.MsgInformativa("Acabou", "Terminou", Inicio)
        
    End If
    
End Function

Public Function AlterarProtocolos()

Dim Arq As Variant
Dim chNFe As IXMLDOMNode
Dim NFe As New DOMDocument60
Dim CNPJEmit As IXMLDOMNode
Dim arrArqs As New ArrayList
Dim Chave As String, Origem$, Destino$, Caminho$, CNPJ$, CNPJRef$
    
    CNPJRef = "41670046000103"
    CNPJ = "23896999000102"
    
    Origem = Util.SelecionarPasta("Selecione a pasta raiz para buscar os XMLS")
    Destino = Util.SelecionarPasta("Selecione a pasta onde deseja salvar os XMLS alterados")
    
    If Origem = "" Or Destino = "" Then Exit Function
    Inicio = Now()
    Call Util.ListarArquivos(arrArqs, Origem)
    
    If arrArqs.Count > 0 Then
        
        For Each Arq In arrArqs
            
            With CamposC170
                
                Set NFe = fnXML.RemoverNamespaces(Arq)
                
                If Not fnXML.ValidarProtocoloCancelamento(NFe) Then GoTo Prx:
                
                Set CNPJEmit = fnXML.SetarTag(NFe, "//infEvento/CNPJ")
                Set chNFe = fnXML.SetarTag(NFe, "//infEvento/chNFe")
                
                If VBA.InStr(1, chNFe.text, CNPJRef) > 0 Then

                    chNFe.text = Util.AlterarChaveAcesso(chNFe.text, CNPJ, CNPJRef)
                    CNPJEmit.text = CNPJ
                    
                    Set chNFe = fnXML.SetarTag(NFe, "//retEvento//chNFe")
                    chNFe.text = Util.AlterarChaveAcesso(chNFe.text, CNPJ, CNPJRef)
                    
                    If VBA.LCase(Arq) Like "*.xml" Then
                        Caminho = VBA.Left(Arq, VBA.InStrRev(Arq, "\") - 1)
                        Arq = VBA.Replace(VBA.Replace(Arq, Caminho, Destino), CNPJRef, CNPJ)
                        Arq = VBA.Replace(Arq, ".xml", "") & " - ALTERADO.xml"
                    End If
                    
                    Call NFe.Save(Arq)
                    
                End If
                
            End With

Prx:
        Next Arq
        
        Call Util.MsgInformativa("Acabou", "Terminou", Inicio)
        
    End If
    
End Function

Public Function AlterarNFeRecebidos()

Dim Arq As Variant
Dim chNFe As IXMLDOMNode
Dim Razao As IXMLDOMNode
Dim refNFe As IXMLDOMNode
Dim ChvNFe As IXMLDOMNode
Dim NFe As New DOMDocument60
Dim FANTASIA As IXMLDOMNode
Dim CNPJDest As IXMLDOMNode
Dim arrArqs As New ArrayList
Dim NotasRef As IXMLDOMNodeList
Dim Chave As String, Origem$, Destino$, Caminho$, CNPJ$, CNPJRef$
    
    CNPJRef = "41670046000103"
    CNPJ = "23896999000102"
    
    Origem = Util.SelecionarPasta("Selecione a pasta raiz para buscar os XMLS")
    Destino = Util.SelecionarPasta("Selecione a pasta onde deseja salvar os XMLS alterados")
    
    If Origem = "" Or Destino = "" Then Exit Function
    Inicio = Now()
    Call Util.ListarArquivos(arrArqs, Origem)
    
    If arrArqs.Count > 0 Then
        
        For Each Arq In arrArqs
            
            With CamposC170
                
                Set NFe = fnXML.RemoverNamespaces(Arq)
                
                Set CNPJDest = fnXML.SetarTag(NFe, "//dest/CNPJ")
                
                If Not fnXML.ValidarNFe(NFe) Then GoTo Prx:
                If CNPJDest Is Nothing Then GoTo Prx:
                
                If CNPJDest.text = CNPJRef Then
                    
                    Set Razao = fnXML.SetarTag(NFe, "//dest/xNome")
                    Set FANTASIA = fnXML.SetarTag(NFe, "//dest/xFant")
                    CNPJDest.text = CNPJ
                    
                    If Not Razao Is Nothing Then Razao.text = "ESCOLA DA AUTOMACAO FISCAL LTDA"
                    If Not FANTASIA Is Nothing Then FANTASIA.text = "ESCOLA DA AUTOMACAO FISCAL"
                    
                    Set NotasRef = NFe.SelectNodes("//NFref")
                    For Each refNFe In NotasRef
                        
                        If VBA.InStr(1, refNFe.text, CNPJRef) > 0 Then refNFe.text = Util.AlterarChaveAcesso(refNFe.text, CNPJ, CNPJRef)
                        
                    Next refNFe
                    
                    If VBA.LCase(Arq) Like "*.xml" Then
                        Caminho = VBA.Left(Arq, VBA.InStrRev(Arq, "\") - 1)
                        Arq = VBA.Replace(VBA.Replace(Arq, Caminho, Destino), CNPJRef, CNPJ)
                        Arq = VBA.Replace(Arq, ".xml", "") & " - ALTERADO.xml"
                    End If
                    
                    Call NFe.Save(Arq)
                    
                End If
                
            End With

Prx:
        Next Arq
        
        Call Util.MsgInformativa("Acabou", "Terminou", Inicio)
        
    End If
    
End Function

Public Function AlterarCTeRecebidos()

Dim Arq As Variant
Dim chCTe As IXMLDOMNode
Dim Razao As IXMLDOMNode
Dim refCTe As IXMLDOMNode
Dim ChvCTe As IXMLDOMNode
Dim CTe As New DOMDocument60
Dim FANTASIA As IXMLDOMNode
Dim CNPJDest As IXMLDOMNode
Dim arrArqs As New ArrayList
Dim NotasRef As IXMLDOMNodeList
Dim Chave As String, Origem$, Destino$, Caminho$, CNPJ$, CNPJRef$
Dim Alterado As Boolean
    
    CNPJRef = "41670046000103"
    CNPJ = "23896999000102"
    
    Origem = Util.SelecionarPasta("Selecione a pasta raiz para buscar os XMLS")
    Destino = Util.SelecionarPasta("Selecione a pasta onde deseja salvar os XMLS alterados")
    
    If Origem = "" Or Destino = "" Then Exit Function
    Inicio = Now()
    Call Util.ListarArquivos(arrArqs, Origem)
    
    If arrArqs.Count > 0 Then
        
        For Each Arq In arrArqs
            
            With CamposC170
                
                CTe.Load (Arq)
                
                Alterado = fnXML.AlterarCNPJTomador(CTe, CNPJ, CNPJRef, "ESCOLA DA AUTOMACAO FISCAL LTDA", "ESCOLA DA AUTOMACAO FISCAL")
                
                If Alterado Then
                    
                    If VBA.LCase(Arq) Like "*.xml" Then
                        Caminho = VBA.Left(Arq, VBA.InStrRev(Arq, "\") - 1)
                        Arq = VBA.Replace(VBA.Replace(Arq, Caminho, Destino), CNPJRef, CNPJ)
                        Arq = VBA.Replace(Arq, ".xml", "") & " - ALTERADO.xml"
                    End If
                    
                    Call CTe.Save(Arq)
                    
                End If
                
            End With
            
Prx:
        Next Arq
        
        Call Util.MsgInformativa("Acabou", "Terminou", Inicio)
        
    End If
    
End Function

Public Sub EstimarEspacoOcupado()
    Dim ws As Worksheet
    Dim UltLin As Long
    Dim DadosAcionamentos As Variant
    Dim Acionamento As Variant
    Dim EspacoOcupado As Double
    Dim TamanhoMaximoBytes As Double
    Dim TamanhoMedioAcionamento As Double
    Dim NumeroMaximoAcionamentos As Double
    
    ' Definir a planilha onde os dados estão armazenados
    Set ws = ThisWorkbook.Sheets("Acionamentos")
    
    ' Encontrar a última linha da coluna A
    UltLin = ws.Cells(ws.Rows.Count, 1).END(xlUp).Row
    
    ' Se não houver dados, exibir mensagem
    If UltLin = 0 Then
        MsgBox "Não há dados de acionamentos na planilha."
        Exit Sub
    End If
    
    ' Carregar os acionamentos em um array
    DadosAcionamentos = ws.Range("A1:A" & UltLin).value
    
    ' Transpor apenas se necessário (quando há múltiplas linhas)
    If UltLin > 1 Then
        DadosAcionamentos = Application.Transpose(DadosAcionamentos)
    End If
    
    ' Calcular o espaço ocupado pelos acionamentos
    EspacoOcupado = 0
    For Each Acionamento In DadosAcionamentos
        EspacoOcupado = EspacoOcupado + Len(CStr(Acionamento))
    Next Acionamento
    
    ' Calcular o espaço disponível e o número máximo de acionamentos
    TamanhoMaximoBytes = 25 * 1024 * 1024# ' 25 MB em bytes, usando # para garantir que é Double
    TamanhoMedioAcionamento = EspacoOcupado / UltLin
    NumeroMaximoAcionamentos = TamanhoMaximoBytes / TamanhoMedioAcionamento
    
    ' Exibir os resultados com valores arredondados
    MsgBox "Espaço ocupado pelos acionamentos: " & EspacoOcupado & " bytes" & vbCrLf & _
           "Tamanho médio de cada acionamento: " & Round(TamanhoMedioAcionamento, 2) & " bytes" & vbCrLf & _
           "Número máximo de acionamentos com 25 MB: " & Round(NumeroMaximoAcionamentos, 0)
    
End Sub

Private Function OperacoesGeradoresCredito()

Dim Registros As Variant, Registro
Dim CFOPS As String, CFOP
Dim Arq As String
    
    Arq = Util.SelecionarArquivo("txt")
    
    Registros = VBA.Split(Util.ImportarConteudo(Arq), vbCrLf)
    
    For Each Registro In Registros
        
        If Registro Like "*|*" Then
            
            CFOP = VBA.Split(Registro, "|")
            CFOPS = CFOPS & """, """ & CFOP(0)
            
        End If
        
    Next Registro
    
End Function

Public Function ValidarPadrao(ByVal Texto As String, ByVal Padrao As String) As Boolean

Dim CustomPart As New clsCustomPartXML
Dim dicRegexCFOP As Dictionary
Dim regex As New RegExp
    
    'Set dicRegexCFOP = JsonConverter.ParseJson(CustomPart.ExtrairJsonXmlPart("RegexCFOP"))
    
    'If dicRegexCFOP.Exists(Padrao) Then
        
        'regex.Pattern = dicRegexCFOP(Padrao)
        regex.Pattern = Padrao
        regex.Global = False
        
    'End If
    
    ValidarPadrao = regex.Test(Texto)
    
End Function

Private Sub CarregarDadosRegistro0000()

Dim regSPED As New clsRegistrosSPED
Dim dicDadosC190 As New Dictionary
    
    Call regSPED.CarregarDadosRegistro0000
    Debug.Print dtoRegSPED.r0000.Count
    Stop
    
End Sub

Sub RenomearColuna_CHV_PAI_Enxuto()

Dim ws As Worksheet
Dim celulaCabecalho As Range
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
    
        If ws.CodeName Like "reg*" Then
        
            Set celulaCabecalho = ws.Rows(3).Find("CHV_PAI_FISCAL", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
            If Not celulaCabecalho Is Nothing Then celulaCabecalho.value = "CHV_PAI_FISCAL"
            
        End If
        
    Next ws
    
    Application.ScreenUpdating = True
    
    MsgBox "Processo finalizado.", vbInformation
    
End Sub

Sub InserirColuna_CHV_PAI_CONTRIBUICOES()

Dim ws As Worksheet
Dim sheetsModified As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For Each ws In ThisWorkbook.Worksheets
        
        If ws.CodeName Like "reg*" Then
            
            ws.Columns("E").Insert Shift:=xlToRight
            ws.Cells(3, "E").value = "CHV_PAI_CONTRIBUICOES"
            
            sheetsModified = sheetsModified + 1
            
        End If
        
    Next ws
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox sheetsModified & " planilhas foram modificadas.", vbInformation, "Processo Concluído"
    
End Sub

Public Sub InserirTitulosPlanilha()

Dim Titulos As Variant
    
    Titulos = Array( _
        "REG", "ARQUIVO", "CHV_REG", "CHV_PAI_FISCAL", "CHV_PAI_CONTRIBUICOES", _
        "PER_APUR_ANT", "NAT_CONT_REC", "VL_CONT_APUR", "VL_CRED_COFINS_DESC", _
        "VL_CONT_DEV", "VL_OUT_DED", "VL_CONT_EXT", "VL_MUL", "VL_JUR", "DT_RECOL" _
    )
        
    With reg1600_Contr
        
        .Range("A3").Resize(1, UBound(Titulos) + 1).value = Titulos
        
        MsgBox "Títulos inseridos com sucesso na planilha '" & .name & "'.", vbInformation, "Concluído"
        
    End With
    
End Sub

Public Sub testarImportacaoRegistros()

Dim importReg As New ImportadorRegistros
    
    Call importReg.ImportarDadosRegistro
    
End Sub

