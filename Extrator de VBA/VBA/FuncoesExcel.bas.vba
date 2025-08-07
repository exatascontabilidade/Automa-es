Attribute VB_Name = "FuncoesExcel"
Option Explicit

Public Sub CarregarDadosDocumentos(ByVal Planilha As Worksheet, ByRef Dicionario As Dictionary, Optional ByRef arrChaves As ArrayList)

Dim Cels As Variant
Dim UltLin As Long, i As Long
Dim Coluna As String

    With Planilha
        If .FilterMode = True Then .AutoFilter.ShowAllData
        UltLin = .Range("A" & Rows.Count).END(xlUp).Row
        Coluna = Util.NomeCol(.Range("XFD3").END(xlToLeft).Address)
        
        If UltLin > 3 Then
            Cels = .Range("A4:" & Coluna & UltLin)
            For i = LBound(Cels) To UBound(Cels)
                
                If arrChaves Is Nothing Then GoTo Pular:
                If Not arrChaves.contains(Replace(Cels(i, 6), "'", "")) Then
Pular:
                    With DadosDoce
                        
                        .nNF = Cels(i, 1)
                        .CNPJPart = Cels(i, 2)
                        .RazaoPart = Cels(i, 3)
                        .dtEmi = Util.FormatarData(Cels(i, 4))
                        .vNF = Cels(i, 5)
                        .chNFe = Replace(Cels(i, 6), "'", "")
                        .UF = Cels(i, 7)
                        .Status = Cels(i, 8)
                        .tpNF = Cels(i, 9)
                        .StatusSPED = Cels(i, 10)
                        .DivergNF = Cels(i, 11)
                        .OBSERVACOES = Cels(i, 12)
                        
                        If .vNF = "" Then .vNF = 0
                        If IsDate(.StatusSPED) Then .StatusSPED = VBA.Format(.StatusSPED, "yyyy-mm-dd")
                        Dicionario(.chNFe) = Array(.nNF, .CNPJPart, .RazaoPart, .dtEmi, CDbl(.vNF), "'" & .chNFe, .UF, .Status, .tpNF, .StatusSPED, .DivergNF, .OBSERVACOES)
                               
                    End With
                    
                End If
            Next i
        End If
    End With
    
End Sub

Public Function CarregarDados(ByRef Plan As Worksheet, ByRef dicDados As Dictionary, ParamArray chaveCampos())

Dim i As Long
Dim arrChave As New ArrayList
Dim dicTitulos As New Dictionary
Dim Dados As Variant, Campos, Campo, Chave
    
    On Error GoTo Tratar:
    
    Inicio = Now()
    If Plan.AutoFilterMode Then Plan.AutoFilter.ShowAllData
    Set dicTitulos = Util.IndexarDados(Util.DefinirTitulos(Plan, 3))
    
    Dados = Util.DefinirDados(Plan, 4, 3)
    
    'verifica se a variável 'chaveCampos' é um array bidimensional e transforma num array unidimensional
    If UBound(chaveCampos) = 0 Then chaveCampos = Application.index(chaveCampos(0), 0, 0)
    If UBound(chaveCampos) > 0 Then
    
        For i = LBound(Dados, 1) To UBound(Dados, 1)
                    
            Campos = Application.index(Dados, i, 0)
        
            For Each Campo In chaveCampos
                arrChave.Add Campos(dicTitulos(Campo))
            Next Campo
    
            Chave = fnSPED.MontarChaveRegistro(VBA.Join(arrChave.toArray, "|"))

            dicDados(Chave) = Campos
            arrChave.Clear
            
        Next i
    
    Else
    
        For i = LBound(Dados, 1) To UBound(Dados, 1)
                    
            Campos = Application.index(Dados, i, 0)
            Chave = fnSPED.MontarChaveRegistro(CStr(Campos(dicTitulos(chaveCampos(0)))))
            
            dicDados(Chave) = Campos
            
        Next i
    
    End If
    
Tratar:

Dim Msg As String

    Select Case Err.Number
        
        Case 9
            If IsEmpty(Campo) Then
                
                Msg = "O Campo '" & chaveCampos(0) & "' não existe na panilha selecionada." & vbCrLf & "Por favor revise os parâmetros informados."
            
            Else
                
                Msg = "O Campo '" & Campo & "' não existe na panilha selecionada." & vbCrLf & "Por favor revise os parâmetros informados."
    
            End If
            
            Call Util.MsgAlerta(Msg, "Campo inválido")
    
    End Select
    
End Function

Public Function CarregarDadosAnaliseSPED(ByRef dicDivergencias As Dictionary)

Dim Intervalo As Variant
Dim UltLin As Long, i&
Dim Dados As Variant
    
    If relInteligenteDivergencias.AutoFilterMode Then relInteligenteDivergencias.AutoFilter.ShowAllData
    Intervalo = relInteligenteDivergencias.Range("A3:" & Util.ConverterNumeroColuna(relInteligenteDivergencias.Range("A3").END(xlToRight).Column) & "3")
    
    UltLin = relInteligenteDivergencias.Range("A" & Rows.Count).END(xlUp).Row
    If UltLin > 3 Then
           
        Dados = relInteligenteDivergencias.Range("A4:" & Util.ConverterNumeroColuna(relInteligenteDivergencias.Range("A3").END(xlToRight).Column) & UltLin)
        For i = LBound(Dados) To UBound(Dados)
            
            Call Util.AntiTravamento(a, 100, "Carregando dados do SPED, por favor aguarde...", UBound(Dados) + 1, Comeco)
            With RelDiverg
                
                If Dados(i, 1) <> "DOC_CONTRIB" Then
                    
                    .DOC_CONTRIB = Dados(i, EncontrarColuna("DOC_CONTRIB", Intervalo))
                    .DOC_PART = Dados(i, EncontrarColuna("DOC_PART", Intervalo))
                    .Modelo = Dados(i, EncontrarColuna("MODELO", Intervalo))
                    .Operacao = Dados(i, EncontrarColuna("OPERACAO", Intervalo))
                    .TP_EMISSAO = Dados(i, EncontrarColuna("TP_EMISSAO", Intervalo))
                    .DOC_PART = Dados(i, EncontrarColuna("DOC_PART", Intervalo))
                    .Situacao = Dados(i, EncontrarColuna("SITUACAO", Intervalo))
                    .SERIE = Dados(i, EncontrarColuna("SERIE", Intervalo))
                    .NUM_DOC = Dados(i, EncontrarColuna("NUM_DOC", Intervalo))
                    .CHV_NFE = Dados(i, EncontrarColuna("CHV_NFE", Intervalo))
                    .DT_DOC = VBA.Format(Dados(i, EncontrarColuna("DT_DOC", Intervalo)), "yyyy-mm-dd")
                    .TP_PAGAMENTO = Dados(i, EncontrarColuna("TP_PAGAMENTO", Intervalo))
                    .TP_FRETE = Dados(i, EncontrarColuna("TP_FRETE", Intervalo))
                    .VL_DOC = Dados(i, EncontrarColuna("VL_DOC", Intervalo))
                    .VL_DESC = Dados(i, EncontrarColuna("VL_DESC", Intervalo))
                    .VL_ABATIMENTO = Dados(i, EncontrarColuna("VL_ABATIMENTO", Intervalo))
                    .VL_PROD = Dados(i, EncontrarColuna("VL_PROD", Intervalo))
                    .VL_FRETE = Dados(i, EncontrarColuna("VL_FRETE", Intervalo))
                    .VL_SEG = Dados(i, EncontrarColuna("VL_SEG", Intervalo))
                    .VL_OUTRO = Dados(i, EncontrarColuna("VL_OUTRO", Intervalo))
                    .VL_BC_ICMS = Dados(i, EncontrarColuna("VL_BC_ICMS", Intervalo))
                    .VL_ICMS = Dados(i, EncontrarColuna("VL_ICMS", Intervalo))
                    .VL_BC_ICMS_ST = Dados(i, EncontrarColuna("VL_BC_ICMS_ST", Intervalo))
                    .VL_ICMS_ST = Dados(i, EncontrarColuna("VL_ICMS_ST", Intervalo))
                    .VL_IPI = Dados(i, EncontrarColuna("VL_IPI", Intervalo))
                    .VL_PIS = Dados(i, EncontrarColuna("VL_PIS", Intervalo))
                    .VL_COFINS = Dados(i, EncontrarColuna("VL_COFINS", Intervalo))
                    .STATUS_ANALISE = Dados(i, EncontrarColuna("STATUS_ANALISE", Intervalo))
                    .OBSERVACOES = Dados(i, EncontrarColuna("OBSERVACOES", Intervalo))
                    
                    dicDivergencias(.CHV_NFE) = Array("'" & .DOC_CONTRIB, "'" & .DOC_PART, "'" & .Modelo, .Operacao, .TP_EMISSAO, .Situacao, _
                                                      "'" & .SERIE, "'" & .NUM_DOC, "'" & .CHV_NFE, .DT_DOC, .TP_PAGAMENTO, .TP_FRETE, CDbl(.VL_DOC), _
                                                      CDbl(.VL_PROD), CDbl(.VL_FRETE), CDbl(.VL_SEG), CDbl(.VL_OUTRO), CDbl(.VL_DESC), _
                                                      CDbl(.VL_ABATIMENTO), CDbl(.VL_BC_ICMS), CDbl(.VL_ICMS), CDbl(.VL_BC_ICMS_ST), _
                                                      CDbl(.VL_ICMS_ST), CDbl(.VL_IPI), CDbl(.VL_PIS), CDbl(.VL_COFINS), _
                                                      .STATUS_ANALISE, .OBSERVACOES)
                End If
                
            End With
            
        Next i
        
    End If
    
End Function

Public Function EncontrarColuna(ByVal NomeColuna As String, ByVal Titulos As Variant) As Integer
    
Dim i As Integer

    Titulos = Application.index(Titulos, 0, 0)
    EncontrarColuna = -1 ' Valor padrão caso a palavra não seja encontrada
    For i = LBound(Titulos) To UBound(Titulos)
        If Titulos(i) = NomeColuna Then
            EncontrarColuna = i
            Exit Function
        End If
    Next i
    
End Function

Public Function ImportarCadastro0200()

Dim dtIni As String, dtFim$, ARQUIVO$, CHV_REG_0000$, CHV_REG_0001$, CHV_REG_0190$, COD_ITEM$, UNID_INV$, Msg$
Dim Caminho As Variant, Campos, Titulo
Dim dicTitulos0200 As New Dictionary
Dim dicDados0190 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim PastaDeTrabalho As Workbook
Dim arrCampos As New ArrayList
Dim arrDados As New ArrayList
Dim Plan As Worksheet
Dim Mapeamento As Byte
    
    If PeriodoImportacao = "" Then
        Call Util.MsgAlerta("Informe o período ('MMAAAA') que deseja inserir os itens para prosseguir com a importação.", "Período de importação não informado")
        Exit Function
    End If
    
    'Carregando a inscrição estadual do contribuinte
    InscContribuinte = CadContrib.Range("InscContribuinte").value
    
    If InscContribuinte = "" Then
        Call Util.MsgAlerta("Informe a Inscrição Estadual do Contribuinte.", "Inscrição Estadual não informada")
        CadContrib.Activate
        CadContrib.Range("InscContribuinte").Activate
        Exit Function
    End If
    
    Caminho = Util.SelecionarArquivo("xlsx")
    If VarType(Caminho) = 11 Then Exit Function
    
    Inicio = Now()
    Set PastaDeTrabalho = Workbooks.Open(Caminho)
    ActiveWindow.visible = False
    
    Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
    Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200)
    
    Set Plan = PastaDeTrabalho.Worksheets(1)
    
    With Plan
    
        On Error Resume Next
            If .AutoFilterMode Then .AutoFilter.ShowAllData
        On Error GoTo 0
        
        Set dicTitulos = Util.MapearTitulos(Plan, 1)
        
        Mapeamento = 0
        
        If dicTitulos.Exists("COD_ITEM") Then Mapeamento = Mapeamento + 1
        If dicTitulos.Exists("DESCR_ITEM") Then Mapeamento = Mapeamento + 1
        If dicTitulos.Exists("UNID_INV") Then Mapeamento = Mapeamento + 1
        
        If Mapeamento < 3 Then
        
            Msg = "As colunas principais não foram mapeadas no arquivo selecionado." & vbCrLf & vbCrLf
            Msg = Msg & "Por favor realize o mapeamento dos dados e tente novamente."
            
            Call Util.MsgAlerta(Msg, "Arquivo sem dados mapeados")
            Exit Function
            
        End If
        
        Set Dados = Util.DefinirIntervalo(Plan, 2, 1)
        If Dados Is Nothing Then
        
            Call Util.MsgAlerta("O arquivo selecionado não possui dados de produtos.", "Arquivo sem dados informados")
            Exit Function
        
        End If
        
        ThisWorkbook.Windows(1).Activate
            
        a = 0
        Comeco = Timer()
        For Each Linha In Dados.Rows
            
            Call Util.AntiTravamento(a, 100, "Importando item " & a + 1 & " de " & Dados.Rows.Count, Dados.Rows.Count, Comeco)
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                For Each Titulo In dicTitulos0200
                    
                    Select Case Titulo
                        
                        Case "REG"
                            arrCampos.Add "'0200"
                            
                        Case "ARQUIVO"
                            ARQUIVO = VBA.Format(PeriodoImportacao, "00/0000") & "-" & CNPJContribuinte
                            arrCampos.Add ARQUIVO
                            
                        Case "COD_ITEM"
                            COD_ITEM = Campos(dicTitulos(Titulo))
                            arrCampos.Add fnExcel.FormatarTexto(COD_ITEM)
                            If dicDados0200.Exists(COD_ITEM) Then GoTo PrxLin:
                            
                        Case "CHV_REG"
                            
                            COD_ITEM = Campos(dicTitulos("COD_ITEM"))
                            CHV_REG_0000 = fnSPED.GerarChvReg0000(VBA.Format(PeriodoImportacao, "00\/0000"))
                            CHV_REG_0001 = fnSPED.GerarChaveRegistro(CHV_REG_0000, "0001")
                            
                            arrCampos.Add fnSPED.GerarChaveRegistro(CHV_REG_0001, COD_ITEM)
                            
                        Case "CHV_PAI_FISCAL"
                            arrCampos.Add CHV_REG_0001
                             
                        Case "TIPO_ITEM"
                            arrCampos.Add ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_TIPO_ITEM(Util.ApenasNumeros(Campos(dicTitulos(Titulo))))
                            
                        Case "COD_NCM", "EX_IPI", "CEST", "COD_BARRA"
                            arrCampos.Add fnExcel.FormatarTexto(Campos(dicTitulos(Titulo)))
                            
                        Case "UNID_INV"
                            UNID_INV = Campos(dicTitulos(Titulo))
                            arrCampos.Add Campos(dicTitulos(Titulo))

                        Case Else
                            If dicTitulos.Exists(Titulo) Then
                                arrCampos.Add Campos(dicTitulos(Titulo))
                            Else
                                arrCampos.Add ""
                            End If
                            
                    End Select
                    
                Next Titulo
                
                If UNID_INV <> "" Then CHV_REG_0190 = fnSPED.GerarChaveRegistro(CHV_REG_0001, UNID_INV)
                
                If Not dicDados0190.Exists(CHV_REG_0190) Then Call IncluirUnid0190(dicDados0190, ARQUIVO, CHV_REG_0001, CHV_REG_0190, UNID_INV)
                If Util.ChecarCamposPreenchidos(arrCampos.toArray()) Then arrDados.Add arrCampos.toArray()
                
            End If
PrxLin:
            arrCampos.Clear
            UNID_INV = ""
            COD_ITEM = ""
            
        Next Linha
        
    End With
    
    Application.DisplayAlerts = False
        PastaDeTrabalho.Close
    Application.DisplayAlerts = True
    
    If arrDados.Count = 0 Then
        Call Util.MsgAlerta("O arquivo selecionado não possui dados de produtos.", "Arquivo sem dados informados")
        Exit Function
    End If
    
    Application.StatusBar = "Exportando dados do registro 0190..."
    Call Util.ExportarDadosDicionario(reg0190, dicDados0190)
    
    Application.StatusBar = "Exportando dados do registro 0200..."
    Call Util.ExportarDadosArrayList(reg0200, arrDados)
    
    Application.StatusBar = "Importação concluída!"
    Call Util.MsgInformativa("Cadastro de produtos importado com sucesso!", "Importação do Cadastro de Produtos", Inicio)
    
End Function

Public Function ImportarCadastroK200()

Dim dtIni As String, dtFim$, ARQUIVO$, CHV_REG_0000$, CHV_REG_K001$, CHV_REG_K010$, CHV_REG_K100$, COD_ITEM$, DT_EST$, IND_EST$, COD_PART$, Msg$
Dim Caminho As Variant, Campos, Titulo
Dim dicTitulosK001 As New Dictionary
Dim dicTitulosK010 As New Dictionary
Dim dicTitulosK100 As New Dictionary
Dim dicTitulosK200 As New Dictionary
Dim dicDadosK001 As New Dictionary
Dim dicDadosK010 As New Dictionary
Dim dicDadosK100 As New Dictionary
Dim dicDadosK200 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim PastaDeTrabalho As Workbook
Dim arrCampos As New ArrayList
Dim arrDados As New ArrayList
Dim Plan As Worksheet
Dim Mapeamento As Byte, i As Byte
    
    'Carregando a inscrição estadual do contribuinte
    InscContribuinte = CadContrib.Range("InscContribuinte").value
    
    If InscContribuinte = "" Then
        Call Util.MsgAlerta("Informe a Inscrição Estadual do Contribuinte.", "Inscrição Estadual não informada")
        CadContrib.Activate
        CadContrib.Range("InscContribuinte").Activate
        Exit Function
    End If
    
    Caminho = Util.SelecionarArquivo("xlsx")
    If VarType(Caminho) = 11 Then Exit Function
    
    Inicio = Now()
    Set PastaDeTrabalho = Workbooks.Open(Caminho)
    ActiveWindow.visible = False
    
    Set dicTitulosK001 = Util.MapearTitulos(regK001, 3)
    Set dicDadosK001 = Util.CriarDicionarioRegistro(regK001, "ARQUIVO")
    
    Set dicTitulosK010 = Util.MapearTitulos(regK010, 3)
    Set dicDadosK010 = Util.CriarDicionarioRegistro(regK010, "ARQUIVO")
    
    Set dicTitulosK100 = Util.MapearTitulos(regK100, 3)
    Set dicDadosK100 = Util.CriarDicionarioRegistro(regK100, "ARQUIVO")
    
    Set dicTitulosK200 = Util.MapearTitulos(regK200, 3)
    Set dicDadosK200 = Util.CriarDicionarioRegistro(regK200)
    
    Set Plan = PastaDeTrabalho.Worksheets(1)
    
    With Plan
        
        On Error Resume Next
            If .AutoFilterMode Then .AutoFilter.ShowAllData
        On Error GoTo 0
        
        Set dicTitulos = Util.MapearTitulos(Plan, 1)
        
        Mapeamento = 0
        
        If dicTitulos.Exists("COD_ITEM") Then Mapeamento = Mapeamento + 1
        If dicTitulos.Exists("DT_EST") Then Mapeamento = Mapeamento + 1
        If dicTitulos.Exists("QTD") Then Mapeamento = Mapeamento + 1
        If dicTitulos.Exists("IND_EST") Then Mapeamento = Mapeamento + 1
        
        If Mapeamento < 4 Then
            
            Msg = "As colunas principais não foram mapeadas no arquivo selecionado." & vbCrLf & vbCrLf
            Msg = Msg & "Por favor realize o mapeamento dos dados e tente novamente."
            
            Call Util.MsgAlerta(Msg, "Arquivo sem dados mapeados")
            Exit Function
            
        End If
        
        Set Dados = Util.DefinirIntervalo(Plan, 2, 1)
        If Dados Is Nothing Then
            
            Call Util.MsgAlerta("O arquivo selecionado não possui dados de produtos.", "Arquivo sem dados informados")
            Exit Function
            
        End If
        
        ThisWorkbook.Windows(1).Activate
        
        a = 0
        Comeco = Timer()
        For Each Linha In Dados.Rows
            
            Call Util.AntiTravamento(a, 100, "Importando item " & a + 1 & " de " & Dados.Rows.Count, Dados.Rows.Count, Comeco)
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                'Carrega campos-chave do registro e preenche a variável ARQUIVO
                DT_EST = fnExcel.FormatarData(Campos(dicTitulos("DT_EST")))
                COD_ITEM = Campos(dicTitulos("COD_ITEM"))
                IND_EST = Campos(dicTitulos("IND_EST"))
                COD_PART = Campos(dicTitulos("COD_PART"))
                
                ARQUIVO = VBA.Format(DT_EST, "mm/yyyy") & "-" & CNPJContribuinte
                Call GerarRegistrosBlocoK(dicDadosK001, dicDadosK010, dicDadosK100, ARQUIVO, DT_EST)
                For Each Titulo In dicTitulosK200
                    
                    Select Case Titulo
                        
                        Case "REG"
                            arrCampos.Add "K200"
                            
                        Case "ARQUIVO"
                            arrCampos.Add ARQUIVO
                            
                        Case "COD_ITEM"
                            arrCampos.Add fnExcel.FormatarTexto(COD_ITEM)
                            
                        Case "CHV_REG"
                            If dicDadosK100.Exists(ARQUIVO) Then
                                If LBound(dicDadosK100(ARQUIVO)) = 0 Then i = 1 Else i = 1
                                CHV_REG_K100 = dicDadosK100(ARQUIVO)(dicTitulosK100("CHV_REG") - i)
                            End If
                            arrCampos.Add fnSPED.GerarChaveRegistro(CHV_REG_K100, DT_EST, COD_ITEM, IND_EST, COD_PART)
                            
                        Case "CHV_PAI_FISCAL"
                            arrCampos.Add CHV_REG_K100
                            
                        Case "DT_EST"
                            arrCampos.Add fnExcel.FormatarData(Campos(dicTitulos(Titulo)))
                            
                        Case "QTD"
                            arrCampos.Add fnExcel.FormatarValores(Campos(dicTitulos(Titulo)))
                            
                        Case Else
                            arrCampos.Add fnExcel.FormatarTexto(Campos(dicTitulos(Titulo)))
'                            If dicTitulos.Exists(Titulo) Then
'                                arrCampos.Add Campos(dicTitulos(Titulo))
'                            Else
'                                arrCampos.Add ""
'                            End If
                            
                    End Select
                    
                Next Titulo
                
                If Util.ChecarCamposPreenchidos(arrCampos.toArray()) Then arrDados.Add arrCampos.toArray()
                
            End If
PrxLin:
            arrCampos.Clear
            COD_ITEM = ""
            
        Next Linha
        
    End With
    
    Application.DisplayAlerts = False
        PastaDeTrabalho.Close
    Application.DisplayAlerts = True
    
    If arrDados.Count = 0 Then
        Call Util.MsgAlerta("O arquivo selecionado não possui dados de produtos.", "Arquivo sem dados informados")
        Exit Function
    End If
        
    Application.StatusBar = "Exportando dados do registro K001..."
    Call Util.ExportarDadosDicionario(regK001, dicDadosK001)
        
    Application.StatusBar = "Exportando dados do registro K010..."
    Call Util.ExportarDadosDicionario(regK010, dicDadosK010)
        
    Application.StatusBar = "Exportando dados do registro K100..."
    Call Util.ExportarDadosDicionario(regK100, dicDadosK100)
        
    Application.StatusBar = "Exportando dados do registro K200..."
    Call Util.ExportarDadosArrayList(regK200, arrDados)
    
    Application.StatusBar = "Importação concluída!"
    Call Util.MsgInformativa("Cadastro de produtos importado com sucesso!", "Importação do Cadastro de Produtos", Inicio)
    
End Function

Private Function GerarRegistrosBlocoK(ByRef dicDadosK001 As Dictionary, ByRef dicDadosK010 As Dictionary, _
    ByRef dicDadosK100 As Dictionary, ByVal ARQUIVO As String, ByVal DT_EST As String)
    
Dim CHV_REG_0000 As String, CHV_REG_K001$, CHV_REG_K010$, CHV_REG_K100$, Periodo$, DT_INI$, DT_FIN$
    
    Periodo = VBA.Format(DT_EST, "mm/yyyy")
    DT_INI = VBA.Format("01/" & Periodo, "yyyy-mm-dd")
    DT_FIN = VBA.Format(Application.WorksheetFunction.EoMonth(DT_INI, 0), "yyyy-mm-dd")
    
    CHV_REG_0000 = fnSPED.GerarChvReg0000(Periodo)
    CHV_REG_K001 = fnSPED.GerarChaveRegistro(CHV_REG_0000, "K001")
    CHV_REG_K010 = fnSPED.GerarChaveRegistro(CHV_REG_K001, "K010")
    CHV_REG_K100 = fnSPED.GerarChaveRegistro(CHV_REG_K001, "K100")
    
    If Not dicDadosK001.Exists(ARQUIVO) Then
    
        dicDadosK001(ARQUIVO) = Array("K001", ARQUIVO, CHV_REG_K001, CHV_REG_0000, "0")
    
    End If
    
    If Not dicDadosK010.Exists(ARQUIVO) Then
    
        dicDadosK010(ARQUIVO) = Array("K010", ARQUIVO, CHV_REG_K010, CHV_REG_K001, "2")
    
    End If
    
    If Not dicDadosK100.Exists(ARQUIVO) Then
    
        dicDadosK100(ARQUIVO) = Array("K100", ARQUIVO, CHV_REG_K100, CHV_REG_K001, DT_INI, DT_FIN)
    
    End If
    
End Function

Public Function ImportarCadastroTributacaoProdutos()

Dim dtIni As String, dtFim$, ARQUIVO$, CHV_REG_0000$, CHV_REG_0001$, CHV_REG_0190$, COD_ITEM$, UNID_INV$, Msg$
Dim Caminho As Variant, Campos, Titulo
Dim dicTitulosCadTribProd As New Dictionary
Dim dicDados0190 As New Dictionary
Dim dicDadosCadTribProd As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim PastaDeTrabalho As Workbook
Dim arrCampos As New ArrayList
Dim arrDados As New ArrayList
Dim Plan As Worksheet
Dim Mapeamento As Byte

'    If InscContribuinte = "" Then
'        Call Util.MsgAlerta("Informe a Inscrição Estadual do Contribuinte.", "Inscrição Estadual não informada")
'        CadContrib.Activate
'        CadContrib.Range("InscContribuinte").Activate
'        Exit Function
'    End If

    Caminho = Util.SelecionarArquivo("xlsx")
    If VarType(Caminho) = 11 Then Exit Function
    
    Inicio = Now()
    Set PastaDeTrabalho = Workbooks.Open(Caminho)
    ActiveWindow.visible = False
    
    Set dicTitulosCadTribProd = Util.MapearTitulos(assTributacaoICMS, 3)
    Set dicDadosCadTribProd = Util.CriarDicionarioRegistro(assTributacaoICMS)
    
    Set Plan = PastaDeTrabalho.Worksheets(1)
    
    With Plan
    
        On Error Resume Next
            If .AutoFilterMode Then .AutoFilter.ShowAllData
        On Error GoTo 0
        
        Set dicTitulos = Util.MapearTitulos(Plan, 1)
        
        Mapeamento = 0
        
        If dicTitulos.Exists("COD_ITEM") Then Mapeamento = Mapeamento + 1
        If dicTitulos.Exists("COD_NCM") Then Mapeamento = Mapeamento + 1
        If dicTitulos.Exists("CFOP") Then Mapeamento = Mapeamento + 1
        If dicTitulos.Exists("CST_ICMS") Or dicTitulos.Exists("CST_PIS_COFINS") _
            Or dicTitulos.Exists("CST_IPI") Then Mapeamento = Mapeamento + 1
        
        If Mapeamento < 4 Then
        
            Msg = "As colunas principais não foram mapeadas no arquivo selecionado." & vbCrLf & vbCrLf
            Msg = Msg & "Por favor realize o mapeamento dos dados e tente novamente."
            
            Call Util.MsgAlerta(Msg, "Arquivo sem dados mapeados")
            Exit Function
            
        End If
        
        Set Dados = Util.DefinirIntervalo(Plan, 2, 1)
        If Dados Is Nothing Then
        
            Call Util.MsgAlerta("O arquivo selecionado não possui dados de produtos.", "Arquivo sem dados informados")
            Exit Function
        
        End If
        
        ThisWorkbook.Windows(1).Activate
            
        a = 0
        Comeco = Timer()
        For Each Linha In Dados.Rows
            
            Call Util.AntiTravamento(a, 100, "Importando item " & a + 1 & " de " & Dados.Rows.Count, Dados.Rows.Count, Comeco)
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                For Each Titulo In dicTitulosCadTribProd
                    
                    Select Case Titulo
                             
                        Case "TIPO_ITEM"
                            arrCampos.Add ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_TIPO_ITEM(Util.ApenasNumeros(Campos(dicTitulos(Titulo))))
                            
                        Case "IND_MOV"
                            arrCampos.Add ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_MOT_INV(Util.ApenasNumeros(Campos(dicTitulos(Titulo))))
                            
                        Case "COD_ITEM", "COD_BARRA", "COD_NCM", "EX_IPI", "CEST", "COD_BARRA", _
                            "CST_ICMS", "CST_IPI", "CST_PIS_COFINS", "COD_NAT_PIS_COFINS", "COD_ENQ_IPI"
                            arrCampos.Add fnExcel.FormatarTexto(Campos(dicTitulos(Titulo)))

                        Case Else
                            If dicTitulos.Exists(Titulo) Then
                                arrCampos.Add Campos(dicTitulos(Titulo))
                            Else
                                arrCampos.Add ""
                            End If
                            
                    End Select
                    
                Next Titulo
                
                If Util.ChecarCamposPreenchidos(arrCampos.toArray()) Then arrDados.Add arrCampos.toArray()
                
            End If
PrxLin:
            arrCampos.Clear
            
        Next Linha
        
    End With
    
    Application.DisplayAlerts = False
        PastaDeTrabalho.Close
    Application.DisplayAlerts = True
    
    If arrDados.Count = 0 Then
        Call Util.MsgAlerta("O arquivo selecionado não possui dados de produtos.", "Arquivo sem dados informados")
        Exit Function
    End If
    
    Application.StatusBar = "Exportando dados do relatório..."
    Call Util.ExportarDadosArrayList(assTributacaoICMS, arrDados)
    
    Application.StatusBar = "Importação concluída!"
    Call Util.MsgInformativa("Cadastro de tributação de produtos importado com sucesso!", "Importação do Cadastro de Tributação de Produtos", Inicio)
    
End Function

Public Function ImportarCadastroCorrelacoes()

Dim Msg As String, CNPJ_FORNECEDOR$, COD_PROD_FORNEC$, UND_FORNEC$, Chave$
Dim Caminho As Variant, Campos, Titulo
Dim dicTitulosCorrelacoes As New Dictionary
Dim dicDadosCorrelacoes As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim PastaDeTrabalho As Workbook
Dim arrCampos As New ArrayList
Dim arrDados As New ArrayList
Dim Plan As Worksheet
Dim Mapeamento As Byte
    
    Caminho = Util.SelecionarArquivo("xlsx")
    If VarType(Caminho) = 11 Then Exit Function
    
    Inicio = Now()
    
    Set PastaDeTrabalho = Workbooks.Open(Caminho)
    ActiveWindow.visible = False
    
    Set dicTitulosCorrelacoes = Util.MapearTitulos(Correlacoes, 3)
    Set dicDadosCorrelacoes = Util.CriarDicionarioCorrelacoes(Correlacoes)
    Set Plan = PastaDeTrabalho.Worksheets(1)
        
        With Plan
            
            If .AutoFilterMode Then .AutoFilter.ShowAllData
            Set dicTitulos = Util.MapearTitulos(Plan, 1)
            
            Mapeamento = 0
            
            If dicTitulos.Exists("CNPJ_FORNECEDOR") Then Mapeamento = Mapeamento + 1
            If dicTitulos.Exists("COD_PROD_FORNEC") Then Mapeamento = Mapeamento + 1
            If dicTitulos.Exists("UND_FORNEC") Then Mapeamento = Mapeamento + 1
            If dicTitulos.Exists("COD_ITEM") Then Mapeamento = Mapeamento + 1
            If dicTitulos.Exists("UND_INV") Then Mapeamento = Mapeamento + 1
            
            If Mapeamento < 5 Then
            
                Msg = "As colunas principais não foram mapeadas no arquivo selecionado." & vbCrLf & vbCrLf
                Msg = Msg & "Por favor realize o mapeamento dos dados e tente novamente."
                
                Call Util.MsgAlerta(Msg, "Arquivo sem dados mapeados")
                Exit Function
                
            End If
        
            Set Dados = Util.DefinirIntervalo(Plan, 2, 1)
            If Dados Is Nothing Then
            
                Call Util.MsgAlerta("O arquivo selecionado não possui dados de correlações.", "Arquivo sem dados informados")
                Exit Function
            
            End If
            
            a = 0
            Comeco = Timer
            For Each Linha In Dados.Rows
                
                Call Util.AntiTravamento(a, 100, "Importando item " & a + 1 & " de " & Dados.Rows.Count, Dados.Rows.Count, Comeco)
                Campos = Application.index(Linha.Value2, 0, 0)
                If Util.ChecarCamposPreenchidos(Campos) Then
                    
                    For Each Titulo In dicTitulosCorrelacoes
                        If dicTitulos.Exists(Titulo) Then
                            
                            Select Case Titulo
                                
                                Case "CNPJ_FORNECEDOR"
                                    CNPJ_FORNECEDOR = Util.FormatarCNPJ(Campos(dicTitulos(Titulo)))
                                    arrCampos.Add Util.FormatarTexto(Util.FormatarCNPJ(Campos(dicTitulos(Titulo))))
                                    
                                Case "COD_PROD_FORNEC"
                                    COD_PROD_FORNEC = Campos(dicTitulos(Titulo))
                                    arrCampos.Add Util.FormatarTexto(Campos(dicTitulos(Titulo)))
                                    
                                Case "COD_ITEM"
                                    arrCampos.Add Util.FormatarTexto(Campos(dicTitulos(Titulo)))
                                    
                                Case "UND_FORNEC"
                                    UND_FORNEC = Campos(dicTitulos(Titulo))
                                    arrCampos.Add Campos(dicTitulos(Titulo))
                                    Chave = CNPJ_FORNECEDOR & COD_PROD_FORNEC & UND_FORNEC
                                    If dicDadosCorrelacoes.Exists(Chave) Then GoTo PrxLin:

                                Case Else
                                    arrCampos.Add Campos(dicTitulos(Titulo))
                            
                            End Select
                            
                        Else
                            arrCampos.Add ""
                            
                        End If
                        
                    Next Titulo
                    
                    Chave = CNPJ_FORNECEDOR & COD_PROD_FORNEC & UND_FORNEC
                    If Not dicDadosCorrelacoes.Exists(Chave) Then arrDados.Add arrCampos.toArray()
                    
                End If
PrxLin:
                arrCampos.Clear
                
            Next Linha
            
        End With
    
    Application.DisplayAlerts = False
        PastaDeTrabalho.Close
    Application.DisplayAlerts = True
    
    If Mapeamento < 5 Then
        Msg = "As colunas principais ('CNPJ_FORNECEDOR', 'COD_PROD_FORNEC', 'UND_FORNEC', 'COD_ITEM' e 'UND_INV') não foram mapeadas no arquivo selecionado." & vbCrLf & vbCrLf
        Msg = Msg & "Por favor verifique o arquivo selecionado e tente novamente."
        
        Call Util.MsgAlerta(Msg, "Arquivo sem dados mapeados")
        Exit Function
    End If
    
    If arrDados.Count = 0 Then
        Call Util.MsgAlerta("O arquivo selecionado não possui dados de Correlacoes ou todos os itens do arquivo já existem na planilha.", "Arquivo sem dados identificados")
        Exit Function
    End If
    
    Call Util.ExportarDadosArrayList(Correlacoes, arrDados)
    Call Util.MsgInformativa("Correlações importadas com sucesso!", "Importação do Correlação de Produtos", Inicio)
    
End Function

Public Function GerarModelo0200()

Dim Campos As Variant
    
    Campos = Array("COD_ITEM", "DESCR_ITEM", "COD_BARRA", "COD_ANT_ITEM", "UNID_INV", "TIPO_ITEM", "COD_NCM", "EX_IPI", "COD_GEN", "COD_LST", "ALIQ_ICMS", "CEST")
    
    Workbooks.Add
    ActiveSheet.name = "Cadastro de Produtos"
    With ActiveSheet.Range("A1").Resize(, UBound(Campos) + 1)
        
        .value = Campos
        .Font.Bold = True
        .Columns.AutoFit
        
    End With
    
End Function

Public Function GerarModeloA100Filhos()

Dim Campos As Variant
        
    Campos = Array("REG", "CNP_ESTABELECIMENTO", "IND_OPER", "IND_EMIT", "COD_PART", "COD_SIT", "SER", _
        "SUB", "NUM_DOC", "CHV_NFSE", "DT_DOC", "DT_EXE_SERV", "VL_DOC", "IND_PGTO", "VL_DESC", "VL_BC_PIS", _
        "VL_PIS", "VL_BC_COFINS", "VL_COFINS", "VL_PIS_RET", "VL_COFINS_RET", "VL_ISS", "NUM_ITEM", "COD_ITEM", _
        "DESCR_COMPL", "VL_ITEM", "VL_DESC", "NAT_BC_CRED", "IND_ORIG_CRED", "CST_PIS", "VL_BC_PIS", "ALIQ_PIS", _
        "VL_PIS", "CST_COFINS", "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS", "COD_CTA", "COD_CCUS")
        
    Workbooks.Add
    ActiveSheet.name = "Documentos A100 e Filhos"
    
    With ActiveSheet.Range("A1").Resize(, UBound(Campos) + 1)
        
        .value = Campos
        .Font.Bold = True
        .Columns.AutoFit
        
    End With
    
End Function

Public Function ExportarCadastro0200()

Dim Campos As Variant
Dim Plan As Worksheet
Dim dicTitulos As New Dictionary
    
    Set dicTitulos = Util.MapearTitulos(reg0200, 3)
    
    Workbooks.Add
    Set Plan = ActiveSheet
    ActiveSheet.name = "Cadastro de Produtos"
    
    reg0200.Range("A3").CurrentRegion.Copy Plan.Range("A1")
    Plan.Columns(dicTitulos("CHV_PAI_FISCAL")).Delete
    Plan.Columns(dicTitulos("CHV_REG")).Delete
    Plan.Columns(dicTitulos("ARQUIVO")).Delete
    Plan.Columns(dicTitulos("REG")).Delete
    
    Plan.Cells.Columns.AutoFit
        
End Function

Public Function ExportarCadastroTributacaoProdutos()

Dim Campos As Variant
Dim Plan As Worksheet
Dim dicTitulos As New Dictionary
    
    Set dicTitulos = Util.MapearTitulos(assTributacaoICMS, 3)
    
    'Adiciona uma noa pasta de trabalho
    Workbooks.Add
    
    'Seta planilha da nova pasta de trabalho
    Set Plan = ActiveSheet
    
    'Definie nome da nova planilha
    ActiveSheet.name = "Tributação de Produtos"
    
    'Copiar dados da planilha atual para a nova pasta de trabalho
    assTributacaoICMS.Range("A3").CurrentRegion.Copy Plan.Range("A1")
    
    'Elimina as linhas 1 e 2 para remover filtros e subtotais
    Plan.Rows("1:2").Delete
    
    On Error Resume Next
    'Remove as colunas desnecessárias
    Plan.Columns(dicTitulos("SUGESTAO")).Delete
    Plan.Columns(dicTitulos("INCONSISTENCIA")).Delete
    Plan.Columns(dicTitulos("CHV_REG")).Delete
    Plan.Columns(dicTitulos("CHV_PAI_FISCAL")).Delete
    Plan.Columns(dicTitulos("ARQUIVO")).Delete
    Plan.Columns(dicTitulos("ITEM")).Delete
    
    'Redimensiona colunas
    Plan.Cells.Columns.AutoFit
        
End Function

Public Function ExportarDadosRegistro()

Dim Campos As Variant
Dim Plan As Worksheet
Dim dicTitulos As New Dictionary
    
    Set dicTitulos = Util.MapearTitulos(reg0200, 3)
    
    Workbooks.Add
    Set Plan = ActiveSheet
    ActiveSheet.name = "Cadastro de Produtos"
    
    reg0200.Range("A3").CurrentRegion.Copy Plan.Range("A1")
    Plan.Columns(dicTitulos("CHV_PAI_FISCAL")).Delete
    Plan.Columns(dicTitulos("CHV_REG")).Delete
    Plan.Columns(dicTitulos("ARQUIVO")).Delete
    Plan.Columns(dicTitulos("REG")).Delete
    
    Plan.Range("$A$1:$L$1048576").RemoveDuplicates Columns:=Array(1, 2, 5), Header:=xlYes
    Plan.Cells.Columns.AutoFit
        
End Function

Public Function GerarModeloInventario()

Dim Campos As Variant
    
    Campos = Array("COD_ITEM", "UNID", "QTD", "VL_UNIT", "VL_ITEM", "IND_PROP", "COD_PART", "TXT_COMPL", "COD_CTA", "VL_ITEM_IR", "CST_ICMS", "VL_BC_ICMS", "VL_ICMS")
    
    Workbooks.Add
    ActiveSheet.name = "Inventário Físico"
    With ActiveSheet.Range("A1").Resize(, UBound(Campos) + 1)
        
        .value = Campos
        .Font.Bold = True
        .Columns.AutoFit
        
    End With
    
End Function

Public Function IncluirUnid0190(ByRef dicDados0190 As Dictionary, ByVal ARQUIVO As String, ByVal CHV_REG_0001 As String, ByVal CHV_REG_0190 As String, ByVal UNID_INV As String)

    dicDados0190(CHV_REG_0190) = Array("'0190", ARQUIVO, CHV_REG_0190, CHV_REG_0001, UNID_INV, UNID_INV)

End Function
