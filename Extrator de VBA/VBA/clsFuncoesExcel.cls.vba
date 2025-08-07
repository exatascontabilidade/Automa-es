Attribute VB_Name = "clsFuncoesExcel"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function ConverterValores(ByVal Valor As Variant, Optional Arredondar As Boolean, Optional CasasDecimais As Byte) As Double
    
    If VBA.InStr(Valor, ".") > 0 And VBA.InStr(Valor, ",") > 0 Then Valor = VBA.Replace(Valor, ".", "")
    If Valor = "" Or Valor = "-" Then Valor = 0
    Valor = VBA.Replace(Valor, ".", ",")
    Valor = VBA.Replace(Valor, "'", "")
    If Arredondar Then Valor = VBA.Round(Valor, CasasDecimais)
    ConverterValores = Valor
    
End Function

Public Function ConverterPercentuais_Old(ByVal Valor As String, Optional Inteiro As Boolean) As Double
    
    Valor = VBA.Replace(Valor, ".", ",")
    If Valor = "" Or Valor = "-" Then Valor = 0
    If VBA.InStr(1, Valor, "%") Then Valor = VBA.Replace(Valor, "%", "") / 100
    If Inteiro Then Valor = Valor * 100
    
    ConverterPercentuais_Old = Valor
    
End Function

Public Function FormatarValores(ByVal Valor As Variant, Optional Arredondar As Boolean, Optional CasasDecimais As Byte) As Double
    
    If VBA.InStr(Valor, ".") > 0 And VBA.InStr(Valor, ",") > 0 Then Valor = VBA.Replace(Valor, ".", "")
    Valor = VBA.Replace(Valor, ".", ",")
    If VBA.Trim(Valor) = "" Or Valor = "-" Then Valor = 0
    If Arredondar Then Valor = VBA.Round(Valor, CasasDecimais)

    FormatarValores = VBA.Format(Valor, "##0.#0")
    
End Function

Public Function FormatarPercentuais(ByVal Valor As String) As Double

Dim pPonto As Integer
Dim pVirgula As Integer
    
    Valor = Trim(Replace(Valor, "%", ""))
    
    If Valor = "" Or Valor = "-" Then
        FormatarPercentuais = 0
        GoTo Finalizar
    End If
    
    pPonto = VBA.InStrRev(Valor, ".")
    pVirgula = VBA.InStrRev(Valor, ",")
    
    Select Case True
        
        Case pVirgula > pPonto
            Valor = VBA.Replace(Valor, ".", "")
            
        Case pPonto > pVirgula
            Valor = VBA.Replace(Valor, ",", "")
            
    End Select
    
    Valor = VBA.Replace(Valor, ".", ",")
        
    If VBA.IsNumeric(Valor) Then FormatarPercentuais = CDbl(Valor) / 100 _
        Else FormatarPercentuais = 0
        
Exit Function
Finalizar:

End Function

Public Function FormatarTexto(ByVal Texto As Variant) As String
    
    If Not IsEmpty(Texto) Then
        
        Texto = VBA.Replace(Texto, "'", "")
        FormatarTexto = "'" & Texto
        
    End If
    
End Function

Public Function FormatarData(ByVal Data As Variant) As String
    
Dim DataTeste
    
    If Data <> "" And Not Data Like "*-*" And Not Data Like "*/*" And VBA.Len(Data) = 8 Then

        DataTeste = VBA.Format(Data, "00/00/0000")
        If IsDate(DataTeste) Then Data = Format(DataTeste, "yyyy-mm-dd")

    End If
    
    FormatarData = Format(Data, "yyyy-mm-dd")
    
End Function

Public Function DefinirTipoCampos(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, Optional SPEDContr As Boolean)

Dim nReg As String
Dim Titulo As Variant
Dim Posicao As Byte, i As Byte
Dim arrCampos As New ArrayList
    
    nReg = Campos(LBound(Campos))
    If LBound(Campos) = 0 Then i = 1
    If arrEnumeracoesSPEDFiscal.Count = 0 Then Call ValidacoesSPED.Fiscal.Enumeracoes.ListarEnumeracoes
    For Each Titulo In dicTitulos.Keys()
        
        If Titulo = "" Then GoTo Prx:
        Posicao = dicTitulos(Titulo) - i
        
        If Posicao > UBound(Campos) Then Exit For
        Select Case True
            
            Case Len(Campos(Posicao)) > 254
                arrCampos.Add Util.FormatarTexto(VBA.Left(Campos(Posicao), 253))
                
            Case arrEnumeracoesSPEDFiscal.contains(Titulo)
                If Campos(Posicao) <> "" Then
                    
                    If SPEDContr Then
                        arrCampos.Add ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracoes(nReg, Titulo, Campos(Posicao))
                    Else
                        arrCampos.Add ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracoes(nReg, Titulo, Campos(Posicao))
                    End If
                    
                Else
                    
                    arrCampos.Add ""
                    
                End If
                
            Case Titulo Like "ESTQ_ABERT", Titulo Like "VOL_ENTR", Titulo Like "VOL_DISP", Titulo Like "VOL_SAIDAS", _
                Titulo Like "ESTQ_ESCR", Titulo Like "VAL_AJ_PERDA", Titulo Like "VAL_AJ_GANHO", Titulo Like "FECH_FISICO"
                If Campos(Posicao) <> "" Then arrCampos.Add fnExcel.ConverterValores(Campos(Posicao), True, 3) Else arrCampos.Add ""
            
            Case Titulo Like "VL_*", Titulo Like "QTD*", Titulo Like "DEB_ESP*", Titulo Like "QUANT_*", Titulo Like "SLD_*", _
                Titulo Like "CRED_*", VBA.UCase(Titulo) Like "VALOR", Titulo Like "VOL_*", Titulo Like "ESTQ_*", Titulo Like "VLR_*", _
                Titulo Like "VAL_*", Titulo Like "FECH_*", Titulo Like "QTD", Titulo Like "ALIQ_*QUANT*", Titulo Like "GT_*", _
                Titulo Like "FAT_*", Titulo Like "*_ORIGINAL", Titulo Like "*_CORRIGIDO", Titulo Like "DIFERENCA_*"
                If Campos(Posicao) <> "" Then arrCampos.Add fnExcel.FormatarValores(Campos(Posicao)) Else arrCampos.Add ""
                
'            Case Titulo Like "ALIQ_PIS"
'                If Campos(Posicao) <> "" Then arrCampos.Add fnExcel.FormatarPercentuaisPIS(Campos(Posicao)) Else arrCampos.Add ""
                
            Case Titulo Like "ALIQ_*", Titulo Like "ALIQ_MARGEM"
                If Campos(Posicao) <> "" Then arrCampos.Add fnExcel.FormatarPercentuais(Campos(Posicao)) Else arrCampos.Add 0
                
            Case Titulo Like "Data de Emissao" Or Titulo Like "Lançada SPED" Or Titulo Like "DT_*"
                If Campos(Posicao) <> "" Then arrCampos.Add fnExcel.FormatarData(Campos(Posicao)) Else arrCampos.Add ""
                
            Case Titulo Like "CFOP", Titulo Like "NUM_ITEM"
                If Campos(Posicao) <> "" Then arrCampos.Add CInt(Campos(Posicao)) Else arrCampos.Add ""
                
            Case Titulo Like "COD_ITEM*"
                If Campos(Posicao) <> "" Then arrCampos.Add "'" & Campos(Posicao) Else arrCampos.Add ""
            
            Case Titulo Like "SUGESTAO*", Titulo Like "INCONSISTENCIA*"
                arrCampos.Add Campos(Posicao)
                
            Case Else
                If Campos(Posicao) <> "" Then arrCampos.Add Util.FormatarTexto(Campos(Posicao)) Else arrCampos.Add ""
                
        End Select
        
Prx:
    Next Titulo
    
    Campos = arrCampos.toArray()
    
End Function

Public Function FormatarTipoDado(ByRef nCampo As Variant, ByRef Valor As Variant)
    
    Select Case True
        
        Case nCampo Like "TIPO_ITEM"
            If Valor <> "" Then Valor = _
                ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_TIPO_ITEM(Util.ApenasNumeros(Valor))
        
        Case nCampo Like "IND_MOV"
            If Valor <> "" Then Valor = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_C170_IND_MOV(Util.ApenasNumeros(Valor))
        
        Case nCampo Like "CST_PIS" Or nCampo Like "CST_COFINS"
            If Valor <> "" Then Valor = _
                ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(Util.ApenasNumeros(Valor))
            
        Case nCampo Like "CST_IPI"
            If Valor <> "" Then Valor = _
                ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_IPI(Util.ApenasNumeros(Valor))
                
        Case nCampo Like "*QUANT", nCampo Like "VL_*", nCampo Like "QTD*", nCampo Like "FAT_*", _
            nCampo Like "*_ORIGINAL", nCampo Like "*_CORRIGIDO", nCampo Like "DIFERENCA_*"
            If Valor <> "" Then Valor = fnExcel.FormatarValores(Valor)
            
        Case nCampo Like "ALIQ_*"
            If Valor <> "" Then Valor = fnExcel.FormatarPercentuais(Valor)
            
        Case nCampo Like "DT_*"
            If Valor <> "" Then Valor = fnExcel.FormatarData(Valor)
            
        Case Else
            If Valor <> "" Then Valor = Util.FormatarTexto(Valor)
            
    End Select
        
    FormatarTipoDado = Valor
    
End Function

Public Function FormatarCampoFiltrado(ByRef nCampo As Variant, ByRef Valor As Variant, Optional TratarAliquotaPIS As Boolean = False)
    
    Select Case True
        
        Case nCampo Like "CST_PIS" Or nCampo Like "CST_COFINS"
            If Valor <> "" Then Valor = _
                ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(Util.ApenasNumeros(Valor))
                
        Case nCampo Like "CST_ICMS"
            If Valor <> "" Then Valor = Util.ApenasNumeros(Valor)
            
        Case nCampo Like "CST_IPI"
            If Valor <> "" Then Valor = _
                ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_IPI(Util.ApenasNumeros(Valor))
                
        Case nCampo Like "ALIQ_*"
            If Valor <> "" Then Valor = VBA.Format(Valor, "#0.00%")
            
    End Select
    
    FormatarCampoFiltrado = Valor
    
End Function

Public Function FormatarIntervalo(ByRef Intervalo As Range, ByRef Plan As Worksheet, Optional LinhaTitulo As Long = 3)

Dim arrCamposFomatadosEsquerda As New ArrayList
Dim dicTitulos As New Dictionary
Dim Campo As Variant
Dim i As Integer
    
    If Intervalo Is Nothing Then Exit Function
    Set arrCamposFomatadosEsquerda = CarregarCamposFormatadosEsquerda()
    
    Set dicTitulos = Util.MapearTitulos(Plan, LinhaTitulo)
    
    Intervalo.VerticalAlignment = xlCenter
    Intervalo.HorizontalAlignment = xlCenter
    For i = 1 To dicTitulos.Count
        
        Campo = dicTitulos.Keys(i - 1)
        Select Case True
            
            Case Campo Like "ESTQ_ABERT", Campo Like "VOL_ENTR", Campo Like "VOL_DISP", Campo Like "VOL_SAIDAS", Campo Like "ESTQ_ESCR", _
                Campo Like "VAL_AJ_PERDA", Campo Like "VAL_AJ_GANHO", Campo Like "FECH_FISICO", Campo Like "VAL_FECHA", Campo Like "VAL_ABERT", _
                Campo Like "VOL_AFERI", Campo Like "VOL_VENDAS"
                Intervalo.Columns(i).NumberFormat = "#,##0.000;-#,##0.000;""-"";@"
                Intervalo.Columns(i).HorizontalAlignment = xlRight
                
            Case Campo Like "VL_*", Campo Like "DEB_ESP*", Campo Like "SLD_*", Campo Like "CRED_*", VBA.UCase(Campo) Like "VALOR", _
                Campo Like "ALIQ_*QUANT*", Campo Like "QTD_*", Campo Like "QUANT_*", Campo Like "VOL_*", Campo Like "ESTQ_*", _
                Campo Like "VAL_*", Campo Like "FECH_*", Campo Like "QTD", Campo Like "*_ORIGINAL", Campo Like "*_CORRIGIDO", _
                Campo Like "DIFERENCA_*", Campo Like "VLR_*", Campo Like "GT_*"
                Intervalo.Columns(i).Style = "Comma"
                
            Case Campo Like "ALIQ_*"
                Intervalo.Columns(i).Style = "Percent"
                Intervalo.Columns(i).NumberFormat = "#0.00%"
                
            Case Campo Like "Data de Emissao" Or Campo Like "Lançada SPED" Or Campo Like "DT_*" Or Campo Like "Data" Or Campo Like "Data_*"
                Intervalo.Columns(i).NumberFormat = "dd/mm/yyyy"
                
            Case Campo Like "CFOP"
                Intervalo.Columns(i).NumberFormat = "0"
                
            Case arrCamposFomatadosEsquerda.contains(Campo)
                Intervalo.Columns(i).HorizontalAlignment = xlLeft
                
            Case Else
                Intervalo.Columns(i).NumberFormat = "@"
                
        End Select
        
    Next i
    
End Function

Private Function CarregarCamposFormatadosEsquerda() As ArrayList

Dim arrCampos As New ArrayList
Dim Campos As Variant, Campo
    
    Campos = Array("CST_PIS", "CST_COFINS", "IND_FRT", "RAZAO_FORNECEDOR", "DESCR_PROD_FORNECEDOR", "DESCR_ITEM", _
        "Razao Social Emitente", "DESCR_ITEM", "TIPO_ITEM", "NOME", "ENDERECO", "IND_NAT_PJ", "IND_ATIV", "COD_SIT", _
        "DESCR_COMPL", "INCONSISTENCIA", "SUGESTAO", "Razao Social Emitente", "Razao Social Destinatário", _
        "DESCR_ITEM_NF", "DESCR_ITEM_SPED", "TXT_COMPL", "CST_PIS_NF", "CST_PIS_SPED", "CST_COFINS_NF", "NOME_RAZAO", _
        "CST_COFINS_SPED", "COD_SIT_NF", "COD_SIT_SPED", "NOME_NF", "NOME_SPED", "NAT_BC_CRED", "COD_NAT", "CST_IPI", _
        "REGIME_TRIBUTARIO", "RECOMENDACAO")
        
    For Each Campo In Campos
        
        If Not arrCampos.contains(Campo) Then arrCampos.Add Campo
        
    Next Campo
    
    Set CarregarCamposFormatadosEsquerda = arrCampos
    
End Function

Public Sub GerarRelatorioQuebraSequenciaXML()

Dim Msg As String, COD_MOD$, SERIE$, CNPJEmit$
Dim dicChaves As New Dictionary
Dim arrQuebras As New ArrayList
Dim arrChaves As New ArrayList
Dim Inicio As Date
    
    Inicio = Now()
    
    Application.StatusBar = "Importando chaves de acesso, por favor aguarde..."
    
    CNPJEmit = Util.FormatarCNPJ(CadContrib.Range("CNPJContribuinte").value)
    
    If CNPJEmit = "" Then
    
        Call Util.MsgAlerta("Informe o CNPJ do Contribuinte para gerar o relatório de quebras de sequência.", "Quebras de Sequência")
        CadContrib.Activate
        CadContrib.Range("CNPJContribuinte").Select
        Exit Sub
        
    End If
    
    'Carrega chaves de acesso dos documentos
    arrChaves.addRange Util.CarregarChavesEmitentes(EntNFe, CNPJEmit, "Chave de Acesso")
    arrChaves.addRange Util.CarregarChavesEmitentes(EntCTe, CNPJEmit, "Chave de Acesso")
    arrChaves.addRange Util.CarregarChavesEmitentes(SaiNFe, CNPJEmit, "Chave de Acesso")
    arrChaves.addRange Util.CarregarChavesEmitentes(SaiNFCe, CNPJEmit, "Chave de Acesso")
    arrChaves.addRange Util.CarregarChavesEmitentes(SaiCTe, CNPJEmit, "Chave de Acesso")
    arrChaves.addRange Util.CarregarChavesEmitentes(SaiCFe, CNPJEmit, "Chave de Acesso")
    
    If arrChaves.Count = 0 Then
        
        Msg = "Não existem chaves de acesso para processar" & vbCrLf & vbCrLf
        Msg = Msg & "Por favor importe XMLS de saída na guia 'Omissões Fiscais' para usar este recurso."
        
        Call Util.MsgAviso(Msg, "Relatório de Quebra de Sequência")
        Exit Sub
        
    End If
    
    'Processa as chaves de aceeo para encontrar as quebras de sequencia
    Call AtribuirNotas(arrChaves, dicChaves)
    
    'Gera relatório de quebra de sequencias
    Call VerificarQuebraSequencia(dicChaves, arrQuebras)
    
    'Limpa dados do relatório de quebra de sequência
    Call Util.LimparDados(QuebraSequencia, 4, False)
    
    'Exporta resultado para planilha
    Call Util.ExportarDadosArrayList(QuebraSequencia, arrQuebras)
    
    If arrQuebras.Count = 0 Then
        
        Msg = "Não foram encontradas quebras de sequência nas chaves informadas"
        
        Call Util.MsgAviso(Msg, "Relatório de Quebra de Sequência XML")
        Exit Sub
        
    End If
    
    QuebraSequencia.Activate
    
    Application.StatusBar = "Processo concluído com sucesso!"
    Call Util.MsgInformativa("Relatório de quebra de sequência gerado com sucesso!", "Relatório de Quebra de Sequência XML", Inicio)
    
    Application.StatusBar = False
    
End Sub

Public Sub GerarRelatorioQuebraSequenciaSPED()

Dim Msg As String, COD_MOD$, SERIE$, CNPJEmit$
Dim dicChaves As New Dictionary
Dim arrQuebras As New ArrayList
Dim arrChaves As New ArrayList
Dim Inicio As Date
    
    Inicio = Now()
    
    Application.StatusBar = "Importando chaves de acesso, por favor aguarde..."
    
    CNPJEmit = Util.FormatarCNPJ(CadContrib.Range("CNPJContribuinte").value)
    
    If CNPJEmit = "" Then
    
        Call Util.MsgAlerta("Informe o CNPJ do Contribuinte para gerar o relatório de quebras de sequência.", "Quebras de Sequência")
        CadContrib.Activate
        CadContrib.Range("CNPJContribuinte").Select
        Exit Sub
        
    End If
    
    'Carrega chaves de acesso dos documentos
    arrChaves.addRange Util.CarregarChavesEmitentes(regB020, CNPJEmit, "CHV_NFE")
    arrChaves.addRange Util.CarregarChavesEmitentes(regC100, CNPJEmit, "CHV_NFE")
    arrChaves.addRange Util.CarregarChavesEmitentes(regC500, CNPJEmit, "CHV_NFE")
    
    arrChaves.addRange Util.CarregarChavesEmitentes(regC465, CNPJEmit, "CHV_CFE")
    arrChaves.addRange Util.CarregarChavesEmitentes(regC800, CNPJEmit, "CHV_CFE")
    
    arrChaves.addRange Util.CarregarChavesEmitentes(regD100, CNPJEmit, "CHV_CTE")
    
    If arrChaves.Count = 0 Then
        
        Msg = "Não existem chaves de acesso para processar" & vbCrLf & vbCrLf
        Msg = Msg & "Por favor importe o SPED Fiscal ou Contribuições para usar este recurso."
        
        Call Util.MsgAviso(Msg, "Relatório de Quebra de Sequência SPED")
        Exit Sub
        
    End If
    
    'Processa as chaves de aceeo para encontrar as quebras de sequencia
    Call AtribuirNotas(arrChaves, dicChaves)
    
    'Gera relatório de quebra de sequencias
    Call VerificarQuebraSequencia(dicChaves, arrQuebras)
    
    'Limpa dados do relatório de quebra de sequência
    Call Util.LimparDados(QuebraSequencia, 4, False)
    
    'Exporta resultado para planilha
    Call Util.ExportarDadosArrayList(QuebraSequencia, arrQuebras)
    
    If arrQuebras.Count = 0 Then
        
        Msg = "Não foram encontradas quebras de sequência nas chaves informadas"
        
        Call Util.MsgAviso(Msg, "Relatório de Quebra de Sequência")
        Exit Sub
        
    End If
    
    QuebraSequencia.Activate
    
    Application.StatusBar = "Processo concluído com sucesso!"
    Call Util.MsgInformativa("Relatório de quebra de sequência gerado com sucesso!", "Relatório de Quebra de Sequência SPED", Inicio)
    
    Application.StatusBar = False
    
End Sub

Private Function AtribuirNotas(ByRef arrChaves As ArrayList, ByRef dicChaves As Dictionary)

Dim COD_MOD As String, SERIE$
Dim chvDoc As Variant
Dim Chave As String
Dim NF As Long
    
    For Each chvDoc In arrChaves
        
        COD_MOD = VBA.Mid(chvDoc, 21, 2)
        SERIE = VBA.Mid(chvDoc, 23, 3)
        NF = VBA.Mid(chvDoc, 26, 9)
        
        Chave = COD_MOD & "/" & SERIE
        
        If Not dicChaves.Exists(Chave) Then
            
            'Inicializar ArrayList, mínimo e máximo
            Set dicChaves(Chave) = New Dictionary
            
            'Adicionar elementos ao dicionário
            Set dicChaves(Chave)("lista") = New ArrayList
            
            If Not dicChaves(Chave)("lista").contains(NF) Then dicChaves(Chave)("lista").Add NF
            dicChaves(Chave)("minimo") = NF
            dicChaves(Chave)("maximo") = NF
            
        End If
        
        'Se o Modelo/Série já existir, atualize o mínimo e máximo conforme necessário
        With dicChaves(Chave)
            
            'Adiciona número da nota a lista
            .item("lista").Add NF
            
            'Atualiza números mínimos e máximos
            If NF < .item("minimo") Then .item("minimo") = NF
            If NF > .item("maximo") Then .item("maximo") = NF
            
        End With
        
    Next chvDoc
    
End Function

Private Function VerificarQuebraSequencia(ByRef dicNotas As Dictionary, ByRef arrQuebras As ArrayList) As ArrayList
    
Dim NUM As Long
Dim Chave As Variant
Dim arrLista As ArrayList
        
    For Each Chave In dicNotas.Keys()
            
        Set arrLista = dicNotas(Chave)("lista")
        
        'Percorre o intervalo en tre o número mínimo e o máximo encontrado
        For NUM = dicNotas(Chave)("minimo") To dicNotas(Chave)("maximo")
            
            'Verifica se a nota está na lista
            If Not arrLista.contains(NUM) Then
                
                'Adiciona o número a lista de quebra de sequência
                arrQuebras.Add Array(Chave, NUM)
            
            End If
                        
        Next NUM
        
    Next Chave

End Function

Public Sub ClassificarColuna(ByRef Plan As Worksheet, dicTitulos As Dictionary, ByVal LinT As Integer, _
    ByVal OrdemDescendente As Boolean, ParamArray Campos() As Variant)
    
Dim Campo As Variant
Dim col As Integer
Dim Ordem As Variant
    
    If OrdemDescendente = True Then Ordem = xlDescending Else Ordem = xlAscending
    
    With Plan.AutoFilter.Sort
        
        .SortFields.Clear
        For Each Campo In Campos
            
            col = dicTitulos(Campo)
            
            .SortFields.Add2 Key:=Range(Cells(LinT, col), Cells(Rows.Count, col)), SortOn:=xlSortOnValues, Order:=Ordem, DataOption:=xlSortNormal
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            
        Next Campo
    
    End With
            
End Sub
