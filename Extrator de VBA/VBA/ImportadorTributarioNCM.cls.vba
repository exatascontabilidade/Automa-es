Attribute VB_Name = "ImportadorTributarioNCM"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private EnumContribuicoes As New clsEnumeracoesSPEDContribuicoes
Private aTributario As New AssistenteTributario
Private dicTitulosTributacao As New Dictionary
Private dicDadosTributarios As New Dictionary
Private dicDadosNCM As New Dictionary
Private dicTitulos As New Dictionary
Private PlanDestino As Worksheet
Private NomeTributo As String
Private CamposTrib As Variant
Private Campos As Variant
Private COD_NCM As String
Private EX_IPI As String

Public Sub AtualizarTributacaoNCM(ByRef PlanTrib As Worksheet)

Dim PastaTrabalho As Workbook
Dim Plan As Worksheet
Dim UltLin As Long
    
    Call LimparObjetos
    
    UltLin = Util.UltimaLinha(PlanTrib, "A")
    If UltLin = 3 Then
        Call Util.MsgAlerta("Para usar essa funionalidade precisa haver dados na planilha.", "Atualização da Tributação por NCM")
        Call LimparObjetos
        Exit Sub
    End If
    
    Set PastaTrabalho = SelecionarArquivoNCM()
    If PastaTrabalho Is Nothing Then
        Call LimparObjetos
        Exit Sub
    End If
    
    Inicio = Now()
    
    Set PlanDestino = PlanTrib
    NomeTributo = aTributario.ExtrairNomeTributo(PlanTrib)
    
    Set dicTitulosTributacao = Util.MapearTitulos(PlanTrib, 3)
    Set dicDadosTributarios = aTributario.CarregarTributacoesSalvas(PlanTrib)
    
    Set Plan = PastaTrabalho.Worksheets(1)
    Call CarregarTributacaoNCM(PastaTrabalho, Plan)
    
    Call IncluirTributacaoNCM
    
    Call Util.LimparDados(PlanTrib, 4, False)
    Call Util.ExportarDadosDicionario(PlanTrib, dicDadosTributarios)
    
    Call Util.MsgInformativa("Atualização de tributação por NCM concluída com sucesso!", "Atualização por NCM", Inicio)
    Call LimparObjetos
    
End Sub

Private Function SelecionarArquivoNCM() As Workbook

Dim Caminho As Variant
Dim wb As Workbook
    
    Caminho = Util.SelecionarArquivo("xlsx")
    If VarType(Caminho) = vbBoolean And Caminho = False Then Exit Function
    
    On Error GoTo ErroAbrir:
        
        Set wb = Workbooks.Open(Caminho, ReadOnly:=True)
        wb.Windows(1).visible = False
        Set SelecionarArquivoNCM = wb
        
    On Error GoTo 0
    
    Exit Function
    
ErroAbrir:
    Set SelecionarArquivoNCM = Nothing
    
End Function

Private Function CarregarTributacaoNCM(ByRef PastaTrabalho As Workbook, ByRef Plan As Worksheet)

Dim Dados As Range, Linha As Range

    With Plan
        
        On Error Resume Next
            If .AutoFilterMode Then .AutoFilter.ShowAllData
        On Error GoTo 0
        
        Set Dados = Util.DefinirIntervalo(Plan, 2, 1)
        If Dados Is Nothing Then
            
            Call Util.MsgAlerta("O arquivo selecionado não possui dados para importar.", "Cadastro Tributário por NCM")
            Exit Function
            
        End If
        
        PastaTrabalho.Windows(1).Activate
        
        Call dicTitulos.RemoveAll
        Set dicTitulos = Util.MapearTitulos(Plan, 1)
        
        a = 0
        Comeco = Timer()
        For Each Linha In Dados.Rows
            
            Call Util.AntiTravamento(a, 100, "Importando item " & a + 1 & " de " & Dados.Rows.Count, Dados.Rows.Count, Comeco)
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                Call ExtrairTributacao
                
            End If
Prx:
        Next Linha
        
        Call FecharPastaTrabalho(PastaTrabalho)
        
    End With
    
End Function

Private Function ExtrairTributacao()

    Select Case True
        
        Case NomeTributo Like "*ICMS*"
            'Call ExtrairTributacaoICMS
            
        Case NomeTributo Like "*IPI*"
            'Call ExtrairTributacaoIPI
            
        Case NomeTributo Like "*PIS/COFINS*"
            Call ExtrairTributacaoPISCOFINS
            
    End Select

End Function

Private Sub ExtrairTributacaoPISCOFINS()
    
    CamposTrib = DTO_TributacaoNCM_PISCOFINS.ObterCampos()
    
    Call CarregarTributacaoPISCOFINS
    Call AtualizarDicDadosNCM_PIS_COFINS

End Sub

Private Sub CarregarTributacaoPISCOFINS()

Dim Campo As Variant, Valor As String
    
    With tribNCM_PIS_COFINS
        
        DTO_TributacaoNCM_PISCOFINS.ResetarTributarioNCM_PIS_COFINS
        For Each Campo In CamposTrib
            
            If dicTitulos.Exists(Campo) Then
            
                Valor = Campos(dicTitulos(Campo))
                Select Case Campo
                
                    Case "COD_NCM": .COD_NCM = Util.ApenasNumeros(Valor)
                    Case "EX_IPI": .EX_IPI = VBA.Format(Util.ApenasNumeros(Valor), "000")
                    Case "CST_PIS_COFINS_ENT": .CST_PIS_COFINS_ENT = EnumContribuicoes.ValidarEnumeracao_CST_PIS_COFINS(Util.ApenasNumeros(Valor))
                    Case "CST_PIS_COFINS_SAI": .CST_PIS_COFINS_SAI = EnumContribuicoes.ValidarEnumeracao_CST_PIS_COFINS(Util.ApenasNumeros(Valor))
                    Case "ALIQ_PIS": .ALIQ_PIS = fnExcel.FormatarPercentuais(Valor)
                    Case "ALIQ_COFINS": .ALIQ_COFINS = fnExcel.FormatarPercentuais(Valor)
                    Case "COD_NAT_PIS_COFINS": .COD_NAT_PIS_COFINS = "'" & VBA.Format(Util.ApenasNumeros(Valor), "000")
                    
                End Select
                
            End If
            
        Next
        
    End With
    
End Sub

Private Sub AtualizarDicDadosNCM_PIS_COFINS()

Dim Campo As Variant, Valor As String, Key As String
    
    With tribNCM_PIS_COFINS
        
        Key = .COD_NCM & "|" & .EX_IPI
        If Key <> "" Then
            
            If Not dicDadosNCM.Exists(Key) Then Set dicDadosNCM(Key) = New Dictionary
            
            For Each Campo In CamposTrib
                
                Select Case Campo
                    
                    Case "COD_NCM": Valor = .COD_NCM
                    Case "EX_IPI": Valor = .EX_IPI
                    Case "CST_PIS_COFINS_ENT": Valor = .CST_PIS_COFINS_ENT
                    Case "CST_PIS_COFINS_SAI": Valor = .CST_PIS_COFINS_SAI
                    Case "ALIQ_PIS": Valor = .ALIQ_PIS
                    Case "ALIQ_COFINS": Valor = .ALIQ_COFINS
                    Case "COD_NAT_PIS_COFINS": Valor = .COD_NAT_PIS_COFINS
                    Case Else: Valor = ""
                    
                End Select
                
                If Campo <> "COD_NCM" And Campo <> "EX_IPI" Then _
                    If Not dicDadosNCM(Key).Exists(Campo) Then dicDadosNCM(Key)(Campo) = Valor
                
            Next
            
        End If
        
    End With
    
End Sub

Private Sub IncluirTributacaoNCM()
    
    Select Case True
        
        Case NomeTributo Like "*ICMS"
            'Call IncluirTributacaoICMS
            
        Case NomeTributo Like "*IPI"
            'Call IncluirTributacaoIPI
            
        Case NomeTributo Like "*PIS/COFINS*"
            Call IncluirTributacaoPISCOFINS
            
    End Select

End Sub

Private Sub IncluirTributacaoPISCOFINS()

Dim Chave As Variant, Key, Operacao
Dim dicTitulosPISCOFINS As New Dictionary
    
    Set dicTitulosPISCOFINS = Util.MapearTitulos(PlanDestino, 3)
    For Each Chave In dicDadosTributarios.Keys
        
        Operacao = dicDadosTributarios(Chave)
            
            COD_NCM = Operacao(dicTitulosPISCOFINS("COD_NCM"))
            EX_IPI = VBA.Format(Operacao(dicTitulosPISCOFINS("EX_IPI")), "000")
            Key = COD_NCM & "|" & EX_IPI
            If Not dicDadosNCM.Exists(Key) Then GoTo Prx:
                
            Call AtualizarTributacaoPISCOFINS(Operacao, dicTitulosPISCOFINS)
            
        dicDadosTributarios(Chave) = Operacao
        
Prx: Next Chave
    
End Sub

Private Function ListarTributacaoNCM_PISCOFINS()
    
    Call CarregarTributacaoPISCOFINS

End Function

Private Function AtualizarTributacaoPISCOFINS(ByRef Operacao As Variant, ByRef dicTitulos As Dictionary)

Dim CFOP As String, Chave$
    
    CFOP = Operacao(dicTitulos("CFOP"))
    
    Chave = COD_NCM & "|" & EX_IPI
    With tribNCM_PIS_COFINS
        
        Select Case CFOP
            
            Case Is < 4000
                Operacao(dicTitulos("CST_PIS")) = dicDadosNCM(Chave)("CST_PIS_COFINS_ENT")
                Operacao(dicTitulos("CST_COFINS")) = dicDadosNCM(Chave)("CST_PIS_COFINS_ENT")
                
                Call AtribuirAliquotaPISCOFINS(Operacao, dicTitulos)
                
            Case Is > 4000
                Operacao(dicTitulos("CST_PIS")) = dicDadosNCM(Chave)("CST_PIS_COFINS_SAI")
                Operacao(dicTitulos("CST_COFINS")) = dicDadosNCM(Chave)("CST_PIS_COFINS_SAI")
                Operacao(dicTitulos("COD_NAT_PIS_COFINS")) = dicDadosNCM(Chave)("COD_NAT_PIS_COFINS")
                
                Call AtribuirAliquotaPISCOFINS(Operacao, dicTitulos)
                
        End Select
                
    End With
    
End Function

Private Function AtribuirAliquotaPISCOFINS(ByRef Operacao As Variant, ByRef dicTitulos As Dictionary)
    
Dim Chave As String
Dim CST_PIS As String
Dim CST_COFINS As String
    
    Chave = COD_NCM & "|" & EX_IPI
    CST_PIS = Operacao(dicTitulos("CST_PIS"))
    CST_COFINS = Operacao(dicTitulos("CST_COFINS"))
    
    Select Case True
        
        Case CST_PIS Like "5*" Or CST_PIS Like "6*"
            Operacao(dicTitulos("ALIQ_PIS")) = fnExcel.FormatarPercentuais(dicDadosNCM(Chave)("ALIQ_PIS"))
            Operacao(dicTitulos("ALIQ_COFINS")) = fnExcel.FormatarPercentuais(dicDadosNCM(Chave)("ALIQ_COFINS"))
            
        Case CST_PIS Like "7*" Or CST_PIS Like "9*"
            Operacao(dicTitulos("ALIQ_PIS")) = 0
            Operacao(dicTitulos("ALIQ_COFINS")) = 0
            
        Case Else
            Operacao(dicTitulos("ALIQ_PIS")) = fnExcel.FormatarPercentuais(dicDadosNCM(Chave)("ALIQ_PIS"))
            Operacao(dicTitulos("ALIQ_COFINS")) = fnExcel.FormatarPercentuais(dicDadosNCM(Chave)("ALIQ_COFINS"))
            
    End Select
    
End Function

Public Function GerarModeloTributacaoNCM_PISCOFINS()

Dim Campos As Variant
    
    Campos = DTO_TributacaoNCM_PISCOFINS.ObterCampos()
    
    Workbooks.Add
    ActiveSheet.name = "Tributação por NCM PIS e COFINS"
    With ActiveSheet.Range("A1").Resize(, UBound(Campos) + 1)
        
        .value = Campos
        .Font.Bold = True
        .Columns.AutoFit
        
    End With
    
End Function

Private Sub FecharPastaTrabalho(ByRef PastaTrabalho As Workbook)
    
    With Application
        
        .DisplayAlerts = False
            PastaTrabalho.Close
            Set PastaTrabalho = Nothing
        .DisplayAlerts = True
        
    End With
    
End Sub

Public Sub GerarPlanilhaTributacaoNCM(ByRef PlanOrig As Worksheet)

Dim PastaTrabalho As Workbook
Dim PlanDest As Worksheet
Dim dicTitulos As Object
Dim arrRegistros As ArrayList
Dim dicExportacao As Dictionary

    Set dicTitulos = Util.MapearTitulos(PlanOrig, 3)
    NomeTributo = aTributario.ExtrairNomeTributo(PlanOrig)

    ' Carregar todos os registros válidos da planilha em memória
    Set arrRegistros = Util.CriarArrayListRegistro(PlanOrig)
    If arrRegistros.Count = 0 Then
        MsgBox "Não há dados para exportar.", vbExclamation
        Exit Sub
    End If

    Set dicExportacao = MontarTributacaoNCM(arrRegistros, dicTitulos)
    Set PlanDest = FormatarPlanilhaDestino()

    Call Util.ExportarDadosDicionario(PlanDest, dicExportacao, "A2")

    Call fnExcel.FormatarIntervalo(PlanDest.UsedRange, PlanDest, 1)
    PlanDest.Cells.Columns.AutoFit
    PlanDest.Activate
    
    Call LimparObjetos

End Sub

Private Function MontarTributacaoNCM(ByRef arrRegistros As ArrayList, ByRef dicTitulos As Dictionary) As Dictionary

    Select Case True
        
        Case NomeTributo Like "*ICMS"
            'Call AtualizarTributacaoICMS
            
        Case NomeTributo Like "*IPI"
            'Call AtualizarTributacaoIPI
            
        Case NomeTributo Like "*PIS/COFINS*"
            Set MontarTributacaoNCM = MontarTributacaoNCM_PISCOFINS(arrRegistros, dicTitulos)
            
    End Select

End Function

Public Function MontarTributacaoNCM_PISCOFINS(ByRef arrRegistros As ArrayList, ByRef dicTitulos As Object) As Dictionary

Dim Registro As Variant
Dim Campos, Chave As String
Dim arrLinha(1 To 7) As Variant
Dim dictUnico As New Dictionary
Dim arrRegistro As New ArrayList
Dim CFOP As Variant, cstPis As String
Dim cstEntrada As String, cstSaida As String
Dim idxNCM As Long, idxEX_IPI As Long, idxCFOP As Long, idxCST_PIS As Long
Dim idxALIQ_PIS As Long, idxALIQ_COFINS As Long, idxNAT_PIS_COFINS As Long

    idxNCM = dicTitulos("COD_NCM")
    idxEX_IPI = dicTitulos("EX_IPI")
    idxCFOP = dicTitulos("CFOP")
    idxCST_PIS = dicTitulos("CST_PIS")
    idxALIQ_PIS = dicTitulos("ALIQ_PIS")
    idxALIQ_COFINS = dicTitulos("ALIQ_COFINS")
    idxNAT_PIS_COFINS = dicTitulos("COD_NAT_PIS_COFINS")
    
    For Each Campos In arrRegistros

        Dim ncmPadrao As String: ncmPadrao = Util.FormatarTexto(VBA.Format(Campos(idxNCM), "00000000"))
        Dim exIpiPadrao As String: exIpiPadrao = Util.FormatarTexto(VBA.Format(Campos(idxEX_IPI), "000"))
        Dim natPadrao As String: natPadrao = Util.FormatarTexto(VBA.Format(Campos(idxNAT_PIS_COFINS), "000"))
        Dim aliqPisPadrao As Variant: aliqPisPadrao = fnExcel.FormatarPercentuais(Campos(idxALIQ_PIS))
        Dim aliqCofinsPadrao As Variant: aliqCofinsPadrao = fnExcel.FormatarPercentuais(Campos(idxALIQ_COFINS))
        CFOP = Campos(idxCFOP)
        cstPis = Campos(idxCST_PIS)
        cstEntrada = ""
        cstSaida = ""

        If IsNumeric(CFOP) Then
            If CLng(CFOP) < 4000 Then
                If Trim(cstPis) <> "" Then
                    cstEntrada = Util.FormatarTexto(VBA.Format(Util.ApenasNumeros(cstPis), "00"))
                End If
            ElseIf CLng(CFOP) > 4000 Then
                If Trim(cstPis) <> "" Then
                    cstSaida = Util.FormatarTexto(VBA.Format(Util.ApenasNumeros(cstPis), "00"))
                End If
            End If
        End If

        Chave = ncmPadrao & "|" & exIpiPadrao
        If dictUnico.Exists(Chave) Then
            
            Registro = dictUnico(Chave)
            
            If Registro(2) = "" And cstEntrada <> "" Then Registro(2) = cstEntrada
            If Registro(3) = "" And cstSaida <> "" Then Registro(3) = cstSaida
            If Registro(4) = "" And aliqPisPadrao <> "" Then Registro(4) = aliqPisPadrao
            If Registro(5) = "" And aliqCofinsPadrao <> "" Then Registro(5) = aliqCofinsPadrao
            If Registro(6) = "" And natPadrao <> "" Then Registro(6) = natPadrao
            
            dictUnico(Chave) = Registro
            
        Else
            
            arrRegistro.Add ncmPadrao
            arrRegistro.Add exIpiPadrao
            arrRegistro.Add cstEntrada
            arrRegistro.Add cstSaida
            arrRegistro.Add aliqPisPadrao
            arrRegistro.Add aliqCofinsPadrao
            arrRegistro.Add natPadrao
            
            dictUnico.Add Chave, arrRegistro.toArray()
            arrRegistro.Clear
            
        End If
    Next Campos

    Set MontarTributacaoNCM_PISCOFINS = dictUnico

End Function

Private Function FormatarPlanilhaDestino() As Worksheet

Dim PastaTrabalho As Workbook
Dim Plan As Worksheet
    
    Set PastaTrabalho = Workbooks.Add
    Set Plan = PastaTrabalho.ActiveSheet
    
    Select Case True
        
        Case NomeTributo Like "*ICMS"
            'Set FormatarPlanilhaDestino = FormatarPlanilhaDestino_ICMS(Plan)
            
        Case NomeTributo Like "*IPI"
            'Set FormatarPlanilhaDestino = FormatarPlanilhaDestino_IPI(Plan)
            
        Case NomeTributo Like "*PIS/COFINS*"
            Set FormatarPlanilhaDestino = FormatarPlanilhaDestino_PISCOFINS(Plan)
            
    End Select

End Function

Private Function FormatarPlanilhaDestino_PISCOFINS(ByRef Plan As Worksheet) As Worksheet

    Dim TitulosCabecalho As Variant
    TitulosCabecalho = DTO_TributacaoNCM_PISCOFINS.ObterCampos()

    With Plan.Range("A1").Resize(1, UBound(TitulosCabecalho) + 1)

        .value = TitulosCabecalho
        .Font.Bold = True
    
    End With

    Set FormatarPlanilhaDestino_PISCOFINS = Plan

End Function

Private Sub LimparObjetos()
    
    Call Util.AtualizarBarraStatus(False)
    
    Set dicTitulosTributacao = Nothing
    Set dicDadosTributarios = Nothing
    Set dicDadosNCM = Nothing
    Set dicTitulos = Nothing
    Set PlanDestino = Nothing
    
End Sub
