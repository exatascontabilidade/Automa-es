Attribute VB_Name = "AssistenteTributario"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Const RegistrosComplementaresPISCOFINS As String = "C181, C185"
Public PIS_COFINS As New AssistenteTributarioPISCOFINS
Public ICMS As New AssistenteTributarioICMS
Public IPI As New AssistenteTributarioIPI
Public dicEstruturaTributaria As Dictionary
Public dicCFOPS_INCORRETOS As Dictionary
Public dicTitulosTributacao As Dictionary
Public dicDadosTributarios As Dictionary
Public dicTitulosApuracao As Dictionary
Public AtualizarTributacao As Boolean
Public arrTributacoes As ArrayList
Public dicTitulos As Dictionary
Public Campo As Variant

Public Function AtualizarDadosRegistro(ByRef dicDados As Dictionary, ByRef dicTitulosReg As Dictionary, ByRef CamposReg As Variant, _
    ByRef dicTitulosRel As Dictionary, ByRef CamposRel As Variant, ByVal CHV_REG As String, Optional ByRef dicCorrelacoes As Dictionary, _
    Optional ByRef dicCorrelacoesInversas As Dictionary)
    
Dim dicCampos As Variant, Campo
Dim CampoDest As String
Dim i As Byte
    
    'Elimina as informações dos campos INCONSISTÊNCIA e SUGESTÃO
    CamposRel(UBound(CamposRel)) = ""
    CamposRel(UBound(CamposRel) - 1) = ""
    
    If dicDados.Exists(CHV_REG) Then
        
        dicCampos = dicDados(CHV_REG)
        
        Call AtualizarCampos(dicTitulosReg, CamposReg, dicTitulosRel, dicCampos, CamposRel, dicCorrelacoes)
        dicDados(CHV_REG) = dicCampos
        
    Else
        
        dicCampos = CriarRegistro(CamposRel, dicTitulosRel, CamposReg, dicTitulosReg, dicCorrelacoesInversas)
        dicDados(CHV_REG) = dicCampos
        
    End If
    
End Function

Private Function AtualizarCampos(ByRef dicTitulosReg As Dictionary, ByRef CamposReg As Variant, ByRef dicTitulos As Dictionary, _
    ByRef dicCampos As Variant, ByRef CamposRel As Variant, Optional ByRef dicCorrelacoes As Dictionary)

Dim CampoDest As String
Dim Campo As Variant
Dim Valor As Variant
Dim i As Long

    If LBound(dicCampos) = 0 Then i = 1 Else i = 0
    
    For Each Campo In CamposReg
        
        If Not dicCorrelacoes Is Nothing Then
            
            If dicCorrelacoes.Exists(Campo) Then CampoDest = dicCorrelacoes(Campo) Else CampoDest = Campo
            dicCampos(dicTitulosReg(CampoDest) - i) = fnExcel.FormatarTipoDado(Campo, CamposRel(dicTitulos(Campo)))
        Else
    
            dicCampos(dicTitulosReg(Campo) - i) = fnExcel.FormatarTipoDado(Campo, CamposRel(dicTitulos(Campo)))
            
        End If
        
    Next Campo
    
End Function

Private Function CriarRegistro(ByRef CamposRel As Variant, ByRef dicTitulosRel As Dictionary, _
    ByRef CamposReg As Variant, ByRef dicTitulosReg As Dictionary, Optional ByRef dicCorrelacoesInversas As Dictionary)

Dim Titulo As Variant
Dim arrRegistro As New ArrayList

    For Each Titulo In dicTitulosReg.Keys()
        
        If Not dicCorrelacoesInversas Is Nothing Then If dicCorrelacoesInversas.Exists(Titulo) Then Titulo = dicCorrelacoesInversas(Titulo)
            
        If dicTitulosRel.Exists(Titulo) Then
            
            arrRegistro.Add fnExcel.FormatarTipoDado(Titulo, CamposRel(dicTitulosRel(Titulo)))
        
        Else
        
            arrRegistro.Add ""
        
        End If
        
    Next Titulo
    
    CriarRegistro = arrRegistro.toArray()
    
End Function

Public Function AtribuirValor(ByVal Titulo As String, ByVal Valor As Variant)

    Campo(dicTitulos(Titulo)) = Valor

End Function

Public Function RedimensionarArray(ByVal NumCampos As Long)
    
    ReDim Campo(1 To NumCampos - 2) As Variant
    
End Function

Public Function CarregarTributacoesSalvas(ByVal Plan As Worksheet) As Dictionary

Dim Chave As String, VIGENCIA_INICIAL$, VIGENCIA_FINAL$, Vigencia$
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim Titulos As Variant, Campos
Dim dicDados As New Dictionary
Dim b As Long
    
    If Plan.AutoFilterMode Then Plan.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(Plan, 4, 3)
    If Dados Is Nothing Then
        
        Set CarregarTributacoesSalvas = New Dictionary
        Exit Function
        
    End If
    
    b = 0
    Comeco = Timer
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 10, "Carregando tributações salvas", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        ReDim Preserve Campos(LBound(Campos) To UBound(Campos) - 2)
        
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            Chave = GerarChaveTributacao(Plan, Campos)
            VIGENCIA_INICIAL = Campos(dicTitulos("VIGENCIA_INICIAL"))
            VIGENCIA_FINAL = Campos(dicTitulos("VIGENCIA_FINAL"))
            Vigencia = VIGENCIA_INICIAL & "|" & VIGENCIA_FINAL
            
            If Not dicDados.Exists(Chave) Then Set dicDados(Chave) = New Dictionary
            dicDados(Chave)(Vigencia) = Campos
            
         End If
         
    Next Linha
    
    Set CarregarTributacoesSalvas = dicDados
    
End Function

Public Function CarregarTributacoesSalvas_Old(ByVal Plan As Worksheet) As Dictionary

Dim Titulos As Variant, Campos
Dim arrCamposChave As New ArrayList
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim dicDados As New Dictionary
Dim Chave As String
Dim b As Long
    
    If Plan.AutoFilterMode Then Plan.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(Plan, 4, 3)
    If Dados Is Nothing Then
        
        Set CarregarTributacoesSalvas_Old = New Dictionary
        Exit Function
        
    End If
    
    b = 0
    Comeco = Timer
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 10, "Carregando tributações salvas", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        ReDim Preserve Campos(LBound(Campos) To UBound(Campos) - 2)
        
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            Chave = GerarChaveTributacao(Plan, Campos)
            
            dicDados(Chave) = Campos
            arrCamposChave.Clear
            
         End If
         
    Next Linha
    
    Set CarregarTributacoesSalvas_Old = dicDados
    
End Function

Public Function CarregarTributacoesImportadas(ByVal Plan As Worksheet) As Dictionary

Dim Titulos As Variant, Campos
Dim arrCamposChave As New ArrayList
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim dicDados As New Dictionary
Dim Chave As String
Dim b As Long
    
    If Plan.AutoFilterMode Then Plan.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(Plan, 4, 3)
    If Dados Is Nothing Then
        
        Set CarregarTributacoesImportadas = New Dictionary
        Exit Function
        
    End If
    
    b = 0
    Comeco = Timer
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 10, "Carregando tributações importadas", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        ReDim Preserve Campos(LBound(Campos) To UBound(Campos) - 2)
        
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            Chave = GerarChaveImportacao(Plan, Campos)
            
            dicDados(Chave) = Campos
            arrCamposChave.Clear
            
         End If
         
    Next Linha
    
    Set CarregarTributacoesImportadas = dicDados
    
End Function

Public Function ObterNomesCamposChave(ByRef Plan As Worksheet, Optional ByVal ManterdicTitulos As Boolean)
    
    Select Case True
        
        Case Plan.CodeName Like "*IPI"
            If Not ManterdicTitulos Then Set dicTitulos = Util.MapearTitulos(Plan, 3)
            ObterNomesCamposChave = Array("CNPJ_ESTABELECIMENTO", "TIPO_PART", "UF_PART", "COD_ITEM", "CFOP")
            
        Case Plan.CodeName Like "*PISCOFINS"
            If Not ManterdicTitulos Then Set dicTitulos = Util.MapearTitulos(Plan, 3)
            ObterNomesCamposChave = Array("CNPJ_ESTABELECIMENTO", "REGIME_TRIBUTARIO", "TIPO_PART", "UF_PART", "COD_ITEM", "CFOP")
            
        Case Plan.CodeName Like "*ICMS"
            If Not ManterdicTitulos Then Set dicTitulos = Util.MapearTitulos(Plan, 3)
            ObterNomesCamposChave = Array("CNPJ_ESTABELECIMENTO", "UF_CONTRIB", "TIPO_PART", "CONTRIBUINTE", "UF_PART", "COD_ITEM", "CFOP")
            
    End Select
    
End Function

Public Function ObterNomesCamposImportacao(ByRef Plan As Worksheet, Optional ByVal ManterdicTitulos As Boolean)
    
    Select Case True
        
        Case Plan.CodeName Like "*IPI"
            If Not ManterdicTitulos Then Set dicTitulos = Util.MapearTitulos(Plan, 3)
            ObterNomesCamposImportacao = Array("CNPJ_ESTABELECIMENTO", "VIGENCIA_INICIAL", "VIGENCIA_FINAL", "TIPO_PART", "UF_PART", "COD_ITEM", "CFOP")
            
        Case Plan.CodeName Like "*PISCOFINS"
            If Not ManterdicTitulos Then Set dicTitulos = Util.MapearTitulos(Plan, 3)
            ObterNomesCamposImportacao = Array("CNPJ_ESTABELECIMENTO", "VIGENCIA_INICIAL", "VIGENCIA_FINAL", "REGIME_TRIBUTARIO", "TIPO_PART", "UF_PART", "COD_ITEM", "CFOP")
            
        Case Plan.CodeName Like "*ICMS"
            If Not ManterdicTitulos Then Set dicTitulos = Util.MapearTitulos(Plan, 3)
            ObterNomesCamposImportacao = Array("CNPJ_ESTABELECIMENTO", "VIGENCIA_INICIAL", "VIGENCIA_FINAL", "UF_CONTRIB", "TIPO_PART", "CONTRIBUINTE", "UF_PART", "COD_ITEM", "CFOP")
            
    End Select
    
End Function

Public Function RegistrarTributacao(ByRef Plan As Worksheet, ByVal REG As String)

Dim Chave As String
    
    Chave = GerarChaveTributacao(Plan, Campo)
        
    If dicDadosTributarios.Exists(Chave) Then
        
        If RegistrosComplementaresPISCOFINS Like "*" & REG & "*" Then Call ComplementarDadosTributariosPISCOFINS(Chave)
        If AtualizarTributacao Then
            
            Campo = AtualizarCadastroTributario(Plan, Chave)
            dicDadosTributarios(Chave) = Campo
            
        End If
        
    Else
        
        Call AtribuirValor("OBSERVACOES", "RECEM CADASTRADO")
        If dicDadosTributarios Is Nothing Then Set dicDadosTributarios = New Dictionary
        
        If Not dicDadosTributarios.Exists(Chave) Then Set dicDadosTributarios(Chave) = New Dictionary
        dicDadosTributarios(Chave)("|") = Campo
        
    End If
    
End Function

Public Function RegistrarTributacao_Old(ByRef Plan As Worksheet, ByVal REG As String)

Dim Chave As String
    
    Chave = GerarChaveTributacao(Plan, Campo)
        
    If dicDadosTributarios.Exists(Chave) Then
        
        If RegistrosComplementaresPISCOFINS Like "*" & REG & "*" Then Call ComplementarDadosTributariosPISCOFINS(Chave)
        If AtualizarTributacao Then
            
            Campo = AtualizarCadastroTributario(Plan, Chave)
            dicDadosTributarios(Chave) = Campo
            
        End If
        
    Else
        
        Call AtribuirValor("OBSERVACOES", "RECEM CADASTRADO")
        If dicDadosTributarios Is Nothing Then Set dicDadosTributarios = New Dictionary
        If Not dicDadosTributarios.Exists(Chave) Then dicDadosTributarios(Chave) = Campo
        
    End If
    
End Function

Private Function ComplementarDadosTributariosPISCOFINS(ByVal Chave As String)

Dim CamposDic As Variant
Dim Titulos As Variant, Titulo
    
    Titulos = Array("CST_PIS", "CST_COFINS", "ALIQ_PIS", "ALIQ_COFINS", "ALIQ_PIS_QUANT", "ALIQ_COFINS_QUANT")
    CamposDic = dicDadosTributarios(Chave)
    
    For Each Titulo In Titulos
        
        Titulo = dicTitulos(Titulo)
        Select Case True
        
            Case Not IsEmpty(CamposDic(Titulo))
                If Util.VerificarStringVazia(Campo(Titulo)) Then Campo(Titulo) = CamposDic(Titulo)
                
        End Select
        
    Next Titulo
    
End Function

Public Function GerarChaveTributacao(ByRef Plan As Worksheet, ByRef Campos As Variant, Optional ByVal ManterdicTitulos As Boolean) As String

Dim CamposChave As Variant, Campo
Dim arrCamposChave As New ArrayList
    
    CamposChave = ObterNomesCamposChave(Plan, ManterdicTitulos)
    
    'Montar chave do registro
    For Each Campo In CamposChave
        arrCamposChave.Add Util.RemoverAspaSimples(Campos(dicTitulos(Campo)))
    Next Campo
    
    GerarChaveTributacao = VBA.Join(arrCamposChave.toArray())
    
End Function

Public Function GerarChaveImportacao(ByRef Plan As Worksheet, ByRef Campos As Variant, Optional ByVal ManterdicTitulos As Boolean) As String

Dim CamposChave As Variant, Campo
Dim arrCamposChave As New ArrayList
    
    CamposChave = ObterNomesCamposImportacao(Plan, ManterdicTitulos)
    
    For Each Campo In CamposChave
        arrCamposChave.Add Util.RemoverAspaSimples(Campos(dicTitulos(Campo)))
    Next Campo
    
    GerarChaveImportacao = VBA.Join(arrCamposChave.toArray())
    
End Function

Public Function ExtrairCNPJContribuinte(ByVal Valor As String)

    ExtrairCNPJContribuinte = fnExcel.FormatarTexto(VBA.Split(Valor, "-")(1))

End Function

Private Function AtualizarCadastroTributario(ByRef Plan As Worksheet, ByVal Chave As String) As Variant

Dim CamposTributacao As Variant, CamposChave, resultado
Dim Chaves() As String
Dim i As Byte
    
    CamposChave = ObterNomesCamposChave(Plan)
    Chaves = IncluirChavesAdicionais(CamposChave)
    
    CamposTributacao = dicDadosTributarios(Chave)
    
    For i = 1 To dicTitulos.Count - 2
        
        resultado = VBA.Filter(Chaves, dicTitulos.Keys(i - 1))
        If UBound(resultado) = 0 Then
            
            Campo(i) = fnExcel.FormatarTipoDado(dicTitulos.Keys(i - 1), CamposTributacao(i))
            
        End If
        
    Next i
    
    AtualizarCadastroTributario = Campo
    
End Function

Private Function IncluirChavesAdicionais(ByRef CamposChave As Variant) As String()

Dim i As Long
Dim Chaves() As String
Dim ChavesAdicionais As Variant
    
    ChavesAdicionais = Array("VIGENCIA_INICIAL", "VIGENCIA_FINAL", "CFOPS_INCORRETOS", "ALIQ_FCP", "ALIQ_FCP_ST", "ICMS_DESON", "ALIQ_RED_BC_ICMS", "ALIQ_MVA")
    
    ReDim Chaves(1 To UBound(CamposChave) + UBound(ChavesAdicionais) + 1)
    
    'Adiciona Campos-Chave ao array
    For i = LBound(CamposChave) To UBound(CamposChave)
        
        Chaves(i + 1) = CamposChave(i)
        
    Next i
    
    'Adiciona Chavs Adicionais ao array
    For i = LBound(ChavesAdicionais) To UBound(ChavesAdicionais)
        
        Chaves(UBound(CamposChave) + i + 1) = ChavesAdicionais(i)
        
    Next i
    
    IncluirChavesAdicionais = Chaves
    
End Function

Public Sub CarregarEstruturaTributaria(ByRef Plan As Worksheet)

Dim CamposChave As Variant, Campos, CampoChave
Dim Dados As Range, Linha As Range
Dim dicBase As New Dictionary
Dim dicAtual As Dictionary
Dim CFOP As String
Dim i As Long

    CamposChave = ObterNomesCamposChave(Plan)

    With Plan
    
        If .AutoFilterMode Then .AutoFilter.ShowAllData
        
        Set Dados = Util.DefinirIntervalo(Plan, 4, 3)
        If Dados Is Nothing Then Exit Sub
        
        For Each Linha In Dados.Rows
        
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                'Reinicia para o dicionário base em cada linha
                Set dicAtual = dicBase
                
                'Itera pelos campos chave
                For i = LBound(CamposChave) To UBound(CamposChave)
                    
                    CampoChave = Campos(dicTitulosTributacao(CamposChave(i)))
                    If CamposChave(i) = "CFOP" Then
                    
                        CFOP = CampoChave
                        
                        'Armazena os campos diretamente no CFOP
                        dicAtual(CFOP) = Campos
                        Exit For
                        
                    Else
                                            
                        If Not dicAtual.Exists(CampoChave) Then
                            Set dicAtual(CampoChave) = New Dictionary
                        End If
                        
                        Set dicAtual = dicAtual(CampoChave)
                        
                    End If
                    
                Next i
                
            End If
            
        Next Linha
        
    End With
    
    Set dicEstruturaTributaria = dicBase
    
End Sub

Public Function ImportarTributacao(ByRef PlanTrib As Worksheet)

Dim Msg As String, ChaveImportacao$, NomeTributo$
Dim Caminho As Variant, Campos, Titulo
Dim Dados As Range, Linha As Range
Dim PastaDeTrabalho As Workbook
Dim arrDados As New ArrayList
Dim Plan As Worksheet
    
    Caminho = Util.SelecionarArquivo("xlsx")
    If VarType(Caminho) = 11 Then Exit Function
    
    Inicio = Now()
    
    Set PastaDeTrabalho = Workbooks.Open(Caminho)
    ActiveWindow.visible = False
    
    Set dicTitulosTributacao = Util.MapearTitulos(PlanTrib, 3)
    
    Set dicDadosTributarios = CarregarTributacoesSalvas(PlanTrib)
    If dicDadosTributarios Is Nothing Then Set dicDadosTributarios = New Dictionary
    
    Set Plan = PastaDeTrabalho.Worksheets(1)
    Set dicTitulos = Util.MapearTitulos(Plan, 1)
    
    NomeTributo = ExtrairNomeTributo(PlanTrib)
    
    With Plan
        
        On Error Resume Next
            If .AutoFilterMode Then .AutoFilter.ShowAllData
        On Error GoTo 0
        
        If Not ValidarCadastroTributario(PastaDeTrabalho, PlanTrib, Plan) Then Exit Function
        
        Set Dados = Util.DefinirIntervalo(Plan, 2, 1)
        If Dados Is Nothing Then
        
            Call Util.MsgAlerta("O arquivo selecionado não possui dados para importar.", "Cadastro Tributário do " & NomeTributo)
            Exit Function
        
        End If
        
        ThisWorkbook.Windows(1).Activate
            
        a = 0
        Comeco = Timer()
        For Each Linha In Dados.Rows
            
            Call Util.AntiTravamento(a, 100, "Importando item " & a + 1 & " de " & Dados.Rows.Count, Dados.Rows.Count, Comeco)
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                ChaveImportacao = GerarChaveImportacao(PlanTrib, Campos, True)
                If dicDadosTributarios.Exists(ChaveImportacao) Then GoTo Prx:
                
                Call RegistrarCadastroTributario(arrDados, Campos)
                
            End If
Prx:
        Next Linha
        
    End With
    
    Application.DisplayAlerts = False
        PastaDeTrabalho.Close
    Application.DisplayAlerts = True
    
    If arrDados.Count = 0 Then
        Call Util.MsgAlerta("Nenhuma tributação nova foi encontrada no arquivo importado.", "Cadastro Tributário do " & NomeTributo)
        Exit Function
    End If
    
    Application.StatusBar = "Exportando dados do relatório..."
    Call Util.ExportarDadosArrayList(PlanTrib, arrDados)
    
    Application.StatusBar = "Importação concluída!"
    Call Util.MsgInformativa("Tributação importada com sucesso!", "Cadastro Tributário do " & NomeTributo, Inicio)
    
End Function

Public Function ValidarCadastroTributario(ByRef PastaDeTrabalho As Workbook, ByRef PlanTrib As Worksheet, ByRef Plan As Worksheet) As Boolean

Dim Mapeamento As Byte
Dim CamposChave As Variant
Dim Result As VbMsgBoxResult
Dim NomeTributo As String, Msg$
    
    CamposChave = ObterNomesCamposChave(PlanTrib, True)
    Set dicTitulos = Util.MapearTitulos(Plan, 1)
    
    Mapeamento = 0
    For Each Campo In CamposChave
        
        If dicTitulos.Exists(Campo) Then Mapeamento = Mapeamento + 1
        
    Next Campo
    
    If Mapeamento < UBound(CamposChave) + 1 Then
        
        Application.DisplayAlerts = False
            PastaDeTrabalho.Close
        Application.DisplayAlerts = True
        
        Msg = "As colunas principais não foram mapeadas no arquivo selecionado." & vbCrLf & vbCrLf
        Msg = Msg & "Por favor realize o mapeamento dos dados e tente novamente."
        
        NomeTributo = ExtrairNomeTributo(PlanTrib)
        
        Result = Util.MsgDecisao(Msg, "Relatório de Tributação " & NomeTributo)
        If Result = vbNo Then Exit Function
        
        Exit Function
        
    End If
    
    ValidarCadastroTributario = True
    
End Function

Public Function ExtrairNomeTributo(ByRef Plan As Worksheet)

Dim NomeTributo As String
    
    NomeTributo = Plan.CodeName
    Select Case True
        
        Case NomeTributo Like "*ICMS"
            ExtrairNomeTributo = "ICMS"
            
        Case NomeTributo Like "*IPI"
            ExtrairNomeTributo = "IPI"
            
        Case NomeTributo Like "*PISCOFINS"
            ExtrairNomeTributo = "PIS/COFINS"
            
    End Select
    
End Function

Private Function RegistrarCadastroTributario(ByRef arrDados As ArrayList, ByRef Campos As Variant)

Dim Titulo As Variant
Dim IgnorarCampos As String
    
    IgnorarCampos = "INCONSISTENCIA,SUGESTAO"
    Call RedimensionarArray(dicTitulosTributacao.Count)
    
    For Each Titulo In dicTitulosTributacao
        
        If Not IgnorarCampos Like "*" & Titulo & "*" Then _
            AtribuirValor Titulo, fnExcel.FormatarTipoDado(Titulo, Campos(dicTitulos(Titulo)))
        
    Next Titulo
    
    If Util.ChecarCamposPreenchidos(Campo) Then arrDados.Add Campo
    
End Function

Public Function ListarTributacoes(ByRef Plan As Worksheet, ByRef PlanTrib As Worksheet)

Dim i As Byte
Dim Intervalo As Range
Dim arrCampo As ArrayList
Dim dicChaves As New Dictionary
Dim Tributacao As New AssistenteTributario
Dim CamposChave As Variant, Campo
    
    CamposChave = Tributacao.ObterNomesCamposChave(PlanTrib)
    For i = LBound(CamposChave) To UBound(CamposChave) - 1
        
        Set dicChaves(CamposChave(i)) = Util.ObterCampoEspecificoFiltrado(Plan, 4, 3, CamposChave(i))
        
    Next i
    
    If Not PlanTrib.AutoFilter Is Nothing Then PlanTrib.AutoFilterMode = False
    Set Intervalo = Util.DefinirIntervalo(PlanTrib, 3, 3)
    Set dicTitulos = Util.MapearTitulos(PlanTrib, 3)
    
    For i = LBound(CamposChave) To UBound(CamposChave) - 1
        
        Set arrCampo = dicChaves(CamposChave(i))
        If arrCampo.Count > 0 Then Intervalo.AutoFilter Field:=dicTitulos(CamposChave(i)), Criteria1:=arrCampo.toArray, Operator:=xlFilterValues
        
    Next i
    
    If dicChaves.Count > 0 Then
        
        Call Application.GoTo(PlanTrib.[H3], True)
        
    Else
        
        Call Util.MsgAlerta("Registro sem  dados informados.", "Registro sem dados")
        
    End If
    
End Function

Public Function LimparInconsistenciasSugestoes(ByRef Campos As Variant)
    
    Campos(dicTitulosApuracao("INCONSISTENCIA")) = Empty
    Campos(dicTitulosApuracao("SUGESTAO")) = Empty
    
End Function

Sub DestacarNovosCadastros(ByRef Plan As Worksheet)

Dim Formulas(0 To 0) As Variant
Dim formula As Variant
Dim rng As Range
Dim Cor As Long
Dim NovoCadastro As String
    
    'Identifica o endereço da célula INCONSISTÊNCIA
    NovoCadastro = IdentificarEnderecoCampo(Plan, 3, "OBSERVACOES")
    If NovoCadastro <> "" Then
    
        NovoCadastro = VBA.Left(NovoCadastro, VBA.Len(NovoCadastro) - 1) & 4
        Formulas(0) = "=E($A4<>"""";$" & NovoCadastro & "=""RECEM CADASTRADO"")"
        
        'Defina a planilha e a faixa de células onde deseja aplicar a formatação condicional
        Set rng = Util.DefinirIntervalo(Plan, 4, 3)
        If Not rng Is Nothing Then
            
            For Each formula In Formulas
                
                If formula <> "" Then
                    
                    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
                        
                        .Font.Bold = True
                        .Font.Color = RGB(255, 255, 255) ' Branco
                        .Interior.Color = RGB(255, 0, 0) ' Vermelho
                        .SetFirstPriority
                        
                    End With
                    
                End If
                
            Next formula
            
        End If
    
    End If
    
End Sub

Public Sub PrepararTributacoesParaExportacao()

Dim Vigencias As Variant, Campos
    
    Set arrTributacoes = New ArrayList
    
    For Each Vigencias In dicDadosTributarios.Items()
        
        If Not VBA.IsEmpty(Vigencias) Then
            
            For Each Campos In Vigencias.Items()
                
                arrTributacoes.Add Campos
                
            Next Campos
            
        End If
        
    Next Vigencias
    
End Sub

Public Function ExtrairDataReferencia(ByRef Campos As Variant)
    
Dim CFOP As String
    
    CFOP = Campos(dicTitulosApuracao("CFOP"))
    
    Select Case True
        
        Case CFOP > 4000
            ExtrairDataReferencia = Campos(dicTitulosApuracao("DT_DOC"))
            
        Case CFOP < 4000
            ExtrairDataReferencia = Campos(dicTitulosApuracao("DT_ENT_SAI"))
            
    End Select
    
End Function

Public Function ExtrairCamposTributarios(ByVal Chave As String, ByVal DT_REF As String) As Variant
    
Dim Vigencia As Variant, Datas
Dim VigenciaInicial As String, VigenciaFinal$
    
    If dicDadosTributarios.Exists(Chave) Then
        
        For Each Vigencia In dicDadosTributarios(Chave).Keys()
            
            Datas = VBA.Split(Vigencia, "|")
            VigenciaInicial = Datas(0)
            VigenciaFinal = Datas(1)
            
            Select Case True
                
                Case VigenciaInicial <> "" And VigenciaFinal <> ""
                    If CDate(DT_REF) >= CDate(VigenciaInicial) And CDate(DT_REF) <= CDate(VigenciaFinal) Then _
                        ExtrairCamposTributarios = dicDadosTributarios(Chave)(Vigencia): Exit Function
                    
                Case VigenciaInicial <> "" And VigenciaFinal = ""
                    If CDate(DT_REF) >= CDate(VigenciaInicial) Then _
                        ExtrairCamposTributarios = dicDadosTributarios(Chave)(Vigencia): Exit Function
                        
                Case VigenciaInicial = "" And VigenciaFinal <> ""
                    If CDate(DT_REF) <= CDate(VigenciaFinal) Then _
                        ExtrairCamposTributarios = dicDadosTributarios(Chave)(Vigencia): Exit Function
                        
                Case VigenciaInicial = "" And VigenciaFinal = ""
                        ExtrairCamposTributarios = dicDadosTributarios(Chave)(Vigencia): Exit Function
                        
            End Select
            
        Next Vigencia
        
    End If
    
End Function

