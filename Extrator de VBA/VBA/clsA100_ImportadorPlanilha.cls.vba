Attribute VB_Name = "clsA100_ImportadorPlanilha"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private EnumContribuicoes As New clsEnumeracoesSPEDContribuicoes
Private EnumFiscal As New clsEnumeracoesSPEDFiscal
Private GerenciadorSPED As New clsRegistrosSPED
Private importPlan As ImportadorPlanilhas
Private ExpReg As ExportadorRegistros
Private arrDadosRegistro As ArrayList
Private dicTitulosOrig As Dictionary
Private dicLayoutA100 As Dictionary
Private dicLayoutA170 As Dictionary
Private PastaTrabalho As Workbook
Private arrDadosOrig As ArrayList
Private arrCampos As ArrayList
Private PlanOrig As Worksheet
Private Periodo As String
Private ARQUIVO As String
Private total As Integer

Public Sub ImportarPlanilha()
    
    Inicio = Now()
    total = 0
    
    If Not Funcoes.CarregarDadosContribuinte Then Exit Sub
    
    Call InicializarObjetos
    If Not SetarPlanilhaOrigem Then Exit Sub
    
    Call Util.AtualizarBarraStatus("Iniciando importação dos registros...")
    If Not ProcessarRegistrosPlanilha Then Exit Sub
    
    Call ExpReg.ExportarRegistros("A001", "A010", "A100", "A170")
    
    Call Util.MsgInformativa("Registros importados com sucesso!", "Importação NFSe", Inicio)
    
    Call LimparObjetos
    
    Debug.Print total
    
End Sub

Private Function ValidarPeriodoImportacao() As Boolean

Dim Mensagem As String
Dim Titulo As String
    
    If PeriodoImportacao = "" Then
        
        Mensagem = "Informe o período ('MMAAAA') que deseja inserir os itens para prosseguir com a importação."
        Titulo = "Período de importação não informado"
        
        Call Util.MsgAlerta(Mensagem, Titulo)
        Exit Function
        
    Else
        
        Periodo = VBA.Format(PeriodoImportacao, "00/0000")
        ARQUIVO = Periodo & "-" & CNPJContribuinte
        If VBA.IsDate(Periodo) Then ValidarPeriodoImportacao = True
        
    End If
    
End Function

Private Function SetarPlanilhaOrigem() As Boolean
    
    Set PastaTrabalho = importPlan.AbrirPastaTrabalho
    If Not PastaTrabalho Is Nothing Then
        
        Set PlanOrig = PastaTrabalho.Worksheets(1)
        SetarPlanilhaOrigem = True
        
    End If
    
End Function

Private Function ProcessarRegistrosPlanilha() As Boolean

Dim b As Long
Dim Comeco As Double
Dim Campos As Variant
    
    Set dicTitulosOrig = Util.MapearTitulos(PlanOrig, 1)
    Set arrDadosOrig = Util.CriarArrayListRegistro(PlanOrig, 2, 1)
    
    If ChecarAusenciaDados(arrDadosOrig) Then Exit Function
    
    Call FecharPastaTrabalho
    Call CarregarRegistrosSPED
    
    b = 0
    Comeco = Timer
    For Each Campos In arrDadosOrig
        
        Call Util.AntiTravamento(b, 100, "Importando Campos " & b + 1 & " de " & arrDadosOrig.Count, arrDadosOrig.Count, Comeco)
        If Util.ChecarCamposPreenchidos(Campos) Then Call ProcessarRegistro(Campos)
        
    Next Campos
    
    ProcessarRegistrosPlanilha = True
    
End Function

Private Sub ProcessarRegistro(ByRef Campos As Variant)

Dim Chave As String
    
    Call ResetarCamposA100
    
    Chave = GerarChaveA100(dicTitulosOrig, Campos)
    If Not dtoRegSPED.rA100.Exists(Chave) Then Call CriarRegistroA100(Campos) Else Stop
    
End Sub

Private Function CriarRegistroA100(ByVal Campos As Variant) As Boolean

Dim i As Integer
Dim Campo As Variant, CamposA100
Dim CHV_REG As String, CHV_PAI_CONTRIBUICOES$, Chave$, CNPJEstabelecimento$, IND_OPER$, IND_EMIT$, COD_SIT$, IND_PGTO$, DT_DOC$, DT_EXE_SERV$, DT_REF$
    
    Call arrCampos.Clear
    i = Util.VerificarPosicaoInicialArray(Campos)
    
    With CamposA100
        
        For Each Campo In dtoTitSPED.tA100.Keys()
            
            Select Case True
                
                Case Campo = "REG"
                    arrCampos.Add "A100"
                    
                Case Campo = "ARQUIVO"
                    DT_DOC = ExtrairCampo(dicTitulosOrig, Campos, "DT_DOC")
                    DT_EXE_SERV = ExtrairCampo(dicTitulosOrig, Campos, "DT_EXE_SERV")
                    DT_REF = IIf(DT_EXE_SERV <> "", DT_EXE_SERV, DT_DOC)
                    Periodo = Util.ExtrairPeriodo(DT_REF)
                    CNPJEstabelecimento = ExtrairCampo(dicTitulosOrig, Campos, "CNPJ_ESTABELECIMENTO")
                    ARQUIVO = Periodo & "-" & CNPJContribuinte
                    If Not dtoRegSPED.r0000_Contr.Exists(ARQUIVO) Then Exit Function
                    arrCampos.Add ARQUIVO
                    
                Case Campo = "CHV_REG"
                    arrCampos.Add ""
                    
                Case Campo = "CHV_PAI_FISCAL"
                    arrCampos.Add ""
                    
                Case Campo = "CHV_PAI_CONTRIBUICOES"
                    CHV_PAI_CONTRIBUICOES = ExtrairChaveRegA010(ARQUIVO, CNPJEstabelecimento)
                    arrCampos.Add CHV_PAI_CONTRIBUICOES
                    
                Case Campo = "IND_OPER"
                    IND_OPER = Campos(dicTitulosOrig(Campo) - i)
                    IND_OPER = EnumContribuicoes.ValidarEnumeracao_IND_OPER(IND_OPER)
                    arrCampos.Add IND_OPER
                    
                Case Campo = "IND_EMIT"
                    IND_EMIT = Campos(dicTitulosOrig(Campo) - i)
                    IND_EMIT = EnumContribuicoes.ValidarEnumeracao_IND_EMIT(IND_EMIT)
                    arrCampos.Add IND_EMIT
                    
                Case Campo = "COD_SIT"
                    COD_SIT = Campos(dicTitulosOrig(Campo) - i)
                    COD_SIT = EnumContribuicoes.ValidarEnumeracao_COD_SIT(COD_SIT)
                    arrCampos.Add COD_SIT
                    
                Case Campo = "IND_PGTO"
                    IND_PGTO = Campos(dicTitulosOrig(Campo) - i)
                    IND_PGTO = EnumContribuicoes.ValidarEnumeracao_IND_PGTO(IND_PGTO)
                    arrCampos.Add IND_PGTO
                    
                Case Campo Like "VL_*"
                    arrCampos.Add fnExcel.ConverterValores(Campos(dicTitulosOrig(Campo) - i))
                    
                Case Campo Like "DT_*"
                    arrCampos.Add fnExcel.FormatarData(Campos(dicTitulosOrig(Campo) - i))
                    
                Case Else
                    If dicTitulosOrig.Exists(CStr(Campo)) Then _
                        arrCampos.Add fnExcel.FormatarTexto(Campos(dicTitulosOrig(Campo) - i)) _
                            Else arrCampos.Add ""
                    
            End Select
            
        Next Campo
        
        CamposA100 = arrCampos.toArray()
        Chave = GerarChaveA100(dtoTitSPED.tA100, CamposA100)
        CHV_REG = fnSPED.CriarChaveRegistro(dicLayoutA100, CHV_PAI_CONTRIBUICOES, CamposA100)
        CamposA100(dtoTitSPED.tA100("CHV_REG") - 1) = CHV_REG
        
        If dtoRegSPED.rA100.Exists(Chave) Then Debug.Print Chave
        dtoRegSPED.rA100(Chave) = CamposA100
        Call fnSPED.AtribuirChaveNivelContribuicoes(dicLayoutA100, CHV_REG)
        
        CriarRegistroA100 = True
        Call arrCampos.Clear
        
        Call CarregarCamposA100(Chave)
        Call CriarRegistroA170(Campos, CHV_REG)
        
    End With
    
End Function

Private Sub CriarRegistroA170(ByVal Campos As Variant, ByVal CHV_PAI As String)

Dim i As Integer
Dim Campo As Variant, CamposA170
Dim CHV_REG As String, CHV_PAI_CONTRIBUICOES$, Chave$, CNPJEstabelecimento$, NAT_BC_CRED$, IND_ORIG_CRED$, CST_PIS$, CST_COFINS$, DT_DOC$, DT_EXE_SERV$, DT_REF$
    
    Call arrCampos.Clear
    i = Util.VerificarPosicaoInicialArray(Campos)
    
    With CamposA170
        
        For Each Campo In dtoTitSPED.tA170.Keys()
            
            Select Case True
                
                Case Campo = "REG"
                    arrCampos.Add "A170"
                    
                Case Campo = "ARQUIVO"
                    DT_DOC = ExtrairCampo(dicTitulosOrig, Campos, "DT_DOC")
                    DT_EXE_SERV = ExtrairCampo(dicTitulosOrig, Campos, "DT_EXE_SERV")
                    DT_REF = IIf(DT_EXE_SERV <> "", DT_EXE_SERV, DT_DOC)
                    Periodo = Util.ExtrairPeriodo(DT_REF)
                    CNPJEstabelecimento = ExtrairCampo(dicTitulosOrig, Campos, "CNPJ_ESTABELECIMENTO")
                    ARQUIVO = Periodo & "-" & CNPJContribuinte
                    If Not dtoRegSPED.r0000_Contr.Exists(ARQUIVO) Then Exit Sub
                    arrCampos.Add ARQUIVO
                    
                Case Campo = "CHV_REG"
                    arrCampos.Add ""
                    
                Case Campo = "CHV_PAI_FISCAL"
                    arrCampos.Add ""
                    
                Case Campo = "CHV_PAI_CONTRIBUICOES"
                    arrCampos.Add CHV_PAI
                    
                Case Campo Like "NAT_BC_CRED"
                    NAT_BC_CRED = Campos(dicTitulosOrig(Campo) - i)
                    NAT_BC_CRED = EnumContribuicoes.ValidarEnumeracao_NAT_BC_CRED(NAT_BC_CRED)
                    arrCampos.Add NAT_BC_CRED
                    
                Case Campo Like "IND_ORIG_CRED"
                    IND_ORIG_CRED = Campos(dicTitulosOrig(Campo) - i)
                    IND_ORIG_CRED = EnumContribuicoes.ValidarEnumeracao_IND_ORIG_CRED(IND_ORIG_CRED)
                    arrCampos.Add IND_ORIG_CRED
                    
                Case Campo Like "CST_PIS"
                    CST_PIS = Campos(dicTitulosOrig(Campo) - i)
                    CST_PIS = EnumContribuicoes.ValidarEnumeracao_CST_PIS_COFINS(CST_PIS)
                    arrCampos.Add CST_PIS
                    
                Case Campo Like "CST_COFINS"
                    CST_COFINS = Campos(dicTitulosOrig(Campo) - i)
                    CST_COFINS = EnumContribuicoes.ValidarEnumeracao_CST_PIS_COFINS(CST_COFINS)
                    arrCampos.Add CST_COFINS
                    
                Case Campo Like "VL_*"
                    arrCampos.Add fnExcel.ConverterValores(Campos(dicTitulosOrig(Campo) - i))
                    
                Case Campo Like "ALIQ_*"
                    arrCampos.Add fnExcel.FormatarPercentuais(Campos(dicTitulosOrig(Campo) - i))
                    
                Case Else
                    If dicTitulosOrig.Exists(CStr(Campo)) Then _
                        arrCampos.Add fnExcel.FormatarTexto(Campos(dicTitulosOrig(Campo) - i)) _
                            Else arrCampos.Add ""
                    
            End Select
            
        Next Campo
        
        CamposA170 = arrCampos.toArray()
        Chave = GerarChaveA170(dtoTitSPED.tA170, CamposA170)
        CHV_REG = fnSPED.CriarChaveRegistro(dicLayoutA170, CHV_PAI, CamposA170)
        CamposA170(dtoTitSPED.tA170("CHV_REG") - 1) = CHV_REG
        
        dtoRegSPED.rA170(Chave) = CamposA170
        Call fnSPED.AtribuirChaveNivelContribuicoes(dicLayoutA170, CHV_REG)
        
        Call arrCampos.Clear
        
    End With
    
End Sub

Private Function ExtrairCampo(ByRef dicTitulos As Dictionary, ByRef Campos As Variant, ByVal nCampo As String)
    
Dim i As Integer
    
    i = Util.VerificarPosicaoInicialArray(Campos)
    If dicTitulos.Exists(nCampo) Then ExtrairCampo = Campos(dicTitulos(nCampo) - i)
    
End Function

Private Function GerarChaveA100(ByRef dicTitulos As Dictionary, ByRef Campos As Variant) As String

Dim i As Integer
    
    i = Util.VerificarPosicaoInicialArray(Campos)
    
    With CamposA100
        
        .IND_OPER = Campos(dicTitulos("IND_OPER") - i)
        .IND_EMIT = Campos(dicTitulos("IND_EMIT") - i)
        .COD_PART = Campos(dicTitulos("COD_PART") - i)
        .SER = Campos(dicTitulos("SER") - i)
        .SUB = Campos(dicTitulos("SUB") - i)
        .NUM_DOC = Campos(dicTitulos("NUM_DOC") - i)
        .CHV_NFSE = Campos(dicTitulos("CHV_NFSE") - i)
        
        GerarChaveA100 = Util.UnirCampos(Util.ApenasNumeros(.IND_OPER), Util.ApenasNumeros(.IND_EMIT), Util.RemoverAspaSimples(.COD_PART), _
            Util.RemoverAspaSimples(.SER), Util.RemoverAspaSimples(.SUB), Util.ApenasNumeros(.NUM_DOC), Util.RemoverAspaSimples(.CHV_NFSE))
        
    End With
    
End Function

Private Function GerarChaveA170(ByRef dicTitulos As Dictionary, ByRef Campos As Variant) As String
    
Dim i As Integer
    
    i = Util.VerificarPosicaoInicialArray(Campos)
    
    With CamposA170
        
        .NUM_ITEM = Util.ApenasNumeros(Campos(dicTitulos("NUM_ITEM") - i))
        .CHV_PAI_CONTRIBUICOES = CamposA100.CHV_REG
        
        GerarChaveA170 = Util.UnirCampos(.CHV_PAI_CONTRIBUICOES, .NUM_ITEM)
        
    End With
    
End Function

Private Sub CarregarCamposA100(ByVal Chave As String)

Dim Campos As Variant
Dim i As Long
    
    Campos = dtoRegSPED.rA100(Chave)
    If VBA.IsEmpty(Campos) Then Exit Sub
    
    i = Util.VerificarPosicaoInicialArray(Campos)
    
    With CamposA100
        
        .REG = Campos(dtoTitSPED.tA100("REG") - i)
        .ARQUIVO = Campos(dtoTitSPED.tA100("ARQUIVO") - i)
        .CHV_REG = Campos(dtoTitSPED.tA100("CHV_REG") - i)
        .CHV_PAI_FISCAL = Campos(dtoTitSPED.tA100("CHV_PAI_FISCAL") - i)
        .CHV_PAI_CONTRIBUICOES = Campos(dtoTitSPED.tA100("CHV_PAI_CONTRIBUICOES") - i)
        .IND_OPER = Campos(dtoTitSPED.tA100("IND_OPER") - i)
        .IND_EMIT = Campos(dtoTitSPED.tA100("IND_EMIT") - i)
        .COD_PART = Campos(dtoTitSPED.tA100("COD_PART") - i)
        .COD_SIT = Campos(dtoTitSPED.tA100("COD_SIT") - i)
        .SER = Campos(dtoTitSPED.tA100("SER") - i)
        .SUB = Campos(dtoTitSPED.tA100("SUB") - i)
        .NUM_DOC = Campos(dtoTitSPED.tA100("NUM_DOC") - i)
        .CHV_NFSE = Campos(dtoTitSPED.tA100("CHV_NFSE") - i)
        .DT_DOC = Campos(dtoTitSPED.tA100("DT_DOC") - i)
        .DT_EXE_SERV = Campos(dtoTitSPED.tA100("DT_EXE_SERV") - i)
        .VL_DOC = Campos(dtoTitSPED.tA100("VL_DOC") - i)
        .IND_PGTO = Campos(dtoTitSPED.tA100("IND_PGTO") - i)
        .VL_DESC = Campos(dtoTitSPED.tA100("VL_DESC") - i)
        .VL_BC_PIS = Campos(dtoTitSPED.tA100("VL_DESC") - i)
        .VL_PIS = Campos(dtoTitSPED.tA100("VL_PIS") - i)
        .VL_BC_COFINS = Campos(dtoTitSPED.tA100("VL_BC_COFINS") - i)
        .VL_COFINS = Campos(dtoTitSPED.tA100("VL_COFINS") - i)
        .VL_PIS_RET = Campos(dtoTitSPED.tA100("VL_PIS_RET") - i)
        .VL_COFINS_RET = Campos(dtoTitSPED.tA100("VL_COFINS_RET") - i)
        .VL_ISS = Campos(dtoTitSPED.tA100("VL_ISS") - i)
        
    End With
    
End Sub

Private Sub InicializarObjetos()
    
Dim COD_VER As String
    
    Call Util.DesabilitarControles
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    Call CarregarRegistrosSPED
    
    Set EnumContribuicoes = New clsEnumeracoesSPEDContribuicoes
    Set EnumFiscal = New clsEnumeracoesSPEDFiscal
    Set GerenciadorSPED = New clsRegistrosSPED
    Set importPlan = New ImportadorPlanilhas
    Set arrDadosRegistro = New ArrayList
    Set ExpReg = New ExportadorRegistros
    Set arrCampos = New ArrayList
    Set dicLayoutA100 = New Dictionary
    Set dicLayoutA170 = New Dictionary
    
    COD_VER = EnumContribuicoes.ValidarEnumeracao_COD_VER(Periodo)
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    Call CarregarRegistrosSPED
    
    Call FuncoesJson.CarregarLayoutSPEDContribuicoes(COD_VER)
    Set dicLayoutA100 = dicLayoutContribuicoes("A100")
    Set dicLayoutA170 = dicLayoutContribuicoes("A170")
    
    Call AtribuirChavesNivel_BlocoA
    
End Sub

Private Sub AtribuirChavesNivel_BlocoA()

Dim dicLayoutA001 As New Dictionary
Dim dicLayoutA010 As New Dictionary
Dim CHV_REG As String
    
    Set dicLayoutA001 = dicLayoutContribuicoes("A001")
    Set dicLayoutA010 = dicLayoutContribuicoes("A010")
    
    CHV_REG = ""
    Call fnSPED.AtribuirChaveNivelContribuicoes(dicLayoutA001, CHV_REG)
    
End Sub

Private Sub CriarRegistroA001(ByVal ARQUIVO As String)

Dim tpCont As String, Chave$
Dim Campos As Variant
    
    If dtoRegSPED.rA001 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroA001("ARQUIVO")
    
    With CamposA001
        
        If Not dtoRegSPED.rA001.Exists(ARQUIVO) Then
            
            .REG = "'A001"
            .ARQUIVO = ARQUIVO
            .IND_MOV = EnumFiscal.ValidarEnumeracao_IND_MOV("0")
            .CHV_PAI_FISCAL = ""
            .CHV_PAI_CONTRIBUICOES = dtoRegSPED.r0000_Contr(.ARQUIVO)(dtoTitSPED.t0000_Contr("CHV_REG"))
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI_CONTRIBUICOES, "A001")
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, .IND_MOV)
            
            dtoRegSPED.rA001(.ARQUIVO) = Campos
            
        End If
        
    End With
    
End Sub

Private Sub CriarRegistroA010(ByVal ARQUIVO As String, ByVal CNPJ As String)

Dim Chave As String
Dim Campos As Variant
    
    If dtoRegSPED.rA010 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroA010("ARQUIVO", "CNPJ")
    
    With CamposA010
        
        Chave = Util.UnirCampos(ARQUIVO, CNPJ)
        If Not dtoRegSPED.rA010.Exists(Chave) Then
            
            .REG = "A010"
            .ARQUIVO = ARQUIVO
            .CNPJ = CNPJ
            .CHV_PAI = ExtrairChaveRegA001(.ARQUIVO)
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CNPJ)
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, "", .CHV_PAI, fnExcel.FormatarTexto(.CNPJ))
            
            dtoRegSPED.rA010(Chave) = Campos
            
        End If
        
    End With
    
End Sub

Private Function ExtrairChaveRegA001(ByVal ARQUIVO As String)

Dim i As Integer
Dim Campos As Variant
    
    With dtoRegSPED
        
        If .rA001 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroA001("ARQUIVO")
        
        If Not .rA001.Exists(ARQUIVO) Then Call CriarRegistroA001(ARQUIVO)
        
        Campos = .rA001(ARQUIVO)
        i = Util.VerificarPosicaoInicialArray(Campos)
        
        ExtrairChaveRegA001 = Campos(dtoTitSPED.tA001("CHV_REG") - i)
        
    End With
        
End Function

Private Function ExtrairChaveRegA010(ByVal ARQUIVO As String, ByVal CNPJ As String) As String

Dim i As Byte
Dim Campos As Variant
Dim Chave As String, CHV_PAI$, CHV_REG$
    
    With dtoRegSPED
        
        If .rA010 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroA010("ARQUIVO", "CNPJ")
        
        Chave = Util.UnirCampos(ARQUIVO, CNPJ)
        If Not .rA010.Exists(Chave) Then Call CriarRegistroA010(ARQUIVO, CNPJ)
        
        Campos = .rA010(Chave)
        i = Util.VerificarPosicaoInicialArray(Campos)
        
        ExtrairChaveRegA010 = Campos(dtoTitSPED.tA010("CHV_REG") - i)
        
    End With
    
End Function

Private Sub CarregarRegistrosSPED()
    
    With GerenciadorSPED
        
        Call .CarregarDadosRegistro0000_Contr("ARQUIVO")
        Call .CarregarDadosRegistroA001("ARQUIVO")
        Call .CarregarDadosRegistroA010("ARQUIVO", "CNPJ")
        Call .CarregarDadosRegistroA100("IND_OPER", "IND_EMIT", "SER", "SUB", "NUM_DOC", "CHV_NFSE")
        Call .CarregarDadosRegistroA170("CHV_PAI_CONTRIBUICOES", "NUM_ITEM")
        
    End With
    
End Sub

Private Function ChecarAusenciaDados(ByRef arrDados As ArrayList) As Boolean

Dim Msg As String
Dim Titulo As String
    
    If arrDados Is Nothing Then
        
        ChecarAusenciaDados = True
        
        Titulo = "Importação de Planilha A100/A170"
        Msg = "A planilha selecionada não possui dados"
        
        Call Util.MsgAlerta(Msg, Titulo)
        
    End If
    
End Function

Private Sub LimparObjetos()
    
    Set EnumContribuicoes = Nothing
    Set arrDadosRegistro = Nothing
    Set dicTitulosOrig = Nothing
    Set PastaTrabalho = Nothing
    Set importPlan = Nothing
    Set EnumFiscal = Nothing
    Set arrCampos = Nothing
    Set PlanOrig = Nothing
    Set ExpReg = Nothing
    
    Call dicLayoutContribuicoes.RemoveAll
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    
    Call Util.AtualizarBarraStatus(False)
    Call Util.HabilitarControles
    
End Sub

Private Sub FecharPastaTrabalho()
    
    PastaTrabalho.Close
    Set PlanOrig = Nothing
    Set PastaTrabalho = Nothing
    
End Sub

Private Function ResetarCamposA100()
    
    Dim CamposVazios As CamposRegA100
    LSet CamposA100 = CamposVazios
    
End Function

Private Function ResetarCamposA170()
    
    Dim CamposVazios As CamposRegA170
    LSet CamposA170 = CamposVazios
    
End Function
