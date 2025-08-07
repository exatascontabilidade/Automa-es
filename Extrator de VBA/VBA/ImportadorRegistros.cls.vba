Attribute VB_Name = "ImportadorRegistros"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Registro As String, ARQUIVO$, Periodo$, CHV_REG$, CHV_PAI_FISCAL$, CHV_PAI_CONTRIBUICOES$, COD_VER_FISCAL$, COD_VER_CONTRIBUICOES$
Private EnumContribuicoes As clsEnumeracoesSPEDContribuicoes
Private EnumFiscal As clsEnumeracoesSPEDFiscal
Private GerenciadorSPED As clsRegistrosSPED
Private importPlan As ImportadorPlanilhas
Private ExpReg As ExportadorRegistros
Private arrDadosRegistro As ArrayList
Private dicTitulosOrig As Dictionary
Private dicTitulosDest As Dictionary
Private RegistroInvalido As Boolean
Private PastaTrabalho As Workbook
Private arrDadosOrig As ArrayList
Private arrCampos As ArrayList
Private PlanOrig As Worksheet
Private PlanDest As Worksheet

Public Sub ImportarDadosRegistro()
    
    If Not Funcoes.CarregarDadosContribuinte Then Exit Sub
    If Not ValidarPeriodoImportacao Then Exit Sub
    
    Call InicializarObjetos
        
        Call SetarPlanilhaOrigem
        Call ProcessarRegistrosPlanilha
        
        Call ExpReg.ExportarRegistros(Registro)
        
    Call LimparObjetos
    
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

Private Function SetarPlanilhaOrigem()
    
    Set PastaTrabalho = importPlan.AbrirPastaTrabalho
    Set PlanOrig = PastaTrabalho.Worksheets(1)
    
End Function

Private Sub ProcessarRegistrosPlanilha()

Dim Campos As Variant
    
    Set dicTitulosDest = Util.MapearTitulos(PlanDest, 3)
    Set dicTitulosOrig = Util.MapearTitulos(PlanOrig, 1)
    
    Set arrDadosOrig = Util.CriarArrayListRegistro(PlanOrig, 2, 1)
    
    For Each Campos In arrDadosOrig
        
        Call ProcessarCamposRegistro(Campos)
        
    Next Campos
    
End Sub

Private Sub ProcessarCamposRegistro(Campos)

Dim i As Integer
Dim Campo As Variant
    
    Call arrCampos.Clear
    i = Util.VerificarPosicaoInicialArray(Campos)
    
    For Each Campo In dicTitulosDest.Keys()
        
        Select Case Campo
            
            Case "ARQUIVO"
                arrCampos.Add ARQUIVO
                
            Case Else
                If dicTitulosOrig.Exists(CStr(Campo)) Then _
                    arrCampos.Add Campos(dicTitulosOrig(Campo) - i) _
                        Else arrCampos.Add ""
                
        End Select
        
    Next Campo
    
    Call CalcularChavesRegistro
    
    arrDadosRegistro.Add arrCampos.toArray()
    Call arrCampos.Clear
    
End Sub

Private Sub CalcularChavesRegistro()
    
Dim REG As String
    
    REG = arrCampos(dicTitulosDest("REG"))
    
    CHV_REG = ""
    CHV_PAI_FISCAL = ""
    CHV_PAI_CONTRIBUICOES = ""
    
    arrCampos(dicTitulosDest("CHV_REG")) = CHV_REG
    arrCampos(dicTitulosDest("CHV_PAI_FISCAL")) = CHV_PAI_FISCAL
    arrCampos(dicTitulosDest("CHV_PAI_CONTRIBUICOES")) = CHV_PAI_CONTRIBUICOES
    
End Sub

Private Sub CalcularChavesRegistro_Blo()
    
Dim REG As String
    
    REG = arrCampos(dicTitulosDest("REG"))
    
    CHV_REG = ""
    CHV_PAI_FISCAL = ""
    CHV_PAI_CONTRIBUICOES = ""
    
    arrCampos(dicTitulosDest("CHV_REG")) = CHV_REG
    arrCampos(dicTitulosDest("CHV_PAI_FISCAL")) = CHV_PAI_FISCAL
    arrCampos(dicTitulosDest("CHV_PAI_CONTRIBUICOES")) = CHV_PAI_CONTRIBUICOES
    
End Sub

Private Sub InicializarObjetos()
    
    Set EnumContribuicoes = New clsEnumeracoesSPEDContribuicoes
    Set EnumFiscal = New clsEnumeracoesSPEDFiscal
    Set GerenciadorSPED = New clsRegistrosSPED
    Set importPlan = New ImportadorPlanilhas
    Set arrDadosRegistro = New ArrayList
    Set ExpReg = New ExportadorRegistros
    Set arrCampos = New ArrayList
    Set PlanDest = ActiveSheet
    
    RegistroInvalido = False
    Registro = VBA.Replace(PlanDest.CodeName, "reg", "")
    COD_VER_FISCAL = EnumFiscal.ValidarEnumeracao_COD_VER(Periodo)
    COD_VER_CONTRIBUICOES = EnumContribuicoes.ValidarEnumeracao_COD_VER(Periodo)
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    Call CarregarRegistrosSPED(Registro)
    
    Call FuncoesJson.CarregarLayoutSPEDFiscal(COD_VER_FISCAL)
    Call FuncoesJson.CarregarLayoutSPEDContribuicoes(COD_VER_CONTRIBUICOES)
    
End Sub

Private Sub CarregarRegistrosSPED(ByVal Registro As String)
    
    With GerenciadorSPED
        
        .CarregarDadosRegistro0000 ("ARQUIVO")
        .CarregarDadosRegistro0000_Contr ("ARQUIVO")
        
    End With
    
    With dtoRegSPED
        
        If .r0000.Count > 0 Then Call CarregarRegistrosSPED_Fiscal(Registro)
        If .r0000_Contr.Count > 0 Then Call CarregarRegistrosSPED_Contribuicoes(Registro)
        
    End With
    
End Sub

Private Sub CarregarRegistrosSPED_Fiscal(ByVal Registro As String)

Dim Msg As String, Titulo$
    
    Select Case VBA.Left(Registro, 1)
        
        Case "0"
            Call CarregarRegistros_Bloco0
            
        Case "K"
            Call CarregarRegistros_BlocoK
            
        Case Else
            Msg = "Registro não mapeado para importação via planilha" & vbCrLf
            Msg = Msg & "Contate o suporte."
            Titulo = "Ação Não Permitida"
            Call Util.MsgAlerta(Msg, Titulo)
            
    End Select
    
End Sub

Private Sub CarregarRegistrosSPED_Contribuicoes(ByVal Registro As String)

Dim Msg As String, Titulo$
    
    Select Case VBA.Left(Registro, 1)
        
        Case "0"
            Call CarregarRegistros_Bloco0(True)
            
        Case "A"
            Call CarregarRegistros_BlocoA
            
        Case Else
            Msg = "Registro não mapeado para importação via planilha" & vbCrLf
            Msg = Msg & "Contate o suporte."
            Titulo = "Ação Não Permitida"
            Call Util.MsgAlerta(Msg, Titulo)
            
    End Select
    
End Sub

Private Sub CarregarRegistros_Bloco0(Optional ByVal Contr As Boolean = False)
    
    With GerenciadorSPED
        
        Call .CarregarDadosRegistro0001("ARQUIVO")
        If Contr Then Call .CarregarDadosRegistro0140("ARQUIVO", "CNPJ")
        
    End With
    
End Sub

Private Sub CarregarRegistros_BlocoA()
    
    With GerenciadorSPED
        
        Call .CarregarDadosRegistroA001("ARQUIVO")
        Call .CarregarDadosRegistroA010("ARQUIVO", "CNPJ")
        Call .CarregarDadosRegistroA100("IND_OPER", "IND_EMIT", "SER", "NUM_DOC", "CHV_NFSE")
        Call .CarregarDadosRegistroA170
        
    End With
    
End Sub

Private Sub CarregarRegistros_BlocoK()
    
    With GerenciadorSPED
        
        Call .CarregarDadosRegistroK001("ARQUIVO")
        Call .CarregarDadosRegistroK010("ARQUIVO")
        Call .CarregarDadosRegistroK100("ARQUIVO")
        
    End With
    
End Sub

Private Sub LimparObjetos()
    
    Set EnumContribuicoes = Nothing
    Set arrDadosRegistro = Nothing
    Set dicTitulosOrig = Nothing
    Set dicTitulosDest = Nothing
    Set PastaTrabalho = Nothing
    Set importPlan = Nothing
    Set EnumFiscal = Nothing
    Set arrCampos = Nothing
    Set PlanOrig = Nothing
    Set PlanDest = Nothing
    Set ExpReg = Nothing
    
    Registro = ""
    RegistroInvalido = False
    
    Call dicLayoutFiscal.RemoveAll
    Call dicLayoutContribuicoes.RemoveAll
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    
    Call Util.AtualizarBarraStatus(False)
    Call Util.HabilitarControles
    
End Sub
