Attribute VB_Name = "clsC500"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Sub gerarC500Filhos_Contribuicoes()

    

End Sub

'Private Sub InicializarObjetos()
'
'Dim COD_VER As String
'
'    Call Util.DesabilitarControles
'
'    Call DTO_RegistrosSPED.ResetarRegistrosSPED
'    Call CarregarRegistrosSPED
'
'    Set EnumContribuicoes = New clsEnumeracoesSPEDContribuicoes
'    Set EnumFiscal = New clsEnumeracoesSPEDFiscal
'    Set GerenciadorSPED = New clsRegistrosSPED
'    Set importPlan = New ImportadorPlanilhas
'    Set arrDadosRegistro = New ArrayList
'    Set ExpReg = New ExportadorRegistros
'    Set arrCampos = New ArrayList
'    Set dicLayoutA100 = New Dictionary
'    Set dicLayoutA170 = New Dictionary
'    Set PlanDest = ActiveSheet
'
'    COD_VER = EnumFiscal.ValidarEnumeracao_COD_VER(Periodo)
'
'    Call DTO_RegistrosSPED.ResetarRegistrosSPED
'    Call CarregarRegistrosSPED
'
'    Call FuncoesJson.CarregarLayoutSPEDContribuicoes(COD_VER)
'    Set dicLayoutA100 = dicLayoutContribuicoes("A100")
'    Set dicLayoutA170 = dicLayoutContribuicoes("A170")
'
'    Call AtribuirChavesNivel_BlocoA
'
'End Sub
'
'Private Sub LimparObjetos()
'
'    Set EnumContribuicoes = Nothing
'    Set arrDadosRegistro = Nothing
'    Set dicTitulosOrig = Nothing
'    Set PastaTrabalho = Nothing
'    Set importPlan = Nothing
'    Set EnumFiscal = Nothing
'    Set arrCampos = Nothing
'    Set PlanOrig = Nothing
'    Set PlanDest = Nothing
'    Set ExpReg = Nothing
'
'    Call dicLayoutContribuicoes.RemoveAll
'    Call DTO_RegistrosSPED.ResetarRegistrosSPED
'
'    Call Util.AtualizarBarraStatus(False)
'    Call Util.HabilitarControles
'
'End Sub
'
'Public Function ResetarCamposA100()
'
'    Dim CamposVazios As CamposRegA100
'    LSet CamposA100 = CamposVazios
'
'End Function
'
'Public Function ResetarCamposA170()
'
'    Dim CamposVazios As CamposRegA170
'    LSet CamposA170 = CamposVazios
'
'End Function
