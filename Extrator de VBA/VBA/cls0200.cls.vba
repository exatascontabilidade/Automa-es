Attribute VB_Name = "cls0200"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Sub AtualizarCodigoGenero(Optional OmitirMsg As Boolean)

Dim Campos As Variant, Campo, nCampo, Titulos
Dim dicTitulos As New Dictionary
Dim arrDados As New ArrayList
Dim arr0200 As New ArrayList
Dim Valores As Variant
Dim Chave As String
Dim Inicio As Date
    
    If Util.ChecarAusenciaDados(reg0200, OmitirMsg) Then Exit Sub
    If Not OmitirMsg Then Inicio = Now()
    
    Call Util.AtualizarBarraStatus("Atualizando os códigos de gênero do registro 0200, por favor aguarde...")
    
    If reg0200.AutoFilterMode Then reg0200.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(reg0200, 3)
    
    Set arrDados = Util.CriarArrayListRegistro(reg0200)
    If arrDados.Count = 0 Then Exit Sub
    
    a = 0
    Comeco = Timer
    For Each Campos In arrDados
        
        Call Util.AntiTravamento(a, 50, "Atualizando código do gênero no registro 0200, por favor aguarde...", arrDados.Count, Comeco)
        Campos(dicTitulos("COD_GEN")) = VBA.Left(Util.ApenasNumeros(Campos(dicTitulos("COD_NCM"))), 2)
        arr0200.Add Campos
        
    Next Campos
    
    Call Util.LimparDados(reg0200, 4, False)
    Call Util.ExportarDadosArrayList(reg0200, arr0200)
    
    Call arr0200.Clear
    
    Application.StatusBar = False
    If Not OmitirMsg Then Call Util.MsgInformativa("Códigos de gênero atualizados com sucesso!", _
        "Atualização dos códigos de gênero do registro 0200", Inicio)
    
End Sub

Public Sub ImportarCadastroProdutos()

Dim Arqs As Variant, Arq, Chave
Dim arrDados As New ArrayList
Dim dicDados As New Dictionary
Dim versao As String
Dim i As Byte
    
    Arqs = Util.SelecionarArquivos("txt", "Selecione o SPED com o Cadastro de Produtos")
    If VarType(Arqs) = vbBoolean Then Exit Sub
    
    For Each Arq In Arqs
        
        If fnSPED.ValidarSPEDFiscal(Arq) Then Call fnSPED.ExtrairCadastroProdutos(Arq)
        
    Next Arq
    
    For Each Chave In dicRegistros.Keys()
        
        Select Case Chave
            
            Case "0200"
                Set arrDados = dicRegistros(Chave)
                Call Util.ExportarDadosArrayList(Worksheets(Chave), arrDados)
                
        End Select
        
    Next Chave
    
End Sub
