Attribute VB_Name = "clsC170_ImportadorDadosFiscal"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private dicTitulosC100 As New Dictionary
Private dicTitulosC170 As New Dictionary
Private dicDadosC100 As New Dictionary
Private dicOperacoes As New Dictionary
Private arrDadosC170 As New ArrayList
Private arrDados As New ArrayList

Public Sub ImportarDadosICMS_SPED_Fiscal(ByVal TipoImportacao As String)
    
    On Error GoTo Notificar:
    
    If Util.ChecarAusenciaDados(regC170, False) Then Exit Sub
    
    Call ProcessarDocumentos.CarregarSPEDs(TipoImportacao)
    If Not ProcessarDocumentos.PossuiSPEDFiscalListado Then Exit Sub
    
    Call CarregarDadosRegistros
    Call ProcessarSPEDsFiscais
    Call IncluirOperacoesNoC170
    
    Call Util.LimparDados(regC170, 4, False)
    Call Util.ExportarDadosArrayList(regC170, arrDados)
    
    Call LimparObjetos
    Call Util.MsgInformativa("Os dados de ICMS do SPED Fiscal foram incluídos com sucesso.", "Inclusão de valores do ICMS do SPED Fiscal", Inicio)
    
Exit Sub
Notificar:
    
    Call TratarExcecoes(TypeName(Me), "ImportarDadosICMS_SPED_Fiscal")
    
End Sub

Private Sub CarregarDadosRegistros()
    
    Call Util.AtualizarBarraStatus("Coletando dados do registro C170, por favor aguarde...")
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    Set arrDadosC170 = Util.CriarArrayListRegistro(regC170)
    
    Call Util.AtualizarBarraStatus("Coletando dados do registro C100, por favor aguarde...")
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    
End Sub

Private Sub ProcessarSPEDsFiscais()

Dim SPED As Variant
    
    Call dicOperacoes.RemoveAll
    
    For Each SPED In DocsFiscais.arrSPEDFiscal
        
        Call ExtrairDadosSPED(SPED)
        
    Next SPED
    
End Sub

Private Sub ExtrairDadosSPED(ByVal SPED As String)

Dim Registros As Variant, Registro, Campos
Dim nReg As String
    
    Registros = fnSPED.ExtrairRegistrosSPED(SPED)
    For Each Registro In Registros
        
        If Registro <> "" Then
            
            nReg = VBA.Mid(Registro, 2, 4)
            Select Case True
                
                Case nReg = "C100"
                    Call ExtrairDadosC100(Registro)
                    
                Case nReg = "C170"
                    If infoICMSC170.CHV_NFE <> "" Then
                    
                        Call ExtrairDadosC170(Registro)
                        Call RegistrarOperacao
                        
                    End If
                    
                Case nReg > "C197"
                    Exit Sub
                    
            End Select
            
        End If
        
    Next Registro
    
End Sub

Private Sub ExtrairDadosC100(ByVal Registro As String)
    
    Call DTO_DadosICMSC170SPEDFiscal.ResetarDadosICMSC170
    
    Campos = fnSPED.ExtrairCamposRegistro(Registro)
    If Campos(5) = "55" Then infoICMSC170.CHV_NFE = Campos(9)
    
End Sub

Private Sub ExtrairDadosC170(ByVal Registro As String)
    
    Campos = fnSPED.ExtrairCamposRegistro(Registro)
    
    With infoICMSC170
        
        If .CHV_NFE <> "" Then
            
            .NUM_ITEM = CInt(fnExcel.ConverterValores(Campos(2), True, 0))
            .COD_ITEM = Campos(3)
            .CFOP = CInt(fnExcel.ConverterValores(Campos(11), True, 0))
            .CST_ICMS = fnExcel.FormatarTexto(Campos(10))
            .VL_BC_ICMS = fnExcel.ConverterValores(Campos(13), True, 2)
            .ALIQ_ICMS = fnExcel.FormatarPercentuais(Campos(14)) / 100
            .VL_ICMS = fnExcel.ConverterValores(Campos(15), True, 2)
            
        End If
        
    End With
    
End Sub

Private Function RegistrarOperacao()

Dim Chave As String
    
    With infoICMSC170
        
        Chave = .CHV_NFE & "|" & .NUM_ITEM & "|" & .COD_ITEM
        dicOperacoes(Chave) = DTO_DadosICMSC170SPEDFiscal.MontarArrayInfoICMSC170()
        
    End With
    
End Function

Private Function IncluirOperacoesNoC170()

Dim Campos As Variant
Dim Chave As String
    
    For Each Campos In arrDadosC170
        
        Chave = ObterChaveOperacao(Campos)
        If dicOperacoes.Exists(Chave) Then Call IncluirCamposICMSC170(Campos, Chave)
        
        arrDados.Add Campos
        
    Next Campos
    
End Function

Private Function ObterChaveOperacao(ByRef Campos As Variant) As String

Dim CHV_PAI As String
    
    CHV_PAI = Campos(dicTitulosC170("CHV_PAI_FISCAL"))
    If Not dicDadosC100.Exists(CHV_PAI) Then Exit Function
    
    Call DTO_DadosICMSC170SPEDFiscal.ResetarDadosICMSC170
    
    With infoICMSC170
        
        .CHV_NFE = dicDadosC100(CHV_PAI)(dicTitulosC100("CHV_NFE"))
        .NUM_ITEM = CInt(fnExcel.ConverterValores(Campos(dicTitulosC170("NUM_ITEM")), True, 0))
        .COD_ITEM = Campos(dicTitulosC170("COD_ITEM"))
        
        ObterChaveOperacao = .CHV_NFE & "|" & .NUM_ITEM & "|" & .COD_ITEM
        
    End With
    
End Function

Private Function IncluirCamposICMSC170(ByRef Campos As Variant, ByVal Chave As String)
    
    Call ExtrairDadosICMSOperacao(Chave)
    
    With infoICMSC170
        
        Campos(dicTitulosC170("CFOP")) = .CFOP
        Campos(dicTitulosC170("CST_ICMS")) = .CST_ICMS
        Campos(dicTitulosC170("VL_BC_ICMS")) = .VL_BC_ICMS
        Campos(dicTitulosC170("ALIQ_ICMS")) = .ALIQ_ICMS
        Campos(dicTitulosC170("VL_ICMS")) = .VL_ICMS
        
    End With
    
End Function

Private Sub ExtrairDadosICMSOperacao(ByVal Chave As String)
    
    With infoICMSC170
        
        .CFOP = dicOperacoes(Chave)(0)
        .CST_ICMS = dicOperacoes(Chave)(1)
        .VL_BC_ICMS = dicOperacoes(Chave)(2)
        .ALIQ_ICMS = dicOperacoes(Chave)(3)
        .VL_ICMS = dicOperacoes(Chave)(4)
        
    End With
    
End Sub

Private Function TratarExcecoes(ByVal NomeClasse As String, ByVal NomeMetodo As String)
    
    Select Case True
        
        Case Else
            Dim infoErro As New clsGerenciadorErros
            Call infoErro.NotificarErroInesperado(NomeClasse, NomeMetodo)
            
    End Select
    
End Function

Private Sub LimparObjetos()
    
    Set dicTitulosC170 = Nothing
    Set arrDadosC170 = Nothing
    
    Set dicTitulosC100 = Nothing
    Set dicDadosC100 = Nothing
    
    Set arrDados = Nothing
    Set dicOperacoes = Nothing
    
    Call Util.AtualizarBarraStatus(False)
    
End Sub
