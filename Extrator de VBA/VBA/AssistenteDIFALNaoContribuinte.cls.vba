Attribute VB_Name = "AssistenteDIFALNaoContribuinte"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private AssistTributario As New AssistenteTributario
Private dicCorrelacoesCTeNFe As New Dictionary
Private dicTitulosTributacao As New Dictionary
Private dicDadosTributarios As New Dictionary
Private dicTitulosApuracao As New Dictionary
Private dicTitulosCTeNFe As New Dictionary
Private dicTitulos0000 As New Dictionary
Private dicTitulosC101 As New Dictionary
Private dicTitulosD100 As New Dictionary
Private dicTitulosD101 As New Dictionary
Private dicTitulosE001 As New Dictionary
Private dicTitulosE300 As New Dictionary
Private dicTitulosE310 As New Dictionary
Private dicTitulosE311 As New Dictionary
Private dicTitulosE313 As New Dictionary
Private dicTitulosE316 As New Dictionary
Private dicDados0000 As New Dictionary
Private dicDadosC101 As New Dictionary
Private dicDadosD100 As New Dictionary
Private dicDadosD101 As New Dictionary
Private dicDadosE001 As New Dictionary
Private dicDadosE300 As New Dictionary
Private dicDadosE310 As New Dictionary
Private dicDadosE311 As New Dictionary
Private dicDadosE313 As New Dictionary
Private dicDadosE316 As New Dictionary
Private DadosE313Carregado As Boolean
Private CHV_0000 As String
Private CHV_E001 As String
Private CHV_E300 As String
Private CHV_E310 As String
Private CHV_E311 As String

Public Sub CalcularDifalNaoContribuinte()

Dim Campos As Variant, DadosC101
Dim arrDadosApuracao As New ArrayList
Dim CFOP As String, REG$, CHV_NFE$, UF$
    
    If Util.ChecarAusenciaDados(assApuracaoICMS, False) Then Exit Sub
    Inicio = Now()
    
    Call CarregarDados
    
    Set arrDadosApuracao = Util.CriarArrayListRegistro(assApuracaoICMS)
    For Each Campos In arrDadosApuracao
        
        Call Util.AntiTravamento(a, 10, "Calculando DIFAL a não contribuinte, por favor aguarde...", arrDadosApuracao.Count, Comeco)
        
        UF = Campos(dicTitulosApuracao("UF_PART"))
        If UF = "EX" Then GoTo Prx:
        
        Call ChecarExistenciaRegistros(Campos)
        
        REG = Campos(dicTitulosApuracao("REG"))
        CHV_NFE = Campos(dicTitulosApuracao("CHV_NFE"))
        CFOP = Campos(dicTitulosApuracao("CFOP"))
        
        Select Case True
            
            Case REG Like "C170" And CFOP Like "22*" And ChecarConsumidorFinal(Campos)
                Call AtualizarCorrelacaoCTeNFe(CHV_NFE, "ESTORNO_DIFAL", "SIM", UF)
                Call IncluirRegistroE311_E313(Campos)
                
            Case REG Like "C170" And CFOP Like "61*" And ChecarConsumidorFinal(Campos)
                Call AtualizarCorrelacaoCTeNFe(CHV_NFE, "DEBITO_DIFAL", "SIM", UF)
                Call IncluirDebitosDIFALC101_D101(Campos)
                
            Case Else
                Call AtualizarCorrelacaoCTeNFe(CHV_NFE, "ESTORNO_DIFAL", "NÃO", UF)
                Call AtualizarCorrelacaoCTeNFe(CHV_NFE, "DEBITO_DIFAL", "NÃO", UF)
                
        End Select
Prx:
    Next Campos
    
    Call CalcularSaldosE310
    
    Call ExportarRegistros
    Call Util.MsgInformativa("Os registros do DIFAL a não contribunte foram atualizados com sucesso!", "Atualização de registros C101", Inicio)
    
End Sub

Private Function ChecarConsumidorFinal(ByRef Campos As Variant) As Boolean

Dim TIPO_PART As String, CONTRIBUINTE$
            
    TIPO_PART = Campos(dicTitulosApuracao("TIPO_PART"))
    CONTRIBUINTE = Campos(dicTitulosApuracao("CONTRIBUINTE"))
        
    Select Case True
        
        Case TIPO_PART Like "PF"
            ChecarConsumidorFinal = True
            
        Case TIPO_PART Like "PJ" And CONTRIBUINTE Like "N*"
            ChecarConsumidorFinal = True
            
    End Select
    
End Function

Private Function IncluirDebitosDIFALC101_D101(ByRef Campos As Variant)

Dim DadosC101 As Variant
Dim CamposTributarios As Variant
Dim ChaveTrib As String, CFOP$, CHV_NFE$, DEBITO_DIFAL$, DT_REF$, CHV_CTE$, COD_AJUSTE_DEBITO_DIFAL$
Dim ALIQ_DEST As Double, ALIQ_FCP#, ALIQ_ICMS#, ALIQ_DIFAL#, VL_BC_ICMS#, VL_DIFAL_NFE#, VL_FCP_NFE#, VL_BC_DIFAL_CTE#, VL_DIFAL_CTE#, VL_FCP_CTE#
    
    ChaveTrib = AssistTributario.GerarChaveTributacao(assApuracaoICMS, Campos)
    VL_BC_ICMS = Campos(dicTitulosApuracao("VL_BC_ICMS"))
    CHV_NFE = Campos(dicTitulosApuracao("CHV_NFE"))
    DT_REF = AssistTributario.ExtrairDataReferencia(Campos)
    
    If dicDadosTributarios.Exists(ChaveTrib) And VL_BC_ICMS > 0 Then
        
        CamposTributarios = AssistTributario.ExtrairCamposTributarios(ChaveTrib, DT_REF)
        
        ALIQ_DEST = fnExcel.FormatarPercentuais(CamposTributarios(dicTitulosTributacao("ALIQ_ICMS_DEST")))
        If ALIQ_DEST > 0 Then
            
            ALIQ_FCP = fnExcel.ConverterValores(CamposTributarios(dicTitulosTributacao("ALIQ_FCP")))
            ALIQ_ICMS = fnExcel.ConverterValores(Campos(dicTitulosApuracao("ALIQ_ICMS")))
            ALIQ_DIFAL = ALIQ_DEST - ALIQ_ICMS
            
            VL_DIFAL_NFE = VBA.Round(VL_BC_ICMS * ALIQ_DIFAL, 2)
            VL_FCP_NFE = VBA.Round(VL_BC_ICMS * ALIQ_FCP, 2)
            
            Call IncluirRegistroC101(Campos, VL_DIFAL_NFE, VL_FCP_NFE)
            
            If dicCorrelacoesCTeNFe.Exists(CHV_NFE) Then
                
                DEBITO_DIFAL = dicCorrelacoesCTeNFe(CHV_NFE)(dicTitulosCTeNFe("DEBITO_DIFAL"))
                VL_BC_DIFAL_CTE = dicCorrelacoesCTeNFe(CHV_NFE)(dicTitulosCTeNFe("VL_BC_DIFAL_CTE"))
                CHV_CTE = dicCorrelacoesCTeNFe(CHV_NFE)(dicTitulosCTeNFe("CHV_CTE"))
                
                If VL_BC_DIFAL_CTE > 0 And DEBITO_DIFAL = "SIM" Then
                    
                    VL_DIFAL_CTE = VBA.Round(VL_BC_DIFAL_CTE * ALIQ_DIFAL, 2)
                    VL_FCP_CTE = 0 ' VBA.Round(VL_BC_DIFAL_CTE * ALIQ_FCP, 2)
                    
                    COD_AJUSTE_DEBITO_DIFAL = GerarCodigoAjusteDebitoDIFAL(Campos)
                    
                    If VL_DIFAL_CTE > 0 Then Call IncluirRegistroE311(Campos, VL_DIFAL_CTE, COD_AJUSTE_DEBITO_DIFAL, "DEBITO PARA AJUSTE DE APURAÇÃO DO DIFAL ICMS", CHV_CTE)
                    If VL_FCP_CTE > 0 Then Call IncluirRegistroE311(Campos, VL_FCP_CTE, "309999", "DEBITO PARA AJUSTE DE APURAÇÃO DO DIFAL FCP", CHV_CTE)
                    
                    If Not dicDadosE310.Exists(CHV_E310) Then Call CriarRegistroE310(Campos)
                    Call AtualizarRegistroE310(CHV_E310, 0, 0, 0, 0, VL_DIFAL_CTE, VL_FCP_CTE)
                    
                End If
            
            End If
            
'            If dicCorrelacoesCTeNFe.Exists(CHV_NFE) Then
'
'                DEBITO_DIFAL = dicCorrelacoesCTeNFe(CHV_NFE)(dicTitulosCTeNFe("DEBITO_DIFAL"))
'                VL_BC_DIFAL_CTE = dicCorrelacoesCTeNFe(CHV_NFE)(dicTitulosCTeNFe("VL_BC_DIFAL_CTE"))
'
'                If VL_BC_DIFAL_CTE > 0 And DEBITO_DIFAL = "SIM" Then
'
'                    VL_DIFAL_CTE = VBA.Round(VL_BC_DIFAL_CTE * ALIQ_DIFAL, 2)
'                    VL_FCP_CTE = 0 ' VBA.Round(VL_BC_DIFAL_CTE * ALIQ_FCP, 2)
'
'                    Call IncluirRegistroD101(Campos, VL_DIFAL_CTE, VL_FCP_CTE)
'
'                End If
'
'            End If
            
        End If
        
    End If
    
End Function

Private Function IncluirRegistroC101(ByRef Campos As Variant, ByVal VL_DIFAL As Double, ByVal VL_FCP As Double)

Dim ALIQ_DEST As Double, ALIQ_FCP#, ALIQ_ICMS#, ALIQ_DIFAL#, VL_BC_ICMS#
    
    With CamposC101
        
        .REG = "C101"
        .ARQUIVO = Campos(dicTitulosApuracao("ARQUIVO"))
        .CHV_PAI = Campos(dicTitulosApuracao("CHV_PAI_FISCAL"))
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "C101")
        .VL_FCP_UF_DEST = VL_FCP
        .VL_ICMS_UF_DEST = VL_DIFAL
        .VL_ICMS_UF_REM = 0
        
        Call AtualizarRegistroE310(CHV_E310, .VL_ICMS_UF_DEST, .VL_FCP_UF_DEST, 0, 0, 0, 0)
        If dicDadosC101.Exists(.CHV_REG) Then Call AtualizarRegistroC101(.CHV_REG)
        
        dicDadosC101(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", _
            CDbl(.VL_FCP_UF_DEST), CDbl(.VL_ICMS_UF_DEST), CDbl(.VL_ICMS_UF_REM))
        
    End With
    
End Function

Private Function AtualizarRegistroC101(ByVal Chave As String)
    
    With CamposC101
        
        .VL_FCP_UF_DEST = dicDadosC101(Chave)(dicTitulosC101("VL_FCP_UF_DEST") - 1) + CDbl(.VL_FCP_UF_DEST)
        .VL_ICMS_UF_DEST = dicDadosC101(Chave)(dicTitulosC101("VL_ICMS_UF_DEST") - 1) + CDbl(.VL_ICMS_UF_DEST)
        .VL_ICMS_UF_REM = dicDadosC101(Chave)(dicTitulosC101("VL_ICMS_UF_REM") - 1) + CDbl(.VL_ICMS_UF_REM)
         
    End With
    
End Function

Private Function IncluirRegistroD101(ByRef Campos As Variant, ByVal VL_DIFAL As Double, ByVal VL_FCP As Double)

Dim CHV_PAI As String, CHV_NFE$, CHV_CTE$, COD_MUN_ORIG$, COD_MUN_DEST$
    
    CHV_NFE = Campos(dicTitulosApuracao("CHV_NFE"))
    If dicCorrelacoesCTeNFe.Exists(CHV_NFE) Then
        
        CHV_CTE = dicCorrelacoesCTeNFe(CHV_NFE)(dicTitulosCTeNFe("CHV_CTE"))
        If dicDadosD100.Exists(CHV_CTE) Then
            
            CHV_PAI = dicDadosD100(CHV_CTE)(dicTitulosD100("CHV_REG"))
            COD_MUN_ORIG = Util.ApenasNumeros(dicDadosD100(CHV_CTE)(dicTitulosD100("COD_MUN_ORIG")))
            COD_MUN_DEST = Util.ApenasNumeros(dicDadosD100(CHV_CTE)(dicTitulosD100("COD_MUN_DEST")))
            
            If VBA.Left(COD_MUN_ORIG, 2) <> VBA.Left(COD_MUN_DEST, 2) Then
                
                With CamposD101
                    
                    .REG = "D101"
                    .ARQUIVO = Campos(dicTitulosApuracao("ARQUIVO"))
                    .CHV_PAI = CHV_PAI
                    .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "D101")
                    .VL_FCP_UF_DEST = VL_FCP
                    .VL_ICMS_UF_DEST = VL_DIFAL
                    .VL_ICMS_UF_REM = 0
                    
                    Call AtualizarRegistroE310(CHV_E310, .VL_ICMS_UF_DEST, .VL_FCP_UF_DEST, 0, 0, 0, 0)
                    If dicDadosD101.Exists(.CHV_REG) Then Call AtualizarRegistroD101(.CHV_REG)
                    
                    dicDadosD101(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", _
                        CDbl(.VL_FCP_UF_DEST), CDbl(.VL_ICMS_UF_DEST), CDbl(.VL_ICMS_UF_REM))
                        
                End With
                
            End If
            
        End If
        
     End If
     
End Function

Private Function AtualizarRegistroD101(ByVal Chave As String)
    
    With CamposD101
        
        .VL_FCP_UF_DEST = dicDadosD101(Chave)(dicTitulosD101("VL_FCP_UF_DEST") - 1) + CDbl(.VL_FCP_UF_DEST)
        .VL_ICMS_UF_DEST = dicDadosD101(Chave)(dicTitulosD101("VL_ICMS_UF_DEST") - 1) + CDbl(.VL_ICMS_UF_DEST)
        .VL_ICMS_UF_REM = dicDadosD101(Chave)(dicTitulosD101("VL_ICMS_UF_REM") - 1) + CDbl(.VL_ICMS_UF_REM)
         
    End With
    
End Function

Private Function IncluirRegistroE311_E313(ByRef Campos As Variant)

Dim CHV_E311 As String, CHV_E313$, UF$
        
    Call IncluirEstornosDIFALE311_E313(Campos)
    
End Function

Private Sub ChecarExistenciaRegistros(ByRef Campos As Variant)

Dim ARQUIVO As String, Periodo$, CHV_PAI$, UF$, DT_INI$, DT_FIN$
    
    ARQUIVO = Campos(dicTitulosApuracao("ARQUIVO"))
    UF = Campos(dicTitulosApuracao("UF_PART"))
    Periodo = VBA.Split(ARQUIVO, "-")(0)
    
    'E001
    CHV_0000 = dicDados0000(ARQUIVO)(dicTitulos0000("CHV_REG"))
    CHV_E001 = fnSPED.GerarChaveRegistro(CHV_0000, "E001")
    If Not dicDadosE001.Exists(CHV_E001) Then Call CriarRegistroE001(Campos)
    
    'E300
    DT_FIN = Util.ConverterPeriodoData(Periodo)
    DT_INI = VBA.Format(DT_FIN, "yyyy-mm-01")
    CHV_E300 = fnSPED.GerarChaveRegistro(CHV_E001, UF, CDate(DT_INI), CDate(DT_FIN))
    If Not dicDadosE300.Exists(CHV_E300) Then Call CriarRegistroE300(Campos)
    
    'E310
    CHV_E310 = fnSPED.GerarChaveRegistro(CHV_E300, "E310")
    If Not dicDadosE310.Exists(CHV_E310) Then Call CriarRegistroE310(Campos)
    
End Sub

Private Sub CarregarDados()
    
    Set dicDadosTributarios = AssistTributario.CarregarTributacoesSalvas(assTributacaoICMS)
    Set dicCorrelacoesCTeNFe = Util.CriarDicionarioRegistro(CorrelacoesCTeNFe, "CHV_NFE")
    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
    Set dicDadosD100 = Util.CriarDicionarioRegistro(regD100, "CHV_CTE")
    Set dicDadosE001 = Util.CriarDicionarioRegistro(regE001)
    
    Set dicTitulosTributacao = Util.MapearTitulos(assTributacaoICMS, 3)
    Set dicTitulosApuracao = Util.MapearTitulos(assApuracaoICMS, 3)
    Set dicTitulosCTeNFe = Util.MapearTitulos(CorrelacoesCTeNFe, 3)
    Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
    Set dicTitulosC101 = Util.MapearTitulos(regC101, 3)
    Set dicTitulosD100 = Util.MapearTitulos(regD100, 3)
    Set dicTitulosD101 = Util.MapearTitulos(regD101, 3)
    Set dicTitulosE001 = Util.MapearTitulos(regE001, 3)
    Set dicTitulosE300 = Util.MapearTitulos(regE300, 3)
    Set dicTitulosE310 = Util.MapearTitulos(regE310, 3)
    Set dicTitulosE311 = Util.MapearTitulos(regE311, 3)
    Set dicTitulosE313 = Util.MapearTitulos(regE313, 3)
    Set dicTitulosE316 = Util.MapearTitulos(regE316, 3)
    
    Set AssistTributario.dicTitulosApuracao = dicTitulosApuracao
    Set AssistTributario.dicTitulosTributacao = dicTitulosTributacao
    Set AssistTributario.dicDadosTributarios = dicDadosTributarios
    
    Call CalcularRateioCTe
    
End Sub

Private Function CriarRegistroE001(ByRef Campos As Variant)
    
    With CamposE001
        
        .REG = "E001"
        .ARQUIVO = Campos(dicTitulosApuracao("ARQUIVO"))
        .CHV_PAI = dicDados0000(.ARQUIVO)(dicTitulos0000("CHV_REG"))
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "E001")
        .IND_MOV = 0
        
        dicDadosE001(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", .IND_MOV)
        
    End With
    
End Function

Private Function CriarRegistroE300(ByRef Campos As Variant)

Dim Periodo As String, Chave$
    
    Periodo = VBA.Split(Campos(dicTitulosApuracao("ARQUIVO")), "-")(0)
    With CamposE300
        
        .REG = "E300"
        .ARQUIVO = Campos(dicTitulosApuracao("ARQUIVO"))
        .CHV_PAI = dicDadosE001(CHV_E001)(dicTitulosE001("CHV_REG"))
        .UF = Campos(dicTitulosApuracao("UF_PART"))
        .DT_FIN = VBA.Format(Util.ConverterPeriodoData(Periodo), "yyyy-mm-dd")
        .DT_INI = VBA.Format(.DT_FIN, "yyyy-mm-01")
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .UF, CDate(.DT_INI), CDate(.DT_FIN))
        
        If .UF <> "EX" Then dicDadosE300(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", .UF, .DT_INI, .DT_FIN)
        
    End With
    
End Function

Private Function CriarRegistroE310(ByRef Campos As Variant)
    
    With CamposE310
        
        .REG = "E310"
        .ARQUIVO = Campos(dicTitulosApuracao("ARQUIVO"))
        .CHV_PAI = CHV_E300
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "E310")
        
        dicDadosE310(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        
    End With
    
End Function

Private Function IncluirEstornosDIFALE311_E313(ByRef Campos As Variant)
    
Dim ChaveTrib As String, CHV_NFE$, ESTORNO_DIFAL$
Dim ALIQ_DEST As Double, ALIQ_FCP#, ALIQ_ICMS#, ALIQ_DIFAL#, VL_BC_ICMS#, VL_FCP#, VL_DIFAL#, VL_BC_DIFAL_CTE#, VL_DIFAL_CTE#, VL_FCP_CTE#
    
    ChaveTrib = AssistTributario.GerarChaveTributacao(assApuracaoICMS, Campos)
    VL_BC_ICMS = Campos(dicTitulosApuracao("VL_BC_ICMS"))
    CHV_NFE = Campos(dicTitulosApuracao("CHV_NFE"))

    If dicDadosTributarios.Exists(ChaveTrib) And VL_BC_ICMS > 0 Then
        
        ALIQ_DEST = fnExcel.FormatarPercentuais(dicDadosTributarios(ChaveTrib)(dicTitulosTributacao("ALIQ_ICMS_DEST")))
        If ALIQ_DEST > 0 Then
            
            ALIQ_FCP = fnExcel.ConverterValores(dicDadosTributarios(ChaveTrib)(dicTitulosTributacao("ALIQ_FCP")))
            ALIQ_ICMS = fnExcel.ConverterValores(Campos(dicTitulosApuracao("ALIQ_ICMS")))
            ALIQ_DIFAL = ALIQ_DEST - ALIQ_ICMS
            
            If ALIQ_FCP > 0 Then VL_FCP = VBA.Round(VL_BC_ICMS * ALIQ_FCP, 2)
            If ALIQ_DIFAL > 0 Then VL_DIFAL = VBA.Round(VL_BC_ICMS * ALIQ_DIFAL, 2)
                        
            If VL_DIFAL > 0 Then Call IncluirRegistroE311(Campos, VL_DIFAL, "239999", "ESTORNO DE DÉBITO PARA AJUSTE DE APURAÇÃO DO DIFAL ICMS")
            If VL_FCP > 0 Then Call IncluirRegistroE311(Campos, VL_FCP, "339999", "ESTORNO DE DÉBITO PARA AJUSTE DE APURAÇÃO DO DIFAL FCP")
            
            If Not dicDadosE310.Exists(CHV_E310) Then Call CriarRegistroE310(Campos)
            Call AtualizarRegistroE310(CHV_E310, 0, 0, VL_DIFAL, VL_FCP, 0, 0)
            
            If dicCorrelacoesCTeNFe.Exists(CHV_NFE) Then
                
                ESTORNO_DIFAL = dicCorrelacoesCTeNFe(CHV_NFE)(dicTitulosCTeNFe("ESTORNO_DIFAL"))
                VL_BC_DIFAL_CTE = dicCorrelacoesCTeNFe(CHV_NFE)(dicTitulosCTeNFe("VL_BC_DIFAL_CTE"))
                
                If VL_BC_DIFAL_CTE > 0 And ESTORNO_DIFAL = "SIM" Then
                    
                    VL_DIFAL_CTE = VBA.Round(VL_BC_DIFAL_CTE * ALIQ_DIFAL, 2)
                    VL_FCP_CTE = 0 ' VBA.Round(VL_BC_DIFAL_CTE * ALIQ_FCP, 2)
                    
                    If VL_DIFAL_CTE > 0 Then Call IncluirRegistroE311(Campos, VL_DIFAL_CTE, "239999", "ESTORNO DE DÉBITO PARA AJUSTE DE APURAÇÃO DO DIFAL ICMS")
                    If VL_FCP_CTE > 0 Then Call IncluirRegistroE311(Campos, VL_FCP_CTE, "339999", "ESTORNO DE DÉBITO PARA AJUSTE DE APURAÇÃO DO DIFAL FCP")
                    
                    If Not dicDadosE310.Exists(CHV_E310) Then Call CriarRegistroE310(Campos)
                    Call AtualizarRegistroE310(CHV_E310, 0, 0, VL_DIFAL_CTE, VL_FCP_CTE, 0, 0)
                    
                End If
            
            End If

        End If
        
    End If
    
End Function

Private Function GerarCodigoAjusteDebitoDIFAL(ByVal Campos As Variant) As String
    
    Select Case Campos(dicTitulosApuracao("UF_PART"))
        
        Case "GO"
            GerarCodigoAjusteDebitoDIFAL = "200000"
            
        Case "SE", "RO"
            GerarCodigoAjusteDebitoDIFAL = "200001"
            
        Case Else
            GerarCodigoAjusteDebitoDIFAL = "209999"
            
    End Select
    
End Function

Private Sub IncluirRegistroE311(ByRef Campos As Variant, ByVal VL_AJUSTE As Double, ByVal COD_AJUSTE As String, ByVal DESCR_AJUSTE As String, Optional CHV_CTE As String = "")

Dim Chave As String, UF As String
            
    UF = Campos(dicTitulosApuracao("UF_PART"))
    
    With CamposE311
        
        .REG = "E311"
        .ARQUIVO = Campos(dicTitulosApuracao("ARQUIVO"))
        .COD_AJ_APUR = Campos(dicTitulosApuracao("UF_PART")) & COD_AJUSTE
        .CHV_PAI = dicDadosE310(CHV_E310)(dicTitulosE310("CHV_REG") - 1)
        .DESCR_COMPL_AJ = DESCR_AJUSTE
        .VL_AJ_APUR = VL_AJUSTE
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .COD_AJ_APUR, .DESCR_COMPL_AJ)
        
        CHV_E311 = .CHV_REG
        If dicDadosE311.Exists(.CHV_REG) Then Call AtualizarRegistroE311(.CHV_REG)
        
        dicDadosE311(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", .COD_AJ_APUR, .DESCR_COMPL_AJ, CDbl(.VL_AJ_APUR))
        
    End With
    
    Call IncluirRegistroE313(Campos, VL_AJUSTE, CHV_CTE)
    
End Sub

Private Function AtualizarRegistroE311(ByVal Chave As String)
    
    With CamposE311
                         
        .VL_AJ_APUR = .VL_AJ_APUR + CDbl(dicDadosE311(Chave)(dicTitulosE311("VL_AJ_APUR") - 1))
        
    End With
    
End Function

Private Sub IncluirRegistroE313(ByRef Campos As Variant, ByVal VL_AJUSTE As Double, Optional CHV_CTE As String = "")
    
Dim Chave As String

    If CHV_CTE <> "" Then
        
        Call CarregarDadosCTeRegistroE313(Campos, VL_AJUSTE, CHV_CTE)
        
    Else
        
        With CamposE313
            
            .REG = "E313"
            .ARQUIVO = Campos(dicTitulosApuracao("ARQUIVO"))
            .CHV_PAI = dicDadosE311(CHV_E311)(dicTitulosE311("CHV_REG") - 1)
            .COD_PART = Campos(dicTitulosApuracao("COD_PART"))
            .COD_MOD = Campos(dicTitulosApuracao("COD_MOD"))
            .SER = VBA.Format(Campos(dicTitulosApuracao("SER")), "000")
            .SUB = ""
            .NUM_DOC = Campos(dicTitulosApuracao("NUM_DOC"))
            .CHV_DOCE = Campos(dicTitulosApuracao("CHV_NFE"))
            .DT_DOC = Campos(dicTitulosApuracao("DT_DOC"))
            .COD_ITEM = Campos(dicTitulosApuracao("COD_ITEM"))
            .VL_AJ_ITEM = VL_AJUSTE
            
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .COD_PART, .COD_MOD, "'" & .SER, .SUB, .NUM_DOC, .CHV_DOCE, .DT_DOC, .COD_ITEM)
            
            If dicDadosE313.Exists(.CHV_REG) Then Call AtualizarRegistroE313(.CHV_REG)
            
            dicDadosE313(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", "'" & .COD_PART, _
                .COD_MOD, "'" & .SER, .SUB, .NUM_DOC, "'" & .CHV_DOCE, .DT_DOC, "'" & .COD_ITEM, CDbl(.VL_AJ_ITEM))
            
        End With
    
    End If
    
End Sub

Private Sub CarregarDadosCTeRegistroE313(ByRef Campos As Variant, ByVal VL_AJUSTE As Double, ByVal CHV_CTE As String)
        
Dim CamposD100 As Variant

    If dicDadosD100.Exists(CHV_CTE) Then
        
        CamposD100 = dicDadosD100(CHV_CTE)
        
        With CamposE313
            
            .REG = "E313"
            .ARQUIVO = Campos(dicTitulosApuracao("ARQUIVO"))
            .CHV_PAI = dicDadosE311(CHV_E311)(dicTitulosE311("CHV_REG") - 1)
            .COD_PART = CamposD100(dicTitulosD100("COD_PART"))
            .COD_MOD = CamposD100(dicTitulosD100("COD_MOD"))
            .SER = VBA.Format(CamposD100(dicTitulosD100("SER")), "000")
            .SUB = ""
            .NUM_DOC = CamposD100(dicTitulosD100("NUM_DOC"))
            .CHV_DOCE = CamposD100(dicTitulosD100("CHV_CTE"))
            .DT_DOC = CamposD100(dicTitulosD100("DT_A_P"))
            .VL_AJ_ITEM = VL_AJUSTE
            
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .COD_PART, .COD_MOD, "'" & .SER, .SUB, .NUM_DOC, .CHV_DOCE, .DT_DOC, .COD_ITEM)
            
            If dicDadosE313.Exists(.CHV_REG) Then Call AtualizarRegistroE313(.CHV_REG)
            
            dicDadosE313(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", "'" & .COD_PART, _
                .COD_MOD, "'" & .SER, .SUB, .NUM_DOC, "'" & .CHV_DOCE, .DT_DOC, "'" & .COD_ITEM, CDbl(.VL_AJ_ITEM))
            
        End With
        
    End If
    
End Sub

Private Function AtualizarRegistroE313(ByVal Chave As String)
    
    With CamposE313
        
        .VL_AJ_ITEM = .VL_AJ_ITEM + CDbl(dicDadosE313(Chave)(dicTitulosE313("VL_AJ_ITEM") - 1))
        
    End With
    
End Function

Private Sub ExportarRegistros()
    
'Dim ExpReg As New ExportadorRegistros
    
    'Call ExpReg.ExportarRegistros("C101", "D101", "E001", "E300", "E310", "E311", "E313", "E316")
    
    Set dicDadosC101 = Util.MesclarDicionarios(dicDadosC101, regC101)
    Call Util.LimparDados(regC101, 4, False)
    Call Util.ExportarDadosDicionario(regC101, dicDadosC101)
            
    Set dicDadosD101 = Util.MesclarDicionarios(dicDadosD101, regD101)
    Call Util.LimparDados(regD101, 4, False)
    Call Util.ExportarDadosDicionario(regD101, dicDadosD101)
    
    Set dicDadosE001 = Util.MesclarDicionarios(dicDadosE001, regE001)
    Call Util.LimparDados(regE001, 4, False)
    Call Util.ExportarDadosDicionario(regE001, dicDadosE001)
    
    Set dicDadosE300 = Util.MesclarDicionarios(dicDadosE300, regE300)
    Call Util.LimparDados(regE300, 4, False)
    Call Util.ExportarDadosDicionario(regE300, dicDadosE300)

    Set dicDadosE310 = Util.MesclarDicionarios(dicDadosE310, regE310)
    Call Util.LimparDados(regE310, 4, False)
    Call Util.ExportarDadosDicionario(regE310, dicDadosE310)

    Set dicDadosE311 = Util.MesclarDicionarios(dicDadosE311, regE311)
    Call Util.LimparDados(regE311, 4, False)
    Call Util.ExportarDadosDicionario(regE311, dicDadosE311)
    
    Set dicDadosE313 = Util.MesclarDicionarios(dicDadosE313, regE313)
    Call Util.LimparDados(regE313, 4, False)
    Call Util.ExportarDadosDicionario(regE313, dicDadosE313)
    
    Set dicDadosE316 = Util.MesclarDicionarios(dicDadosE316, regE316)
    Call Util.LimparDados(regE316, 4, False)
    Call Util.ExportarDadosDicionario(regE316, dicDadosE316)
    
    Call Util.LimparDados(CorrelacoesCTeNFe, 4, False)
    Call Util.ExportarDadosDicionario(CorrelacoesCTeNFe, dicCorrelacoesCTeNFe)
    
End Sub

Private Function AtualizarRegistroE310(ByRef Chave As String, ByVal VL_DIFAL As Double, ByVal VL_FCP As Double, ByVal EST_DIFAL As Double, _
    ByVal EST_FCP As Double, ByVal DEBITO_DIFAL As Double, ByVal DEBITO_FCP As Double, Optional CalcularSaldo As Boolean = False)
    
Dim Campos As Variant
    
    Campos = dicDadosE310(Chave)
    With CamposE310
        
        'DIFAL ICMS
        .REG = Campos(dicTitulosE310("REG") - 1)
        .ARQUIVO = Campos(dicTitulosE310("ARQUIVO") - 1)
        .CHV_PAI = Campos(dicTitulosE310("CHV_PAI_FISCAL") - 1)
        .CHV_REG = Campos(dicTitulosE310("CHV_REG") - 1)
        .IND_MOV_FCP_DIFAL = Campos(dicTitulosE310("IND_MOV_FCP_DIFAL") - 1)
        .VL_SLD_CRED_ANT_DIFAL = CDbl(Campos(dicTitulosE310("VL_SLD_CRED_ANT_DIFAL") - 1))
        .VL_TOT_DEBITOS_DIFAL = CDbl(Campos(dicTitulosE310("VL_TOT_DEBITOS_DIFAL") - 1)) + VL_DIFAL
        .VL_OUT_DEB_DIFAL = CDbl(Campos(dicTitulosE310("VL_OUT_DEB_DIFAL") - 1)) + DEBITO_DIFAL
        .VL_TOT_CREDITOS_DIFAL = CDbl(Campos(dicTitulosE310("VL_TOT_CREDITOS_DIFAL") - 1))
        .VL_OUT_CRED_DIFAL = CDbl(Campos(dicTitulosE310("VL_OUT_CRED_DIFAL") - 1)) + EST_DIFAL
        .VL_SLD_DEV_ANT_DIFAL = (.VL_TOT_DEBITOS_DIFAL + CDbl(.VL_OUT_DEB_DIFAL)) - (.VL_SLD_CRED_ANT_DIFAL + .VL_TOT_CREDITOS_DIFAL + CDbl(.VL_OUT_CRED_DIFAL))
        .VL_DEDUCOES_DIFAL = CDbl(Campos(dicTitulosE310("VL_DEDUCOES_DIFAL") - 1))
        .VL_RECOL_DIFAL = .VL_SLD_DEV_ANT_DIFAL - CDbl(.VL_DEDUCOES_DIFAL)
        .VL_SLD_CRED_TRANSPORTAR_DIFAL = CDbl(Campos(dicTitulosE310("VL_SLD_CRED_TRANSPORTAR_DIFAL") - 1))
        .DEB_ESP_DIFAL = CDbl(Campos(dicTitulosE310("DEB_ESP_DIFAL") - 1))
        
        If .VL_SLD_DEV_ANT_DIFAL < 0 And CalcularSaldo Then
            
            .VL_SLD_CRED_TRANSPORTAR_DIFAL = VBA.Abs(.VL_SLD_DEV_ANT_DIFAL)
            .VL_SLD_DEV_ANT_DIFAL = 0
            .VL_RECOL_DIFAL = 0
            
        End If
        
        'DIFAL FCP
        .VL_SLD_CRED_ANT_FCP = CDbl(Campos(dicTitulosE310("VL_SLD_CRED_ANT_FCP") - 1))
        .VL_TOT_DEB_FCP = CDbl(Campos(dicTitulosE310("VL_TOT_DEB_FCP") - 1)) + VL_FCP
        .VL_OUT_DEB_FCP = CDbl(Campos(dicTitulosE310("VL_OUT_DEB_FCP") - 1)) + DEBITO_FCP
        .VL_TOT_CRED_FCP = CDbl(Campos(dicTitulosE310("VL_TOT_CRED_FCP") - 1))
        .VL_OUT_CRED_FCP = CDbl(Campos(dicTitulosE310("VL_OUT_CRED_FCP") - 1)) + EST_FCP
        .VL_SLD_DEV_ANT_FCP = (.VL_TOT_DEB_FCP + CDbl(.VL_OUT_DEB_FCP)) - (.VL_SLD_CRED_ANT_FCP + .VL_TOT_CRED_FCP + CDbl(.VL_OUT_CRED_FCP))
        .VL_DEDUCOES_FCP = CDbl(Campos(dicTitulosE310("VL_DEDUCOES_FCP") - 1))
        .VL_RECOL_FCP = .VL_SLD_DEV_ANT_FCP - CDbl(.VL_DEDUCOES_FCP)
        .VL_SLD_CRED_TRANSPORTAR_FCP = CDbl(Campos(dicTitulosE310("VL_SLD_CRED_TRANSPORTAR_FCP") - 1))
        .DEB_ESP_FCP = CDbl(Campos(dicTitulosE310("DEB_ESP_FCP") - 1))
        
        If .VL_SLD_DEV_ANT_FCP < 0 And CalcularSaldo Then
            
            .VL_SLD_CRED_TRANSPORTAR_FCP = VBA.Abs(.VL_SLD_DEV_ANT_FCP)
            .VL_SLD_DEV_ANT_FCP = 0
            .VL_RECOL_FCP = 0
            
        End If
        
        dicDadosE310(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", .IND_MOV_FCP_DIFAL, _
            CDbl(.VL_SLD_CRED_ANT_DIFAL), CDbl(.VL_TOT_DEBITOS_DIFAL), CDbl(.VL_OUT_DEB_DIFAL), _
            CDbl(.VL_TOT_CREDITOS_DIFAL), CDbl(.VL_OUT_CRED_DIFAL), CDbl(.VL_SLD_DEV_ANT_DIFAL), _
            CDbl(.VL_DEDUCOES_DIFAL), CDbl(.VL_RECOL_DIFAL), CDbl(.VL_SLD_CRED_TRANSPORTAR_DIFAL), _
            CDbl(.DEB_ESP_DIFAL), CDbl(.VL_SLD_CRED_ANT_FCP), CDbl(.VL_TOT_DEB_FCP), CDbl(.VL_OUT_DEB_FCP), _
            CDbl(.VL_TOT_CRED_FCP), CDbl(.VL_OUT_CRED_FCP), CDbl(.VL_SLD_DEV_ANT_FCP), CDbl(.VL_DEDUCOES_FCP), _
            CDbl(.VL_RECOL_FCP), CDbl(.VL_SLD_CRED_TRANSPORTAR_FCP), CDbl(.DEB_ESP_FCP))
        
        If CalcularSaldo Then
        
            If .VL_RECOL_DIFAL > 0 Then Call IncluirRegistroE316(.ARQUIVO, "'000", "100110", .VL_RECOL_DIFAL)
            If .VL_RECOL_FCP > 0 Then Call IncluirRegistroE316(.ARQUIVO, "'006", "100137", .VL_RECOL_FCP)
        
        End If
        
    End With
    
End Function

Private Function CalcularSaldosE310()

Dim Chave As Variant
    
    For Each Chave In dicDadosE310.Keys()
        
        CHV_E310 = Chave
        Call AtualizarRegistroE310(CStr(Chave), 0, 0, 0, 0, 0, 0, True)

    Next Chave

End Function

Private Function IncluirRegistroE316(ByRef ARQUIVO As String, ByVal COD_OR As String, ByVal COD_REC As String, ByVal VL_OR As Double)
    
    With CamposE316
        
        .REG = "E316"
        .ARQUIVO = ARQUIVO
        .CHV_PAI = CHV_E310
        .COD_OR = COD_OR
        .VL_OR = VL_OR
        .DT_VCTO = VBA.Format("09/" & VBA.Split(ARQUIVO, "-")(0), "yyyy-mm-dd")
        .COD_REC = COD_REC
        .NUM_PROC = ""
        .IND_PROC = ""
        .PROC = ""
        .TXT_COMPL = ""
        .MES_REF = fnExcel.FormatarTexto(VBA.Replace(VBA.Split(ARQUIVO, "-")(0), "/", ""))
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .COD_OR, .COD_REC, .IND_PROC)
        dicDadosE316(.CHV_REG) = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, "", .COD_OR, _
            CDbl(.VL_OR), .DT_VCTO, .COD_REC, .NUM_PROC, .IND_PROC, .PROC, .TXT_COMPL, .MES_REF)
        
    End With
    
End Function

Private Function CalcularRateioCTe()

Dim chCTe As String
Dim Campos As Variant, Chave
Dim dicDadosRateio As New Dictionary
Dim vCTe As Double, vMercadorias#, vProdutos#, vRateio#
    
    Call AtualizarValorMercadoriasC100
    
    For Each Campos In dicCorrelacoesCTeNFe.Items()
        
        chCTe = Campos(dicTitulosCTeNFe("CHV_CTE"))
        vCTe = Campos(dicTitulosCTeNFe("VL_CTE"))
        vMercadorias = Campos(dicTitulosCTeNFe("VL_MERCADORIAS"))
        
        If dicDadosRateio.Exists(chCTe) Then vMercadorias = vMercadorias + dicDadosRateio(chCTe)(1)
        dicDadosRateio(chCTe) = Array(vCTe, vMercadorias)
        
    Next Campos
    
    For Each Chave In dicCorrelacoesCTeNFe.Keys()
        
        If Chave <> "" Then
            
            Campos = dicCorrelacoesCTeNFe(Chave)
            
            chCTe = Campos(dicTitulosCTeNFe("CHV_CTE"))
            vProdutos = Campos(dicTitulosCTeNFe("VL_MERCADORIAS"))
            
            vCTe = dicDadosRateio(chCTe)(0)
            vMercadorias = dicDadosRateio(chCTe)(1)
            
            If vMercadorias > 0 Then vRateio = VBA.Round(vProdutos / vMercadorias * vCTe, 2) Else vRateio = 0
            
            Campos(dicTitulosCTeNFe("VL_BC_DIFAL_CTE")) = vRateio
            
            dicCorrelacoesCTeNFe(Chave) = Campos
        
        End If
        
    Next Chave
    
End Function

Private Sub AtualizarValorMercadoriasC100()

Dim VL_MERC As Double
Dim Chave As Variant, Campos
Dim dicDadosC100 As New Dictionary
Dim COD_SIT As String, OBSERVACOES$
Dim dicTitulosC100 As New Dictionary
    
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100, "CHV_NFE")
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    
    For Each Chave In dicDadosC100.Keys()
        
        VL_MERC = dicDadosC100(Chave)(dicTitulosC100("VL_MERC"))
        COD_SIT = dicDadosC100(Chave)(dicTitulosC100("COD_SIT"))
        
        Select Case True
            
            Case COD_SIT Like "02*" Or COD_SIT Like "03*"
                OBSERVACOES = "DOCUMENTO CANCELADO"
                
            Case COD_SIT Like "04*"
                OBSERVACOES = "DOCUMENTO DENEGADO"
                
            Case COD_SIT Like "05*"
                OBSERVACOES = "DOCUMENTO INUTILIZADO"
                
            Case Else
                OBSERVACOES = ""
                
        End Select
        
        If dicCorrelacoesCTeNFe.Exists(Chave) Then
            
            Campos = dicCorrelacoesCTeNFe(Chave)
            Campos(dicTitulosCTeNFe("VL_MERCADORIAS")) = VL_MERC
            Campos(dicTitulosCTeNFe("DEBITO_DIFAL")) = "NÃO"
            Campos(dicTitulosCTeNFe("ESTORNO_DIFAL")) = "NÃO"
            If VBA.Len(OBSERVACOES) > 0 Then Campos(dicTitulosCTeNFe("OBSERVACOES")) = OBSERVACOES
            
            dicCorrelacoesCTeNFe(Chave) = Campos
            
        End If
        
    Next Chave
    
End Sub

Private Sub AtualizarCorrelacaoCTeNFe(ByVal Chave As String, ByVal Campo As String, ByVal Valor As String, ByVal UF As String)

    Dim Campos As Variant
    
    If dicCorrelacoesCTeNFe.Exists(Chave) Then
        
        Campos = dicCorrelacoesCTeNFe(Chave)
        Campos(dicTitulosCTeNFe("UF_DESTINO")) = UF
        Campos(dicTitulosCTeNFe(Campo)) = Valor
        
        dicCorrelacoesCTeNFe(Chave) = Campos
        
    End If

End Sub




