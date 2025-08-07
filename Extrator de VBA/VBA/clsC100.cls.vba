Attribute VB_Name = "clsC100"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Public Function ImportarParaAnalise(ByVal Registro As String, ByRef dicDadosC100 As Dictionary)

Dim Campos
Dim Chave As String
    
    Campos = fnSPED.ExtrairCamposRegistro(Registro)
    
    With RelDiverg
    
        .DOC_PART = ""
        .DOC_CONTRIB = "'" & fnSPED.GerarChaveRegistro(Campos0000.CNPJ, Campos0000.CPF)
        .Operacao = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_OPER(Campos(2))
        .TP_EMISSAO = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_EMIT(Campos(3))
        .COD_PART = Util.FormatarTexto(Campos(4))
        .Modelo = Campos(5)
        .Situacao = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_SIT(Campos(6))
        .SERIE = Util.FormatarTexto(VBA.Format(Campos(7), "000"))
        .NUM_DOC = Util.FormatarTexto(VBA.Format(Campos(8), String(9, "0")))
        .CHV_NFE = Util.FormatarTexto(Campos(9))
        .DT_DOC = Util.FormatarData(Campos(10))
        .TP_PAGAMENTO = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_PGTO(Campos(13))
        .TP_FRETE = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_IND_FRT(Campos(17))
        .VL_DOC = Util.FormatarValores(Campos(12))
        .VL_DESC = Util.FormatarValores(Campos(14))
        .VL_ABATIMENTO = Util.FormatarValores(Campos(15))
        .VL_PROD = Util.FormatarValores(Campos(16))
        .VL_FRETE = Util.FormatarValores(Campos(18))
        .VL_SEG = Util.FormatarValores(Campos(19))
        .VL_OUTRO = Util.FormatarValores(Campos(20))
        .VL_BC_ICMS = Util.FormatarValores(Campos(21))
        .VL_ICMS = Util.FormatarValores(Campos(22))
        .VL_BC_ICMS_ST = Util.FormatarValores(Campos(23))
        .VL_ICMS_ST = Util.FormatarValores(Campos(24))
        .VL_IPI = Util.FormatarValores(Campos(25))
        .VL_PIS = Util.FormatarValores(Campos(26))
        .VL_COFINS = Util.FormatarValores(Campos(27))
        
        Chave = Util.RemoverAspaSimples(.COD_PART)
        If regEFD.dic0150.Exists(Chave) Then
            .DOC_PART = regEFD.dic0150(Chave)(5) & regEFD.dic0150(Chave)(6)
        End If
        
        Select Case True
            
            Case DesconsiderarPISCOFINS = True
                .VL_PIS = 0
                .VL_COFINS = 0
                
            Case DesconsiderarAbatimento = True
                .VL_ABATIMENTO = 0
                
        End Select
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.COD_PART, .Modelo, .Situacao, .SERIE, .NUM_DOC, .CHV_NFE)
        If (.Modelo = "55") Or (.Modelo = "65") Then
        
            dicDadosC100(.CHV_REG) = Array(.DOC_CONTRIB, .DOC_PART, .Modelo, .Operacao, .TP_EMISSAO, _
                .Situacao, .SERIE, .NUM_DOC, .CHV_NFE, .DT_DOC, .TP_PAGAMENTO, .TP_FRETE, CDbl(.VL_DOC), _
                CDbl(.VL_PROD), CDbl(.VL_FRETE), CDbl(.VL_SEG), CDbl(.VL_OUTRO), CDbl(.VL_DESC), _
                CDbl(.VL_ABATIMENTO), CDbl(.VL_BC_ICMS), CDbl(.VL_ICMS), CDbl(.VL_BC_ICMS_ST), _
                CDbl(.VL_ICMS_ST), CDbl(.VL_IPI), CDbl(.VL_PIS), CDbl(.VL_COFINS), "XML AUSENTE")
        
        End If
        
    End With
    
End Function

Public Sub RatearDivergenciasC100ParaC190()

Dim VL_DIF As Double, VL_DOC#, VL_MERC#, VL_DESC#, VL_FRT#, VL_SEG#, VL_OUT_DA#, VL_ICMS_ST#, VL_IPI#, VL_OPR#, VL_DOC_CALC#, VL_OPR_TOTAL#, Margem#, VL_ABAT_NT#
Dim CHV_REG_C100 As String, CHV_REG$, IND_OPER$, IND_EMIT$, COD_SIT$
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC190 As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim dicDadosC190 As New Dictionary
Dim dicValorOpr As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim Campos As Variant
    
    Inicio = Now()
    Call Util.AtualizarBarraStatus("Ajustando divergencias, por favor aguarde...")
    
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    
    Set dicValorOpr = Util.CriarDicionarioValorOperacoesC190(regC190)
    Set dicTitulosC190 = Util.MapearTitulos(regC190, 3)
    Set Dados = Util.DefinirIntervalo(regC190, 4, 3)
    
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não existem dados nos registros C190", "Dados indisponíveis")
        Exit Sub
    End If
    
    a = 0
    Comeco = Timer
    Margem = 0.02
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Ajustando divergencias, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            CHV_REG_C100 = Campos(dicTitulosC190("CHV_PAI_FISCAL"))
            If dicDadosC100.Exists(CHV_REG_C100) Then
                
                IND_OPER = VBA.Left(dicDadosC100(CHV_REG_C100)(dicTitulosC100("IND_OPER")), 1)
                IND_EMIT = VBA.Left(dicDadosC100(CHV_REG_C100)(dicTitulosC100("IND_EMIT")), 1)
                COD_SIT = VBA.Left(dicDadosC100(CHV_REG_C100)(dicTitulosC100("COD_SIT")), 2)
                VL_DOC = dicDadosC100(CHV_REG_C100)(dicTitulosC100("VL_DOC"))

                'If IND_OPER <> "0" And IND_EMIT <> "1" Then GoTo Prx:
                If (Not COD_SIT Like "00") And (Not COD_SIT Like "01") And (Not COD_SIT Like "08") Then GoTo Prx:
                
                If dicValorOpr.Exists(CHV_REG_C100) Then
                    VL_OPR_TOTAL = dicValorOpr(CHV_REG_C100)
                End If
                
                If VBA.Abs(VL_OPR_TOTAL - VL_DOC) < Margem Then GoTo Prx:
                
                VL_MERC = dicDadosC100(CHV_REG_C100)(dicTitulosC100("VL_MERC"))
                VL_DESC = dicDadosC100(CHV_REG_C100)(dicTitulosC100("VL_DESC"))
                VL_FRT = dicDadosC100(CHV_REG_C100)(dicTitulosC100("VL_FRT"))
                VL_OUT_DA = dicDadosC100(CHV_REG_C100)(dicTitulosC100("VL_OUT_DA"))
                VL_ICMS_ST = dicDadosC100(CHV_REG_C100)(dicTitulosC100("VL_ICMS_ST"))
                VL_IPI = dicDadosC100(CHV_REG_C100)(dicTitulosC100("VL_IPI"))
                VL_ABAT_NT = dicDadosC100(CHV_REG_C100)(dicTitulosC100("VL_ABAT_NT"))
                VL_DOC_CALC = VL_MERC + VL_FRT + VL_SEG + VL_OUT_DA + VL_ICMS_ST + VL_IPI - VL_DESC
                
                VL_DIF = VBA.Round(VL_DOC - VL_DOC_CALC, 2)
                If VBA.Abs(VL_DIF) > Margem Then
                
                    VL_DOC_CALC = VL_DOC_CALC - VL_ABAT_NT
                    VL_DIF = VBA.Round(VL_DOC - VL_DOC_CALC, 2)
                
                End If
                
            End If
            
            VL_DIF = VBA.Round(VL_DOC - VL_OPR_TOTAL, 2)
            If VL_DIF <> 0 Then
                
                VL_OPR = Campos(dicTitulosC190("VL_OPR"))
                Select Case VL_OPR
                
                    Case Is > 0
                        Campos(dicTitulosC190("VL_OPR")) = VL_OPR + VBA.Round(VBA.Round(VL_OPR / VL_OPR_TOTAL, 15) * VL_DIF, 2)
                    
                    Case Is = 0
                        Campos(dicTitulosC190("VL_OPR")) = VBA.Round(VL_DIF, 2)
                
                End Select
                
            End If
            
Prx:
            CHV_REG = Campos(dicTitulosC190("CHV_REG"))
            dicDadosC190(CHV_REG) = Campos
            VL_DIF = 0
            
        End If
        
    Next Linha
    
    Call Util.AtualizarBarraStatus("Atualizando valores do C190, por favor aguarde...")
    Call Util.LimparDados(regC190, 4, False)
    Call Util.ExportarDadosDicionario(regC190, dicDadosC190)

    Application.StatusBar = False
    Call Util.MsgInformativa("Rateios efetuados com sucesso!", "Rateio de divergências do C100 para o C190", Inicio)
    
End Sub
