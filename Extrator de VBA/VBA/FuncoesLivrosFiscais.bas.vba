Attribute VB_Name = "FuncoesLivrosFiscais"
Option Explicit
Option Base 1

Public Function GerarLivroICMS()

Dim dicDados0000 As New Dictionary
Dim dicDados0150 As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim dicDadosC190 As New Dictionary
Dim dicDadosC500 As New Dictionary
Dim dicDadosC590 As New Dictionary
Dim dicDadosC800 As New Dictionary
Dim dicDadosC850 As New Dictionary
Dim dicDadosD100 As New Dictionary
Dim dicDadosD190 As New Dictionary
Dim dicDadosD500 As New Dictionary
Dim dicDadosD590 As New Dictionary
Dim dicLivroICMS As New Dictionary
Dim dicTitulos0000 As New Dictionary
Dim dicTitulos0150 As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC190 As New Dictionary
Dim dicTitulosC500 As New Dictionary
Dim dicTitulosC590 As New Dictionary
Dim dicTitulosC800 As New Dictionary
Dim dicTitulosC850 As New Dictionary
Dim dicTitulosD100 As New Dictionary
Dim dicTitulosD190 As New Dictionary
Dim dicTitulosD500 As New Dictionary
Dim dicTitulosD590 As New Dictionary
Dim dicTitulosRel As New Dictionary
Dim Msg As String
    
    Inicio = Now()
    Application.StatusBar = "Gerando livro ICMS, por favor aguarde..."
    
    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
    Set dicDados0150 = Util.CriarDicionarioRegistro(reg0150)
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    Set dicDadosC190 = Util.CriarDicionarioRegistro(regC190)
    Set dicDadosC500 = Util.CriarDicionarioRegistro(regC500)
    Set dicDadosC590 = Util.CriarDicionarioRegistro(regC590)
    Set dicDadosC800 = Util.CriarDicionarioRegistro(regC800)
    Set dicDadosC850 = Util.CriarDicionarioRegistro(regC850)
    Set dicDadosD100 = Util.CriarDicionarioRegistro(regD100)
    Set dicDadosD190 = Util.CriarDicionarioRegistro(regD190)
    Set dicDadosD500 = Util.CriarDicionarioRegistro(regD500)
    Set dicDadosD590 = Util.CriarDicionarioRegistro(regD590)
    
    Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
    Set dicTitulos0150 = Util.MapearTitulos(reg0150, 3)
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    Set dicTitulosC190 = Util.MapearTitulos(regC190, 3)
    Set dicTitulosC500 = Util.MapearTitulos(regC500, 3)
    Set dicTitulosC590 = Util.MapearTitulos(regC590, 3)
    Set dicTitulosC800 = Util.MapearTitulos(regC800, 3)
    Set dicTitulosC850 = Util.MapearTitulos(regC850, 3)
    Set dicTitulosD100 = Util.MapearTitulos(regD100, 3)
    Set dicTitulosD190 = Util.MapearTitulos(regD190, 3)
    Set dicTitulosD500 = Util.MapearTitulos(regD500, 3)
    Set dicTitulosD590 = Util.MapearTitulos(regD590, 3)
    Set dicTitulosRel = Util.MapearTitulos(relICMS, 3)
    
    Call MontarRelatorioICMS(dicLivroICMS, dicDados0000, dicDados0150, dicDadosC100, dicDadosC190, dicTitulos0000, dicTitulos0150, dicTitulosC100, dicTitulosC190, dicTitulosRel)
    Call MontarRelatorioICMS(dicLivroICMS, dicDados0000, dicDados0150, dicDadosC500, dicDadosC590, dicTitulos0000, dicTitulos0150, dicTitulosC500, dicTitulosC590, dicTitulosRel)
    Call MontarRelatorioICMS(dicLivroICMS, dicDados0000, dicDados0150, dicDadosC800, dicDadosC850, dicTitulos0000, dicTitulos0150, dicTitulosC800, dicTitulosC850, dicTitulosRel)
    Call MontarRelatorioICMS(dicLivroICMS, dicDados0000, dicDados0150, dicDadosD100, dicDadosD190, dicTitulos0000, dicTitulos0150, dicTitulosD100, dicTitulosD190, dicTitulosRel)
    
    Call Util.LimparDados(relICMS, 4, False)
    Call Util.ExportarDadosDicionario(relICMS, dicLivroICMS)
    
    Call FuncoesFormatacao.AplicarFormatacao(relICMS)
    Application.StatusBar = "Auditando escrituração, por favor aguarde..."
    Call FuncoesFormatacao.FormatarInconsistencias(relICMS)
    
    Application.StatusBar = False
    
    If dicLivroICMS.Count > 0 Then
        
        Call Util.MsgInformativa("Livro de operações do ICMS gerado com sucesso!", "Geração do Livro ICMS", Inicio)
        
    Else
        
        Msg = "Nenhum dado foi importado na automação!"
        Msg = Msg & vbCrLf & vbCrLf & "É preciso importar os dados do SPED Fiscal antes de gerar o livro."
        
        Call Util.MsgAlerta(Msg, "Geração do Livro ICMS")
        
    End If
    
End Function

Public Function MontarRelatorioICMS(ByRef dicDadosRel As Dictionary, ByRef dicDadosArq As Dictionary, ByRef dicDadosParticipantes As Dictionary, _
                                    ByRef dicDadosPai As Dictionary, ByRef dicDadosFilho As Dictionary, ByRef dicTitulosArq As Dictionary, _
                                    ByRef dicTitulosParticipantes As Dictionary, ByRef dicTitulosPai As Dictionary, _
                                    ByRef dicTitulosFilho As Dictionary, ByRef dicTitulosRel As Dictionary)

Dim Campos As Variant, Campo, nCampo
Dim arrCampos As New ArrayList
Dim Valores As Variant
Dim Chave As String
    
    For Each Campos In dicDadosFilho.Items
        
        With relatICMS
            
            .REG = Campos(dicTitulosFilho("REG"))
            If dicDadosPai.Exists(Campos(dicTitulosFilho("CHV_PAI_FISCAL"))) Then .COD_MOD = dicDadosPai(Campos(dicTitulosFilho("CHV_PAI_FISCAL")))(dicTitulosPai("COD_MOD")) Else .COD_MOD = ""
            .CFOP = VBA.Replace(Campos(dicTitulosFilho("CFOP")), "'", "")
            .CST_ICMS = fnExcel.FormatarTexto(Campos(dicTitulosFilho("CST_ICMS")))
            .ALIQ_ICMS = fnExcel.FormatarPercentuais(Campos(dicTitulosFilho("ALIQ_ICMS")))
            .VL_OPR = Util.ValidarValores(Campos(dicTitulosFilho("VL_OPR")))
            .VL_BC_ICMS = Util.ValidarValores(Campos(dicTitulosFilho("VL_BC_ICMS")))
            .VL_ICMS = Util.ValidarValores(Campos(dicTitulosFilho("VL_ICMS")))
            .ALIQ_FCP = 0
            .VL_FCP = 0
            
            If Not IsEmpty(dicTitulosFilho("VL_BC_ICMS_ST")) Then .VL_BC_ICMS_ST = Campos(dicTitulosFilho("VL_BC_ICMS_ST")) Else .VL_BC_ICMS_ST = 0
            If Not IsEmpty(dicTitulosFilho("VL_ICMS_ST")) Then .VL_ICMS_ST = Campos(dicTitulosFilho("VL_ICMS_ST")) Else .VL_ICMS_ST = 0
            If Not IsEmpty(dicTitulosFilho("VL_RED_BC")) Then .VL_RED_BC = Campos(dicTitulosFilho("VL_RED_BC")) Else .VL_RED_BC = 0
            If Not IsEmpty(dicTitulosFilho("VL_IPI")) Then .VL_IPI = Campos(dicTitulosFilho("VL_IPI")) Else .VL_IPI = 0
            
            Call CalcularIsentasOutras
            If dicDadosArq.Exists(Campos(dicTitulosFilho("ARQUIVO"))) Then
            
            End If
            'TODO: Criar lógica para trazer a UF da operação
            'If Not IsEmpty(dicTitulosFilho("UF")) Then .UF = Campos(dicTitulosFilho("UF")) Else .UF = dicDadosArq(Campos(dicTitulosFilho("ARQUIVO"))(DICTOTULOSARQ())
            
            .CHV_REG = fnSPED.GerarChaveRegistro(.REG, .COD_MOD, .UF, .CFOP, .CST_ICMS, .ALIQ_ICMS)
            If dicDadosRel.Exists(.CHV_REG) Then
                
                .VL_OPR = dicDadosRel(.CHV_REG)(dicTitulosRel("VL_OPR")) + CDbl(.VL_OPR)
                .VL_BC_ICMS = dicDadosRel(.CHV_REG)(dicTitulosRel("VL_BC_ICMS")) + CDbl(.VL_BC_ICMS)
                .VL_ICMS = dicDadosRel(.CHV_REG)(dicTitulosRel("VL_ICMS")) + CDbl(.VL_ICMS)
                .VL_FCP = dicDadosRel(.CHV_REG)(dicTitulosRel("VL_FCP")) + CDbl(.VL_FCP)
                .VL_BC_ICMS_ST = dicDadosRel(.CHV_REG)(dicTitulosRel("VL_BC_ICMS_ST")) + CDbl(.VL_BC_ICMS_ST)
                .VL_ICMS_ST = dicDadosRel(.CHV_REG)(dicTitulosRel("VL_ICMS_ST")) + CDbl(.VL_ICMS_ST)
                .VL_RED_BC = dicDadosRel(.CHV_REG)(dicTitulosRel("VL_RED_BC")) + CDbl(.VL_RED_BC)
                .VL_IPI = dicDadosRel(.CHV_REG)(dicTitulosRel("VL_IPI")) + CDbl(.VL_IPI)
                .VL_ISENTAS = dicDadosRel(.CHV_REG)(dicTitulosRel("VL_ISENTAS")) + CDbl(.VL_ISENTAS)
                .VL_OUTRAS = dicDadosRel(.CHV_REG)(dicTitulosRel("VL_OUTRAS")) + CDbl(.VL_OUTRAS)
                
            End If
            
            dicDadosRel(.CHV_REG) = Array(.REG, "'" & .COD_MOD, .UF, CInt(.CFOP), .CST_ICMS, .ALIQ_ICMS, _
                                          CDbl(.VL_OPR), CDbl(.VL_BC_ICMS), CDbl(.VL_ICMS), CDbl(.ALIQ_FCP), _
                                          CDbl(.VL_FCP), CDbl(.VL_BC_ICMS_ST), CDbl(.VL_ICMS_ST), CDbl(.VL_RED_BC), _
                                          CDbl(.VL_IPI), CDbl(.VL_ISENTAS), CDbl(.VL_OUTRAS))
            
        End With
        
    Next Campos

End Function

Private Function CalcularIsentasOutras()
    
    With relatICMS
        
        .VL_ISENTAS = 0
        .VL_OUTRAS = 0
        
        Select Case True
            
            Case .CST_ICMS Like "*20", .CST_ICMS Like "*30", .CST_ICMS Like "*40,", .CST_ICMS Like "*41", .CST_ICMS Like "*70"
                .VL_ISENTAS = .VL_OPR - .VL_BC_ICMS - .VL_ICMS_ST - .VL_FCP - .VL_IPI
                If .VL_ISENTAS < 0 Then .VL_ISENTAS = 0
                
            Case Else
                .VL_OUTRAS = .VL_OPR - .VL_BC_ICMS - .VL_ICMS_ST - .VL_FCP - .VL_IPI - .VL_ISENTAS
                If .VL_OUTRAS < 0 Then .VL_OUTRAS = 0
                
        End Select
        
    End With
    
End Function
