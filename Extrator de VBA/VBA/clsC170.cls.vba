Attribute VB_Name = "clsC170"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Option Base 1

Public ImportDadosICMS As New clsC170_ImportadorDadosFiscal

Public Function GerarC175(Optional ByVal OmitirMsg As Boolean)

Dim VL_FRT As Double, VL_SEG#, VL_OUT_DA#, VL_DESP#, VL_MERC#, VL_ITEM#, VL_DESC#, VL_ADIC#, VL_ICMS#
Dim Campos As Variant, Campo, nCampo, Titulos, Chave
Dim ARQUIVO As String, CHV_REG$, COD_MOD$, COD_SIT$
Dim Dados As Range, Linha As Range
Dim dic0000 As New Dictionary
Dim dic0150 As New Dictionary
Dim dicC100 As New Dictionary
Dim dicC170 As New Dictionary
Dim dicC175 As New Dictionary
Dim dicTitulos As New Dictionary
Dim dicTitulosC175 As New Dictionary
Dim arrExcluir As New ArrayList
Dim Valores As Variant
Dim Inicio As Date
Dim i As Long
    
    If Not OmitirMsg Then Inicio = Now()
    
    If Util.ChecarAusenciaDados(regC170, OmitirMsg) Then Exit Function
    Call Util.AtualizarBarraStatus("Gerando registros C175, por favor aguarde...")
    
    If regC100.AutoFilterMode Then regC100.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC100, 3)
    Set Dados = Util.DefinirIntervalo(regC100, 4, 3)
    If Not Dados Is Nothing Then
        
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                COD_MOD = Campos(dicTitulos("COD_MOD"))
                If COD_MOD = "65" Then
                    
                    COD_SIT = Campos(dicTitulos("COD_SIT"))
                    Select Case VBA.Left(COD_SIT, 2)
                        
                        Case "00", "01", "06", "07", "08"
                        
                            CHV_REG = Campos(dicTitulos("CHV_REG"))
                            VL_MERC = Util.ValidarValores(Campos(dicTitulos("VL_MERC")))
                            VL_FRT = Util.ValidarValores(Campos(dicTitulos("VL_FRT")))
                            VL_SEG = Util.ValidarValores(Campos(dicTitulos("VL_SEG")))
                            VL_OUT_DA = Util.ValidarValores(Campos(dicTitulos("VL_OUT_DA")))
                            VL_DESP = VL_FRT + VL_SEG + VL_OUT_DA
                            
                            If Not dicC100.Exists(CHV_REG) Then Set dicC100(CHV_REG) = New Dictionary
                            
                            dicC100(CHV_REG)("VL_MERC") = VL_MERC
                            dicC100(CHV_REG)("VL_DESP") = VL_DESP
                            dicC100(CHV_REG)("COD_MOD") = COD_MOD
                            
                    End Select
                
                End If
                
            End If
            
        Next Linha

    End If
    
    If regC170.AutoFilterMode Then regC170.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC170, 3)
    Set dicTitulosC175 = Util.MapearTitulos(regC175_Contr, 3)
    Set Dados = Util.DefinirIntervalo(regC170, 4, 3)
    
    'Carrega dados do C170 e gera registro C175
    For Each Linha In Dados.Rows
        
        With CamposC175Contrib
        
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                .CHV_PAI = Campos(dicTitulos("CHV_PAI_FISCAL"))
                If Not dicC100.Exists(.CHV_PAI) Then GoTo Prx:
                    
                arrExcluir.Add .CHV_PAI
                
                VL_MERC = dicC100(.CHV_PAI)("VL_MERC")
                VL_DESP = dicC100(.CHV_PAI)("VL_DESP")
                ARQUIVO = Campos(dicTitulos("ARQUIVO"))
                VL_ITEM = Util.ValidarValores(Campos(dicTitulos("VL_ITEM")))
                VL_DESC = Util.ValidarValores(Campos(dicTitulos("VL_DESC")))
                VL_ICMS = Util.ValidarValores(Campos(dicTitulos("VL_ICMS")))
                
                If VL_MERC > 0 Then VL_ADIC = VBA.Round(VL_ITEM / VL_MERC * VL_DESP, 2)
                                                
                .REG = "C175"
                .CFOP = Campos(dicTitulos("CFOP"))
                .VL_OPER = VBA.Round(VL_ITEM + VL_ADIC, 2)
                .VL_DESC = VL_DESC + VL_ICMS
                .CST_PIS = Campos(dicTitulos("CST_PIS"))
                .VL_BC_PIS = Util.ValidarValores(Campos(dicTitulos("VL_BC_PIS")))
                .ALIQ_PIS = Util.ValidarValores(Campos(dicTitulos("ALIQ_PIS")))
                .QUANT_BC_PIS = Campos(dicTitulos("QUANT_BC_PIS"))
                .ALIQ_PIS_QUANT = Campos(dicTitulos("ALIQ_PIS_QUANT"))
                .VL_PIS = Util.ValidarValores(Campos(dicTitulos("VL_PIS")))
                .CST_COFINS = Campos(dicTitulos("CST_COFINS"))
                .VL_BC_COFINS = Util.ValidarValores(Campos(dicTitulos("VL_BC_COFINS")))
                .ALIQ_COFINS = Util.ValidarValores(Campos(dicTitulos("ALIQ_COFINS")))
                .QUANT_BC_COFINS = Campos(dicTitulos("QUANT_BC_COFINS"))
                .ALIQ_COFINS_QUANT = Campos(dicTitulos("ALIQ_COFINS_QUANT"))
                .VL_COFINS = Util.ValidarValores(Campos(dicTitulos("VL_COFINS")))
                .COD_CTA = Campos(dicTitulos("COD_CTA"))
                .INFO_COMPL = Campos(dicTitulos("OBS"))
                
                .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CFOP, .CST_PIS, .CST_COFINS, .ALIQ_PIS, .ALIQ_COFINS)
                
                Campos = Array(.REG, ARQUIVO, .CHV_REG, .CHV_PAI, .CHV_PAI, CInt(.CFOP), CDbl(.VL_OPER), CDbl(.VL_DESC), .CST_PIS, _
                    CDbl(.VL_BC_PIS), CDbl(.ALIQ_PIS), .QUANT_BC_PIS, .ALIQ_PIS_QUANT, CDbl(.VL_PIS), .CST_COFINS, CDbl(.VL_BC_COFINS), _
                    CDbl(.ALIQ_COFINS), .QUANT_BC_COFINS, .ALIQ_COFINS_QUANT, CDbl(.VL_COFINS), .COD_CTA, .INFO_COMPL)
                    
                If dicC175.Exists(.CHV_REG) Then
                    
                    'Soma valores valores do C175 para registros com a mesma chave
                    For Each Chave In dicTitulosC175.Keys()
                        
                        If Chave Like "VL_*" Then
                            Campos(dicTitulosC175(Chave)) = dicC175(.CHV_REG)(dicTitulosC175(Chave)) + CDbl(Campos(dicTitulosC175(Chave)))
                        End If
                        
                    Next Chave
                    
                End If
                
                dicC175(.CHV_REG) = Campos
                
            End If
            
        End With
Prx:
    Next Linha
    
    'Atualiza os valores dos Valores no registro C100
    If regC175_Contr.AutoFilterMode Then regC175_Contr.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC175_Contr, 3)
    
    Set Dados = Util.DefinirIntervalo(regC175_Contr, 4, 3)
    If Not Dados Is Nothing Then
        
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
            
                Chave = Campos(dicTitulos("CHV_PAI_FISCAL"))
                If Not arrExcluir.contains(Chave) Then
                    
                    Chave = Campos(dicTitulos("CHV_REG"))
                    dicC175(Chave) = Campos
                    
                End If
            
            End If
            
        Next Linha
        
    End If
    
    Call Util.LimparDados(regC175_Contr, 4, False)
    Call Util.ExportarDadosDicionario(regC175_Contr, dicC175)
    
    Call dicC100.RemoveAll
    Call dicC170.RemoveAll
    Call dicC175.RemoveAll
    
    Application.StatusBar = False
    If Not OmitirMsg Then Call Util.MsgInformativa("Registro C175 gerado/atualizado com sucesso!", "Geração do registro C175", Inicio)
    
End Function

Public Function GerarC190(Optional ByVal OmitirMsg As Boolean)

Dim VL_FRT As Double, VL_SEG#, VL_OUT_DA#, VL_DESP#, VL_MERC#, VL_ITEM#, VL_DESC#, VL_ADIC#
Dim Dados As Range, Linha As Range
Dim Campos As Variant, Campo, nCampo, Titulos, Chave
Dim dic0000 As New Dictionary
Dim dic0150 As New Dictionary
Dim dicC100 As New Dictionary
Dim dicC170 As New Dictionary
Dim dicC190 As New Dictionary
Dim dicTitulos As New Dictionary
Dim dicTitulosC190 As New Dictionary
Dim arrExcluir As New ArrayList
Dim Valores As Variant
Dim Inicio As Date
Dim i As Long
    
    If Not OmitirMsg Then Inicio = Now()
    
    If Util.ChecarAusenciaDados(regC170, OmitirMsg) Then Exit Function
    Call Util.AtualizarBarraStatus("Gerando registros C190, por favor aguarde...")
    
    If regC100.AutoFilterMode Then regC100.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC100, 3)
    Set Dados = Util.DefinirIntervalo(regC100, 4, 3)
    
    If Not Dados Is Nothing Then
        
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            Chave = Campos(dicTitulos("CHV_REG"))
            
            VL_MERC = Util.ValidarValores(Campos(dicTitulos("VL_MERC")))
            VL_FRT = Util.ValidarValores(Campos(dicTitulos("VL_FRT")))
            VL_SEG = Util.ValidarValores(Campos(dicTitulos("VL_SEG")))
            VL_OUT_DA = Util.ValidarValores(Campos(dicTitulos("VL_OUT_DA")))
            VL_DESP = VL_FRT + VL_SEG + VL_OUT_DA
            
            If Not dicC100.Exists(Chave) Then Set dicC100(Chave) = New Dictionary
            
            dicC100(Chave)("VL_MERC") = VL_MERC
            dicC100(Chave)("VL_DESP") = VL_DESP
            
        Next Linha
        
    End If
    
    If regC170.AutoFilterMode Then regC170.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC170, 3)
    Set dicTitulosC190 = Util.MapearTitulos(regC190, 3)
    Set Dados = Util.DefinirIntervalo(regC170, 4, 3)
    
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        With CamposC190
            
            .CHV_PAI = Campos(dicTitulos("CHV_PAI_FISCAL"))
            If Not dicC100.Exists(.CHV_PAI) Then GoTo Prx:
            
            VL_MERC = dicC100(.CHV_PAI)("VL_MERC")
            VL_DESP = dicC100(.CHV_PAI)("VL_DESP")
            ARQUIVO = Campos(dicTitulos("ARQUIVO"))
            VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM")), True, 2)
            VL_DESC = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESC")), True, 2)
            If VL_MERC > 0 Then VL_ADIC = fnExcel.ConverterValores(VL_ITEM / VL_MERC * VL_DESP, True, 2)
            
            arrExcluir.Add .CHV_PAI
            
            .REG = "C190"
            .CFOP = Campos(dicTitulos("CFOP"))
            .CST_ICMS = Campos(dicTitulos("CST_ICMS"))
            .ALIQ_ICMS = Campos(dicTitulos("ALIQ_ICMS"))
            .VL_BC_ICMS = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_ICMS")), True, 2)
            .VL_ICMS = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS")), True, 2)
            .VL_BC_ICMS_ST = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_ICMS_ST")), True, 2)
            .VL_ICMS_ST = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS_ST")), True, 2)
            .VL_IPI = fnExcel.ConverterValores(Campos(dicTitulos("VL_IPI")), True, 2)
            .COD_OBS = ""
            
            .VL_OPR = fnExcel.ConverterValores(VL_ITEM + VL_ADIC + .VL_IPI + .VL_ICMS_ST - VL_DESC, True, 2)
            If .ALIQ_ICMS <> "" Then .ALIQ_ICMS = fnExcel.ConverterValores(.ALIQ_ICMS)
            
            If .CST_ICMS Like "*20" Or .CST_ICMS Like "*70" Then
                
                .VL_RED_BC = .VL_OPR - .VL_BC_ICMS - .VL_ICMS_ST - .VL_IPI
                If .VL_RED_BC < 0 Then .VL_RED_BC = 0
                
            Else
                
                .VL_RED_BC = 0
                
            End If
            
            If .VL_OPR < 0 Then .VL_OPR = 0
            
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CFOP, .CST_ICMS, .ALIQ_ICMS)
            
            Campos = Array(.REG, ARQUIVO, .CHV_REG, .CHV_PAI, "", "'" & .CST_ICMS, CInt(.CFOP), .ALIQ_ICMS, _
                           CDbl(.VL_OPR), CDbl(.VL_BC_ICMS), CDbl(.VL_ICMS), CDbl(.VL_BC_ICMS_ST), _
                           CDbl(.VL_ICMS_ST), CDbl(.VL_RED_BC), CDbl(.VL_IPI), .COD_OBS)
            
            If .ALIQ_ICMS <> "" Then Campos(dicTitulosC190("ALIQ_ICMS")) = CDbl(.ALIQ_ICMS)
            If dicC190.Exists(.CHV_REG) Then
                
                'Soma valores valores do C190 para registros com a mesma chave
                For Each Chave In dicTitulosC190.Keys()
                    
                    If Chave Like "VL_*" Then
                        Campos(dicTitulosC190(Chave)) = dicC190(.CHV_REG)(dicTitulosC190(Chave)) + CDbl(Campos(dicTitulosC190(Chave)))
                    End If
                    
                Next Chave
                
            End If
            
            If Util.ChecarCamposPreenchidos(Campos) Then dicC190(.CHV_REG) = Campos
            
        End With
Prx:
    Next Linha
    
    'Atualiza os valores dos Valores no registro C100
    If regC190.AutoFilterMode Then regC190.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC190, 3)
    Set Dados = Util.DefinirIntervalo(regC190, 4, 3)
    
    If Not Dados Is Nothing Then
        
        For Each Linha In Dados.Rows
            Campos = Application.index(Linha.Value2, 0, 0)
            
            Chave = Campos(dicTitulos("CHV_PAI_FISCAL"))
            If Not arrExcluir.contains(Chave) Then
                
                Campos(dicTitulos("CST_ICMS")) = Util.FormatarTexto(Campos(dicTitulos("CST_ICMS")))
                Chave = Campos(dicTitulos("CHV_REG"))
                If Util.ChecarCamposPreenchidos(Campos) Then dicC190(Chave) = Campos
                
            End If
            
        Next Linha
        
    End If
    
    Call Util.LimparDados(regC190, 4, False)
    Call Util.ExportarDadosDicionario(regC190, dicC190)
    
    Call dicC100.RemoveAll
    Call dicC170.RemoveAll
    Call dicC190.RemoveAll
    
    Application.StatusBar = False
    If Not OmitirMsg Then Call Util.MsgInformativa("Registro C190 gerado/atualizado com sucesso!", "Geração do registro C190", Inicio)
    
End Function

Public Function CalcularPISCOFINS(Optional ByVal ExcluirICMS As Boolean, Optional ByVal ExcluirICMS_ST As Boolean)

Dim VL_FRT As Double, VL_SEG#, VL_OUT_DA#, VL_DESP#, VL_MERC#, VL_ITEM#, VL_DESC#, VL_ADIC#, VL_ICMS#, VL_ICMS_ST#, ALIQ_PIS#, ALIQ_COFINS#
Dim CST_PIS As String
Dim Dados As Range, Linha As Range
Dim Campos As Variant, Campo, nCampo, Titulos
Dim dicC100 As New Dictionary
Dim dicC170 As New Dictionary
Dim dicTitulos As New Dictionary
Dim Chave As String
Dim Valores As Variant
Dim Inicio As Date
Dim i As Long
    
    Inicio = Now()
    
    If Util.ChecarAusenciaDados(regC100, False) Then Exit Function
    Call Util.AtualizarBarraStatus("Calculando PIS e COFINS, por favor aguarde...")
    
    If regC100.AutoFilterMode Then regC100.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC100, 3)
    Set Dados = Util.DefinirIntervalo(regC100, 4, 3)
    
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
        
            Chave = Campos(dicTitulos("CHV_REG"))
            
            VL_MERC = fnExcel.ConverterValores(Campos(dicTitulos("VL_MERC")))
            VL_FRT = fnExcel.ConverterValores(Campos(dicTitulos("VL_FRT")))
            VL_SEG = fnExcel.ConverterValores(Campos(dicTitulos("VL_SEG")))
            VL_OUT_DA = fnExcel.ConverterValores(Campos(dicTitulos("VL_OUT_DA")))
            VL_DESP = VL_FRT + VL_SEG + VL_OUT_DA
            
            If Not dicC100.Exists(Chave) Then Set dicC100(Chave) = New Dictionary
            
            dicC100(Chave)("VL_MERC") = VL_MERC
            dicC100(Chave)("VL_DESP") = VL_DESP
                        
        End If
        
    Next Linha
            
    If regC170.AutoFilterMode Then regC170.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC170, 3)
    Set Dados = Util.DefinirIntervalo(regC170, 4, 3)
    
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não existem dados para processar no registro C170", "Registro C170 sem dados informados")
        Exit Function
    End If
    
    'Carrega dados do C170 e recalcula o PIS e o COFINS
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            With CamposC170
                
                .CHV_PAI = Campos(dicTitulos("CHV_PAI_FISCAL"))
                .CHV_REG = Campos(dicTitulos("CHV_REG"))
                ARQUIVO = Campos(dicTitulos("ARQUIVO"))
                CST_PIS = Util.ApenasNumeros(Campos(dicTitulos("CST_PIS")))
                VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM")))
                VL_DESC = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESC")))
                VL_ICMS = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS")))
                VL_ICMS_ST = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS_ST")))
                VL_MERC = dicC100(.CHV_PAI)("VL_MERC")
                VL_DESP = dicC100(.CHV_PAI)("VL_DESP")
                If VL_MERC > 0 Then VL_ADIC = VBA.Round((VL_ITEM / VL_MERC) * VL_DESP, 2)
                
                If ExcluirICMS Then VL_ITEM = VL_ITEM - VL_ICMS
                If ExcluirICMS_ST Then VL_ITEM = VL_ITEM - VL_ICMS - VL_ICMS_ST
                
                Select Case True
                    
                    Case CST_PIS Like "01*", CST_PIS Like "02*", CST_PIS Like "50*"
                        ALIQ_PIS = fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_PIS")))
                        Campos(dicTitulos("VL_BC_PIS")) = VBA.Round(VL_ITEM + VL_ADIC - VL_DESC, 2)
                        Campos(dicTitulos("VL_PIS")) = VBA.Round(Campos(dicTitulos("VL_BC_PIS")) * fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_PIS"))), 2)
                        
                        If ALIQ_PIS = 0.0065 Then
                            
                            Campos(dicTitulos("ALIQ_COFINS")) = 0.03
                            
                        ElseIf ALIQ_PIS = 0.0165 Then
                            
                            Campos(dicTitulos("ALIQ_COFINS")) = 0.076
                            
                        End If
                        
                        Campos(dicTitulos("VL_BC_COFINS")) = VBA.Round(VL_ITEM + VL_ADIC - VL_DESC, 2)
                        Campos(dicTitulos("VL_COFINS")) = VBA.Round(Campos(dicTitulos("VL_BC_COFINS")) * fnExcel.FormatarPercentuais(Campos(dicTitulos("ALIQ_COFINS"))), 2)
                        
                    Case Else
                        Campos(dicTitulos("VL_BC_PIS")) = 0
                        Campos(dicTitulos("ALIQ_PIS")) = 0
                        Campos(dicTitulos("VL_PIS")) = 0
                        
                        Campos(dicTitulos("VL_BC_COFINS")) = 0
                        Campos(dicTitulos("ALIQ_COFINS")) = 0
                        Campos(dicTitulos("VL_COFINS")) = 0
                        
                End Select
                
                dicC170(.CHV_REG) = Campos
                
            End With
            
        End If
        
    Next Linha
    
    If dicC170.Count > 0 Then
        
        Call Util.LimparDados(regC170, 4, False)
        Call Util.ExportarDadosDicionario(regC170, dicC170)
        
        Call dicC100.RemoveAll
        Call dicC170.RemoveAll
        
        Call Util.MsgInformativa("Valores de PIS e COFINS recalculados com sucesso!", "Cálculo do PIS e COFINS", Inicio)
        
    End If
    
    Application.StatusBar = False
    
End Function

Public Function AtualizarImpostosC100(Optional ByVal OmitirMsg As Boolean)
    
Dim Dados As Range, Linha As Range
Dim Campos As Variant, Campo, nCampo, Titulos
Dim dicC100 As New Dictionary
Dim dicC170 As New Dictionary
Dim dicTitulos As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim Chave As Variant
Dim Valores As Variant
Dim Inicio As Date
    
    If Not OmitirMsg Then Inicio = Now()
    Call Util.AtualizarBarraStatus("Atualizando os valores dos Valores no registro C100, por favor aguarde...")
    
    Valores = Array("VL_ITEM", "VL_DESC", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "VL_ICMS_ST", "VL_IPI", "VL_PIS", "VL_COFINS", "VL_ABAT_NT")
    If regC170.AutoFilterMode Then regC170.AutoFilter.ShowAllData
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    Set Dados = Util.DefinirIntervalo(regC170, 4, 3)
    If Dados Is Nothing Then Exit Function
    
    With CamposC170
        
        'Carrega dados do C170 e soma os valores dos campos selecionados
        For Each Linha In Dados.Rows
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
            
                .CHV_PAI = Campos(dicTitulosC170("CHV_PAI_FISCAL"))
                
                If dicC170.Exists(.CHV_PAI) Then
                
                    'Soma valores valores do C170 para registros com a mesma chave
                    For Each Chave In dicTitulosC170.Keys()
                        
                        If Chave Like "VL_*" Then
                            If Campos(dicTitulosC170(Chave)) = "" Then Campos(dicTitulosC170(Chave)) = 0
                            Campos(dicTitulosC170(Chave)) = dicC170(.CHV_PAI)(dicTitulosC170(Chave)) + CDbl(Campos(dicTitulosC170(Chave)))
                        End If
                        
                    Next Chave
                                        
                End If
                
                dicC170(.CHV_PAI) = Campos
            
            End If
            
        Next Linha
    
    End With
    
    'Atualiza os valores dos Valores no registro C100
    If regC100.AutoFilterMode Then regC100.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC100, 3)
    Set Dados = Util.DefinirIntervalo(regC100, 4, 3)
    
    For Each Linha In Dados.Rows
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            Chave = Campos(dicTitulos("CHV_REG"))
            If dicC170.Exists(Chave) Then
                
                For Each nCampo In Valores
                    If nCampo = "VL_ITEM" Then
                        Campos(dicTitulos("VL_MERC")) = CDbl(dicC170(Chave)(dicTitulosC170("VL_ITEM")))
                    Else
                        Campos(dicTitulos(nCampo)) = CDbl(dicC170(Chave)(dicTitulosC170(nCampo)))
                    End If
                    
                Next nCampo
            
            End If
            
            dicC100(Chave) = Campos
        
        End If
        
    Next Linha
    
    Call Util.LimparDados(regC100, 4, False)
    Call Util.ExportarDadosDicionario(regC100, dicC100)
    
    Call dicC170.RemoveAll
    Call dicC100.RemoveAll
    
    Application.StatusBar = False
    If Not OmitirMsg Then Call Util.MsgInformativa("Valores atualizados com sucesso!", "Atualização de Valores do C100", Inicio)
              
End Function

Public Function CalcularAjustesDecretoAtacadistaBA(ByVal Registro As String, ByRef arrProdutosExcluidos As ArrayList, _
                                                   ByRef dicAjustes As Dictionary, ByRef dicE113 As Dictionary, _
                                                   ByVal cPart As String, ByVal Modelo As String, ByVal SERIE As String, _
                                                   ByVal nNF As String, ByVal Emissao As String, ByVal chNFe As String)

Dim CFOP As Integer
Dim Campos As Variant
Dim Chave As String, CST$, cProd$
Dim bcICMS As Double, vICMS#, pCred#, vAjuste#
    
    Campos = fnSPED.ExtrairCamposRegistro(Registro)
    Emissao = VBA.Format(VBA.Format(Emissao, "00/00/0000"), "ddmmyyyy")
    cProd = Campos(3)
    bcICMS = Util.ValidarValores(Campos(7)) - Util.ValidarValores(Campos(8))
    CFOP = Campos(11)
    vICMS = Util.ValidarValores(Campos(15))
    pCred = CalcularPerncentualCredito(bcICMS, vICMS)
    
    If Not arrProdutosExcluidos.contains(cProd) Then
        
        Select Case True
            
            Case (VBA.Left(CFOP, 2) = "11" Or VBA.Left(CFOP, 2) = "19" Or VBA.Left(CFOP, 2) = "21" Or VBA.Left(CFOP, 2) = "29") And (pCred > 0.1)
                vAjuste = FuncoesSEFAZ_BA.CalcularEstornoCreditoDecretoAtacadistaBA(bcICMS, vICMS)
                Call CriarAtualizarAjustes(dicAjustes, "BA010005", vAjuste)
                Call rE113.CriarAtualizarE113(dicE113, "BA010005", vAjuste, cPart, Modelo, SERIE, nNF, Emissao, cProd, chNFe)
                
            Case (CFOP > 6000) And (pCred >= 0.12)
                vAjuste = FuncoesSEFAZ_BA.CalcularCreditoPresumidoInsterestadualDecretoAtacadistaBA(bcICMS, vICMS)
                Call CriarAtualizarAjustes(dicAjustes, "BA020010", vAjuste)
                Call rE113.CriarAtualizarE113(dicE113, "BA020010", vAjuste, cPart, Modelo, SERIE, nNF, Emissao, cProd, chNFe)
                
        End Select
        
    End If
    
End Function

Public Function CalcularCreditoPresumidoArt269Inc10BA(ByVal Registro As String, ByRef dicFornecSN As Dictionary, _
                                                      ByRef dicAjustes As Dictionary, ByRef dicE113 As Dictionary, _
                                                      ByVal cPart As String, ByVal Modelo As String, ByVal SERIE As String, _
                                                      ByVal nNF As String, ByVal Emissao As String, ByVal chNFe As String) As String

Dim CFOP As Integer
Dim Campos As Variant
Dim Chave As String, CST$, cProd$, NITEM$
Dim vItem As Double, vICMS#, pCredSN#, vAjuste#
    
    Campos = fnSPED.ExtrairCamposRegistro(Registro)
    NITEM = Campos(2)
    Emissao = VBA.Format(VBA.Format(Emissao, "00/00/0000"), "ddmmyyyy")
    cProd = Campos(3)
    vItem = Util.ValidarValores(Campos(7)) - Util.ValidarValores(Campos(8))
    CFOP = Campos(11)
    vICMS = Util.ValidarValores(Campos(15))
    
    Chave = chNFe & NITEM
    Select Case True
        
        Case (CFOP < 2000) And (dicFornecSN.Exists(Chave))
            If (dicFornecSN(Chave)(0) = "5101") And (VBA.Left(CFOP, 2) = "11" Or VBA.Left(CFOP, 2) = "21" Or VBA.Right(CFOP, 3) = "910") Then
                vAjuste = FuncoesSEFAZ_BA.CalcularCreditoPresumidoArt269Inc10BA(vItem)
                Call CriarAtualizarAjustes(dicAjustes, "BA020009", vAjuste)
                Call rE113.CriarAtualizarE113(dicE113, "BA020009", vAjuste, cPart, Modelo, SERIE, nNF, Emissao, cProd, chNFe)
                
            End If
            
            Campos(13) = 0
            Campos(14) = 0
            Campos(15) = 0
            
    End Select
    
    CalcularCreditoPresumidoArt269Inc10BA = Join(Campos, "|")
    
End Function

Public Function CalcularCreditoAquisicaoSimplesNacionalBA(ByVal Registro As String, ByRef dicFornecSN As Dictionary, _
                                                          ByRef dicAjustes As Dictionary, ByRef dicE113 As Dictionary, _
                                                          ByVal cPart As String, ByVal Modelo As String, ByVal SERIE As String, _
                                                          ByVal nNF As String, ByVal Emissao As String, ByVal chNFe As String) As String

Dim CFOP As Integer
Dim Campos As Variant
Dim Chave As String, CST$, cProd$, NITEM$
Dim vItem As Double, vICMS#, pCredSN#, vAjuste#

    Campos = fnSPED.ExtrairCamposRegistro(Registro)
    NITEM = Campos(2)
    Emissao = VBA.Format(VBA.Format(Emissao, "00/00/0000"), "ddmmyyyy")
    cProd = Campos(3)
    vItem = Util.ValidarValores(Campos(7)) - Util.ValidarValores(Campos(8))
    CFOP = Campos(11)
    vICMS = Util.ValidarValores(Campos(15))
    
    Chave = chNFe & NITEM
    Select Case True
            
        Case (CFOP < 3000) And (dicFornecSN.Exists(Chave))
        
            If (dicFornecSN(Chave)(0) <> "5101") Then
            
                pCredSN = dicFornecSN(Chave)(1)
                If (Mid(dicFornecSN(Chave)(0), 2, 1) = "1") And (pCredSN > 0) _
                And (VBA.Left(CFOP, 2) = "11" Or VBA.Left(CFOP, 2) = "21" Or VBA.Right(CFOP, 3) = "910") Then
                            
                    vAjuste = FuncoesSEFAZ_BA.CalcularCreditoAquisicaoSimplesNacionalBA(vItem, pCredSN)
                    Call CriarAtualizarAjustes(dicAjustes, "BA020008", vAjuste)
                    Call rE113.CriarAtualizarE113(dicE113, "BA020008", vAjuste, cPart, Modelo, SERIE, nNF, Emissao, cProd, chNFe)
            
                End If
            
                Campos(13) = 0
                Campos(14) = 0
                Campos(15) = 0
            
            End If
            
    End Select

    CalcularCreditoAquisicaoSimplesNacionalBA = Join(Campos, "|")
    
End Function

Public Function SomarIPIeSTaosItens(ByVal tipo As String, Optional OmitirMsg As Boolean)
    
Dim VL_ITEM#, VL_ICMS_ST#, VL_IPI#
Dim Dados As Range, Linha As Range
Dim Campos As Variant
Dim dicC170 As New Dictionary
Dim dicTitulos As New Dictionary
Dim CHV_REG As String
Dim Inicio As Date
Dim Msg As String
    
    Inicio = Now()
    If Util.ChecarAusenciaDados(regC170, OmitirMsg) Then Exit Function
    
    If tipo = "IPI" Then Msg = "Incluido valores do IPI ao valor dos itens, por favor aguarde..."
    If tipo = "ST" Then Msg = "Incluido valores do ICMS-ST ao valor dos itens, por favor aguarde..."
    If tipo = "IPI-ST" Then Msg = "Incluido valores do IPI e ICMS-ST ao valor dos itens, por favor aguarde..."
    
    Application.StatusBar = Msg
    
    If regC170.AutoFilterMode Then regC170.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC170, 3)
    Set Dados = Util.DefinirIntervalo(regC170, 4, 3)
    
    'Carrega dados do C170 e soma os valores dos campos selecionados
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            Select Case tipo
                
                Case "IPI-ST"
                    Msg = "Inclusão do IPI e ICMS-ST ao valor dos itens do C170"
                    Call ZerarValoresIPI_ST(Campos, dicTitulos)
                    
                Case "ST"
                    Msg = "Inclusão ICMS-ST ao valor dos itens do C170"
                    Call ZerarValoresST(Campos, dicTitulos)
                    
                Case "IPI"
                    Msg = "Inclusão do IPI ao valor dos itens do C170"
                    Call ZerarValoresIPI(Campos, dicTitulos)
                    
            End Select
            
            CHV_REG = Campos(dicTitulos("CHV_REG"))
            dicC170(CHV_REG) = Campos
            
        End If
        
    Next Linha
    
    'Atualiza os dados do registro C170
    Call Util.LimparDados(regC170, 4, False)
    Call Util.ExportarDadosDicionario(regC170, dicC170)
    
    Call dicC170.RemoveAll
    
    Application.StatusBar = False
    
    If OmitirMsg Then Exit Function
    Call Util.MsgInformativa("Valor dos Itens atualizados com sucesso!", Msg, Inicio)
    
End Function

Private Sub ZerarValoresIPI(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim VL_ITEM As Double, VL_IPI As Double
    
    VL_ITEM = Campos(dicTitulos("VL_ITEM"))
    VL_IPI = Campos(dicTitulos("VL_IPI"))
    
    Campos(dicTitulos("VL_ITEM")) = VL_ITEM + VL_IPI
    Campos(dicTitulos("VL_BC_IPI")) = 0
    Campos(dicTitulos("ALIQ_IPI")) = 0
    Campos(dicTitulos("VL_IPI")) = 0
    Campos(dicTitulos("CST_IPI")) = ""
    Campos(dicTitulos("COD_ENQ")) = ""
    Campos(dicTitulos("IND_APUR")) = ""
    
End Sub

Private Sub ZerarValoresST(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim VL_ITEM As Double, VL_ICMS_ST As Double
    
    VL_ITEM = Campos(dicTitulos("VL_ITEM"))
    VL_ICMS_ST = Campos(dicTitulos("VL_ICMS_ST"))
    
    Campos(dicTitulos("VL_ITEM")) = VL_ITEM + VL_ICMS_ST
    Campos(dicTitulos("VL_BC_ICMS_ST")) = 0
    Campos(dicTitulos("VL_ICMS_ST")) = 0
    Campos(dicTitulos("ALIQ_ST")) = 0
    
End Sub

Private Sub ZerarValoresIPI_ST(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)
    
Dim VL_ITEM As Double, VL_IPI As Double, VL_ICMS_ST As Double
    
    VL_IPI = Campos(dicTitulos("VL_IPI"))
    VL_ITEM = Campos(dicTitulos("VL_ITEM"))
    VL_ICMS_ST = Campos(dicTitulos("VL_ICMS_ST"))
    
    Campos(dicTitulos("VL_ITEM")) = VL_ITEM + VL_ICMS_ST + VL_IPI
    
    'Campos IPI
    Campos(dicTitulos("VL_BC_IPI")) = 0
    Campos(dicTitulos("ALIQ_IPI")) = 0
    Campos(dicTitulos("VL_IPI")) = 0
    Campos(dicTitulos("CST_IPI")) = ""
    Campos(dicTitulos("COD_ENQ")) = ""
    Campos(dicTitulos("IND_APUR")) = ""
    
    'Campos ST
    Campos(dicTitulos("VL_BC_ICMS_ST")) = 0
    Campos(dicTitulos("VL_ICMS_ST")) = 0
    Campos(dicTitulos("ALIQ_ST")) = 0
        
End Sub

Public Function AgruparRegistros(Optional OmitirMsg As Boolean = False)

Dim Campos As Variant, Campo, nCampo, Titulos, Valores
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim dicC170 As New Dictionary
Dim Chave As String
Dim Inicio As Date
    
    If Not OmitirMsg Then Inicio = Now()
    
    If Util.ChecarAusenciaDados(regC170, OmitirMsg) Then Exit Function
    Call Util.AtualizarBarraStatus("Agrupando dados do registro C170, por favor aguarde...")
    
    Set dicTitulos = Util.MapearTitulos(regC170, 3)
    Titulos = Array("QTD", "VL_ITEM", "VL_DESC", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "VL_ICMS_ST", "VL_BC_IPI", "VL_IPI", "VL_BC_PIS", "QUANT_BC_PIS", "VL_PIS", "VL_BC_COFINS", "QUANT_BC_COFINS", "VL_COFINS")

    If regC170.AutoFilterMode Then regC170.AutoFilter.ShowAllData
    Call fnExcel.ClassificarColuna(regC170, dicTitulos, 3, False, "CHV_PAI_FISCAL", "NUM_ITEM")
    
    Set Dados = Util.DefinirIntervalo(regC170, 4, 3)
    
    With CamposC170
        
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                .CHV_PAI = Campos(dicTitulos("CHV_PAI_FISCAL"))
                .COD_ITEM = Campos(dicTitulos("COD_ITEM"))
                .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .COD_ITEM)
                
                Call Util.SomarValoresCamposSelecionados(dicC170, dicTitulos, Titulos, Campos, .CHV_REG)
                
            End If
            
        Next Linha
        
    End With
    
    'Reorganiza a numeração dos itens no C170
    Call ReEnumerarItens(dicC170, dicTitulos)
    
    'Atualiza os dados do registro C170
    Call Util.LimparDados(regC170, 4, False)
    Call Util.ExportarDadosDicionario(regC170, dicC170)
    
    Call dicC170.RemoveAll
    
    Application.StatusBar = False
    If Not OmitirMsg Then Call Util.MsgInformativa("registros agrupados com sucesso!", "Agrupamento de registros C170", Inicio)
    
End Function

Public Function UnificarProdutosDuplicadosEmNotasSelecionadas(ByVal arrChaves As ArrayList) As Boolean
    
Dim Campos As Variant, Campo, nCampo, Titulos, Valores
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim dicC170 As New Dictionary
Dim arrDados As New ArrayList
Dim Chave As String
Dim Inicio As Date
    
    Call Util.AtualizarBarraStatus("Unificando produtos duplicados no registro C170, por favor aguarde...")
    
    Set dicTitulos = Util.MapearTitulos(regC170, 3)
    Titulos = Array("QTD", "VL_ITEM", "VL_DESC", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "VL_ICMS_ST", "VL_BC_IPI", "VL_IPI", "VL_BC_PIS", "QUANT_BC_PIS", "VL_PIS", "VL_BC_COFINS", "QUANT_BC_COFINS", "VL_COFINS")

    If regC170.AutoFilterMode Then regC170.AutoFilter.ShowAllData
    Call fnExcel.ClassificarColuna(regC170, dicTitulos, 3, False, "CHV_PAI_FISCAL", "NUM_ITEM")
    
    Set arrDados = Util.CriarArrayListRegistro(regC170)
    If arrDados.Count = 0 Then Exit Function
    
    With CamposC170
        
        For Each Campos In arrDados
            
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                .CHV_PAI = Campos(dicTitulos("CHV_PAI_FISCAL"))
                .CHV_REG = Campos(dicTitulos("CHV_REG"))
                
                If arrChaves.contains(.CHV_PAI) Then
                    
                    .COD_ITEM = Campos(dicTitulos("COD_ITEM"))
                    .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .COD_ITEM)
                    
                End If
                
                Call Util.SomarValoresCamposSelecionados(dicC170, dicTitulos, Titulos, Campos, .CHV_REG)
                
            End If
            
        Next Campos
        
    End With
    
    Call ReEnumerarItens(dicC170, dicTitulos)
    
    Call Util.LimparDados(regC170, 4, False)
    Call Util.ExportarDadosDicionario(regC170, dicC170)
    
    If dicC170.Count < arrDados.Count Then UnificarProdutosDuplicadosEmNotasSelecionadas = True
    Call dicC170.RemoveAll
    
    Application.StatusBar = False
    
End Function

Private Function ReEnumerarItens(ByRef dicC170 As Dictionary, ByRef dicTitulos As Dictionary)

Dim Chave As Variant, Campos
Dim NUM_ITEM As String
Dim NovoNumItem As Integer
Dim CHV_PAI As String, CHV_PAI_Anterior$
    
    'Inicializa as variáveis
    NovoNumItem = 1
    CHV_PAI_Anterior = ""
    
    For Each Chave In dicC170.Keys()
        
        Campos = dicC170(Chave)
        
        CHV_PAI = Campos(dicTitulos("CHV_PAI_FISCAL"))
        
        'Verifica se mudou a nota fiscal e reinicia a contagem
        If CHV_PAI <> CHV_PAI_Anterior Then
            NovoNumItem = 1
            CHV_PAI_Anterior = CHV_PAI
        End If
                        
        Campos(dicTitulos("NUM_ITEM")) = NovoNumItem
                
        dicC170(Chave) = Campos
        
        NovoNumItem = NovoNumItem + 1
        
    Next Chave
    
End Function
