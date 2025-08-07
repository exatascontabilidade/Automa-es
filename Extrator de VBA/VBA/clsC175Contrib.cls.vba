Attribute VB_Name = "clsC175Contrib"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function AtualizarImpostosC100(Optional ByVal OmitirMsg As Boolean)

Dim Valores As Variant, CamposC100, Chave, Campos
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC175 As New Dictionary
Dim dicAcumulador As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim Dados As Range, Linha As Range
Dim CHV_C100 As String, CST_PIS$
    
    If Not OmitirMsg Then Inicio = Now()
    
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    
    Set dicTitulosC175 = Util.MapearTitulos(regC175_Contr, 3)
    
    Application.StatusBar = "Atualizando os valores dos Valores no registro C100, por favor aguarde..."
    
    If regC175_Contr.AutoFilterMode Then regC175_Contr.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(regC175_Contr, 4, 3)
    If Dados Is Nothing Then Exit Function
    
    With CamposC175
        
        'Carrega dados do C175 e soma os valores dos campos selecionados
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                CHV_C100 = Campos(dicTitulosC175("CHV_PAI_FISCAL"))
                If Not dicAcumulador.Exists(CHV_C100) Then
                    dicAcumulador.Add CHV_C100, Array(0, 0, 0, 0, 0, 0)
                End If
                
                'Carrega valores acumulados
                Valores = dicAcumulador(CHV_C100)
                
                CST_PIS = Campos(dicTitulosC175("CST_PIS"))
                If CST_PIS Like "05*" Or CST_PIS Like "75*" Then
                    
                    Valores(4) = Valores(4) + CDbl(fnExcel.ConverterValores(Campos(dicTitulosC175("VL_PIS"))))
                    Valores(5) = Valores(5) + CDbl(fnExcel.ConverterValores(Campos(dicTitulosC175("VL_COFINS"))))
                    
                Else
                    
                    Valores(2) = Valores(2) + CDbl(fnExcel.ConverterValores(Campos(dicTitulosC175("VL_PIS"))))
                    Valores(3) = Valores(3) + CDbl(fnExcel.ConverterValores(Campos(dicTitulosC175("VL_COFINS"))))
                    
                End If
                
                Valores(0) = Valores(0) + CDbl(fnExcel.ConverterValores(Campos(dicTitulosC175("VL_OPER"))))
                Valores(1) = Valores(1) + CDbl(fnExcel.ConverterValores(Campos(dicTitulosC175("VL_DESC"))))
                
            End If
                        
            'Atualiza o acumulador
            dicAcumulador(CHV_C100) = Valores
            
        Next Linha
        
        For Each Chave In dicDadosC100.Keys()
        
            If dicDadosC100.Exists(Chave) Then
                
                CamposC100 = dicDadosC100(Chave)
                
                If dicAcumulador.Exists(Chave) Then
                    
                    Valores = dicAcumulador(Chave)
                    
                    CamposC100(dicTitulosC100("VL_MERC")) = Valores(0)
                    'CamposC100(dicTitulosC100("VL_DESC")) = Valores(1)
                    CamposC100(dicTitulosC100("VL_PIS")) = Valores(2)
                    CamposC100(dicTitulosC100("VL_COFINS")) = Valores(3)
                    CamposC100(dicTitulosC100("VL_PIS_ST")) = Valores(4)
                    CamposC100(dicTitulosC100("VL_COFINS_ST")) = Valores(5)
                
                    'Atualizando o dicionário com os novos valores
                    dicDadosC100(Chave) = CamposC100
                
                End If
            
            End If
            
        Next Chave
        
    End With
    
    Call Util.LimparDados(regC100, 4, False)
    Call Util.ExportarDadosDicionario(regC100, dicDadosC100)
    
    Application.StatusBar = False
    If Not OmitirMsg Then Call Util.MsgInformativa("Valores atualizados com sucesso!", "Atualização de Valores do C100", Inicio)
    
End Function

Public Function AgruparRegistros(Optional ByVal OmitirMsg As Boolean)
    
Dim Campos As Variant, Campo, nCampo, Titulos
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim dicC175 As New Dictionary
Dim Chave As Variant, Valores
Dim CHV_REG As String
Dim Inicio As Date
    
    If Not OmitirMsg Then Inicio = Now()
    
    If Util.ChecarAusenciaDados(regC175_Contr, OmitirMsg) Then Exit Function
    Call Util.AtualizarBarraStatus("Agrupando dados do registro C175, por favor aguarde...")
    
    Valores = Array("VL_OPER", "VL_DESC", "VL_BC_PIS", "VL_PIS", "VL_BC_COFINS", "VL_COFINS")
    
    If regC175_Contr.AutoFilterMode Then regC175_Contr.AutoFilter.ShowAllData
    Set dicTitulos = Util.MapearTitulos(regC175_Contr, 3)
    Set Dados = Util.DefinirIntervalo(regC175_Contr, 4, 3)
    
    With CamposC175Contrib
        
        'Carrega dados do C175 e soma os valores dos campos selecionados
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                .CHV_REG = Campos(dicTitulos("CHV_REG"))
                .CHV_PAI = Campos(dicTitulos("CHV_PAI_FISCAL"))
                .CFOP = Campos(dicTitulos("CFOP"))
                .CST_PIS = Campos(dicTitulos("CST_PIS"))
                .CST_COFINS = Campos(dicTitulos("CST_COFINS"))
                .ALIQ_PIS = Campos(dicTitulos("ALIQ_PIS"))
                .ALIQ_COFINS = Campos(dicTitulos("ALIQ_COFINS"))
                .ALIQ_PIS_QUANT = Campos(dicTitulos("ALIQ_PIS_QUANT"))
                .ALIQ_COFINS_QUANT = Campos(dicTitulos("ALIQ_COFINS_QUANT"))

                .COD_CTA = Campos(dicTitulos("COD_CTA"))
                
                CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, .CFOP, .CST_PIS, .ALIQ_PIS, .ALIQ_PIS_QUANT, .CST_COFINS, .ALIQ_COFINS, .ALIQ_COFINS_QUANT, .COD_CTA)
                If dicC175.Exists(CHV_REG) Then
                    
                    'Soma valores valores do C175 para registros com a mesma chave
                    For Each Chave In dicTitulos.Keys()
                        
                        If Chave Like "VL_*" Then
                            Campos(dicTitulos(Chave)) = dicC175(CHV_REG)(dicTitulos(Chave)) + CDbl(Campos(dicTitulos(Chave)))
                        End If
                        
                    Next Chave
                    
                End If
                
                'Call fnExcel.DefinirTipoCampos(Campos, dicTitulos)
                dicC175(CHV_REG) = Campos
            
            End If
            
        Next Linha
        
    End With
    
    'Atualiza os dados do registro C175
    Call Util.LimparDados(regC175_Contr, 4, False)
    Call Util.ExportarDadosDicionario(regC175_Contr, dicC175)
    
    Call dicC175.RemoveAll
    
    Application.StatusBar = False
    If Not OmitirMsg Then Call Util.MsgInformativa("Valores atualizados com sucesso!", "Agrupamento de registro do C175", Inicio)
    
End Function

Public Sub ExcluirICMSBasePIS_COFINS()

Dim CHV_PAI As String, CST_PIS$
Dim dicDadosC100 As New Dictionary
Dim dicDadosC175 As New Dictionary
Dim arrChavesUnicas As New ArrayList
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC175 As New Dictionary
Dim Campos As Variant, CamposPai, Chave
    
    Inicio = Now()
    
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    
    Set dicDadosC175 = Util.CriarDicionarioRegistro(regC175_Contr)
    Set dicTitulosC175 = Util.MapearTitulos(regC175_Contr, 3)
    
    Set arrChavesUnicas = Util.ListarValoresOcorrenciaUnica(regC175_Contr, 4, 3, "CHV_PAI_FISCAL")
    
    For Each Chave In dicDadosC175.Keys()
        
        Campos = dicDadosC175(Chave)
        Call ExtrairCamposC175(Campos, dicTitulosC175)
        
        With CamposC175Contrib
            
            If arrChavesUnicas.contains(.CHV_PAI) And .CST_PIS Like "01*" And .ALIQ_PIS > 0 Then
                
                Call ExtrairCamposC100(dicDadosC100(.CHV_PAI), dicTitulosC100)
                
                .VL_DESC = CDbl(CamposC100.VL_ICMS) + CDbl(CamposC100.VL_DESC)
                .VL_BC_PIS = CDbl(CamposC100.VL_MERC) + CDbl(CamposC100.VL_FRT) + CDbl(CamposC100.VL_SEG) + CDbl(CamposC100.VL_OUT_DA) - CDbl(.VL_DESC)
                .VL_BC_COFINS = CDbl(.VL_BC_PIS)
                .VL_PIS = CDbl(.VL_BC_PIS) * CDbl(.ALIQ_PIS)
                .VL_COFINS = CDbl(.VL_BC_COFINS) * CDbl(.ALIQ_COFINS)
                
                Campos(dicTitulosC175("VL_OPER")) = CDbl(CamposC100.VL_MERC)
                Campos(dicTitulosC175("VL_DESC")) = CDbl(.VL_DESC)
                Campos(dicTitulosC175("VL_BC_PIS")) = CDbl(.VL_BC_PIS)
                Campos(dicTitulosC175("VL_PIS")) = CDbl(.VL_PIS)
                Campos(dicTitulosC175("VL_BC_COFINS")) = CDbl(.VL_BC_COFINS)
                Campos(dicTitulosC175("VL_COFINS")) = CDbl(.VL_COFINS)
                
                dicDadosC175(Chave) = Campos
                
            End If
        
        End With
        
    Next Chave
    
    Call Util.ExportarDadosDicionario(regC175_Contr, dicDadosC175, "A4")
    Call AtualizarImpostosC100(True)
    
    Call Util.MsgInformativa("ICMS excluído da base do PIS e COFINS com sucesso!", "Exclusão do ICMS da base do PIS e COFINS", Inicio)
    
End Sub

Private Function ExtrairCamposC100(ByRef Campos As Variant, ByRef dicTitulosC100 As Dictionary)
    
    With CamposC100
        
        .VL_ICMS = fnExcel.ConverterValores(Campos(dicTitulosC100("VL_ICMS")))
        .VL_DESC = fnExcel.ConverterValores(Campos(dicTitulosC100("VL_DESC")))
        .VL_FRT = fnExcel.ConverterValores(Campos(dicTitulosC100("VL_FRT")))
        .VL_SEG = fnExcel.ConverterValores(Campos(dicTitulosC100("VL_SEG")))
        .VL_OUT_DA = fnExcel.ConverterValores(Campos(dicTitulosC100("VL_OUT_DA")))
        .VL_MERC = fnExcel.ConverterValores(Campos(dicTitulosC100("VL_MERC")))
        .VL_ICMS = fnExcel.ConverterValores(Campos(dicTitulosC100("VL_ICMS")))
        .VL_ICMS = fnExcel.ConverterValores(Campos(dicTitulosC100("VL_ICMS")))
        
    End With
    
End Function

Private Function ExtrairCamposC175(ByRef Campos As Variant, ByRef dicTitulosC175 As Dictionary)
    
    With CamposC175Contrib
        
        .CHV_PAI = Campos(dicTitulosC175("CHV_PAI_FISCAL"))
        .VL_OPER = fnExcel.ConverterValores(Campos(dicTitulosC175("VL_OPER")))
        .CST_PIS = Util.ApenasNumeros(Campos(dicTitulosC175("CST_PIS")))
        .ALIQ_PIS = fnExcel.FormatarPercentuais(Campos(dicTitulosC175("ALIQ_PIS")))
        .VL_PIS = fnExcel.ConverterValores(Campos(dicTitulosC175("VL_PIS")))
        .CST_COFINS = Util.ApenasNumeros(Campos(dicTitulosC175("CST_COFINS")))
        .ALIQ_COFINS = fnExcel.FormatarPercentuais(Campos(dicTitulosC175("ALIQ_COFINS")))
        .VL_COFINS = fnExcel.ConverterValores(Campos(dicTitulosC175("VL_COFINS")))
        
    End With
    
End Function
