Attribute VB_Name = "clsE111"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function GerarEstornoDecretoAtacadistaBA()

Dim ARQUIVO As String
Dim dicDadosC100 As New Dictionary
Dim dicDadosC170 As New Dictionary
Dim dicDadosE110 As New Dictionary
Dim dicDadosE111 As New Dictionary
Dim dicDadosE113 As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicTitulosE110 As New Dictionary
Dim Chave As Variant, ChaveE110, Campos
    
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    Set dicDadosC170 = Util.CriarDicionarioRegistro(regC170)
    Set dicDadosE110 = Util.CriarDicionarioRegistro(regE110)
    
    Call Util.IndexarCampos(dicDadosC100("CHV_NFE"), dicTitulosC100)
    Call dicDadosC100.Remove("CHV_NFE")
    
    Call Util.IndexarCampos(dicDadosC170("CHV_PAI|NUM_ITEM"), dicTitulosC170)
    Call dicDadosC170.Remove("CHV_PAI|NUM_ITEM")
    
    Call Util.IndexarCampos(dicDadosE110("ARQUIVO|CHV_PAI"), dicTitulosE110)
    Call dicDadosE110.Remove("ARQUIVO|CHV_PAI")
    
    CamposE111.VL_AJ_APUR = 0
    CamposE113.VL_AJ_ITEM = 0
    For Each ChaveE110 In dicDadosE110.Keys
        
        With CamposE110
            
            ARQUIVO = dicDadosE110(ChaveE110)(dicTitulosE110("ARQUIVO"))
            .CHV_PAI = dicDadosE110(ChaveE110)(dicTitulosE110("CHV_PAI_FISCAL"))
            
        End With
        
        With CamposE111
            
            .REG = "E111"
            .COD_AJ_APUR = "BA010005"
            .DESCR_COMPL_AJ = "ESTORNO DOS CRÉDITOS DE ICMS ACIMA DE 10% REF. AO DECRETO ATACADISTA"
            .CHV_REG = fnSPED.MontarChaveRegistro(ARQUIVO, CamposE110.CHV_PAI, .COD_AJ_APUR, .DESCR_COMPL_AJ)
            
        End With
        
        For Each Chave In dicDadosC170.Keys
            
            CamposC170.CFOP = dicDadosC170(Chave)(dicTitulosC170("CFOP"))
            CamposC170.COD_ITEM = dicDadosC170(Chave)(dicTitulosC170("COD_ITEM"))
            CamposC170.VL_ITEM = dicDadosC170(Chave)(dicTitulosC170("VL_ITEM"))
            CamposC170.VL_DESC = dicDadosC170(Chave)(dicTitulosC170("VL_DESC"))
            CamposC170.VL_ICMS = dicDadosC170(Chave)(dicTitulosC170("VL_ICMS"))
            CamposC170.CHV_PAI = dicDadosC170(Chave)(dicTitulosC170("CHV_PAI_FISCAL"))
            
            If CamposC170.CFOP < 3000 Then
                
                With CamposE113
                    
                    .VL_AJ_ITEM = CalcularEstornoDecretoAtacadistaBA(CamposC170.VL_ITEM - CamposC170.VL_DESC, CamposC170.VL_ICMS)
                    
                    If .VL_AJ_ITEM > 0 Then
                        
                        .REG = "E113"
                        .COD_PART = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("COD_PART"))
                        .COD_MOD = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("COD_MOD"))
                        .SER = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("SER"))
                        .SUB = ""
                        .NUM_DOC = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("NUM_DOC"))
                        .DT_DOC = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("DT_DOC"))
                        .COD_ITEM = CamposC170.COD_ITEM
                        .CHV_DOCE = CamposC170.CHV_PAI
                        
                        CamposE111.VL_AJ_APUR = CDbl(CamposE111.VL_AJ_APUR) + CDbl(.VL_AJ_ITEM)
                        
                        If .VL_AJ_ITEM > 0 Then
                            Campos = Array(.REG, ARQUIVO, CamposE111.CHV_REG, .COD_PART, .COD_MOD, _
                                           .SER, .SUB, .NUM_DOC, .DT_DOC, .COD_ITEM, CDbl(.VL_AJ_ITEM), .CHV_DOCE)
                            
                            .CHV_REG = fnSPED.MontarChaveRegistro(ARQUIVO, CamposE111.CHV_REG, .COD_MOD, .SER, .NUM_DOC, _
                                                                  .CHV_DOCE, .COD_ITEM, .VL_AJ_ITEM, dicDadosE113.Count)
                            dicDadosE113(.CHV_REG) = Campos
                        
                        End If
                    
                    End If
                    
                End With
                
            End If
            
        Next Chave
        
        With CamposE111
                            
            Campos = Array(.REG, ARQUIVO, CamposE110.CHV_PAI, .COD_AJ_APUR, .DESCR_COMPL_AJ, CDbl(.VL_AJ_APUR))
            dicDadosE111(.CHV_REG) = Campos
            .VL_AJ_APUR = 0
        
        End With
        
    Next ChaveE110
    
    Call Util.ExportarDadosDicionario(regE111, dicDadosE111)
    Call Util.ExportarDadosDicionario(regE113, dicDadosE113)
    
    Call dicDadosE111.RemoveAll
    Call dicDadosE113.RemoveAll
    
    Call regE111.Activate
    
End Function

Public Function GerarCreditoSIMPLESNACIONAL()

Dim arrXML As New ArrayList
Dim arrXMLs As New ArrayList
Dim ARQUIVO As String, Caminho$
Dim dicDadosC100 As New Dictionary
Dim dicDadosC170 As New Dictionary
Dim dicDadosE110 As New Dictionary
Dim dicDadosE111 As New Dictionary
Dim dicDadosE113 As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicTitulosE110 As New Dictionary
Dim dicFornecSN As New Dictionary

Dim Chave As Variant, ChaveE110, Campos
    
    Caminho = Util.SelecionarPasta("Selecione a pasta onde estão os XML")
    
    Call Util.ListarArquivos(arrXMLs, Caminho)
    Call fnXML.ColetarFornecedoresSN(dicFornecSN, arrXMLs)
    
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    Set dicDadosC170 = Util.CriarDicionarioRegistro(regC170)
    Set dicDadosE110 = Util.CriarDicionarioRegistro(regE110)
    
'    Call rC100.CarregarDados(dicDadosC100)
'    Call rC170.CarregarDados(dicDadosC170)
'    Call rE110.CarregarDados(dicDadosE110)
    
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    Set dicTitulosE110 = Util.MapearTitulos(regE110, 3)
    
'    Call Util.IndexarCampos(dicDadosC100("CHV_REG"), dicTitulosC100)
'    Call dicDadosC100.Remove("CHV_REG")
'
'    Call Util.IndexarCampos(dicDadosC170("CHV_PAI|NUM_ITEM"), dicTitulosC170)
'    Call dicDadosC170.Remove("CHV_PAI|NUM_ITEM")
'
'    Call Util.IndexarCampos(dicDadosE110("ARQUIVO|CHV_PAI"), dicTitulosE110)
'    Call dicDadosE110.Remove("ARQUIVO|CHV_PAI")
    
    CamposE111.VL_AJ_APUR = 0
    CamposE113.VL_AJ_ITEM = 0
    For Each ChaveE110 In dicDadosE110.Keys
        
        With CamposE110
            
            ARQUIVO = dicDadosE110(ChaveE110)(dicTitulosE110("ARQUIVO"))
            .CHV_PAI = dicDadosE110(ChaveE110)(dicTitulosE110("CHV_PAI_FISCAL"))
            
        End With
        
        With CamposE111
            
            .REG = "E111"
            .COD_AJ_APUR = "BA020008"
            .DESCR_COMPL_AJ = "CRÉDITOS NA AQUISIÇÃO DE EMPRESAS DO SIMPLES NACIONAL"
            .CHV_REG = fnSPED.MontarChaveRegistro(ARQUIVO, CamposE110.CHV_PAI, .COD_AJ_APUR, .DESCR_COMPL_AJ)
            
        End With
        
        For Each Chave In dicDadosC170.Keys
            
            CamposC170.CFOP = dicDadosC170(Chave)(dicTitulosC170("CFOP"))
            CamposC170.NUM_ITEM = dicDadosC170(Chave)(dicTitulosC170("NUM_ITEM"))
            CamposC170.COD_ITEM = dicDadosC170(Chave)(dicTitulosC170("COD_ITEM"))
            CamposC170.VL_ITEM = dicDadosC170(Chave)(dicTitulosC170("VL_ITEM"))
            CamposC170.VL_DESC = dicDadosC170(Chave)(dicTitulosC170("VL_DESC"))
            CamposC170.VL_ICMS = dicDadosC170(Chave)(dicTitulosC170("VL_ICMS"))
            CamposC170.CHV_PAI = dicDadosC170(Chave)(dicTitulosC170("CHV_PAI_FISCAL"))
            
            If CamposC170.CFOP < 3000 Then
                
                With CamposE113
                    
                    Chave = CamposC170.CHV_PAI & CInt(CamposC170.NUM_ITEM)
                    If dicFornecSN.Exists(Chave) Then
                    
                        If dicFornecSN(Chave)(0) <> "5101" Then
                            .VL_AJ_ITEM = CalcularCreditoAquisicaoSimplesNacional(CamposC170.VL_ITEM - CamposC170.VL_DESC, CDbl(dicFornecSN(Chave)(1)))
                        End If
                        
                    End If
                    
                    If .VL_AJ_ITEM > 0 Then
                        
                        .REG = "E113"
                        .COD_PART = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("COD_PART"))
                        .COD_MOD = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("COD_MOD"))
                        .SER = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("SER"))
                        .SUB = ""
                        .NUM_DOC = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("NUM_DOC"))
                        .DT_DOC = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("DT_DOC"))
                        .COD_ITEM = CamposC170.COD_ITEM
                        .CHV_DOCE = CamposC170.CHV_PAI
                        
                        CamposE111.VL_AJ_APUR = CDbl(CamposE111.VL_AJ_APUR) + CDbl(.VL_AJ_ITEM)
                        
                        If .VL_AJ_ITEM > 0 Then
                            Campos = Array(.REG, ARQUIVO, CamposE111.CHV_REG, .COD_PART, .COD_MOD, _
                                           .SER, .SUB, .NUM_DOC, .DT_DOC, .COD_ITEM, CDbl(.VL_AJ_ITEM), .CHV_DOCE)
                                            
                            .CHV_REG = fnSPED.MontarChaveRegistro(ARQUIVO, CamposE111.CHV_REG, .COD_MOD, .SER, .NUM_DOC, _
                                                                  .CHV_DOCE, .COD_ITEM, .VL_AJ_ITEM, dicDadosE113.Count)
                            dicDadosE113(.CHV_REG) = Campos
                        
                        End If
                        
                        .VL_AJ_ITEM = 0
                        
                    End If
                    
                End With
                
            End If
            
        Next Chave
        
        With CamposE111
                            
            Campos = Array(.REG, ARQUIVO, CamposE110.CHV_PAI, .COD_AJ_APUR, .DESCR_COMPL_AJ, CDbl(.VL_AJ_APUR))
            dicDadosE111(.CHV_REG) = Campos
            .VL_AJ_APUR = 0
        
        End With
        
    Next ChaveE110
    
    Call Util.ExportarDadosDicionario(regE111, dicDadosE111)
    Call Util.ExportarDadosDicionario(regE113, dicDadosE113)
    
    Call dicDadosE111.RemoveAll
    Call dicDadosE113.RemoveAll
    
    Call regE111.Activate
    
End Function

Public Function GerarCreditoArt269Inc10SEFAZBA()

Dim arrXML As New ArrayList
Dim arrXMLs As New ArrayList
Dim ARQUIVO As String, Caminho$
Dim dicDadosC100 As New Dictionary
Dim dicDadosC170 As New Dictionary
Dim dicDadosE110 As New Dictionary
Dim dicDadosE111 As New Dictionary
Dim dicDadosE113 As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicTitulosE110 As New Dictionary
Dim dicFornecSN As New Dictionary

Dim Chave As Variant, ChaveE110, Campos
    
    Caminho = Util.SelecionarPasta("Selecione a pasta onde estão os XML")
    
    Call Util.ListarArquivos(arrXMLs, Caminho)
    Call fnXML.ColetarFornecedoresSN(dicFornecSN, arrXMLs)
    
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    Set dicDadosC170 = Util.CriarDicionarioRegistro(regC170)
    Set dicDadosE110 = Util.CriarDicionarioRegistro(regE110)
        
    Call Util.IndexarCampos(dicDadosC100("CHV_NFE"), dicTitulosC100)
    Call dicDadosC100.Remove("CHV_NFE")
    
    Call Util.IndexarCampos(dicDadosC170("CHV_PAI|NUM_ITEM"), dicTitulosC170)
    Call dicDadosC170.Remove("CHV_PAI|NUM_ITEM")
    
    Call Util.IndexarCampos(dicDadosE110("ARQUIVO|CHV_PAI"), dicTitulosE110)
    Call dicDadosE110.Remove("ARQUIVO|CHV_PAI")
    
    CamposE111.VL_AJ_APUR = 0
    CamposE113.VL_AJ_ITEM = 0
    For Each ChaveE110 In dicDadosE110.Keys
        
        With CamposE110
            
            ARQUIVO = dicDadosE110(ChaveE110)(dicTitulosE110("ARQUIVO"))
            .CHV_PAI = dicDadosE110(ChaveE110)(dicTitulosE110("CHV_PAI_FISCAL"))
            
        End With
        
        With CamposE111
            
            .REG = "E111"
            .COD_AJ_APUR = "BA020009"
            .DESCR_COMPL_AJ = "CRÉDITOS NA AQUISIÇÃO INTERNA DE INDÚSTRIAS DO SIMPLES NACIONAL (ART. 269 INC X DO RICMS/BA)"
            .CHV_REG = fnSPED.MontarChaveRegistro(ARQUIVO, CamposE110.CHV_PAI, .COD_AJ_APUR, .DESCR_COMPL_AJ)
            
        End With
        
        For Each Chave In dicDadosC170.Keys
            
            CamposC170.CFOP = dicDadosC170(Chave)(dicTitulosC170("CFOP"))
            CamposC170.NUM_ITEM = dicDadosC170(Chave)(dicTitulosC170("NUM_ITEM"))
            CamposC170.COD_ITEM = dicDadosC170(Chave)(dicTitulosC170("COD_ITEM"))
            CamposC170.VL_ITEM = dicDadosC170(Chave)(dicTitulosC170("VL_ITEM"))
            CamposC170.VL_DESC = dicDadosC170(Chave)(dicTitulosC170("VL_DESC"))
            CamposC170.VL_ICMS = dicDadosC170(Chave)(dicTitulosC170("VL_ICMS"))
            CamposC170.CHV_PAI = dicDadosC170(Chave)(dicTitulosC170("CHV_PAI_FISCAL"))
            
            If CamposC170.CFOP < 3000 Then
                
                With CamposE113
                    
                    Chave = CamposC170.CHV_PAI & CInt(CamposC170.NUM_ITEM)
                    If dicFornecSN.Exists(Chave) Then
                    
                        If dicFornecSN(Chave)(0) = "5101" Then
                            .VL_AJ_ITEM = FuncoesSEFAZ_BA.CalcularCreditoPresumidoArt269Inc10BA(CamposC170.VL_ITEM - CamposC170.VL_DESC)
                        End If
                        
                    End If
                    
                    If .VL_AJ_ITEM > 0 Then
                        
                        .REG = "E113"
                        .COD_PART = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("COD_PART"))
                        .COD_MOD = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("COD_MOD"))
                        .SER = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("SER"))
                        .SUB = ""
                        .NUM_DOC = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("NUM_DOC"))
                        .DT_DOC = dicDadosC100(CamposC170.CHV_PAI)(dicTitulosC100("DT_DOC"))
                        .COD_ITEM = CamposC170.COD_ITEM
                        .CHV_DOCE = CamposC170.CHV_PAI
                        
                        CamposE111.VL_AJ_APUR = CDbl(CamposE111.VL_AJ_APUR) + CDbl(.VL_AJ_ITEM)
                        
                        If .VL_AJ_ITEM > 0 Then
                            Campos = Array(.REG, ARQUIVO, CamposE111.CHV_REG, .COD_PART, .COD_MOD, _
                                           .SER, .SUB, .NUM_DOC, .DT_DOC, .COD_ITEM, CDbl(.VL_AJ_ITEM), .CHV_DOCE)
                                            
                            .CHV_REG = fnSPED.MontarChaveRegistro(ARQUIVO, CamposE111.CHV_REG, .COD_MOD, .SER, .NUM_DOC, _
                                                                  .CHV_DOCE, .COD_ITEM, .VL_AJ_ITEM, dicDadosE113.Count)
                            dicDadosE113(.CHV_REG) = Campos
                        
                        End If
                        
                        .VL_AJ_ITEM = 0
                        
                    End If
                    
                End With
                
            End If
            
        Next Chave
        
        With CamposE111
                            
            Campos = Array(.REG, ARQUIVO, CamposE110.CHV_PAI, .COD_AJ_APUR, .DESCR_COMPL_AJ, CDbl(.VL_AJ_APUR))
            dicDadosE111(.CHV_REG) = Campos
            .VL_AJ_APUR = 0
        
        End With
        
    Next ChaveE110
    
    Call Util.ExportarDadosDicionario(regE111, dicDadosE111)
    Call Util.ExportarDadosDicionario(regE113, dicDadosE113)
    
    Call dicDadosE111.RemoveAll
    Call dicDadosE113.RemoveAll
    
    Call regE111.Activate
    
End Function

Public Function CriarE111(ByRef dicE111 As Dictionary, ByVal cAjuste As String, ByVal vEstorno As Double)
   
    If dicE111.Exists(cAjuste) Then vEstorno = vEstorno + CDbl(dicE111(cAjuste)) Else
    dicE111(cAjuste) = vEstorno
    
End Function

Public Function CalcularEstornoDecretoAtacadistaBA(ByVal vOperacao As Double, ByVal vICMS As Double) As Double
    
    'A base legal para o estorno de créditos que exceda 10% do valor da operação está no art. 6º do Decreto 7.799/2000 da SEFAZ/BA
    'Observação: O estorno de crédito não se aplicará as operações de entradas de mercadorias decorrentes de importação do exterior [Base Legal: §2º do art. 6º do Decreto 7.799/2000 da SEFAZ/BA]
    'Importante!: O estorno não se aplica as operações com papel higiênico [Base Legal: Parágrafo Único do art. 2º-A do Decreto 7.799/2000 da SEFAZ/BA]
    'Data da consulta: 06/02/2023
    
    Select Case True
    
        Case (vOperacao > 0) And (Round(vICMS / vOperacao, 2) > 0.1)
            CalcularEstornoDecretoAtacadistaBA = Round(vICMS - ((vOperacao * 0.1)), 2)
    
    End Select
    
End Function

Public Function CalcularCreditoAquisicaoSimplesNacional(ByVal vOperacao As Double, pCredSN As Double) As Double
    CalcularCreditoAquisicaoSimplesNacional = Round(vOperacao * pCredSN, 2)
End Function
