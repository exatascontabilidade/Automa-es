Attribute VB_Name = "clsM400"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function Calcular_M400_M800(ByVal ExcluirICMS As Boolean)

Dim dicTitulos0000 As New Dictionary
Dim dicTitulosA100 As New Dictionary
Dim dicTitulosA170 As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicTitulosC185 As New Dictionary
Dim dicTitulosC385 As New Dictionary
Dim dicTitulosC485 As New Dictionary
Dim dicTitulosC495 As New Dictionary
Dim dicTitulosC605 As New Dictionary
Dim dicTitulosD205 As New Dictionary
Dim dicTitulosD300 As New Dictionary
Dim dicTitulosD350 As New Dictionary
Dim dicTitulosD605 As New Dictionary
Dim dicTitulosF100 As New Dictionary
Dim dicTitulosF200 As New Dictionary
Dim dicTitulosF500 As New Dictionary
Dim dicTitulosF510 As New Dictionary
Dim dicTitulosF550 As New Dictionary
Dim dicTitulosF560 As New Dictionary
Dim dicTitulosI100 As New Dictionary
Dim dicTitulosM400 As New Dictionary
Dim dicTitulosM800 As New Dictionary

Dim dicDados0000 As New Dictionary
Dim dicDadosA100 As New Dictionary
Dim dicDadosA170 As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim dicDadosC170 As New Dictionary
Dim dicDadosC185 As New Dictionary
Dim dicDadosC385 As New Dictionary
Dim dicDadosC485 As New Dictionary
Dim dicDadosC495 As New Dictionary
Dim dicDadosC605 As New Dictionary
Dim dicDadosD205 As New Dictionary
Dim dicDadosD300 As New Dictionary
Dim dicDadosD350 As New Dictionary
Dim dicDadosD605 As New Dictionary
Dim dicDadosF100 As New Dictionary
Dim dicDadosF200 As New Dictionary
Dim dicDadosF500 As New Dictionary
Dim dicDadosF510 As New Dictionary
Dim dicDadosF550 As New Dictionary
Dim dicDadosF560 As New Dictionary
Dim dicDadosI100 As New Dictionary
Dim dicDadosM400 As New Dictionary
Dim dicDadosM800 As New Dictionary

Dim Campos As Variant, Campo, nCampo, Titulos
Dim Chave As String, CHV_0000$, CHV_M001$
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim VL_TOTAL As Double
Dim Valores As Variant
Dim Inicio As Date
Dim i As Long
    
    Inicio = Now()
    Application.StatusBar = "gerando registros M400 e M800, por favor aguarde..."
    
    VL_TOTAL = 0
    
    Set dicTitulos0000 = Util.MapearTitulos(reg0000_Contr, 3)
    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000_Contr, "ARQUIVO")
    
    Set dicTitulosA100 = Util.MapearTitulos(regA100, 3)
    Set dicDadosA100 = Util.CriarDicionarioRegistro(regA100)
    
    Set dicTitulosA170 = Util.MapearTitulos(regA170, 3)
    Set dicDadosA170 = Util.CriarDicionarioRegistro(regA170)
    
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    Set dicDadosC170 = Util.CriarDicionarioRegistro(regC170)
    
    Set dicTitulosC185 = Util.MapearTitulos(regC185, 3)
    Set dicDadosC185 = Util.CriarDicionarioRegistro(regC185)
    
    Set dicTitulosC385 = Util.MapearTitulos(regC381, 3)
    Set dicDadosC385 = Util.CriarDicionarioRegistro(regC385)
    
    Set dicTitulosC485 = Util.MapearTitulos(regC485, 3)
    Set dicDadosC485 = Util.CriarDicionarioRegistro(regC485)
    
    Set dicTitulosC495 = Util.MapearTitulos(regC495, 3)
    Set dicDadosC495 = Util.CriarDicionarioRegistro(regC495)
    
    Set dicTitulosC605 = Util.MapearTitulos(regC605, 3)
    Set dicDadosC605 = Util.CriarDicionarioRegistro(regC605)
    
    Set dicTitulosD205 = Util.MapearTitulos(regD205, 3)
    Set dicDadosD205 = Util.CriarDicionarioRegistro(regD205)
    
    Set dicTitulosD605 = Util.MapearTitulos(regD605, 3)
    Set dicDadosD605 = Util.CriarDicionarioRegistro(regD605)
    
    Set dicTitulosD300 = Util.MapearTitulos(regD300, 3)
    Set dicDadosD300 = Util.CriarDicionarioRegistro(regD300)
    
    Set dicTitulosD350 = Util.MapearTitulos(regD350, 3)
    Set dicDadosD350 = Util.CriarDicionarioRegistro(regD350)
    
    Set dicTitulosF100 = Util.MapearTitulos(regF100, 3)
    Set dicDadosF100 = Util.CriarDicionarioRegistro(regF100)
    
    Set dicTitulosF200 = Util.MapearTitulos(regF200, 3)
    Set dicDadosF200 = Util.CriarDicionarioRegistro(regF200)
    
    Set dicTitulosF500 = Util.MapearTitulos(regF500, 3)
    Set dicDadosF500 = Util.CriarDicionarioRegistro(regF500)
    
    Set dicTitulosF510 = Util.MapearTitulos(regF510, 3)
    Set dicDadosF510 = Util.CriarDicionarioRegistro(regF510)
    
    Set dicTitulosF550 = Util.MapearTitulos(regF550, 3)
    Set dicDadosF550 = Util.CriarDicionarioRegistro(regF550)
    
    Set dicTitulosF560 = Util.MapearTitulos(regF560, 3)
    Set dicDadosF560 = Util.CriarDicionarioRegistro(regF560)
    
    Set dicTitulosI100 = Util.MapearTitulos(regI100, 3)
    Set dicDadosI100 = Util.CriarDicionarioRegistro(regI100)
    
    Set dicTitulos = Util.MapearTitulos(regC010, 3)
    Set Dados = Util.DefinirIntervalo(regC010, 4, 3)
    If Not Dados Is Nothing Then
                
        For Each Linha In Dados.Rows
            
            Campos = Application.Transpose(Application.Transpose(Linha))
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                ARQUIVO = Campos(dicTitulos("ARQUIVO"))
                CHV_0000 = dicDados0000(ARQUIVO)(dicTitulos0000("CHV_REG"))
                CHV_M001 = fnSPED.GerarChaveRegistro(CHV_0000, "M001")
                VL_TOTAL = VL_TOTAL + CalcularReceitaIndividualizada(dicDadosA100, dicDadosA170, dicTitulosA100, dicTitulosA170, dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosC385, dicTitulosC385, "CST_COFINS", "VL_ITEM", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosC605, dicTitulosC605, "CST_COFINS", "VL_ITEM", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosD205, dicTitulosD205, "CST_COFINS", "VL_ITEM", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosD605, dicTitulosD605, "CST_COFINS", "VL_ITEM", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosD300, dicTitulosD300, "CST_COFINS", "VL_DOC", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosD350, dicTitulosD350, "CST_COFINS", "VL_BRT", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosF100, dicTitulosF100, "CST_PIS", "VL_OPER", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosF200, dicTitulosF200, "CST_COFINS", "VL_TOT_REC", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosF500, dicTitulosF500, "CST_COFINS", "VL_REC_CAIXA", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosF510, dicTitulosF510, "CST_COFINS", "VL_REC_CAIXA", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosF550, dicTitulosF550, "CST_COFINS", "VL_REC_COMP", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosF560, dicTitulosF560, "CST_COFINS", "VL_REC_COMP", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosI100, dicTitulosI100, "CST_COFINS", "VL_REC", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                
                If Campos(dicTitulos("IND_ESCRI")) = "1" Then
                    
                    VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosC185, dicTitulosC185, "CST_COFINS", "VL_ITEM", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                    VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosC495, dicTitulosC495, "CST_COFINS", "VL_ITEM", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                    
                Else
                                        
                    VL_TOTAL = VL_TOTAL + CalcularReceitaIndividualizada(dicDadosC100, dicDadosC170, dicTitulosC100, dicTitulosC170, dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                    VL_TOTAL = VL_TOTAL + CalcularReceitas(dicDadosC485, dicTitulosC485, "CST_COFINS", "VL_ITEM", "VL_COFINS", dicDadosM400, dicDadosM800, CHV_M001, ARQUIVO)
                    
                End If
            
            End If
            'Debug.Print VL_TOTAL
            
        Next Linha
        
    End If
    
    Application.StatusBar = False
    Call Util.MsgInformativa("Registros M400 e M800 gerados com sucesso!", "Geração dos Registros M400 e M800", Inicio)
              
End Function

Private Function CalcularReceitas(ByRef dicDados As Dictionary, ByRef dicTitulos As Dictionary, ByVal TituloCST As String, _
    ByVal TituloBase As String, ByVal TituloImposto As String, ByRef dicDadosM400 As Dictionary, ByRef dicDadosM800 As Dictionary, _
    ByVal CHV_PAI As String, ByVal ARQUIVO As String) As Double

Dim Chave As Variant
Dim CST As String, REG$, IND_OPER$, Arq$, CHV_REG$, COD_CTA$, DESC_COMPL$
Dim VL_TOTAL As Double, VL_PROD#, VL_IMP#, VL_TOT_REC#
Dim arrCSTs As New ArrayList
    
    arrCSTs.Add "04": arrCSTs.Add "05": arrCSTs.Add "06": arrCSTs.Add "07": arrCSTs.Add "08": arrCSTs.Add "09"
    VL_TOTAL = 0
    
    If dicDados.Count > 0 Then
    
        For Each Chave In dicDados.Keys()
            
            Arq = dicDados(Chave)(dicTitulos("ARQUIVO"))
            If Arq = ARQUIVO Then
                
                REG = dicDados(Chave)(dicTitulos("REG"))
                CST = dicDados(Chave)(dicTitulos(TituloCST))
                VL_PROD = dicDados(Chave)(dicTitulos(TituloBase))
                VL_IMP = dicDados(Chave)(dicTitulos(TituloImposto))
                If REG = "F100" Then
                    IND_OPER = dicDados(Chave)(dicTitulos("IND_OPER"))
                    If IND_OPER = "0" Then VL_PROD = 0
                End If
                
            End If
            
            If arrCSTs.contains(CST) And VL_IMP = 0 Then
                
                CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, CST, COD_CTA, DESC_COMPL)
                If dicDadosM400.Exists(CHV_REG) Then
                    VL_TOT_REC = CDbl(dicDadosM400(CHV_REG)(4)) + VL_PROD
                    dicDadosM400(CHV_REG) = Array("M400", ARQUIVO, CHV_REG, CHV_PAI, CST, VL_TOT_REC, COD_CTA, DESC_COMPL)
                End If
                
            End If
            
        Next Chave
    
    End If
    
End Function

Private Function CalcularReceitaIndividualizada(ByRef dicPai As Dictionary, ByRef dicFilho As Dictionary, _
    ByRef dicTitulosPai As Dictionary, ByRef dicTitulosFilho As Dictionary, ByRef dicDadosM400 As Dictionary, _
    ByRef dicDadosM800 As Dictionary, ByVal CHV_PAI As String, ByVal ARQUIVO As String) As Double

Dim Chave As Variant
Dim arrCSTs As New ArrayList
Dim VL_TOTAL As Double, VL_PROD#, VL_IMP#, VL_TOT_REC#
Dim CST As String, REG$, IND_OPER$, Arq$, COD_CTA$, DESC_COMPL$, CHV_REG$, chvPai$
   
    arrCSTs.Add "04": arrCSTs.Add "05": arrCSTs.Add "06": arrCSTs.Add "07": arrCSTs.Add "08": arrCSTs.Add "09"
    If dicFilho.Count > 0 Then
        
        For Each Chave In dicFilho.Keys()
            
            chvPai = dicFilho(Chave)(dicTitulosFilho("CHV_PAI_FISCAL"))
            Arq = dicFilho(Chave)(dicTitulosFilho("ARQUIVO"))
            If dicPai(chvPai)(dicTitulosPai("IND_OPER")) = "0" Or Arq <> ARQUIVO Then GoTo Prx:
            
            REG = dicFilho(Chave)(dicTitulosFilho("REG"))
            CST = dicFilho(Chave)(dicTitulosFilho("CST_COFINS"))
            VL_PROD = dicFilho(Chave)(dicTitulosFilho("VL_ITEM"))
            VL_IMP = dicFilho(Chave)(dicTitulosFilho("VL_COFINS"))
            
            If arrCSTs.contains(CST) And VL_IMP = 0 Then
            
                CHV_REG = fnSPED.GerarChaveRegistro(CHV_PAI, CST, COD_CTA, DESC_COMPL)
                If dicDadosM400.Exists(CHV_REG) Then
                
                    VL_TOT_REC = CDbl(dicDadosM400(CHV_REG)(4)) + VL_PROD
                    dicDadosM400(CHV_REG) = Array("M400", ARQUIVO, CHV_REG, CHV_PAI, CST, VL_TOT_REC, COD_CTA, DESC_COMPL)
                    dicDadosM800(CHV_REG) = Array("M800", ARQUIVO, CHV_REG, CHV_PAI, CST, VL_TOT_REC, COD_CTA, DESC_COMPL)
                    
                End If
                
            End If
            
Prx:
        Next Chave
    
    End If
    
End Function
