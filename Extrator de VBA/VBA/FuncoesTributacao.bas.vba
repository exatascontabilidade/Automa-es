Attribute VB_Name = "FuncoesTributacao"
Option Explicit

Public Function ImportarItensSPED()

Dim nReg As String, CNPJ$, cod$, DESC$, NCM$, CEST$, EX_TIPI$, TIPO_ITEM$
Dim Arqs As Variant, Arq, Registros, Registro
Dim dicTributacao As New Dictionary
    
    Arqs = Util.SelecionarArquivos("txt")
    If VarType(Arqs) <> 11 Then
        
        Inicio = Now()
        For Each Arq In Arqs
            
            a = Util.AntiTravamento(a)
            Registros = Util.ImportarTxt(Arq)
            For Each Registro In Registros
                
                With regEFD
                    
                    If Registro <> "" Then
                        
                        nReg = Mid(Registro, 2, 4)
                        Select Case True
                            
                            Case nReg = "0000"
                                CNPJ = fnSPED.ExtrairCampo(Registro, 7)
                                
                            Case nReg = "0200"
                                cod = Util.FormatarTexto(fnSPED.ExtrairCampo(Registro, 2))
                                DESC = Util.FormatarTexto(fnSPED.ExtrairCampo(Registro, 3))
                                TIPO_ITEM = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_TIPO_ITEM(fnSPED.ExtrairCampo(Registro, 7))
                                NCM = Util.FormatarTexto(VBA.Format(fnSPED.ExtrairCampo(Registro, 8), String(8, "0")))
                                EX_TIPI = Util.FormatarTexto(VBA.Format(fnSPED.ExtrairCampo(Registro, 9), "00"))
                                CEST = Util.FormatarTexto(VBA.Format(fnSPED.ExtrairCampo(Registro, 13), String(7, "0")))
                                
                                dicTributacao(cod) = Array(cod, DESC, TIPO_ITEM, NCM, EX_TIPI, CEST, "", "", "", "", "", "", "", "")
                                
                        End Select
                        
                    End If
                    
                End With
                
            Next Registro
            
        Next Arq
        
        Call Util.LimparDados(Tributacao, 4, False)
        Call Util.ExportarDadosDicionario(Tributacao, dicTributacao)
        
        Call Util.HabilitarControles
        Call Util.MsgInformativa("Cadastro extraído com sucesso", "Extração de Itens do SPED", Inicio)
        
    End If
    
End Function

Public Function ImportarCadastroTributacao()

Dim Arqs As Variant, Arq, Registros, Registro, Campos
Dim dicTributacao As New Dictionary
Dim Chave As String

    Arqs = Util.SelecionarArquivos("txt")
    If VarType(Arqs) <> 11 Then
        
        Inicio = Now()
        For Each Arq In Arqs
            
            a = Util.AntiTravamento(a)
            Registros = Util.ImportarTxt(Arq)
            For Each Registro In Registros
                
                With regEFD
                    
                    If Registro <> "" Then
                        
                        Campos = VBA.Split(Registro, "|")
                        Chave = fnSPED.MontarChaveRegistro(CStr(Campos(0)), CStr(Campos(6)))
                        dicTributacao(Chave) = Campos
                        
                    End If
                    
                End With
                
            Next Registro
            
        Next Arq
        
        Call Util.ExportarDadosDicionario(Tributacao, dicTributacao)
        
        Call Util.HabilitarControles
        Call Util.MsgInformativa("Cadastro importado com sucesso", "Importação de Cadastro de Tributação", Inicio)
        
    End If
    
End Function

Public Sub VerificarTributacao()

Dim CFOP As String, CST_ICMS$, CST_PIS_COFINS$, CST_IPI$, CHV_TRIB$, OBS$
Dim arrMsg As New ArrayList
Dim arrOBS As New ArrayList
Dim dicTributacao As New Dictionary
Dim dicDadosC170 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicTitulos As New Dictionary
Dim Chave As Variant, Campos
    
    Inicio = Now()
    Call FuncoesExcel.CarregarDados(Tributacao, dicTributacao, "CODIGO", "CFOP")
    Call FuncoesExcel.CarregarDados(regC170, dicDadosC170, "CHV_PAI_FISCAL", "NUM_ITEM")
    
    Set dicTitulos = Util.IndexarDados(Util.DefinirTitulos(Tributacao, 3))
    Set dicTitulosC170 = Util.IndexarDados(Util.DefinirTitulos(regC170, 3))
    
    If dicTributacao.Count = 1 Or dicDadosC170.Count = 2 Then
        
        If dicTributacao.Count = 1 Then arrMsg.Add "O relatório de tributação não possui informações."
        If dicDadosC170.Count = 2 Then arrMsg.Add "O registro C170 não possui informações."
        arrMsg.Add "Por favor preencha os dados para continuar a análise."
        
        Call Util.MsgAlerta(VBA.Join(arrMsg.toArray, vbCrLf), "Dados Insuficientes")
        Exit Sub
    
    End If
    
    For Each Chave In dicDadosC170.Keys()
        
        With CamposC170
            
            Campos = dicDadosC170(Chave)
            
                .COD_ITEM = Campos(dicTitulosC170("COD_ITEM"))
                .CFOP = Campos(dicTitulosC170("CFOP"))
                .CST_ICMS = Campos(dicTitulosC170("CST_ICMS"))
                .CST_IPI = Campos(dicTitulosC170("CST_IPI"))
                .CST_PIS = Campos(dicTitulosC170("CST_PIS"))
                .CST_COFINS = Campos(dicTitulosC170("CST_COFINS"))
                
                CHV_TRIB = fnSPED.MontarChaveRegistro(.COD_ITEM, .CFOP)
                If dicTributacao.Exists(CHV_TRIB) Then
                    
                    CST_ICMS = dicTributacao(CHV_TRIB)(dicTitulos("CST_ICMS"))
                    If CST_ICMS = .CST_ICMS Then arrOBS.Add "CST_ICMS = OK" Else arrOBS.Add "CST_ICMS = DIVERGENTE"
                    
                    CST_PIS_COFINS = dicTributacao(CHV_TRIB)(dicTitulos("CST_PIS_COFINS"))
                    If CST_PIS_COFINS = .CST_PIS Then arrOBS.Add "CST_PIS = OK" Else arrOBS.Add "CST_PIS = DIVERGENTE"
                    If CST_PIS_COFINS = .CST_COFINS Then arrOBS.Add "CST_COFINS = OK" Else arrOBS.Add "CST_COFINS = DIVERGENTE"
                    
                    CST_IPI = dicTributacao(CHV_TRIB)(dicTitulos("CST_IPI"))
                    If CST_IPI = .CST_IPI Then arrOBS.Add "CST_IPI = OK" Else arrOBS.Add "CST_IPI = DIVERGENTE"
                    
                Else
                
                    arrOBS.Add "PRODUTO / OPERAÇÃO NÃO CADASTRADA"
                
                End If
                
                OBS = VBA.Join(arrOBS.toArray, " / ")
                Campos(dicTitulosC170("OBS")) = OBS
            
            dicDadosC170(Chave) = Campos
            
        End With
        arrOBS.Clear
        
    Next Chave
        
    Call Util.LimparDados(regC170, 4, False)
    Call Util.ExportarDadosDicionario(regC170, dicDadosC170, "A4")
    Call Util.MsgInformativa("Análise de tributação concluída!", "Análise de Tributação", Inicio)
    Call regC170.Activate
    
End Sub
