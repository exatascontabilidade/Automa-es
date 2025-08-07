Attribute VB_Name = "AssistentesOtimizacoesFiscais"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public OtimizacoesAtivas As Boolean

Public Function OtimizarAtualizacaoRegistros(ByRef Plan As Worksheet) As Boolean

Dim Inconsistencias As Long
    
    If Not ChecarInconsistenciasApuracao(Plan, Inconsistencias) Then
        
        OtimizarAtualizacaoRegistros = True
        
        If Inconsistencias > 0 Then Exit Function
        Call SugerirAtualizacaoTributaria(Plan)
        
    End If
    
End Function

Private Function SugerirAtualizacaoTributaria(ByRef Plan As Worksheet)

Dim Msg As String, Tributo$
Dim Result As VbMsgBoxResult
    
    Msg = "Gostaria de Salvar/Atualizar a tributação produzida nessa apuração?"
    Result = Util.MsgInformativaDecisao(Msg, "Assistente de Otimizações Fiscais")
    If Result = vbNo Then Exit Function
    
    Select Case True
        
        Case Plan.CodeName Like "*ICMS"
            Assistente.Tributario.ICMS.SalvarTributacaoICMS
            
        Case Plan.CodeName Like "*IPI"
            Assistente.Tributario.IPI.SalvarTributacaoIPI
            
        Case Plan.CodeName Like "*PISCOFINS"
            Assistente.Tributario.PIS_COFINS.SalvarTributacaoPISCOFINS
            
    End Select
    
End Function

Private Function ChecarInconsistenciasApuracao(ByRef Plan As Worksheet, ByRef Inconsistencias As Long) As Boolean

Dim dicTitulos As Dictionary
Dim Result As VbMsgBoxResult
Dim Msg As String
    
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    Inconsistencias = Util.ApenasNumeros(Plan.Cells(2, dicTitulos("INCONSISTENCIA")))
    
    If Inconsistencias = 0 Then Exit Function
    
    Msg = "Ainda existem inconsistências nessa apuração!" & vbCrLf & vbCrLf
    Msg = Msg & "Você ainda pode utilizar o recurso de ignorar inconsistências para garantir que não deixou nenhuma analise passar." & vbCrLf & vbCrLf
    Msg = Msg & "Deseja atualizar os registros do SPED mesmo assim?" & vbCrLf & vbCrLf
    Msg = Msg & "Clique em SIM para continuar ou em NÃO para corrigir as inconsistências."
    
    Result = Util.MsgInformativaDecisao(Msg, "Assistente de Otimizações Fiscais")
    If Result = vbNo Then ChecarInconsistenciasApuracao = True
    
End Function

Public Function SugerirCarregamentoTributacao(ByRef Plan As Worksheet)

Dim PlanTrib As Worksheet
    
    Select Case Plan.CodeName
        
        Case "assApuracaoICMS"
            Set PlanTrib = assTributacaoICMS
            
        Case "assApuracaoIPI"
            Set PlanTrib = assTributacaoIPI
            
        Case "assApuracaoPISCOFINS"
            Set PlanTrib = assTributacaoPISCOFINS
            
    End Select
    
    Call ChecarDadosTributarios(PlanTrib)

End Function

Private Sub ChecarDadosTributarios(ByRef Plan As Worksheet)

Dim UltLin As Long
Dim Msg As String, Tributo$
Dim Result As VbMsgBoxResult
    
    UltLin = Util.UltimaLinha(Plan, "A")
    If UltLin < 4 Then
        
        Tributo = VBA.Replace(Plan.CodeName, "assTributacao", "")
        Tributo = VBA.IIf(Tributo = "PISCOFINS", "PIS/COFINS", Tributo)
        
        Msg = "O Assistente de Tributação do " & Tributo & " não possui dados informados!" & vbCrLf & vbCrLf
        Msg = Msg & "Gostaria de importar a tributação do " & Tributo & " para auxiliar na apuração dos impostos?"
        Result = Util.MsgInformativaDecisao(Msg, "Assistente de Otimizações Fiscais")
        
        If Result = vbNo Then Exit Sub
        Call ImportarTributacao(Plan)
        
    End If
    
End Sub

Private Function ImportarTributacao(ByRef Plan As Worksheet)

    Select Case Plan.CodeName
        
        Case "assTributacaoICMS"
            Call Assistente.Tributario.ICMS.ImportarTributacaoICMS
            
        Case "assTributacaoIPI"
            Call Assistente.Tributario.IPI.ImportarTributacaoIPI
            
        Case "assTributacaoPISCOFINS"
            Call Assistente.Tributario.PIS_COFINS.ImportarTributacaoPISCOFINS
            
    End Select
        
End Function

Public Function SugerirSomadoIPIaoItem()

Dim dicDados0000 As New Dictionary
Dim arrValoresIND_ATIV As New ArrayList
        
    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000)
    If dicDados0000 Is Nothing Then Exit Function
    
    Set arrValoresIND_ATIV = Util.ListarValoresUnicos(reg0000, 4, 3, "IND_ATIV")
    If Not arrValoresIND_ATIV.contains("1 - Outros") Then Exit Function
    
    Call SomarIPIaoValorItem
    
End Function

Private Sub SomarIPIaoValorItem()

Dim Msg As String
Dim TituloMsg As String
Dim Result As VbMsgBoxResult
Dim arrCSTIPI As New ArrayList
Dim arrValoresIPI As New ArrayList
    
    Set arrValoresIPI = Util.ListarValoresUnicos(regC170, 4, 3, "VL_IPI")
    Set arrCSTIPI = Util.ListarValoresUnicos(regC170, 4, 3, "CST_IPI")
    
    If Not ChecarDadosIPI(arrCSTIPI, arrValoresIPI) Then Exit Sub
    
    TituloMsg = "Assistente de Otimizações Fiscais"
    Msg = "O SPED importado é de um NÃO contribuinte do IPI!" & vbCrLf & vbCrLf
    Msg = Msg & "Nestes casos o Perguntas Frequentes da EFD-ICMS/IPI, na questão 11.13.2.1 "
    Msg = Msg & "orienta a somar o valor do IPI ao campo VL_ITEM do Registro C170, VL_OPR do C190 "
    Msg = Msg & "e VL_MERC do registro C100." & vbCrLf & vbCrLf
    Msg = Msg & "Posso efetuar esse procedimento para você?"
    
    Result = Util.MsgInformativaDecisao(Msg, TituloMsg)
    If Result = vbNo Then Exit Sub
    Msg = Msg & ""
    Inicio = Now()
    
    Call rC170.SomarIPIeSTaosItens("IPI", True)
    Call rC170.AtualizarImpostosC100(True)
    Call rC170.GerarC190(True)
    
    Msg = "Valores do IPI somados ao valor do item do registro C170 com sucesso!"
    Call Util.MsgInformativa(Msg, TituloMsg, Inicio)
    
End Sub

Private Function ChecarDadosIPI(ByRef arrCSTIPI As ArrayList, ByRef arrValoresIPI As ArrayList) As Boolean
    
    If arrCSTIPI.Count = 0 And arrValoresIPI.Count = 0 Then Exit Function
    
    If arrCSTIPI.Count > 0 Then
        
        Select Case arrCSTIPI.Count
            
            Case 1
                If Not arrCSTIPI.contains("") Then
                    
                    ChecarDadosIPI = True
                    Exit Function
                    
                End If
                
            Case Else
                ChecarDadosIPI = True
                Exit Function
                
        End Select
        
    End If
    
    Select Case arrValoresIPI.Count
        
        Case 1
            If Not arrValoresIPI.contains("0") And Not arrValoresIPI.contains("") Then
                
                ChecarDadosIPI = True
                Exit Function
            
            End If
            
        Case 2
            If arrValoresIPI(0) <> "0" And Not IsEmpty(arrValoresIPI("0")) And _
               arrValoresIPI(1) <> "0" And Not IsEmpty(arrValoresIPI("1")) Then
                
                ChecarDadosIPI = True
                Exit Function
                
            End If
            
        Case Else
            ChecarDadosIPI = True
            Exit Function
            
    End Select

End Function
