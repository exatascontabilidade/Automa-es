Attribute VB_Name = "clsFuncoesSeguranca"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function VerificarDadosTributarios(ByVal tipo As String) As Boolean

Dim dicPlanilhas As New Dictionary
Dim dicPlanDados As New Dictionary
Dim dicPlanErro As New Dictionary
Dim Result As VbMsgBoxResult
Dim Tributo As String, Msg$
Dim Exportado As Boolean
Dim Plan As Worksheet
Dim UltLin As Long
Dim i As Long
    
    'Adicionar planilhas específicas a uma coleção
    dicPlanilhas.Add "ICMS", assTributacaoICMS
    dicPlanilhas.Add "IPI", assTributacaoIPI
    dicPlanilhas.Add "PISCOFINS", assTributacaoPISCOFINS
    
    Set dicPlanDados = ListarAssistentesTributariosComDados(dicPlanilhas)
    If dicPlanDados.Count = 0 Then
        VerificarDadosTributarios = True
        Exit Function
    End If
    
    Result = ApresentarMensagemUsuario(dicPlanDados, tipo)
    If Result = vbNo Then
        VerificarDadosTributarios = True
        Exit Function
    End If
    
    For i = 0 To dicPlanDados.Count - 1
        
        Set Plan = dicPlanDados.Items(i)
        Tributo = dicPlanDados.Keys()(i)
        Tributo = VBA.IIf(Tributo = "PIS/COFINS", "PIS-COFINS", Tributo)
        
        Exportado = Util.ExportarDadosRelatorio(Plan, "Tributação " & Tributo)
        If Exportado Then Call Util.LimparDados(Plan, 4, False) Else dicPlanErro.Add dicPlanDados.Keys()(i), Plan
        
    Next i
    
    If dicPlanErro.Count > 0 Then
        
        Call ApresentarMensagemErro(dicPlanErro)
        
    Else
        
        Msg = "Dados tributários exportados com sucesso!"
        Call Util.MsgAviso(Msg, "Assistente de Segurança de Dados")
        VerificarDadosTributarios = True
        
    End If
    
End Function

Private Function ListarAssistentesTributariosComDados(ByRef dicPlanilhas As Dictionary) As Dictionary

Dim dicDados As New Dictionary
Dim Result As VbMsgBoxResult
Dim Tributo As String, Msg$
Dim Exportado As Boolean
Dim Plan As Worksheet
Dim Chave As Variant
Dim UltLin As Long
Dim i As Long

    For Each Chave In dicPlanilhas.Keys()
        
        Set Plan = dicPlanilhas(Chave)
        If Not Plan Is Nothing Then
            
            UltLin = Util.UltimaLinha(Plan, "A")
            If UltLin > 3 Then

                Tributo = VBA.IIf(Chave = "PISCOFINS", "PIS/COFINS", Chave)
                dicDados.Add Tributo, Plan
                
            End If
            
        End If
        
    Next Chave
    
    Set ListarAssistentesTributariosComDados = dicDados
    
End Function

Private Function ApresentarMensagemUsuario(ByRef dicDados As Dictionary, ByVal tipo As String) As VbMsgBoxResult

Dim Msg As String
    
    Select Case dicDados.Count
        
        Case 1
            Msg = "O Assistente Tributário de " & dicDados.Keys()(0) & " possui dados informados!" & vbCrLf & vbCrLf
            
        Case 2
            Msg = "Os Assistentes Tributários de " & dicDados.Keys()(0) & " e " & dicDados.Keys()(1) & " possuem dados informados!" & vbCrLf & vbCrLf
            
        Case 3
            Msg = "Os Assistentes Tributários de " & dicDados.Keys()(0) & ", " & dicDados.Keys()(1) & " e " & dicDados.Keys()(2) & " possuem dados informados!" & vbCrLf & vbCrLf
            
    End Select
    
    If dicDados.Count > 1 Then
        
        Msg = Msg & "Deseja exportar essas informações antes de " & tipo & " o ControlDocs?"
        
    ElseIf dicDados.Count = 1 Then
        
        Msg = Msg & "Deseja exportar essa informação antes de " & tipo & " o ControlDocs?"
        
    End If
    
    ApresentarMensagemUsuario = Util.MsgDecisao(Msg, "Assistente de Segurança dos Dados")
    
End Function

Private Sub ApresentarMensagemErro(ByRef dicPlanErro As Dictionary)

Dim i As Long
Dim Msg As String
Dim Tributo As String
    
    'Construir a mensagem de erro
    Select Case dicPlanErro.Count
        
        Case 1
            Tributo = dicPlanErro.Keys()(0)
            Msg = "Ocorreu um erro ao exportar os dados do Assistente Tributário de " & Tributo & "."
            
        Case 2
            Msg = "Não foi possível exportar as tributações dos Assistentes Tributários de " & _
                  dicPlanErro.Keys()(0) & " e " & dicPlanErro.Keys()(1) & "."
            
        Case 3
            Msg = "Não foi possível exportar as tributações dos Assistentes Tributários de " & _
                  dicPlanErro.Keys()(0) & ", " & dicPlanErro.Keys()(1) & " e " & dicPlanErro.Keys()(2) & "."
            
    End Select
    
    If dicPlanErro.Count > 1 Then
        
        Msg = Msg & vbCrLf & vbCrLf & _
              "Por favor, considere executar o procedimento de exportação diretamente nos assistentes afetados."
    
    ElseIf dicPlanErro.Count = 1 Then
        
        Msg = Msg & vbCrLf & vbCrLf & _
              "Por favor, considere executar o procedimento de exportação diretamente no assistente afetado."
        
    End If
    
    Call Util.MsgCritica(Msg, "Erro na Exportação dos Dados")
    
End Sub

