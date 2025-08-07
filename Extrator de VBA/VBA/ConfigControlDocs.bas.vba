Attribute VB_Name = "ConfigControlDocs"
Option Explicit

Public Sub RemoverLinhasGrade(ByVal MostrarGrades As Boolean)

Dim ws As Worksheet, PlanAtiva As Worksheet
    
    Set PlanAtiva = ActiveSheet
    
    Call Util.DesabilitarControles
    
        For Each ws In ThisWorkbook.Worksheets
            ws.Activate
            ActiveWindow.DisplayGridlines = MostrarGrades
        Next ws
    
    Call Util.HabilitarControles
    
    PlanAtiva.Activate
    
    Call AtualizarRibbon("Configurações")
    
    If MostrarGrades Then
        
        Call Util.MsgAviso("Linhas de grade ativadas com sucesso!", "Configuração das Linhas de Grade")
    
    Else
    
        Call Util.MsgAviso("Linhas de grade removidas com sucesso!", "Configuração das Linhas de Grade")
        
    End If
    
End Sub

Public Function ResetarConfiguracoes()
    
    With ConfiguracoesControlDocs
        
        .Range("IgnorarQtdUnidXML").value = False
        
    End With
    
End Function
