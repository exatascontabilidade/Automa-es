Attribute VB_Name = "TratamentoErros"
Option Explicit

Public Sub TratarErros(ByRef Erro As ErrObject, ByVal nRotina As String)

Dim Msg As String, Tabela As String
    
    Select Case Erro.Number
            
        Case Else
            Msg = "Erro sem tratamento na rotina: " & nRotina & vbCrLf & vbCrLf
            Msg = Msg & "Código do erro: " & Err.Number & vbCrLf
            Msg = Msg & "Descrição do erro: " & Err.Description
            
            MsgBox Msg, vbCritical, "Erro não identificado"
            
    End Select
    
End Sub
