Attribute VB_Name = "clsTratamentoErros"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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
