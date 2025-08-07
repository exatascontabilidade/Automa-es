Attribute VB_Name = "FuncoesFiscais"
Option Explicit

Public Function CalcularPerncentualCredito(ByVal vOperacao As Double, ByVal vICMS As Double) As Double
    If vOperacao > 0 Then CalcularPerncentualCredito = Round((vICMS / vOperacao), 2)
End Function

Public Function CriarAtualizarAjustes(ByRef dicAjustes As Dictionary, ByVal cAjuste As String, ByVal vAjuste As Double)
   
    If dicAjustes.Exists(cAjuste) Then vAjuste = Round(vAjuste + dicAjustes(cAjuste), 2)
    dicAjustes(cAjuste) = vAjuste
    
End Function

Public Function IncluirAjustesE111eE113(ByRef dicAjustes As Dictionary, ByRef dicE113 As Dictionary, ByRef EFD As ArrayList)

Dim cAjuste As Variant

    For Each cAjuste In dicAjustes
        EFD.Add Join(Array("", "E111", cAjuste, "", dicAjustes(cAjuste), ""), "|")
        EFD.Add Join(dicE113(cAjuste).toArray, vbCrLf)
    Next cAjuste
    
End Function

Public Function CalcularReducaoBaseICMS(ByRef Campos As Variant)
    
    Select Case True
    
        Case (VBA.Right(Campos(2), 2) = "20") Or (VBA.Right(Campos(2), 2) = "70")
            Campos(10) = VBA.Round(CDbl(Campos(5)) - CDbl(Campos(6)) - CDbl(Campos(9)) - CDbl(Campos(11)), 2)
            
    End Select
    
End Function
