Attribute VB_Name = "ConfigOtimizacoesFiscais"
Option Explicit

Public Function AlternarOtimizacoesFiscais(ByRef control As IRibbonControl, Check As Boolean)
    
    ConfiguracoesControlDocs.Range("OtimizacoesFiscais").value = Check
    Otimizacoes.OtimizacoesAtivas = Check
    
End Function

Public Function ObterStatusOtimizacoesFiscais(ByRef control As IRibbonControl, Check)
    
    Check = CBool(ConfiguracoesControlDocs.Range("OtimizacoesFiscais").value)
    Otimizacoes.OtimizacoesAtivas = Check
    
End Function
