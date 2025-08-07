Attribute VB_Name = "ConfigAssistenteTributario"
Option Explicit

Public Function AlternarAtualizacaoM400M800(ByRef control As IRibbonControl, Check As Boolean)
    
    ConfiguracoesControlDocs.Range("ManterM400M800").value = Check
    AtualizarM400M800 = Check
    
End Function

Public Function ObterStatusAtualizacaoM400M800(ByRef control As IRibbonControl, Check)
    
    Check = CBool(ConfiguracoesControlDocs.Range("ManterM400M800").value)
    AtualizarM400M800 = Check
    
End Function

