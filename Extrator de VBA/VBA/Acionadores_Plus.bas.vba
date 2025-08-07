Attribute VB_Name = "Acionadores_Plus"
Option Explicit

Public Function FuncoesPlus(control As IRibbonControl)
    
    Select Case control.id
        
        Case "btnImportarSPEDContribuicoes"
            Call FuncoesSPEDContribuicoes.ImportarSPEDContribuicoes
    
        Case "btnExportarSPEDContribuicoes"
            Call FuncoesSPEDContribuicoes.GerarEFDContribuicoes
                
    End Select
    
End Function


