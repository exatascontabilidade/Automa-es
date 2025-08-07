Attribute VB_Name = "clsC190Contrib"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Processador As Iprocessador
Private Executor As New clsExecutorMetodos

Public Sub AtualizarNCM_C190()
    
    Inicio = Now()
    Call Executor.ExecutarMetodo("AtualizarNCM_C190", New clsC190Contrib_AtualizarNCM, "Contribuições", "C190")
    
End Sub

Public Sub AgruparRegistros()
    
    Inicio = Now()
    Call Executor.ExecutarMetodo("AgruparRegistros", New fnSPED_AgruparRegistros, "Contribuições", "C190")
    
End Sub
