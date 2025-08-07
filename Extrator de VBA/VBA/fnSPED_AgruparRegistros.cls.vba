Attribute VB_Name = "fnSPED_AgruparRegistros"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Implements Iprocessador

Public TipoSPED As String
Public Registro As String

Public Sub IProcessador_Executar(ByVal tpSPED As String, ByVal nReg As String)
    
    Me.TipoSPED = tpSPED
    Me.Registro = nReg
    
    Call AgruparRegistros
    
End Sub

Private Sub AgruparRegistros()

Dim Agrupador As IAgruparRegistros
    
    Select Case True
        
        Case TipoSPED = "Contribuições" And Registro = "C180"
            Set Agrupador = New fnSPED_AgruparRegistrosC180Cont
            Call Agrupador.AgruparRegistros
            
        Case TipoSPED = "Contribuições" And Registro = "C190"
            Set Agrupador = New fnSPED_AgruparRegistrosC190Cont
            Call Agrupador.AgruparRegistros
            
        Case Else
            Call Util.MsgAlerta("Registro não mapeado para função de Agrupamento de registros", "Registro Não Mapeado")
            
    End Select
    
End Sub
