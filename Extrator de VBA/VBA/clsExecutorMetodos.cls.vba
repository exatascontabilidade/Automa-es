Attribute VB_Name = "clsExecutorMetodos"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Sub ExecutarMetodo(ByVal NomeMetodo As String, ByVal Processador As Iprocessador, ByVal TipoSPED As String, ByVal Registro As String)
    
    On Error GoTo TratarErro:
    
    Call Processador.Executar(TipoSPED, Registro)
    Set Processador = Nothing
    
Exit Sub
TratarErro:

Dim infoErro As New clsGerenciadorErros
    
    Call infoErro.NotificarErroInesperado(TypeName(Processador), NomeMetodo)
    
End Sub
