Attribute VB_Name = "AssistenteClassificacaoSPEDs"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private ClassificadorArquivos As New AssistenteClassificacaoArquivos

Public Function ListarArquivosSPED(ByVal Caminho As String)

Dim ArquivosListados As Variant
    
    On Error GoTo Notificar:
        
    If Not ClassificadorArquivos.EncontrouPowerShell() Then Exit Function
    
    Call Util.AtualizarBarraStatus("Listando arquivos selecionados...")
    ArquivosListados = ClassificadorArquivos.ListarArquivos(Caminho, "txt")
    
    Call ClassificarSPEDS(ArquivosListados)
    
Exit Function
Notificar:

    Call TratarExcecoes
    
End Function

Public Function ClassificarSPEDS(ByVal ArquivosListados As Variant)

Dim TotalArquivos As Long
Dim SPED As Variant
Dim Arq As String
    
    On Error GoTo Notificar:
    
    ArquivosListados = Util.ConverterArrayListEmArray(ArquivosListados)
    
    a = 0
    Comeco = Timer()
    TotalArquivos = UBound(ArquivosListados) + 1
    For Each SPED In ArquivosListados
        
        Arq = CStr(VBA.Trim(SPED))
        
        If Arq <> "" Then
            
            Call Util.AntiTravamento(a, 50, "Listando arquivos SPED, por favor aguarde...", TotalArquivos, Comeco)
            Select Case IdentificarTipoSPED(Arq)
                
                Case "Fiscal"
                    DocsFiscais.arrSPEDFiscal.Add Arq
                    
                Case "Contribuições"
                    DocsFiscais.arrSPEDContribuicoes.Add Arq
                    
                Case "Desconhecido"
                    DocsFiscais.arrSPEDsInvalidos.Add Arq
                    
            End Select
            
        End If
        
    Next SPED
    
Exit Function
Notificar:
    
    Call TratarExcecoes
    
End Function

Private Function IdentificarTipoSPED(ByVal SPED As String) As String
    
Dim Registro As String
    
    Registro = ExtrairRegistroAbertura(SPED)
    If Registro Like "|0000*" Then
    
        IdentificarTipoSPED = ClassificarSPED(Registro)
    
    Else
        
        IdentificarTipoSPED = "Desconhecido"
        
    End If
    
End Function

Private Function ClassificarSPED(ByVal Registro As String) As String

Dim Campos As Variant
    
    Campos = VBA.Split(Registro, "|")
    Select Case True
    
        Case ValidarSPEDFiscal(Campos)
            ClassificarSPED = "Fiscal"
        
        Case ValidarSPEDContribuicoes(Campos)
            ClassificarSPED = "Contribuições"
        
        Case Else
            ClassificarSPED = "Desconhecido"
        
    End Select
    
End Function

Private Function ValidarSPEDFiscal(ByVal Campos As Variant) As Boolean

    If IsDate(Util.FormatarData(Campos(4))) And IsDate(Util.FormatarData(Campos(5))) Then ValidarSPEDFiscal = True
    
End Function

Private Function ValidarSPEDContribuicoes(ByVal Campos As Variant) As String
    
    If IsDate(Util.FormatarData(Campos(6))) And IsDate(Util.FormatarData(Campos(7))) Then ValidarSPEDContribuicoes = True
    
End Function

Function ExtrairRegistroAbertura(ByVal Arq As String) As String
    
Dim fso As New FileSystemObject
Dim Registro As String
Dim ts As Object
    
    Set ts = fso.OpenTextFile(Arq, 1)
    Registro = ts.ReadLine
    
    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing
    
    ExtrairRegistroAbertura = Registro
    
End Function

Private Function TratarExcecoes()

    Select Case True

        Case Else
            With infNotificacao
        
                .Funcao = "ListarArquivosSPED"
                .Classe = "AssistenteClassificacaoSPEDs"
                .MensagemErro = Err.Number & " - " & Err.Description
                .OBSERVACOES = "Erro Inesperado"
                
            End With
            
            Call Notificacoes.NotificarErroInesperado
            
    End Select

End Function
