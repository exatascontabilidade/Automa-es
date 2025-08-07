Attribute VB_Name = "AnalistaApuracaoICMS"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private dicResumoICMS As Dictionary
Private Analista As AnalistaApuracao
Private ValidacoesICMS As AnalistaApuracaoICMS_Validacoes

Public Sub GerarResumoApuracaoICMS()

Dim arrDadosApuracao As New ArrayList
Dim Comeco As Double, VL_OPR#
Dim Campos As Variant, nCampo
Dim Chave As String
    
    Inicio = Now()
    If Util.ChecarAusenciaDados(assApuracaoICMS, False) Then Exit Sub
        
    Set dicResumoICMS = New Dictionary
    Set Analista = New AnalistaApuracao
    Set ValidacoesICMS = New AnalistaApuracaoICMS_Validacoes
    
    Set arrDadosApuracao = Util.CriarArrayListRegistro(assApuracaoICMS)
    Set Analista.dicTitulosApuracao = Util.MapearTitulos(assApuracaoICMS, 3)
    Set Analista.dicTitulosResumo = Util.MapearTitulos(resICMS, 3)
    Set Analista.dicTitulos = Analista.dicTitulosResumo
    
    Call ValidacoesICMS.InicializarObjetos
    
    a = 0
    Comeco = Timer
    With Analista
        
        For Each Campos In arrDadosApuracao
            
            .RedimensionarArray (.dicTitulosResumo.Count)
            Call Util.AntiTravamento(a, 100, "Gerando resumo de apuração do ICMS, por favor aguarde...", arrDadosApuracao.Count, Comeco)
            
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                For Each nCampo In .dicTitulosResumo.Keys()
                    
                    Call MontarRegistroICMS(Campos, nCampo)
                    
                Next nCampo
                
            End If
            
            .Campos = ValidacoesICMS.ValidarRegrasResumoICMS(.Campos)
            
            Chave = GerarChaveResumoICMS(Campos)
            If dicResumoICMS.Exists(Chave) Then Call AtualizarResumoICMS(Chave)
            
            dicResumoICMS(Chave) = .Campos
            
        Next Campos
        
        Call Util.LimparDados(resICMS, 4, False)
        Call Util.ExportarDadosDicionario(resICMS, dicResumoICMS)
        Call FuncoesFormatacao.DestacarInconsistencias(resICMS)
        
        resICMS.Activate
        Call Util.MsgInformativa("Resumo ICMS gerado com sucesso!", "Resumo Apuração do ICMS", Inicio)
        
    End With
    
End Sub

Private Function MontarRegistroICMS(ByVal Campos As Variant, ByVal nCampo As String)

Dim vCampo As Variant
Dim nTitulo As String
    
    With Analista
        
        Select Case nCampo
            
            Case "INCONSISTENCIA", "SUGESTAO"
                vCampo = Empty
                
            Case Else
                vCampo = TratarCamposResumoICMS(nCampo, Campos)
                
        End Select
        
        .AtribuirValor nCampo, vCampo
        
    End With
    
End Function

Private Function TratarCamposResumoICMS(ByVal nCampo As String, ByVal Campos As Variant)

Dim vCampo As Variant
    
    With Analista
    
        vCampo = Campos(.dicTitulosApuracao(nCampo))
        
        Select Case True
            
            Case nCampo = "VL_ITEM"
                TratarCamposResumoICMS = CalcularCampoVL_ITEM(Campos)
                
            Case nCampo Like "VL_*"
                TratarCamposResumoICMS = VBA.Round(vCampo, 2)
            
            Case nCampo Like "*CST*"
                TratarCamposResumoICMS = fnExcel.FormatarTexto(vCampo)
                
            Case Else
                TratarCamposResumoICMS = vCampo
                
        End Select
        
    End With
    
End Function

Private Function CalcularCampoVL_ITEM(ByVal Campos As Variant) As Double
    
Dim VL_ITEM As Double, VL_DESP#, VL_DESC#
    
    With Analista
        
        VL_ITEM = Campos(.dicTitulosApuracao("VL_ITEM"))
        VL_DESP = Campos(.dicTitulosApuracao("VL_DESP"))
        VL_DESC = Campos(.dicTitulosApuracao("VL_DESC"))
        
        CalcularCampoVL_ITEM = VL_ITEM + VL_DESP - VL_DESC
        
    End With
    
End Function

Private Function GerarChaveResumoICMS(ByVal Campos As Variant) As String

Dim CamposChave As Variant, Campo
Dim arrCampos As New ArrayList

    CamposChave = Array("CFOP", "CST_ICMS", "ALIQ_ICMS")
    
    With Analista
        
        For Each Campo In CamposChave
            
            arrCampos.Add Campos(.dicTitulosApuracao(Campo))
            
        Next Campo
        
        GerarChaveResumoICMS = fnSPED.GerarChaveRegistro(VBA.Join(arrCampos.toArray()))
        
    End With
    
End Function

Private Function AtualizarResumoICMS(ByVal Chave As String)

Dim CamposResumo As Variant, nCampo
Dim vCampo As Double
    
    With Analista
        
        CamposResumo = dicResumoICMS(Chave)
        For Each nCampo In .dicTitulosResumo.Keys()
            
            If CStr(nCampo) Like "VL_*" Then
                
                vCampo = CDbl(CamposResumo(.dicTitulosResumo(CStr(nCampo)))) + CDbl(.Campos(.dicTitulosResumo(CStr(nCampo))))
                .AtribuirValor CStr(nCampo), vCampo
                
            End If
            
        Next nCampo
        
    End With
    
End Function

Public Sub FiltrarRegistros()

Dim CamposOrig As Variant, CamposDest
    
    CamposOrig = Array("CFOP", "CST_ICMS", "ALIQ_ICMS")
    CamposDest = Array("CFOP", "CST_ICMS", "ALIQ_ICMS")
    
    Call Util.FiltrarRegistros(resICMS, assApuracaoICMS, CamposOrig, CamposDest)
    Call Application.GoTo(assApuracaoICMS.[AA3], True)
    
End Sub
