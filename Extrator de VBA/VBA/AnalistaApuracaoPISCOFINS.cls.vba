Attribute VB_Name = "AnalistaApuracaoPISCOFINS"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private dicResumoPISCOFINS As Dictionary
Private Analista As AnalistaApuracao
Private ValidacoesPISCOFINS As AnalistaApuracaoPISCOFINS_Valid

Public Sub GerarResumoApuracaoPISCOFINS()

Dim arrDadosApuracao As New ArrayList
Dim Comeco As Double, VL_OPR#
Dim Campos As Variant, nCampo
Dim Chave As String
    
    Inicio = Now()
    If Util.ChecarAusenciaDados(assApuracaoPISCOFINS, False) Then Exit Sub
        
    Set dicResumoPISCOFINS = New Dictionary
    Set Analista = New AnalistaApuracao
    Set ValidacoesPISCOFINS = New AnalistaApuracaoPISCOFINS_Valid
    
    Set arrDadosApuracao = Util.CriarArrayListRegistro(assApuracaoPISCOFINS)
    Set Analista.dicTitulosApuracao = Util.MapearTitulos(assApuracaoPISCOFINS, 3)
    Set Analista.dicTitulosResumo = Util.MapearTitulos(resPISCOFINS, 3)
    Set Analista.dicTitulos = Analista.dicTitulosResumo
    
    Call ValidacoesPISCOFINS.InicializarObjetos
    
    a = 0
    Comeco = Timer
    With Analista
        
        For Each Campos In arrDadosApuracao
            
            .RedimensionarArray (.dicTitulosResumo.Count)
            Call Util.AntiTravamento(a, 100, "Gerando resumo de apuração do PISCOFINS, por favor aguarde...", arrDadosApuracao.Count, Comeco)
            
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                For Each nCampo In .dicTitulosResumo.Keys()
                    
                    Call MontarRegistroPISCOFINS(Campos, nCampo)
                    
                Next nCampo
                
            End If
            
            .Campos = ValidacoesPISCOFINS.ValidarResumoPISCOFINS(.Campos)
            
            Chave = GerarChaveResumoPISCOFINS(Campos)
            If dicResumoPISCOFINS.Exists(Chave) Then Call AtualizarResumoPISCOFINS(Chave)
            
            dicResumoPISCOFINS(Chave) = .Campos
            
        Next Campos
        
        Call Util.LimparDados(resPISCOFINS, 4, False)
        Call Util.ExportarDadosDicionario(resPISCOFINS, dicResumoPISCOFINS)
        Call FuncoesFormatacao.DestacarInconsistencias(resPISCOFINS)
        
        resPISCOFINS.Activate
        Call Util.MsgInformativa("Resumo PISCOFINS gerado com sucesso!", "Resumo Apuração do PISCOFINS", Inicio)
        
    End With
    
End Sub

Private Function MontarRegistroPISCOFINS(ByVal Campos As Variant, ByVal nCampo As String)

Dim vCampo As Variant
Dim nTitulo As String
    
    With Analista
        
        Select Case nCampo
            
            Case "INCONSISTENCIA", "SUGESTAO"
                vCampo = Empty
                
            Case Else
                vCampo = TratarCamposResumoPISCOFINS(nCampo, Campos)
                
        End Select
        
        .AtribuirValor nCampo, vCampo
        
    End With
    
End Function

Private Function TratarCamposResumoPISCOFINS(ByVal nCampo As String, ByVal Campos As Variant)

Dim vCampo As Variant
    
    With Analista
    
        vCampo = Campos(.dicTitulosApuracao(nCampo))
        
        Select Case True
            
            'Case nCampo = "VL_ITEM"
                'TratarCamposResumoPISCOFINS = CalcularCampoVL_ITEM(Campos)
                
            Case nCampo Like "VL_*"
                TratarCamposResumoPISCOFINS = VBA.Round(vCampo, 2)
            
            Case nCampo Like "*CST*"
                TratarCamposResumoPISCOFINS = fnExcel.FormatarTexto(vCampo)
                
            Case Else
                TratarCamposResumoPISCOFINS = vCampo
                
        End Select
        
    End With
    
End Function

Private Function CalcularCampoVL_ITEM(ByVal Campos As Variant) As Double

Dim VL_ITEM As Double, VL_DESP#, VL_DESC#, VL_ICMS#
    
    With Analista
        
        VL_ITEM = Campos(.dicTitulosApuracao("VL_ITEM"))
        VL_DESP = Campos(.dicTitulosApuracao("VL_DESP"))
        VL_DESC = Campos(.dicTitulosApuracao("VL_DESC"))
        VL_ICMS = Campos(.dicTitulosApuracao("VL_ICMS"))
        
        CalcularCampoVL_ITEM = VL_ITEM + VL_DESP - VL_DESC - VL_ICMS
        
    End With
    
End Function

Private Function GerarChaveResumoPISCOFINS(ByVal Campos As Variant) As String

Dim CamposChave As Variant, Campo
Dim arrCampos As New ArrayList

    CamposChave = Array("CFOP", "CST_PIS", "CST_COFINS", "ALIQ_PIS", "ALIQ_COFINS", "ALIQ_PIS_QUANT", "ALIQ_PIS_QUANT")
    
    With Analista
        
        For Each Campo In CamposChave
            
            arrCampos.Add Campos(.dicTitulosApuracao(Campo))
            
        Next Campo
        
        GerarChaveResumoPISCOFINS = fnSPED.GerarChaveRegistro(VBA.Join(arrCampos.toArray()))
        
    End With
    
End Function

Private Function AtualizarResumoPISCOFINS(ByVal Chave As String)

Dim CamposResumo As Variant, nCampo
Dim vCampo As Double
    
    With Analista
        
        CamposResumo = dicResumoPISCOFINS(Chave)
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
    
    CamposOrig = Array("CFOP", "CST_PIS", "CST_COFINS", "ALIQ_PIS", "ALIQ_COFINS", "ALIQ_PIS_QUANT", "ALIQ_PIS_QUANT")
    CamposDest = Array("CFOP", "CST_PIS", "CST_COFINS", "ALIQ_PIS", "ALIQ_COFINS", "ALIQ_PIS_QUANT", "ALIQ_PIS_QUANT")
    
    Call Util.FiltrarRegistros(resPISCOFINS, assApuracaoPISCOFINS, CamposOrig, CamposDest)
    Call Application.GoTo(assApuracaoPISCOFINS.[AA3], True)
    
End Sub

