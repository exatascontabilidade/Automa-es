Attribute VB_Name = "clsValidacoesSPED"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Fiscal As New clsValidacoesSPEDFiscal
Public Contribuicoes As New clsValidacoesSPEDContribuicoes

Public Function ValidarValores(ByVal Valor As Variant) As Double
    If Valor = "" Then ValidarValores = 0 Else ValidarValores = Valor
End Function

Public Function FormatarValores(ByVal Valor As Variant) As Double
    
    If Valor = "" Then Valor = "#0.00"
    Valor = VBA.Replace(Valor, ",", ".")
    FormatarValores = Valor
    
End Function

Public Function ValidarMotivoInventario(ByVal codInv As Variant) As String
    
    Select Case codInv
    
        Case "01"
            ValidarMotivoInventario = "01 - No final no período"
            
        Case "02"
            ValidarMotivoInventario = "02 - Na mudança de forma de tributação da mercadoria (ICMS)"
            
        Case "03"
            ValidarMotivoInventario = "03 - Na solicitação da baixa cadastral, paralisação temporária e outras situações"
            
        Case "04"
            ValidarMotivoInventario = "04 - Na alteração de regime de pagamento – condição do contribuinte"
            
        Case "05"
            ValidarMotivoInventario = "05 - Por determinação dos fiscos"
            
        Case "06"
            ValidarMotivoInventario = "06 - Para controle das mercadorias sujeitas ao regime de substituição tributária – restituição/ ressarcimento/ complementação"
            
        Case Is <> ""
            ValidarMotivoInventario = codInv & " - Código Inválido"
            
    End Select
        
End Function

Public Function ValidarPropriedade(ByVal codProp As String)
    
    Select Case Replace(codProp, "'", "")
        
        Case "0"
            ValidarPropriedade = "0 - Item de propriedade do informante e em seu poder"
            
        Case "1"
            ValidarPropriedade = "1 - Item de propriedade do informante em posse de terceiros"
            
        Case "2"
            ValidarPropriedade = "2 - Item de propriedade de terceiros em posse do informante"
            
        Case Else
            ValidarPropriedade = codProp & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarSituacaoEspecial(ByVal IND_SIT_ESP As String)
    
    Select Case IND_SIT_ESP
        
        Case "0"
            ValidarSituacaoEspecial = "0 - Abertura"
            
        Case "1"
            ValidarSituacaoEspecial = "1 - Cisão"
            
        Case "2"
            ValidarSituacaoEspecial = "2 - Fusão"
            
        Case "3"
            ValidarSituacaoEspecial = "3 - Incorporação"
            
        Case "4"
            ValidarSituacaoEspecial = "4 - Encerramento"
            
        Case Else
            ValidarSituacaoEspecial = IND_SIT_ESP & " - Código Inválido"
            
    End Select
    
End Function

Public Function ValidarEnumeracao_COD_FIN(ByVal tpEsc As String)
    
    Select Case Replace(tpEsc, "'", "")
        
        Case "0"
            ValidarEnumeracao_COD_FIN = "0 - Original"
            
        Case "1"
            ValidarEnumeracao_COD_FIN = "1 - Retificadora"
            
        Case Else
            ValidarEnumeracao_COD_FIN = tpEsc & " - Código Inválido"
            
    End Select
    
End Function


Public Function ValidarNaturezaJuridica(ByVal indNat As String)
            
        Select Case Replace(indNat, "'", "")
        
            Case "00"
                ValidarNaturezaJuridica = "00 - Pessoa jurídica em geral"
                
            Case "01"
                ValidarNaturezaJuridica = "01 - Sociedade cooperativa"
                
            Case "02"
                ValidarNaturezaJuridica = "02 - Entidade sujeita ao PIS/Pasep exclusivamente com base na Folha de Salários"
                
            Case "03"
                ValidarNaturezaJuridica = "03 - Pessoa jurídica em geral participante de SCP como sócia ostensiva"
                
            Case "04"
                ValidarNaturezaJuridica = "04 - Sociedade cooperativa participante de SCP como sócia ostensiva"
                
            Case "05"
                ValidarNaturezaJuridica = "05 - Sociedade em Conta de Participação - SCP"
                
            Case Else
                ValidarNaturezaJuridica = indNat & " - Código Inválido"
                
        End Select
    
End Function

Public Function ValidarTipoTitulo(ByVal tpTitulo As String)
            
        Select Case Replace(tpTitulo, "'", "")
        
            Case "00"
                ValidarTipoTitulo = "00 - Duplicata"
                
            Case "01"
                ValidarTipoTitulo = "01 - Cheque"
                
            Case "02"
                ValidarTipoTitulo = "02 - Promissória"
                
            Case "03"
                ValidarTipoTitulo = "03 - Recibo"
                
            Case "99"
                ValidarTipoTitulo = "99 - Outros"
            
            Case Else
                ValidarTipoTitulo = tpTitulo & " - Código Inválido"
                
        End Select
    
End Function

Public Function ValidarPeriodoArquivo() As Boolean
    
Dim dicDados0000 As New Dictionary
Dim Periodo As String, ARQUIVO$, Msg$

    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
    
    CNPJContribuinte = Util.FormatarCNPJ(CadContrib.Range("CNPJContribuinte").value)
    CNPJBase = VBA.Left(CNPJContribuinte, 8)
    Periodo = VBA.Format(PeriodoEspecifico, "00/0000")
            
    ARQUIVO = Periodo & "-" & CNPJContribuinte
    If dicDados0000.Exists(ARQUIVO) Then ValidarPeriodoArquivo = True

End Function

