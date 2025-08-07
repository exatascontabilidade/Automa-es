Attribute VB_Name = "clsRegrasFiscaisNCM"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private TabelasFiscais As New AtualizadorTabelasSPED

Private Sub Class_Initialize()

Dim dicDadosNCM As New Dictionary
Dim CustomPart As New clsCustomPartXML

    If TabelaNCM.Count = 0 Then
        
        Call Util.AtualizarBarraStatus("Carregando informações da tabela NCM, por favor aguarde...")
        Set dicDadosNCM = JsonConverter.ParseJson(CustomPart.ExtrairJsonXmlPart("TabelaNCM"))
        Call CarregarTabelaNCM(dicDadosNCM)
        
    End If
    
End Sub

Private Function CarregarTabelaNCM(ByRef dicDadosNCM As Dictionary)

Dim Campos As Variant
    
    With CamposNCM
        
        For Each Campos In dicDadosNCM("Nomenclaturas")
            
            .COD_NCM = Util.ApenasNumeros(Campos("Codigo"))
            .DESCRICAO = Campos("Descricao")
            .VIGENCIA_INICIAL = fnExcel.FormatarData(Campos("Data_Inicio"))
            .VIGENCIA_FINAL = fnExcel.FormatarData(Campos("Data_Fim"))
            
            TabelaNCM(.COD_NCM) = Array(.DESCRICAO, .VIGENCIA_INICIAL, .VIGENCIA_FINAL)
            
        Next Campos
        
    End With
    
End Function

Public Function verificarNCM(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim COD_NCM As String, TIPO_ITEM$
Dim i As Integer
    
    If LBound(Campos) = 0 Then i = 1
    
    COD_NCM = fnSPED.FormatarNCM(Util.ApenasNumeros(Campos(dicTitulos("COD_NCM") - i)))
    TIPO_ITEM = Util.FormatarValores(Util.ApenasNumeros(Campos(dicTitulos("TIPO_ITEM") - i)))
    
    Select Case True
        
        Case COD_NCM = "" And TIPO_ITEM < 7
            Call Util.GravarSugestao(Campos, dicTitulos, _
                INCONSISTENCIA:="O campo COD_NCM não foi informado", _
                SUGESTAO:="informar um valor válido para o campo COD_NCM", _
                dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
                
        Case Not CarregarDadosNCM(COD_NCM) And TIPO_ITEM < 7
            Call Util.GravarSugestao(Campos, dicTitulos, _
                INCONSISTENCIA:="O NCM informado é inválido", _
                SUGESTAO:="Informar um valor válido para o campo COD_NCM", _
                dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
                
    End Select
    
End Function

Public Function CarregarDadosNCM(ByVal NCM As String) As Boolean

Dim Campos As Variant
    
    If NCM <> "" And TabelaNCM.Exists(NCM) Then
        
        Campos = TabelaNCM(NCM)
        Call DadosValidacaoNCM.CarregarDadosTabelaNCM(Campos)
        CarregarDadosNCM = True
        
    End If
    
End Function

Public Function ValidarCampo_COD_NCM(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim i As Integer
Dim ExisteNCM As Boolean
Dim TamanhoNCM As Boolean
    
    Call DadosValidacaoNCM.ResetarCamposNCM
    Call DadosValidacaoNCM.CarregarDadosApuracaoNCM(Campos, ActiveSheet)
    Set DadosValidacaoNCM.dicTitulosApuracao = dicTitulos
    
    With CamposNCM
            
        TamanhoNCM = VBA.Len(.COD_NCM) = 8
        ExisteNCM = CarregarDadosNCM(.COD_NCM)
        
        Select Case True
            
            Case Not TamanhoNCM And Not Util.VerificarStringVazia(.COD_NCM)
                .INCONSISTENCIA = "O campo COD_NCM precisa ter 8 dígitos"
                .SUGESTAO = "Adicionar zeros a esquerda do campo COD_NCM"
                
            Case Not ExisteNCM And Not Util.VerificarStringVazia(.COD_NCM)
                .INCONSISTENCIA = "O NCM (" & .COD_NCM & ") não existe na tabela NCM vigente"
                .SUGESTAO = "Apagar valor informado no campo COD_NCM"
                
            Case ExisteNCM
                Call ValidarNCM
                
        End Select
        
        If .INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, INCONSISTENCIA:=.INCONSISTENCIA, _
            SUGESTAO:=.SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
    End With
    
End Function

Private Function ValidarNCM()
    
    With CamposNCM
        
        Call DefinirDataReferencia
        
        Select Case True
            
            Case .DT_REF < .VIGENCIA_INICIAL
                .INCONSISTENCIA = "O NCM (" & .COD_NCM & ") é inválido para a data da operação (" & .DT_REF & "). A vigência deste código inicia em " & .VIGENCIA_INICIAL
                .SUGESTAO = "Informe um código NCM válido para a data da operação. Consulte a tabela de NCMs vigentes, se necessário."
                
            Case .DT_REF > .VIGENCIA_FINAL
                .INCONSISTENCIA = "O NCM (" & .COD_NCM & ") é inválido para a data da operação (" & .DT_REF & "). A vigência deste código expirou em " & .VIGENCIA_FINAL
                .SUGESTAO = "Informe um código NCM válido para a data da operação. Consulte a tabela de NCMs vigentes, se necessário."
                
        End Select
        
    End With
    
End Function

Private Function DefinirDataReferencia()
    
    With CamposNCM
        
        Select Case True
            
            Case .IND_OPER Like "*Entrada"
                .DT_REF = .DT_ENT_SAI
                
            Case .IND_OPER Like "*Saida"
                .DT_REF = .DT_DOC
                
        End Select
        
    End With
    
End Function

Public Sub BaixarTabelaNCM()
    
    Call TabelasFiscais.BaixarTabela(UrlTabelaNCM, "TabelaNCM")
    
End Sub
