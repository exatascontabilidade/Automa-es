Attribute VB_Name = "clsRegrasDivergenciasNotas"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Jaccard As New clsSimilaridadeJaccard
Private arrFuncoesValidacao As New ArrayList
Private dicTitulosRelatorio As New Dictionary
Private CamposRelatorio As Variant

Private Sub Class_Initialize()
    
    Set dicTitulosRelatorio = Util.MapearTitulos(relDivergenciasNotas, 3)
    Call DadosSPEDFiscal.CarregarDadosRegistro0000
    Call CarregarValidacoes
    
End Sub

Private Sub CarregarValidacoes()
    
    With arrFuncoesValidacao
        
        .Clear
        
        .Add "ValidarCampo_COD_MOD"
        .Add "ValidarCampo_NUM_DOC"
        .Add "ValidarCampo_SER"
        .Add "ValidarCampo_COD_PART"
        .Add "ValidarCampo_NOME_RAZAO"
        .Add "ValidarCampo_INSC_EST"
        .Add "ValidarCampo_VL_FRT"
        .Add "ValidarCampo_VL_SEG"
        .Add "ValidarCampo_VL_OUT_DA"
        .Add "ValidarCampo_VL_DOC"
        .Add "ValidarCampo_VL_BC_ICMS"
        .Add "ValidarCampo_VL_ICMS"
        .Add "ValidarCampo_VL_BC_ICMS_ST"
        .Add "ValidarCampo_VL_ICMS_ST"
        .Add "ValidarCampo_VL_IPI"
        .Add "ValidarCampo_VL_DESC"
        .Add "ValidarCampo_VL_ABAT_NT"
        .Add "ValidarCampo_VL_MERC"
        .Add "ValidarCampo_IND_OPER"
        .Add "ValidarCampo_IND_EMIT"
        .Add "ValidarCampo_DT_DOC"
        .Add "ValidarCampo_DT_E_S"
        .Add "ValidarCampo_IND_PGTO"
        .Add "ValidarCampo_IND_FRT"
        
    End With
    
End Sub

Public Function IdentificarDivergenciasNotas(ByRef Registros As ArrayList)

Dim i As Long
    
    For i = 0 To Registros.Count - 1
        
        Call DadosDivergenciasNotas.ResetarCamposNota
        Call DadosDivergenciasNotas.CarregarDadosRegistroDivergenciaNotas(Registros(i))
        
        CamposRelatorio = Registros(i)
            
            Call IdentificarDivergenciaNota
            
        Registros(i) = CamposRelatorio
        
    Next i
    
End Function

Private Sub IdentificarDivergenciaNota()

Dim Validacao As Variant
    
    With CamposNota
        
        Call ValidarCampo_COD_SIT
        
        For Each Validacao In arrFuncoesValidacao
            
            If .INCONSISTENCIA <> "" Then
                
                Call Util.GravarSugestao(CamposRelatorio, dicTitulosRelatorio, .INCONSISTENCIA, .SUGESTAO, dicInconsistenciasIgnoradas)
                Exit Sub
                
            End If
            
            If Not .COD_SIT_SPED Like "*02*" Then CallByName Me, CStr(Validacao), VbMethod
            
        Next Validacao
        
        If .INCONSISTENCIA <> "" Then _
            Call Util.GravarSugestao(CamposRelatorio, dicTitulosRelatorio, .INCONSISTENCIA, .SUGESTAO, dicInconsistenciasIgnoradas)
            
    End With
    
End Sub

Public Function ValidarCampo_COD_MOD()
    
    With CamposNota
        
        Select Case True
            
            Case .COD_MOD_NF <> .COD_MOD_SPED And .COD_MOD_NF <> ""
                .INCONSISTENCIA = "Os campos COD_MOD_NF (" & .COD_MOD_NF & ") e COD_MOD_SPED (" & .COD_MOD_SPED & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo modelo do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_COD_SIT()
    
    With CamposNota
        
        Select Case True
            
            Case .COD_SIT_NF <> .COD_SIT_SPED And .COD_SIT_NF <> ""
                .INCONSISTENCIA = "Os campos COD_SIT_NF (" & .COD_SIT_NF & ") e COD_SIT_SPED (" & .COD_SIT_SPED & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo código de situação do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_NUM_DOC()
    
    With CamposNota
        
        Select Case True
            
            Case .NUM_DOC_NF <> .NUM_DOC_SPED And .NUM_DOC_NF <> ""
                .INCONSISTENCIA = "Os campos NUM_DOC_NF (" & .NUM_DOC_NF & ") e NUM_DOC_SPED (" & .NUM_DOC_SPED & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo número de documento do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_SER()
    
    With CamposNota
        
        Select Case True
            
            Case .SER_NF <> .SER_SPED And .SER_NF <> ""
                .INCONSISTENCIA = "Os campos SER_NF (" & .SER_NF & ") e SER_SPED (" & .SER_SPED & ") estão divergentes"
                .SUGESTAO = "Informar a mesma série do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_IND_OPER()
    
    With CamposNota
        
        Select Case True
            
            Case .IND_OPER_NF <> .IND_OPER_SPED And .IND_OPER_NF <> ""
                .INCONSISTENCIA = "Os campos IND_OPER_NF (" & .IND_OPER_NF & ") e IND_OPER_SPED (" & .IND_OPER_SPED & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo tipo de operação do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_IND_EMIT()
    
    With CamposNota
        
        Select Case True
            
            Case .IND_EMIT_NF <> .IND_EMIT_SPED And .IND_EMIT_NF <> ""
                .INCONSISTENCIA = "Os campos IND_EMIT_NF (" & .IND_EMIT_NF & ") e IND_EMIT_SPED (" & .IND_EMIT_SPED & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo tipo emissão do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_COD_PART()
    
    With CamposNota
        
        If .COD_MOD_NF <> "65" And Not .COD_SIT_SPED Like "*02*" Then
            
            Select Case True
                
                Case .COD_PART_NF <> .COD_PART_SPED And .COD_PART_NF <> ""
                    .INCONSISTENCIA = "Os campos COD_PART_NF (" & .COD_PART_NF & ") e COD_PART_SPED (" & .COD_PART_SPED & ") estão divergentes"
                    .SUGESTAO = "Informar o mesmo participante do XML para o SPED"
                    
            End Select
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_NOME_RAZAO()
    
    With CamposNota
        
        If .COD_MOD_NF <> "65" And Not .COD_SIT_SPED Like "*02*" Then
            
            Select Case True
                
                Case Jaccard.CalcularSimilaridadeJaccard(.NOME_RAZAO_NF, .NOME_RAZAO_SPED) < 0.8
                    .INCONSISTENCIA = "Os campos NOME_RAZAO_NF (" & .NOME_RAZAO_NF & ") e NOME_RAZAO_SPED (" & .NOME_RAZAO_SPED & ") estão divergentes"
                    .SUGESTAO = "Informar a mesma razão do participante do XML para o SPED"
                    
            End Select
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_INSC_EST()
    
    With CamposNota
        
        If .COD_MOD_NF <> "65" And Not .COD_SIT_SPED Like "*02*" And .INSC_EST_NF <> "" Then
            
            If .INSC_EST_SPED = "" Then .INSC_EST_SPED = "VAZIO"
            
            Select Case True
                
                Case .INSC_EST_SPED = ""
                    .INCONSISTENCIA = "A inscrição estadual do participante não foi informada"
                    .SUGESTAO = "Informar a inscrição estadual do XML para o SPED"
                    
                Case fnExcel.ConverterValores(Util.ApenasNumeros(.INSC_EST_NF), True, 0) <> fnExcel.ConverterValores(Util.ApenasNumeros(.INSC_EST_SPED), True, 0)
                    .INCONSISTENCIA = "Os campos INSC_EST_NF (" & .INSC_EST_NF & ") e INSC_EST_SPED (" & .INSC_EST_SPED & ") estão divergentes"
                    .SUGESTAO = "Informar a mesma inscrição estadual do XML para o SPED"
                    
            End Select
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_DT_DOC()
    
    With CamposNota
        
        If .DT_DOC_NF <> "" And .DT_DOC_SPED <> "" Then
            
            Select Case True
                
                Case CDate(.DT_DOC_NF) <> CDate(.DT_DOC_SPED)
                    .INCONSISTENCIA = "Os campos DT_DOC_NF (" & .DT_DOC_NF & ") e DT_DOC_SPED (" & .DT_DOC_SPED & ") estão divergentes"
                    .SUGESTAO = "Informar a mesma data de emissão do XML para o SPED"
                    
            End Select
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_DT_E_S()
    
    With CamposNota
        
        If .DT_E_S_NF <> "" And .DT_E_S_SPED <> "" Then
            
            Select Case True
                
                Case CDate(.DT_E_S_NF) > CDate(.DT_E_S_SPED)
                    .INCONSISTENCIA = "Os campos DT_E_S_NF (" & .DT_E_S_NF & ") e DT_E_S_SPED (" & .DT_E_S_SPED & ") estão divergentes"
                    .SUGESTAO = "Informar a mesma data de entrada/saída do XML para o SPED"
                    
            End Select
        
        End If
        
    End With
    
End Function

Public Function ValidarCampo_IND_PGTO()
    
    With CamposNota
        
        Select Case True
            
            Case .IND_PGTO_NF <> .IND_PGTO_SPED And .IND_PGTO_NF <> ""
                .INCONSISTENCIA = "Os campos IND_PGTO_NF (" & .IND_PGTO_NF & ") e IND_PGTO_SPED (" & .IND_PGTO_SPED & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo tipo de pagamento do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_IND_FRT()
    
    With CamposNota
        
        Select Case True
            
            Case .IND_FRT_NF <> .IND_FRT_SPED And .IND_FRT_NF <> ""
                .INCONSISTENCIA = "Os campos IND_FRT_NF (" & .IND_FRT_NF & ") e IND_FRT_SPED (" & .IND_FRT_SPED & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo tipo de frete do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_DOC()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_DOC_NF > 0 And .VL_DOC_SPED = 0 And .COD_SIT_SPED Like "*08*"
                .INCONSISTENCIA = "Operação especial (COD_SIT_SPED = " & .COD_SIT_SPED & ") com valor do documento zerado (VL_DOC_SPED: R$ " & VBA.Round(.VL_DOC_SPED, 2) & ")"
                .SUGESTAO = "Informar o mesmo valor de documento do XML para o SPED"
                
            Case .VL_DOC_NF <> .VL_DOC_SPED And .VL_DOC_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_DOC_NF (R$ " & VBA.Round(.VL_DOC_NF, 2) & ") e VL_DOC_SPED (R$ " & VBA.Round(.VL_DOC_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor de documento do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_DESC()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_DESC_NF <> .VL_DESC_SPED And .VL_DESC_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_DESC_NF (R$ " & VBA.Round(.VL_DESC_NF, 2) & ") e VL_DESC_SPED (R$ " & VBA.Round(.VL_DESC_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor de desconto do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_ABAT_NT()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_ABAT_NT_NF <> .VL_ABAT_NT_SPED And .VL_ABAT_NT_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_ABAT_NT_NF (R$ " & VBA.Round(.VL_ABAT_NT_NF, 2) & ") e VL_ABAT_NT_SPED (R$ " & VBA.Round(.VL_ABAT_NT_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor de abatimento do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_MERC()

Dim VL_MERC_NF As Double
Dim VL_MERC_SPED As Double
    
    With CamposNota
        
        VL_MERC_NF = VBA.Round(.VL_IPI_NF + .VL_ICMS_ST_NF + .VL_MERC_NF, 2)
        VL_MERC_SPED = VBA.Round(.VL_IPI_SPED + .VL_ICMS_ST_SPED + .VL_MERC_SPED, 2)
        
        Select Case True
            
            Case .VL_MERC_NF <> .VL_MERC_SPED
                .INCONSISTENCIA = "Os valores dos campos VL_MERC_NF (R$ " & VBA.Round(.VL_MERC_NF, 2) & ") e VL_MERC_SPED (R$ " & VBA.Round(.VL_MERC_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor de mercadorias do XML para o SPED"

'            Case Not VerificarPossibilidadesValidasCampo_VL_MERC(VL_MERC_SPED)
'                .INCONSISTENCIA = "Os valores dos campos VL_MERC_NF (R$ " & VBA.Round(.VL_MERC_NF, 2) & ") e VL_MERC_SPED (R$ " & VBA.Round(.VL_MERC_SPED, 2) & ") estão divergentes"
'                .SUGESTAO = "Informar o mesmo valor de mercadorias do XML para o SPED"
'
'            Case VerificarPossibilidadesInvalidasCampo_VL_MERC(VL_MERC_SPED)
'                .INCONSISTENCIA = "Os valores dos campos VL_MERC_NF (R$ " & VBA.Round(.VL_MERC_NF, 2) & ") e VL_MERC_SPED (R$ " & VBA.Round(.VL_MERC_SPED, 2) & ") estão divergentes"
'                .SUGESTAO = "Informar o mesmo valor de mercadorias do XML para o SPED"
                

                
        End Select
        
    End With
    
End Function

Private Function VerificarPossibilidadesValidasCampo_VL_MERC(ByVal VL_MERC_SPED As Double) As Boolean
    
Dim Valor As Variant

    With CamposNota
        
        For Each Valor In CarregarPossibilidadesValidas
            
            If Valor = VL_MERC_SPED Then
                
                VerificarPossibilidadesValidasCampo_VL_MERC = True
                Exit Function
                
            End If
            
        Next Valor
        
    End With
    
End Function

Private Function CarregarPossibilidadesValidas() As Variant
    
    With CamposNota
        
        CarregarPossibilidadesValidas = Array( _
            CDbl(VBA.Round(.VL_MERC_NF, 2)), _
            CDbl(VBA.Round(.VL_MERC_NF + .VL_IPI_NF, 2)), _
            CDbl(VBA.Round(.VL_MERC_NF + .VL_ICMS_ST_NF, 2)), _
            CDbl(VBA.Round(.VL_MERC_NF + .VL_ICMS_ST_NF + .VL_IPI_NF, 2)))
            
    End With
    
End Function

Private Function VerificarPossibilidadesInvalidasCampo_VL_MERC(ByVal VL_MERC_SPED As Double) As Boolean

Dim Valor As Double
Dim Chave As Variant
Dim dicPossibilidades As New Dictionary
    
    With CamposNota
        
        Set dicPossibilidades = CarregarPossibilidadesInvalidas
        
        For Each Chave In dicPossibilidades.Keys()
            
            Valor = dicPossibilidades(Chave)(0)
            If Valor = VL_MERC_SPED Then
                
                .INCONSISTENCIA = dicPossibilidades(Chave)(1)
                .SUGESTAO = dicPossibilidades(Chave)(2)
                
                VerificarPossibilidadesInvalidasCampo_VL_MERC = True
                Exit Function
                
            End If
            
        Next Chave
        
    End With
    
End Function

Private Function CarregarPossibilidadesInvalidas_Old() As Dictionary

Dim dicPossibilidadesInvalidas As New Dictionary
    
    With CamposNota
        
        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF, 2), "O valor do campo VL_FRT_NF foi somado ao valor do campo VL_MERC_SPED", "Informar o mesmo valor de mercadorias do XML para o SPED"), "VL_FRT_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF, 2), "VL_OUT_DA_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_SEG_NF, 2), "VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF, 2), "VL_FRT_NF, VL_OUT_DA_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_SEG_NF, 2), "VL_FRT_NF, VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF + .VL_SEG_NF, 2), "VL_OUT_DA_NF, VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF + .VL_SEG_NF, 2), "VL_FRT_NF, VL_OUT_DA_NF, VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_IPI_NF, 2), "VL_FRT_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF + .VL_IPI_NF, 2), "VL_OUT_DA_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_SEG_NF + .VL_IPI_NF, 2), "VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF + .VL_IPI_NF, 2), "VL_FRT_NF, VL_OUT_DA_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_SEG_NF + .VL_IPI_NF, 2), "VL_FRT_NF, VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF + .VL_SEG_NF + .VL_IPI_NF, 2), "VL_OUT_DA_NF, VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF + .VL_SEG_NF + .VL_IPI_NF, 2), "VL_FRT_NF, VL_OUT_DA_NF, VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_ICMS_ST_NF, 2), "VL_FRT_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF + .VL_ICMS_ST_NF, 2), "VL_OUT_DA_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_SEG_NF + .VL_ICMS_ST_NF, 2), "VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF + .VL_ICMS_ST_NF, 2), "VL_FRT_NF, VL_OUT_DA_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_SEG_NF + .VL_ICMS_ST_NF, 2), "VL_FRT_NF, VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF + .VL_SEG_NF + .VL_ICMS_ST_NF, 2), "VL_OUT_DA_NF, VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF + .VL_SEG_NF + .VL_ICMS_ST_NF, 2), "VL_FRT_NF, VL_OUT_DA_NF, VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_IPI_NF + .VL_ICMS_ST_NF, 2), "VL_FRT_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF + .VL_IPI_NF + .VL_ICMS_ST_NF, 2), "VL_OUT_DA_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_SEG_NF + .VL_IPI_NF + .VL_ICMS_ST_NF, 2), "VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF + .VL_IPI_NF + .VL_ICMS_ST_NF, 2), "VL_FRT_NF, VL_OUT_DA_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_SEG_NF + .VL_IPI_NF + .VL_ICMS_ST_NF, 2), "VL_FRT_NF, VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF + .VL_SEG_NF + .VL_IPI_NF + .VL_ICMS_ST_NF, 2), "VL_OUT_DA_NF, VL_SEG_NF"
        dicPossibilidadesInvalidas.Add VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF + .VL_SEG_NF + .VL_IPI_NF + .VL_ICMS_ST_NF, 2), "VL_FRT_NF, VL_OUT_DA_NF, VL_SEG_NF"
        
    End With
    
    Set CarregarPossibilidadesInvalidas_Old = dicPossibilidadesInvalidas
    
End Function

Private Function CarregarPossibilidadesInvalidas() As Dictionary

Dim dicPossibilidadesInvalidas As New Dictionary
    
    With CamposNota
        
        ' Combinações com VL_FRT_NF
        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF, 2), _
                                            "O valor do campo VL_FRT_NF foi somado ao valor do campo VL_MERC_SPED", _
                                            "Informar o mesmo valor de mercadorias do XML para o SPED"), _
                                            "VL_FRT_NF"

        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF, 2), _
                                            "Os valores dos campos VL_FRT_NF e VL_OUT_DA_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de frete, outras despesas e mercadorias no XML e ajuste o SPED"), _
                                            "VL_FRT_NF_OUT_DA_NF" ' Chave única

        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_SEG_NF, 2), _
                                             "Os valores dos campos VL_FRT_NF e VL_SEG_NF foram somados ao valor do campo VL_MERC_SPED", _
                                             "Verifique os valores de frete, seguro e mercadorias no XML e ajuste o SPED"), _
                                             "VL_FRT_NF_SEG_NF" ' Chave única
        
        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_IPI_NF, 2), _
                                           "Os valores dos campos VL_FRT_NF e VL_IPI_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de frete, IPI e mercadorias no XML e ajuste o SPED"), _
                                            "VL_FRT_NF_IPI_NF" ' Chave única


        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_ICMS_ST_NF, 2), _
                                            "Os valores dos campos VL_FRT_NF e VL_ICMS_ST_NF foram somados ao valor do campo VL_MERC_SPED", _
                                             "Verifique os valores de frete, ICMS-ST e mercadorias no XML e ajuste o SPED"), _
                                             "VL_FRT_NF_ICMS_ST_NF" ' Chave única
        
        ' Combinações com VL_OUT_DA_NF
        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF, 2), _
                                            "O valor do campo VL_OUT_DA_NF foi somado ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de outras despesas acessórias e mercadorias no XML e ajuste o SPED"), _
                                            "VL_OUT_DA_NF"


        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF + .VL_SEG_NF, 2), _
                                           "Os valores dos campos VL_OUT_DA_NF e VL_SEG_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de outras despesas, seguro e mercadorias no XML e ajuste o SPED"), _
                                            "VL_OUT_DA_NF_SEG_NF" ' Chave única


        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF + .VL_IPI_NF, 2), _
                                             "Os valores dos campos VL_OUT_DA_NF e VL_IPI_NF foram somados ao valor do campo VL_MERC_SPED", _
                                             "Verifique os valores de outras despesas, IPI e mercadorias no XML e ajuste o SPED"), _
                                             "VL_OUT_DA_NF_IPI_NF" ' Chave única

         dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF + .VL_ICMS_ST_NF, 2), _
                                            "Os valores dos campos VL_OUT_DA_NF e VL_ICMS_ST_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de outras despesas, ICMS-ST e mercadorias no XML e ajuste o SPED"), _
                                            "VL_OUT_DA_NF_ICMS_ST_NF" ' Chave única

         ' Combinações com VL_SEG_NF
         dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_SEG_NF, 2), _
                                            "O valor do campo VL_SEG_NF foi somado ao valor do campo VL_MERC_SPED", _
                                             "Verifique os valores de seguro e mercadorias no XML e ajuste o SPED"), _
                                             "VL_SEG_NF"

        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_SEG_NF + .VL_IPI_NF, 2), _
                                             "Os valores dos campos VL_SEG_NF e VL_IPI_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de seguro, IPI e mercadorias no XML e ajuste o SPED"), _
                                            "VL_SEG_NF_IPI_NF" ' Chave única

        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_SEG_NF + .VL_ICMS_ST_NF, 2), _
                                            "Os valores dos campos VL_SEG_NF e VL_ICMS_ST_NF foram somados ao valor do campo VL_MERC_SPED", _
                                             "Verifique os valores de seguro, ICMS-ST e mercadorias no XML e ajuste o SPED"), _
                                            "VL_SEG_NF_ICMS_ST_NF" ' Chave única

       ' Combinações com VL_FRT_NF, VL_OUT_DA_NF, VL_SEG_NF
        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF + .VL_SEG_NF, 2), _
                                            "Os valores dos campos VL_FRT_NF, VL_OUT_DA_NF e VL_SEG_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de frete, outras despesas, seguro e mercadorias no XML e ajuste o SPED"), _
                                             "VL_FRT_NF_OUT_DA_NF_SEG_NF" ' Chave única

        ' Combinações com VL_FRT_NF, VL_OUT_DA_NF, VL_IPI_NF
        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF + .VL_IPI_NF, 2), _
                                             "Os valores dos campos VL_FRT_NF, VL_OUT_DA_NF e VL_IPI_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de frete, outras despesas, IPI e mercadorias no XML e ajuste o SPED"), _
                                             "VL_FRT_NF_OUT_DA_NF_IPI_NF" ' Chave única
        ' Combinações com VL_FRT_NF, VL_OUT_DA_NF, VL_ICMS_ST_NF
        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF + .VL_ICMS_ST_NF, 2), _
                                             "Os valores dos campos VL_FRT_NF, VL_OUT_DA_NF e VL_ICMS_ST_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de frete, outras despesas, ICMS ST e mercadorias no XML e ajuste o SPED"), _
                                             "VL_FRT_NF_OUT_DA_NF_ICMS_ST_NF" ' Chave única

         ' Combinações com VL_FRT_NF, VL_SEG_NF, VL_IPI_NF
        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_SEG_NF + .VL_IPI_NF, 2), _
                                             "Os valores dos campos VL_FRT_NF, VL_SEG_NF e VL_IPI_NF foram somados ao valor do campo VL_MERC_SPED", _
                                             "Verifique os valores de frete, seguro, IPI e mercadorias no XML e ajuste o SPED"), _
                                             "VL_FRT_NF_SEG_NF_IPI_NF" ' Chave única
        
        ' Combinações com VL_FRT_NF, VL_SEG_NF, VL_ICMS_ST_NF
        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_SEG_NF + .VL_ICMS_ST_NF, 2), _
                                             "Os valores dos campos VL_FRT_NF, VL_SEG_NF e VL_ICMS_ST_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de frete, seguro, ICMS ST e mercadorias no XML e ajuste o SPED"), _
                                            "VL_FRT_NF_SEG_NF_ICMS_ST_NF" ' Chave única
        
         ' Combinações com VL_OUT_DA_NF, VL_SEG_NF, VL_IPI_NF
        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF + .VL_SEG_NF + .VL_IPI_NF, 2), _
                                             "Os valores dos campos VL_OUT_DA_NF, VL_SEG_NF e VL_IPI_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de outras despesas, seguro, IPI e mercadorias no XML e ajuste o SPED"), _
                                             "VL_OUT_DA_NF_SEG_NF_IPI_NF" ' Chave única
       
        ' Combinações com VL_OUT_DA_NF, VL_SEG_NF, VL_ICMS_ST_NF
        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF + .VL_SEG_NF + .VL_ICMS_ST_NF, 2), _
                                             "Os valores dos campos VL_OUT_DA_NF, VL_SEG_NF e VL_ICMS_ST_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de outras despesas, seguro, ICMS ST e mercadorias no XML e ajuste o SPED"), _
                                            "VL_OUT_DA_NF_SEG_NF_ICMS_ST_NF" ' Chave única
        
        ' Combinações com VL_FRT_NF, VL_OUT_DA_NF, VL_SEG_NF, VL_IPI_NF
         dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF + .VL_SEG_NF + .VL_IPI_NF, 2), _
                                            "Os valores dos campos VL_FRT_NF, VL_OUT_DA_NF, VL_SEG_NF e VL_IPI_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de frete, outras despesas, seguro, IPI e mercadorias no XML e ajuste o SPED"), _
                                            "VL_FRT_NF_OUT_DA_NF_SEG_NF_IPI_NF" ' Chave única

         ' Combinações com VL_FRT_NF, VL_OUT_DA_NF, VL_SEG_NF, VL_ICMS_ST_NF
         dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF + .VL_SEG_NF + .VL_ICMS_ST_NF, 2), _
                                             "Os valores dos campos VL_FRT_NF, VL_OUT_DA_NF, VL_SEG_NF e VL_ICMS_ST_NF foram somados ao valor do campo VL_MERC_SPED", _
                                             "Verifique os valores de frete, outras despesas, seguro, ICMS ST e mercadorias no XML e ajuste o SPED"), _
                                              "VL_FRT_NF_OUT_DA_NF_SEG_NF_ICMS_ST_NF" ' Chave única
        
          ' Combinações com VL_FRT_NF, VL_IPI_NF, VL_ICMS_ST_NF
        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_IPI_NF + .VL_ICMS_ST_NF, 2), _
                                             "Os valores dos campos VL_FRT_NF,  VL_IPI_NF e VL_ICMS_ST_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de frete, IPI, ICMS ST e mercadorias no XML e ajuste o SPED"), _
                                             "VL_FRT_NF_IPI_NF_ICMS_ST_NF" ' Chave única

        ' Combinações com VL_OUT_DA_NF, VL_IPI_NF, VL_ICMS_ST_NF
         dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_OUT_DA_NF + .VL_IPI_NF + .VL_ICMS_ST_NF, 2), _
                                             "Os valores dos campos VL_OUT_DA_NF,  VL_IPI_NF e VL_ICMS_ST_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de outras despesas, IPI, ICMS ST e mercadorias no XML e ajuste o SPED"), _
                                            "VL_OUT_DA_NF_IPI_NF_ICMS_ST_NF" ' Chave única

         ' Combinações com VL_SEG_NF, VL_IPI_NF, VL_ICMS_ST_NF
        dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_SEG_NF + .VL_IPI_NF + .VL_ICMS_ST_NF, 2), _
                                            "Os valores dos campos VL_SEG_NF,  VL_IPI_NF e VL_ICMS_ST_NF foram somados ao valor do campo VL_MERC_SPED", _
                                            "Verifique os valores de seguro, IPI, ICMS ST e mercadorias no XML e ajuste o SPED"), _
                                            "VL_SEG_NF_IPI_NF_ICMS_ST_NF" ' Chave única

         ' Combinações com todos os campos
         dicPossibilidadesInvalidas.Add Array(VBA.Round(.VL_MERC_NF + .VL_FRT_NF + .VL_OUT_DA_NF + .VL_SEG_NF + .VL_IPI_NF + .VL_ICMS_ST_NF, 2), _
                                             "Os valores dos campos VL_FRT_NF, VL_OUT_DA_NF, VL_SEG_NF, VL_IPI_NF e VL_ICMS_ST_NF foram somados ao valor do campo VL_MERC_SPED", _
                                             "Verifique todos os valores (frete, outras despesas, seguro, IPI, ICMS ST e mercadorias) no XML e ajuste o SPED"), _
                                            "TODOS_CAMPOS" ' Chave única

    End With

    Set CarregarPossibilidadesInvalidas = dicPossibilidadesInvalidas

End Function

Public Function ValidarCampo_VL_FRT()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_FRT_NF <> .VL_FRT_SPED And .VL_FRT_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_FRT_NF (R$ " & VBA.Round(.VL_FRT_NF, 2) & ") e VL_FRT_SPED (R$ " & VBA.Round(.VL_FRT_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor de frete do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_SEG()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_SEG_NF <> .VL_SEG_SPED And .VL_SEG_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_SEG_NF (R$ " & VBA.Round(.VL_SEG_NF, 2) & ") e VL_SEG_SPED (R$ " & VBA.Round(.VL_SEG_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor de seguro do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_OUT_DA()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_OUT_DA_NF <> .VL_OUT_DA_SPED And .VL_OUT_DA_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_OUT_DA_NF (R$ " & VBA.Round(.VL_OUT_DA_NF, 2) & ") e VL_OUT_DA_SPED (R$ " & VBA.Round(.VL_OUT_DA_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor de outras despesas do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_BC_ICMS()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_BC_ICMS_NF = 0 And .VL_BC_ICMS_SPED > 0
                .INCONSISTENCIA = "O valor do campo VL_BC_ICMS_NF está zerado e o campo VL_BC_ICMS_SPED (R$ " & VBA.Round(.VL_BC_ICMS_SPED, 2) & ") está maior que zero"
                .SUGESTAO = "Zerar valor do campo VL_BC_ICMS_SPED"
                
            Case .VL_BC_ICMS_NF < .VL_BC_ICMS_SPED And .VL_BC_ICMS_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_BC_ICMS_NF (R$ " & VBA.Round(.VL_BC_ICMS_NF, 2) & ") e VL_BC_ICMS_SPED (R$ " & VBA.Round(.VL_BC_ICMS_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor de base do ICMS do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_ICMS()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_ICMS_NF = 0 And .VL_ICMS_SPED > 0
                .INCONSISTENCIA = "O valor do campo VL_ICMS_NF está zerado e o campo VL_ICMS_SPED (R$ " & VBA.Round(.VL_ICMS_SPED, 2) & ") está maior que zero"
                .SUGESTAO = "Zerar valor do campo VL_ICMS_SPED"
                
            Case .VL_ICMS_NF < .VL_ICMS_SPED And .VL_ICMS_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_ICMS_NF (R$ " & VBA.Round(.VL_ICMS_NF, 2) & ") e VL_ICMS_SPED (R$ " & VBA.Round(.VL_ICMS_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor de ICMS do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_BC_ICMS_ST()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_BC_ICMS_ST_NF < .VL_BC_ICMS_ST_SPED And .VL_BC_ICMS_ST_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_BC_ICMS_ST_NF (R$ " & VBA.Round(.VL_BC_ICMS_ST_NF, 2) & ") e VL_BC_ICMS_ST_SPED (R$ " & VBA.Round(.VL_BC_ICMS_ST_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor de base do ICMS-ST do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_ICMS_ST()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_ICMS_ST_NF < .VL_ICMS_ST_SPED And .VL_ICMS_ST_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_ICMS_ST_NF (R$ " & VBA.Round(.VL_ICMS_ST_NF, 2) & ") e VL_ICMS_ST_SPED (R$ " & VBA.Round(.VL_ICMS_ST_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor de ICMS-ST do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_IPI()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_IPI_NF < .VL_IPI_SPED And .VL_IPI_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_IPI_NF (R$ " & VBA.Round(.VL_IPI_NF, 2) & ") e VL_IPI_SPED (R$ " & VBA.Round(.VL_IPI_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor de IPI do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_PIS()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_PIS_NF <> .VL_PIS_SPED And .VL_PIS_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_PIS_NF (R$ " & VBA.Round(.VL_PIS_NF, 2) & ") e VL_PIS_SPED (R$ " & VBA.Round(.VL_PIS_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor do PIS do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_COFINS()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_COFINS_NF <> .VL_COFINS_SPED And .VL_COFINS_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_COFINS_NF (R$ " & VBA.Round(.VL_COFINS_NF, 2) & ") e VL_COFINS_SPED (R$ " & VBA.Round(.VL_COFINS_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor da COFINS do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_PIS_ST()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_PIS_NF <> .VL_PIS_SPED And .VL_PIS_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_PIS_NF (R$ " & VBA.Round(.VL_PIS_NF, 2) & ") e VL_PIS_SPED (R$ " & VBA.Round(.VL_PIS_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor do PIS-ST do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_VL_COFINS_ST()
    
    With CamposNota
        
        Select Case True
            
            Case .VL_COFINS_NF <> .VL_COFINS_SPED And .VL_COFINS_NF > 0
                .INCONSISTENCIA = "Os valores dos campos VL_COFINS_NF (R$ " & VBA.Round(.VL_COFINS_NF, 2) & ") e VL_COFINS_SPED (R$ " & VBA.Round(.VL_COFINS_SPED, 2) & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo valor da COFINS-ST do XML para o SPED"
                
        End Select
        
    End With
    
End Function
