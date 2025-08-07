Attribute VB_Name = "clsRegrasDivergenciasProdutos"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private CamposRelatorio As Variant
Private AtividadeComercial As Boolean
Private AtividadeIndustrial As Boolean
Private dicDados0000 As New Dictionary
Private dicTitulos0000 As New Dictionary
Private arrFuncoesValidacao As New ArrayList
Private dicTitulosRelatorio As New Dictionary
Private ValidacoesCFOP As New clsRegrasFiscaisCFOP
Private ValidacoesGerais As New clsRegrasFiscaisGerais
Private ValidacoesCOD_BARRAS As New clsRegrasFiscaisCodigoBarras

Private Sub Class_Initialize()

    Set dicTitulosRelatorio = Util.MapearTitulos(relDivergenciasProdutos, 3)
    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
    Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
    
End Sub

Public Function IdentificarDivergenciasProdutos(ByRef Registros As ArrayList)

Dim i As Long
    
    For i = 0 To Registros.Count - 1
                
        Call DadosDivergenciasProdutos.ResetarCamposProduto
        Call DadosDivergenciasProdutos.CarregarDadosRegistroDivergenciaProdutos(Registros(i))
        
        CamposRelatorio = Registros(i)
        Call IdentificarDivergenciaProduto
        
        Registros(i) = CamposRelatorio
        
    Next i
    
End Function

Private Sub IdentificarDivergenciaProduto()

Dim Validacao As Variant
    
    With CamposProduto
        
        Call CarregarValidacoes
        
        For Each Validacao In arrFuncoesValidacao
            
            If .INCONSISTENCIA <> "" Then
                Call Util.GravarSugestao(CamposRelatorio, dicTitulosRelatorio, .INCONSISTENCIA, .SUGESTAO, dicInconsistenciasIgnoradas)
                Exit Sub
                
            End If
            
            CallByName Me, CStr(Validacao), VbMethod
            
        Next Validacao
        
    End With
    
End Sub

Private Sub CarregarValidacoes()
    
    With arrFuncoesValidacao
                
        If Not .contains("ValidarCampo_VL_ITEM") Then .Add "ValidarCampo_VL_ITEM"
        If Not .contains("ValidarCampo_COD_ITEM") Then .Add "ValidarCampo_COD_ITEM"
        If Not .contains("ValidarCampo_VL_DESC") Then .Add "ValidarCampo_VL_DESC"
        If Not .contains("ValidarCampo_DESCR_ITEM") Then .Add "ValidarCampo_DESCR_ITEM"
        If Not .contains("ValidarCampo_COD_BARRA") Then .Add "ValidarCampo_COD_BARRA"
        If Not .contains("ValidarCampo_CEST") Then .Add "ValidarCampo_CEST"
        If Not .contains("ValidarCampo_CFOP") Then .Add "ValidarCampo_CFOP"
        If Not .contains("ValidarCampo_CST_ICMS") Then .Add "ValidarCampo_CST_ICMS"
        If Not .contains("ValidarCampo_UNID") Then .Add "ValidarCampo_UNID"
        If Not .contains("ValidarCampo_QTD") Then .Add "ValidarCampo_QTD"
        If Not .contains("ValidarCampo_VL_BC_IPI") Then .Add "ValidarCampo_VL_BC_IPI"
        If Not .contains("ValidarCampo_ALIQ_IPI") Then .Add "ValidarCampo_ALIQ_IPI"
        If Not .contains("ValidarCampo_VL_IPI") Then .Add "ValidarCampo_VL_IPI"
        If Not .contains("ValidarCampo_VL_BC_ICMS") Then .Add "ValidarCampo_VL_BC_ICMS"
        If Not .contains("ValidarCampo_ALIQ_ICMS") Then .Add "ValidarCampo_ALIQ_ICMS"
        If Not .contains("ValidarCampo_VL_ICMS") Then .Add "ValidarCampo_VL_ICMS"
        If Not .contains("ValidarCampo_VL_BC_ICMS_ST") Then .Add "ValidarCampo_VL_BC_ICMS_ST"
        If Not .contains("ValidarCampo_ALIQ_ST") Then .Add "ValidarCampo_ALIQ_ST"
        If Not .contains("ValidarCampo_VL_ICMS_ST") Then .Add "ValidarCampo_VL_ICMS_ST"
        If Not .contains("ValidarCampo_VL_OPER") Then .Add "ValidarCampo_VL_OPER"
        
    End With
    
End Sub

Public Function ValidarCampo_CFOP()
    
    Call ValidacoesCFOP.ValidarCampo_CFOP(CamposRelatorio, "DIVERGENCIAS")
    
End Function

Public Function ValidarCampo_COD_BARRA()

Dim vCOD_BARRA_NF As Boolean, vCOD_BARRA_SPED As Boolean
    
    With CamposProduto
                        
        'Carrega Informações da NF
        vCOD_BARRA_NF = ValidacoesCOD_BARRAS.ValidarCodigoBarras(.COD_BARRA_NF)
        
        'Carrega informações do SPED
        vCOD_BARRA_SPED = ValidacoesCOD_BARRAS.ValidarCodigoBarras(.COD_BARRA_SPED)
        
        Select Case True
            
            Case Not vCOD_BARRA_SPED And .COD_BARRA_SPED <> ""
                .INCONSISTENCIA = "O valor informado no campo COD_BARRA_SPED (" & .COD_BARRA_SPED & ") está inválido"
                .SUGESTAO = "Apagar código de barras informado no SPED"
                
            Case .COD_BARRA_NF <> .COD_BARRA_SPED And vCOD_BARRA_NF
                .INCONSISTENCIA = "Os campos COD_BARRA_NF (" & .COD_BARRA_NF & ") e COD_BARRA_SPED (" & .COD_BARRA_SPED & ") estão divergentes"
                .SUGESTAO = "Informar o mesmo código de barras do XML para o SPED"
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_CEST()

Dim vCEST_NF As Boolean, vCEST_SPED As Boolean
    
    With CamposProduto
        
        'Carrega Informações da NF
        vCEST_NF = ValidacoesGerais.ValidarCEST(.CEST_NF)
        
        'Carrega informações do SPED
        vCEST_SPED = ValidacoesGerais.ValidarCEST(.CEST_SPED)
        
        'Verifica se o CEST informado na NF é válido
        If Not vCEST_NF And .CEST_NF <> "" Then
            .INCONSISTENCIA = "O valor informado no campo CEST_NF está inválido"
            .SUGESTAO = "Apagar valor do CEST informado no campo CEST_NF"
            
        'Verifica se o CEST informado no SPED é válido
        ElseIf Not vCEST_SPED And .CEST_SPED <> "" Then
            .INCONSISTENCIA = "O valor informado no campo CEST_SPED (" & .CEST_SPED & ") está inválido"
            .SUGESTAO = "Apagar CEST informado no SPED"
            
        'Verifica se há divergências entre os campos CEST_NF e CEST_SPED
        ElseIf .CEST_NF <> .CEST_SPED And vCEST_NF Then
            .INCONSISTENCIA = "Os campos CEST_NF (" & .CEST_NF & ") e CEST_SPED (" & .CEST_SPED & ") estão divergentes"
            .SUGESTAO = "Informar o mesmo código CEST do XML para o SPED"
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_COD_ITEM()
    
    With CamposProduto
        
        If .COD_ITEM_NF = .COD_ITEM_SPED Then
            
            .INCONSISTENCIA = "Os campos COD_ITEM_NF (" & .COD_ITEM_NF & ") e COD_ITEM_SPED (" & .COD_ITEM_SPED & ") estão iguais"
            .SUGESTAO = "O lançamento de itens no SPED deve conter o COD_ITEM do contribuinte do arquivo não o do fornecedor"
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_QTD()
    
    If CBool(ConfiguracoesControlDocs.Range("IgnorarQtdUnidXML").value) Then Exit Function
    
    With CamposProduto
        
        If .QTD_NF <> .QTD_SPED Then
            
            .INCONSISTENCIA = "Os valores dos campos QTD_NF (" & .QTD_NF & ") e QTD_SPED (" & .QTD_SPED & ") estão divergentes"
            .SUGESTAO = "Informar o mesmo valor do campo QTD_NF para o campo QTD_SPED"
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_UNID()
        
    If CBool(ConfiguracoesControlDocs.Range("IgnorarQtdUnidXML").value) Then Exit Function
    
    With CamposProduto
        
        If Not .UNID_SPED Like "*" & .UNID_NF & "*" Then
            
            .INCONSISTENCIA = "Os valores dos campos UNID_NF (" & .UNID_NF & ") e UNID_SPED (" & .UNID_SPED & ") estão divergentes"
            .SUGESTAO = "Informar o mesmo valor do campo UNID_NF para o campo UNID_SPED"
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_DESCR_ITEM()
    
    With CamposProduto
        
        If .DESCR_ITEM_SPED = "ITEM NÃO IDENTIFICADO" Then
            
            .INCONSISTENCIA = "O código " & .COD_ITEM_SPED & " informado no campo COD_ITEM_SPED não possui cadastro no registro 0200"
            .SUGESTAO = "Cadastrar item no registro 0200"
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_VL_ITEM()
        
    With CamposProduto
        
        If .VL_ITEM_NF <> .VL_ITEM_SPED And .VL_ITEM_NF > 0 Then
            
            .INCONSISTENCIA = "Os valores dos campos VL_ITEM_NF (" & .VL_ITEM_NF & ") e VL_ITEM_SPED (" & .VL_ITEM_SPED & ") estão divergentes"
            .SUGESTAO = "Informar o mesmo valor do campo VL_ITEM_NF para o campo VL_ITEM_SPED"
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_VL_DESC()
    
    With CamposProduto
        
        If .VL_DESC_NF <> .VL_DESC_SPED Then
            
            .INCONSISTENCIA = "Os valores dos campos VL_DESC_NF (" & .VL_DESC_NF & ") e VL_DESC_SPED (" & .VL_DESC_SPED & ") estão divergentes"
            .SUGESTAO = "Informar o mesmo valor do campo VL_DESC_NF para o campo VL_DESC_SPED"
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_VL_OPER()
    
    With CamposProduto
        
        If .VL_OPER_NF <> .VL_OPER_SPED And .VL_OPER_NF > 0 Then
            
            .INCONSISTENCIA = "Os valores dos campos VL_OPER_NF (" & .VL_OPER_NF & ") e VL_OPER_SPED (" & .VL_OPER_SPED & ") estão divergentes"
            .SUGESTAO = "Provável erro no lançamento do item no SPED, corrija a escrituração do item"
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_VL_BC_IPI()
    
    With CamposProduto
        
        If .VL_BC_IPI_NF <> .VL_BC_IPI_SPED Then
            
            Select Case True
                
                Case Not VerificarOperacaoUsoConsumo And Not VerificarOperacaoAtivo
                    .INCONSISTENCIA = "Os valores dos campos VL_BC_IPI_NF (" & .VL_BC_IPI_NF & ") e VL_BC_IPI_SPED (" & .VL_BC_IPI_SPED & ") estão divergentes"
                    .SUGESTAO = "Informar o mesmo valor do campo VL_BC_IPI_NF para o campo VL_BC_IPI_SPED"
                    
            End Select
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_ALIQ_IPI()
    
    With CamposProduto
        
        If .ALIQ_IPI_NF <> .ALIQ_IPI_SPED Then
            
            Select Case True
                
                Case Not VerificarOperacaoUsoConsumo And Not VerificarOperacaoAtivo
                    .INCONSISTENCIA = "Os valores dos campos ALIQ_IPI_NF (" & .ALIQ_IPI_NF & ") e ALIQ_IPI_SPED (" & .ALIQ_IPI_SPED & ") estão divergentes"
                    .SUGESTAO = "Informar o mesmo valor do campo ALIQ_IPI_NF para o campo ALIQ_IPI_SPED"
                    
            End Select
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_VL_IPI()
    
    With CamposProduto
        
        If .VL_IPI_NF <> .VL_IPI_SPED Then
            
            Select Case True
                
                Case Not VerificarOperacaoUsoConsumo And Not VerificarOperacaoAtivo
                    .INCONSISTENCIA = "Os valores dos campos VL_IPI_NF (" & .VL_IPI_NF & ") e VL_IPI_SPED (" & .VL_IPI_SPED & ") estão divergentes"
                    .SUGESTAO = "Informar o mesmo valor do campo VL_IPI_NF para o campo VL_IPI_SPED"
                    
            End Select
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_VL_BC_ICMS()
    
    With CamposProduto
        
        If .VL_BC_ICMS_NF <> .VL_BC_ICMS_SPED Then
            
            Select Case True
                
                Case Not VerificarOperacaoUsoConsumo And Not VerificarOperacaoAtivo
                    .INCONSISTENCIA = "Os valores dos campos VL_BC_ICMS_NF (" & .VL_BC_ICMS_NF & ") e VL_BC_ICMS_SPED (" & .VL_BC_ICMS_SPED & ") estão divergentes"
                    .SUGESTAO = "Informar o mesmo valor do campo VL_BC_ICMS_NF para o campo VL_BC_ICMS_SPED"
                    
            End Select
            
        End If
        
    End With
    
End Function

Private Function VerificarOperacaoUsoConsumo() As Boolean
    
    With CamposProduto
        
        Select Case True
            
            Case .CFOP_SPED Like "#407", .CFOP_SPED Like "#556", .CFOP_SPED Like "#653"
                VerificarOperacaoUsoConsumo = True
                
        End Select
        
    End With
    
End Function

Private Function VerificarOperacaoAtivo() As Boolean
    
    With CamposProduto
        
        Select Case True
            
            Case .CFOP_SPED Like "#406", .CFOP_SPED Like "#551"
                VerificarOperacaoAtivo = True
                
        End Select
        
    End With
    
End Function

Public Function ValidarCampo_ALIQ_ICMS()
    
    With CamposProduto
        
        If .ALIQ_ICMS_NF <> .ALIQ_ICMS_SPED Then
            
            Select Case True
                
                Case Not VerificarOperacaoUsoConsumo And Not VerificarOperacaoAtivo
                    .INCONSISTENCIA = "Os valores dos campos ALIQ_ICMS_NF (" & .ALIQ_ICMS_NF & ") e ALIQ_ICMS_SPED (" & .ALIQ_ICMS_SPED & ") estão divergentes"
                    .SUGESTAO = "Informar o mesmo valor do campo ALIQ_ICMS_NF para o campo ALIQ_ICMS_SPED"
                    
                Case VerificarOperacaoUsoConsumo And .ALIQ_ICMS_SPED > 0
                    .INCONSISTENCIA = "Operação de uso e consumo (" & .ALIQ_ICMS_NF & ") e ALIQ_ICMS_SPED (" & .ALIQ_ICMS_SPED & ") estão divergentes"
                    .SUGESTAO = "Informar o mesmo valor do campo ALIQ_ICMS_NF para o campo ALIQ_ICMS_SPED"
                    
            End Select
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_VL_ICMS()
    
    With CamposProduto
        
        If .VL_ICMS_NF <> .VL_ICMS_SPED Then
            
            Select Case True
                
                Case Not VerificarOperacaoUsoConsumo And Not VerificarOperacaoAtivo
                    .INCONSISTENCIA = "Os valores dos campos VL_ICMS_NF (" & .VL_ICMS_NF & ") e VL_ICMS_SPED (" & .VL_ICMS_SPED & ") estão divergentes"
                    .SUGESTAO = "Informar o mesmo valor do campo VL_ICMS_NF para o campo VL_ICMS_SPED"
                    
            End Select
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_VL_BC_ICMS_ST()
    
    With CamposProduto
        
        If .VL_BC_ICMS_ST_NF <> .VL_BC_ICMS_ST_SPED Then
            
            .INCONSISTENCIA = "Os valores dos campos VL_BC_ICMS_ST_NF (" & .VL_BC_ICMS_ST_NF & ") e VL_BC_ICMS_ST_SPED (" & .VL_BC_ICMS_ST_SPED & ") estão divergentes"
            .SUGESTAO = "Informar o mesmo valor do campo VL_BC_ICMS_ST_NF para o campo VL_BC_ICMS_ST_SPED"
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_ALIQ_ST()
    
    With CamposProduto
        
        If .ALIQ_ST_NF <> .ALIQ_ST_SPED Then
            
            .INCONSISTENCIA = "Os valores dos campos ALIQ_ST_NF (" & .ALIQ_ST_NF & ") e ALIQ_ST_SPED (" & .ALIQ_ST_SPED & ") estão divergentes"
            .SUGESTAO = "Informar o mesmo valor do campo ALIQ_ST_NF para o campo ALIQ_ST_SPED"
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_VL_ICMS_ST()
    
    With CamposProduto
        
        If .VL_ICMS_ST_NF <> .VL_ICMS_ST_SPED Then
            
            .INCONSISTENCIA = "Os valores dos campos VL_ICMS_ST_NF (" & .VL_ICMS_ST_NF & ") e VL_ICMS_ST_SPED (" & .VL_ICMS_ST_SPED & ") estão divergentes"
            .SUGESTAO = "Informar o mesmo valor do campo VL_ICMS_ST_NF para o campo VL_ICMS_ST_SPED"
            
        End If
        
    End With
    
End Function

Public Function ValidarCampo_CST_ICMS()
    
    With CamposProduto
        
        Select Case True
            
            'Identifica CSOSN de operações com permissão de crédito
            Case .CST_ICMS_NF Like "#101" And Not .CST_ICMS_SPED Like "*90"
                .INCONSISTENCIA = "Campo CST_ICMS_NF (" & .CST_ICMS_NF & ") indica indica operação com permissão de crédito do ICMS"
                .SUGESTAO = "Informar o CST 90 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações sem permissão de crédito
            Case .CST_ICMS_NF Like "#102" And Not .CST_ICMS_SPED Like "*90"
                .INCONSISTENCIA = "Campo CST_ICMS_NF (" & .CST_ICMS_NF & ") indica indica operação sem permissão de crédito do ICMS"
                .SUGESTAO = "Informar o CST 90 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações com isenção do ICMS
            Case .CST_ICMS_NF Like "#103" And Not .CST_ICMS_SPED Like "*40"
                .INCONSISTENCIA = "Campo CST_ICMS_NF (" & .CST_ICMS_NF & ") indica indica operação isenta"
                .SUGESTAO = "Informar o CST 40 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações sujeitas a cobrança da Substituição tributária do ICMS
            Case .CST_ICMS_NF Like "#20#" And Not .CST_ICMS_SPED Like "*60"
                .INCONSISTENCIA = "Campo CST_ICMS_NF (" & .CST_ICMS_NF & ") indica indica operação com cobrança da Substituição Tributária do ICMS"
                .SUGESTAO = "Informar o CST 60 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações com imunidade
            Case .CST_ICMS_NF Like "#300" And Not .CST_ICMS_SPED Like "*41"
                .INCONSISTENCIA = "Campo CST_ICMS_NF (" & .CST_ICMS_NF & ") indica indica operação com imune"
                .SUGESTAO = "Informar o CST 41 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações não-tributadas
            Case .CST_ICMS_NF Like "#400" And Not .CST_ICMS_SPED Like "*41"
                .INCONSISTENCIA = "Campo CST_ICMS_NF (" & .CST_ICMS_NF & ") indica indica operação não-tributada"
                .SUGESTAO = "Informar o CST 41 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações sujeitas a Substituição tributária do ICMS
            Case .CST_ICMS_NF Like "#500" And Not .CST_ICMS_SPED Like "*60"
                .INCONSISTENCIA = "Campo CST_ICMS_NF (" & .CST_ICMS_NF & ") indica indica operação com ICMS cobrado anteriormente por substituição"
                .SUGESTAO = "Informar o CST 60 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações com tributação do ICMS
            Case .CST_ICMS_NF Like "#900" And Not .CST_ICMS_SPED Like "*90"
                .INCONSISTENCIA = "Campo CST_ICMS_NF (" & .CST_ICMS_NF & ") indica indica outras operações"
                .SUGESTAO = "Informar o CST 90 da tabela B para o campo CST_ICMS_SPED"
                
        End Select
        
    End With
    
End Function

Public Function VerificarCampoCST_ICMS(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim ORIG_NF As String, ORIG_SPED As String
Dim CST_ICMS_NF As String, CST_ICMS_SPED$, INCONSISTENCIA$, SUGESTAO$
Dim vCSOSN_NF As Boolean
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    'Carrega Informações da NF
    CST_ICMS_NF = Util.ApenasNumeros(Campos(dicTitulos("CST_ICMS_NF") - i))
    ORIG_NF = VBA.Left(CST_ICMS_NF, 1)
    
    'Carrega informações do SPED
    CST_ICMS_SPED = Util.ApenasNumeros(Campos(dicTitulos("CST_ICMS_SPED") - i))
    ORIG_SPED = VBA.Left(CST_ICMS_SPED, 1)
    
    'Verificações
    vCSOSN_NF = VBA.Len(CST_ICMS_NF) = 4
    
    If ORIG_NF = "1" And ORIG_SPED <> "2" Then
        INCONSISTENCIA = "O dígito de origem do CST_ICMS_SPED deve ser igual a 2"
        SUGESTAO = "Mudar o dígito de origem do CST_ICMS_SPED para 2"
        
    ElseIf ORIG_NF = "6" And ORIG_SPED <> "7" Then
        INCONSISTENCIA = "O dígito de origem do CST_ICMS_SPED deve ser igual a 7"
        SUGESTAO = "Mudar o dígito de origem do CST_ICMS_SPED para 7"
    
    'Verifica as operações com CSOSN
    ElseIf vCSOSN_NF Then
        
        Select Case True
                
            'Identifica CSOSN de operações com permissão de crédito
            Case CST_ICMS_NF Like "#101" And Not CST_ICMS_SPED Like "*90"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica operação com permissão de crédito do ICMS"
                SUGESTAO = "Informar o CST 90 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações sem permissão de crédito
            Case CST_ICMS_NF Like "#102" And Not CST_ICMS_SPED Like "*90"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica operação sem permissão de crédito do ICMS"
                SUGESTAO = "Informar o CST 90 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações com isenção do ICMS
            Case CST_ICMS_NF Like "#103" And Not CST_ICMS_SPED Like "*40"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica operação isenta"
                SUGESTAO = "Informar o CST 40 da tabela B para o campo CST_ICMS_SPED"
            
            'Identifica CSOSN de operações sujeitas a cobrança da Substituição tributária do ICMS
            Case CST_ICMS_NF Like "#20#" And Not CST_ICMS_SPED Like "*60"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica operação com cobrança da Substituição Tributária do ICMS"
                SUGESTAO = "Informar o CST 60 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações com imunidade
            Case CST_ICMS_NF Like "#300" And Not CST_ICMS_SPED Like "*41"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica operação com imune"
                SUGESTAO = "Informar o CST 41 da tabela B para o campo CST_ICMS_SPED"
            
            'Identifica CSOSN de operações não-tributadas
            Case CST_ICMS_NF Like "#400" And Not CST_ICMS_SPED Like "*41"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica operação não-tributada"
                SUGESTAO = "Informar o CST 41 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações sujeitas a Substituição tributária do ICMS
            Case CST_ICMS_NF Like "#500" And Not CST_ICMS_SPED Like "*60"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica operação com ICMS cobrado anteriormente por substituição"
                SUGESTAO = "Informar o CST 60 da tabela B para o campo CST_ICMS_SPED"
                
            'Identifica CSOSN de operações com tributação do ICMS
            Case CST_ICMS_NF Like "#900" And Not CST_ICMS_SPED Like "*90"
                INCONSISTENCIA = "Campo CST_ICMS_NF (" & CST_ICMS_NF & ") indica indica outras operações"
                SUGESTAO = "Informar o CST 90 da tabela B para o campo CST_ICMS_SPED"

        End Select
'
'    ElseIf (CST_ICMS_NF Like "#500" Or CST_ICMS_NF Like "#20#") And vCSOSN_NF And Not CST_ICMS_SPED Like "*60" Then
'
'        Call Util.GravarSugestao(Campos, dicTitulos, _
'            Inconsistencia:="Operação sujeita a ST de fornecedor optante pelo SMIPLES NACIONAL", _
'            Sugestao:="Informar o CST 60 da tabela B para o campo CST_ICMS_SPED", _
'            dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
'
'    ElseIf CST_ICMS_NF Like "#103" And vCSOSN_NF And Not CST_ICMS_SPED Like "*40" Then
'
'        Call Util.GravarSugestao(Campos, dicTitulos, _
'            Inconsistencia:="Operação Isenta de fornecedor optante pelo SMIPLES NACIONAL", _
'            Sugestao:="Informar o CST 40 da tabela B para o campo CST_ICMS_SPED", _
'            dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
'
'    ElseIf (CST_ICMS_NF Like "#300" Or CST_ICMS_NF Like "#400" Or CST_ICMS_NF Like "#103") And vCSOSN_NF And Not CST_ICMS_SPED Like "*41" Then
'
'        Call Util.GravarSugestao(Campos, dicTitulos, _
'            Inconsistencia:="Operação Imune ou Não Tributada de fornecedor optante pelo SMIPLES NACIONAL", _
'            Sugestao:="Informar o CST 41 da tabela B para o campo CST_ICMS_SPED", _
'            dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
'
'    ElseIf (CST_ICMS_NF Like "#900" Or CST_ICMS_NF Like "#10#") And Not CST_ICMS_NF Like "#103" And vCSOSN_NF And Not CST_ICMS_SPED Like "*90" Then
'
'        Call Util.GravarSugestao(Campos, dicTitulos, _
'            Inconsistencia:="Operação sem ST de fornecedor optante pelo SMIPLES NACIONAL", _
'            Sugestao:="Informar o CST 90 da tabela B para o campo CST_ICMS_SPED", _
'            dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
                                
    End If

    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulos, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)

End Function

