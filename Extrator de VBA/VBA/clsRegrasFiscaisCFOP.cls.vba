Attribute VB_Name = "clsRegrasFiscaisCFOP"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private arrFuncoesValidacao As New ArrayList
Private TabelasFiscais As New AtualizadorTabelasSPED

Private Sub Class_Initialize()

Dim dicDadosCFOP As New Dictionary
    
    If TabelaCFOP.Count = 0 Then
        
        Call Util.AtualizarBarraStatus("Carregando informações da tabela CFOP, por favor aguarde...")
        Call CarregarTabelaCFOP
        
    End If
    
    Call CarregarValidacoes
    
End Sub

Private Sub CarregarValidacoes()
    
    Call arrFuncoesValidacao.Clear
    
    With arrFuncoesValidacao
        
        .Add "VerificarCFOPVazio"
        .Add "VerificarTamanhoCFOP"
        .Add "VerificarExistenciaCFOP"
        .Add "VerificarOrigemOperacao"
        .Add "VerificarCFOPImposto"
        
    End With
    
End Sub

Public Sub CarregarTabelaCFOP()

Dim RegistrosCFOP As String
Dim CustomPart As New clsCustomPartXML
Dim Registros As Variant, Registro, Campos
    
    With CamposCFOP
        
        RegistrosCFOP = CustomPart.ExtrairTXTPartXML("TabelaCFOP")
        
        Registros = VBA.Split(RegistrosCFOP, vbLf)
        For Each Registro In Registros
            
            Campos = VBA.Split(Registro, "|")
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                If Not Campos(0) Like "*CFOP*" Then
                    
                    Call DadosValidacaoCFOP.CarregarDadosTabelaCFOP(Campos)
                    If .COD_CFOP <> "" Then TabelaCFOP(.COD_CFOP) = _
                        Array(.COD_CFOP, .DESCRICAO, .VIGENCIA_INICIAL, .VIGENCIA_FINAL)
                    
                End If
                
            End If
            
        Next Registro
        
    End With
    
End Sub

Public Function VerificarExistenciaCFOP() As Boolean
    
    If TabelaCFOP.Count = 0 Then Call CarregarTabelaCFOP
    
    With CamposCFOP
        
        If TabelaCFOP.Exists(.COD_CFOP) Then
            
            VerificarExistenciaCFOP = True
            ExisteCFOP = True
            
        End If
        
    End With
    
End Function

Public Function VerificarCFOP(ByRef Campos As Variant, ByRef dicTitulos As Dictionary)

Dim CFOP As String
Dim i As Integer
    
    Call DadosValidacaoCFOP.CarregarCamposCFOP(Campos, ActiveSheet)
    
    With CamposCFOP
        
        Select Case True
            
            Case CStr(.COD_CFOP) = ""
                Call Util.GravarSugestao(Campos, dicTitulos, _
                    INCONSISTENCIA:="O campo CFOP não foi informado", _
                    SUGESTAO:="informar um valor válido para o campo CFOP", _
                    dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
                    
            Case Else
                If Not VerificarExistenciaCFOP Then
                    Call Util.GravarSugestao(Campos, dicTitulos, _
                        INCONSISTENCIA:="O CFOP informado não existe na tabela CFOP", _
                        SUGESTAO:="Informar um valor válido para o campo CFOP", _
                        dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
                End If
                
        End Select
        
    End With
    
End Function

Public Function ValidarCFOPCompraIndustrializacao(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#101"
            ValidarCFOPCompraIndustrializacao = True
            
        Case CFOP < 4000 And CFOP Like "#111"
            ValidarCFOPCompraIndustrializacao = True
            
        Case CFOP < 4000 And CFOP Like "#116"
            ValidarCFOPCompraIndustrializacao = True
        
        Case CFOP < 4000 And CFOP Like "#120"
            ValidarCFOPCompraIndustrializacao = True
            
        Case CFOP < 4000 And CFOP Like "#122"
            ValidarCFOPCompraIndustrializacao = True
        
        'CHECAR: Verificar se o CFOP #124 é obrigatório para industriais ou equiparados a industariais
        Case CFOP < 4000 And CFOP Like "#124"
            ValidarCFOPCompraIndustrializacao = True
            
        'CHECAR: Verificar se o CFOP #125 é obrigatório para industriais ou equiparados a industariais
        Case CFOP < 4000 And CFOP Like "#125"
            ValidarCFOPCompraIndustrializacao = True
        
        Case CFOP < 4000 And CFOP Like "#401"
            ValidarCFOPCompraIndustrializacao = True
        
            
    End Select
    
End Function

Public Function ValidarCFOPCompraIndustrializacaoSemST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#101"
            ValidarCFOPCompraIndustrializacaoSemST = True
            
        Case CFOP < 4000 And CFOP Like "#111"
            ValidarCFOPCompraIndustrializacaoSemST = True
            
        Case CFOP < 4000 And CFOP Like "#116"
            ValidarCFOPCompraIndustrializacaoSemST = True
        
        Case CFOP < 4000 And CFOP Like "#120"
            ValidarCFOPCompraIndustrializacaoSemST = True
            
        Case CFOP < 4000 And CFOP Like "#122"
            ValidarCFOPCompraIndustrializacaoSemST = True
        
        'CHECAR: Verificar se o CFOP #124 é obrigatório para industriais ou equiparados a industariais
        Case CFOP < 4000 And CFOP Like "#124"
            ValidarCFOPCompraIndustrializacaoSemST = True
            
        'CHECAR: Verificar se o CFOP #125 é obrigatório para industriais ou equiparados a industariais
        Case CFOP < 4000 And CFOP Like "#125"
            ValidarCFOPCompraIndustrializacaoSemST = True
            
    End Select
    
End Function

Public Function ValidarCFOPCompraIndustrializacaoComST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#401"
            ValidarCFOPCompraIndustrializacaoComST = True
                    
    End Select
    
End Function

Public Function ValidarCFOPCompraRevendaSemST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#102"
            ValidarCFOPCompraRevendaSemST = True
            
        Case CFOP < 4000 And CFOP Like "#113"
            ValidarCFOPCompraRevendaSemST = True
            
        Case CFOP < 4000 And CFOP Like "#117"
            ValidarCFOPCompraRevendaSemST = True
        
        Case CFOP < 4000 And CFOP Like "#118"
            ValidarCFOPCompraRevendaSemST = True
            
        Case CFOP < 4000 And CFOP Like "#121"
            ValidarCFOPCompraRevendaSemST = True
            
        Case CFOP < 4000 And CFOP Like "#126"
            ValidarCFOPCompraRevendaSemST = True
            
    End Select
    
End Function

Public Function ValidarCFOPCompraRevendaComST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#403"
            ValidarCFOPCompraRevendaComST = True
            
    End Select
    
End Function

Public Function ValidarCFOPEntradaInterna(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "1###"
            ValidarCFOPEntradaInterna = True
            
    End Select
    
End Function

Public Function ValidarCFOPEntradaInterestadual(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "2###"
            ValidarCFOPEntradaInterestadual = True
            
    End Select
    
End Function

Public Function ValidarCFOPImportacao(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "3###"
            ValidarCFOPImportacao = True
            
    End Select
    
End Function

Public Function ValidarCFOPSaidaInterna(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP > 4000 And CFOP Like "5###"
            ValidarCFOPSaidaInterna = True
            
    End Select
    
End Function

Public Function ValidarCFOPSaidaInterestadual(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP > 4000 And CFOP Like "6###"
            ValidarCFOPSaidaInterestadual = True
            
    End Select
    
End Function

Public Function ValidarCFOPExportacao(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP > 4000 And CFOP Like "7###"
            ValidarCFOPExportacao = True
            
    End Select
    
End Function

Public Function ValidarCFOPCompraCombustiveisIndustrializacao(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#651"
            ValidarCFOPCompraCombustiveisIndustrializacao = True
            
    End Select
    
End Function

Public Function ValidarCFOPEntradaBonificacao(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#910"
            ValidarCFOPEntradaBonificacao = True
            
    End Select
    
End Function

Public Function ValidarCFOPCompraCombustiveisRevenda(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#652"
            ValidarCFOPCompraCombustiveisRevenda = True
            
    End Select
    
End Function

Public Function ValidarCFOPCompraCombustiveisConsumo(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#653"
            ValidarCFOPCompraCombustiveisConsumo = True
            
    End Select
    
End Function

Public Function ValidarCFOPVendaCombustiveis(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP > 4000 And CFOP Like "#651"
            ValidarCFOPVendaCombustiveis = True
            
        Case CFOP > 4000 And CFOP Like "#652"
            ValidarCFOPVendaCombustiveis = True
            
        Case CFOP > 4000 And CFOP Like "#653"
            ValidarCFOPVendaCombustiveis = True
            
        Case CFOP > 4000 And CFOP Like "#654"
            ValidarCFOPVendaCombustiveis = True
            
        Case CFOP > 4000 And CFOP Like "#655"
            ValidarCFOPVendaCombustiveis = True
            
        Case CFOP > 4000 And CFOP Like "#656"
            ValidarCFOPVendaCombustiveis = True
            
    End Select
    
End Function

Public Function ValidarCFOPCompraPrestacaoServico(ByVal CFOP As Integer) As Boolean
    
    Select Case True
                    
        Case CFOP < 4000 And CFOP Like "#128"
            ValidarCFOPCompraPrestacaoServico = True
            
    End Select
    
End Function

Public Function ValidarCFOPCompraAtivoImobilizadoSemST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#551"
            ValidarCFOPCompraAtivoImobilizadoSemST = True
            
    End Select
    
End Function

Public Function ValidarAquisicaoUsoConsumo(ByVal CFOP As Integer) As Boolean
    
    If CFOP < 4000 Then
        
        Select Case True
            
            Case CFOP Like "#407"
                ValidarAquisicaoUsoConsumo = True
                
            Case CFOP Like "#556"
                ValidarAquisicaoUsoConsumo = True
                
        End Select
        
    End If
    
End Function

Public Function ValidarAquisicaoAtivoImobilizado(ByVal CFOP As Integer) As Boolean
    
    If CFOP < 4000 Then
        
        Select Case True
            
            Case CFOP Like "#406"
                ValidarAquisicaoAtivoImobilizado = True
                
            Case CFOP Like "#551"
                ValidarAquisicaoAtivoImobilizado = True
                
        End Select
        
    End If
    
End Function

Public Function ValidarCFOPCompraUsoConsumoSemST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#556"
            ValidarCFOPCompraUsoConsumoSemST = True
            
    End Select
    
End Function

Public Function ValidarCFOPCompraUsoConsumoComST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#407"
            ValidarCFOPCompraUsoConsumoComST = True
            
    End Select
    
End Function

Public Function ValidarCFOPCompraAtivoImobilizadoComST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#406"
            ValidarCFOPCompraAtivoImobilizadoComST = True
            
    End Select
    
End Function

Public Function ValidarCFOPFaturamento(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP > 4000 And CFOP Like "#10#"
            ValidarCFOPFaturamento = True
            
        Case CFOP > 4000 And CFOP Like "#11#" And VBA.Right(CFOP, 3) <> 117
            ValidarCFOPFaturamento = True
            
        Case CFOP > 4000 And CFOP Like "#12#"
            ValidarCFOPFaturamento = True
        
        Case CFOP > 4000 And VBA.Right(CFOP, 3) > 400 And VBA.Right(CFOP, 3) < 406
            ValidarCFOPFaturamento = True
            
        Case CFOP > 4000 And VBA.Right(CFOP, 3) > 650 And VBA.Right(CFOP, 3) < 657
            ValidarCFOPFaturamento = True
        
        Case CFOP > 4000 And CFOP Like "#667"
            ValidarCFOPFaturamento = True
        
        Case CFOP > 4000 And CFOP Like "#922"
            ValidarCFOPFaturamento = True
            
    End Select
    
End Function

Public Function ValidarCFOPVendaSemST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP > 4000 And CFOP Like "#10#"
            ValidarCFOPVendaSemST = True
            
        Case CFOP > 4000 And CFOP Like "#11#" And VBA.Right(CFOP, 3) <> 117
            ValidarCFOPVendaSemST = True
            
        Case CFOP > 4000 And CFOP Like "#12#"
            ValidarCFOPVendaSemST = True
            
        Case CFOP > 4000 And CFOP Like "#922"
            ValidarCFOPVendaSemST = True
            
    End Select
    
End Function

Public Function ValidarCFOPVendaComST(ByVal CFOP As Integer) As Boolean
    
    Select Case True

        Case CFOP > 4000 And VBA.Right(CFOP, 3) > 400 And VBA.Right(CFOP, 3) < 406
            ValidarCFOPVendaComST = True
            
        Case CFOP > 4000 And VBA.Right(CFOP, 3) > 650 And VBA.Right(CFOP, 3) < 657
            ValidarCFOPVendaComST = True
            
    End Select
    
End Function

Public Function ValidarCFOPDevolucaoCompra(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP > 4000 And CFOP Like "#20#"
            ValidarCFOPDevolucaoCompra = True
            
        Case CFOP > 4000 And CFOP Like "#21#"
            ValidarCFOPDevolucaoCompra = True
        
        Case CFOP > 4000 And VBA.Right(CFOP, 3) > 409 And VBA.Right(CFOP, 3) < 412
            ValidarCFOPDevolucaoCompra = True
            
        Case CFOP > 4000 And CFOP Like "#503"
            ValidarCFOPDevolucaoCompra = True
            
        Case CFOP > 4000 And CFOP Like "#553"
            ValidarCFOPDevolucaoCompra = True
            
        Case CFOP > 4000 And CFOP Like "#555"
            ValidarCFOPDevolucaoCompra = True
            
        Case CFOP > 4000 And CFOP Like "#556"
            ValidarCFOPDevolucaoCompra = True
            
        Case CFOP > 4000 And VBA.Right(CFOP, 3) > 659 And VBA.Right(CFOP, 3) < 663
            ValidarCFOPDevolucaoCompra = True
            
    End Select
    
End Function

Public Function ValidarCFOPDevolucaoIndustrializacaoSemST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#201"
            ValidarCFOPDevolucaoIndustrializacaoSemST = True
            
        Case CFOP < 4000 And CFOP Like "#203"
            ValidarCFOPDevolucaoIndustrializacaoSemST = True
        
        Case CFOP < 4000 And CFOP Like "#212"
            ValidarCFOPDevolucaoIndustrializacaoSemST = True
            
        Case CFOP < 4000 And CFOP Like "#213"
            ValidarCFOPDevolucaoIndustrializacaoSemST = True
            
        Case CFOP < 4000 And CFOP Like "#214"
            ValidarCFOPDevolucaoIndustrializacaoSemST = True
            
        Case CFOP < 4000 And CFOP Like "#215"
            ValidarCFOPDevolucaoIndustrializacaoSemST = True
            
    End Select
    
End Function

Public Function ValidarCFOPDevolucaoTransferenciaIndustrializacaoSemST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#208"
            ValidarCFOPDevolucaoTransferenciaIndustrializacaoSemST = True

    End Select
    
End Function

Public Function ValidarCFOPDevolucaoTransferenciaIndustrializacaoComST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#408"
            ValidarCFOPDevolucaoTransferenciaIndustrializacaoComST = True

    End Select
    
End Function

Public Function ValidarCFOPDevolucaoIndustrializacaoComST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#410"
            ValidarCFOPDevolucaoIndustrializacaoComST = True
            
    End Select
    
End Function

Public Function ValidarCFOPDevolucaoRevendaSemST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#202"
            ValidarCFOPDevolucaoRevendaSemST = True
            
        Case CFOP < 4000 And CFOP Like "#204"
            ValidarCFOPDevolucaoRevendaSemST = True
        
        Case CFOP < 4000 And CFOP Like "#216"
            ValidarCFOPDevolucaoRevendaSemST = True
            
    End Select
    
End Function

Public Function ValidarCFOPDevolucaoTransferenciaRevendaSemST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#209"
            ValidarCFOPDevolucaoTransferenciaRevendaSemST = True

    End Select
    
End Function

Public Function ValidarCFOPDevolucaoRevendaComST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#411"
            ValidarCFOPDevolucaoRevendaComST = True

    End Select
    
End Function

Public Function ValidarCFOPDevolucaoTransferenciaRevendaComST(ByVal CFOP As Integer) As Boolean
    
    Select Case True
        
        Case CFOP < 4000 And CFOP Like "#409"
            ValidarCFOPDevolucaoTransferenciaRevendaComST = True

    End Select
    
End Function

Public Sub ValidarCampo_CFOP(ByRef Campos As Variant, ByVal Imposto As String)

Dim Validacao As Variant
    
    Call DadosValidacaoCFOP.ResetarCamposCFOP
    Call DadosValidacaoCFOP.CarregarCamposCFOP(Campos, Imposto)
    Set dicTitulosRelatorio = Util.MapearTitulos(ActiveSheet, 3)
    
    With CamposCFOP
        
        For Each Validacao In arrFuncoesValidacao
            
            If .INCONSISTENCIA <> "" Then
                
                Call Util.GravarSugestao(Campos, DadosValidacaoCFOP.dicTitulosRelatorio, .INCONSISTENCIA, .SUGESTAO, dicInconsistenciasIgnoradas)
                Exit Sub
                
            End If
            
            CallByName Me, CStr(Validacao), VbMethod
            
        Next Validacao
        
        If .INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, DadosValidacaoCFOP.dicTitulosRelatorio, .INCONSISTENCIA, .SUGESTAO, dicInconsistenciasIgnoradas)
        
    End With
    
End Sub

Public Function VerificarCFOPImposto() As Boolean
    
    With CamposCFOP
        
        Select Case True
            
            Case DadosValidacaoCFOP.TipoRelatorio = "IPI"
                VerificarCFOPImposto = Validar_CFOP_ICMS
                
            Case DadosValidacaoCFOP.TipoRelatorio = "ICMS"
                VerificarCFOPImposto = Validar_CFOP_ICMS
                
            Case DadosValidacaoCFOP.TipoRelatorio = "PISCOFINS"
                'VerificarCFOPImposto = Validar_CFOP_ICMS
                
            Case DadosValidacaoCFOP.TipoRelatorio = "DIVERGENCIAS"
                VerificarCFOPImposto = Validar_CFOP_DIVERGENCIAS
                
        End Select
        
    End With
    
End Function

Public Function VerificarCFOPVazio() As Boolean
    
    If CamposCFOP.COD_CFOP = "" Then
        
        CamposCFOP.INCONSISTENCIA = "O campo CFOP não foi informado"
        CamposCFOP.SUGESTAO = "Informe um CFOP válido para a operação"
        VerificarCFOPVazio = True
        
    End If
    
End Function

Public Function VerificarTamanhoCFOP() As Boolean
    
    If VBA.Len(CamposCFOP.COD_CFOP) = 4 Then
        
        VerificarTamanhoCFOP = True
        DadosValidacaoCFOP.TamanhoCFOP = True
        
    End If
    
End Function

Public Function VerificarOrigemOperacao() As Boolean
    
    Select Case True
        
        Case CamposCFOP.COD_CFOP > 4000 And CamposCFOP.UF_CONTRIB = CamposCFOP.UF_PART And Not CamposCFOP.COD_CFOP Like "5*"
            CamposCFOP.INCONSISTENCIA = "CFOP (" & CamposCFOP.COD_CFOP & ") incompatível com operação interna - UF_CONTRIB (" & CamposCFOP.UF_CONTRIB & ") = UF_PART (" & CamposCFOP.UF_PART & ")"
            CamposCFOP.SUGESTAO = "Informe um CFOP começando com o dígito 5"
            VerificarOrigemOperacao = True
            
        Case CamposCFOP.COD_CFOP > 4000 And CamposCFOP.UF_CONTRIB <> CamposCFOP.UF_PART And Not CamposCFOP.COD_CFOP Like "6*"
            CamposCFOP.INCONSISTENCIA = "CFOP (" & CamposCFOP.COD_CFOP & ") incompatível com a operação interestadual - UF_CONTRIB (" & CamposCFOP.UF_CONTRIB & ") <> UF_PART (" & CamposCFOP.UF_PART & ")"
            CamposCFOP.SUGESTAO = "Informe um CFOP começando com o dígito 6"
            VerificarOrigemOperacao = True
            
        Case CamposCFOP.COD_CFOP < 4000 And CamposCFOP.UF_CONTRIB = CamposCFOP.UF_PART And CamposCFOP.UF_CONTRIB <> "" And Not CamposCFOP.COD_CFOP Like "1*"
            CamposCFOP.INCONSISTENCIA = "CFOP (" & CamposCFOP.COD_CFOP & ") incompatível com operação interna - UF_CONTRIB (" & CamposCFOP.UF_CONTRIB & ") = UF_PART (" & CamposCFOP.UF_PART & ")"
            CamposCFOP.SUGESTAO = "Informe um CFOP começando com o dígito 1"
            VerificarOrigemOperacao = True
            
        Case CamposCFOP.COD_CFOP < 4000 And CamposCFOP.UF_CONTRIB <> CamposCFOP.UF_PART And CamposCFOP.UF_CONTRIB <> "" And CamposCFOP.COD_CFOP Like "1*"
            CamposCFOP.INCONSISTENCIA = "CFOP (" & CamposCFOP.COD_CFOP & ") incompatível com a operação - UF_CONTRIB (" & CamposCFOP.UF_CONTRIB & ") <> UF_PART (" & CamposCFOP.UF_PART & ")"
            CamposCFOP.SUGESTAO = "Informe um CFOP começando com o dígito 2"
            VerificarOrigemOperacao = True
            
    End Select
    
End Function

Private Function Validar_CFOP_IPI() As Boolean
    
    Select Case True
        
        Case DadosValidacaoCFOP.CST_IPITributado
            Validar_CFOP_IPI = False
            
        Case DadosValidacaoCFOP.CST_IPIAliqZero
            Validar_CFOP_IPI = False
            
        Case DadosValidacaoCFOP.CST_IPIIsento
            Validar_CFOP_IPI = False
                        
        Case DadosValidacaoCFOP.CST_IPINaoTributado
            Validar_CFOP_IPI = False
                                    
        Case DadosValidacaoCFOP.CST_IPIImune
            Validar_CFOP_IPI = False
                                                
        Case DadosValidacaoCFOP.CST_IPISuspensao
            Validar_CFOP_IPI = False
                                                            
        Case DadosValidacaoCFOP.CST_IPIOutrasSaidas
            Validar_CFOP_IPI = False
            
        Case DadosValidacaoCFOP.CST_IPIEntradaTributada
            Validar_CFOP_IPI = False
            
        Case DadosValidacaoCFOP.CST_IPIEntradaZero
            Validar_CFOP_IPI = False
                                                                                                
        Case DadosValidacaoCFOP.CST_IPIEntradaIsenta
            Validar_CFOP_IPI = False
                                                                                                            
        Case DadosValidacaoCFOP.CST_IPIEntradaNaoTributada
            Validar_CFOP_IPI = False
                                                                                                            
        Case DadosValidacaoCFOP.CST_IPIEntradaImune
            Validar_CFOP_IPI = False
                                                                                                            
        Case DadosValidacaoCFOP.CST_IPIEntradaSuspensa
            Validar_CFOP_IPI = False
                                                                                                            
        Case DadosValidacaoCFOP.CST_IPIOutrasEntradas
            Validar_CFOP_IPI = False
                                                                                                            
    End Select
    
End Function

Private Function ValidarSaidaTributadaIPI() As Boolean
    'TODO: Verificar regra que deveria ser para o CST_IPI e não para o CFOP
    Select Case True
        
        Case DadosValidacaoCFOP.CST_IPITributado And CamposCFOP.VL_IPI = 0
            CamposCFOP.INCONSISTENCIA = "O campo CST_IPI (" & CamposCFOP.CST_IPI & ") indica saída tributada mas não há débito destacado (VL_IPI = R$ " & VBA.Format(CamposCFOP.VL_IPI, "#0.00") & ")"
            CamposCFOP.SUGESTAO = "Informe um CFOP válido para a operação"
            ValidarSaidaTributadaIPI = True
            
    End Select
    
End Function

Private Function Validar_CFOP_ICMS() As Boolean
    
    Select Case True
        
        Case DadosValidacaoCFOP.EntradaSemST
            Validar_CFOP_ICMS = ValidarEntradasICMSSemST
            
        Case DadosValidacaoCFOP.EntradaComST
            Validar_CFOP_ICMS = ValidarEntradasICMSComST
            
    End Select
    
End Function

Private Function ValidarEntradasICMSSemST() As Boolean
    
    Select Case True
        
        Case DadosValidacaoCFOP.CST_ICMSComST And CamposCFOP.VL_ICMS = 0
            CamposCFOP.INCONSISTENCIA = "O CFOP (" & CamposCFOP.COD_CFOP & ") indica entrada sem ST com o campo CST_ICMS (" & CamposCFOP.CST_ICMS & ") indicando operacao com ST sem aproveitamento de crédito de ICMS (VL_ICMS = R$ " & VBA.Format(CamposCFOP.VL_ICMS, "#0.00") & ")"
            CamposCFOP.SUGESTAO = "Informe um CFOP válido para a operação"
            ValidarEntradasICMSSemST = True
            
    End Select
    
End Function

Private Function ValidarEntradasICMSComST() As Boolean
    
    Select Case True
        
        Case DadosValidacaoCFOP.CST_ICMSTributado And CamposCFOP.VL_ICMS > 0
            CamposCFOP.INCONSISTENCIA = "O CFOP (" & CamposCFOP.COD_CFOP & ") indica entrada com ST com o campo CST_ICMS (" & CamposCFOP.CST_ICMS & ") indicando operacao tributada com aproveitamento de crédito de ICMS (VL_ICMS = R$ " & VBA.Format(CamposCFOP.VL_ICMS, "#0.00") & ")"
            CamposCFOP.SUGESTAO = "Informe um CFOP válido para a operação"
            ValidarEntradasICMSComST = True
            
    End Select
    
End Function

Private Function Validar_CFOP_DIVERGENCIAS() As Boolean
    
    Select Case True
        
        Case DadosValidacaoCFOP.SaidaInterna
            Validar_CFOP_DIVERGENCIAS = VerificarSaidasInternas_Divergencias
            
        Case DadosValidacaoCFOP.SaidaInterestadual
            Validar_CFOP_DIVERGENCIAS = VerificarSaidasInterestaduais_Divergencias
            
        Case DadosValidacaoCFOP.Importacao
            Validar_CFOP_DIVERGENCIAS = VerificarImportacoes_Divergencias
            
        Case DadosValidacaoCFOP.Comercializacao
            Validar_CFOP_DIVERGENCIAS = VerificarOperacoesComerciais_Divergencias
            
    End Select
    
End Function

Private Function VerificarSaidasInternas_Divergencias() As Boolean
    
    Select Case True
        
        Case DadosValidacaoCFOP.EntradaInterestadual
            CamposCFOP.INCONSISTENCIA = "CFOP_NF (" & CamposCFOP.CFOP_NF & ") indicando saída interna com Camposcfop.CFOP_SPED (" & CamposCFOP.CFOP_SPED & ") indicando entrada interestadual"
            CamposCFOP.SUGESTAO = "Informe um CFOP de entrada interna no campo CFOP_SPED"
            VerificarSaidasInternas_Divergencias = True
            
        Case DadosValidacaoCFOP.Importacao
            CamposCFOP.INCONSISTENCIA = "CFOP_NF (" & CamposCFOP.CFOP_NF & ") indicando saída interna com Camposcfop.CFOP_SPED (" & CamposCFOP.CFOP_SPED & ") indicando importação"
            CamposCFOP.SUGESTAO = "Informe um CFOP de entrada interna no campo CFOP_SPED"
            VerificarSaidasInternas_Divergencias = True
            
        Case DadosValidacaoCFOP.VendaSemST And (Not DadosValidacaoCFOP.CompraRevendaSemST And _
            Not DadosValidacaoCFOP.CompraIndustrializacaoSemST And Not DadosValidacaoCFOP.CompraUsoConsumoSemST And _
            Not DadosValidacaoCFOP.CompraImobilizadoSemST) And Not CamposCFOP.CFOP_SPED Like "#9##"
            CamposCFOP.INCONSISTENCIA = "CFOP_NF (" & CamposCFOP.CFOP_NF & ") indicando venda interna sem ST com CFOP_SPED (" & CamposCFOP.CFOP_SPED & ") divergente de operação de compra sem ST"
            CamposCFOP.SUGESTAO = "Informe um CFOP de compra interna sem ST"
            VerificarSaidasInternas_Divergencias = True
            
        Case DadosValidacaoCFOP.VendaComST And Not DadosValidacaoCFOP.CompraRevendaComST And _
            Not DadosValidacaoCFOP.CompraUsoConsumoComST And Not DadosValidacaoCFOP.CompraImobilizadoComST And Not CamposCFOP.CFOP_SPED Like "#9##" And Not CamposCFOP.CFOP_SPED Like "#6##"
            CamposCFOP.INCONSISTENCIA = "CFOP_NF (" & CamposCFOP.CFOP_NF & ") indicando venda interna com ST com CFOP_SPED (" & CamposCFOP.CFOP_SPED & ") divergente de operação de compra com ST"
            CamposCFOP.SUGESTAO = "Informe um CFOP de compra interna com ST"
            VerificarSaidasInternas_Divergencias = True
            
        Case DadosValidacaoCFOP.VendaCombustivel And Not DadosValidacaoCFOP.CompraCombustivelRevenda And _
            Not DadosValidacaoCFOP.CompraCombustivelConsumo
            CamposCFOP.INCONSISTENCIA = "CFOP_NF (" & CamposCFOP.CFOP_NF & ") indicando venda de combustíveis e lubrificantes com CFOP_SPED (" & CamposCFOP.CFOP_SPED & ") divergente de operação de compra de combustíveis e lubrificantes"
            CamposCFOP.SUGESTAO = "Informe um CFOP de compra interna de combustíveis e lubrificantes"
            VerificarSaidasInternas_Divergencias = True
            
    End Select
    
End Function

Private Function VerificarSaidasInterestaduais_Divergencias() As Boolean

    Select Case True
        
        Case DadosValidacaoCFOP.EntradaInterna
            CamposCFOP.INCONSISTENCIA = "CFOP_NF (" & CamposCFOP.CFOP_NF & ") indicando saída interestadual com CFOP_SPED (" & CamposCFOP.CFOP_SPED & ") indicando entrada interna"
            CamposCFOP.SUGESTAO = "Informe um CFOP de entrada interestadual no campo CFOP_SPED"
            VerificarSaidasInterestaduais_Divergencias = True
            
        Case DadosValidacaoCFOP.Importacao
            CamposCFOP.INCONSISTENCIA = "CFOP_NF (" & CamposCFOP.CFOP_NF & ") indicando saída interestadual com CFOP_SPED (" & CamposCFOP.CFOP_SPED & ") indicando importação"
            CamposCFOP.SUGESTAO = "Informe um CFOP de entrada interestadual no campo CFOP_SPED"
            VerificarSaidasInterestaduais_Divergencias = True
            
        Case DadosValidacaoCFOP.VendaComST And Not DadosValidacaoCFOP.CompraRevendaComST And Not DadosValidacaoCFOP.CompraUsoConsumoComST And _
            Not DadosValidacaoCFOP.CompraImobilizadoComST And Not DadosValidacaoCFOP.CompraIndustrializacaoComST And Not CamposCFOP.CFOP_SPED Like "#9##"
            CamposCFOP.INCONSISTENCIA = "CFOP_NF (" & CamposCFOP.CFOP_NF & ") indicando venda interestadual com ST com CFOP_SPED (" & CamposCFOP.CFOP_SPED & ") divergente de operação de compra com ST"
            CamposCFOP.SUGESTAO = "Informe um CFOP de compra interestadual com ST"
            VerificarSaidasInterestaduais_Divergencias = True
            
        Case DadosValidacaoCFOP.VendaCombustivel And Not DadosValidacaoCFOP.CompraCombustivelRevenda And Not DadosValidacaoCFOP.CompraCombustivelConsumo
            CamposCFOP.INCONSISTENCIA = "CFOP_NF (" & CamposCFOP.CFOP_NF & ") indicando venda de combustíveis e lubrificantes com CFOP_SPED (" & CamposCFOP.CFOP_SPED & ") divergente de operação de compra de combustíveis e lubrificantes"
            CamposCFOP.SUGESTAO = "Informe um CFOP de compra interestadual de combustíveis e lubrificantes"
            VerificarSaidasInterestaduais_Divergencias = True
            
    End Select

End Function

Private Function VerificarImportacoes_Divergencias() As Boolean
    
    Select Case True
        
        Case DadosValidacaoCFOP.EntradaInterna
            CamposCFOP.INCONSISTENCIA = "CFOP_NF (" & CamposCFOP.CFOP_NF & ") indicando importação com CFOP_SPED (" & CamposCFOP.CFOP_SPED & ") indicando entrada interna"
            CamposCFOP.SUGESTAO = "Informe um CFOP de importação no campo CFOP_SPED"
            
        Case DadosValidacaoCFOP.EntradaInterestadual
            CamposCFOP.INCONSISTENCIA = "CFOP_NF (" & CamposCFOP.CFOP_NF & ") indicando importação com CFOP_SPED (" & CamposCFOP.CFOP_SPED & ") indicando entrada interestadual"
            CamposCFOP.SUGESTAO = "Informe um CFOP de importação no campo CFOP_SPED"
            
    End Select
    
End Function

Private Function VerificarOperacoesComerciais_Divergencias() As Boolean
    
    Select Case True
        
        Case DadosValidacaoCFOP.CompraIndustrializacao
            CamposCFOP.INCONSISTENCIA = "CFOP_SPED (" & CamposCFOP.CFOP_SPED & ") indicando compra para industrialização, mas contribuinte não é indústria (Campo IND_ATIV do registro 0000 = 1 - Outros)"
            CamposCFOP.SUGESTAO = "Informe um CFOP de compra para revenda"
            VerificarOperacoesComerciais_Divergencias = True
            
    End Select
    
End Function

Public Sub BaixarTabelaCFOP()
    
    Call TabelasFiscais.BaixarTabela(UrlTabelaCFOP, "TabelaCFOP")
    
End Sub

