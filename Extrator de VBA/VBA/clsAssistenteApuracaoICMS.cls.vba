Attribute VB_Name = "clsAssistenteApuracaoICMS"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private ValidacoesCFOP As New clsRegrasFiscaisCFOP
Private ValidacoesNCM As New clsRegrasFiscaisNCM
Private Apuracao As clsAssistenteApuracao
Private arrRelatorio As New ArrayList
Private dicTitulos As New Dictionary

Public Function GerarApuracaoAssistidaICMS()

Dim arrDocsC100 As New ArrayList
Dim Msg As String
    
    Inicio = Now()
    
    Call arrRelatorio.Clear
    Call dicInconsistenciasIgnoradas.RemoveAll
    Call DadosValidacaoCFOP.CarregarTitulosRelatorio(ActiveSheet)
    
    Set Apuracao = New clsAssistenteApuracao
    Set dicTitulos = Util.MapearTitulos(assApuracaoICMS, 3)
    
    Call CarregarDadosC170(arrDocsC100, Msg)
    Call CarregarDadosC190(arrDocsC100, Msg)
    
    Application.StatusBar = "Processo concluído com sucesso!"
    If arrRelatorio.Count > 0 Then
        
        On Error Resume Next
            If assApuracaoICMS.AutoFilter.FilterMode Then assApuracaoICMS.ShowAllData
        On Error GoTo 0
        Call Util.LimparDados(assApuracaoICMS, 4, False)
        
        Call Util.ExportarDadosArrayList(assApuracaoICMS, arrRelatorio)
        Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoICMS)
        Call Util.MsgInformativa("Relatório gerado com sucesso", "Assistente de Apuração do ICMS", Inicio)
        
    Else
        
        Msg = "Nenhum dado encontrado para geração do relatório." & vbCrLf & vbCrLf
        Msg = Msg & "Por favor verifique se o SPED foi importado e tente novamente."
        Call Util.MsgAlerta(Msg, "Assistente de Apuração do ICMS")
        
    End If
    
    Application.StatusBar = False
    
    DTO_RegistrosSPED.ResetarRegistrosSPED
    Set Apuracao = Nothing
    
End Function

Public Sub CarregarDadosC170(ByRef arrDocsC100 As ArrayList, ByVal Msg As String)

Dim ARQUIVO As String, CHV_REG$, CHV_0001$, CHV_C100$, COD_ITEM$, COD_PART$, COD_NAT$, UF_CONTRIB$, CONTRIBUINTE$, CNPJ_ESTABELECIMENTO$
Dim dicTitulosC170 As New Dictionary
Dim Dados As Range, Linha As Range
Dim Comeco As Double, VL_ITEM#
Dim Campos As Variant
Dim b As Long
        
    Set Dados = Util.DefinirIntervalo(regC170, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    Set Apuracao.dicTitulos = dicTitulos
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    
    b = 0
    Comeco = Timer
    With Apuracao
        
        Call .CarregarDadosRegistro0000
        
        For Each Linha In Dados.Rows
            
            .RedimensionarArray (dicTitulos.Count)
            Call Util.AntiTravamento(b, 100, Msg & "Carregando dados do registro C170", Dados.Rows.Count, Comeco)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                'Carrega as variáveis necessárias
                ARQUIVO = Campos(dicTitulosC170("ARQUIVO"))
                UF_CONTRIB = .ExtrairUFContribuinte(ARQUIVO)
                CNPJ_ESTABELECIMENTO = fnExcel.FormatarTexto(VBA.Split(ARQUIVO, "-")(1))
                
                CHV_C100 = Campos(dicTitulosC170("CHV_PAI_FISCAL"))
                If CHV_C100 <> "" Then arrDocsC100.Add CHV_C100
                
                COD_PART = .ExtrairCOD_PART_C100(CHV_C100)
                CHV_0001 = .ExtrairCHV_0001(ARQUIVO)
                COD_ITEM = Campos(dicTitulosC170("COD_ITEM"))
                COD_NAT = Campos(dicTitulosC170("COD_NAT"))
                VL_ITEM = fnExcel.FormatarValores(Campos(dicTitulosC170("VL_ITEM")))
                
                'Extrai dados do registro C100
                Call .ExtrairDadosC100(CHV_C100)
                
                'Extrai dados do registro 0150
                Call .ExtrairDados0150(CHV_0001, COD_PART)
                
                'Extrai dados do registro 0200
                Call .ExtrairDados0200(CHV_0001, COD_ITEM)
                
                'Extrai dados do registro C177
                Call .ExtrairDadosC177(CHV_C100)
                
                'Extrai dados do registro 0400
                Call .ExtrairDados0400(ARQUIVO, COD_NAT)
                
                'Atribui valores aos campos do relatório
                .AtribuirValor "REG", Campos(dicTitulosC170("REG"))
                .AtribuirValor "ARQUIVO", ARQUIVO
                .AtribuirValor "CHV_PAI_FISCAL", CHV_C100
                .AtribuirValor "CHV_REG", Campos(dicTitulosC170("CHV_REG"))
                .AtribuirValor "CNPJ_ESTABELECIMENTO", CNPJ_ESTABELECIMENTO
                .AtribuirValor "UF_CONTRIB", UF_CONTRIB
                .AtribuirValor "COD_ITEM", fnExcel.FormatarTexto(COD_ITEM)
                .AtribuirValor "IND_MOV", Campos(dicTitulosC170("IND_MOV"))
                .AtribuirValor "CFOP", Campos(dicTitulosC170("CFOP"))
                .AtribuirValor "CST_ICMS", fnExcel.FormatarTexto(Campos(dicTitulosC170("CST_ICMS")))
                .AtribuirValor "VL_ITEM", VL_ITEM
                .AtribuirValor "VL_DESP", .ExtrairVL_DESP_C100(CHV_C100, VL_ITEM)
                .AtribuirValor "VL_DESC", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_DESC")))
                .AtribuirValor "VL_BC_ICMS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_BC_ICMS")))
                .AtribuirValor "ALIQ_ICMS", fnExcel.ConverterValores(Campos(dicTitulosC170("ALIQ_ICMS")))
                .AtribuirValor "VL_ICMS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_ICMS")))
                .AtribuirValor "VL_BC_ICMS_ST", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_BC_ICMS_ST")))
                .AtribuirValor "ALIQ_ST", fnExcel.ConverterValores(Campos(dicTitulosC170("ALIQ_ST")))
                .AtribuirValor "VL_ICMS_ST", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_ICMS_ST")))
                
            End If
            
            Campos = ValidarRegrasFiscais(.Campo, dicTitulos, dtoRegSPED.r0000, dtoTitSPED.t0000)
            arrRelatorio.Add Campos
            
        Next Linha
        
    End With
    
End Sub

Public Sub CarregarDadosC190(ByRef arrDocsC100 As ArrayList, ByVal Msg As String)

Dim ARQUIVO As String, CHV_REG$, CHV_0001$, CHV_C100$, COD_PART$, COD_NAT$, UF_CONTRIB$, CONTRIBUINTE$, CNPJ_ESTABELECIMENTO$
Dim dicTitulosC190 As New Dictionary
Dim Dados As Range, Linha As Range
Dim Comeco As Double, VL_OPR#
Dim Campos As Variant
Dim b As Long
        
    Set Dados = Util.DefinirIntervalo(regC190, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    Set Apuracao.dicTitulos = dicTitulos
    Set dicTitulosC190 = Util.MapearTitulos(regC190, 3)
    
    b = 0
    Comeco = Timer
    With Apuracao
        
        Call .CarregarDadosRegistro0000
        
        For Each Linha In Dados.Rows
            
            .RedimensionarArray (dicTitulos.Count)
            Call Util.AntiTravamento(b, 100, Msg & "Carregando dados do registro C190", Dados.Rows.Count, Comeco)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                'Carrega as variáveis necessárias
                ARQUIVO = Campos(dicTitulosC190("ARQUIVO"))
                UF_CONTRIB = .ExtrairUFContribuinte(ARQUIVO)
                CNPJ_ESTABELECIMENTO = fnExcel.FormatarTexto(VBA.Split(ARQUIVO, "-")(1))
                CHV_0001 = .ExtrairCHV_0001(ARQUIVO)
                CHV_C100 = Campos(dicTitulosC190("CHV_PAI_FISCAL"))
                If arrDocsC100.contains(CHV_C100) Then GoTo Prx:
                
                COD_PART = .ExtrairCOD_PART_C100(CHV_C100)
                
                VL_OPR = fnExcel.FormatarValores(Campos(dicTitulosC190("VL_OPR")))
                
                'Extrai dados do registro C100
                Call .ExtrairDadosC100(CHV_C100)
                
                'Extrai dados do registro 0150
                Call .ExtrairDados0150(CHV_0001, COD_PART)
                
                'Extrai dados do registro 0200
                Call .ExtrairDados0200(CHV_0001, "")
                
                'Extrai dados do registro C177
                Call .ExtrairDadosC177(CHV_C100)
                
                'Extrai dados do registro 0400
                Call .ExtrairDados0400(ARQUIVO, COD_NAT)
                
                'Atribui valores aos campos do relatório
                .AtribuirValor "DESCR_ITEM", "O REGISTRO C190 NÃO POSSUI DADOS DE PRODUTOS"
                .AtribuirValor "REG", Campos(dicTitulosC190("REG"))
                .AtribuirValor "ARQUIVO", ARQUIVO
                .AtribuirValor "CHV_PAI_FISCAL", CHV_C100
                .AtribuirValor "CHV_REG", Campos(dicTitulosC190("CHV_REG"))
                .AtribuirValor "CNPJ_ESTABELECIMENTO", CNPJ_ESTABELECIMENTO
                .AtribuirValor "UF_CONTRIB", UF_CONTRIB
                .AtribuirValor "CFOP", Campos(dicTitulosC190("CFOP"))
                .AtribuirValor "CST_ICMS", fnExcel.FormatarTexto(Campos(dicTitulosC190("CST_ICMS")))
                .AtribuirValor "VL_ITEM", VL_OPR
                .AtribuirValor "VL_DESP", .ExtrairVL_DESP_C100(CHV_C100, VL_OPR)
                .AtribuirValor "VL_BC_ICMS", fnExcel.ConverterValores(Campos(dicTitulosC190("VL_BC_ICMS")))
                .AtribuirValor "ALIQ_ICMS", fnExcel.ConverterValores(Campos(dicTitulosC190("ALIQ_ICMS")))
                .AtribuirValor "VL_ICMS", fnExcel.ConverterValores(Campos(dicTitulosC190("VL_ICMS")))
                .AtribuirValor "VL_BC_ICMS_ST", fnExcel.ConverterValores(Campos(dicTitulosC190("VL_BC_ICMS_ST")))
                .AtribuirValor "VL_ICMS_ST", fnExcel.ConverterValores(Campos(dicTitulosC190("VL_ICMS_ST")))
                
            End If
            
            Campos = ValidarRegrasFiscais(.Campo, dicTitulos, dtoRegSPED.r0000, dtoTitSPED.t0000)
            arrRelatorio.Add Campos
Prx:
        Next Linha
        
    End With
    
End Sub

Public Function ValidarRegrasFiscais(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, _
    ByRef dicDados0000 As Dictionary, ByRef dicTitulos0000 As Dictionary) As Variant
    
Dim Registro As String, ARQUIVO$, UF$
Dim Campos0000 As Variant
Dim i As Integer
    
    If UBound(Campos) = -1 Then
        ValidarRegrasFiscais = Campos
        Exit Function
    End If
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    ARQUIVO = Campos(dicTitulos("ARQUIVO") - i)
    Registro = Campos(dicTitulos("REG") - i)
    
    If dicDados0000.Exists(ARQUIVO) Then
        
        Campos0000 = dicDados0000(ARQUIVO)
        If LBound(Campos0000) = 0 Then i = 1 Else i = 0
        
        UF = Campos0000(dicTitulos0000("UF") - i)
        
    End If
    
    Select Case Registro
        
        Case "C170"
            Call ValidarRegrasFiscaisC170(Campos, dicTitulos, UF)
            
        Case "C190"
            Call ValidarRegrasFiscaisC190(Campos, dicTitulos, UF)

    End Select
    
    ValidarRegrasFiscais = Campos
    
End Function

Public Function ValidarRegrasFiscaisC170(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal UFContrib As String)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidacoesCFOP.ValidarCampo_CFOP(Campos, "ICMS")
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoICMS.VerificarCampoCEST(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoICMS.ValidarCampo_DT_ENT_SAI(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidacoesNCM.ValidarCampo_COD_NCM(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_IND_MOV(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_TIPO_ITEM(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoICMS.ValidarCampo_CST_ICMS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoICMS.ValidarCampo_VL_ICMS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoICMS.ValidarCampo_VL_ICMS_ST(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoICMS.ValidarCampo_ALIQ_ICMS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoICMS.VerificarCampoCONTRIBUINTE(Campos, dicTitulos)
    
End Function

Public Function ValidarRegrasFiscaisC190(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal UFContrib As String)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidacoesCFOP.ValidarCampo_CFOP(Campos, "ICMS")
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoICMS.ValidarCampo_CST_ICMS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoICMS.ValidarCampo_VL_ICMS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoICMS.ValidarCampo_VL_ICMS_ST(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoICMS.ValidarCampo_ALIQ_ICMS(Campos, dicTitulos)
    
End Function

Public Function ReprocessarSugestoes()

Dim dicTitulos0000 As New Dictionary
Dim dicDados0000 As New Dictionary
Dim Dados As Range, Linha As Range
Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim Campos As Variant

    Set dicTitulos = Util.MapearTitulos(assApuracaoICMS, 3)
    If assApuracaoICMS.AutoFilterMode Then assApuracaoICMS.AutoFilter.ShowAllData
    
    Call DadosValidacaoCFOP.CarregarTitulosRelatorio(ActiveSheet)
    Set Dados = Util.DefinirIntervalo(assApuracaoICMS, 4, 3)
    If Dados Is Nothing Then Exit Function
    
    Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 100, "Reprocessando sugestões, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            Campos(dicTitulos("INCONSISTENCIA")) = Empty
            Campos(dicTitulos("SUGESTAO")) = Empty
            Call ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
            
            arrRelatorio.Add Campos
            
        End If
        
    Next Linha
    
    Call Util.LimparDados(assApuracaoICMS, 4, False)
    Call Util.ExportarDadosArrayList(assApuracaoICMS, arrRelatorio)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoICMS)
    
    Call Util.AtualizarBarraStatus("Processamento Concluído!")
    
End Function

Public Function AceitarSugestoes()

Dim dicTitulos0000 As New Dictionary
Dim dicDados0000 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim Campos As Variant
Dim arrDados As New ArrayList
Dim UltimaSugestao As String

    Inicio = Now()
    Application.StatusBar = "Implementando sugestões selecionadas, por favor aguarde..."
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoICMS, 3)
    Set Dados = assApuracaoICMS.Range("A4").CurrentRegion
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não foi encontrada nenhuma inconsistência!", "Inconsistências Fiscais")
        Exit Function
    End If
    
    Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
    
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
                        
            If Linha.EntireRow.Hidden = False And Campos(dicTitulos("SUGESTAO")) <> "" And Linha.Row > 3 Then
                 
'                If UltimaSugestao = Campos(dicTitulos("SUGESTAO")) Then Trava = Trava + 1
'                UltimaSugestao = Campos(dicTitulos("SUGESTAO"))
'                Debug.Print UltimaSugestao
                
VERIFICAR:
                Call Util.AntiTravamento(a, 10, "Aplicando sugestões sugeridas...", Dados.Rows.Count, Comeco)
                Select Case Campos(dicTitulos("SUGESTAO"))
                    
                    Case "Alterar valor do campo COD_SIT para: 08 - Regime Especial ou Norma Específica"
                        Campos(dicTitulos("COD_SIT")) = "08 - Regime Especial ou Norma Específica"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                    Case "Somar valor do campo VL_ICMS_ST ao campo VL_ITEM"
                        Campos(dicTitulos("VL_ITEM")) = CDbl(Campos(dicTitulos("VL_ITEM"))) + CDbl(Campos(dicTitulos("VL_ICMS_ST")))
                        Campos(dicTitulos("VL_BC_ICMS_ST")) = 0
                        Campos(dicTitulos("ALIQ_ST")) = 0
                        Campos(dicTitulos("VL_ICMS_ST")) = 0
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                    
                    Case "Informar ""SIM"" no campo CONTRIBUINTE"
                        Campos(dicTitulos("CONTRIBUINTE")) = "SIM"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                    
                    Case "Informar ""NÃO"" no campo CONTRIBUINTE"
                        Campos(dicTitulos("CONTRIBUINTE")) = "NÃO"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                    Case "Alterar dígitos da Tabela B do CST/ICMS para 00"
                        Campos(dicTitulos("CST_ICMS")) = VBA.Left(Campos(dicTitulos("CST_ICMS")), 1) & "00"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                                                
                    Case "Alterar dígitos da Tabela B do CST/ICMS para 40", "Informar o CST 40 da tabela B para o campo CST_ICMS"
                        Campos(dicTitulos("CST_ICMS")) = VBA.Left(Campos(dicTitulos("CST_ICMS")), 1) & "40"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                    Case "Alterar dígitos da Tabela B do CST/ICMS para 41", "Informar o CST 41 da tabela B para o campo CST_ICMS"
                        Campos(dicTitulos("CST_ICMS")) = VBA.Left(Campos(dicTitulos("CST_ICMS")), 1) & "41"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                    Case "Alterar dígitos da Tabela B do CST/ICMS para 60", "Informar o CST 60 da tabela B para o campo CST_ICMS"
                        Campos(dicTitulos("CST_ICMS")) = VBA.Left(Campos(dicTitulos("CST_ICMS")), 1) & "60"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                        
                    Case "Alterar dígitos da Tabela B do CST/ICMS para 90", "Informar o CST 90 da tabela B para o campo CST_ICMS"
                        Campos(dicTitulos("CST_ICMS")) = VBA.Left(Campos(dicTitulos("CST_ICMS")), 1) & "90"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                        
                    Case "Mudar o dígito de origem do campo CST_ICMS para 2"
                        Campos(dicTitulos("CST_ICMS")) = "2" & VBA.Right(Campos(dicTitulos("CST_ICMS")), VBA.Len(Campos(dicTitulos("CST_ICMS"))) - 1)
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                        
                    Case "Mudar o dígito de origem do campo CST_ICMS para 7"
                        Campos(dicTitulos("CST_ICMS")) = "7" & VBA.Right(Campos(dicTitulos("CST_ICMS")), VBA.Len(Campos(dicTitulos("CST_ICMS"))) - 1)
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Call ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                        
                    Case "Zerar campos do ICMS"
                        Campos(dicTitulos("VL_BC_ICMS")) = 0
                        Campos(dicTitulos("ALIQ_ICMS")) = 0
                        Campos(dicTitulos("VL_ICMS")) = 0
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        GoTo VERIFICAR:
                        
                    Case "Alterar o valor do campo TIPO_ITEM para 00"
                        Campos(dicTitulos("TIPO_ITEM")) = "00 - Mercadoria para Revenda"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                        
                    Case "Alterar o valor do campo TIPO_ITEM para 07"
                        Campos(dicTitulos("TIPO_ITEM")) = "07 - Material de Uso e Consumo"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                        
                    Case "Alterar o valor do campo TIPO_ITEM para 08"
                        Campos(dicTitulos("TIPO_ITEM")) = "08 - Ativo Imobilizado"
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                        
                    Case "Apagar valor do CEST informado no campo CEST"
                        Campos(dicTitulos("CEST")) = ""
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                        
                    Case "Apagar valor informado no campo COD_NCM"
                        Campos(dicTitulos("COD_NCM")) = ""
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                        
                    Case "Adicionar zeros a esquerda do CEST"
                        Campos(dicTitulos("CEST")) = "'" & VBA.Format(Util.ApenasNumeros(Campos(dicTitulos("CEST"))), VBA.String(7, "0"))
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                        
                    Case "Adicionar zeros a esquerda do campo COD_NCM"
                        Campos(dicTitulos("CEST")) = "'" & VBA.Format(Util.ApenasNumeros(Campos(dicTitulos("CEST"))), VBA.String(7, "0"))
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        
                        
                    Case "Zerar Alíquota do ICMS"
                        Campos(dicTitulos("ALIQ_ICMS")) = 0
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        GoTo VERIFICAR:
                    
                    Case "Recalcular o campo VL_ICMS"
                        Campos(dicTitulos("VL_ICMS")) = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_ICMS")) * Campos(dicTitulos("ALIQ_ICMS")), True, 2)
                        Campos(dicTitulos("INCONSISTENCIA")) = Empty
                        Campos(dicTitulos("SUGESTAO")) = Empty
                        Campos = ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                        GoTo VERIFICAR:

                End Select
                
            End If
            
            If Linha.Row > 3 Then arrDados.Add Campos
            
        End If
        
    Next Linha

    If arrDados.Count = 0 Then
        Call Util.MsgAlerta("Não existem sugestões para processar!", "Sugestões Fiscais")
        Exit Function
    End If
        
    If assApuracaoICMS.AutoFilterMode Then assApuracaoICMS.AutoFilter.ShowAllData
    Call Util.LimparDados(assApuracaoICMS, 4, False)
    Call Util.ExportarDadosArrayList(assApuracaoICMS, arrDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoICMS)
    
    Call Util.MsgInformativa("Sugestões aplicadas com sucesso!", "Inclusão de Sugestões", Inicio)
    Application.StatusBar = False
    
End Function

Public Function IgnorarInconsistencias()

Dim dicTitulos0000 As New Dictionary
Dim dicDados0000 As New Dictionary
Dim Dados As Range, Linha As Range
Dim CHV_REG As String, INCONSISTENCIA$
Dim dicTitulos As New Dictionary
Dim dicDados As New Dictionary
Dim Resposta As VbMsgBoxResult
Dim Campos As Variant

    Resposta = MsgBox("Tem certeza que deseja ignorar as inconsistências selecionadas?" & vbCrLf & _
                      "Essa operação NÃO pode ser desfeita.", vbExclamation + vbYesNo, "Ignorar Inconsistências")
    
    If Resposta = vbNo Then Exit Function

    Inicio = Now()
    Application.StatusBar = "Ignorando as sugestões selecionadas, por favor aguarde..."
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoICMS, 3)
    Set Dados = assApuracaoICMS.Range("A4").CurrentRegion
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não foi encontrada nenhuma inconsistência!", "Inconsistências Fiscais")
        Exit Function
    End If
    
    Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
        
    For Each Linha In Dados.Rows
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
                        
            If Linha.EntireRow.Hidden = False And Campos(dicTitulos("INCONSISTENCIA")) <> "" And Linha.Row > 3 Then
                
                CHV_REG = Campos(dicTitulos("CHV_REG"))
                INCONSISTENCIA = Campos(dicTitulos("INCONSISTENCIA"))
                
                'Verifica se o registro já possui inconsistências ignoradas, caso não exista cria
                If Not dicInconsistenciasIgnoradas.Exists(CHV_REG) Then Set dicInconsistenciasIgnoradas(CHV_REG) = New ArrayList
                
                'Verifica se a inconsistência já foi ignorada e caso contrário adiciona ela na lista
                If Not dicInconsistenciasIgnoradas(CHV_REG).contains(INCONSISTENCIA) Then _
                    dicInconsistenciasIgnoradas(CHV_REG).Add INCONSISTENCIA
                
                Campos(dicTitulos("INCONSISTENCIA")) = Empty
                Campos(dicTitulos("SUGESTAO")) = Empty
                Call ValidarRegrasFiscais(Campos, dicTitulos, dicDados0000, dicTitulos0000)
                
            End If
            
            If Linha.Row > 3 Then dicDados(Campos(dicTitulos("CHV_REG"))) = Campos
            
        End If
        
    Next Linha

    If dicInconsistenciasIgnoradas.Count = 0 Then
        Call Util.MsgAlerta("Não existem Inconsistêncais a ignorar!", "Ignorar Inconsistências")
        Exit Function
    End If
    
    If assApuracaoICMS.AutoFilterMode Then assApuracaoICMS.AutoFilter.ShowAllData
    Call Util.LimparDados(assApuracaoICMS, 4, False)
    Call Util.ExportarDadosDicionario(assApuracaoICMS, dicDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoICMS)
    
    Call Util.MsgInformativa("Inconsistências ignoradas com sucesso!", "Ignorar Inconsistências", Inicio)
    Application.StatusBar = False
    
End Function

Public Sub AtualizarRegistros()
'TODO: refatorar rotina para torná-la mais simples e eficiente
Dim Campos As Variant, Campos0200, CamposC100, CamposC170, CamposC177, CamposC190, dicCampos, regCampo
Dim CHV_REG As String, CHV_C177$, CHV_PAI$, CHV_0200$
Dim dicTitulos0200 As New Dictionary
Dim dicTitulosC100 As New Dictionary
Dim dicTitulosC170 As New Dictionary
Dim dicTitulosC177 As New Dictionary
Dim dicTitulosC190 As New Dictionary
Dim dicDados0200 As New Dictionary
Dim dicDadosC100 As New Dictionary
Dim dicDadosC170 As New Dictionary
Dim dicDadosC177 As New Dictionary
Dim dicDadosC190 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicTitulos As New Dictionary
Dim Status As Boolean
    
'    If Otimizacoes.OtimizacoesAtivas Then
'
'        Status = Otimizacoes.OtimizarAtualizacaoRegistros(assApuracaoICMS)
'        If Not Status Then Exit Sub
'
'    End If
    
    Inicio = Now()
    Util.AtualizarBarraStatus ("Preparando dados para atualização do SPED, por favor aguarde...")
    
    Campos0200 = Array("REG", "COD_BARRA", "COD_NCM", "EX_IPI", "CEST", "TIPO_ITEM")
    CamposC100 = Array("CHV_NFE", "NUM_DOC", "SER", "COD_PART")
    CamposC170 = Array("COD_NAT", "IND_MOV", "CFOP", "VL_ITEM", "CST_ICMS", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "ALIQ_ST", "VL_ICMS_ST")
    CamposC177 = Array("COD_INF_ITEM")
    CamposC190 = Array("CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "VL_ICMS_ST")
    
    Set dicTitulos0200 = Util.MapearTitulos(reg0200, 3)
    Set dicDados0200 = Util.CriarDicionarioRegistro(reg0200, "ARQUIVO", "COD_ITEM")
    
    Set dicTitulosC100 = Util.MapearTitulos(regC100, 3)
    Set dicDadosC100 = Util.CriarDicionarioRegistro(regC100)
    
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    Set dicDadosC170 = Util.CriarDicionarioRegistro(regC170)
    
    Set dicTitulosC190 = Util.MapearTitulos(regC190, 3)
    Set dicDadosC190 = Util.CriarDicionarioRegistro(regC190)
        
    Set dicTitulos = Util.MapearTitulos(assApuracaoICMS, 3)
    If assApuracaoICMS.AutoFilterMode Then assApuracaoICMS.AutoFilter.ShowAllData
    
    Set Dados = Util.DefinirIntervalo(assApuracaoICMS, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    a = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(a, 10, "Preparando dados para atualização do SPED, por favor aguarde...", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            CHV_0200 = Util.UnirCampos(CStr(Campos(dicTitulos("ARQUIVO"))), CStr(Campos(dicTitulos("COD_ITEM"))))
            CHV_PAI = Campos(dicTitulos("CHV_PAI_FISCAL"))
            CHV_REG = Campos(dicTitulos("CHV_REG"))
            
            'Atualizar dados do 0200
            If dicDados0200.Exists(CHV_0200) Then
                
                dicCampos = dicDados0200(CHV_0200)
                For Each regCampo In Campos0200
                    
                    If regCampo = "CEST" Or regCampo = "COD_BARRA" Or regCampo = "COD_NCM" Or regCampo = "EX_TIPI" Then
                        Campos(dicTitulos(regCampo)) = Util.FormatarTexto(Campos(dicTitulos(regCampo)))
                    End If
                    
                    If regCampo = "REG" Then Campos(dicTitulos(regCampo)) = "'0200"
                    dicCampos(dicTitulos0200(regCampo)) = Campos(dicTitulos(regCampo))
                    
                Next regCampo
                
                dicDados0200(CHV_0200) = dicCampos
            
            End If
            
            'Atualizar dados do C100
            If dicDadosC100.Exists(CHV_PAI) Then
                
                dicCampos = dicDadosC100(CHV_PAI)
                For Each regCampo In CamposC100
                    
                    dicCampos(dicTitulosC100(regCampo)) = Campos(dicTitulos(regCampo))
                    
                Next regCampo
                
                dicDadosC100(CHV_PAI) = dicCampos
            
            End If
                
            'Atualizar dados do C170
            If dicDadosC170.Exists(CHV_REG) Then
                
                dicCampos = dicDadosC170(CHV_REG)
                For Each regCampo In CamposC170
                    
                    If regCampo = "CST_ICMS" Or regCampo = "COD_BARRA" Then
                        Campos(dicTitulos(regCampo)) = fnExcel.FormatarTexto(Campos(dicTitulos(regCampo)))
                    End If
                    
                    If regCampo Like "VL_*" Then Campos(dicTitulos(regCampo)) = VBA.Round(Campos(dicTitulos(regCampo)), 2)
                    dicCampos(dicTitulosC170(regCampo)) = Campos(dicTitulos(regCampo))
                
                Next regCampo
                
                dicDadosC170(CHV_REG) = dicCampos
            
            End If
            
            'Atualizar dados do C177
            CHV_C177 = fnSPED.GerarChaveRegistro(CHV_REG, "C177")
            If dicDadosC177.Exists(CHV_PAI) Then
                
                dicCampos = dicDadosC177(CHV_C177)
                For Each regCampo In CamposC177

                    dicCampos(dicTitulosC177(regCampo)) = Campos(dicTitulos(regCampo))
                    
                Next regCampo
                
                dicDadosC177(CHV_C177) = dicCampos
                
            ElseIf Campos(dicTitulos("COD_INF_ITEM")) <> "" Then
                
                dicCampos = Array("C177", Campos(dicTitulos("ARQUIVO")), CHV_C177, CHV_REG, "", Campos(dicTitulos("COD_INF_ITEM")))
                dicDadosC177(CHV_C177) = dicCampos
                
            End If
            
            'Atualizar dados do C190
            If dicDadosC190.Exists(CHV_REG) Then
                
                dicCampos = dicDadosC190(CHV_REG)
                For Each regCampo In CamposC190

                    Select Case True
                        
                        Case regCampo Like "CFOP"
                            dicCampos(dicTitulosC190(regCampo)) = Campos(dicTitulos(regCampo))
                            
                        Case regCampo Like "CST_ICMS"
                            dicCampos(dicTitulosC190(regCampo)) = fnExcel.FormatarTexto(Campos(dicTitulos(regCampo)))
                            
                        Case regCampo Like "VL_OPR"
                            dicCampos(dicTitulosC190(regCampo)) = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM")), True, 2)
                            
                        Case regCampo Like "VL_*"
                            dicCampos(dicTitulosC190(regCampo)) = fnExcel.ConverterValores(Campos(dicTitulos(regCampo)), True, 2)
                            
                    End Select
                    
                Next regCampo
                
                dicDadosC190(CHV_REG) = dicCampos
            
            End If
            
        End If

    Next Linha
    
    Util.AtualizarBarraStatus ("Atualizando dados do registro 0200, por favor aguarde...")
    Call Util.LimparDados(reg0200, 4, False)
    Call Util.ExportarDadosDicionario(reg0200, dicDados0200)
    
    Util.AtualizarBarraStatus ("Atualizando dados do registro C100, por favor aguarde...")
    Call Util.LimparDados(regC100, 4, False)
    Call Util.ExportarDadosDicionario(regC100, dicDadosC100)
    
    Util.AtualizarBarraStatus ("Atualizando dados do registro C170, por favor aguarde...")
    Call Util.LimparDados(regC170, 4, False)
    Call Util.ExportarDadosDicionario(regC170, dicDadosC170)
        
    Util.AtualizarBarraStatus ("Atualizando dados do registro C177, por favor aguarde...")
    Call Util.LimparDados(regC177, 4, False)
    Call Util.ExportarDadosDicionario(regC177, dicDadosC177)
    
    Util.AtualizarBarraStatus ("Atualizando dados do registro C190, por favor aguarde...")
    Call Util.LimparDados(regC190, 4, False)
    Call Util.ExportarDadosDicionario(regC190, dicDadosC190)
    Call rC190.AgruparRegistros(True)
    Call rC190.AtualizarImpostosC100(True)
    Call rC170.GerarC190(True)
    
    Util.AtualizarBarraStatus ("Atualizando valores dos impostos no registro C100, por favor aguarde...")
    Call rC170.AtualizarImpostosC100(True)
    Call r0200.AtualizarCodigoGenero(True)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoICMS)
        
    Util.AtualizarBarraStatus ("Atualização concluída com sucesso!")
    Call Util.MsgInformativa("Registros atualizados com sucesso!", "Atualização de dados", Inicio)
    Util.AtualizarBarraStatus (False)
    
End Sub

Public Function ListarTributacoesICMS()

Dim Tributacao As New AssistenteTributario
        
    Call Tributacao.ListarTributacoes(assApuracaoICMS, assTributacaoICMS)
    
End Function

Public Sub CalcularDifalNaoContribuinte()

Dim assDIFAL As New AssistenteDIFALNaoContribuinte
    
    Inicio = Now()
    assDIFAL.CalcularDifalNaoContribuinte
    
End Sub
