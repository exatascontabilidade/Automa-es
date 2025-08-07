Attribute VB_Name = "clsAssistenteApuracaoPISCOFINS"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private ValidacoesGerais As New clsAssistenteApuracao_Regras
Private ValidacoesCFOP As New clsRegrasFiscaisCFOP
Private ValidacoesNCM As New clsRegrasFiscaisNCM
Public dicDiferencaC181C185 As New Dictionary
Public dicDiferencaC191C195 As New Dictionary
Private GerenciadorSPED As clsRegistrosSPED
Private ExpReg As ExportadorRegistros
Private arrRelatorio As New ArrayList
Private dicTitulos As New Dictionary

Public Function GerarApuracaoAssistidaPISCOFINS()

Dim Msg As String
    
    Inicio = Now()
    
    Call Util.DesabilitarControles
    
    Call arrRelatorio.Clear
    Call dicInconsistenciasIgnoradas.RemoveAll
    Call DadosValidacaoCFOP.CarregarTitulosRelatorio(ActiveSheet)
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoPISCOFINS, 3)
    
    Call CarregarDadosA170(Msg)
    Call CarregarDadosC170(Msg)
    Call CarregarDadosC175(Msg)
    Call CarregarDadosC181(Msg)
    Call CarregarDadosC185(Msg)
    Call CarregarDadosF100(Msg)
    
    Application.StatusBar = "Processo concluído com sucesso!"
    If arrRelatorio.Count > 0 Then
        
        Call RegistrarDivergenciasC181C185
        
        On Error Resume Next
            If assApuracaoPISCOFINS.AutoFilter.FilterMode Then assApuracaoPISCOFINS.ShowAllData
        On Error GoTo 0
        Call Util.LimparDados(assApuracaoPISCOFINS, 4, False)
        
        Call Util.ExportarDadosArrayList(assApuracaoPISCOFINS, arrRelatorio)
        Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoPISCOFINS)
        Call Util.MsgInformativa("Relatório gerado com sucesso", "Assistente de Apuração do PIS/COFINS", Inicio)
        
    Else
        
        Msg = "Nenhum dado encontrado para geração do relatório." & vbCrLf & vbCrLf
        Msg = Msg & "Por favor verifique se o SPED foi importado e tente novamente."
        Call Util.MsgAlerta(Msg, "Assistente de Apuração do PIS/COFINS")
        
    End If
    
    Call Util.AtualizarBarraStatus(False)
    Call Util.HabilitarControles
    
End Function

Public Sub CarregarDadosA170(ByVal Msg As String)

Dim REG As String, ARQUIVO$, CHV_REG$, CHV_0140$, CHV_0150$, CHV_A100$, COD_ITEM$, COD_PART$, COD_INC_TRIB$, CNPJ_ESTABELECIMENTO$, UF_CONTRIB$
Dim Apuracao As New clsAssistenteApuracao
Dim dicTitulosA170 As New Dictionary
Dim Dados As Range, Linha As Range
Dim Comeco As Double, VL_ITEM#
Dim Campos As Variant
Dim b As Long
        
    Set Dados = Util.DefinirIntervalo(regA170, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    Set Apuracao.dicTitulos = dicTitulos
    Set dicTitulosA170 = Util.MapearTitulos(regA170, 3)
    
    b = 0
    Comeco = Timer
    With Apuracao
        
        Call .CarregarDadosRegistro0000
        
        For Each Linha In Dados.Rows
            
            .RedimensionarArray (dicTitulos.Count)
            Call Util.AntiTravamento(b, 10, Msg & "Carregando dados do registro A170", Dados.Rows.Count, Comeco)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                'Carrega as variáveis necessárias
                REG = Campos(dicTitulosA170("REG"))
                ARQUIVO = Campos(dicTitulosA170("ARQUIVO"))
                CHV_A100 = Campos(dicTitulosA170("CHV_PAI_CONTRIBUICOES"))
                CNPJ_ESTABELECIMENTO = .ExtrairCNPJ_ESTABELECIMENTO(REG, CHV_A100, ARQUIVO)
                UF_CONTRIB = .ExtrairUF_Estabelecimento(ARQUIVO, CNPJ_ESTABELECIMENTO)
                COD_PART = .ExtrairCOD_PART_A100(CHV_A100)
                CHV_0140 = .ExtrairCHV_0140(Util.UnirCampos(ARQUIVO, CNPJ_ESTABELECIMENTO))
                CHV_0150 = Util.UnirCampos(ARQUIVO, COD_PART)
                COD_ITEM = Campos(dicTitulosA170("COD_ITEM"))
                VL_ITEM = fnExcel.FormatarValores(Campos(dicTitulosA170("VL_ITEM")))
                COD_INC_TRIB = .ExtrairREGIME_TRIBUTARIO(ARQUIVO, True)
                
                'Extrai dados do registro A100
                Call .ExtrairDadosA100(CHV_A100)
                
                'Extrai dados do registro 0150
                Call .ExtrairDados0150(CHV_0150, COD_PART, True)
                
                'Extrai dados do registro 0200
                Call .ExtrairDados0200(CHV_0140, COD_ITEM, True)
                
                'Atribui valores aos campos do relatório
                .AtribuirValor "REG", REG
                .AtribuirValor "ARQUIVO", ARQUIVO
                .AtribuirValor "CHV_PAI_CONTRIBUICOES", CHV_A100
                .AtribuirValor "CHV_REG", Campos(dicTitulosA170("CHV_REG"))
                .AtribuirValor "CNPJ_ESTABELECIMENTO", fnExcel.FormatarTexto(CNPJ_ESTABELECIMENTO)
                .AtribuirValor "UF_CONTRIB", UF_CONTRIB
                .AtribuirValor "COD_ITEM", fnExcel.FormatarTexto(COD_ITEM)
                .AtribuirValor "VL_ITEM", VL_ITEM
                .AtribuirValor "VL_DESC", fnExcel.ConverterValores(Campos(dicTitulosA170("VL_DESC")))
                .AtribuirValor "CST_PIS", fnExcel.FormatarTexto(Campos(dicTitulosA170("CST_PIS")))
                .AtribuirValor "CST_COFINS", fnExcel.FormatarTexto(Campos(dicTitulosA170("CST_COFINS")))
                .AtribuirValor "VL_BC_PIS", fnExcel.ConverterValores(Campos(dicTitulosA170("VL_BC_PIS")))
                .AtribuirValor "ALIQ_PIS", fnExcel.ConverterValores(Campos(dicTitulosA170("ALIQ_PIS")))
                .AtribuirValor "VL_PIS", fnExcel.ConverterValores(Campos(dicTitulosA170("VL_PIS")))
                .AtribuirValor "VL_BC_COFINS", fnExcel.ConverterValores(Campos(dicTitulosA170("VL_BC_COFINS")))
                .AtribuirValor "ALIQ_COFINS", fnExcel.ConverterValores(Campos(dicTitulosA170("ALIQ_COFINS")))
                .AtribuirValor "VL_COFINS", fnExcel.ConverterValores(Campos(dicTitulosA170("VL_COFINS")))
                .AtribuirValor "COD_CTA", fnExcel.FormatarTexto(Campos(dicTitulosA170("COD_CTA")))
                
            End If
            
            Campos = ValidarRegrasFiscaisPISCOFINS(.Campo, COD_INC_TRIB)
            arrRelatorio.Add Campos
            
        Next Linha
        
    End With
    
End Sub

Public Sub CarregarDadosC175(ByVal Msg As String)

Dim ARQUIVO As String, CHV_REG$, CHV_PAI$, CHV_0140$, CHV_0150$, CHV_C100$, COD_ITEM$, COD_PART$, UF_CONTRIB$, COD_INC_TRIB$, CNPJ_ESTABELECIMENTO$
Dim Apuracao As New clsAssistenteApuracao
Dim dicTitulosC175 As New Dictionary
Dim Dados As Range, Linha As Range
Dim arrChavesPai As New ArrayList
Dim Comeco As Double, VL_ITEM#
Dim Campos As Variant
Dim b As Long
        
    Set Dados = Util.DefinirIntervalo(regC175_Contr, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    Set Apuracao.dicTitulos = dicTitulos
    Set dicTitulosC175 = Util.MapearTitulos(regC175_Contr, 3)
    
    b = 0
    Comeco = Timer
    With Apuracao
        
        Call .CarregarDadosRegistro0000
        Set arrChavesPai = Util.ListarValoresUnicos(regC170, 4, 3, "CHV_PAI_CONTRIBUICOES")
        
        For Each Linha In Dados.Rows
            
            .RedimensionarArray (dicTitulos.Count)
            Call Util.AntiTravamento(b, 50, Msg & "Carregando dados do registro C175", Dados.Rows.Count, Comeco)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                CHV_PAI = Campos(dicTitulosC175("CHV_PAI_CONTRIBUICOES"))
                If arrChavesPai.contains(CHV_PAI) Then GoTo Prx: Stop
                
                'Carrega as variáveis necessárias
                ARQUIVO = Campos(dicTitulosC175("ARQUIVO"))
                UF_CONTRIB = .ExtrairUFContribuinte(ARQUIVO, True)
                CHV_C100 = Campos(dicTitulosC175("CHV_PAI_CONTRIBUICOES"))
                CNPJ_ESTABELECIMENTO = .ExtrairCNPJ_ESTABELECIMENTO_C100(CHV_C100)
                UF_CONTRIB = .ExtrairUF_Estabelecimento(ARQUIVO, CNPJ_ESTABELECIMENTO)
                COD_PART = .ExtrairCOD_PART_C100(CHV_C100)
                CHV_0140 = .ExtrairCHV_0140(Util.UnirCampos(ARQUIVO, CNPJ_ESTABELECIMENTO))
                CHV_0150 = Util.UnirCampos(ARQUIVO, COD_PART)
                VL_ITEM = fnExcel.FormatarValores(Campos(dicTitulosC175("VL_OPER")))
                COD_INC_TRIB = .ExtrairREGIME_TRIBUTARIO(ARQUIVO, True)
                
                'Extrai dados do registro C100
                Call .ExtrairDadosC100(CHV_C100)
                
                'Extrai dados do registro 0150
                Call .ExtrairDados0150(CHV_0150, COD_PART, True)
                
                'Extrai dados do registro 0200
                Call .ExtrairDados0200(CHV_0140, COD_ITEM, True)
                
                'Atribui valores aos campos do relatório
                .AtribuirValor "DESCR_ITEM", "O REGISTRO C175 NÃO POSSUI DADOS DO PRODUTO"
                .AtribuirValor "NOME_RAZAO", "O REGISTRO C175 NÃO POSSUI DADOS DO PARTICIPANTE"
                .AtribuirValor "REG", Campos(dicTitulosC175("REG"))
                .AtribuirValor "ARQUIVO", ARQUIVO
                .AtribuirValor "CHV_PAI_CONTRIBUICOES", CHV_C100
                .AtribuirValor "CHV_REG", Campos(dicTitulosC175("CHV_REG"))
                .AtribuirValor "CNPJ_ESTABELECIMENTO", fnExcel.FormatarTexto(CNPJ_ESTABELECIMENTO)
                .AtribuirValor "UF_CONTRIB", UF_CONTRIB
                .AtribuirValor "UF_PART", .ExtrairUF_PART(CHV_0150)
                .AtribuirValor "CFOP", Campos(dicTitulosC175("CFOP"))
                .AtribuirValor "VL_ITEM", VL_ITEM
                .AtribuirValor "VL_DESP", .ExtrairVL_DESP_C100(CHV_C100, VL_ITEM)
                .AtribuirValor "VL_DESC", fnExcel.ConverterValores(Campos(dicTitulosC175("VL_DESC")))
                .AtribuirValor "CST_PIS", fnExcel.FormatarTexto(Campos(dicTitulosC175("CST_PIS")))
                .AtribuirValor "CST_COFINS", fnExcel.FormatarTexto(Campos(dicTitulosC175("CST_COFINS")))
                .AtribuirValor "VL_BC_PIS", fnExcel.ConverterValores(Campos(dicTitulosC175("VL_BC_PIS")))
                .AtribuirValor "ALIQ_PIS", fnExcel.ConverterValores(Campos(dicTitulosC175("ALIQ_PIS")))
                .AtribuirValor "QUANT_BC_PIS", Campos(dicTitulosC175("QUANT_BC_PIS"))
                .AtribuirValor "ALIQ_PIS_QUANT", Campos(dicTitulosC175("ALIQ_PIS_QUANT"))
                .AtribuirValor "VL_PIS", fnExcel.ConverterValores(Campos(dicTitulosC175("VL_PIS")))
                .AtribuirValor "VL_BC_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC175("VL_BC_COFINS")))
                .AtribuirValor "ALIQ_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC175("ALIQ_COFINS")))
                .AtribuirValor "QUANT_BC_COFINS", Campos(dicTitulosC175("QUANT_BC_COFINS"))
                .AtribuirValor "ALIQ_COFINS_QUANT", Campos(dicTitulosC175("ALIQ_COFINS_QUANT"))
                .AtribuirValor "VL_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC175("VL_COFINS")))
                .AtribuirValor "COD_CTA", fnExcel.FormatarTexto(Campos(dicTitulosC175("COD_CTA")))
                
            End If
            
            Campos = ValidarRegrasFiscaisPISCOFINS(.Campo, COD_INC_TRIB)
            arrRelatorio.Add Campos
Prx:
        Next Linha
        
    End With
    
End Sub

Public Sub CarregarDadosC170(ByVal Msg As String)

Dim ARQUIVO As String, REG$, CHV_REG$, CHV_0140$, CHV_C100$, COD_ITEM$, COD_PART$, COD_INC_TRIB$, CNPJ_ESTABELECIMENTO$, UF_CONTRIB$
Dim Apuracao As New clsAssistenteApuracao
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
        
        Call .CarregarDadosRegistro0000(True)
        
        For Each Linha In Dados.Rows
            
            .RedimensionarArray (dicTitulos.Count)
            Call Util.AntiTravamento(b, 10, Msg & "Carregando dados do registro C170", Dados.Rows.Count, Comeco)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                'Carrega as variáveis necessárias
                REG = Campos(dicTitulosC170("REG"))
                ARQUIVO = Campos(dicTitulosC170("ARQUIVO"))
                CHV_C100 = Campos(dicTitulosC170("CHV_PAI_CONTRIBUICOES"))
                CNPJ_ESTABELECIMENTO = .ExtrairCNPJ_ESTABELECIMENTO(REG, CHV_C100, ARQUIVO)
                CHV_0140 = .ExtrairCHV_0140(Util.UnirCampos(ARQUIVO, CNPJ_ESTABELECIMENTO))
                .UFContrib = .ExtrairUF_Estabelecimento(ARQUIVO, CNPJ_ESTABELECIMENTO)
                COD_PART = .ExtrairCOD_PART_C100(CHV_C100)
                COD_ITEM = Campos(dicTitulosC170("COD_ITEM"))
                VL_ITEM = fnExcel.FormatarValores(Campos(dicTitulosC170("VL_ITEM")))
                COD_INC_TRIB = .ExtrairREGIME_TRIBUTARIO(ARQUIVO, True)
                
                'Extrai dados do registro C100
                Call .ExtrairDadosC100(CHV_C100)
                
                'Extrai dados do registro 0150
                Call .ExtrairDados0150(CHV_0140, COD_PART, True)
                                
                'Extrai dados do registro 0200
                Call .ExtrairDados0200(CHV_0140, COD_ITEM, True)
                
                'Atribui valores aos campos do relatório
                .AtribuirValor "REG", REG
                .AtribuirValor "ARQUIVO", ARQUIVO
                .AtribuirValor "CHV_PAI_CONTRIBUICOES", CHV_C100
                .AtribuirValor "CHV_REG", Campos(dicTitulosC170("CHV_REG"))
                .AtribuirValor "CNPJ_ESTABELECIMENTO", fnExcel.FormatarTexto(CNPJ_ESTABELECIMENTO)
                .AtribuirValor "UF_CONTRIB", .UFContrib
                .AtribuirValor "COD_ITEM", fnExcel.FormatarTexto(COD_ITEM)
                .AtribuirValor "IND_MOV", Campos(dicTitulosC170("IND_MOV"))
                .AtribuirValor "CFOP", Campos(dicTitulosC170("CFOP"))
                .AtribuirValor "VL_ITEM", VL_ITEM
                .AtribuirValor "VL_DESP", .ExtrairVL_DESP_C100(CHV_C100, VL_ITEM)
                .AtribuirValor "VL_DESC", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_DESC")))
                .AtribuirValor "VL_ICMS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_ICMS")))
                .AtribuirValor "CST_PIS", fnExcel.FormatarTexto(Campos(dicTitulosC170("CST_PIS")))
                .AtribuirValor "CST_COFINS", fnExcel.FormatarTexto(Campos(dicTitulosC170("CST_COFINS")))
                .AtribuirValor "VL_BC_PIS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_BC_PIS")))
                .AtribuirValor "ALIQ_PIS", fnExcel.ConverterValores(Campos(dicTitulosC170("ALIQ_PIS")))
                .AtribuirValor "QUANT_BC_PIS", IIf(Campos(dicTitulosC170("QUANT_BC_PIS")) = 0, "", Campos(dicTitulosC170("QUANT_BC_PIS")))
                .AtribuirValor "ALIQ_PIS_QUANT", IIf(Campos(dicTitulosC170("ALIQ_PIS_QUANT")) = 0, "", Campos(dicTitulosC170("ALIQ_PIS_QUANT")))
                .AtribuirValor "VL_PIS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_PIS")))
                .AtribuirValor "VL_BC_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_BC_COFINS")))
                .AtribuirValor "ALIQ_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC170("ALIQ_COFINS")))
                .AtribuirValor "QUANT_BC_COFINS", IIf(Campos(dicTitulosC170("QUANT_BC_COFINS")) = 0, "", Campos(dicTitulosC170("QUANT_BC_COFINS")))
                .AtribuirValor "ALIQ_COFINS_QUANT", IIf(Campos(dicTitulosC170("ALIQ_COFINS_QUANT")) = 0, "", Campos(dicTitulosC170("ALIQ_COFINS_QUANT")))
                .AtribuirValor "VL_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_COFINS")))
                .AtribuirValor "COD_CTA", fnExcel.FormatarTexto(Campos(dicTitulosC170("COD_CTA")))
                
            End If
            
            Campos = ValidarRegrasFiscaisPISCOFINS(.Campo, COD_INC_TRIB)
            arrRelatorio.Add Campos
            
        Next Linha
        
    End With
    
    Set Apuracao = Nothing
    
End Sub

Public Sub CarregarDadosC181(ByVal Msg As String)

Dim REG As String, ARQUIVO$, CHV_REG$, CHV_0140$, CHV_C180$, COD_ITEM$, COD_INC_TRIB$, CNPJ_ESTABELECIMENTO$, UF_CONTRIB$
Dim Apuracao As New clsAssistenteApuracao
Dim dicTitulosC181 As New Dictionary
Dim Dados As Range, Linha As Range
Dim Comeco As Double, VL_ITEM#
Dim Campos As Variant
Dim b As Long
    
    Set Dados = Util.DefinirIntervalo(regC181_Contr, 4, 3)
    If Dados Is Nothing Then Exit Sub
        
    Set Apuracao.dicTitulos = dicTitulos
    Set dicDiferencaC181C185 = New Dictionary
    Set dicTitulosC181 = Util.MapearTitulos(regC181_Contr, 3)
    
    b = 0
    Comeco = Timer
    With Apuracao
        
        Call .CarregarDadosRegistro0000
        
        For Each Linha In Dados.Rows
            
            .RedimensionarArray (dicTitulos.Count)
            Call Util.AntiTravamento(b, 10, Msg & "Carregando dados do registro C181", Dados.Rows.Count, Comeco)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                'Carrega as variáveis necessárias
                REG = Campos(dicTitulosC181("REG"))
                ARQUIVO = Campos(dicTitulosC181("ARQUIVO"))
                UF_CONTRIB = .ExtrairUFContribuinte(ARQUIVO, True)
                CHV_0140 = .ExtrairCHV_0140(Util.UnirCampos(ARQUIVO, CNPJ_ESTABELECIMENTO))
                CHV_C180 = Campos(dicTitulosC181("CHV_PAI_CONTRIBUICOES"))
                CNPJ_ESTABELECIMENTO = .ExtrairCNPJ_ESTABELECIMENTO(REG, CHV_C180, ARQUIVO)
                UF_CONTRIB = .ExtrairUF_Estabelecimento(ARQUIVO, CNPJ_ESTABELECIMENTO)
                VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulosC181("VL_ITEM")), True, 2)
                COD_INC_TRIB = .ExtrairREGIME_TRIBUTARIO(ARQUIVO, True)
                
                'Extrai dados do registro C180
                Call .ExtrairDadosC180(CHV_C180)
                
                'Extrai dados do registro 0200
                COD_ITEM = .ExtrairValor("COD_ITEM")
                Call .ExtrairDados0200(CHV_0140, COD_ITEM, True)
                
                'Atribui valores aos campos do relatório
                .AtribuirValor "REG", Campos(dicTitulosC181("REG"))
                .AtribuirValor "ARQUIVO", ARQUIVO
                .AtribuirValor "CHV_PAI_CONTRIBUICOES", CHV_C180
                .AtribuirValor "CHV_REG", Campos(dicTitulosC181("CHV_REG"))
                .AtribuirValor "CHV_NFE", "O REGISTRO C181 NÃO POSSUI DADOS DO DOCUMENTO FISCAL"
                .AtribuirValor "NOME_RAZAO", "O REGISTRO C181 NÃO INFORMA DADOS DO PARTICIPANTE"
                .AtribuirValor "CNPJ_ESTABELECIMENTO", fnExcel.FormatarTexto(CNPJ_ESTABELECIMENTO)
                .AtribuirValor "UF_CONTRIB", UF_CONTRIB
                .AtribuirValor "CFOP", Campos(dicTitulosC181("CFOP"))
                .AtribuirValor "VL_ITEM", Campos(dicTitulosC181("VL_ITEM"))
                .AtribuirValor "UF_PART", UF_CONTRIB
                .AtribuirValor "VL_DESC", fnExcel.ConverterValores(Campos(dicTitulosC181("VL_DESC")))
                .AtribuirValor "CST_PIS", fnExcel.FormatarTexto(Campos(dicTitulosC181("CST_PIS")))
                .AtribuirValor "VL_BC_PIS", fnExcel.ConverterValores(Campos(dicTitulosC181("VL_BC_PIS")))
                .AtribuirValor "ALIQ_PIS", fnExcel.ConverterValores(Campos(dicTitulosC181("ALIQ_PIS")))
                .AtribuirValor "QUANT_BC_PIS", Campos(dicTitulosC181("QUANT_BC_PIS"))
                .AtribuirValor "ALIQ_PIS_QUANT", Campos(dicTitulosC181("ALIQ_PIS_QUANT"))
                .AtribuirValor "VL_PIS", fnExcel.ConverterValores(Campos(dicTitulosC181("VL_PIS")))
                .AtribuirValor "COD_CTA", fnExcel.FormatarTexto(Campos(dicTitulosC181("COD_CTA")))
                
            End If
            
            ProcessarValoresC181 .Campo
            
            Campos = ValidarRegrasFiscaisPISCOFINS(.Campo, COD_INC_TRIB)
            arrRelatorio.Add Campos
            
        Next Linha
        
    End With
    
End Sub

Public Sub CarregarDadosC185(ByVal Msg As String)

Dim REG As String, ARQUIVO$, CHV_REG$, CHV_0140$, CHV_C180$, COD_ITEM$, COD_INC_TRIB$, CNPJ_ESTABELECIMENTO$, UF_CONTRIB$
Dim Apuracao As New clsAssistenteApuracao
Dim dicTitulosC185 As New Dictionary
Dim Dados As Range, Linha As Range
Dim Comeco As Double, VL_ITEM#
Dim Campos As Variant
Dim b As Long
    
    Set Dados = Util.DefinirIntervalo(regC185_Contr, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    Set Apuracao.dicTitulos = dicTitulos
    Set dicTitulosC185 = Util.MapearTitulos(regC185_Contr, 3)
    
    b = 0
    Comeco = Timer
    With Apuracao
        
        Call .CarregarDadosRegistro0000
        
        For Each Linha In Dados.Rows
            
            .RedimensionarArray (dicTitulos.Count)
            Call Util.AntiTravamento(b, 10, Msg & "Carregando dados do registro C185", Dados.Rows.Count, Comeco)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                'Carrega as variáveis necessárias
                REG = Campos(dicTitulosC185("REG"))
                ARQUIVO = Campos(dicTitulosC185("ARQUIVO"))
                CHV_0140 = .ExtrairCHV_0140(Util.UnirCampos(ARQUIVO, CNPJ_ESTABELECIMENTO))
                CHV_C180 = Campos(dicTitulosC185("CHV_PAI_CONTRIBUICOES"))
                CNPJ_ESTABELECIMENTO = .ExtrairCNPJ_ESTABELECIMENTO(REG, CHV_C180, ARQUIVO)
                UF_CONTRIB = .ExtrairUF_Estabelecimento(ARQUIVO, CNPJ_ESTABELECIMENTO)
                VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulosC185("VL_ITEM")), True, 2)
                COD_INC_TRIB = .ExtrairREGIME_TRIBUTARIO(ARQUIVO, True)
                
                'Extrai dados do registro C180
                Call .ExtrairDadosC180(CHV_C180)
                
                'Extrai dados do registro 0200
                COD_ITEM = .ExtrairValor("COD_ITEM")
                Call .ExtrairDados0200(CHV_0140, COD_ITEM, True)
                
                'Atribui valores aos campos do relatório
                .AtribuirValor "REG", Campos(dicTitulosC185("REG"))
                .AtribuirValor "ARQUIVO", ARQUIVO
                .AtribuirValor "CHV_PAI_CONTRIBUICOES", CHV_C180
                .AtribuirValor "CHV_REG", Campos(dicTitulosC185("CHV_REG"))
                .AtribuirValor "CHV_NFE", "O REGISTRO C185 NÃO POSSUI DADOS DO DOCUMENTO FISCAL"
                .AtribuirValor "NOME_RAZAO", "O REGISTRO C185 NÃO INFORMA DADOS DO PARTICIPANTE"
                .AtribuirValor "CNPJ_ESTABELECIMENTO", fnExcel.FormatarTexto(CNPJ_ESTABELECIMENTO)
                .AtribuirValor "UF_CONTRIB", UF_CONTRIB
                .AtribuirValor "CFOP", Campos(dicTitulosC185("CFOP"))
                .AtribuirValor "VL_ITEM", Campos(dicTitulosC185("VL_ITEM"))
                .AtribuirValor "UF_PART", UF_CONTRIB
                .AtribuirValor "VL_DESC", fnExcel.ConverterValores(Campos(dicTitulosC185("VL_DESC")))
                .AtribuirValor "CST_COFINS", fnExcel.FormatarTexto(Campos(dicTitulosC185("CST_COFINS")))
                .AtribuirValor "VL_BC_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC185("VL_BC_COFINS")))
                .AtribuirValor "ALIQ_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC185("ALIQ_COFINS")))
                .AtribuirValor "QUANT_BC_COFINS", Campos(dicTitulosC185("QUANT_BC_COFINS"))
                .AtribuirValor "ALIQ_COFINS_QUANT", Campos(dicTitulosC185("ALIQ_COFINS_QUANT"))
                .AtribuirValor "VL_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC185("VL_COFINS")))
                .AtribuirValor "COD_CTA", fnExcel.FormatarTexto(Campos(dicTitulosC185("COD_CTA")))
                
            End If
            
            ProcessarValoresC185 .Campo
            
            Campos = ValidarRegrasFiscaisPISCOFINS(.Campo, COD_INC_TRIB)
            arrRelatorio.Add Campos
            
        Next Linha
        
    End With
    
End Sub

Public Sub CarregarDadosC191(ByVal Msg As String)

Dim REG As String, ARQUIVO$, CHV_REG$, CHV_0140$, CHV_C190$, COD_ITEM$, COD_INC_TRIB$, CNPJ_ESTABELECIMENTO$, UF_CONTRIB$
Dim Apuracao As New clsAssistenteApuracao
Dim dicTitulosC191 As New Dictionary
Dim Dados As Range, Linha As Range
Dim Comeco As Double, VL_ITEM#
Dim Campos As Variant
Dim b As Long
    
    Set Dados = Util.DefinirIntervalo(regC191_Contr, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    Set Apuracao.dicTitulos = dicTitulos
    Set dicDiferencaC191C195 = New Dictionary
    Set dicTitulosC191 = Util.MapearTitulos(regC191_Contr, 3)
    
    b = 0
    Comeco = Timer
    With Apuracao
        
        Call .CarregarDadosRegistro0000
        
        For Each Linha In Dados.Rows
            
            .RedimensionarArray (dicTitulos.Count)
            Call Util.AntiTravamento(b, 10, Msg & "Carregando dados do registro C191", Dados.Rows.Count, Comeco)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                'Carrega as variáveis necessárias
                REG = Campos(dicTitulosC191("REG"))
                ARQUIVO = Campos(dicTitulosC191("ARQUIVO"))
                CHV_0140 = .ExtrairCHV_0140(Util.UnirCampos(ARQUIVO, CNPJ_ESTABELECIMENTO))
                CHV_C190 = Campos(dicTitulosC191("CHV_PAI_CONTRIBUICOES"))
                CNPJ_ESTABELECIMENTO = .ExtrairCNPJ_ESTABELECIMENTO(REG, CHV_C190, ARQUIVO)
                UF_CONTRIB = .ExtrairUF_Estabelecimento(ARQUIVO, CNPJ_ESTABELECIMENTO)
                VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulosC191("VL_ITEM")), True, 2)
                COD_INC_TRIB = .ExtrairREGIME_TRIBUTARIO(ARQUIVO, True)
                
                'Extrai dados do registro C190
                'Call .ExtrairDadosC190(CHV_C190)
                
                'Extrai dados do registro 0200
                COD_ITEM = .ExtrairValor("COD_ITEM")
                Call .ExtrairDados0200(CHV_0140, COD_ITEM, True)
                
                'Atribui valores aos campos do relatório
                .AtribuirValor "REG", Campos(dicTitulosC191("REG"))
                .AtribuirValor "ARQUIVO", ARQUIVO
                .AtribuirValor "CHV_PAI_CONTRIBUICOES", CHV_C190
                .AtribuirValor "CHV_REG", Campos(dicTitulosC191("CHV_REG"))
                .AtribuirValor "CHV_NFE", "O REGISTRO C191 NÃO POSSUI DADOS DO DOCUMENTO FISCAL"
                .AtribuirValor "NOME_RAZAO", "O REGISTRO C191 NÃO INFORMA DADOS DO PARTICIPANTE"
                .AtribuirValor "CNPJ_ESTABELECIMENTO", fnExcel.FormatarTexto(CNPJ_ESTABELECIMENTO)
                .AtribuirValor "UF_CONTRIB", UF_CONTRIB
                .AtribuirValor "CFOP", Campos(dicTitulosC191("CFOP"))
                .AtribuirValor "VL_ITEM", Campos(dicTitulosC191("VL_ITEM"))
                .AtribuirValor "VL_DESC", Campos(dicTitulosC191("VL_DESC"))
                .AtribuirValor "UF_PART", UF_CONTRIB
                .AtribuirValor "VL_DESC", fnExcel.ConverterValores(Campos(dicTitulosC191("VL_DESC")))
                .AtribuirValor "CST_PIS", fnExcel.FormatarTexto(Campos(dicTitulosC191("CST_PIS")))
                .AtribuirValor "VL_BC_PIS", fnExcel.ConverterValores(Campos(dicTitulosC191("VL_BC_PIS")))
                .AtribuirValor "ALIQ_PIS", fnExcel.ConverterValores(Campos(dicTitulosC191("ALIQ_PIS")))
                .AtribuirValor "QUANT_BC_PIS", Campos(dicTitulosC191("QUANT_BC_PIS"))
                .AtribuirValor "ALIQ_PIS_QUANT", Campos(dicTitulosC191("ALIQ_PIS_QUANT"))
                .AtribuirValor "VL_PIS", fnExcel.ConverterValores(Campos(dicTitulosC191("VL_PIS")))
                .AtribuirValor "COD_CTA", fnExcel.FormatarTexto(Campos(dicTitulosC191("COD_CTA")))
                
            End If
            
            'ProcessarValoresC191 .Campo
            
            Campos = ValidarRegrasFiscaisPISCOFINS(.Campo, COD_INC_TRIB)
            arrRelatorio.Add Campos
            
        Next Linha
        
    End With
    
End Sub

Public Function CarregarDadosC195(ByRef dicDados0200 As Dictionary, ByRef dicTitulos0200 As Dictionary, _
    ByRef arrRelatorio As ArrayList, ByRef dicTitulos As Dictionary, Optional ByVal Msg As String)
    
'TODO: REFATORAR ROTINA PARA FICAR IGUAL A CarregarDadosC185
Dim CHV_REG As String, ARQUIVO$, COD_ITEM$, COD_NCM$, EX_IPI$, CST_PIS$, CST_COFINS$, CHV_CNPJ$, COD_INC_TRIB$, CHV_0110$, CHV_0140$, CHV_PAI$, DESCR_ITEM$
Dim Campos As Variant, Campos0200, CamposC190
Dim dicTitulos0110 As New Dictionary
Dim dicTitulos0140 As New Dictionary
Dim dicTitulosC010 As New Dictionary
Dim dicTitulosC190 As New Dictionary
Dim dicTitulosC195 As New Dictionary
Dim dicDados0110 As New Dictionary
Dim dicDados0140 As New Dictionary
Dim dicDadosC010 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicDadosC190 As New Dictionary
Dim arrDados As New ArrayList
Dim Comeco As Double
Dim b As Long
Dim i As Byte
    
    Set dicTitulos0110 = Util.MapearTitulos(reg0110, 3)
    Set dicDados0110 = Util.CriarDicionarioRegistro(reg0110, "ARQUIVO")
    
    Set dicTitulos0140 = Util.MapearTitulos(reg0140, 3)
    Set dicDados0140 = Util.CriarDicionarioRegistro(reg0140, "ARQUIVO")
    
    Set dicTitulosC010 = Util.MapearTitulos(regC010, 3)
    Set dicDadosC010 = Util.CriarDicionarioRegistro(regC010)
    
    Set dicTitulosC190 = Util.MapearTitulos(regC190_Contr, 3)
    Set dicDadosC190 = Util.CriarDicionarioRegistro(regC190_Contr)
    
    Set dicTitulosC195 = Util.MapearTitulos(regC195_Contr, 3)
    Set Dados = Util.DefinirIntervalo(regC195_Contr, 4, 3)
    
    If Dados Is Nothing Then Exit Function
    
    b = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(b, 10, Msg & "Carregando dados do registro C195", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            ARQUIVO = Campos(dicTitulosC195("ARQUIVO"))
            'COD_INC_TRIB = ExtrairRegimeTributario(dicDados0110, dicTitulos0110, ARQUIVO, COD_INC_TRIB)
                        
            arrDados.Add Campos(dicTitulosC195("REG"))
            arrDados.Add ARQUIVO
            arrDados.Add Campos(dicTitulosC195("CHV_PAI_CONTRIBUICOES"))
            arrDados.Add Campos(dicTitulosC195("CHV_REG"))
            arrDados.Add COD_INC_TRIB
            
            CHV_PAI = Campos(dicTitulosC195("CHV_PAI_CONTRIBUICOES"))
            If dicDadosC190.Exists(CHV_PAI) Then
                
                CamposC190 = dicDadosC190(CHV_PAI)
                If LBound(CamposC190) = 0 Then i = 1 Else i = 0
                
                CHV_CNPJ = CamposC190(dicTitulosC190("CHV_PAI_CONTRIBUICOES") - i)
                If dicDadosC010.Exists(CHV_CNPJ) Then
                    
                    arrDados.Add "'" & dicDadosC010(CHV_CNPJ)(dicTitulosC010("CNPJ"))
                    
                Else
                    
                    arrDados.Add ""
                    
                End If
                
                COD_ITEM = CamposC190(dicTitulosC190("COD_ITEM") - i)
                COD_NCM = CamposC190(dicTitulosC190("COD_NCM") - i)
                EX_IPI = CamposC190(dicTitulosC190("EX_IPI") - i)
                
                If LBound(CamposC190) = 0 Then i = 1 Else i = 0
                'Coleta dados do registro 0200
                
                If dicDados0140.Exists(ARQUIVO) Then CHV_0140 = dicDados0140(ARQUIVO)(dicTitulos0140("CHV_REG") - i)
                
                CHV_REG = fnSPED.GerarChaveRegistro(CHV_0140, COD_ITEM)
                If dicDados0200.Exists(CHV_REG) Then
                    
                    Campos0200 = dicDados0200(CHV_REG)
                    If LBound(Campos0200) = 0 Then i = 1 Else i = 0
                    DESCR_ITEM = Campos0200(dicTitulos0200("DESCR_ITEM") - i)
                    
                End If
                
            Else
                
                arrDados.Add ""
                
            End If
            
            arrDados.Add "" 'CHV_NFE
            arrDados.Add "" 'NUM_DOC
            arrDados.Add "" 'SER
            arrDados.Add COD_ITEM
            arrDados.Add DESCR_ITEM
            arrDados.Add "" 'COD_BARRA
            arrDados.Add COD_NCM
            arrDados.Add EX_IPI
            arrDados.Add "" 'TIPO_ITEM
            arrDados.Add "" 'IND_MOV
            arrDados.Add Campos(dicTitulosC195("CFOP"))
            arrDados.Add Campos(dicTitulosC195("VL_ITEM"))
            arrDados.Add 0 'VL_DESP
            arrDados.Add Campos(dicTitulosC195("VL_DESC"))
            arrDados.Add 0 'VL_ICMS
            arrDados.Add 0 'VL_ICMS_ST
            arrDados.Add "" 'CST_PIS
            arrDados.Add "'" & Campos(dicTitulosC195("CST_COFINS"))
            arrDados.Add "" 'VL_BC_PIS
            arrDados.Add "" 'ALIQ_PIS
            arrDados.Add "" 'QUANT_BC_PIS
            arrDados.Add "" 'ALIQ_PIS_QUANT
            arrDados.Add "" 'VL_PIS
            arrDados.Add Campos(dicTitulosC195("VL_BC_COFINS"))
            arrDados.Add Campos(dicTitulosC195("ALIQ_COFINS"))
            arrDados.Add Campos(dicTitulosC195("QUANT_BC_COFINS"))
            arrDados.Add Campos(dicTitulosC195("ALIQ_COFINS_QUANT"))
            arrDados.Add Campos(dicTitulosC195("VL_COFINS"))
            arrDados.Add Util.FormatarTexto(Campos(dicTitulosC195("COD_CTA")))
            arrDados.Add "" 'COD_NAT_PIS_COFINS
            arrDados.Add Empty 'INCONSISTENCIA
            arrDados.Add Empty 'SUGESTAO
            
        End If
        
        Campos = ValidarRegrasFiscaisPISCOFINS(arrDados.toArray(), COD_INC_TRIB)
        arrRelatorio.Add Campos
        arrDados.Clear
        
    Next Linha
    
End Function

Public Function CarregarDadosD201(ByRef arrRelatorio As ArrayList, ByRef dicTitulos As Dictionary, Optional ByVal Msg As String)
'TODO: REFATORAR ROTINA PARA FICAR IGUAL A CarregarDadosC181
Dim CHV_REG As String, ARQUIVO$, CST_PIS$, CST_COFINS$, CHV_CNPJ$, COD_INC_TRIB$, CHV_0110$, CHV_PAI$
Dim dicTitulos0110 As New Dictionary
Dim dicTitulosD010 As New Dictionary
Dim dicTitulosD200 As New Dictionary
Dim dicTitulosD201 As New Dictionary
Dim dicDados0110 As New Dictionary
Dim dicDadosD010 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicDadosD200 As New Dictionary
Dim arrDados As New ArrayList
Dim Campos As Variant, CamposD200
Dim Comeco As Double
Dim b As Long
Dim i As Byte
    
    Set dicTitulos0110 = Util.MapearTitulos(reg0110, 3)
    Set dicDados0110 = Util.CriarDicionarioRegistro(reg0110, "ARQUIVO")
        
    Set dicTitulosD010 = Util.MapearTitulos(regD010, 3)
    Set dicDadosD010 = Util.CriarDicionarioRegistro(regD010)
        
    Set dicTitulosD200 = Util.MapearTitulos(regD200, 3)
    Set dicDadosD200 = Util.CriarDicionarioRegistro(regD200)
    
    Set dicTitulosD201 = Util.MapearTitulos(regD201, 3)
    Set Dados = Util.DefinirIntervalo(regD201, 4, 3)
    
    If Dados Is Nothing Then Exit Function
    
    b = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(b, 10, Msg & "Carregando dados do registro D201", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            ARQUIVO = Campos(dicTitulosD201("ARQUIVO"))
            'COD_INC_TRIB = ExtrairRegimeTributario(dicDados0110, dicTitulos0110, ARQUIVO, COD_INC_TRIB)
            
            arrDados.Add Campos(dicTitulosD201("REG"))
            arrDados.Add ARQUIVO
            arrDados.Add Campos(dicTitulosD201("CHV_PAI_CONTRIBUICOES"))
            arrDados.Add Campos(dicTitulosD201("CHV_REG"))
            arrDados.Add COD_INC_TRIB
            
            CHV_PAI = Campos(dicTitulosD201("CHV_PAI_CONTRIBUICOES"))
            If dicDadosD200.Exists(CHV_PAI) Then
                
                CamposD200 = dicDadosD200(CHV_PAI)
                If LBound(CamposD200) = 0 Then i = 1 Else i = 0
                
                CHV_CNPJ = CamposD200(dicTitulosD200("CHV_PAI_CONTRIBUICOES") - i)
                If dicDadosD010.Exists(CHV_CNPJ) Then
                    
                    arrDados.Add "'" & dicDadosD010(CHV_CNPJ)(dicTitulosD010("CNPJ"))
                    
                Else
                    
                    arrDados.Add ""
                    
                End If
            
            Else
            
                arrDados.Add ""
            
            End If
            
            arrDados.Add "" 'CHV_NFE
            arrDados.Add "" 'NUM_DOC
            arrDados.Add "" 'SER
            arrDados.Add "" 'COD_ITEM
            arrDados.Add "O REGISTRO D201 POR PADRÃO NÃO IDENTIFICA OS ITENS" 'DESC_ITEM
            arrDados.Add "" 'COD_BARRA
            arrDados.Add "" 'COD_NCM
            arrDados.Add "" 'EX_IPI
            arrDados.Add "" 'TIPO_ITEM
            arrDados.Add "" 'IND_MOV
            arrDados.Add "" 'CFOP
            arrDados.Add Campos(dicTitulosD201("VL_ITEM"))
            arrDados.Add 0 'VL_DESP
            arrDados.Add 0 'VL_DESC
            arrDados.Add 0 'VL_ICMS
            arrDados.Add 0 'VL_ICMS_ST
            arrDados.Add "'" & Campos(dicTitulosD201("CST_PIS"))
            arrDados.Add "" 'CST_COFINS
            arrDados.Add Campos(dicTitulosD201("VL_BC_PIS"))
            arrDados.Add Campos(dicTitulosD201("ALIQ_PIS"))
            arrDados.Add "" 'QUANT_BC_PIS
            arrDados.Add "" 'ALIQ_PIS_QUANT
            arrDados.Add Campos(dicTitulosD201("VL_PIS"))
            arrDados.Add "" 'VL_BC_COFINS
            arrDados.Add "" 'ALIQ_COFINS
            arrDados.Add "" 'QUANT_BC_COFINS
            arrDados.Add "" 'ALIQ_COFINS_QUANT
            arrDados.Add "" 'VL_COFINS
            arrDados.Add Util.FormatarTexto(Campos(dicTitulosD201("COD_CTA")))
            arrDados.Add "" 'COD_NAT_PIS_COFINS
            arrDados.Add Empty 'INCONSISTENCIA
            arrDados.Add Empty 'SUGESTAO
            
        End If
        
        Campos = ValidarRegrasFiscaisPISCOFINS(arrDados.toArray(), COD_INC_TRIB)
        arrRelatorio.Add Campos
        arrDados.Clear
        
    Next Linha
    
End Function

Public Function CarregarDadosD205(ByRef arrRelatorio As ArrayList, ByRef dicTitulos As Dictionary, Optional ByVal Msg As String)
'TODO: REFATORAR ROTINA PARA FICAR IGUAL A CarregarDadosC185
Dim CHV_REG As String, ARQUIVO$, CST_PIS$, CST_COFINS$, CHV_CNPJ$, COD_INC_TRIB$, CHV_0110$, CHV_PAI$
Dim dicTitulos0110 As New Dictionary
Dim dicTitulosD010 As New Dictionary
Dim dicTitulosD200 As New Dictionary
Dim dicTitulosD205 As New Dictionary
Dim dicDados0110 As New Dictionary
Dim dicDadosD010 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicDadosD200 As New Dictionary
Dim arrDados As New ArrayList
Dim Campos As Variant, CamposD200
Dim Comeco As Double
Dim b As Long
Dim i As Byte
    
    Set dicTitulos0110 = Util.MapearTitulos(reg0110, 3)
    Set dicDados0110 = Util.CriarDicionarioRegistro(reg0110, "ARQUIVO")
        
    Set dicTitulosD010 = Util.MapearTitulos(regD010, 3)
    Set dicDadosD010 = Util.CriarDicionarioRegistro(regD010)
        
    Set dicTitulosD200 = Util.MapearTitulos(regD200, 3)
    Set dicDadosD200 = Util.CriarDicionarioRegistro(regD200)
    
    Set dicTitulosD205 = Util.MapearTitulos(regD205, 3)
    Set Dados = Util.DefinirIntervalo(regD205, 4, 3)
    
    If Dados Is Nothing Then Exit Function
    
    b = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(b, 10, Msg & "Carregando dados do registro D205", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            ARQUIVO = Campos(dicTitulosD205("ARQUIVO"))
            'COD_INC_TRIB = ExtrairRegimeTributario(dicDados0110, dicTitulos0110, ARQUIVO, COD_INC_TRIB)
            
            arrDados.Add Campos(dicTitulosD205("REG"))
            arrDados.Add ARQUIVO
            arrDados.Add Campos(dicTitulosD205("CHV_PAI_CONTRIBUICOES"))
            arrDados.Add Campos(dicTitulosD205("CHV_REG"))
            arrDados.Add COD_INC_TRIB
            
            CHV_PAI = Campos(dicTitulosD205("CHV_PAI_CONTRIBUICOES"))
            If dicDadosD200.Exists(CHV_PAI) Then
                
                CamposD200 = dicDadosD200(CHV_PAI)
                If LBound(CamposD200) = 0 Then i = 1 Else i = 0
                
                CHV_CNPJ = CamposD200(dicTitulosD200("CHV_PAI_CONTRIBUICOES") - i)
                If dicDadosD010.Exists(CHV_CNPJ) Then
                    
                    arrDados.Add "'" & dicDadosD010(CHV_CNPJ)(dicTitulosD010("CNPJ"))
                    
                Else
                    
                    arrDados.Add ""
                    
                End If
            
            Else
            
                arrDados.Add ""
            
            End If
            
            arrDados.Add "" 'CHV_NFE
            arrDados.Add "" 'NUM_DOC
            arrDados.Add "" 'SER
            arrDados.Add "" 'COD_ITEM
            arrDados.Add "O REGISTRO D205 POR PADRÃO NÃO IDENTIFICA OS ITENS" 'DESC_ITEM
            arrDados.Add "" 'COD_BARRA
            arrDados.Add "" 'COD_NCM
            arrDados.Add "" 'EX_IPI
            arrDados.Add "" 'TIPO_ITEM
            arrDados.Add "" 'IND_MOV
            arrDados.Add "" 'CFOP
            arrDados.Add Campos(dicTitulosD205("VL_ITEM"))
            arrDados.Add 0 'VL_DESP
            arrDados.Add 0 'VL_DESC
            arrDados.Add 0 'VL_ICMS
            arrDados.Add 0 'VL_ICMS_ST
            arrDados.Add "" 'CST_PIS
            arrDados.Add "" & Campos(dicTitulosD205("CST_COFINS"))
            arrDados.Add "" 'VL_BC_PIS
            arrDados.Add "" 'ALIQ_PIS
            arrDados.Add "" 'QUANT_BC_PIS
            arrDados.Add "" 'ALIQ_PIS_QUANT
            arrDados.Add "" 'VL_PIS
            arrDados.Add Campos(dicTitulosD205("VL_BC_COFINS"))
            arrDados.Add Campos(dicTitulosD205("ALIQ_COFINS"))
            arrDados.Add "" 'QUANT_BC_COFINS
            arrDados.Add "" 'ALIQ_COFINS_QUANT
            arrDados.Add Campos(dicTitulosD205("VL_COFINS"))
            arrDados.Add Util.FormatarTexto(Campos(dicTitulosD205("COD_CTA")))
            arrDados.Add "" 'COD_NAT_PIS_COFINS
            arrDados.Add Empty 'INCONSISTENCIA
            arrDados.Add Empty 'SUGESTAO
            
        End If
        
        Campos = ValidarRegrasFiscaisPISCOFINS(arrDados.toArray(), COD_INC_TRIB)
        arrRelatorio.Add Campos
        arrDados.Clear
        
    Next Linha
    
End Function

Public Sub CarregarDadosF100(ByVal Msg As String)

Dim REG As String, ARQUIVO$, CHV_REG$, CHV_0140$, CHV_0150$, CHV_F010$, COD_ITEM$, COD_PART$, COD_INC_TRIB$, CNPJ_ESTABELECIMENTO$, UF_CONTRIB$
Dim Apuracao As New clsAssistenteApuracao
Dim dicTitulosF100 As New Dictionary
Dim Dados As Range, Linha As Range
Dim Comeco As Double, VL_ITEM#
Dim Campos As Variant
Dim b As Long
    
    Set Dados = Util.DefinirIntervalo(regF100, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    Set Apuracao.dicTitulos = dicTitulos
    Set dicTitulosF100 = Util.MapearTitulos(regF100, 3)
    
    b = 0
    Comeco = Timer
    With Apuracao
        
        Call .CarregarDadosRegistro0000
        
        For Each Linha In Dados.Rows
            
            .RedimensionarArray (dicTitulos.Count)
            Call Util.AntiTravamento(b, 10, Msg & "Carregando dados do registro F100", Dados.Rows.Count, Comeco)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                'Carrega as variáveis necessárias
                REG = Campos(dicTitulosF100("REG"))
                ARQUIVO = Campos(dicTitulosF100("ARQUIVO"))
                COD_PART = Campos(dicTitulosF100("COD_PART"))
                CHV_0140 = .ExtrairCHV_0140(Util.UnirCampos(ARQUIVO, CNPJ_ESTABELECIMENTO))
                CHV_F010 = Campos(dicTitulosF100("CHV_PAI_CONTRIBUICOES"))
                CNPJ_ESTABELECIMENTO = .ExtrairCNPJ_ESTABELECIMENTO(REG, CHV_F010, ARQUIVO)
                UF_CONTRIB = .ExtrairUF_Estabelecimento(ARQUIVO, CNPJ_ESTABELECIMENTO)
                CHV_0150 = Util.UnirCampos(ARQUIVO, COD_PART)
                COD_ITEM = Campos(dicTitulosF100("COD_ITEM"))
                VL_ITEM = fnExcel.FormatarValores(Campos(dicTitulosF100("VL_OPER")))
                COD_INC_TRIB = .ExtrairREGIME_TRIBUTARIO(ARQUIVO, True)
                
                'Extrai dados do registro 0150
                Call .ExtrairDados0150(CHV_0150, True)
                
                If COD_ITEM = "" Then
                    
                    .AtribuirValor "DESCR_ITEM", "NENHUM ITEM INFORMADO PARA ESSE REGISTRO"
                    
                Else
                    
                    'Extrai dados do registro 0200
                    Call .ExtrairDados0200(CHV_0140, COD_ITEM, True)
                    
                End If
                
                'Atribui valores aos campos do relatório
                .AtribuirValor "CHV_NFE", "O REGISTRO F100 NÃO POSSUI NOTAS FISCAIS"
                .AtribuirValor "REG", Campos(dicTitulosF100("REG"))
                .AtribuirValor "ARQUIVO", ARQUIVO
                .AtribuirValor "CHV_PAI_CONTRIBUICOES", Campos(dicTitulosF100("CHV_PAI_CONTRIBUICOES"))
                .AtribuirValor "CHV_REG", Campos(dicTitulosF100("CHV_REG"))
                .AtribuirValor "CNPJ_ESTABELECIMENTO", fnExcel.FormatarTexto(CNPJ_ESTABELECIMENTO)
                .AtribuirValor "UF_CONTRIB", UF_CONTRIB
                .AtribuirValor "COD_ITEM", fnExcel.FormatarTexto(COD_ITEM)
                .AtribuirValor "VL_ITEM", VL_ITEM
                .AtribuirValor "CST_PIS", fnExcel.FormatarTexto(Campos(dicTitulosF100("CST_PIS")))
                .AtribuirValor "CST_COFINS", fnExcel.FormatarTexto(Campos(dicTitulosF100("CST_COFINS")))
                .AtribuirValor "VL_BC_PIS", fnExcel.ConverterValores(Campos(dicTitulosF100("VL_BC_PIS")))
                .AtribuirValor "ALIQ_PIS", fnExcel.ConverterValores(Campos(dicTitulosF100("ALIQ_PIS")))
                .AtribuirValor "VL_PIS", fnExcel.ConverterValores(Campos(dicTitulosF100("VL_PIS")))
                .AtribuirValor "VL_BC_COFINS", fnExcel.ConverterValores(Campos(dicTitulosF100("VL_BC_COFINS")))
                .AtribuirValor "ALIQ_COFINS", fnExcel.ConverterValores(Campos(dicTitulosF100("ALIQ_COFINS")))
                .AtribuirValor "VL_COFINS", fnExcel.ConverterValores(Campos(dicTitulosF100("VL_COFINS")))
                .AtribuirValor "COD_CTA", fnExcel.FormatarTexto(Campos(dicTitulosF100("COD_CTA")))
                
            End If
            
            Campos = ValidarRegrasFiscaisPISCOFINS(.Campo, COD_INC_TRIB)
            arrRelatorio.Add Campos
            
        Next Linha
        
    End With
    
End Sub

Public Function CarregarDadosF120(ByRef arrRelatorio As ArrayList, ByRef dicTitulos As Dictionary, Optional ByVal Msg As String)
'TODO: REFATORAR ROTINA PARA FICAR IGUAL A CarregarDadosF100
Dim CHV_REG As String, ARQUIVO$, CST_PIS$, CST_COFINS$, CHV_CNPJ$, COD_INC_TRIB$, CHV_0110$
Dim dicTitulos0110 As New Dictionary
Dim dicTitulosF120 As New Dictionary
Dim dicTitulosF010 As New Dictionary
Dim dicDados0110 As New Dictionary
Dim dicDadosF010 As New Dictionary
Dim Dados As Range, Linha As Range
Dim dicDadosA100 As New Dictionary
Dim arrDados As New ArrayList
Dim Campos As Variant
Dim Comeco As Double
Dim b As Long
    
    Set dicTitulos0110 = Util.MapearTitulos(reg0110, 3)
    Set dicDados0110 = Util.CriarDicionarioRegistro(reg0110, "ARQUIVO")
    
    Set dicTitulosF010 = Util.MapearTitulos(regF010, 3)
    Set dicDadosF010 = Util.CriarDicionarioRegistro(regF010)
    
    Set dicTitulosF120 = Util.MapearTitulos(regF120, 3)
    Set Dados = Util.DefinirIntervalo(regF120, 4, 3)
    
    If Dados Is Nothing Then Exit Function
    
    b = 0
    Comeco = Timer
    For Each Linha In Dados.Rows
        
        Call Util.AntiTravamento(b, 10, Msg & "Carregando dados do registro F120", Dados.Rows.Count, Comeco)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            ARQUIVO = Campos(dicTitulosF120("ARQUIVO"))
            'COD_INC_TRIB = ExtrairRegimeTributario(dicDados0110, dicTitulos0110, ARQUIVO, COD_INC_TRIB)
            
            arrDados.Add Campos(dicTitulosF120("REG"))
            arrDados.Add ARQUIVO
            arrDados.Add Campos(dicTitulosF120("CHV_PAI_CONTRIBUICOES"))
            arrDados.Add Campos(dicTitulosF120("CHV_REG"))
            arrDados.Add COD_INC_TRIB
            
            CHV_CNPJ = Campos(dicTitulosF120("CHV_PAI_CONTRIBUICOES"))
            If dicDadosF010.Exists(CHV_CNPJ) Then
                
                arrDados.Add "'" & dicDadosF010(CHV_CNPJ)(dicTitulosF010("CNPJ"))
                
            Else
                
                arrDados.Add ""
                
            End If
            
            arrDados.Add "" 'CHV_NFE
            arrDados.Add "" 'NUM_DOC
            arrDados.Add "" 'SER
            arrDados.Add "" 'COD_ITEM
            arrDados.Add "" 'DESC_ITEM
            arrDados.Add "" 'COD_BARRA
            arrDados.Add "" 'COD_NCM
            arrDados.Add "" 'EX_IPI
            arrDados.Add "" 'TIPO_ITEM
            arrDados.Add "" 'IND_MOV
            arrDados.Add "" 'CFOP
            arrDados.Add Campos(dicTitulosF120("VL_OPER_DEP"))
            arrDados.Add 0 'VL_DESP
            arrDados.Add Campos(dicTitulosF120("PARC_OPER_NAO_BC_CRED"))
            arrDados.Add 0 'VL_ICMS
            arrDados.Add 0 'VL_ICMS_ST
            arrDados.Add "'" & Campos(dicTitulosF120("CST_PIS"))
            arrDados.Add "'" & Campos(dicTitulosF120("CST_COFINS"))
            arrDados.Add Campos(dicTitulosF120("VL_BC_PIS"))
            arrDados.Add Campos(dicTitulosF120("ALIQ_PIS"))
            arrDados.Add "" 'QUANT_BC_PIS
            arrDados.Add "" 'ALIQ_PIS_QUANT
            arrDados.Add Campos(dicTitulosF120("VL_PIS"))
            arrDados.Add Campos(dicTitulosF120("VL_BC_COFINS"))
            arrDados.Add Campos(dicTitulosF120("ALIQ_COFINS"))
            arrDados.Add "" 'QUANT_BC_COFINS
            arrDados.Add "" 'ALIQ_COFINS_QUANT
            arrDados.Add Campos(dicTitulosF120("VL_COFINS"))
            arrDados.Add Util.FormatarTexto(Campos(dicTitulosF120("COD_CTA")))
            arrDados.Add "" 'COD_NAT_PIS_COFINS
            arrDados.Add Empty 'INCONSISTENCIA
            arrDados.Add Empty 'SUGESTAO
            
        End If
        
        Campos = ValidarRegrasFiscaisPISCOFINS(arrDados.toArray(), COD_INC_TRIB)
        arrRelatorio.Add Campos
        arrDados.Clear
        
    Next Linha
    
End Function

Public Function ValidarRegrasFiscaisPISCOFINS(ByVal Campos As Variant, ByVal COD_INC_TRIB As String) As Variant

Dim REG As String
Dim i As Integer
    
    If UBound(Campos) = -1 Then
        ValidarRegrasFiscaisPISCOFINS = Campos
        Exit Function
    End If
    
    If dicTitulos.Count = 0 Then Set dicTitulos = Util.MapearTitulos(assApuracaoPISCOFINS, 3)
    If LBound(Campos) = 0 Then i = 1
    
    Call ValidacoesGerais.ExecutarValidacoesGerais(Campos, dicTitulos)
    
    REG = Campos(dicTitulos("REG") - i)
    
    Select Case REG
        
        Case "A170"
            Call ValidarRegrasFiscaisA170(Campos, dicTitulos, COD_INC_TRIB)
            
        Case "C170"
            Call ValidarRegrasFiscaisC170(Campos, dicTitulos, COD_INC_TRIB)
            
        Case "C175"
            Call ValidarRegrasFiscaisC175(Campos, dicTitulos, COD_INC_TRIB)
            
        Case "C181"
            Call ValidarRegrasFiscaisC181(Campos, dicTitulos, COD_INC_TRIB)
            
        Case "C185"
            Call ValidarRegrasFiscaisC185(Campos, dicTitulos, COD_INC_TRIB)
            
        Case "C191"
            Call ValidarRegrasFiscaisC191(Campos, dicTitulos, COD_INC_TRIB)
            
        Case "C195"
            Call ValidarRegrasFiscaisC195(Campos, dicTitulos, COD_INC_TRIB)
            
        Case "D201"
            Call ValidarRegrasFiscaisD201(Campos, dicTitulos, COD_INC_TRIB)
            
        Case "D205"
            Call ValidarRegrasFiscaisD205(Campos, dicTitulos, COD_INC_TRIB)
            
        Case "F100"
            Call ValidarRegrasFiscaisF100(Campos, dicTitulos, COD_INC_TRIB)
            
        Case "F120"
            Call ValidarRegrasFiscaisF120(Campos, dicTitulos, COD_INC_TRIB)
            
    End Select
    
Sair:
    ValidarRegrasFiscaisPISCOFINS = Campos
    
End Function

Public Sub ReprocessarSugestoes()

Dim arrRelatorio As New ArrayList
Dim dicTitulos As New Dictionary
Dim COD_INC_TRIB As String
Dim arrDados As ArrayList
Dim Campos As Variant
    
    Call Util.DesabilitarControles
    
    If Util.ChecarAusenciaDados(assApuracaoPISCOFINS) Then GoTo Finalizar:
    If assApuracaoPISCOFINS.AutoFilterMode Then assApuracaoPISCOFINS.AutoFilter.ShowAllData
    
    Set arrDados = Util.CriarArrayListRegistro(assApuracaoPISCOFINS)
    If arrDados.Count = 0 Then GoTo Finalizar:
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoPISCOFINS, 3)
    
    Call DadosValidacaoCFOP.CarregarTitulosRelatorio(ActiveSheet)
    
    a = 0
    Comeco = Timer
    For Each Campos In arrDados
        
        Call Util.AntiTravamento(a, 50, "Reprocessando inconsistências, por favor aguarde...", arrDados.Count, Comeco)
        
        If Util.ChecarCamposPreenchidos(Campos) Then
            
            COD_INC_TRIB = Util.ApenasNumeros(Campos(dicTitulos("REGIME_TRIBUTARIO")))
            
            Campos(dicTitulos("INCONSISTENCIA")) = Empty
            Campos(dicTitulos("SUGESTAO")) = Empty
            Campos = ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
            arrRelatorio.Add Campos
            
        End If
        
    Next Campos
    
    Call Util.AtualizarBarraStatus("Atualizando relatório de inconsistências, isso pode levar alguns segundos! Por favor aguarde...")
    
    Call Util.LimparDados(assApuracaoPISCOFINS, 4, False)
    Call Util.ExportarDadosArrayList(assApuracaoPISCOFINS, arrRelatorio)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoPISCOFINS)
    
    Call Util.AtualizarBarraStatus("Processamento Concluído! Já pode mexer nos dados novamente.")
    
Finalizar:
    Call Util.HabilitarControles

End Sub

Public Function AceitarSugestoes()

Dim Campos As Variant, Campos0200, CamposC170, dicCampos, regCampo, CamposTrib
Dim ARQUIVO As String, COD_INC_TRIB$, CNPJ$, COD_ITEM$, CFOP$, CHV_REG$
Dim VL_BC_PIS_CALC As Double, VL_BC_COFINS_CALC#
Dim Apuracao As New clsAssistenteApuracao
Dim dicDados0110 As New Dictionary
Dim Dados As Range, Linha As Range
Dim arrDados As New ArrayList
    
    Inicio = Now()
    
    Call Util.DesabilitarControles
    Call Util.AtualizarBarraStatus("Implementando sugestões selecionadas, por favor aguarde...")
    
    With Apuracao
        
        Set .dicTitulos = Util.MapearTitulos(assApuracaoPISCOFINS, 3)
        Set Dados = assApuracaoPISCOFINS.Range("A4").CurrentRegion
        
        If Dados Is Nothing Then
            Call Util.MsgAlerta("Não foi encontrada nenhuma inconsistência!", "Inconsistências Fiscais")
            GoTo Finalizar:
        End If
        
        a = 0
        For Each Linha In Dados.Rows
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                If Linha.EntireRow.Hidden = False And Campos(.dicTitulos("SUGESTAO")) <> "" And Linha.Row > 3 Then
                    
                    ARQUIVO = Campos(.dicTitulos("ARQUIVO"))
                    COD_INC_TRIB = Util.ApenasNumeros(Campos(.dicTitulos("REGIME_TRIBUTARIO")))
                    
VERIFICAR:
                    Call Util.AntiTravamento(a, 10, "Aplicando sugestões sugeridas...", Dados.Rows.Count, Comeco)
                    Select Case Campos(.dicTitulos("SUGESTAO"))
                        
                        Case "Alterar valor do campo COD_SIT para: 08 - Regime Especial ou Norma Específica"
                            Campos(dicTitulos("COD_SIT")) = "08 - Regime Especial ou Norma Específica"
                            Campos(dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                        
                        Case "Recalcular bases de PIS e COFINS"
                            Campos(.dicTitulos("VL_BC_PIS")) = RegrasFiscais.ApuracaoPISCOFINS.CalcularBasePISCOFINS(.dicTitulos, Campos)
                            Campos(.dicTitulos("VL_BC_COFINS")) = RegrasFiscais.ApuracaoPISCOFINS.CalcularBasePISCOFINS(.dicTitulos, Campos)
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Gerar base de cálculo do PIS"
                            VL_BC_PIS_CALC = RegrasFiscais.ApuracaoPISCOFINS.CalcularBasePISCOFINS(.dicTitulos, Campos)
                            If VL_BC_PIS_CALC > 0 Then Campos(.dicTitulos("VL_BC_PIS")) = VL_BC_PIS_CALC Else Campos(.dicTitulos("VL_BC_PIS")) = 0
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Gerar base de cálculo da COFINS"
                            VL_BC_COFINS_CALC = RegrasFiscais.ApuracaoPISCOFINS.CalcularBasePISCOFINS(.dicTitulos, Campos)
                            If VL_BC_COFINS_CALC > 0 Then Campos(.dicTitulos("VL_BC_COFINS")) = VL_BC_COFINS_CALC Else Campos(.dicTitulos("VL_BC_COFINS")) = 0
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Recalcular valor do PIS"
                            Campos(.dicTitulos("VL_PIS")) = VBA.Round(Campos(.dicTitulos("VL_BC_PIS")) * fnExcel.FormatarPercentuais(Campos(.dicTitulos("ALIQ_PIS"))), 2)
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Recalcular valor da COFINS"
                            Campos(.dicTitulos("VL_COFINS")) = VBA.Round(Campos(.dicTitulos("VL_BC_COFINS")) * fnExcel.FormatarPercentuais(Campos(.dicTitulos("ALIQ_COFINS"))), 2)
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Zerar base, alíquota e valor do PIS", "Zerar base, alíquota e valor da COFINS", "Zerar valores do PIS", "Zerar valores da COFINS"
                            Campos(.dicTitulos("VL_BC_PIS")) = 0
                            Campos(.dicTitulos("ALIQ_PIS")) = 0
                            Campos(.dicTitulos("VL_PIS")) = 0
                            Campos(.dicTitulos("VL_BC_COFINS")) = 0
                            Campos(.dicTitulos("ALIQ_COFINS")) = 0
                            Campos(.dicTitulos("VL_COFINS")) = 0
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Zerar base e valores de PIS e COFINS"
                            Campos(.dicTitulos("VL_BC_PIS")) = 0
                            Campos(.dicTitulos("VL_PIS")) = 0
                            Campos(.dicTitulos("VL_BC_COFINS")) = 0
                            Campos(.dicTitulos("VL_COFINS")) = 0
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Informar alíquota de 1,65% para o PIS"
                            Campos(.dicTitulos("ALIQ_PIS")) = 0.0165
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Informar alíquota de 0,65% para o PIS"
                            Campos(.dicTitulos("ALIQ_PIS")) = 0.0065
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Informar alíquota de 7,60% para a COFINS"
                            Campos(.dicTitulos("ALIQ_COFINS")) = 0.076
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Informar alíquota de 3,00% para a COFINS"
                            Campos(.dicTitulos("ALIQ_COFINS")) = 0.03
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Alterar CST_PIS para 49", "Informar CST_PIS 49 - Outras Operações de Saída"
                            Campos(.dicTitulos("CST_PIS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(49)
                            Campos(.dicTitulos("CST_COFINS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(49)
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                                                                        
                        Case "Informar CST_PIS igual a 70 - Operação de Aquisição sem Direito a Crédito"
                            Campos(.dicTitulos("CST_PIS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(70)
                            Campos(.dicTitulos("CST_COFINS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(70)
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Informar CST_COFINS igual a 70 - Operação de Aquisição sem Direito a Crédito"
                            Campos(.dicTitulos("CST_PIS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(70)
                            Campos(.dicTitulos("CST_COFINS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(70)
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Informar CST_PIS 98 - Outras Operações de Entrada"
                            Campos(.dicTitulos("CST_PIS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(98)
                            Campos(.dicTitulos("CST_COFINS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(98)
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Alterar CST_COFINS para 49", "Informar CST_COFINS 49 - Outras Operações de Saída"
                            Campos(.dicTitulos("CST_PIS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(49)
                            Campos(.dicTitulos("CST_COFINS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(49)
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
    
                        Case "Informar CST_COFINS 98 - Outras Operações de Entrada"
                            Campos(.dicTitulos("CST_PIS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(98)
                            Campos(.dicTitulos("CST_COFINS")) = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_CST_PIS_COFINS(98)
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Alterar o valor do campo TIPO_ITEM para 00"
                            Campos(.dicTitulos("TIPO_ITEM")) = "00 - Mercadoria para Revenda"
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Alterar o valor do campo TIPO_ITEM para 07"
                            Campos(.dicTitulos("TIPO_ITEM")) = "07 - Material de Uso e Consumo"
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Alterar o valor do campo TIPO_ITEM para 08"
                            Campos(.dicTitulos("TIPO_ITEM")) = "08 - Ativo Imobilizado"
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            Campos = Assistente.Fiscal.Apuracao.PISCOFINS.ValidarRegrasFiscaisPISCOFINS(Campos, COD_INC_TRIB)
                            GoTo VERIFICAR:
                            
                        Case "Alterar campo REGIME_TRIBUTARIO para 2 - Cumulativo"
                            Campos(.dicTitulos("REGIME_TRIBUTARIO")) = "2 - Cumulativo"
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                            
                        Case "Alterar campo REGIME_TRIBUTARIO para 1 - Não-Cumulativo"
                            Campos(.dicTitulos("REGIME_TRIBUTARIO")) = "1 - Não-Cumulativo"
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                        
                        Case "Apagar valor informado no campo COD_NCM"
                            Campos(.dicTitulos("COD_NCM")) = ""
                            Campos(.dicTitulos("INCONSISTENCIA")) = Empty
                            Campos(.dicTitulos("SUGESTAO")) = Empty
                        
                    End Select
                    
                End If
                
                If Linha.Row > 3 Then arrDados.Add Campos
            
            Else
            
                Stop
            
            End If
            
        Next Linha
    
    End With
    
    If assApuracaoPISCOFINS.AutoFilterMode Then assApuracaoPISCOFINS.AutoFilter.ShowAllData
    Call Util.LimparDados(assApuracaoPISCOFINS, 4, False)
    Call Util.ExportarDadosArrayList(assApuracaoPISCOFINS, arrDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoPISCOFINS)
    
    Call Util.MsgInformativa("Sugestões aplicadas com sucesso!", "Inclusão de Sugestões", Inicio)
    Application.StatusBar = False
    
Finalizar:
    Call Util.HabilitarControles
    
End Function

Public Sub AtualizarRegistrosPIS_COFINS()

Dim Campos As Variant, Campos0200, CamposA170, CamposC170, CamposC175, CamposC181, Camposc185, CamposD201, CamposD205
Dim ARQUIVO As String, CHV_REG$, REG$, CHV_0000$, CHV_0140$, CHV_M001$, CNPJ$, COD_ITEM$, Msg$
Dim dicCorrelacoes As New Dictionary
Dim dicReceitaCST As New Dictionary
Dim dicReceitaCSTNAT As New Dictionary
Dim dicTitulos As New Dictionary
Dim arrDados As New ArrayList
Dim Status As Boolean
Dim i As Byte
    
    Inicio = Now()
    Call Util.DesabilitarControles
    
    Campos0200 = Array("COD_BARRA", "COD_NCM", "EX_IPI", "TIPO_ITEM")
    CamposA170 = Array("VL_ITEM", "VL_DESC", "CST_PIS", "VL_BC_PIS", "ALIQ_PIS", "VL_PIS", "CST_COFINS", "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS", "COD_CTA")
    CamposC170 = Array("IND_MOV", "CFOP", "VL_ITEM", "VL_DESC", "VL_ICMS", "CST_PIS", "VL_BC_PIS", "ALIQ_PIS", "QUANT_BC_PIS", "ALIQ_PIS_QUANT", "VL_PIS", "CST_COFINS", "VL_BC_COFINS", "ALIQ_COFINS", "QUANT_BC_COFINS", "ALIQ_COFINS_QUANT", "VL_COFINS", "COD_CTA")
    CamposC175 = Array("CFOP", "VL_ITEM", "VL_DESC", "CST_PIS", "VL_BC_PIS", "ALIQ_PIS", "QUANT_BC_PIS", "ALIQ_PIS_QUANT", "VL_PIS", "CST_COFINS", "VL_BC_COFINS", "ALIQ_COFINS", "QUANT_BC_COFINS", "ALIQ_COFINS_QUANT", "VL_COFINS", "COD_CTA")
    CamposC181 = Array("CST_PIS", "VL_ITEM", "VL_DESC", "VL_BC_PIS", "ALIQ_PIS", "QUANT_BC_PIS", "ALIQ_PIS_QUANT", "VL_PIS", "COD_CTA")
    Camposc185 = Array("CST_COFINS", "VL_ITEM", "VL_BC_COFINS", "ALIQ_COFINS", "QUANT_BC_COFINS", "ALIQ_COFINS_QUANT", "VL_COFINS", "COD_CTA")
    CamposD201 = Array("CST_PIS", "VL_ITEM", "VL_BC_PIS", "ALIQ_PIS", "VL_PIS", "COD_CTA")
    CamposD205 = Array("CST_COFINS", "VL_ITEM", "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS", "COD_CTA")
    
    Call InicializarObjetos_Atualizacao

    Set dicTitulos = Util.MapearTitulos(assApuracaoPISCOFINS, 3)
    If assApuracaoPISCOFINS.AutoFilterMode Then assApuracaoPISCOFINS.AutoFilter.ShowAllData
    
    Set arrDados = Util.CriarArrayListRegistro(assApuracaoPISCOFINS)
    If arrDados.Count = 0 Then GoTo Finalizar:
    
    With dtoRegSPED
        
        a = 0
        Comeco = Timer
        For Each Campos In arrDados
            
            Call Util.AntiTravamento(a, 50, "Processando as atualizações dos registros, por favor aguarde...", arrDados.Count, Comeco)
            If LBound(Campos) = 0 Then i = 1 Else i = 0
            
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                REG = Campos(dicTitulos("REG") - i)
                
                If REG Like "*70" Then
                    
                    ARQUIVO = Campos(dicTitulos("ARQUIVO"))
                    CNPJ = Campos(dicTitulos("CNPJ_ESTABELECIMENTO"))
                    CHV_0140 = .r0140(Util.UnirCampos(ARQUIVO, CNPJ))(dtoTitSPED.t0140("CHV_REG"))
                    COD_ITEM = Campos(dicTitulos("COD_ITEM"))
                    CHV_REG = Util.UnirCampos(CHV_0140, COD_ITEM)
                    Call AtualizarDadosRegistro(.r0200, dtoTitSPED.t0200, Campos0200, dicTitulos, Campos, CHV_REG)
                    
                End If
                
                Select Case REG
                    
                    Case "A170"
                        CHV_REG = Util.RemoverAspaSimples(Campos(dicTitulos("CHV_REG") - i))
                        Call AtualizarDadosRegistro(.rA170, dtoTitSPED.tA170, CamposA170, dicTitulos, Campos, CHV_REG)
                        
                    Case "C170"
                        CHV_REG = Util.RemoverAspaSimples(Campos(dicTitulos("CHV_REG") - i))
                        Call AtualizarDadosRegistro(.rC170, dtoTitSPED.tC170, CamposC170, dicTitulos, Campos, CHV_REG)
                        
                    Case "C175"
                        CHV_REG = Util.RemoverAspaSimples(Campos(dicTitulos("CHV_REG") - i))
                        
                        'Correlaciona campos de origem (relatório) e destino (registro)
                        dicCorrelacoes("VL_ITEM") = "VL_OPER"
                        
                        Call AtualizarDadosRegistro(.rC175_Contr, dtoTitSPED.tC175_Contr, CamposC175, dicTitulos, Campos, CHV_REG, dicCorrelacoes)
                        
                    Case "C181"
                        CHV_REG = Util.RemoverAspaSimples(Campos(dicTitulos("CHV_REG") - i))
                        Call AtualizarDadosRegistro(.rC181_Contr, dtoTitSPED.tC181_Contr, CamposC181, dicTitulos, Campos, CHV_REG)
                        
                    Case "C185"
                        CHV_REG = Util.RemoverAspaSimples(Campos(dicTitulos("CHV_REG") - i))
                        Call AtualizarDadosRegistro(.rC185_Contr, dtoTitSPED.tC185_Contr, Camposc185, dicTitulos, Campos, CHV_REG)
                        
                    Case "D201"
                        CHV_REG = Util.RemoverAspaSimples(Campos(dicTitulos("CHV_REG") - i))
                        Call AtualizarDadosRegistro(.rD201, dtoTitSPED.tD201, CamposD201, dicTitulos, Campos, CHV_REG)
                        
                    Case "D205"
                        CHV_REG = Util.RemoverAspaSimples(Campos(dicTitulos("CHV_REG") - i))
                        Call AtualizarDadosRegistro(.rD205, dtoTitSPED.tD205, CamposD205, dicTitulos, Campos, CHV_REG)
                        
                End Select
                
                ARQUIVO = Util.RemoverAspaSimples(Util.LimparTexto(CStr(Campos(dicTitulos("ARQUIVO") - i))))
                
            End If
            
            If AtualizarM400M800 Then
                
                If .rM001.Exists(ARQUIVO) Then
                    
                    If LBound(.rM001(ARQUIVO)) = 0 Then i = 1 Else i = 0
                    CHV_M001 = .rM001(ARQUIVO)(dtoTitSPED.tM001("CHV_REG") - i)
                    
                Else
                    
                    'Carrega campo CHV_REG do registro 0000
                    If LBound(.r0000_Contr(ARQUIVO)) = 0 Then i = 1 Else i = 0
                    CHV_0000 = Util.LimparTexto(CStr(.r0000_Contr(ARQUIVO)(dtoTitSPED.t0000_Contr("CHV_REG") - i)))
                    
                    'Cria chave do registro M001
                    CHV_M001 = fnSPED.GerarChaveRegistro(CHV_0000, "M001")
                    
                    .rM001(ARQUIVO) = Array("M001", ARQUIVO, CHV_M001, "", CHV_0000, "0")
                    
                End If
                
                Call AtualizarReceitasBlocoM(Campos, dicTitulos, CHV_M001)
                
                Call dicCorrelacoes.RemoveAll
                
            End If
            
        Next Campos
        
    End With
    
    Call ExpReg.ExportarRegistros("0200", "C100", "C170", "C175_Contr", "C181_Contr", "C185_Contr", "D201", "D205")
    
    If AtualizarM400M800 Then Call ExpReg.ExportarRegistros("M001", "M400", "M410", "M800", "M810")
    
    Call Util.AtualizarBarraStatus("Atualizando código do gênero no registro 0200, por favor aguarde...")
    Call r0200.AtualizarCodigoGenero(True)
    
    Call Util.AtualizarBarraStatus("Gerando dados do registro C175, por favor aguarde...")
    Call rC170.GerarC175(True)
    Call rC175Contr.AgruparRegistros(True)
    
    Call Util.AtualizarBarraStatus("Atualizando valores dos impostos no registro C100, por favor aguarde...")
    Call rC170.AtualizarImpostosC100(True)
    Call rC175Contr.AtualizarImpostosC100(True)
    
    Call Util.AtualizarBarraStatus("Atualização concluída com sucesso!")
    Call Util.MsgInformativa("Registros atualizados com sucesso!", "Atualização de dados", Inicio)
    Call Util.AtualizarBarraStatus(False)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoPISCOFINS)
    
Finalizar:
    Call LimparObjetos_Atualizacao
    
End Sub

Public Function AtualizarReceitasBlocoM(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, CHV_M001 As String)

Dim CST_PIS As String, CST_COFINS$, COD_CTA$, COD_NAT$, CHV_0000$, CHV_M400$, CHV_M800$, CHV_PAI$, CHV_REG$, Chave$
Dim VL_ITEM As Double, VL_ITEM00#, VL_ITEM10#, VL_PIS#, VL_COFINS#
Dim dicCampos As Variant
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    ARQUIVO = Util.RemoverAspaSimples(Campos(dicTitulos("ARQUIVO")))
    CST_PIS = Util.ApenasNumeros(Campos(dicTitulos("CST_PIS")))
    CST_COFINS = Util.ApenasNumeros(Campos(dicTitulos("CST_COFINS")))
    VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM")))
    VL_PIS = fnExcel.ConverterValores(Campos(dicTitulos("VL_PIS")))
    VL_COFINS = fnExcel.ConverterValores(Campos(dicTitulos("VL_COFINS")))
    COD_CTA = Util.RemoverAspaSimples(Campos(dicTitulos("COD_CTA")))
    COD_NAT = Util.RemoverAspaSimples(Campos(dicTitulos("COD_NAT_PIS_COFINS")))
    ARQUIVO = Campos(dicTitulos("ARQUIVO"))
    
    With dtoRegSPED
        
        If CST_PIS > "03" And CST_PIS < "10" And VL_PIS = 0 Then
            
            CHV_M400 = fnSPED.GerarChaveRegistro(CHV_M001, CST_PIS, COD_CTA)
            If .rM400.Exists(CHV_M400) Then
                
                dicCampos = .rM400(CHV_M400)
                i = Util.VerificarPosicaoInicialArray(dicCampos)
                VL_ITEM00 = dicCampos(dtoTitSPED.tM400("VL_TOT_REC") - i)
                
            End If
            
            .rM400(CHV_M400) = Array("M400", ARQUIVO, CHV_M400, "", CHV_M001, "'" & CST_PIS, VL_ITEM + VL_ITEM00, COD_CTA, "")
                        
            CHV_REG = fnSPED.GerarChaveRegistro(CHV_M400, COD_NAT, COD_CTA)
            If .rM410.Exists(CHV_REG) Then
                
                dicCampos = .rM410(CHV_REG)
                i = Util.VerificarPosicaoInicialArray(dicCampos)
                VL_ITEM10 = dicCampos(dtoTitSPED.tM410("VL_REC") - i)
                
            End If
            
            CHV_REG = fnSPED.GerarChaveRegistro(CHV_M400, COD_NAT, COD_CTA)
            .rM410(CHV_REG) = Array("M410", ARQUIVO, CHV_REG, "", CHV_M400, "'" & COD_NAT, VL_ITEM + VL_ITEM10, COD_CTA, "")
            
        End If
        
        If CST_COFINS > "03" And CST_COFINS < "10" And VL_COFINS = 0 Then
            
            CHV_M800 = fnSPED.GerarChaveRegistro(CHV_M001, CST_COFINS, COD_CTA)
            If .rM800.Exists(CHV_M800) Then
                
                dicCampos = .rM800(CHV_M800)
                If LBound(dicCampos) = 0 Then i = 1 Else i = 0
                VL_ITEM00 = dicCampos(dtoTitSPED.tM800("VL_TOT_REC") - i)
                
            End If
            
            .rM800(CHV_M800) = Array("M800", ARQUIVO, CHV_M800, "", CHV_M001, "'" & CST_COFINS, VL_ITEM + VL_ITEM00, COD_CTA, "")
            
            CHV_REG = fnSPED.GerarChaveRegistro(CHV_M800, COD_NAT, COD_CTA)
            If .rM810.Exists(CHV_REG) Then
                
                dicCampos = .rM810(CHV_REG)
                If LBound(dicCampos) = 0 Then i = 1 Else i = 0
                VL_ITEM10 = dicCampos(dtoTitSPED.tM810("VL_REC") - i)
                
            End If
            
            .rM810(CHV_REG) = Array("M810", ARQUIVO, CHV_REG, "", CHV_M800, "'" & COD_NAT, VL_ITEM + VL_ITEM10, COD_CTA, "")
            
        End If
        
    End With
    
End Function

'Public Sub AtualizarRegistroM400(ByRef Dados As Range, ByRef dicTitulos As Dictionary, ByRef CamposM400 As Variant, Optional ByVal Msg As String)
'
'Dim Campos As Variant, Campo, dicCampos, regCampo
'Dim dicTitulosM400 As New Dictionary
'Dim dicDadosM400 As New Dictionary
'Dim CHV_M400 As String, REG$
'Dim Linha As Range
'Dim b As Long, i&
'
'    Set dicTitulosM400 = Util.MapearTitulos(regM400, 3)
'    Set dicDadosM400 = Util.CriarDicionarioRegistro(regM400)
'
'    If dicDadosM400.Count = 0 Then Exit Sub
'
'    b = 0
'    Comeco = Timer
'    For Each Linha In Dados.Rows
'
'        Call Util.AntiTravamento(b, 10, Msg & "Atualizando registro " & b & " de " & Dados.Rows.Count, Dados.Rows.Count, Comeco)
'
'        Campos = Application.index(Linha.Value2, 0, 0)
'        If Util.ChecarCamposPreenchidos(Campos) Then
'
'            REG = Campos(dicTitulos("REG"))
'            If REG <> "M400" Then GoTo Prx:
'
'            CHV_M400 = Campos(dicTitulos("CHV_REG"))
'
'            'Atualizar dados do M400
'            If dicDadosM400.Exists(CHV_M400) Then
'
'                If LBound(Campos) = 0 Then i = 1 Else i = 0
'
'                dicCampos = dicDadosM400(CHV_M400)
'                For Each regCampo In CamposM400
'
'                    Campo = regCampo
'                    dicCampos(dicTitulosM400(regCampo)) = Campos(dicTitulos(Campo))
'
'                Next regCampo
'
'                dicDadosM400(CHV_M400) = dicCampos
'
'            End If
'
'        End If
'Prx:
'    Next Linha
'
'    Application.StatusBar = Msg & "Salvando alterações..."
'    Call Util.LimparDados(regM400, 4, False)
'    Call Util.ExportarDadosDicionario(regM400, dicDadosM400)
'
'End Sub

Public Function ValidarRegrasFiscaisA170(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal COD_INC_TRIB As String)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_REGIME_TRIBUTARIO(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidacoesCampo_CST_PIS_COFINS(Campos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_VL_DESC(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_ALIQ_PIS_COFINS(Campos, dicTitulos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_BC_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_COD_CTA(Campos, dicTitulos)
    
    
End Function

Public Function ValidarRegrasFiscaisC170(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal COD_INC_TRIB As String)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_REGIME_TRIBUTARIO(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidacoesCFOP.ValidarCampo_CFOP(Campos, "PISCOFINS")
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidacoesCampo_CST_PIS_COFINS(Campos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_BC_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_VL_DESC(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_ALIQ_PIS_COFINS(Campos, dicTitulos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_DT_ENT_SAI(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidacoesNCM.ValidarCampo_COD_NCM(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_TIPO_ITEM(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_COD_NAT_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_COD_CTA(Campos, dicTitulos)
    
End Function

Public Function ValidarRegrasFiscaisC175(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal COD_INC_TRIB As String)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call ValidacoesCFOP.ValidarCampo_CFOP(Campos, "PISCOFINS")
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_VL_DESC(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidacoesCampo_CST_PIS_COFINS(Campos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_ALIQ_PIS_COFINS(Campos, dicTitulos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_BC_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_COD_CTA(Campos, dicTitulos)
    
End Function

Public Function ValidarRegrasFiscaisC181(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal COD_INC_TRIB As String)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_VL_DESC(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidacoesCampo_CST_PIS_COFINS(Campos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_ALIQ_PIS_COFINS(Campos, dicTitulos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_BC_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_COD_CTA(Campos, dicTitulos)
    
End Function

Public Function ValidarRegrasFiscaisC185(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal COD_INC_TRIB As String)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_VL_DESC(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidacoesCampo_CST_PIS_COFINS(Campos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_ALIQ_PIS_COFINS(Campos, dicTitulos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_BC_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_COD_CTA(Campos, dicTitulos)
    
End Function

Public Function ValidarRegrasFiscaisC191(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal COD_INC_TRIB As String)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_VL_DESC(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidacoesCampo_CST_PIS_COFINS(Campos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_ALIQ_PIS_COFINS(Campos, dicTitulos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_BC_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_COD_CTA(Campos, dicTitulos)
    
End Function

Public Function ValidarRegrasFiscaisC195(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal COD_INC_TRIB As String)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_VL_DESC(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidacoesCampo_CST_PIS_COFINS(Campos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_ALIQ_PIS_COFINS(Campos, dicTitulos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_BC_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_COD_CTA(Campos, dicTitulos)
    
End Function

Public Function ValidarRegrasFiscaisD201(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal COD_INC_TRIB As String)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_VL_DESC(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidacoesCampo_CST_PIS_COFINS(Campos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_ALIQ_PIS_COFINS(Campos, dicTitulos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_BC_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_COD_CTA(Campos, dicTitulos)
    
End Function

Public Function ValidarRegrasFiscaisD205(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal COD_INC_TRIB As String)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_VL_DESC(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidacoesCampo_CST_PIS_COFINS(Campos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_ALIQ_PIS_COFINS(Campos, dicTitulos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_BC_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_COD_CTA(Campos, dicTitulos)
    
End Function

Public Function ValidarRegrasFiscaisF100(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal COD_INC_TRIB As String)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_VL_DESC(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidacoesCampo_CST_PIS_COFINS(Campos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_ALIQ_PIS_COFINS(Campos, dicTitulos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_BC_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_COD_CTA(Campos, dicTitulos)
    
End Function

Public Function ValidarRegrasFiscaisF120(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal COD_INC_TRIB As String)

Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_VL_DESC(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidacoesCampo_CST_PIS_COFINS(Campos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_ALIQ_PIS_COFINS(Campos, dicTitulos, COD_INC_TRIB)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_BC_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.ApuracaoPISCOFINS.ValidarCampo_VL_PIS_COFINS(Campos, dicTitulos)
    If Campos(dicTitulos("INCONSISTENCIA") - i) = "" Then Call RegrasFiscais.Geral.ValidarCampo_COD_CTA(Campos, dicTitulos)
    
End Function

Public Function ModificarRegimeTributarioPISCOFINS()

Dim dicDados As New Dictionary
Dim dicTitulos As New Dictionary
Dim dicDados0110 As New Dictionary
Dim dicTitulos0110 As New Dictionary
Dim Chave As Variant, Campos
    
    Set dicTitulos0110 = Util.MapearTitulos(reg0110, 3)
    Set dicDados0110 = Util.CriarDicionarioRegistro(reg0110, "ARQUIVO")
    
    RegimePISCOFINS = 1
    For Each Chave In dicDados0110.Keys
        
        Campos = dicDados0110(Chave)
        
        If RegimePISCOFINS = 1 Then Campos(dicTitulos0110("COD_INC_TRIB")) = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_COD_INC_TRIB("1")
        If RegimePISCOFINS = 2 Then Campos(dicTitulos0110("COD_INC_TRIB")) = ValidacoesSPED.Contribuicoes.Enumeracoes.ValidarEnumeracao_COD_INC_TRIB("2")
        
        dicDados0110(Chave) = Campos
        
    Next Chave
    
    Call Util.LimparDados(reg0110, 4, False)
    Call Util.ExportarDadosDicionario(reg0110, dicDados0110)
    
    Call ReprocessarSugestoes
    
End Function

Public Function IgnorarInconsistencias()

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
    
    Set dicTitulos = Util.MapearTitulos(assApuracaoPISCOFINS, 3)
    Set Dados = assApuracaoPISCOFINS.Range("A4").CurrentRegion
    If Dados Is Nothing Then
        Call Util.MsgAlerta("Não foi encontrada nenhuma inconsistência!", "Inconsistências Fiscais")
        Exit Function
    End If
    
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
                
            End If
            
            If Linha.Row > 3 Then dicDados(Campos(dicTitulos("CHV_REG"))) = Campos
            
        End If
        
    Next Linha
    
    Call ReprocessarSugestoes
    
    If dicInconsistenciasIgnoradas.Count = 0 Then
        Call Util.MsgAlerta("Não existem Inconsistêncais a ignorar!", "Ignorar Inconsistências")
        Exit Function
    End If
    
    If assApuracaoPISCOFINS.AutoFilterMode Then assApuracaoPISCOFINS.AutoFilter.ShowAllData
    Call Util.LimparDados(assApuracaoPISCOFINS, 4, False)
    Call Util.ExportarDadosDicionario(assApuracaoPISCOFINS, dicDados)
    
    Call FuncoesFormatacao.AplicarFormatacao(ActiveSheet)
    Call FuncoesFormatacao.DestacarInconsistencias(assApuracaoPISCOFINS)
    
    Call Util.MsgInformativa("Inconsistências ignoradas com sucesso!", "Ignorar Inconsistências", Inicio)
    Application.StatusBar = False
    
End Function

Private Function AtualizarDadosRegistro(ByRef dicDados As Dictionary, ByRef dicTitulosReg As Dictionary, ByRef CamposReg As Variant, _
    ByRef dicTitulosRel As Dictionary, ByRef CamposRel As Variant, ByVal CHV_REG As String, Optional ByRef dicCorrelacoes As Dictionary, _
    Optional ByRef dicCorrelacoesInversas As Dictionary)
    
Dim dicCampos As Variant, Campo
Dim CampoDest As String
Dim i As Byte
    
    If dicDados.Exists(CHV_REG) Then
        
        dicCampos = dicDados(CHV_REG)
        
        Call AtualizarCampos(dicTitulosReg, CamposReg, dicTitulosRel, dicCampos, CamposRel, dicCorrelacoes)
        dicDados(CHV_REG) = dicCampos
        
    Else

        dicCampos = CriarRegistro(CamposRel, dicTitulosRel, CamposReg, dicTitulosReg, dicCorrelacoesInversas)
        dicDados(CHV_REG) = dicCampos

    End If
    
End Function

Private Function GerarChave0200(ByRef dicDados0140 As Dictionary, ByRef dicTitulos0140 As Dictionary, ByRef dicTitulos As Dictionary, _
    ByRef Campos As Variant) As String

Dim CNPJ As String, CHV_0140$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1 Else i = 0
    
    CNPJ = Campos(dicTitulos("CNPJ_ESTABELECIMENTO") - i)
    If dicDados0140.Exists(CNPJ) Then CHV_0140 = dicDados0140(CNPJ)(dicTitulos0140("CHV_REG"))
    
    GerarChave0200 = fnSPED.GerarChaveRegistro(CHV_0140, CStr(Campos(dicTitulos("COD_ITEM") - i)))
    
End Function

Private Function AtualizarCampos(ByRef dicTitulosReg As Dictionary, ByRef CamposReg As Variant, ByRef dicTitulos As Dictionary, _
    ByRef dicCampos As Variant, ByRef CamposRel As Variant, Optional ByRef dicCorrelacoes As Dictionary)

Dim CampoDest As String
Dim Campo As Variant
Dim Valor As Variant
Dim i As Long

    If LBound(dicCampos) = 0 Then i = 1 Else i = 0
    
    For Each Campo In CamposReg
        
        If Not dicCorrelacoes Is Nothing Then
            
            If dicCorrelacoes.Exists(Campo) Then CampoDest = dicCorrelacoes(Campo) Else CampoDest = Campo
            dicCampos(dicTitulosReg(CampoDest) - i) = CamposRel(dicTitulos(Campo))
        Else
    
            dicCampos(dicTitulosReg(Campo) - i) = CamposRel(dicTitulos(Campo))
            
        End If
        
    Next Campo
    
End Function

Private Function CriarRegistro(ByRef CamposRel As Variant, ByRef dicTitulosRel As Dictionary, _
    ByRef CamposReg As Variant, ByRef dicTitulosReg As Dictionary, Optional ByRef dicCorrelacoesInversas As Dictionary)

Dim Titulo As Variant
Dim arrRegistro As New ArrayList

    For Each Titulo In dicTitulosReg.Keys()
        
        If Not dicCorrelacoesInversas Is Nothing Then If dicCorrelacoesInversas.Exists(Titulo) Then Titulo = dicCorrelacoesInversas(Titulo)
            
        If dicTitulosRel.Exists(Titulo) Then
            
            arrRegistro.Add fnExcel.FormatarTipoDado(Titulo, CamposRel(dicTitulosRel(Titulo)))
        
        Else
        
            arrRegistro.Add ""
        
        End If
        
    Next Titulo
    
    CriarRegistro = arrRegistro.toArray()
    
End Function

Public Function ListarTributacoesPISCOFINS()

Dim Tributacao As New AssistenteTributario
    
    Call Tributacao.ListarTributacoes(assApuracaoPISCOFINS, assTributacaoPISCOFINS)
    
End Function

Private Function ProcessarValoresC181(ByVal Campos As Variant) As Variant

Dim VL_ITEM As Double
Dim CHV_PAI As String
Dim CamposC181 As Variant
    
    CHV_PAI = Util.RemoverAspaSimples(Campos(dicTitulos("CHV_PAI_CONTRIBUICOES")))
    VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM")), True, 2)
    
    CamposC181 = Array(VL_ITEM, 0, True, False)
    RegistrarValoresC181 CamposC181, CHV_PAI
    
End Function

Private Sub RegistrarValoresC181(ByVal Campos As Variant, ByVal Chave As String)

Dim CamposDic As Variant
    
    If dicDiferencaC181C185.Exists(Chave) Then
        
        CamposDic = dicDiferencaC181C185(Chave)
        Campos(0) = CDbl(Campos(0)) + CDbl(CamposDic(0))
        
    End If
    
    dicDiferencaC181C185(Chave) = Campos
    
End Sub

Private Function ProcessarValoresC185(ByVal Campos As Variant) As Variant

Dim VL_ITEM As Double
Dim CHV_PAI As String
Dim Camposc185 As Variant
    
    CHV_PAI = Util.RemoverAspaSimples(Campos(dicTitulos("CHV_PAI_CONTRIBUICOES")))
    VL_ITEM = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM")), True, 2)
    
    Camposc185 = Array(0, VL_ITEM, False, True)
    RegistrarValoresC185 Camposc185, CHV_PAI
    
End Function

Private Sub RegistrarValoresC185(ByVal Campos As Variant, ByVal Chave As String)

Dim CamposDic As Variant
    
    If dicDiferencaC181C185.Exists(Chave) Then
        
        CamposDic = dicDiferencaC181C185(Chave)
        Campos(0) = CDbl(Campos(1)) + CDbl(CamposDic(1))
        
    End If
    
    dicDiferencaC181C185(Chave) = Campos
    
End Sub

Private Function RegistrarDivergenciasC181C185()

Dim REG As String, CHV_PAI$
Dim Campos As Variant
Dim i As Long
    
    For i = 0 To arrRelatorio.Count - 1
        
        Campos = arrRelatorio.item(i)
        
        REG = Util.RemoverAspaSimples(Campos(dicTitulos("REG")))
        CHV_PAI = Util.RemoverAspaSimples(Campos(dicTitulos("CHV_PAI_CONTRIBUICOES")))
        
        If REG = "C181" Then RegistrarInconsistenciaC181 CHV_PAI, Campos
        If REG = "C185" Then RegistrarInconsistenciaC185 CHV_PAI, Campos
        
        arrRelatorio.item(i) = Campos
        
    Next i
    
End Function

Private Sub RegistrarInconsistenciaC181(ByVal Chave As String, ByRef Campos As Variant)

Dim VL_PIS As Double
Dim SUGESTAO As Byte
Dim VL_COFINS As Double
Dim CamposDic As Variant
Dim INCONSISTENCIA As Byte
Dim POSSUI_C185 As Boolean
    
    INCONSISTENCIA = UBound(Campos) - 1
    SUGESTAO = UBound(Campos)
    
    If dicDiferencaC181C185.Exists(Chave) Then
        
        CamposDic = dicDiferencaC181C185(Chave)
        
        VL_PIS = CamposDic(0)
        VL_COFINS = CamposDic(1)
        POSSUI_C185 = CamposDic(3)
        
        Select Case True
            
            Case VL_COFINS = 0 And Not POSSUI_C185
                Campos(INCONSISTENCIA) = "Nenhum registro C185 foi informado para o documento do campo CHV_PAI"
                Campos(SUGESTAO) = "Inclua os registros C185 ausentes para resolver os erros gerados no PVA"
                
            Case VL_PIS > VL_COFINS
                Campos(INCONSISTENCIA) = "O campo VL_ITEM do registro atual está maior que o do registro C185"
                Campos(SUGESTAO) = "Provável ausência de um ou mais registros C185 para o documento do campo CHV_PAI"
            
            Case VL_PIS < VL_COFINS
                Campos(INCONSISTENCIA) = "O campo VL_ITEM do registro atual está menor que o do registro C185"
                Campos(SUGESTAO) = "Provável ausência de um ou mais registros C181 para o documento do campo CHV_PAI"
                
        End Select
        
    End If
    
End Sub

Private Sub RegistrarInconsistenciaC185(ByVal Chave As String, ByRef Campos As Variant)

Dim VL_PIS As Double
Dim SUGESTAO As Byte
Dim VL_COFINS As Double
Dim CamposDic As Variant
Dim INCONSISTENCIA As Byte
Dim POSSUI_C181 As Boolean

    INCONSISTENCIA = UBound(Campos) - 1
    SUGESTAO = UBound(Campos)
    
    If dicDiferencaC181C185.Exists(Chave) Then
        
        CamposDic = dicDiferencaC181C185(Chave)
        
        VL_PIS = CamposDic(0)
        VL_COFINS = CamposDic(1)
        POSSUI_C181 = CamposDic(2)
        
        Select Case True
            
            Case VL_PIS = 0 And Not POSSUI_C181
                Campos(INCONSISTENCIA) = "Nenhum registro C181 foi informado para o documento do campo CHV_PAI"
                Campos(SUGESTAO) = "Inclua os registros C181 ausentes para resolver os erros gerados no PVA"
                
            Case VL_PIS > VL_COFINS
                Campos(INCONSISTENCIA) = "O campo VL_ITEM do registro atual está menor que o do registro C181"
                Campos(SUGESTAO) = "Provável ausência de um ou mais registros C185 para o documento do campo CHV_PAI"
                
            Case VL_PIS < VL_COFINS
                Campos(INCONSISTENCIA) = "O campo VL_ITEM do registro atual está maior que o do registro C181"
                Campos(SUGESTAO) = "Provável ausência de um ou mais registros C181 para o documento do campo CHV_PAI"
                
        End Select
        
    End If
    
End Sub

Public Function ObterStatusAtualizacaoM400M800() As Boolean
    
    ObterStatusAtualizacaoM400M800 = CBool(ConfiguracoesControlDocs.Range("ManterM400M800").value)
    
End Function

Public Sub DesativarAtualizacaoM400M800()
    
    ConfiguracoesControlDocs.Range("ManterM400M800").value = False
    If Not Rib Is Nothing Then Rib.InvalidateControl "chManterM400M800"
    
End Sub

Private Sub InicializarObjetos_Atualizacao()
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    Set GerenciadorSPED = New clsRegistrosSPED
    Set ExpReg = New ExportadorRegistros
    
    With GerenciadorSPED
        
        Call .CarregarDadosRegistro0000_Contr("ARQUIVO")
        Call .CarregarDadosRegistro0140("ARQUIVO", "CNPJ")
        Call .CarregarDadosRegistro0200("CHV_PAI_CONTRIBUICOES", "COD_ITEM")
        Call .CarregarDadosRegistroC100
        Call .CarregarDadosRegistroC170
        Call .CarregarDadosRegistroC175_Contr
        Call .CarregarDadosRegistroC181_Contr
        Call .CarregarDadosRegistroC185_Contr
        Call .CarregarDadosRegistroD201
        Call .CarregarDadosRegistroD205
        Call .CarregarDadosRegistroM001("ARQUIVO")
        Call .CarregarDadosRegistroM400("CHV_PAI_CONTRIBUICOES", "CST_PIS", "COD_CTA")
        Call .CarregarDadosRegistroM410("CHV_PAI_CONTRIBUICOES", "NAT_REC", "COD_CTA")
        Call .CarregarDadosRegistroM800("CHV_PAI_CONTRIBUICOES", "CST_COFINS", "COD_CTA")
        Call .CarregarDadosRegistroM810("CHV_PAI_CONTRIBUICOES", "NAT_REC", "COD_CTA")
        
    End With
    
End Sub

Private Sub LimparObjetos_Atualizacao()
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    Set GerenciadorSPED = Nothing
    Set ExpReg = Nothing
    
    Call Util.HabilitarControles
    
End Sub
