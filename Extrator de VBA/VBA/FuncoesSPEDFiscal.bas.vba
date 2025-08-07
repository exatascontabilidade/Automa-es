Attribute VB_Name = "FuncoesSPEDFiscal"
Option Explicit
Option Base 1

Private EnumContrib As New clsEnumeracoesSPEDContribuicoes

Public Function ImportarSPED(Optional ByVal SelReg As String, Optional Periodo As String, Optional Centralizar As Boolean)

Dim CNPJ As String, UF$, IE$, DT_INI$, reg0000$, versao$, versaoContrib$
Dim Arqs As Variant, Arq, Registros, Registro, Campos
Dim Comeco As Double, Comeco2#, Comeco3#
Dim arrDados As New ArrayList
Dim MsgArq As String
Dim MsgReg As String
Dim Chave As Variant
Dim nReg As String
Dim b As Long
    
    Arqs = Util.SelecionarArquivos("txt")
    If VarType(Arqs) = vbBoolean Then Exit Function
    
    Inicio = Now()
    
    Call Util.DesabilitarControles
    
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    Call dicChavesRegistroSPED.RemoveAll
    Call Util.DesabilitarControles
    Call CarregarNiveisRegistros
    
    If Centralizar Then
        
        For Each Arq In Arqs
            
            If fnSPED.ClassificarSPED(Arq) = "Fiscal" Then
                
                Registro = fnSPED.ExtrairRegistroAbertura(Arq)
                Campos = VBA.Split(Registro, "|")
                nReg = Campos(1)
                
                Select Case nReg
                    
                    Case "0000"
                        
                        CNPJ = Campos(7)
                        If VBA.Mid(CNPJ, 9, 4) = "0001" Then
                            
                            UF = Campos(9)
                            DT_INI = Campos(4)
                            reg0000 = Registro
                            Exit For
                            
                        End If
                        
                    Case Else
                        Exit For
                        
                End Select
                
                If reg0000 <> "" Then Exit For
                
            End If
            
        Next Arq
        
    End If
    
    a = 0
    Comeco = Timer
    For Each Arq In Arqs
        
        MsgArq = "Importando SPED Fiscal " & a + 1 & " de " & UBound(Arqs)
        
        Application.StatusBar = "Verificando o tipo do SPED importado."
        If fnSPED.ClassificarSPED(Arq, versao, Periodo) <> "Fiscal" Then
            
            Call Util.MsgAlerta("O SPED importado não é o fiscal.", "Verificação Tipo SPED")
            Application.StatusBar = False
            
            Exit Function
            
        Else
            
            Application.StatusBar = "Carregando layout do SPED Fiscal, por favor aguarde..."
            Call CarregarLayoutSPEDFiscal(versao)
            
            versaoContrib = EnumContrib.ValidarEnumeracao_COD_VER(Periodo)
            Call CarregarLayoutSPEDContribuicoes(versaoContrib)
            
        End If
        
        Call Util.AntiTravamento(a, 1, MsgArq, UBound(Arqs) + 1, Comeco)
        Registros = Util.ImportarTxt(Arq)
        
        b = 0
        Comeco2 = Timer
        For Each Registro In Registros
            
            Call Util.AntiTravamento(b, 100, MsgArq & " [Importando registros do bloco " & VBA.Left(nReg, 1) & "]", UBound(Registros) + 1, Comeco2)
            If VBA.Trim(Registro) <> "" And Registro Like "*|*" And VBA.Len(Registro) > 6 Then
                
                Campos = VBA.Split(Registro, "|")
                nReg = Campos(1)
                
                If Centralizar And nReg = "0000" Then
                    
                    If Campos(7) Like VBA.Left(CNPJ, 8) & "*" And Campos(4) Like DT_INI And Campos(9) = UF Then
                        Registro = reg0000
                        Campos = VBA.Split(reg0000, "|")
                    End If
                    
                End If
                
                If SelReg <> "" Then
                    
                    Select Case True
                        
                        Case SelReg = nReg Or nReg = "0000" Or (VBA.Right(nReg, 3) = "001" And nReg <> "9001")
                            If Not dicRegistros.Exists(nReg) Then Set dicRegistros(nReg) = New ArrayList
                            Call fnSPED.ProcessarRegistro(Registro, Periodo)
                            
                    End Select
                    
                Else
                    
                    If Not dicRegistros.Exists(nReg) Then Set dicRegistros(nReg) = New ArrayList
                    Call fnSPED.ProcessarRegistro(Registro, Unificar:=Centralizar)
                    If nReg = "9001" Then Exit For
                    
                End If
                
            End If
            
            If SelReg <> "" Then If VBA.Left(nReg, 2) > VBA.Left(SelReg, 2) Then Exit For
            
        Next Registro
        
    Next Arq
    
    b = 0
    Comeco3 = Timer
    For Each Chave In dicRegistros.Keys()
        Call Util.AntiTravamento(b, 10, MsgArq & " [Exportando dados do bloco " & VBA.Left(Chave, 1) & " para o Excel]")
        Set arrDados = dicRegistros(Chave)
        If VBA.Len(Chave) < 4 Then Chave = VBA.Format(Chave, "0000")
        On Error Resume Next
            Call Util.ExportarDadosArrayList(Worksheets(Chave), arrDados)
        On Error GoTo 0
    Next Chave
    
    Call dicRegistros.RemoveAll
    
    If Centralizar Then
        Call Util.AgruparRegistros(reg1400)
    End If
    
    Application.StatusBar = "Exportação finalizada com sucesso!"
    Call Util.MsgInformativa("SPED Fiscal importado com sucesso!", "Importação do SPED Fiscal", Inicio)
    
    If Otimizacoes.OtimizacoesAtivas Then Call Otimizacoes.SugerirSomadoIPIaoItem
    
    Call Util.AtualizarBarraStatus(False)
    
    Call Util.HabilitarControles
    
End Function

Public Function ImportarSPEDFiscalparaAnalise(Optional ByVal regSelect As String)

Dim Arqs As Variant, Arq, Registros, Registro
Dim nReg As String
Dim b As Long
    
    Arqs = Util.SelecionarArquivos("txt")
    If VarType(Arqs) <> 11 Then
        
        Inicio = Now()
        Call ZerarDicionariosEFD
        Call dicChavesRegistroSPED.RemoveAll
        
        a = 0
        Comeco = Timer
        For Each Arq In Arqs
            
            Call Util.AntiTravamento(a, 1, "Importando dados do SPED Fiscal, por favor aguarde...", UBound(Arqs) + 1, Comeco)
            
            b = 0
            Application.StatusBar = "Carregando dados do SPED para memória do computador, por favor aguarde..."
            Registros = Util.ImportarTxt(Arq)
            For Each Registro In Registros
                
                Call Util.AntiTravamento(b, 100, "Importando registros do SPED Fiscal", UBound(Registros) + 1, Comeco)
                With regEFD
                    
                    If Registro <> "" Then
                        nReg = Mid(Registro, 2, 4)
                        Select Case True
                            
                            Case (regSelect = nReg And nReg = "0000") Or (regSelect = "" And nReg = "0000")
                                Call r0000.ImportarParaExcel(Registro, .dic0000)
                                
                            Case (regSelect = nReg And nReg = "0150") Or (regSelect = "" And nReg = "0150")
                                Call r0150.ImportarParaAnalise(Registro, .dic0150)
                                
                            Case (regSelect = nReg And nReg = "C100") Or (regSelect = "" And nReg = "C100")
                                Call rC100.ImportarParaAnalise(Registro, .dicC100)
                                
                        End Select
                        
                    End If
                    
                End With
                
            Next Registro
            
        Next Arq
        
        Application.StatusBar = "Exportando registros do SPED para o relatório, por favor aguarde..."
        
        Call Util.LimparDados(relDivergencias, 4, False)
        Call Util.ExportarDadosDicionario(relDivergencias, regEFD.dicC100)
        Call FuncoesFormatacao.DeletarFormatacao
        Call FuncoesFormatacao.AplicarFormatacao(relDivergencias)
        Call FuncoesFormatacao.FormatarDivergencias(relDivergencias)
        
        'Call ZerarDicionariosEFD
        
        Call Util.MsgInformativa("Dados extraídos com sucesso", "Extração de Dados do SPED", Inicio)
        
    End If
    
End Function

Public Sub AlterarChavesSPED()

Dim Registro As Variant, Registros, Campos
Dim EFD As New ArrayList
Dim nReg As String
Dim Arqs, Arq
    
    Arqs = Util.SelecionarArquivos("txt")
    If VarType(Arqs) <> 11 Then
        
        Inicio = Now()
        For Each Arq In Arqs
            
            a = Util.AntiTravamento(a)
            
            Registros = Util.ImportarTxt(Arq)
            For Each Registro In Registros
                
                a = Util.AntiTravamento(a)
                nReg = VBA.Mid(Registro, 2, 4)
                Select Case True
                    
                    Case nReg = "0000"
                        
                        Campos = fnSPED.ExtrairCamposRegistro(Registro)
                        If Campos(7) <> "" Then
                            Campos(6) = "ESCOLA DA AUTOMAÇÃO FISCAL LTDA"
                            Campos(7) = "23896999000102"
                        End If
                        Registro = fnSPED.GerarRegistro(Campos)
                        EFD.Add Registro
                        
                    Case nReg = "0005"
                        
                        Campos = fnSPED.ExtrairCamposRegistro(Registro)
                            Campos(2) = "ESCOLA DA AUTOMAÇÃO FISCAL"
                        Registro = fnSPED.GerarRegistro(Campos)
                        EFD.Add Registro
                        
                    Case nReg = "C100"
                        
                        Campos = fnSPED.ExtrairCamposRegistro(Registro)
                        If Campos(3) = "0" And Campos(9) <> "" Then Campos(9) = FuncoesTXT.AlterarChaveEAF(Campos(9))
                        Registro = fnSPED.GerarRegistro(Campos)
                        EFD.Add Registro
                        
                    Case nReg = "C113"
                        
                        Campos = fnSPED.ExtrairCamposRegistro(Registro)
                        If Campos(3) = "0" And Campos(10) <> "" Then Campos(10) = FuncoesTXT.AlterarChaveEAF(Campos(10))
                        Registro = fnSPED.GerarRegistro(Campos)
                        EFD.Add Registro
                        
                    Case nReg = "D100"
                        
                        Campos = fnSPED.ExtrairCamposRegistro(Registro)
                        If Campos(3) = "0" And Campos(10) <> "" Then Campos(10) = FuncoesTXT.AlterarChaveEAF(Campos(9))
                        Registro = fnSPED.GerarRegistro(Campos)
                        EFD.Add Registro
                        
                    Case Registro <> ""
                        EFD.Add Registro
                        
                End Select
                
            Next Registro
            
            Arq = Replace(VBA.UCase(Arq), ".TXT", " - ESTRUTURADO.txt")
            Call fnSPED.ExportarSPED(Arq, fnSPED.TotalizarRegistrosSPED(EFD))
            
        Next Arq
        
    End If
    
    regC190.Activate
    Call Util.MsgInformativa("SPED estruturado com sucesso!", "Estruturação do SPED Fiscal", Inicio)
    
End Sub

Public Sub EstruturarSPED()

Dim Registro As Variant, Registros As Variant
Dim EFD As New ArrayList
Dim nReg As String
Dim Arqs, Arq
    
    Arqs = Util.SelecionarArquivos("txt")
    If VarType(Arqs) <> 11 Then
        
        Inicio = Now()
        For Each Arq In Arqs
            
            a = Util.AntiTravamento(a)
            
            Registros = Util.ImportarTxt(Arq)
            For Each Registro In Registros
                
                a = Util.AntiTravamento(a)
                nReg = VBA.Mid(Registro, 2, 4)
                Select Case True
                                        
                    Case Registro <> ""
                        EFD.Add Registro
                        
                End Select
                
            Next Registro
            
            Arq = Replace(VBA.UCase(Arq), ".TXT", " - ESTRUTURADO.txt")
            Call fnSPED.ExportarSPED(Arq, fnSPED.TotalizarRegistrosSPED(EFD))
            
        Next Arq
        
        Call Util.MsgInformativa("SPED estruturado com sucesso!", "Estruturação do SPED Fiscal", Inicio)
        
    End If
    
End Sub

Public Sub ExtrairrelICMS()

Dim Registro As Variant, Registros As Variant
Dim EFD As New ArrayList
Dim nReg As String
Dim Arqs, Arq
    
    Arqs = Util.SelecionarArquivos("txt")
    If VarType(Arqs) <> 11 Then
        
        Inicio = Now()
        For Each Arq In Arqs
            
            a = Util.AntiTravamento(a)
            
            Registros = Util.ImportarTxt(Arq)
            For Each Registro In Registros
                
                a = Util.AntiTravamento(a)
                nReg = VBA.Mid(Registro, 2, 4)
                Select Case True
                                        
                    Case Registro <> ""
                        EFD.Add Registro
                        
                End Select
                
            Next Registro
            
            Arq = Replace(VBA.UCase(Arq), ".TXT", " - ESTRUTURADO.txt")
            Call fnSPED.ExportarSPED(Arq, fnSPED.TotalizarRegistrosSPED(EFD))
            
        Next Arq
        
    End If
    
    regC190.Activate
    Call Util.MsgInformativa("SPED estruturado com sucesso!", "Estruturação do SPED Fiscal", Inicio)
    
End Sub

Public Sub ListarRegistros()

Dim Registro As Variant, Registros As Variant
Dim EFD As New Dictionary
Dim nReg As String
Dim Arqs, Arq
    
    Arqs = Util.SelecionarArquivos("txt")
    If VarType(Arqs) <> 11 Then
        
        Inicio = Now()
        For Each Arq In Arqs
            
            a = Util.AntiTravamento(a)
            Registros = Util.ImportarTxt(Arq)
            For Each Registro In Registros
                
                nReg = VBA.Mid(Registro, 2, 4)
                Select Case True
                                        
                    Case Registro <> ""
                        EFD(fnSPED.ExtrairCampo(Registro, 1)) = Array(fnSPED.ExtrairCampo(Registro, 1))
                        
                End Select
                
            Next Registro
            
        Next Arq
        
        Call Util.ExportarDadosDicionario(reg0000, EFD)
    End If
    
    reg0000.Activate
    Call Util.MsgInformativa("SPED estruturado com sucesso!", "Estruturação do SPED Fiscal", Inicio)
    
End Sub

Public Sub GerarEFDICMSIPI()

Dim dicEstruturaSPED As New Dictionary
Dim dicRegistros As New Dictionary
Dim Msg As String, Caminho$, Arq$, versao$
Dim dicTitulos As New Dictionary
Dim Registros As New Dictionary
Dim Registro As Variant, Campos
Dim dtRef As String, dtFin$
Dim EFD As New ArrayList
Dim UltLin As Long
    
    UltLin = Util.UltimaLinha(reg0000, "A")
    If UltLin < 4 Then
        Call Util.MsgAlerta("Sem dados do SPED Fiscal para exportar.", "Exportação do SPED Fiscal")
        Exit Sub
    End If
    
    If fnSPED.ChecarErrosEstruturaisICMSIPI(Msg) Then
        
        Call Util.MsgAlerta(Msg, "Validação de Erros de Estrutura da EFD-ICMS/IPI")
        Exit Sub
        
    End If
    
    If Otimizacoes.OtimizacoesAtivas Then Call Otimizacoes.SugerirSomadoIPIaoItem
    
    Caminho = Util.SelecionarPasta("Selecione a pasta para Salvar os arquivos")
    If Caminho = "" Then Exit Sub
    
    Inicio = Now()
    Call Util.DesabilitarControles
        
        Set dicTitulos = Util.MapearTitulos(reg0000, 3)
        Set Registros = Util.CriarDicionarioRegistro(reg0000)
        
        For Each Campos In Registros.Items
            
            versao = Campos(dicTitulos("COD_VER"))
            
            Application.StatusBar = "Etapa 1/5 - Carregando layout do SPED Fiscal, por favor aguarde..."
            Call FuncoesJson.CarregarEstruturaSPEDFiscal(dicEstruturaSPED, versao)
            
            Application.StatusBar = "Etapa 2/5 - Carregando registros do SPED Contribuições, por favor aguarde..."
            Set dicRegistros = Util.CarregarRegistrosSPED(dicEstruturaSPED)
            
            Call Util.TratarParticularidadesRegistros(dicRegistros)
            
            Call Util.AntiTravamento(a, 1, "Exportando arquivo " & a + 1 & " de " & Registros.Count, Registros.Count, Comeco)
            
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                With Campos0000
                    
                    .CHV_REG = VBA.Replace(Campos(dicTitulos("CHV_REG")), "'", "")
                    .COD_VER = VBA.Replace(Campos(dicTitulos("COD_VER")), "'", "")
                    .IND_ATIV = Util.ApenasNumeros(Campos(dicTitulos("IND_ATIV")))
                    .ARQUIVO = Util.RemoverAspaSimples(Campos(dicTitulos("ARQUIVO")))
                    .CNPJ = Util.ApenasNumeros(Campos(dicTitulos("CNPJ")))
                    .COD_FIN = Util.ApenasNumeros(Campos(dicTitulos("COD_FIN")))
                    .DT_INI = VBA.Format(Util.ApenasNumeros(Campos(dicTitulos("DT_INI"))), "yyyymm")
                    dtFin = VBA.Format(Util.ApenasNumeros(Campos(dicTitulos("DT_FIN"))), "yyyy-mm-dd")
                    If .COD_FIN = 0 Then .COD_FIN = " Remessa do arquivo original " Else .COD_FIN = " Remessa do arquivo substituto "
                    
                    dtRef = VBA.Format(.DT_INI & "01", "0000-00-00")
                    Arq = Caminho & "\EFD ICMS-IPI " & .CNPJ & .COD_FIN & .DT_INI & " - ControlDocs.txt"
                    
                    Call fnSPED.IncluirRegistro(EFD, dicEstruturaSPED, dicRegistros, "0000", .CHV_REG)
                    Call FuncoesAjusteSPED.RemoverCadastrosNaoReferenciados(Arq, EFD)
                    
                End With
                
            End If
            
            EFD.Clear
            
        Next Campos
        
        If dicRegistros("0000").Count = 1 Then Msg = "Arquivo gerado com sucesso!" Else Msg = "Arquivos gerados com sucesso!"
        Call Util.MsgInformativa(Msg, "Geração da EFD-ICMS/IPI", Inicio)
        
    Call Util.HabilitarControles
    Application.StatusBar = False
    
End Sub

Public Function ZerarDicionariosEFD()

    With regEFD

        'Bloco 0
        .dic0000.RemoveAll
        .dic0001.RemoveAll
        .dic0002.RemoveAll
        .dic0005.RemoveAll
        .dic0015.RemoveAll
        .dic0100.RemoveAll
        .dic0150.RemoveAll
        .dic0175.RemoveAll
        .dic0190.RemoveAll
        .dic0200.RemoveAll
        .dic0205.RemoveAll
        .dic0206.RemoveAll
        .dic0210.RemoveAll
        .dic0220.RemoveAll
        .dic0221.RemoveAll
        .dic0300.RemoveAll
        .dic0305.RemoveAll
        .dic0400.RemoveAll
        .dic0450.RemoveAll
        .dic0460.RemoveAll
        .dic0500.RemoveAll
        .dic0600.RemoveAll
        .dic0990.RemoveAll


        'Bloco B
        .dicB001.RemoveAll
        .dicB020.RemoveAll
        .dicB025.RemoveAll
        .dicB030.RemoveAll
        .dicB035.RemoveAll
        .dicB350.RemoveAll
        .dicB420.RemoveAll
        .dicB440.RemoveAll
        .dicB460.RemoveAll
        .dicB470.RemoveAll
        .dicB500.RemoveAll
        .dicB510.RemoveAll
        .dicB990.RemoveAll


        'Bloco C
        .dicC001.RemoveAll
        .dicC100.RemoveAll
        .dicC101.RemoveAll
        .dicC105.RemoveAll
        .dicC110.RemoveAll
        .dicC111.RemoveAll
        .dicC112.RemoveAll
        .dicC113.RemoveAll
        .dicC114.RemoveAll
        .dicC115.RemoveAll
        .dicC116.RemoveAll
        .dicC120.RemoveAll
        .dicC130.RemoveAll
        .dicC140.RemoveAll
        .dicC141.RemoveAll
        .dicC160.RemoveAll
        .dicC165.RemoveAll
        .dicC170.RemoveAll
        .dicC171.RemoveAll
        .dicC172.RemoveAll
        .dicC173.RemoveAll
        .dicC174.RemoveAll
        .dicC175.RemoveAll
        .dicC176.RemoveAll
        .dicC177.RemoveAll
        .dicC178.RemoveAll
        .dicC179.RemoveAll
        .dicC180.RemoveAll
        .dicC181.RemoveAll
        .dicC185.RemoveAll
        .dicC186.RemoveAll
        .dicC190.RemoveAll
        .dicC191.RemoveAll
        .dicC195.RemoveAll
        .dicC197.RemoveAll
        .dicC300.RemoveAll
        .dicC310.RemoveAll
        .dicC320.RemoveAll
        .dicC321.RemoveAll
        .dicC330.RemoveAll
        .dicC350.RemoveAll
        .dicC370.RemoveAll
        .dicC380.RemoveAll
        .dicC390.RemoveAll
        .dicC400.RemoveAll
        .dicC405.RemoveAll
        .dicC410.RemoveAll
        .dicC420.RemoveAll
        .dicC425.RemoveAll
        .dicC430.RemoveAll
        .dicC460.RemoveAll
        .dicC465.RemoveAll
        .dicC470.RemoveAll
        .dicC480.RemoveAll
        .dicC490.RemoveAll
        .dicC495.RemoveAll
        .dicC500.RemoveAll
        .dicC510.RemoveAll
        .dicC590.RemoveAll
        .dicC591.RemoveAll
        .dicC595.RemoveAll
        .dicC597.RemoveAll
        .dicC600.RemoveAll
        .dicC601.RemoveAll
        .dicC610.RemoveAll
        .dicC690.RemoveAll
        .dicC700.RemoveAll
        .dicC790.RemoveAll
        .dicC791.RemoveAll
        .dicC800.RemoveAll
        .dicC810.RemoveAll
        .dicC815.RemoveAll
        .dicC850.RemoveAll
        .dicC855.RemoveAll
        .dicC857.RemoveAll
        .dicC860.RemoveAll
        .dicC870.RemoveAll
        .dicC880.RemoveAll
        .dicC890.RemoveAll
        .dicC895.RemoveAll
        .dicC897.RemoveAll
        .dicC990.RemoveAll


        'Bloco D
        .dicD001.RemoveAll
        .dicD100.RemoveAll
        .dicD101.RemoveAll
        .dicD110.RemoveAll
        .dicD120.RemoveAll
        .dicD130.RemoveAll
        .dicD140.RemoveAll
        .dicD150.RemoveAll
        .dicD160.RemoveAll
        .dicD161.RemoveAll
        .dicD162.RemoveAll
        .dicD170.RemoveAll
        .dicD180.RemoveAll
        .dicD190.RemoveAll
        .dicD195.RemoveAll
        .dicD197.RemoveAll
        .dicD300.RemoveAll
        .dicD301.RemoveAll
        .dicD310.RemoveAll
        .dicD350.RemoveAll
        .dicD355.RemoveAll
        .dicD360.RemoveAll
        .dicD365.RemoveAll
        .dicD370.RemoveAll
        .dicD390.RemoveAll
        .dicD400.RemoveAll
        .dicD410.RemoveAll
        .dicD411.RemoveAll
        .dicD420.RemoveAll
        .dicD500.RemoveAll
        .dicD510.RemoveAll
        .dicD530.RemoveAll
        .dicD590.RemoveAll
        .dicD600.RemoveAll
        .dicD610.RemoveAll
        .dicD690.RemoveAll
        .dicD695.RemoveAll
        .dicD696.RemoveAll
        .dicD697.RemoveAll
        .dicD700.RemoveAll
        .dicD730.RemoveAll
        .dicD731.RemoveAll
        .dicD735.RemoveAll
        .dicD737.RemoveAll
        .dicD750.RemoveAll
        .dicD760.RemoveAll
        .dicD761.RemoveAll
        .dicD990.RemoveAll


        'Bloco E
        .dicE001.RemoveAll
        .dicE100.RemoveAll
        .dicE110.RemoveAll
        .dicE111.RemoveAll
        .dicE112.RemoveAll
        .dicE113.RemoveAll
        .dicE115.RemoveAll
        .dicE116.RemoveAll
        .dicE200.RemoveAll
        .dicE210.RemoveAll
        .dicE220.RemoveAll
        .dicE230.RemoveAll
        .dicE240.RemoveAll
        .dicE250.RemoveAll
        .dicE300.RemoveAll
        .dicE310.RemoveAll
        .dicE311.RemoveAll
        .dicE312.RemoveAll
        .dicE313.RemoveAll
        .dicE316.RemoveAll
        .dicE500.RemoveAll
        .dicE510.RemoveAll
        .dicE520.RemoveAll
        .dicE530.RemoveAll
        .dicE531.RemoveAll
        .dicE990.RemoveAll


        'Bloco G
        .dicG001.RemoveAll
        .dicG110.RemoveAll
        .dicG125.RemoveAll
        .dicG126.RemoveAll
        .dicG130.RemoveAll
        .dicG140.RemoveAll
        .dicG990.RemoveAll


        'Bloco H
        .dicH001.RemoveAll
        .dicH005.RemoveAll
        .dicH010.RemoveAll
        .dicH020.RemoveAll
        .dicH030.RemoveAll
        .dicH990.RemoveAll


        'Bloco K
        .dicK001.RemoveAll
        .dicK010.RemoveAll
        .dicK100.RemoveAll
        .dicK200.RemoveAll
        .dicK210.RemoveAll
        .dicK215.RemoveAll
        .dicK220.RemoveAll
        .dicK230.RemoveAll
        .dicK235.RemoveAll
        .dicK250.RemoveAll
        .dicK255.RemoveAll
        .dicK260.RemoveAll
        .dicK265.RemoveAll
        .dicK270.RemoveAll
        .dicK275.RemoveAll
        .dicK280.RemoveAll
        .dicK290.RemoveAll
        .dicK291.RemoveAll
        .dicK292.RemoveAll
        .dicK300.RemoveAll
        .dicK301.RemoveAll
        .dicK302.RemoveAll
        .dicK990.RemoveAll

        'Bloco M
        .dicM001.RemoveAll
        .dicM100.RemoveAll
        .dicM200.RemoveAll
        .dicM210.RemoveAll
        .dicM500.RemoveAll
        .dicM600.RemoveAll
        .dicM610.RemoveAll
        .dicM990.RemoveAll


        'Bloco 1
        .dic1001.RemoveAll
        .dic1010.RemoveAll
        .dic1100.RemoveAll
        .dic1105.RemoveAll
        .dic1110.RemoveAll
        .dic1200.RemoveAll
        .dic1210.RemoveAll
        .dic1250.RemoveAll
        .dic1255.RemoveAll
        .dic1300.RemoveAll
        .dic1310.RemoveAll
        .dic1320.RemoveAll
        .dic1350.RemoveAll
        .dic1360.RemoveAll
        .dic1370.RemoveAll
        .dic1390.RemoveAll
        .dic1391.RemoveAll
        .dic1400.RemoveAll
        .dic1500.RemoveAll
        .dic1510.RemoveAll
        .dic1600.RemoveAll
        .dic1601.RemoveAll
        .dic1700.RemoveAll
        .dic1710.RemoveAll
        .dic1800.RemoveAll
        .dic1900.RemoveAll
        .dic1910.RemoveAll
        .dic1920.RemoveAll
        .dic1921.RemoveAll
        .dic1922.RemoveAll
        .dic1923.RemoveAll
        .dic1925.RemoveAll
        .dic1926.RemoveAll
        .dic1960.RemoveAll
        .dic1970.RemoveAll
        .dic1975.RemoveAll
        .dic1980.RemoveAll
        .dic1990.RemoveAll

    End With

End Function

Public Function ExportarRegistroSelecionado(ByVal nReg As String)
    
Dim ExpReg As New ExportadorRegistros
    
    Call ExpReg.ExportarRegistros(nReg)
    Exit Function
    
    With regEFD
        
        Select Case nReg
            
            'Bloco 0
            Case "0000"
                Call Util.ExportarDadosDicionario(reg0000, .dic0000)
            
            Case "0001"
                Call Util.ExportarDadosDicionario(reg0000, .dic0001)
            
            Case "0002"
                Call Util.ExportarDadosDicionario(reg0002, .dic0002)
                
            Case "0005"
                Call Util.ExportarDadosDicionario(reg0005, .dic0005)
                
            Case "0015"
                Call Util.ExportarDadosDicionario(reg0015, .dic0015)
                
            Case "0100"
                Call Util.ExportarDadosDicionario(reg0100, .dic0100)
                
            Case "0150"
                Call Util.ExportarDadosDicionario(reg0150, .dic0150)
                
            Case "0175"
                Call Util.ExportarDadosDicionario(reg0175, .dic0175)
                
            Case "0190"
                Call Util.ExportarDadosDicionario(reg0190, .dic0190)
                
            Case "0200"
                Call Util.ExportarDadosDicionario(reg0200, .dic0200)
                
            Case "0205"
                Call Util.ExportarDadosDicionario(reg0205, .dic0205)
            
            Case "0206"
                Call Util.ExportarDadosDicionario(reg0206, .dic0206)
                
            Case "0210"
                Call Util.ExportarDadosDicionario(reg0210, .dic0210)
                
            Case "0220"
                Call Util.ExportarDadosDicionario(reg0220, .dic0220)
            
            Case "0221"
                Call Util.ExportarDadosDicionario(reg0221, .dic0221)
                
            Case "0300"
                Call Util.ExportarDadosDicionario(reg0300, .dic0300)
                
            Case "0305"
                Call Util.ExportarDadosDicionario(reg0305, .dic0305)
                
            Case "0400"
                Call Util.ExportarDadosDicionario(reg0400, .dic0400)
                
            Case "0450"
                Call Util.ExportarDadosDicionario(reg0450, .dic0450)
                
            Case "0460"
                Call Util.ExportarDadosDicionario(reg0460, .dic0460)
                
            Case "0500"
                Call Util.ExportarDadosDicionario(reg0500, .dic0500)
            
            Case "0600"
                Call Util.ExportarDadosDicionario(reg0600, .dic0600)
                
            Case "0990"
                Call Util.ExportarDadosDicionario(reg0990, .dic0990)
            
            'Bloco B
            Case "B001"
                Call Util.ExportarDadosDicionario(regB001, .dicB001)
            
            Case "B990"
                Call Util.ExportarDadosDicionario(regB990, .dicB990)
                    
            'Bloco C
            Case "C001"
                Call Util.ExportarDadosDicionario(regC001, .dicC001)
            
            Case "C100"
                Call Util.ExportarDadosDicionario(regC100, .dicC100)
                
                    
            Case "C101"
                Call Util.ExportarDadosDicionario(regC101, .dicC101)
                    
            Case "C110"
                Call Util.ExportarDadosDicionario(regC110, .dicC110)
                        
            Case "C111"
                Call Util.ExportarDadosDicionario(regC111, .dicC111)
                
            Case "C112"
                Call Util.ExportarDadosDicionario(regC112, .dicC112)
                
            Case "C113"
                Call Util.ExportarDadosDicionario(regC113, .dicC113)
                    
            Case "C114"
                Call Util.ExportarDadosDicionario(regC114, .dicC114)
                
            Case "C115"
                Call Util.ExportarDadosDicionario(regC115, .dicC115)
            
            Case "C116"
                Call Util.ExportarDadosDicionario(regC116, .dicC116)
            
            Case "C120"
                Call Util.ExportarDadosDicionario(regC116, .dicC120)
                
            Case "C140"
                Call Util.ExportarDadosDicionario(regC140, .dicC140)
                
            Case "C141"
                Call Util.ExportarDadosDicionario(regC141, .dicC141)
                
            Case "C170"
                Call Util.ExportarDadosDicionario(regC170, .dicC170)
                
            Case "C171"
                Call Util.ExportarDadosDicionario(regC171, .dicC171)
                
            Case "C190"
                Call Util.ExportarDadosDicionario(regC190, .dicC190)
                
            Case "C191"
                Call Util.ExportarDadosDicionario(regC191, .dicC191)
                
            Case "C195"
                Call Util.ExportarDadosDicionario(regC195, .dicC195)
                
            Case "C197"
                Call Util.ExportarDadosDicionario(regC197, .dicC197)
                
            Case "C400"
                Call Util.ExportarDadosDicionario(regC400, .dicC400)
                                
            Case "C405"
                Call Util.ExportarDadosDicionario(regC405, .dicC405)
                
            Case "C410"
                Call Util.ExportarDadosDicionario(regC410, .dicC410)
                
            Case "C420"
                Call Util.ExportarDadosDicionario(regC420, .dicC420)
                
            Case "C425"
                Call Util.ExportarDadosDicionario(regC425, .dicC425)
                
            Case "C430"
                Call Util.ExportarDadosDicionario(regC430, .dicC430)
                
            Case "C460"
                Call Util.ExportarDadosDicionario(regC460, .dicC460)
                
            Case "C465"
                Call Util.ExportarDadosDicionario(regC465, .dicC465)
                
            Case "C470"
                Call Util.ExportarDadosDicionario(regC470, .dicC470)
                
            Case "C480"
                Call Util.ExportarDadosDicionario(regC480, .dicC480)
                
            Case "C490"
                Call Util.ExportarDadosDicionario(regC490, .dicC490)
                
            Case "C500"
                Call Util.ExportarDadosDicionario(regC500, .dicC500)
                                
            Case "C510"
                Call Util.ExportarDadosDicionario(regC510, .dicC510)
                
            Case "C590"
                Call Util.ExportarDadosDicionario(regC590, .dicC590)
                                        
            Case "C990"
                Call Util.ExportarDadosDicionario(regC990, .dicC990)
                    
            'Bloco D
                
            Case "D001"
                Call Util.ExportarDadosDicionario(regD001, .dicD001)
                            
            Case "D100"
                Call Util.ExportarDadosDicionario(regD100, .dicD100)
                                                
            Case "D101"
                Call Util.ExportarDadosDicionario(regD101, .dicD101)
                            
            Case "D190"
                Call Util.ExportarDadosDicionario(regD190, .dicD190)
                            
            Case "D195"
                Call Util.ExportarDadosDicionario(regD195, .dicD195)
                            
            Case "D197"
                Call Util.ExportarDadosDicionario(regD197, .dicD197)
                                            
            Case "D990"
                Call Util.ExportarDadosDicionario(regD990, .dicD990)
            
            'Bloco E
                                            
            Case "E001"
                Call Util.ExportarDadosDicionario(regE001, .dicE001)
                                                            
            Case "E100"
                Call Util.ExportarDadosDicionario(regE100, .dicE100)
                                                            
            Case "E110"
                Call Util.ExportarDadosDicionario(regE110, .dicE110)
                                                            
            Case "E111"
                Call Util.ExportarDadosDicionario(regE111, .dicE111)
                                                            
            Case "E112"
                Call Util.ExportarDadosDicionario(regE112, .dicE112)
                                                            
            Case "E113"
                Call Util.ExportarDadosDicionario(regE113, .dicE113)
                                                            
            Case "E115"
                Call Util.ExportarDadosDicionario(regE115, .dicE115)
                                                            
            Case "E116"
                Call Util.ExportarDadosDicionario(regE116, .dicE116)
                                                            
            Case "E200"
                Call Util.ExportarDadosDicionario(regE200, .dicE200)
                                                            
            Case "E210"
                Call Util.ExportarDadosDicionario(regE210, .dicE210)
                                                            
            Case "E220"
                Call Util.ExportarDadosDicionario(regE220, .dicE220)
                                                            
            Case "E230"
                Call Util.ExportarDadosDicionario(regE230, .dicE230)
                                                            
            Case "E240"
                Call Util.ExportarDadosDicionario(regE240, .dicE240)
                                                            
            Case "E250"
                Call Util.ExportarDadosDicionario(regE250, .dicE250)
                                                            
            Case "E300"
                Call Util.ExportarDadosDicionario(regE300, .dicE300)
                                                            
            Case "E310"
                Call Util.ExportarDadosDicionario(regE310, .dicE310)
                                                            
            Case "E311"
                Call Util.ExportarDadosDicionario(regE311, .dicE311)
                                                            
            Case "E312"
                Call Util.ExportarDadosDicionario(regE312, .dicE312)
                                                            
            Case "E313"
                Call Util.ExportarDadosDicionario(regE313, .dicE313)
                                                            
            Case "E316"
                Call Util.ExportarDadosDicionario(regE316, .dicE316)
                                                            
            Case "E500"
                Call Util.ExportarDadosDicionario(regE500, .dicE500)
                                                            
            Case "E510"
                Call Util.ExportarDadosDicionario(regE510, .dicE510)
                                                            
            Case "E520"
                Call Util.ExportarDadosDicionario(regE520, .dicE520)
                                                            
            Case "E530"
                Call Util.ExportarDadosDicionario(regE530, .dicE530)
                                                            
            Case "E531"
                Call Util.ExportarDadosDicionario(regE531, .dicE531)
                                                                                        
            Case "E990"
                Call Util.ExportarDadosDicionario(regE990, .dicE990)
            
            'Bloco G
            Case "G001"
                Call Util.ExportarDadosDicionario(regG001, .dicG001)
            
            Case "G990"
                Call Util.ExportarDadosDicionario(regG990, .dicG990)
            
            'Bloco H
            Case "H001"
                Call Util.ExportarDadosDicionario(regH001, .dicH001)
            
            Case "H005"
                Call Util.ExportarDadosDicionario(regH005, .dicH005)
            
            Case "H010"
                Call Util.ExportarDadosDicionario(regH010, .dicH010)
            
            Case "H020"
                Call Util.ExportarDadosDicionario(regH020, .dicH020)
            
            Case "H030"
                Call Util.ExportarDadosDicionario(regH030, .dicH030)
            
            Case "H990"
                Call Util.ExportarDadosDicionario(regH990, .dicH990)
            
            'Bloco K
            Case "K001"
                Call Util.ExportarDadosDicionario(regK001, .dicK001)
            
            Case "K010"
                Call Util.ExportarDadosDicionario(regK010, .dicK010)
            
            Case "K100"
                Call Util.ExportarDadosDicionario(regK100, .dicK100)
            
            Case "K200"
                Call Util.ExportarDadosDicionario(regK200, .dicK200)
            
            Case "K230"
                Call Util.ExportarDadosDicionario(regK230, .dicK230)
            
            Case "K280"
                Call Util.ExportarDadosDicionario(regK280, .dicK280)
            
            Case "K990"
                Call Util.ExportarDadosDicionario(regK990, .dicK990)
            
            'Bloco M
            Case "M001"
                Call Util.ExportarDadosDicionario(regM001, .dicM001)
            
            Case "M100"
                Call Util.ExportarDadosDicionario(regM100, .dicM100)
            
            Case "M200"
                Call Util.ExportarDadosDicionario(regM200, .dicM200)
            
            Case "M210"
                Call Util.ExportarDadosDicionario(regM210, .dicM210)
            
            Case "M500"
                Call Util.ExportarDadosDicionario(regM500, .dicM500)
            
            Case "M600"
                Call Util.ExportarDadosDicionario(regM600, .dicM600)
            
            Case "M610"
                Call Util.ExportarDadosDicionario(regM610, .dicM610)
            
            Case "M990"
                Call Util.ExportarDadosDicionario(regM990, .dicM990)
                
            'Bloco 1
            Case "1001"
                Call Util.ExportarDadosDicionario(reg1001, .dic1001)
            
            Case "1010"
                Call Util.ExportarDadosDicionario(reg1010, .dic1010)
            
            Case "1300"
                Call Util.ExportarDadosDicionario(reg1300, .dic1300)
            
            Case "1310"
                Call Util.ExportarDadosDicionario(reg1310, .dic1310)
            
            Case "1320"
                Call Util.ExportarDadosDicionario(reg1320, .dic1320)
            
            Case "1350"
                Call Util.ExportarDadosDicionario(reg1350, .dic1350)
            
            Case "1360"
                Call Util.ExportarDadosDicionario(reg1360, .dic1360)
            
            Case "1370"
                Call Util.ExportarDadosDicionario(reg1370, .dic1370)
            
            Case "1400"
                Call Util.ExportarDadosDicionario(reg1400, .dic1400)
            
            Case "1601"
                Call Util.ExportarDadosDicionario(reg1601, .dic1601)
            
            Case "1990"
                Call Util.ExportarDadosDicionario(reg1990, .dic1990)
        
        End Select
        
    End With
        
End Function

Private Function IncluirRegistro(ByRef EFD As ArrayList, ByVal nReg As String, ByVal Chave As String)

Dim Registro As Variant, Dados, nCampos, Campos, Campo, REG, regFilho, Titulos
Dim CHV_REG, CHV_PAI
Dim titCampos As New ArrayList
Dim arrCampos As New ArrayList
Dim arrRegistros As New ArrayList
Dim dicRegistro As New Dictionary
Dim dicTitulos As New Dictionary
Dim arrTitulos As New ArrayList
Dim i As Long, UltLin As Long
Dim Intervalo As Range
Dim teste As Variant
    
    'dicNomes.Exists(nReg) And
    If dicRegistros.Exists(nReg) Then
        
        'Definindo dicionário com os títulos das colunas da planilha de dados
        Set dicTitulos = Util.MapearTitulos(Worksheets(nReg), 3)
        Set arrTitulos = Util.TransformarDicionarioArrayList(dicTitulos)

        'Definindo a lista de campos que vão para o SPED
        Set titCampos = dicLayoutFiscal(nReg)("NomeCampos")
        
        'Selecionando o registro para pegar os dados para exportação
        Set dicRegistro = dicRegistros(nReg)
        
        'Criando ArrayList com todos os dados de exportação de uma chave específica
        Set arrRegistros = dicRegistro(Chave)
        
        For Each Campos In arrRegistros
                
            arrCampos.Add ""
            For i = LBound(Campos) To UBound(Campos)
                
                'Coletando CHV_REG e CHV_PAI de cada registro
                CHV_REG = Campos(dicTitulos("CHV_REG"))
                CHV_PAI = Campos(dicTitulos("CHV_PAI_FISCAL"))
                
                'Verificando se o campo em questão faz parte do SPED e o adiciona em caso afirmativo
                If titCampos.contains(arrTitulos.item(i - 1)) Then arrCampos.Add Campos(i)
                
            Next i
            arrCampos.Add ""
                    
            'Inserindo registro montado no SPED
            EFD.Add fnSPED.GerarRegistro(arrCampos.toArray)
            
            'Limpando dados do registro montado
            arrCampos.Clear
            
            'Identificando registros filhos do registro Atual
            If dicLayoutFiscal.Exists(nReg) Then
                If dicLayoutFiscal(nReg).Exists("Filhos") Then
                    For Each regFilho In dicLayoutFiscal(nReg)("Filhos")
                        'Chama a rotina para incluir os dados do registro filho caso ele exista
                        If dicRegistros.Exists(regFilho) Then
                            If dicRegistros.Exists(regFilho) Then If dicRegistros(regFilho).Exists(CHV_REG) Then Call IncluirRegistro(EFD, regFilho, CHV_REG)
                        End If
                    Next regFilho
                End If
            End If
            
        Next Campos
        
    End If
    
End Function

Public Function ListarArquivosCentralizados(ByVal Arqs As Variant, ByRef reg0000 As String) As ArrayList

Dim Arq As Variant, Campos
Dim Registro As String, nReg$, CNPJ_MATRIZ$, CNPJ$, UF$, DT_INI$
Dim arrSPEDSCentralizados As New ArrayList
    
    For Each Arq In Arqs
        
        Registro = fnSPED.ExtrairRegistroAbertura(Arq)
        Campos = VBA.Split(Registro, "|")
        nReg = Campos(1)
        
        Select Case True
            
            Case nReg = "0000"
                
                CNPJ = Campos(7)
                If VBA.Mid(CNPJ, 9, 4) = "0001" Then
                    
                    CNPJ_MATRIZ = Campos(7)
                    UF = Campos(9)
                    DT_INI = Campos(4)
                    reg0000 = Registro
                    arrSPEDSCentralizados.Add Arq
                    GoTo PrxArq:
                    
                Else
                    
                    Select Case True
                        
                        Case CNPJ Like VBA.Left(CNPJ_MATRIZ, 8) & "*" And Campos(4) Like DT_INI And Campos(9) = UF
                            arrSPEDSCentralizados.Add Arq
                        
                        Case Else
                            GoTo PrxArq:
                        
                    End Select
                    
                End If
                
            Case Else
                Exit For
                
        End Select
PrxArq:
    Next Arq
    
    Set ListarArquivosCentralizados = arrSPEDSCentralizados
    
End Function


