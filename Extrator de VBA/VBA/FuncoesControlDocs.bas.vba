Attribute VB_Name = "FuncoesControlDocs"
Option Explicit

Public Function ObterUuidComputador() As String

Dim objWMIService As Object
Dim colItems As Object
Dim objItem As Object
Dim strComputer As String
    
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct", , 48)
    
    For Each objItem In colItems
        ObterUuidComputador = objItem.Uuid
    Next
    
End Function

Public Function VersionarProjeto()

Dim Projeto As VBIDE.VBProject
Dim VersaoProjeto As String

    On Error Resume Next
    Set Projeto = ThisWorkbook.VBProject
    
    VersaoProjeto = "ControlDocsProject_v" & Util.ApenasNumeros(ThisWorkbook.name)
    If Projeto.name <> VersaoProjeto Then
        Projeto.name = VersaoProjeto
    End If
    
End Function

Public Function ResetarAssinatura() As Boolean

Dim Nomes As Variant, NOME
    
    On Error Resume Next
    Nomes = Array("Vencimento_Assinatura", "Ultima_Consulta", "Email_Assinante", "status", "plano", "uuid")

    For Each NOME In Nomes
        Names.Add NOME, ""
    Next NOME
    
End Function

Public Function ConsultarStatusAssinatura(ByVal Funcao As String, Optional ByVal EmitirMsg As Boolean, Optional ByVal URLTeste As Boolean)

Dim https As New WinHttp.WinHttpRequest
Dim Uuid As String, versao$, Msg$, URL$, dispositivo$
Dim Resposta As VbMsgBoxResult
Dim json As Object
    
    On Error Resume Next
    
    Set https = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    'Coleta email do assinante
    EmailAssinante = VBA.LCase(VBA.Trim(relGestaoAssinatura.Range("email_cliente").value))
    
    'Verifica se o e-mail do Assinante está preenchido e se é valido
    If EmailAssinante = "" Or Not Funcoes.ValidarEmail(EmailAssinante) Then
        
        If EmitirMsg Then
            
            MsgBox "Informe um e-mail válido para autenticar sua assinatura.", vbExclamation, "Email não informado ou inválido"
            relGestaoAssinatura.Activate
            relGestaoAssinatura.Range("email_cliente").Activate
        
        End If
        
        GoTo Desativar:
    
    End If
    

    'Coleta a versão da ferramenta
    versao = ExtrairVersao
    
    'Coleta o Uuid do Computador
    Uuid = ObterUuidComputador
    dispositivo = VBA.Environ("COMPUTERNAME")
    
    With https
        
        'Define se a url utilizada será a padrão ou a de teste
        If URLTeste Then URL = urlTestesControlDocs Else URL = urlControlDocs
        
        'Configura os parâmetros da requisição
        .Open "POST", URL, False
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "TokenControlDocs", TokenControlDocs
        .Send "{""versao"": " & versao & ", ""funcao"": """ & Funcao & """, ""email"": """ & EmailAssinante & """, ""dispositivo"": """ & dispositivo & """, ""uuid"": """ & Uuid & """}"
        
        '#Verificações
'        Debug.Print url
'        Debug.Print "{""versao"": " & Versao & ", ""funcao"": """ & Funcao & """, ""email"": """ & EmailAssinante & """, ""dispositivo"": """ & dispositivo & """, ""uuid"": """ & Uuid & """}"
        
        Do
            If .Status <> 0 Then Exit Do
        Loop
        
        'Debug.Print .ResponseText
        'Converte a resposta da API num dicionário
        Set json = JsonConverter.ParseJson(.ResponseText)
        
        'Avisa ao usuário caso ocorre algum erro na requisição
        If json Is Nothing Then
            
            Msg = "Ocorreu um erro na autenticação." & vbCrLf
            Msg = Msg & "Por favor, tire um print dessa mensagem e envie ao suporte para que possamos te ajudar." & vbCrLf & vbCrLf
            Msg = Msg & .ResponseText
            
            Call Util.MsgAlerta(Msg, "Erro na autenticação")
            GoTo Desativar:
            
        End If
        
        'Atualiza informações da assinatura do usuário
        If json.Exists("status") Then Application.Names.Add "status", json("status") Else Application.Names.Add "status", ""
        If json.Exists("plano") Then Application.Names.Add "plano", json("plano") Else Application.Names.Add "plano", ""
        If json.Exists("vencimento") Then Application.Names.Add "vencimento", json("vencimento") Else Application.Names.Add "vencimento", ""
        
        If json.Exists("status") Then
            
            Select Case True
                
                Case json("status") = "ACTIVE"
                    If EmitirMsg Then
                    
                        If json("plano") Like "*experimental*" And Funcao = "ASSINATURA_EXPERIMENTAL" Then
                        
                            Msg = "Sua assinatura Experimental do ControlDocs foi ativada com sucesso!" & vbCrLf & vbCrLf
                            Msg = Msg & "Em até 5 minutos você receberá um e-mail com o link para nossas aulas práticas. Ele contém o guia necessário para você iniciar o uso da ferramenta."
                            Call Util.MsgAviso(Msg, "Teste Experimental do ControlDocs")

                        Else
                            
                            Msg = "Sua assinatura ControlDocs foi ativada com sucesso!" & vbCrLf & vbCrLf
                            Msg = Msg & "Agora você já pode desfrutar de uma rotina mais rápida, prática e segura."
                            Call Util.MsgAviso(Msg, "Assinatura Ativada")
                        
                        End If
                        
                    End If
                    
                Case json("status") Like "*CANCELLED*"
                    Msg = "A sua assinatura do ControlDocs está cancelada!" & vbCrLf & vbCrLf
                    Msg = Msg & "Por favor, contrate um plano ControlDocs para começar a aproveitar uma rotina mais rápida, prática e segura."
                    Call Util.MsgAlerta(Msg, "Assinatura Cancelada")
                    GoTo Desativar:
                    
                Case json("status") = "DELAYED"
                    Msg = "A sua assinatura do ControlDocs está atrasada!" & vbCrLf & vbCrLf
                    Msg = Msg & "Por favor, renove a sua assinatura para continuar aproveitando uma rotina mais rápida, prática e segura."
                    Call Util.MsgAlerta(Msg, "Assinatura Atrasada")
                    GoTo Desativar:
                    
                Case json("status") = "FINISH"
                    Msg = "A sua assinatura Experimental do ControlDocs está finalizada!" & vbCrLf & vbCrLf
                    Msg = Msg & "Por favor, contrate um plano ControlDocs para começar a aproveitar uma rotina mais rápida, prática e segura."
                    Call Util.MsgAlerta(Msg, "Assinatura Atrasada")
                    GoTo Desativar:
                    
                Case json("status") = "INACTIVE"
                    Msg = "A contratação da sua assinatura ControlDocs não foi finalizada!" & vbCrLf & vbCrLf
                    Msg = Msg & "Por favor, contrate um plano ControlDocs para começar a aproveitar uma rotina mais rápida, prática e segura."
                    Call Util.MsgAlerta(Msg, "Assinatura Atrasada")
                    GoTo Desativar:
                    
                Case Else
                    Call VerificarMensagemAPI(json("mensagem"))
                    GoTo Desativar:
                    
            End Select

        End If
        
        If json.Exists("mensagem") Then
            
            Call VerificarMensagemAPI(json("mensagem"))
            
        End If
        
    End With
    
Exit Function
Desativar:

    If Err.Number = -2147012867 Then
        
        Msg = "Não foi possível estabelecer uma conexão com o servidor." & vbCrLf & vbCrLf
        Msg = Msg & "Por favor, verifique sua conexão com a internet e refaça a sua autenticação do ControlDocs."
        Call Util.MsgAlerta(Msg, "Falha na conexão com o servidor")
    
    End If
    
    Call FuncoesControlDocs.ResetarAssinatura
    
End Function

Public Function ExtrairVersao()

Dim Projeto As VBIDE.VBProject

    Set Projeto = ThisWorkbook.VBProject
    ExtrairVersao = Util.ApenasNumeros(Projeto.name)

End Function

Public Function ExtrairVersaoProjeto() As String

Dim Projeto As VBIDE.VBProject
    
    Set Projeto = ThisWorkbook.VBProject
    ExtrairVersaoProjeto = VBA.Split(Projeto.name, "_")(1)
    
End Function

Public Function ObterEmailAssinante() As String
    
    ObterEmailAssinante = VBA.LCase(VBA.Trim(relGestaoAssinatura.Range("email_cliente").value))
    
End Function

Public Function ListarDispositivosControlDocs()

Dim Uuid As String, versao$, Msg$, URL$, Funcao$
Dim https As New WinHttp.WinHttpRequest
Dim arrDispositivos As New ArrayList
Dim Projeto As VBIDE.VBProject
Dim Resposta As VbMsgBoxResult
Dim dispositivo As Variant
Dim json As Object
Dim i As Byte

    On Error GoTo Tratar:
    
    Set https = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    'Coleta email do assinante
    EmailAssinante = ObterEmailAssinante
    
    'Verifica se o e-mail do Assinante está preenchido e se é valido
    If EmailAssinante = "" Or Not Funcoes.ValidarEmail(EmailAssinante) Then
            
        MsgBox "Informe um e-mail válido para autenticar sua assinatura.", vbExclamation, "Email não informado ou inválido"
        relGestaoAssinatura.Activate
        relGestaoAssinatura.Range("email_cliente").Activate
        GoTo Tratar:
    
    End If
    
    Call Util.LimparDados(relGestaoAssinatura, 4, False)
    
    With relGestaoAssinatura
    
        .Range("C2").value = ""
        .Range("E2").value = ""
        .Range("G2").value = ""
        .Range("I2").value = ""
        
    End With
        
    'Coleta a versão da ferramenta
    Set Projeto = ThisWorkbook.VBProject
    versao = Util.ApenasNumeros(Projeto.name)
    Funcao = "LISTAR_DISPOSITIVOS"
    
    'Coleta o Uuid do Computador
    Uuid = ObterUuidComputador
    dispositivo = Environ("COMPUTERNAME")
    
    With https
                
        'Configura os parâmetros da requisição
        .Open "POST", urlControlDocs, False
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "TokenControlDocs", TokenControlDocs
        .Send "{""versao"": " & versao & ", ""funcao"": """ & Funcao & """, ""email"": """ & EmailAssinante & """, ""dispositivo"": """ & dispositivo & """, ""uuid"": """ & Uuid & """}"
        
        Application.StatusBar = "Consultando dados da sua assinatura ControlDocs, por favor aguarde..."
        Do
            If .Status <> 0 Then Exit Do
        Loop
        
        'Converte a resposta da API num dicionário
        Set json = JsonConverter.ParseJson(.ResponseText)
        
        'Avisa ao usuário caso ocorre algum erro na requisição
        If json Is Nothing Then
            
            Msg = "Ocorreu um erro na autenticação." & vbCrLf
            Msg = Msg & "Por favor, tire um print dessa mensagem e envie ao suporte para que possamos te ajudar." & vbCrLf & vbCrLf
            Msg = Msg & .ResponseText
            
            Call Util.MsgAlerta(Msg, "Erro na autenticação")
            GoTo Tratar:
            
        End If
        
        i = 0
        For Each dispositivo In json("dispositivos")
            i = i + 1
            arrDispositivos.Add Array(i, dispositivo("dispositivo"), dispositivo("uuid"))
        Next dispositivo
        
        With relGestaoAssinatura

            .Range("E2").value = VBA.UCase(json("plano"))
            .Range("G2").value = json("vencimento")
            .Range("I2").value = json("qtdDispositivos")
            .Range("K2").value = VBA.UCase(json("status"))

        End With

        Call Util.ExportarDadosArrayList(relGestaoAssinatura, arrDispositivos)
        
    End With
    
    Call Util.MsgAviso("Os dados da sua assinatura ControlDocs foram baixados com sucesso!", "Consulta de Assinatura ControlDocs")
    Application.StatusBar = False
    
Exit Function
Tratar:

    If Err.Number = -2147012867 Then
        
        Msg = "Não foi possível estabelecer uma conexão com o servidor." & vbCrLf & vbCrLf
        Msg = Msg & "Por favor, verifique sua conexão com a internet e refaça a sua autenticação do ControlDocs."
        Call Util.MsgAlerta(Msg, "Falha na conexão com o servidor")
    
    End If
    
    Call FuncoesControlDocs.ResetarAssinatura
    
End Function

Public Function ResetarControlDocs()

Const Excecoes As String = "ConfiguracoesControlDocs, CadContrib"
Dim Resposta As VbMsgBoxResult
Dim pasta As New Workbook
Dim Plan As Worksheet
    
    Resposta = MsgBox("Tem certeza que deseja apagar os dados de TODOS registros do ControlDocs?" & vbCrLf & _
                      "Essa operação NÃO pode ser desfeita.", vbCritical + vbYesNo, "Apagar registros ControlDocs")
    
    If Resposta = vbYes Then
        
        Inicio = Now()
        Call Util.DesabilitarControles
            
            Application.StatusBar = "Resetando ControlDocs, por favor aguarde..."
            
            Call FuncoesControlDocs.EnviarAcionamentos("RELATORIO_ACIONAMENTOS")
            
            Set pasta = ThisWorkbook
            For Each Plan In pasta.Sheets
                
                If Not Excecoes Like "*" & Plan.CodeName & "*" Then Call Util.OtimizarCelulas(Plan, 3)
                
            Next Plan
            
            Call Util.LimparDados(CadContrib, 4, False)
            
            Call Util.MsgInformativa("Registros deletados com sucesso!", "Limpeza de registros ControlDocs", Inicio)
            Application.StatusBar = False
        
        Call Util.HabilitarControles
        
    End If
    
End Function

Public Sub ExtrairChavesSPED()

Dim Vencimento As String, Plano$, Status$, versao$, Uuid$, Conteudo$, jsonString$
Dim arrChaves As New ArrayList
Dim http As New WinHttp.WinHttpRequest
Dim Arqs As Variant, Arq, Campos
Dim Projeto As VBIDE.VBProject
Dim vbResult As VbMsgBoxResult
Dim arrArqs As New ArrayList
Dim json As New Dictionary
Dim Maquinas As Byte
Dim Msg As String
    
    Arqs = Util.SelecionarArquivos("txt", "Selecione os SPEDs")
    For Each Arq In Arqs
        jsonString = Util.GerarJson(Arq)
    Next Arq
    
    Call Util.ExportarTxt("C:\Users\marcu\OneDrive\Pasta5\Área de Trabalho\SOS BORRACHAS\APURAÇÃO 12-2023\TesteJson.txt", jsonString)
    Set Projeto = ThisWorkbook.VBProject
    versao = VBA.Right(Projeto.name, 3)
    
    EmailAssinante = relGestaoAssinatura.Range("email_cliente").value
    If Not Funcoes.ValidarEmail(EmailAssinante) Then
        MsgBox "Informe um e-mail válido para consultar sua assinatura.", vbExclamation, "Email não informado ou inválido"
        Call FuncoesControlDocs.ResetarAssinatura
        relGestaoAssinatura.Range("email_cliente").Activate
        Exit Sub
    End If

    Application.StatusBar = "Enviando dados dos arquivos selecionados, por favor aguarde..."
    
    With http
        
        .Open "POST", urlControlDocs, False
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "TokenControlDocs", TokenControlDocs
        .Send "{""versao"": " & versao & ", ""funcao"": ""EXTRAIR_CHAVES"", ""arquivos"": " & jsonString & ", ""email"": """ & EmailAssinante & """}"
        
        Do
            If .Status <> 0 Then Exit Do
        Loop
        
        If .Status = 200 Then
            
            Set json = JsonConverter.ParseJson(.ResponseText)
            
            If json.Exists("chaves") Then
                Set arrChaves = json("chaves")
            End If
            
            Call Util.LimparDados(relGestaoAssinatura, 4, False)
            Call Util.ExportarListaJson(relGestaoAssinatura, arrChaves)
            
            Select Case json("mensagem")
                
                Case "Chaves extraídas com sucesso"
                    Msg = "Chaves extraídas com sucesso!" & vbCrLf & vbCrLf
                    Msg = Msg & "As chaves de acesso do registro C100 foram extraídas com sucesso."
                    
                    Call Util.MsgAviso(Msg, "Extração de Chaves")
                    
                Case "Versão desatualizada"
                    Call FuncoesControlDocs.ResetarAssinatura
                    
                    Msg = "Sua versão do ControlDocs está desatualizada!" & vbCrLf & vbCrLf
                    Msg = Msg & "Clique em SIM para fazer o download da última versão agora mesmo!"
                    
                    vbResult = MsgBox(Msg, vbExclamation + vbYesNo, "Versão Desatualizada!")
                    If vbResult = vbYes Then Call FuncoesLinks.AbrirUrl(DownloadControlDocs)
                    
                Case "Assinatura não encontrada"
                    Call FuncoesControlDocs.ResetarAssinatura
                    Msg = "Não foi encontrada nenhuma assinatura para o e-mail informado!" & vbCrLf & vbCrLf
                    Msg = Msg & "Por favor, verifique o e-mail digitado e tente novamente."
                    
                    Call Util.MsgAlerta(Msg, "Assinatura não Encontrada")
                    
            End Select
            
        End If
        
    End With
    
    Application.StatusBar = False
    
End Sub

Function ConversonJson(jsonString As String) As Dictionary
    
Dim json As New Dictionary
Dim ParesJson() As String
Dim Par As Variant
Dim ChaveValor() As String

    jsonString = Replace(jsonString, "{", "")
    jsonString = Replace(jsonString, "}", "")
    jsonString = Replace(jsonString, """", "")
    
    ParesJson = Split(jsonString, ", ")
    For Each Par In ParesJson
        
        ChaveValor = Split(Par, ": ")
        If UBound(ChaveValor) = 1 Then
            json(ChaveValor(0)) = ChaveValor(1)
        End If
        
    Next
    
    Set ConversonJson = json
    
End Function

Public Function RegistrarAcionamento(ByVal Botao As String)

Dim PlanAcionamentos As Worksheet
Dim NovoAcionamento As String
Dim Acionamentos As String
Dim UltLin As Long
    
    'Define a planilha onde os dados serão armazenados
    Set PlanAcionamentos = ThisWorkbook.Sheets("Acionamentos")
    
    'Formata o novo acionamento como JSON
    NovoAcionamento = "{""acionamento"":""" & Botao & """, ""horario"":""" & Format(Now(), "dd/mm/yyyy hh:nn:ss") & """}"
    
    'Encontra a última linha preenchida
    UltLin = Util.UltimaLinha(PlanAcionamentos, "A")
    
    'Adiciona o novo acionamento
    PlanAcionamentos.Cells(UltLin + 1, 1).value = NovoAcionamento
    
End Function

Public Function ContarAcionamentos() As Long
    ContarAcionamentos = Util.UltimaLinha(ThisWorkbook.Sheets("Acionamentos"), "A")
End Function

Public Function CarregarAcionamentos() As String

Dim DadosAcionamentos As Variant, Acionamento
Dim PlanAcionamentos As Worksheet
Dim Acionamentos As String
Dim UltLin As Long
    
    'Define a planilha onde os dados estão armazenados
    Set PlanAcionamentos = ThisWorkbook.Sheets("Acionamentos")
    
    'Encontrar a última linha da coluna A
    UltLin = Util.UltimaLinha(PlanAcionamentos, "A")
    
    'Se não houver dados, retornar uma string vazia
    If UltLin = 0 Then
    
        CarregarAcionamentos = "[]"
        Exit Function
        
    End If
    
    'Carrega os acionamentos em um array
    DadosAcionamentos = Application.Transpose(PlanAcionamentos.Range("A2:A" & UltLin).value)
    
    If VarType(DadosAcionamentos) = vbString Then
    
        CarregarAcionamentos = "[" & DadosAcionamentos & "]"
        Exit Function
        
    End If
    
    'Concatena todos os acionamentos em uma única string JSON
    Acionamentos = "["
    For Each Acionamento In DadosAcionamentos
        
        If Acionamentos = "[" Then
            
            Acionamentos = Acionamentos & Acionamento
        
        Else
        
            Acionamentos = Acionamentos & "," & Acionamento
        
        End If
        
    Next Acionamento
    
    Acionamentos = Acionamentos & "]"
    
    ' Retornar os acionamentos concatenados
    CarregarAcionamentos = Acionamentos
    
End Function

Public Function EnviarAcionamentos(ByVal Funcao As String, Optional ByVal URLTeste As Boolean)

Dim Uuid As String, versao$, URL$, jsonAcionamentos$
Dim https As New WinHttp.WinHttpRequest
Dim Projeto As VBIDE.VBProject
        
    Set https = CreateObject("WinHttp.WinHttpRequest.5.1")
            
    'Coleta a versão da ferramenta
    Set Projeto = ThisWorkbook.VBProject
    versao = Util.ApenasNumeros(Projeto.name)
    
    'Coleta o Uuid do Computador
    Uuid = ObterUuidComputador
    
    'Coleta os dados dos acionamentos
    jsonAcionamentos = CarregarAcionamentos()
    
    With https
        
        'Define se a url utilizada será a padrão ou a de teste
        If URLTeste Then URL = urlTestesControlDocs Else URL = urlControlDocs
        
        'Configura os parâmetros da requisição
        .Open "POST", URL, False
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "TokenControlDocs", TokenControlDocs
        .Send "{""versao"": " & versao & ", ""funcao"": """ & Funcao & """, ""uuid"": """ & Uuid & """, ""acionamentos"": " & jsonAcionamentos & "}"
        
        Do
            If .Status <> 0 Then Exit Do
        Loop

        If .Status = 200 Then Call Util.LimparDados(Acionamentos, 1, False)
    
    End With
    
End Function

Public Sub CarregarDadosPlanilha()

Dim arrChaves As New ArrayList
Dim pasta As New Workbook
Dim Plan As Worksheet
Dim Arq As Variant
Dim dicNomesPlanilhas As New Dictionary
    
    Call Util.CarregarChavesAcessoDoces(arrChaves)
    Arq = Application.GetOpenFilename("Arquivos Excel Binarios (*.xlsb), *.xlsb,Arquivos Excel (*.xlsm), *.xlsm ", , "Selecione a planilha", , False)
    
    If Arq <> False Then
        
        Inicio = Now()
        
        Set pasta = Workbooks.Open(Arq)
        
        Application.StatusBar = "Listando dados da pasta de trabalho selecionada."
        For Each Plan In pasta.Sheets
            
            dicNomesPlanilhas(Plan.CodeName) = Plan.name
            
        Next Plan
        
        For Each Plan In ThisWorkbook.Sheets
            
            If dicNomesPlanilhas.Exists(Plan.CodeName) Then
                'Debug.Print "É sim" & Plan.CodeName
                'Call Util.CriarDicionarioRegistro(Plan)
                'Call Util.ExportarDadosDicionario(this)
            End If
            
        Next Plan
        
        pasta.Close
        
        Util.MsgInformativa "Dados importados com sucesso!", "Importação de dados", Inicio
        
    End If
    
End Sub

Private Function VerificarMensagemAPI(ByRef Mensagem As String)

Dim Resposta As VbMsgBoxResult
Dim Msg As String
    
    Select Case True
    
        Case VBA.LCase(Mensagem) = "versão desatualizada"
            Msg = "A sua versão do ControlDocs está desatualizada!" & vbCrLf & vbCrLf
            Msg = Msg & "Clique em 'SIM' para atualizar para a última versão."
            Resposta = MsgBox(Msg, vbExclamation + vbYesNo, "Versão Desatualizada")
            If Resposta = vbYes Then Call FuncoesLinks.AbrirUrl(DownloadControlDocs)
            GoTo Desativar:
        
        Case VBA.LCase(Mensagem) = "a assinatura experimental do controldocs já foi ativada para esse dispositivo"
            Msg = "A Assinatura Experimental do ControlDocs já foi ativada para esse dispositivo."
            Call Util.MsgAlerta(Msg, "Dispositivo já Cadastrado")
            GoTo Desativar:
        
        Case VBA.LCase(Mensagem) = "assinatura vencida"
            Msg = "Sua assinatura ControlDocs está vencida!" & vbCrLf & vbCrLf
            Msg = Msg & "Por favor, renove sua assinatura para continuar aproveitando uma rotina mais rápida, prática e segura."
            Call Util.MsgAlerta(Msg, "Assinatura Vencida")
            GoTo Desativar:
            
        Case VBA.LCase(Mensagem) = "assinatura não encontrada"
            Msg = "Não foi encontrada uma assinatura para o e-mail informado!" & vbCrLf & vbCrLf
            Msg = Msg & "Por favor, verifique se o e-mail foi digitado corretamente e tente novamente."
            Call Util.MsgAlerta(Msg, "Assinatura Não Encontrada")
            GoTo Desativar:
            
        Case VBA.LCase(Mensagem) = "limite de dispositivos atingido"
            Msg = "O limite de dispositivos para sua assinatura do ControlDocs foi atingido!" & vbCrLf & vbCrLf
            Msg = Msg & "Se precisar utilizar o ControlDocs em mais máquinas entre em contato o suporte para contratar licenças adicionais."
            Call Util.MsgAlerta(Msg, "Limite de Dispositivos Atingido")
            GoTo Desativar:
        
        Case VBA.LCase(Mensagem) = "dispositivo não cadastrado para essa assinatura! por favor faça o processo de autenticação para registrar o dispositivo."
            Msg = "Este dispositivo não possui cadastro para essa assinatura do ControlDocs!" & vbCrLf & vbCrLf
            Msg = Msg & "Por favor, faça o processo de autenticação para registrar o dispositivo."
            Call Util.MsgAlerta(Msg, "Dispositivo sem cadastro")
            GoTo Desativar:
        
        Case VBA.LCase(Mensagem) = "essa assinatura ainda não possui nenhum dispositivo cadastrado!"
            Msg = "Essa Assinatura ainda não possui nenhum dispositivo cadastrado!" & vbCrLf & vbCrLf
            Msg = Msg & "Por favor, faça o processo de autenticação para registrar o dispositivo."
            Call Util.MsgAlerta(Msg, "Dispositivo sem cadastro")
            GoTo Desativar:
            
        Case VBA.LCase(Mensagem) = "erro ao processar a sua solicitação. por favor entre em contato com o suporte."
            Msg = "Erro ao processar a sua solicitação!" & vbCrLf & vbCrLf
            Msg = Msg & "Por favor entre em contato com o suporte"
            GoTo Desativar:
            
    End Select
    
Exit Function
Desativar:

    If Err.Number = -2147012867 Then
        
        Msg = "Não foi possível estabelecer uma conexão com o servidor." & vbCrLf & vbCrLf
        Msg = Msg & "Por favor, verifique sua conexão com a internet e refaça a sua autenticação do ControlDocs."
        Call Util.MsgAlerta(Msg, "Falha na conexão com o servidor")
    
    End If
    
    Call FuncoesControlDocs.ResetarAssinatura
    
End Function
