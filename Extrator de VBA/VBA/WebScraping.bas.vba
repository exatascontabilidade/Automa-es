Attribute VB_Name = "WebScraping"
Option Explicit

Public Function ExtrairCadastroContribuinteWeb()

Dim URL As String, CNPJ$, InscEst$, Razao$, UF$
Dim Tabelas As MSHTML.IHTMLElementCollection
Dim Linhas As MSHTML.IHTMLElementCollection
Dim colunas As MSHTML.IHTMLElementCollection
Dim AplicarFiltro As MSHTML.IHTMLElement
Dim txtCNPJ As MSHTML.IHTMLElement
Dim CmdUF As MSHTML.IHTMLSelectElement
Dim IE As New SHDocVw.InternetExplorer
Dim HTMLDoc As New MSHTML.HTMLDocument
Dim Opcao As Variant
Dim i As Long
Dim StartTime As Date
Dim Timeout As Byte

        CNPJ = Util.FormatarCNPJ(CadContrib.Range("A4").value)
        UF = VBA.UCase(CadContrib.Range("C4").value)
        If CNPJ <> "" Then
            
            On Error GoTo Tratar:
            
            URL = "http://hnfe.sefaz.ba.gov.br/servicos/nfenc/Modulos/Geral/NFENC_consulta_cadastro_ccc.aspx"
                   
            Application.StatusBar = "Acessando o site..."
            Set IE = CreateObject("InternetExplorer.Application")
            IE.visible = False
            IE.Navigate URL
            
            Call IniciarContagem(StartTime, Timeout)
            Do While IE.Busy
                If IE.Busy = False Then Exit Do
                If DateDiff("s", StartTime, Now) > Timeout Then Call TempoLimiteExcedido: Exit Function
            Loop
        
            Set HTMLDoc = IE.Document
            
            Application.StatusBar = "Realizando consulta..."
            Call IniciarContagem(Now(), Timeout)
            Do
                Set txtCNPJ = HTMLDoc.getElementById("txtCNPJ")
                Set CmdUF = HTMLDoc.getElementById("CmdUF")
                Set AplicarFiltro = HTMLDoc.getElementById("AplicarFiltro")
                
                If Not txtCNPJ Is Nothing Then
                    
                    txtCNPJ.value = CNPJ
                    CmdUF.value = UF
                    
                    If Not AplicarFiltro Is Nothing Then AplicarFiltro.Click: Exit Do
                    
                End If
                
                If DateDiff("s", StartTime, Now) > Timeout Then Call TempoLimiteExcedido: Exit Function
                
            Loop
            
            Call IniciarContagem(Now(), Timeout)
            Do While IE.Busy
                If IE.Busy = False Then Exit Do
                If DateDiff("s", StartTime, Now) > Timeout Then Call TempoLimiteExcedido: Exit Function
            Loop
            
            Application.StatusBar = "Obtendo dados..."
            Call IniciarContagem(Now(), Timeout)
            Do
                'Application.Wait DateAdd("s", 1, Now) 'Para esperar 1 segundo antes de avançar a execução do código
                Set Tabelas = HTMLDoc.getElementsByTagName("tbody")
                If Not Tabelas Is Nothing And Tabelas.Length > 9 Then
                    Set Linhas = Tabelas(10).getElementsByTagName("tr")
                    If Not Linhas Is Nothing Then Exit Do
                End If
                
                If DateDiff("s", StartTime, Now) > Timeout Then Call TempoLimiteExcedido: Exit Function
                
            Loop
            
            Call IniciarContagem(Now(), Timeout)
            Do
                If Linhas.Length = 2 Then
                    
                    Set colunas = Linhas(1).getElementsByTagName("td")
                    If Not colunas Is Nothing And colunas(0).innerText <> "" Then Exit Do
                
                Else
                
                    MsgBox "O CNPJ informado não retornou dados!" & vbCrLf & vbCrLf & _
                           "Por favor verifique se o CNPJ informado está correto." & vbCrLf & vbCrLf & _
                           "Isso também pode ocorrer caso o CNPJ informado possua inscrição estadual em mais de um estado. " & _
                           "Se esse for o seu caso, além do CNPJ informe também a UF do estabelecimento.", vbCritical, "Dados não retornados"
                    Exit Function
                    
                End If
                
                If DateDiff("s", StartTime, Now) > Timeout Then Call TempoLimiteExcedido: Exit Function
                
            Loop
            
            InscEst = colunas(1).innerText
            Razao = colunas(2).innerText
            UF = colunas(3).innerText
            
            ActiveSheet.Range("B4:D4").value = Array("'" & InscEst, UF, Razao)
            IE.Quit
            
            Application.StatusBar = "Dados obtidos com sucesso!"
            
            MsgBox "Dados de cadastro importados com sucesso!", vbInformation, "Contribuinte Cadastrado"

        Else

            MsgBox "Informe CNPJ do Contribuinte que deseja trabalhar", vbExclamation, "CNPJ não informado"
            CadContrib.Activate

        End If
        
        Application.StatusBar = False
        
Exit Function
Tratar:

    IE.Quit
    If Err.Number <> 0 Then Call TratarErros(Err, "WebScraping.ExtrairCadastroContribuinteWeb")
    Application.StatusBar = False
     
End Function

Private Function TempoLimiteExcedido()

Dim Msg As String
    
    Msg = "O site excedeu o tempo limite de resposta e a operação foi suspensa." + vbCrLf + vbCrLf
    Msg = Msg & "Verifique a sua conexão com a internet e tente novamente." + vbCrLf
    Msg = Msg & "Caso a situação persista, o site pode estar fora do ar."
    
    Call Util.MsgAlerta(Msg, "Tempo Limite Excedido")
    Application.StatusBar = False
     
End Function

Private Function IniciarContagem(ByRef StartTime As Date, ByRef Timeout As Byte, Optional Limite As Byte)
    StartTime = Now()
    If Limite > 0 Then Timeout = Limite Else Timeout = 10
End Function

