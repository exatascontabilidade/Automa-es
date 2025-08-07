Attribute VB_Name = "Rotinas"
Option Explicit
Option Base 1

Public Function AssinaturaControlDocs(control As IRibbonControl)
                
Dim URL As String
                
    Select Case control.id
        
'        Case "btnBasicoMensal"
'            Url = urlControlDocsBasicoMensal
'
'        Case "btnBasicoSemestral"
'            Url = urlControlDocsBasicoSemestral
'
'        Case "btnBasicoAnual"
'            Url = urlControlDocsBasicoAnual
'
'        Case "btnPlusMensal"
'            Url = urlControlDocsPlusMensal
'
'        Case "btnPlusSemestral"
'            Url = urlControlDocsPlusSemestral
'
'        Case "btnPlusAnual"
'            Url = urlControlDocsPlusAnual
'
'        Case "btnPremiumMensal"
'            Url = urlControlDocsPremiumMensal
'
'        Case "btnPremiumSemestral"
'            Url = urlControlDocsPremiumSemestral
'
'        Case "btnPremiumAnual"
'            Url = urlControlDocsPremiumAnual
        
        Case "btnIndividualMensal"
            URL = urlAssinaturaIndividualMensal
        
        Case "btnIndividualAnual"
            URL = urlAssinaturaIndividualAnual
        
        Case "btnEmpresarialMensal"
            URL = urlAssinaturaEmpresarialMensal
        
        Case "btnEmpresarialAnual"
            URL = urlAssinaturaEmpresarialAnual
        
    End Select
                
    Call FuncoesLinks.AbrirUrl(URL)

End Function

Public Sub CruzarDadosComSPED()

Dim Arqs As Variant, Arq As Variant
Dim UltLin As Long, i As Long
Dim Registro As Variant
Dim Dados As Variant
Dim chNFe As String, StatusDoc As String
Dim dicEntradasNFe As New Dictionary
Dim dicSaidasNFe As New Dictionary
Dim dicSaidasNFCe As New Dictionary
Dim dicEntradasCTes As New Dictionary
Dim dicSaidasCTes As New Dictionary
Dim dicSaidasCFes As New Dictionary
Dim DicXMLsFaltantes As New Dictionary
Dim arrCanceladas As New ArrayList
Dim StatusSPED As Boolean
    
    Arqs = Util.SelecionarArquivos("txt")
    
    If VarType(Arqs) <> 11 Then
        
        Inicio = Now
        
        Call Util.LimparDados(XMLSFaltantes, 4, False)
        
        Application.StatusBar = "Carregando dados das NFe de entrada, por favor aguarde..."
        Set dicEntradasNFe = Util.CriarDicionarioRegistro(EntNFe, "Chave de Acesso")
        
        Application.StatusBar = "Carregando dados dos CTe de entrada, por favor aguarde..."
        Set dicEntradasCTes = Util.CriarDicionarioRegistro(EntCTe, "Chave de Acesso")
        
        Application.StatusBar = "Carregando dados das NFe de saída, por favor aguarde..."
        Set dicSaidasNFe = Util.CriarDicionarioRegistro(SaiNFe, "Chave de Acesso")
        
        Application.StatusBar = "Carregando dados das NFCe, por favor aguarde..."
        Set dicSaidasNFCe = Util.CriarDicionarioRegistro(SaiNFCe, "Chave de Acesso")
        
        Application.StatusBar = "Carregando dados dos CFe de saída, por favor aguarde..."
        Set dicSaidasCFes = Util.CriarDicionarioRegistro(SaiCFe, "Chave de Acesso")
        
        Application.StatusBar = "Carregando dados dos CTe de saída, por favor aguarde..."
        Set dicSaidasCTes = Util.CriarDicionarioRegistro(SaiCTe, "Chave de Acesso")
        
        a = 0
        Comeco = Timer
        Application.StatusBar = "Identificando omissões de documentos fiscais nos SPEDs, por favor aguarde..."
        For Each Arq In Arqs
            Call Util.AntiTravamento(a, 1, "Identificando omissões de documentos fiscais nos SPEDs, por favor aguarde...", UBound(Arqs) + 1, Comeco)
            StatusSPED = VerificarStatusSPED(Arq, dicEntradasNFe, dicSaidasNFe, dicSaidasNFCe, dicEntradasCTes, dicSaidasCTes, dicSaidasCFes, DicXMLsFaltantes)
        Next Arq
        
        Application.StatusBar = "Exportando resultados das análises, por favor aguarde..."
        Call Util.ExportarDadosDicionario(EntNFe, dicEntradasNFe, "A4")
        Call Util.ExportarDadosDicionario(SaiNFe, dicSaidasNFe, "A4")
        Call Util.ExportarDadosDicionario(SaiNFCe, dicSaidasNFCe, "A4")
        Call Util.ExportarDadosDicionario(EntCTe, dicEntradasCTes, "A4")
        Call Util.ExportarDadosDicionario(SaiCTe, dicSaidasCTes, "A4")
        Call Util.ExportarDadosDicionario(SaiCFe, dicSaidasCFes, "A4")
        Call Util.ExportarDadosDicionario(XMLSFaltantes, DicXMLsFaltantes, "A4")
        
        Call ZerarDicionarios(dicEntradasNFe, dicSaidasNFe, dicSaidasNFCe, dicEntradasCTes, dicSaidasCTes)
        
        MsgBox "Cruzamento de dados concluído com sucesso." & vbCrLf & vbCrLf & "Tempo decorrido: " & VBA.Format(Now - Inicio, "ttttt"), vbInformation, "Cruzamento de dados com o SPED Fiscal"
        
    End If
    
    Application.StatusBar = False
    
End Sub

Public Sub ImportarDadosSEFAZBA()

Dim dicEntradas As New Dictionary
Dim dicSaidas As New Dictionary
Dim arrChaves As New ArrayList
Dim Arqs
    
    Arqs = Util.SelecionarArquivos("CSV")
    
    If VarType(Arqs) <> 11 Then
        
        Inicio = Now
        Call CarregarDadosContribuinte
        Call Util.CarregarChavesAcessoDoces(arrChaves)
        
        Call fnCSV.ImportarCSV(Arqs, arrChaves, dicEntradas, dicSaidas)
        If dicEntradas.Count > 0 Then Call Util.ExportarDadosDicionario(EntNFe, dicEntradas)
        If dicSaidas.Count > 0 Then Call Util.ExportarDadosDicionario(SaiNFe, dicSaidas)
        
        MsgBox "Importação concluída" & vbCrLf & vbCrLf & "Tempo decorrido: " & VBA.Format(Now - Inicio, "ttttt"), vbInformation, "Importação de dados da SEFAZ/BA"
        
    End If
    
End Sub

Public Function VerificarStatusSPED(ByVal Arq As String, ByRef dicEntradasNFe As Dictionary, ByRef dicSaidasNFe As Dictionary, _
                                    ByRef dicSaidasNFCe As Dictionary, ByRef dicEntradasCTe As Dictionary, ByRef dicSaidasCTe As Dictionary, _
                                    ByRef dicSaidasCFe As Dictionary, ByRef DicionarioPendencias As Dictionary) As Boolean

Dim Modelo As String, tpOperacao$, Situacao$, nReg$, Data$
Dim EFD As Variant
Dim Campos As Variant, Registro As Variant
Dim DataSPED As String
Dim b As Long
    
    Application.StatusBar = "Carregando dados do SPED para memória do computador, esse processo travar o Excel por alguns minutos, mas o processo continua acontecendo"
    EFD = Util.ImportarTxt(Arq)
    For Each Registro In EFD
        
        Call Util.AntiTravamento(b, 50, "Processando registros do bloco " & VBA.Mid(Registro, 2, 1) & " por favor aguarde", UBound(EFD) + 1, Comeco)
        If Registro <> "" Then
            
            Campos = Split(Registro, "|")
            Select Case Campos(1)
                
                Case "0000"
                    DataSPED = Campos(4)
                
                Case "C100", "D100"

                    nReg = Campos(1)
                    tpOperacao = Campos(2)
                    Modelo = Campos(5)
                    Situacao = Campos(6)
                    
                    Select Case tpOperacao
                        
                        Case "0"
                            If Modelo = "55" Then Call AdicionarDocumentoDicionario(Registro, dicEntradasNFe, DicionarioPendencias, DataSPED)
                            If Modelo = "57" Then Call AdicionarDocumentoDicionario(Registro, dicEntradasCTe, DicionarioPendencias, DataSPED)
                            
                        Case "1"
                            If Modelo = "55" Then Call AdicionarDocumentoDicionario(Registro, dicSaidasNFe, DicionarioPendencias, DataSPED)
                            If Modelo = "65" Then Call AdicionarDocumentoDicionario(Registro, dicSaidasNFCe, DicionarioPendencias, DataSPED)
                            If Modelo = "57" Then Call AdicionarDocumentoDicionario(Registro, dicSaidasCTe, DicionarioPendencias, DataSPED)
                            
                    End Select
                
                Case "C800"
                    Call AdicionarDocumentoDicionario(Registro, dicSaidasCFe, DicionarioPendencias, DataSPED)
                    
                Case "D990"
                    Exit For
                        
            End Select
            
        End If
        
    Next Registro
    
End Function

Private Sub ZerarDicionarios(ParamArray dicionarios() As Variant)

Dim Dicionario As Variant

    For Each Dicionario In dicionarios
        Dicionario.RemoveAll
    Next Dicionario
    
End Sub

Public Sub ImportarCTes()

Dim dicEntradasCTes As New Dictionary
    
    Call CarregarDadosDocumentos(EntCTe, dicEntradasCTes)
    Call fnXML.ImportarCTe(Util.SelecionarArquivos("xml"), dicEntradasCTes)
    Call Util.ExportarDadosDicionario(EntCTe, dicEntradasCTes, "A4", 12)
    EntCTe.Activate
    
End Sub

Public Function VerificarDataSPED(ByVal DataSPED As String, chNFe As String, ByRef Dicionario As Dictionary)

    If DataSPED <> "" Then
        VerificarDataSPED = Util.FormatarData(VBA.Format(DataSPED, "00/00/0000"))
    Else
        If Dicionario.Exists(chNFe) Then VerificarDataSPED = Dicionario(chNFe)(9)
    End If
    
End Function

Public Function ValidarvNF(ByVal vNFSPED As String, chNFe As String, ByRef Dicionario As Dictionary)
    
    If vNFSPED <> "" Then
        If Round(vNFSPED - fnExcel.ConverterValores(Dicionario(chNFe)(5)), 2) <> 0 Then
            ValidarvNF = Round(vNFSPED - Dicionario(chNFe)(5), 2)
        Else
            ValidarvNF = "OK"
        End If
    End If
    
End Function

Private Sub AdicionarDocumentoDicionario(ByVal Registro As String, ByRef Dicionario As Dictionary, _
    ByRef DicXMLsFaltantes As Dictionary, ByVal DataSPED As String)
    
Dim Campos As Variant, dicCampos
Dim dtEnt As String, dtEmi$, nReg$, chDoc$, tpOperacao$, tpEmissao$, Situacao$, Modelo$, SERIE$, nNF$, indOper$
Dim vNF As Double
    
    Campos = VBA.Split(Registro, "|")
    
    nReg = Campos(1)
    Situacao = Campos(6)
    If nReg = "C100" Then
        
        indOper = Campos(2)
        nNF = "'" & Campos(8)
        dtEmi = Util.FormatarData(Campos(10))
        dtEnt = Util.FormatarData(Campos(11))
        vNF = Util.ValidarValores(Campos(12))
        chDoc = Campos(9)
        Campos(9) = "'" & Campos(9)
                
    ElseIf nReg = "C800" Then
        
        Situacao = Campos(3)
        nNF = "'" & Campos(4)
        dtEnt = Util.FormatarData(Campos(5))
        dtEmi = Util.FormatarData(Campos(5))
        vNF = Util.ValidarValores(Campos(6))
        chDoc = Campos(11)
        Campos(11) = "'" & Campos(11)
        
    ElseIf nReg = "D100" Then
        
        indOper = Campos(2)
        nNF = "'" & Campos(9)
        dtEmi = Util.FormatarData(Campos(11))
        dtEnt = Util.FormatarData(Campos(12))
        vNF = Util.ValidarValores(Campos(15))
        chDoc = Campos(10)
        Campos(10) = "'" & Campos(10)
        
    End If
    
    If Dicionario.Exists(chDoc) Then
        
        If indOper <> "" Then If indOper = 1 Then dtEnt = dtEmi
        dicCampos = Dicionario(chDoc)
        dicCampos(10) = dtEnt
        dicCampos(11) = ValidarvNF(vNF, chDoc, Dicionario)
        
        If Situacao Like "02*" Or Situacao Like "03*" Or Situacao Like "04*" Then _
            dicCampos(10) = VBA.Format(VBA.Format(DataSPED, "00/00/0000"), "yyyy-mm-dd")
            
        Dicionario(dicCampos(6)) = dicCampos
        If DicXMLsFaltantes.Exists(chDoc) Then Call DicXMLsFaltantes.Remove(chDoc)
        
    Else
        
        Modelo = Campos(5)
        SERIE = "'" & Campos(7)
        tpOperacao = Util.ValidarOperacao(Campos(2))
        tpEmissao = Util.ValidarEmissao(Campos(3))
        Situacao = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_SIT(Campos(6))
        
        If nReg = "C800" Then
            
            Modelo = Campos(2)
            SERIE = "'" & Campos(10)
            tpOperacao = Util.ValidarOperacao("1")
            tpEmissao = Util.ValidarEmissao("0")
            Situacao = ValidacoesSPED.Fiscal.Enumeracoes.ValidarEnumeracao_COD_SIT(Campos(3))
            
        End If
        
        DicXMLsFaltantes(chDoc) = Array(Modelo, SERIE, nNF, dtEmi, dtEnt, "'" & chDoc, CDbl(vNF), Situacao, tpOperacao, tpEmissao)
        
    End If
    
End Sub

Public Sub ListarXMLsAusentes()
    
    Select Case ActiveSheet.CodeName
        
        Case "relInteligenteDivergencias"
            ActiveSheet.Range("A3:AB" & Rows.Count).AutoFilter Field:=27, Criteria1:="XML AUSENTE"
    
    End Select
    
End Sub

Public Sub ListarEscrituracoesDivergentes()
    
    Select Case ActiveSheet.CodeName
        
        Case "Divergencias"
            ActiveSheet.Range("A3:AB" & Rows.Count).AutoFilter Field:=27, Criteria1:="DIVERGÊNCIA"
    
    End Select
    
End Sub

Public Sub ClassificarCFOPs()

    On Error Resume Next
    
        With ActiveSheet.AutoFilter.Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=Range("A2:A" & Rows.Count)
            .Apply
        End With

End Sub

Public Function AtualizarDadosXML(ByVal vbResult As VbMsgBoxResult) As Long
    If vbResult = vbYes Then Call Util.LimparDados(relInteligenteDivergencias, 4, True)
End Function

'IDENTIFICAR FUROS DE NUMERAÇÃO NAS NOTAS FISCAIS
Public Sub ListarFurosNumeracao()

Dim dicNotas As New Dictionary
Dim Nota, Notas, item, Chaves
Dim Furos As New ArrayList
Dim nMin As Long, nMax As Long, i As Long, UltLin As Long
    
    Inicio = Now
    
    UltLin = SaiNFe.Range("F" & Rows.Count).END(xlUp).Row
    
    If UltLin > 3 Then

        With DadosDoce
        
            Chaves = SaiNFe.Range("F4:F" & UltLin)
            For i = LBound(Chaves) To UBound(Chaves)
            
                .Modelo = VBA.Mid(Chaves(i, 1), 21, 2)
                .SERIE = VBA.Mid(Chaves(i, 1), 23, 3)
                .nNF = VBA.Mid(Chaves(i, 1), 26, 9)
                .Chave = .Modelo & .SERIE
                
                If Not dicNotas.Exists(.Chave) Then Set dicNotas(.Chave) = CreateObject("System.Collections.ArrayList")
                dicNotas(.Chave).Add CLng(.nNF)
                
            Next i
            
        End With
            
        For Each item In dicNotas.Keys()
            
            Notas = dicNotas(item).toArray
            nMin = WorksheetFunction.Min(Notas)
            nMax = WorksheetFunction.Max(Notas)
            
            For i = nMin To nMax
                If dicNotas(item).contains(i) = False Then Furos.Add i
            Next i
            
        Next item
        
        With Application
            If Furos.Count > 0 Then
                relInteligenteDivergencias.Range("T4").Resize(Furos.Count, 1).value = .Transpose(Furos.toArray)
                MsgBox "Processamento concluído!" & vbCrLf & "Tempo decorrido: " & VBA.Format(Now() - Inicio, "ttttt"), vbInformation, "Furos de Numeração"
            Else
                MsgBox "Não foram encontrados furos na numeração!", vbInformation, "Furos de Numeração"
            End If
        End With
        
    End If
    
    
End Sub

Public Function ListarArquivos(ByRef ListaXMLS As ArrayList) As ArrayList

Dim Caminho As String, Arq
Dim fso As New FileSystemObject
Dim pasta As Folder
Dim ARQUIVO As file
    
    Caminho = Util.SelecionarPasta()
    If Caminho <> "" Then
        
        Inicio = Now()
        
        Set pasta = fso.GetFolder(Caminho)
        Application.StatusBar = "Listando os arquivos do diretório, por favor aguarde."
        
        For Each ARQUIVO In pasta.Files
            If InStr(1, VBA.LCase(ARQUIVO.Path), ".xml") Then ListaXMLS.Add ARQUIVO.Path
        Next ARQUIVO
        
        Call Funcoes.IndentificarSubPastas(Caminho, ListaXMLS)
        
    End If
    
    Set ListarArquivos = ListaXMLS
    
End Function

Public Sub LimparDados(control As IRibbonControl)
    Call Util.LimparDados(ActiveSheet, 4, True)
End Sub

Public Sub LimparFiltros(control As IRibbonControl)
    Call Util.LimparFiltros(ActiveSheet)
End Sub
