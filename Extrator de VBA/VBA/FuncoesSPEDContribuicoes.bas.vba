Attribute VB_Name = "FuncoesSPEDContribuicoes"
Option Explicit
Option Base 1

Private EnumFiscal As New clsEnumeracoesSPEDFiscal

Public Function CarregarNiveisRegistros()

Dim Chaves As Variant, Chave
    
    Call dicChavesNivel.RemoveAll
    
    Chaves = Array("0", "1", "2", "3", "4", "5", "6")
    For Each Chave In Chaves
        dicChavesNivel.Add Chave, ""
    Next Chave
    
End Function

Public Function ImportarSPEDContribuicoes(Optional ByVal SelReg As String, Optional Periodo As String, Optional ByVal Unificar As Boolean)

Dim Arqs As Variant, Arq, Registros, Registro, Campos
Dim arrDados As New ArrayList
Dim nReg As String, versao$
Dim MsgEtapa As String
Dim Comeco2 As Double
Dim Comeco As Double
Dim MsgArq As String
Dim MsgReg As String
Dim Chave As Variant
Dim b As Long
    
    Arqs = Util.SelecionarArquivos("txt")
    If VarType(Arqs) = vbBoolean Then Exit Function
    
    Inicio = Now()
    Call Util.DesabilitarControles
    Call CarregarNiveisRegistros
    Call dicChavesRegistroSPED.RemoveAll
    
    Application.StatusBar = "Carregrando estrutura de dados do SPED Contribuições, por favor aguarde..."
    
    a = 0
    Comeco = Timer
    For Each Arq In Arqs
        
        Application.StatusBar = "Verificando o tipo do SPED importado."
        If fnSPED.ClassificarSPED(Arq, versao, Periodo) <> "Contribuições" Then
            
            Call Util.MsgAlerta("O SPED importado não é o Contribuições.", "Verificação Tipo SPED")
            Application.StatusBar = False
            Exit Function
            
        Else
            
            Call CarregarLayoutSPEDContribuicoes(versao)
            
            VersaoFiscal = EnumFiscal.ValidarEnumeracao_COD_VER(Periodo)
            Call CarregarLayoutSPEDFiscal(VersaoFiscal)
            
        End If
        
        Application.StatusBar = "Carregando dados do SPED para memória do computador, por favor aguarde..."
        Registros = Util.ImportarTxt(Arq)
        
        MsgArq = "Importando SPED Contribuições " & a + 1 & " de " & UBound(Arqs)
        Call Util.AntiTravamento(a, 1, MsgArq, UBound(Arqs) + 1, Comeco)
        
        b = 0
        Comeco2 = Timer
        For Each Registro In Registros
            
            Call Util.AntiTravamento(b, 100, MsgArq & " [Importando registros do bloco " & VBA.Left(nReg, 1) & "]", UBound(Registros) + 1, Comeco2)
            If VBA.Trim(Registro) <> "" And Registro Like "*|*" And VBA.Len(Registro) > 6 Then
                
                Campos = VBA.Split(Registro, "|")
                
                nReg = Campos(1)
                If dicLayoutContribuicoes.Exists(nReg & "_Contr") Then nReg = nReg & "_Contr"
                If dicLayoutContribuicoes.Exists(nReg & "_INI") Then nReg = nReg & "_INI"
                
                If SelReg <> "" Then
                    
                    Select Case True
                        
                        Case SelReg = nReg Or nReg = "0000_Contr" Or (VBA.Right(nReg, 3) = "001" And nReg <> "9001")
                            If Not dicRegistros.Exists(nReg) Then Set dicRegistros(nReg) = New ArrayList
                            Call fnSPED.ProcessarRegistro(Registro, Periodo, True)
                            
                    End Select
                    
                Else
                    
                    If Not dicRegistros.Exists(nReg) Then Set dicRegistros(nReg) = New ArrayList
                    If nReg = "0000_Contr" Then
                        
                        If Unificar Then Campos(9) = Util.GerarCNPJCompleto(VBA.Left(Campos(9), 8), "0001")
                        Registro = VBA.Join(Campos, "|")
                        
                    End If
                    
                    Call fnSPED.ProcessarRegistro(Registro, SPEDContr:=True, Unificar:=Unificar)
                    If nReg = "9001" Then Exit For
                    
                End If
                
            End If
            
            If SelReg <> "" Then If VBA.Left(nReg, 2) > VBA.Left(SelReg, 2) Then Exit For
            
        Next Registro
        
    Next Arq
    
    b = 0
    For Each Chave In dicRegistros.Keys()
        Call Util.AntiTravamento(b, 10, MsgArq & " [Exportando dados do bloco " & VBA.Left(Chave, 1) & " para o Excel]")
        Set arrDados = dicRegistros(Chave)
        On Error Resume Next
            Call Util.ExportarDadosArrayList(Worksheets(Chave), arrDados)
        On Error GoTo 0
    Next Chave
    
    Call dicRegistros.RemoveAll
    
    Application.StatusBar = "Exportação finalizada com sucesso!"
    Call Util.MsgInformativa("SPED Contribuições importado com sucesso!", "Importação do SPED Contribuições", Inicio)
    
    Application.StatusBar = False
            
End Function

Private Function ProcessarCampos(ByVal Registro As String)

Dim Campos() As Variant
Dim Dados As Variant
Dim i As Byte
    
    Registro = VBA.Trim(VBA.Replace(VBA.Replace(Registro, vbCr, ""), vbLf, ""))
    Registro = VBA.Mid(Registro, 2, Len(Registro) - 2)
    ProcessarCampos = VBA.Split(Registro, "|")
    
End Function

Public Function ResetarDicionarios()
    
    Call dicHierarquiaSPEDFiscal.RemoveAll
    Call dicMapaChavesSPEDFiscal.RemoveAll
    Call dicChavesNivel.RemoveAll
    Call dicRegistros.RemoveAll
    Call ListaChaves.RemoveAll
    'Call dicTitulos.RemoveAll
    Call dicFilhos.RemoveAll
    Call dicNomes.RemoveAll
    Call dicPais.RemoveAll
    
End Function

Public Sub GerarEFDContribuicoes()

Dim dicEstruturaSPED As New Dictionary
Dim dicRegistros As New Dictionary
Dim Msg As String, Caminho$, Arq$, MsgEtapa$, versao$
Dim dicTitulos As New Dictionary
Dim Registros As New Dictionary
Dim Registro As Variant, Campos
Dim dtRef As String, dtFin$
Dim EFD As New ArrayList
Dim UltLin As Long
    
    UltLin = Util.UltimaLinha(reg0000_Contr, "A")
    If UltLin < 4 Then
        Call Util.MsgAlerta("Sem dados do SPED Contribuições para exportar.", "Exportação do SPED Contribuições")
        Exit Sub
    End If
    
    Caminho = Util.SelecionarPasta("Selecione a pasta para Salvar os arquivos")
    If Caminho = "" Then Exit Sub
    
    Inicio = Now()
    Call Util.DesabilitarControles
                
        Set dicTitulos = Util.MapearTitulos(reg0000_Contr, 3)
        Set Registros = Util.CriarDicionarioRegistro(reg0000_Contr)
                
        a = 0
        Comeco = Timer
        For Each Campos In Registros.Items
            
            versao = Campos(dicTitulos("COD_VER"))
            
            Application.StatusBar = "Etapa 1/5 - Carregando estrutura de dados do SPED Contribuições, por favor aguarde..."
            Call FuncoesJson.CarregarEstruturaSPEDContribuicoes(dicEstruturaSPED, versao)
            
            Application.StatusBar = "Etapa 2/5 - Carregando registros do SPED Contribuições, por favor aguarde..."
            Set dicRegistros = Util.CarregarRegistrosSPED(dicEstruturaSPED, True)
            Call Util.TratarParticularidadesEFDContribuicoes(dicRegistros)
                        
            MsgEtapa = "Etapa 3/5 - "
            Call Util.AntiTravamento(a, 1, MsgEtapa & "Exportando arquivo " & a + 1 & " de " & Registros.Count, Registros.Count, Comeco)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                With Campos0000_Contr
                                        
                    .CHV_REG = Campos(dicTitulos("CHV_REG"))
                    .IND_ATIV = Campos(dicTitulos("IND_ATIV"))
                    .ARQUIVO = Campos(dicTitulos("ARQUIVO"))
                    .CNPJ = Campos(dicTitulos("CNPJ"))
                    .TIPO_ESCRIT = VBA.Left(VBA.Replace(Campos(dicTitulos("TIPO_ESCRIT")), "'", ""), 1)
                    .DT_INI = VBA.Format(Campos(dicTitulos("DT_INI")), "yyyymm")
                    dtFin = VBA.Format(Campos(dicTitulos("DT_FIN")), "yyyy-mm-dd")
                    If .TIPO_ESCRIT = 0 Then .TIPO_ESCRIT = " Remessa do arquivo original " Else .TIPO_ESCRIT = " Remessa do arquivo substituto "
                    
                    dtRef = VBA.Format(.DT_INI & "01", "0000-00-00")
                    Arq = Caminho & "\EFD Contribuicoes " & .CNPJ & .TIPO_ESCRIT & .DT_INI & " - ControlDocs.txt"
                    
                    Call fnSPED.IncluirRegistro(EFD, dicEstruturaSPED, dicRegistros, "0000_Contr", .CHV_REG, True)
                    Call FuncoesAjusteSPED.RemoverCadastrosNaoReferenciados(Arq, EFD, True)
                    
                End With
                
            End If
            
            EFD.Clear
            
        Next Campos
        
        If dicRegistros("0000_Contr").Count = 1 Then Msg = "Arquivo gerado com sucesso!" Else Msg = "Arquivos gerados com sucesso!"
        Call Util.MsgInformativa(Msg, "Geração da EFD Contribuições", Inicio)
        
    Call Util.HabilitarControles
    Application.StatusBar = False
    
End Sub
