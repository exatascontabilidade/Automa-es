Attribute VB_Name = "clsFuncoesSPED"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Fiscal As New clsFuncoesSPEDFiscal
Public dicChavesNivelFiscal As New Dictionary
Private GerenciadorSPED As New clsRegistrosSPED
Public dicChavesNivelContribuicoes As New Dictionary
Public Contribuicoes As New clsFuncoesSPEDContribuicoes
Private EnumContrib As New clsEnumeracoesSPEDContribuicoes

Public Sub ExportarSPED(ByVal Arq As String, EFD As ArrayList)

Dim TXT As String
    
    TXT = VBA.Join(EFD.toArray(), vbCrLf)
    Open Arq For Output As #1
        Print #1, TXT
    Close #1
    
End Sub

Public Function TotalizarRegistrosSPED(ByRef EFD As ArrayList, Optional ByVal SPEDContr As Boolean) As ArrayList

Dim dicRegistros As New Dictionary
Dim dicBlocos As New Dictionary
Dim totReg As Long, totBloco&
Dim nReg As String, nBloco$
Dim Registro As Variant
    
    For Each Registro In EFD
        
        If Registro <> "" Then
            
            nReg = VBA.Mid(Registro, 2, 4)
            nBloco = VBA.Left(nReg, 1)
            If dicRegistros.Exists(nReg) Then totReg = dicRegistros(nReg)
            If dicBlocos.Exists(nBloco) Then totBloco = dicBlocos(nBloco)
            
            dicRegistros(nReg) = totReg + 1
            dicBlocos(nBloco) = totBloco + 1
            totReg = 0: totBloco = 0
            
            If nReg = "9001" Then Exit For
            
        End If
        
    Next Registro
    
    Set TotalizarRegistrosSPED = TotalizarBlocos(EFD, dicRegistros, dicBlocos, SPEDContr)
    
End Function

Public Function TotalizarBlocos(ByRef EFD As ArrayList, ByRef dicRegistros As Dictionary, _
    ByRef dicBlocos As Dictionary, Optional ByVal SPEDContr As Boolean) As ArrayList
    
Dim Registro As Variant, REG, Campos
Dim nReg As String, nBloco$, COD_VER$
Dim nEFD As New ArrayList

    For Each Registro In EFD
        
        If Registro <> "" Then
            
            nReg = VBA.Mid(Registro, 2, 4)
            Select Case True
                
                Case nReg Like "0000"
                    COD_VER = fnSPED.ExtrairCampo(Registro, 2)
                    nEFD.Add Registro
                    
                Case nReg Like "*001" And Not nReg Like "E001" And Not nReg Like "1001" And Not nReg Like "9001"
                    nBloco = VBA.Left(nReg, 1)
                    Campos = VBA.Split(Registro, "|")
                    If dicBlocos(nBloco) > 2 Then Campos(2) = 0 Else Campos(2) = 1
                    
                    Registro = VBA.Join(Campos, "|")
                    nEFD.Add Registro
                    
                Case nReg Like "*990"
                    
                    nBloco = VBA.Left(nReg, 1)
                    Campos = VBA.Split(Registro, "|")
                        
                        Campos(2) = dicBlocos(nBloco)
                        
                    Registro = VBA.Join(Campos, "|")
                    nEFD.Add Registro
                    
                Case nReg = "E001" And Not SPEDContr
                    
                    nEFD.Add Registro
                    If Not dicRegistros.Exists("E100") Then
                        Call GerarRegistroE100eFilhos(dicRegistros, nEFD(0), nEFD)
                        If Not dicRegistros.Exists("E100") Then
                            dicRegistros("E100") = 1
                            dicRegistros("E110") = 1
                            dicBlocos("E") = CInt(dicBlocos("E")) + 2
                        End If
                    End If
                    
                Case nReg = "1001" And Not SPEDContr
                                         
                    If COD_VER > "005" Then
                        
                        nEFD.Add "|1001|0|"
                        nEFD.Add GerarRegistro1010(dicRegistros, COD_VER)
                        If Not dicRegistros.Exists("1010") Then
                            dicRegistros("1010") = 1
                            dicBlocos("1") = CInt(dicBlocos("1")) + 1
                        End If
                        
                    Else
                        
                        nEFD.Add "|1001|1|"
                        
                    End If
                    
                Case nReg = "1010" And Not SPEDContr
                    
                Case nReg = "9001"
                    
                    nEFD.Add Registro
                    For Each REG In dicRegistros.Keys()
                        If REG <> "" Then nEFD.Add "|9900|" & REG & "|" & dicRegistros(REG) & "|"
                    Next REG
                    
                    nEFD.Add "|9900|9900|" & dicRegistros.Count + 3 & "|"
                    nEFD.Add "|9900|9990|1|"
                    nEFD.Add "|9900|9999|1|"
                    nEFD.Add "|9990|" & dicRegistros.Count + 6 & "|"
                    nEFD.Add "|9999|" & nEFD.Count + 1 & "|"
                    Exit For
                    
                Case Else
                    If Registro <> "" Then nEFD.Add Registro
                    
            End Select
        
        End If
        
    Next Registro
    
    Set TotalizarBlocos = nEFD
    
End Function

Public Function GerarRegistroE100eFilhos(ByRef dicRegistros As Dictionary, ByVal reg0000 As String, ByRef EFD As ArrayList)
    
Dim DT_INI As String
Dim DT_FIM As String
Dim Campos As Variant
    
    Campos = VBA.Split(reg0000, "|")
    
    DT_INI = Campos(4)
    DT_FIM = Campos(5)
    
    EFD.Add "|E100|" & DT_INI & "|" & DT_FIM & "|"
    EFD.Add "|E110|0|0|0|0|0|0|0|0|0|0|0|0|0|0|"
    
End Function

Public Function GerarRegistro1010(ByRef dicRegistros As Dictionary, ByVal COD_VER As String)

Dim Registros As Variant, Registro
Dim arrCampos As New ArrayList
    
    arrCampos.Add ""
    arrCampos.Add "1010"
    
    Select Case True
    
        Case COD_VER > "013"
            Registros = Array("1100", "1200", "1300", "1390", "1400", "1500", "1601", "1700", "1800", "1960", "1970", "1980", "1250")
            
        Case COD_VER > "012"
            Registros = Array("1100", "1200", "1300", "1390", "1400", "1500", "1601", "1700", "1800", "1960", "1970", "1980")
            
        Case Else
            Registros = Array("1100", "1200", "1300", "1390", "1400", "1500", "1601", "1700", "1800")
            
    End Select
    
    For Each Registro In Registros
        If dicRegistros.Exists(Registro) Then arrCampos.Add "S" Else arrCampos.Add "N"
    Next Registro
    
    arrCampos.Add ""
    
    GerarRegistro1010 = VBA.Join(arrCampos.toArray, "|")
    
End Function

Private Function AdicionarRegistro(ByRef EFD As Variant, ByRef Registro As Variant)
    If EFD = "" Then EFD = Registro Else EFD = EFD & vbCrLf & Registro
End Function

Public Function GerarChvReg0000(ByVal Periodo As String) As String

Dim dtIni As String, dtFim$

    CNPJContribuinte = Util.FormatarCNPJ(CadContrib.Range("CNPJContribuinte").value)
    CNPJBase = VBA.Left(CNPJContribuinte, 8)
    InscContribuinte = fnExcel.FormatarValores(Util.ApenasNumeros(CadContrib.Range("InscContribuinte").value)) * 1
    
    dtIni = Util.FormatarData("01" & VBA.Replace(Periodo, "/", ""))
    dtFim = VBA.Format(Util.FimMes(dtIni), "ddmmyyyy")
    dtIni = VBA.Format(dtIni, "ddmmyyyy")
    GerarChvReg0000 = Cripto.MD5(Util.UnirCampos(dtIni, dtFim, CNPJContribuinte, InscContribuinte))
    'GerarChvReg0000 = Cripto.MD5(fnSPED.MontarChaveRegistro(Array("", dtIni, dtFim, CNPJContribuinte, "", InscContribuinte)))

End Function

Public Function IncluirRegistro(ByRef EFD As ArrayList, ByRef dicEstruturaSPED As Dictionary, _
    ByRef dicRegistros As Dictionary, ByVal nReg As String, ByVal Chave As String, Optional SPEDContr As Boolean = False)
    
Dim Registro As Variant, Dados, nCampos, Campos, Campo, REG, regFilho, Titulos
Dim CHV_REG, CHV_PAI
Dim MsgEtapa As String
Dim titCampos As New ArrayList
Dim arrCampos As New ArrayList
Dim dicCampos As New Dictionary
Dim dicRegistro As New Dictionary
Dim dicTitulos As New Dictionary
Dim arrTitulos As New ArrayList
Dim i As Long, UltLin As Long
Dim Intervalo As Range
Dim teste As Variant
    
    MsgEtapa = "Etapa 4/5 - "
    Call Util.AntiTravamento(a, 1, MsgEtapa & "Incluindo o registro " & nReg, dicRegistros(nReg).Count, Comeco)
    
    Chave = VBA.Replace(Chave, "'", "")
    
    If dicRegistros.Exists(nReg) Then
        
        'Definindo dicionário com os títulos das colunas da planilha de dados
        Set dicTitulos = Util.MapearTitulos(Worksheets(nReg), 3)
        Set arrTitulos = Util.TransformarDicionarioArrayList(dicTitulos)
        
        'Definindo a lista de campos que vão para o SPED
        Set titCampos = dicEstruturaSPED(nReg)("NomeCampos")
        
        'Selecionando o registro para pegar os dados para exportação
        Set dicRegistro = dicRegistros(nReg)
        
        'Criando ArrayList com todos os dados de exportação de uma chave específica
        Set dicCampos = dicRegistro(Chave)
        
        'Exit Function
        For Each Campos In dicCampos.Items
            
            'Coletando CHV_REG e CHV_PAI de cada registro
            CHV_REG = Campos(dicTitulos("CHV_REG"))
            CHV_PAI = Campos(dicTitulos(IIf(SPEDContr = True, "CHV_PAI_CONTRIBUICOES", "CHV_PAI_FISCAL")))
                
            arrCampos.Add ""
            For i = LBound(Campos) To UBound(Campos)

                'Verificando se o campo em questão faz parte do SPED e o adiciona em caso afirmativo
                If titCampos.contains(arrTitulos.item(i - 1)) Then arrCampos.Add Campos(i)
                
            Next i
            arrCampos.Add ""
            
            'Inserindo registro montado no SPED
            EFD.Add GerarRegistro(arrCampos.toArray)
            
            'Limpando dados do registro montado
            arrCampos.Clear
            On Error Resume Next
            'Identificando registros filhos do registro Atual
            If dicEstruturaSPED.Exists(nReg) Then
                If dicEstruturaSPED(nReg).Exists("Filhos") Then
                    For Each regFilho In dicEstruturaSPED(nReg)("Filhos")
                        'Chama a rotina para incluir os dados do registro filho caso ele exista
                        If dicRegistros.Exists(regFilho) Then
                            If dicRegistros.Exists(regFilho) Then If dicRegistros(regFilho).Exists(CHV_REG) Then Call IncluirRegistro(EFD, dicEstruturaSPED, dicRegistros, regFilho, CHV_REG, SPEDContr)
                        End If
                    Next regFilho
                End If
            End If
            
        Next Campos
        
    End If
    
End Function

Public Function ProcessarRegistro(ByVal Registro As String, Optional Periodo As String, Optional SPEDContr As Boolean, Optional Unificar As Boolean)

Dim arrInconsistencias As New ArrayList
Dim dicTitulos As New Dictionary
Dim Campos As Variant, Titulo
Dim dicReg As New Dictionary
Dim dicRegFiscal As Dictionary
Dim dicRegContribuicoes As Dictionary
Dim arrConversoes As New ArrayList
Dim nCampos As New ArrayList
Dim nReg As String, nRegContrib$, CHV_PAI$, CHV_REG$, CHV_PAI_FISCAL$, CHV_PAI_CONTRIBUICOES$
Dim Posicao As Long
    
    arrConversoes.Add "0001"
    arrConversoes.Add "0150"
    arrConversoes.Add "C100"
    arrConversoes.Add "D100"
    arrConversoes.Add "9001"
    
    Campos = ProcessarCampos(Registro)
    nReg = Campos(0)
    
    If SPEDContr Then
        
        If dicLayoutContribuicoes.Exists(nReg & "_Contr") Then nReg = nReg & "_Contr"
        If dicLayoutContribuicoes.Exists(nReg & "_INI") Then nReg = nReg & "_INI"
        If Not dicLayoutContribuicoes.Exists(nReg) Then Exit Function
        Set dicReg = dicLayoutContribuicoes(nReg)
        
    Else
        
        If Not dicLayoutFiscal.Exists(nReg) Then Exit Function
        Set dicReg = dicLayoutFiscal(nReg)
        
    End If
    
    If dicLayoutFiscal.Exists(nReg) Then Set dicRegFiscal = dicLayoutFiscal(nReg)
    
    nRegContrib = nReg
    If dicLayoutContribuicoes.Exists(nRegContrib & "_Contr") Then nRegContrib = nRegContrib & "_Contr"
    If dicLayoutContribuicoes.Exists(nRegContrib & "_INI") Then nRegContrib = nRegContrib & "_INI"
    If dicLayoutContribuicoes.Exists(nRegContrib) Then Set dicRegContribuicoes = dicLayoutContribuicoes(nRegContrib)
    
    Call ListarChavesRegistrosSPED(nReg)
    
    Call nCampos.Clear
    Set nCampos = dicReg("NomeCampos")
    
    'Insere os campos chave do ControlDocs
    If nCampos(1) <> "ARQUIVO" Then
        
        nCampos.Insert 1, "CHV_PAI_CONTRIBUICOES"
        nCampos.Insert 1, "CHV_PAI_FISCAL"
        nCampos.Insert 1, "CHV_REG"
        nCampos.Insert 1, "ARQUIVO"
        
    End If
    
    Set dicTitulos = Util.TransformarArrayListDicionario(nCampos)
    
    For Each Titulo In dicTitulos.Keys()
        
        Posicao = dicTitulos(Titulo) - 1
        
        Select Case Titulo
            
            Case "REG"
                If nReg = "0000" Or nReg = "0000_Contr" Then Campos0000.ARQUIVO = fnSPED.GerarCodigoArquivo(Campos, SPEDContr, Unificar)
                
            Case "ARQUIVO"
                If Periodo = "" Then Campos(Posicao) = Campos0000.ARQUIVO Else Campos(Posicao) = Periodo & "-" & CNPJContribuinte
                
            Case "CHV_REG"
                CHV_PAI = IdentificarChavePai(dicReg)
                CHV_REG = fnSPED.CriarChaveRegistro(dicReg, CHV_PAI, Campos)
                
                If VerificarExistenciaChaveRegistro(nReg, CHV_REG) Then Exit Function
                If arrConversoes.contains(nReg) And Not SPEDContr Then Call IncluirRegistrosContribuicoes(nReg)
                
                If Not dicRegFiscal Is Nothing Then CHV_PAI_FISCAL = IdentificarChavePaiFiscal(dicRegFiscal)
                If Not dicRegFiscal Is Nothing Then Call AtribuirChaveNivelFiscal(dicRegFiscal, CHV_REG)
                If Not dicRegContribuicoes Is Nothing Then CHV_PAI_CONTRIBUICOES = IdentificarChavePaiContribuicoes(dicRegContribuicoes)
                If Not dicRegContribuicoes Is Nothing Then Call AtribuirChaveNivelContribuicoes(dicRegContribuicoes, CHV_REG)
                
                Campos(Posicao) = CHV_REG
                If nReg = "F100" Or nReg = "F120" Then
                    Campos(Posicao) = fnSPED.GerarChaveRegistro(CStr(Campos(Posicao)), CInt(dicRegistros(nReg).Count + 1))
                End If
                
            Case "CHV_PAI_FISCAL"
                Campos(Posicao) = CHV_PAI_FISCAL
                
            Case "CHV_PAI_CONTRIBUICOES"
                Campos(Posicao) = CHV_PAI_CONTRIBUICOES
                
            Case Else
                Call fnExcel.DefinirTipoCampos(Campos, dicTitulos, SPEDContr)
                Exit For
                
        End Select
        
    Next Titulo
    
    If IgnorarEmissoesProprias Then
        
        Select Case True
            
            Case nReg = "C100", nReg = "C500", nReg = "D100", nReg = "D500", nReg = "D700"
                If Campos(6) Like "0*" Then
                    IngorarCHV_PAI = Util.RemoverAspaSimples(Campos(2))
                    Exit Function
                End If
                
            Case nReg = "C300", nReg = "C350", nReg = "C400", nReg = "C495", nReg = "C400", nReg = "C600", nReg = "C700", _
                nReg = "C800", nReg = "C860", nReg = "D300", nReg = "D350", nReg = "D400", nReg = "D600", nReg = "D695"
                IngorarCHV_PAI = Util.RemoverAspaSimples(Campos(2))
                Exit Function
                
        End Select
        
        If (CHV_PAI_FISCAL = IngorarCHV_PAI And CHV_PAI_FISCAL <> "") _
            Or (CHV_PAI_CONTRIBUICOES = IngorarCHV_PAI And CHV_PAI_CONTRIBUICOES <> "") Then
            Exit Function
        End If
        
    End If
    
    If nReg Like "0000*" Then
        
        If dicRegistros(nReg).Count < 1 Or Not Unificar Then dicRegistros(nReg).Add Campos
        
    Else
        
        dicRegistros(nReg).Add Campos
        
    End If
    
    If dicRegistros(nReg).Count > 50000 Then
        Call Util.ExportarDadosArrayList(Worksheets(nReg), dicRegistros(nReg))
        Call dicRegistros(nReg).Clear
    End If
    
End Function

Private Function IncluirRegistrosContribuicoes(ByVal nReg As String)

Dim dicReg As Dictionary
Dim Registro As Variant, Campos, r0140, rC010
Dim CHV_REG As String, CHV_PAI_CONTRIBUICOES$
        
    With dtoTitSPED
        
        If .t0000 Is Nothing Then Set .t0000 = Util.MapearTitulos(reg0000, 3)
        If .t0001 Is Nothing Then Set .t0001 = Util.MapearTitulos(reg0001, 3)
        
        Select Case nReg
            
            Case "0001"
                Call IncluirRegistro0000Contribuicoes
                
            Case "0150"
                Call IncluirRegistro0110
                Call IncluirRegistro0140
                
            Case "C100"
                Call IncluirRegistroC010
                
            Case "D100"
                Call IncluirRegistroD010
                
            Case "9001"
                'Call IncluirRegistroF001
                
        End Select
        
    End With
    
End Function

Private Sub IncluirRegistro0000Contribuicoes()

Dim dicReg As Dictionary
Dim arrRegistros As ArrayList
Dim Campos As Variant, r0000
Dim CHV_REG As String, COD_VER$, TIPO_ESCRIT$, Periodo$, CHV_PAI_CONTRIBUICOES$
    
    With dtoTitSPED
        
        Set dicReg = dicLayoutContribuicoes("0000_Contr")
        CHV_PAI_CONTRIBUICOES = IdentificarChavePaiContribuicoes(dicReg)
        
        Set arrRegistros = dicRegistros("0000")
        Campos = arrRegistros(arrRegistros.Count - 1)
        
        If .t0000 Is Nothing Then Set .t0000 = Util.MapearTitulos(reg0000, 3)
        
        Periodo = VBA.Format(Campos(.t0000("DT_INI") - 1), "mmaaaa")
        COD_VER = "'" & EnumContrib.ValidarEnumeracao_COD_VER(Periodo)
        TIPO_ESCRIT = Campos(.t0000("COD_FIN") - 1)
        
        r0000 = Array("'0000", Campos(.t0000("ARQUIVO") - 1), "", "", "", COD_VER, TIPO_ESCRIT, "", "", _
            Campos(.t0000("DT_INI") - 1), Campos(.t0000("DT_FIN") - 1), Campos(.t0000("NOME") - 1), Campos(.t0000("CNPJ") - 1), _
            Campos(.t0000("UF") - 1), Campos(.t0000("COD_MUN") - 1), Campos(.t0000("SUFRAMA") - 1), "", "")
            
        CHV_REG = fnSPED.CriarChaveRegistro(dicReg, "", r0000)
        
        r0000(.t0000("CHV_REG") - 1) = CHV_REG
        
        If Not dicRegistros.Exists("0000_Contr") Then Set dicRegistros("0000_Contr") = New ArrayList
        dicRegistros("0000_Contr").Add r0000
        
        Call AtribuirChaveNivelContribuicoes(dicReg, CHV_REG)
        
    End With
    
End Sub

Private Sub IncluirRegistro0110()

Dim dicReg As Dictionary
Dim arrRegistros As ArrayList
Dim Campos As Variant, r0110
Dim CHV_REG As String, CHV_PAI_CONTRIBUICOES$
    
    With dtoTitSPED
        
        If dicRegistros.Exists("0110") Then
            
            If dicRegistros("0110").Count >= dicRegistros("0000_Contr").Count Then Exit Sub
            
        End If
        
        Set dicReg = dicLayoutContribuicoes("0110")
        CHV_PAI_CONTRIBUICOES = IdentificarChavePaiContribuicoes(dicReg)
        
        Set arrRegistros = dicRegistros("0000")
        Campos = arrRegistros(arrRegistros.Count - 1)
        
        If .t0110 Is Nothing Then Set .t0110 = Util.MapearTitulos(reg0110, 3)
        
        r0110 = Array("0110", Campos(.t0000("ARQUIVO") - 1), "", "", CHV_PAI_CONTRIBUICOES, "", "", "", "")
        
        CHV_REG = fnSPED.CriarChaveRegistro(dicReg, CHV_PAI_CONTRIBUICOES, r0110)
        
        r0110(.t0110("REG") - 1) = "'0110"
        r0110(.t0110("CHV_REG") - 1) = CHV_REG
        
        If Not dicRegistros.Exists("0110") Then Set dicRegistros("0110") = New ArrayList
        dicRegistros("0110").Add r0110
        
        Call AtribuirChaveNivelContribuicoes(dicReg, CHV_REG)
        
    End With
    
End Sub

Private Sub IncluirRegistro0140()

Dim dicReg As Dictionary
Dim arrRegistros As ArrayList
Dim Campos As Variant, r0140
Dim CHV_REG As String, CHV_PAI_CONTRIBUICOES$
    
    With dtoTitSPED
        
        If dicRegistros.Exists("0140") Then
            
            If dicRegistros("0140").Count >= dicRegistros("0000").Count Then Exit Sub
            
        End If
        
        Set dicReg = dicLayoutContribuicoes("0140")
        CHV_PAI_CONTRIBUICOES = IdentificarChavePaiContribuicoes(dicReg)
        
        Set arrRegistros = dicRegistros("0000")
        Campos = arrRegistros(arrRegistros.Count - 1)
        
        If .t0140 Is Nothing Then Set .t0140 = Util.MapearTitulos(reg0140, 3)

        r0140 = Array("0140", Campos(.t0000("ARQUIVO") - 1), "", "", CHV_PAI_CONTRIBUICOES, "", _
            Campos(.t0000("NOME") - 1), Campos(.t0000("CNPJ") - 1), Campos(.t0000("UF") - 1), _
            Campos(.t0000("IE") - 1), Campos(.t0000("COD_MUN") - 1), Campos(.t0000("IM") - 1), Campos(.t0000("SUFRAMA") - 1))
            
        CHV_REG = fnSPED.CriarChaveRegistro(dicReg, CHV_PAI_CONTRIBUICOES, r0140)
        
        r0140(.t0140("REG") - 1) = "'0140"
        r0140(.t0140("CHV_REG") - 1) = CHV_REG
        
        If Not dicRegistros.Exists("0140") Then Set dicRegistros("0140") = New ArrayList
        dicRegistros("0140").Add r0140
        
        Call AtribuirChaveNivelContribuicoes(dicReg, CHV_REG)
        
    End With
    
End Sub

Private Sub IncluirRegistroC010()

Dim dicReg As Dictionary
Dim arrRegistros As ArrayList
Dim Campos As Variant, rC010
Dim CHV_REG As String, CHV_PAI_CONTRIBUICOES$
    
    With dtoTitSPED
        
        If dicRegistros.Exists("C010") Then
            
            If dicRegistros("C010").Count >= dicRegistros("0000").Count Then Exit Sub
            
        End If
        
        Set dicReg = dicLayoutContribuicoes("C010")
        CHV_PAI_CONTRIBUICOES = IdentificarChavePaiContribuicoes(dicReg)
        
        Set arrRegistros = dicRegistros("0000")
        Campos = arrRegistros(arrRegistros.Count - 1)
        
        If .tC010 Is Nothing Then Set .tC010 = Util.MapearTitulos(regC010, 3)
        rC010 = Array("C010", Campos(.t0000("ARQUIVO") - 1), "", "", CHV_PAI_CONTRIBUICOES, Campos(.t0000("CNPJ") - 1), "2")
        
        CHV_REG = fnSPED.CriarChaveRegistro(dicReg, CHV_PAI_CONTRIBUICOES, rC010)
        
        rC010(.tC010("REG") - 1) = "'C010"
        rC010(.tC010("CHV_REG") - 1) = CHV_REG
        
        If Not dicRegistros.Exists("C010") Then Set dicRegistros("C010") = New ArrayList
        dicRegistros("C010").Add rC010
        
        Call AtribuirChaveNivelContribuicoes(dicReg, CHV_REG)
        
    End With
    
End Sub

Private Sub IncluirRegistroD010()

Dim dicReg As Dictionary
Dim arrRegistros As ArrayList
Dim Campos As Variant, rD010
Dim CHV_REG As String, CHV_PAI_CONTRIBUICOES$
    
    With dtoTitSPED
        
        If dicRegistros.Exists("D010") Then
            
            If dicRegistros("D010").Count >= dicRegistros("0000").Count Then Exit Sub
            
        End If
        
        Set dicReg = dicLayoutContribuicoes("D010")
        CHV_PAI_CONTRIBUICOES = IdentificarChavePaiContribuicoes(dicReg)
        
        Set arrRegistros = dicRegistros("0000")
        Campos = arrRegistros(arrRegistros.Count - 1)
        
        If .tD010 Is Nothing Then Set .tD010 = Util.MapearTitulos(regD010, 3)
        rD010 = Array("D010", Campos(.t0000("ARQUIVO") - 1), "", "", CHV_PAI_CONTRIBUICOES, Campos(.t0000("CNPJ") - 1), "2")
        
        CHV_REG = fnSPED.CriarChaveRegistro(dicReg, CHV_PAI_CONTRIBUICOES, rD010)
        
        rD010(.tD010("REG") - 1) = "'D010"
        rD010(.tD010("CHV_REG") - 1) = CHV_REG
        
        If Not dicRegistros.Exists("D010") Then Set dicRegistros("D010") = New ArrayList
        dicRegistros("D010").Add rD010
        
        Call AtribuirChaveNivelContribuicoes(dicReg, CHV_REG)
        
    End With
    
End Sub

Private Sub IncluirRegistroF001()

Dim dicReg As Dictionary
Dim arrRegistros As ArrayList
Dim Campos As Variant, rF001
Dim CHV_REG As String, CHV_PAI$, CHV_PAI_CONTRIBUICOES$, IND_MOV$
    
    With dtoTitSPED
        
        Set dicReg = dicLayoutContribuicoes("F001")
        CHV_PAI_CONTRIBUICOES = IdentificarChavePaiContribuicoes(dicReg)
        
        Set arrRegistros = dicRegistros("0000")
        Campos = arrRegistros(arrRegistros.Count - 1)
        
        If .tF001 Is Nothing Then Set .tF001 = Util.MapearTitulos(regF001, 3)
        IND_MOV = EnumContrib.ValidarEnumeracao_IND_MOV("1")
        rF001 = Array("F001", Campos(.tF001("ARQUIVO") - 1), CHV_REG, "", CHV_PAI_CONTRIBUICOES, "1")
            
        CHV_REG = fnSPED.CriarChaveRegistro(dicReg, CHV_PAI_CONTRIBUICOES, rF001)
        
        If Not dicRegistros.Exists("F001") Then Set dicRegistros("F001") = New ArrayList
        dicRegistros("F001").Add rF001
        
        Call AtribuirChaveNivelContribuicoes(dicReg, CHV_REG)
        
    End With
    
End Sub

Private Function ProcessarCampos(ByVal Registro As String)
    
Dim nReg As String
    
    Registro = VBA.Trim(VBA.Replace(VBA.Replace(Registro, vbCr, ""), vbLf, ""))
    nReg = VBA.Left(Registro, 5) & "||||"
    Registro = nReg & VBA.Right(Registro, VBA.Len(Registro) - 5)
    Registro = VBA.Mid(Registro, 2, Len(Registro) - 2)
    ProcessarCampos = VBA.Split(Registro, "|")
    
End Function

Public Function ExtrairCampo(ByVal Registro As String, NumCampo As Byte)

Dim Campos
    
    Campos = Split(Registro, "|")
    
        ExtrairCampo = Campos(NumCampo)
    
End Function

Public Function ExtrairCamposRegistro(ByVal Registro As String)
    ExtrairCamposRegistro = Split(Registro, "|")
End Function

Public Function GerarRegistro(ByVal Campos As Variant)
    GerarRegistro = Join(Campos, "|")
End Function
 
Public Function GerarCodigoSituacao(ByVal Situacao As Variant) As String
    
    Select Case Situacao
        Case "Autorizada"
            GerarCodigoSituacao = "00"
        
        Case "Cancelada"
            GerarCodigoSituacao = "02"
        
        Case "Denegada"
            GerarCodigoSituacao = "04"
        
        Case "Inutilizada"
            GerarCodigoSituacao = "05"
            
    End Select
    
End Function

Public Function MontarChaveRegistro(ParamArray Campos() As Variant)
    
    Campos = Application.index(Campos(0), 0, 0)
    MontarChaveRegistro = VBA.Replace(VBA.Join(Campos, ""), "'", "")
    
End Function

Private Function MapearCamposChave(ByRef dicReg As Dictionary, ByVal nReg As String, ByVal CHV_PAI As String, ByVal Campos As Variant)
        
Dim arrCamposChave As New ArrayList
Dim Posicao As Integer
Dim nCampo As Variant
Dim nivel As String

    arrCamposChave.Add CHV_PAI
    For Each nCampo In dicReg("CamposChave")
    
        If nCampo > 1 Then Posicao = nCampo + 3 Else Posicao = 0
        If Campos(0) = "0000" And nCampo = 11 Then arrCamposChave.Add Campos(Posicao) * 1 Else arrCamposChave.Add CStr(Campos(Posicao))
        
    Next nCampo
    
    MapearCamposChave = arrCamposChave.toArray()
    
End Function

Public Function IdentificarChavePai(ByRef dicReg As Dictionary) As String
        
Dim nivel As String
    
    nivel = dicReg("Nivel")
    If nivel = "0" Then
        IdentificarChavePai = ""
        Exit Function
    End If
    
    nivel = CInt(nivel) - 1
    IdentificarChavePai = dicChavesNivel(nivel)
    
End Function

Public Function IdentificarChavePaiFiscal(ByRef dicRegFiscal As Dictionary) As String

Dim nivel As String
    
    nivel = dicRegFiscal("Nivel")
    If nivel = "0" Then
        IdentificarChavePaiFiscal = ""
        Exit Function
    End If
    
    nivel = CInt(nivel) - 1
    IdentificarChavePaiFiscal = dicChavesNivelFiscal(nivel)
    
End Function

Public Function IdentificarChavePaiContribuicoes(ByRef dicRegContribuicoes As Dictionary) As String

Dim nivel As String
    
    nivel = dicRegContribuicoes("Nivel")
    If nivel = "0" Then
        IdentificarChavePaiContribuicoes = ""
        Exit Function
    End If
    
    nivel = CInt(nivel) - 1
    IdentificarChavePaiContribuicoes = dicChavesNivelContribuicoes(nivel)
    
End Function

Public Function AtribuirChaveNivel(ByRef dicReg As Dictionary, ByVal CHV_REG As String) As String

Dim nivel As String
    
    nivel = dicReg("Nivel")
    dicChavesNivel(nivel) = CHV_REG

End Function

Public Function AtribuirChaveNivelFiscal(ByRef dicRegFiscal As Dictionary, ByVal CHV_REG As String) As String

Dim nivel As String
    
    nivel = dicRegFiscal("Nivel")
    dicChavesNivelFiscal(nivel) = CHV_REG
    
End Function

Public Function AtribuirChaveNivelContribuicoes(ByRef dicRegContribuicoes As Dictionary, ByVal CHV_REG As String) As String

Dim nivel As String
    
    nivel = dicRegContribuicoes("Nivel")
    dicChavesNivelContribuicoes(nivel) = CHV_REG
    
End Function

Public Function LimparRegistrosEFD()
    
Dim Resposta As VbMsgBoxResult
Dim pasta As New Workbook
Dim Plan As Worksheet
        
    Resposta = MsgBox("Tem certeza que deseja apagar os dados de TODOS registros da EFD?" & vbCrLf & _
                      "Essa operação NÃO pode ser desfeita.", vbCritical + vbYesNo, "Apagar registros da EFD")
                
    If Resposta = vbYes Then
        
        Inicio = Now()
        Call Util.DesabilitarControles
            
            Application.StatusBar = "Limpando registros do SPED, por favor aguarde..."
            
            Set pasta = ThisWorkbook
            For Each Plan In pasta.Sheets
                            
                Select Case True
                    
                    Case (VBA.Left(Plan.CodeName, 3) = "reg") Or (VBA.Left(Plan.CodeName, 3) = "rel")
                        Call Util.DeletarDados(Plan, 4, False)
                        
                End Select
                        
            Next Plan
            
            Call Util.MsgInformativa("Registros deletados com sucesso!", "Limpeza de registros do SPED", Inicio)
            Application.StatusBar = False
        
        Call Util.HabilitarControles
        
    End If
    
End Function

Public Sub ImportarDadosProdutoContribuinte(ByRef dicRelatorio As Dictionary, ByVal Arqs As Variant)

Dim Arq
Dim NFe As New DOMDocument60
Dim Produto As IXMLDOMNode
Dim Produtos As IXMLDOMNodeList
Dim dicProdutos As New Dictionary
Dim vProd As Double, pICMS#, bcICMS#, vICMS#
Dim dhEmi As Variant, Registros, Registro, Campos, CamposDic
Dim nNF As String, chNFe$, cProd$, xProd$, CFOP$, CSTICMS$, CNPJ_FORNEC$, Razao$, uCom$, Chave$, NITEM$, uInv$
    
    For Each Arq In Arqs
        
        Open Arq For Input As #1
            Registros = Util.ImportarTxt(Arq)
        Close #1
        
        For Each Registro In Registros
            
            Select Case Mid(Registro, 2, 4)
                
                Case "0200"
                    Campos = Split(Registro, "|")
                    cProd = Campos(2)
                    xProd = Campos(3)
                    uInv = VBA.UCase(Campos(6))
                    
                    dicProdutos(cProd) = Array(xProd, uInv)
                    
                Case "C100"
                    Campos = Split(Registro, "|")
                    chNFe = Campos(9)
                    
                Case "C170"
                    Campos = Split(Registro, "|")
                    NITEM = Campos(2)
                    
                    Chave = chNFe & CInt(NITEM)
                    If dicRelatorio.Exists(Chave) Then
                        
                        CamposDic = dicRelatorio(Chave)
                            
                            cProd = Campos(3)
                            If dicProdutos.Exists(cProd) Then
                                CamposDic(5) = Util.FormatarTexto(cProd)
                                CamposDic(6) = dicProdutos(cProd)(0)
                                CamposDic(7) = dicProdutos(cProd)(1)
                            End If
                            
                        dicRelatorio(Chave) = CamposDic
                        
                    End If
                    
            End Select
            
        Next Registro
        
    Next Arq
    
End Sub

Public Function GerarChaveRegistro(ParamArray Campos() As Variant) As Variant
    GerarChaveRegistro = Cripto.MD5(fnSPED.MontarChaveRegistro(Campos))
End Function

Public Function CriarChaveRegistro(ByRef dicReg As Dictionary, ByVal CHV_PAI As String, ByVal Campos As Variant) As Variant

Dim nReg As String
Dim CamposChave As Variant
Dim CHV_REG As String
    
    nReg = Campos(LBound(Campos))
    CamposChave = MapearCamposChave(dicReg, nReg, CHV_PAI, Campos)
    CHV_REG = Cripto.MD5(fnSPED.MontarChaveRegistro(CamposChave))
    Call AtribuirChaveNivel(dicReg, CHV_REG)
    CriarChaveRegistro = CHV_REG
    
End Function

Public Function GerarCodigoArquivo(ByRef Campos As Variant, Optional SPEDContr As Boolean, Optional Unificar As Boolean)

Dim Periodo As String, CNPJ$, CPF$, Ano$, Mes$
    
    If SPEDContr Then
        
        Periodo = VBA.Format(Util.FormatarData(Campos(9)), "mm/yyyy")
        CNPJ = Campos(12)
        
    Else
    
        Periodo = VBA.Format(Util.FormatarData(Campos(7)), "mm/yyyy")
        CNPJ = Campos(10)
        CPF = Campos(11)
    
    End If
    
    If Unificar Then
        
        If SPEDContr Then
            
            Ano = VBA.Right(Campos(9), 4)
            Mes = VBA.Mid(Campos(9), 3, 2)
            Campos(9) = "01" & VBA.Right(Campos(9), 6)
            Campos(10) = VBA.Format(Application.WorksheetFunction.EoMonth(Ano & "-" & Mes & "-" & "01", 0), "ddmmyyyy")
        
        Else
        
            Ano = VBA.Right(Campos(7), 4)
            Mes = VBA.Mid(Campos(7), 3, 2)
            Campos(7) = "01" & VBA.Right(Campos(7), 6)
            Campos(8) = VBA.Format(Application.WorksheetFunction.EoMonth(Ano & "-" & Mes & "-" & "01", 0), "ddmmyyyy")
                    
        End If
        
    End If
    
    If CNPJ <> "" Then GerarCodigoArquivo = Periodo & "-" & CNPJ Else GerarCodigoArquivo = Periodo & "-" & CPF
    
End Function

Public Function FormatarNCM(ByVal COD_NCM As String) As String
    If COD_NCM <> "" Then FormatarNCM = VBA.Format(COD_NCM, "0000\.00\.00")
End Function

Public Function FormatarCEST(ByVal COD_CEST As String) As String
    If COD_CEST <> "" Then FormatarCEST = VBA.Format(COD_CEST, "0000\.00\.00")
End Function

Public Function FormatarDataSPED(ByVal Data As Variant) As String
    
Dim DataTeste

    If Data <> "" And Not Data Like "*-*" And Not Data Like "*/*" And VBA.Len(Data) = 8 Then

        DataTeste = VBA.Format(Data, "00/00/0000")
        If IsDate(DataTeste) Then Data = Format(DataTeste, "yyyy-mm-dd")

    End If
    
    FormatarDataSPED = Data
    
End Function

Public Function FormatarPercentuais(ByVal Valor As String) As Double
    
    If Valor = "" Or Valor = "-" Then Valor = 0
    Valor = VBA.Replace(Valor, ".", ",")
    
    If Valor Like "*%*" Then
        Valor = VBA.Replace(Valor, "%", "")
    Else
        Valor = Valor * 100
    End If
    
    FormatarPercentuais = Valor
    
End Function

Public Function CarregarChaveRegistro(ByVal ARQUIVO As String, dicDadosPai As Dictionary, _
    dicDadosAvo As Dictionary, dicTitulosPai As Dictionary, dicTitulosAvo As Dictionary) As String
        
    If dicDadosPai.Exists(ARQUIVO) Then
        
        CarregarChaveRegistro = dicDadosPai(ARQUIVO)(dicTitulosPai("CHV_REG"))
        
    ElseIf dicDadosAvo.Exists(ARQUIVO) Then
    
        CarregarChaveRegistro = dicDadosAvo(ARQUIVO)(dicTitulosAvo("CHV_REG"))
    
    End If

End Function

Public Function ExtrairCadastroSPEDFiscal()

Dim dicDados0000 As New Dictionary
Dim dicTitulos0000 As New Dictionary
Dim ARQUIVO As String
    
    Set dicTitulos0000 = Util.MapearTitulos(reg0000, 3)
    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000, "ARQUIVO")
    
    If dicDados0000.Count = 0 Then
        
        Call Util.MsgAlerta("Não existem dados do SPED Fiscal importados", "SPED não importado")
        Exit Function
        
    End If
    
    ARQUIVO = dicDados0000.Items(0)(dicTitulos0000("ARQUIVO"))
    Call AtuailzarCadastroContribuinte(dicDados0000, dicTitulos0000, ARQUIVO)
    
    Call Util.MsgAviso("Dados importados com sucesso!", "Importação de dados do SPED Fiscal")
    
End Function

Public Function ExtrairCadastroSPEDContribuicoes()

Dim dicDados0000 As New Dictionary
Dim dicTitulos0000 As New Dictionary
Dim ARQUIVO As String
    
    Set dicTitulos0000 = Util.MapearTitulos(reg0000_Contr, 3)
    Set dicDados0000 = Util.CriarDicionarioRegistro(reg0000_Contr, "ARQUIVO")
    
    If dicDados0000.Count = 0 Then
        
        Call Util.MsgAlerta("Não existem dados do SPED Fiscal importados", "SPED não importado")
        Exit Function
        
    End If
    
    ARQUIVO = dicDados0000.Items(0)(dicTitulos0000("ARQUIVO"))
    Call AtuailzarCadastroContribuinte(dicDados0000, dicTitulos0000, ARQUIVO)
    
    Call Util.MsgAviso("Dados importados com sucesso!", "Importação de dados do SPED Fiscal")
    
End Function

Public Function AtuailzarCadastroContribuinte(ByRef dicDados0000 As Dictionary, _
    ByRef dicTitulos0000 As Dictionary, ByRef ARQUIVO As String, Optional ByRef Periodo As String)
    
    CadContrib.Range("CNPJContribuinte").value = dicDados0000(ARQUIVO)(dicTitulos0000("CNPJ"))
    CadContrib.Range("InscContribuinte").value = dicDados0000(ARQUIVO)(dicTitulos0000("IE"))
    CadContrib.Range("UFContribuinte").value = dicDados0000(ARQUIVO)(dicTitulos0000("UF"))
    CadContrib.Range("RazaoContribuinte").value = dicDados0000(ARQUIVO)(dicTitulos0000("NOME"))
    Periodo = ARQUIVO
    
End Function

Public Function ExtrairCamposDicionario(ByRef dicDados As Dictionary, ByVal CHV_REG As String)
    If dicDados.Exists(CHV_REG) Then ExtrairCamposDicionario = dicDados(CHV_REG)
End Function

Public Function ExtrairCampoDicionario(ByRef dicDados As Dictionary, ByRef dicTitulos As Dictionary, ByVal CHV_REG As String, ByVal nCampo As String) As String

Dim Campos
Dim i As Byte
     
    If dicDados.Exists(CHV_REG) Then
        
        Campos = dicDados(CHV_REG)
        i = Util.VerificarPosicaoInicialArray(Campos)
        
        ExtrairCampoDicionario = Campos(dicTitulos(nCampo) - i)
        
    End If
    
End Function

Public Function ExtrairCampoArray(ByRef Campos As Variant, ByRef dicTitulos As Dictionary, ByVal nCampo As String) As String
        
Dim i As Integer
    
    i = Util.VerificarPosicaoInicialArray(Campos)
    ExtrairCampoArray = Campos(dicTitulos(nCampo) - i)
    
End Function

Public Function ChecarErrosEstruturaisICMSIPI(ByRef Msg As String) As Boolean
    
    Select Case True
        
        Case ChecarErrosReg0000(Msg), ChecarErrosReg0005(Msg)
            ChecarErrosEstruturaisICMSIPI = True
            
    End Select
    
End Function

Private Function ChecarErrosReg0000(ByRef Msg As String) As Boolean

Dim dicTitulos As New Dictionary
Dim dicDados As New Dictionary
Dim Campos As Variant
Dim IND_PERFIL As String, IND_ATIV$
Dim i As Long, Coluna&
Dim regex As Object

    Set dicTitulos = Util.MapearTitulos(reg0000, 3)
    Set dicDados = Util.CriarDicionarioRegistro(reg0000)
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^[ABC]$"
    
    i = 0
    For Each Campos In dicDados.Items
        
        i = i + 1
        
        IND_PERFIL = Campos(dicTitulos("IND_PERFIL"))
        If CorrigirPerfilMinusculasSPEDFiscal(IND_PERFIL) Then
            
            Coluna = dicTitulos("IND_PERFIL")
            reg0000.Cells(i + 3, Coluna).value = IND_PERFIL
        
        End If
        
        IND_ATIV = Campos(dicTitulos("IND_ATIV"))
        
        Select Case True
            
            Case IND_PERFIL = ""
                Msg = "O campo 'IND_PERFIL' do registro 0000 não foi informado." & vbCrLf & vbCrLf
                Coluna = dicTitulos("IND_PERFIL")
                Exit For
                
            Case Not regex.Test(IND_PERFIL)
                Msg = "O campo 'IND_PERFIL' do registro 0000 deve possui apenas os valores: 'A', 'B' ou 'C' em maiúsculas." & vbCrLf & vbCrLf
                Coluna = dicTitulos("IND_PERFIL")
                Exit For
                
            Case IND_ATIV = ""
                Msg = "O campo 'IND_ATIV' do registro 0000 não foi informado." & vbCrLf & vbCrLf
                Coluna = dicTitulos("IND_ATIV")
                Exit For
                
        End Select
                
    Next Campos
    
    If Msg <> "" Then
        
        Msg = Msg & "Este campo é obrigatório para a validação do SPED no PVA." & vbCrLf & "Por favor, preencha o campo e tente novamente."
    
        reg0000.Activate
        reg0000.Cells(i + 3, Coluna).Select
        ChecarErrosReg0000 = True
        
    End If
    
End Function

Private Function ChecarErrosReg0005(ByRef Msg As String) As Boolean

Dim dicTitulos As New Dictionary
Dim dicDados As New Dictionary
Dim Campos As Variant
Dim FANTASIA As String
Dim i As Long, Coluna&
Dim regex As Object

    Set dicTitulos = Util.MapearTitulos(reg0005, 3)
    Set dicDados = Util.CriarDicionarioRegistro(reg0005)
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^[ABC]$"
    
    i = 0
    For Each Campos In dicDados.Items
        
        i = i + 1
        FANTASIA = Campos(dicTitulos("FANTASIA"))
                    
        Select Case True
            
            Case FANTASIA = ""
                Msg = "O campo 'FANTASIA' do registro 0005 não foi informado." & vbCrLf & vbCrLf
                Coluna = dicTitulos("FANTASIA")
                Exit For
                
        End Select
                
    Next Campos
    
    If Msg <> "" Then
        
        Msg = Msg & "Este campo é obrigatório para a validação do SPED no PVA." & vbCrLf & "Por favor, preencha o campo e tente novamente."
    
        reg0005.Activate
        reg0005.Cells(i + 3, Coluna).Select
        ChecarErrosReg0005 = True
        
    End If
    
End Function

Public Function ExtrairDadosContador(ByVal Arq As String) As String

Dim NUM As Byte
Dim Registro As String
    
    NUM = FreeFile
    
    Open Arq For Input As #NUM
        
        Do While Not EOF(NUM)
            
            Line Input #NUM, Registro
            If Registro Like "|0100|*" Then
                
                ExtrairDadosContador = Registro
                Exit Do
                
            End If
            
        Loop
        
    Close #NUM
    
End Function

Public Function ExtrairDadosLMC(ByVal Arq As String) As String

Dim NUM As Byte
Dim Chave As Variant
Dim arrDados As New ArrayList
Dim Registro As String, nReg$
    
    NUM = FreeFile
    
    Open Arq For Input As #NUM
        
        Do While Not EOF(NUM)
            
            Line Input #NUM, Registro
            nReg = VBA.Mid(Registro, 2, 4)
            Select Case nReg
                
                Case "0000", "1001", "1300", "1310", "1320"
                    If Not dicRegistros.Exists(nReg) Then Set dicRegistros(nReg) = New ArrayList
                    Call fnSPED.ProcessarRegistro(Registro)

            End Select
            
        Loop
        
    Close #NUM
    
End Function

Public Function ExtrairCadastroProdutos(ByVal Arq As String) As String

Dim NUM As Byte
Dim Registro As String, nReg$
    
    NUM = FreeFile
    
    Open Arq For Input As #NUM
        
        Do While Not EOF(NUM)
            
            Line Input #NUM, Registro
            nReg = VBA.Mid(Registro, 2, 4)
            Select Case True
                
                Case nReg = "0000", nReg = "0001", nReg = "0200"
                    If Not dicRegistros.Exists(nReg) Then Set dicRegistros(nReg) = New ArrayList
                    Call fnSPED.ProcessarRegistro(Registro)
                    
                Case nReg > "0221"
                    Exit Do
                    
            End Select
            
        Loop
        
    Close #NUM
    
End Function

Public Function ValidarSPEDFiscal(ByVal Arq As String) As Boolean
    
Dim versao As String
    
    If fnSPED.ClassificarSPED(Arq, versao) = "Fiscal" Then
        
        Application.StatusBar = "Carregando layout do SPED Fiscal, por favor aguarde..."
        Call CarregarLayoutSPEDFiscal(versao)
        
        ValidarSPEDFiscal = True
        
    End If
    
End Function

Function ClassificarSPED(ByVal Arq As String, Optional ByRef versao As String, Optional ByRef Periodo As String) As String

Dim Campos As Variant
Dim Registro As String
    
    Registro = ExtrairRegistroAbertura(Arq)
    
    If Not Registro Like "|0000*" Then
        ClassificarSPED = "Desconhecido"
        Exit Function
    End If
    
    Campos = VBA.Split(Registro, "|")
    
    If IsDate(Util.FormatarData(Campos(4))) And IsDate(Util.FormatarData(Campos(5))) Then
        
        ClassificarSPED = "Fiscal"
        versao = Campos(2)
        Periodo = VBA.Right(Campos(4), 6)
        
    ElseIf IsDate(Util.FormatarData(Campos(6))) And IsDate(Util.FormatarData(Campos(7))) Then
        
        ClassificarSPED = "Contribuições"
        versao = Campos(2)
        Periodo = VBA.Right(Campos(6), 6)
        
    End If
    
End Function

Public Function PrepararSPEDsImportacao(Optional ByVal SelReg As String, Optional Periodo As String, _
    Optional Unificar As Boolean, Optional Lote As Boolean, Optional TipoSPED As String)
    
Dim Arqs As Variant
Dim arrFiscal As New ArrayList
Dim Caminho As String, reg0000$
Dim arrContribuicoes As New ArrayList
    
    If Lote Then
        
        Caminho = Util.SelecionarPasta("Selecione a pasta que contém os SPEDs")
        If Caminho = "" Then Exit Function
        
        Call fnSPED.ListarSPEDs(Caminho, arrFiscal, arrContribuicoes, Comeco)
        
    Else
        
        Arqs = Util.SelecionarArquivos("txt")
        If VarType(Arqs) = vbBoolean Then Exit Function
        
        If TipoSPED = "Fiscal" Then Call fnSPED.Fiscal.ListarSPEDsFiscais(Arqs, arrFiscal)
        If TipoSPED = "Contribuicoes" Then Call fnSPED.Contribuicoes.ListarSPEDsContribuicoes(Arqs, arrContribuicoes)
        
    End If
    
    If arrFiscal.Count > 0 Then
        
        Select Case True
            
            Case Unificar
                Set arrFiscal = ListarArquivosCentralizados(arrFiscal, reg0000)
                
        End Select
        
        Call ImportarSPED(SelReg, Periodo, Unificar)
        
    End If
    
End Function

Private Function CorrigirPerfilMinusculasSPEDFiscal(ByRef IND_PERFIL As String) As Boolean
    
    Select Case IND_PERFIL
        
        Case "a", "b", "c"
            IND_PERFIL = VBA.UCase(IND_PERFIL)
            CorrigirPerfilMinusculasSPEDFiscal = True
            
    End Select
    
End Function

Private Function ListarChavesRegistrosSPED(ByVal nReg As String)
    
    If Not dicChavesRegistroSPED.Exists(nReg) Then Set dicChavesRegistroSPED(nReg) = Util.ListarValoresUnicos(Worksheets(nReg), 4, 3, "CHV_REG")
    
End Function

Private Function VerificarExistenciaChaveRegistro(ByVal nReg As String, ByVal CHV_REG As String) As Boolean

Dim arrChaves As New ArrayList
    
    If dicChavesRegistroSPED.Exists(nReg) Then
        
        Set arrChaves = dicChavesRegistroSPED(nReg)
        If arrChaves.contains(CHV_REG) Then VerificarExistenciaChaveRegistro = True
        
    End If
    
End Function

Public Function ExtrairRegistroAbertura(ByVal Arq As String) As String

Dim NUM As Byte
Dim Registro As String
    
    NUM = FreeFile
    
    Open Arq For Input As #NUM
        
        Line Input #NUM, Registro
        
    Close #NUM
    
    ExtrairRegistroAbertura = Registro
    
End Function

Public Sub ListarSPEDs(ByVal Caminho As String, ByRef arrFiscal As ArrayList, _
    ByRef arrContribuicoes As ArrayList, Optional ByVal Comeco As Double)
    
Dim fso As New FileSystemObject
Dim ARQUIVO As Scripting.file
Dim subpasta As Folder
Dim TipoSPED As String
Dim pasta As Folder

    Set pasta = fso.GetFolder(Caminho)
    For Each ARQUIVO In pasta.Files
        
        Call Util.AntiTravamento(a, 1, "Listando SPEDs para importação, por favor aguarde", pasta.Files.Count, Comeco)
        
        If VBA.LCase(ARQUIVO.Path) Like "*.txt" Then
            
            TipoSPED = ClassificarSPED(ARQUIVO.Path)
            Select Case TipoSPED
                
                Case "Fiscal"
                    arrFiscal.Add ARQUIVO.Path
                    
                Case "Contribuições"
                    arrContribuicoes.Add ARQUIVO.Path
                    
            End Select
            
        End If
        
    Next ARQUIVO
    
    For Each subpasta In pasta.SubFolders
        Call ListarSPEDs(subpasta.Path, arrFiscal, arrContribuicoes, Comeco)
    Next subpasta
    
End Sub

Public Function ExtrairRegistrosSPED(ByVal Arq As String)

Dim fso As New FileSystemObject
Dim ARQUIVO As TextStream
Dim Conteudo As String
    
    Set ARQUIVO = fso.OpenTextFile(Arq, ForReading, False)
    Conteudo = ARQUIVO.ReadAll
    ARQUIVO.Close
    
    ExtrairRegistrosSPED = VBA.Split(Conteudo, vbCrLf)
    If UBound(ExtrairRegistrosSPED) <= 1 Then ExtrairRegistrosSPED = VBA.Split(Conteudo, vbLf)
    
End Function
