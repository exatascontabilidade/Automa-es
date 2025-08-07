Attribute VB_Name = "FuncoesTXT"
Option Explicit

Public Function ImportarListaProdutosExcluidos(ByRef arrProdutosExcluidos As ArrayList)

Dim Registros As Variant, Registro, Arq

    Arq = Util.SelecionarArquivo("txt", "Selecione a lista de produtos que deseja excluir dos ajustes")
    
    If (Arq <> False) And (Arq <> "") Then
    
        Registros = "" ' fnSPED.ImportarDadosEFD(Arq)
        For Each Registro In Registros
            arrProdutosExcluidos.Add Registro
        Next Registro
        
    End If
    
End Function

Public Function ImportarListaProdutos(ByRef arrProdutos As ArrayList)

Dim Registros As Variant, Registro, Arq

    Arq = Util.SelecionarArquivo("txt", "Selecione a lista de produtos que deseja importar")
    
    Inicio = Now()
    If (Arq <> False) And (Arq <> "") Then
    
        Registros = "" ' fnSPED.ImportarDadosEFD(Arq)
        For Each Registro In Registros
            arrProdutos.Add Registro
        Next Registro
        
    End If
    
End Function

Function AlterarChaveEAF(ByVal Chave As String) As String

    Dim soma As Integer
    Dim resto As Integer
    Dim Digito As Integer
    Dim i As Integer
    Dim Peso As Integer
    Dim CNPJEmit As String
    
    soma = 0
    Peso = 2
    CNPJEmit = VBA.Mid(Chave, 7, 14)
    Chave = VBA.Replace(Chave, CNPJEmit, "23896999000102")
    Chave = VBA.Left(Chave, 43)
    
    'Loop reverso na chave, começando do último dígito
    For i = Len(Chave) To 1 Step -1
        soma = soma + (Mid(Chave, i, 1) * Peso)
        Peso = Peso + 1
        If Peso > 9 Then Peso = 2
    Next i
    
    resto = soma Mod 11
    
    If resto < 2 Then
        Digito = 0
    Else
        Digito = 11 - resto
    End If
    
    AlterarChaveEAF = Chave & Digito
    
End Function

Public Function ExportarParaTxt(ByRef Plan As Worksheet, ParamArray Campos() As Variant)

Dim dicDados As New Dictionary
Dim arrTXT As New ArrayList
Dim Caminho As String, Arq$
Dim Chave As Variant
    
    Caminho = Util.SelecionarPasta("Selecione a pasta para onde deseja exportar o arquivo TXT.")
    Inicio = Now()
    
    Application.StatusBar = "Exportando os dados, por favor aguarde..."
    
    Arq = "\CadastroTributacao.txt"
    Caminho = Caminho & Arq
    
    Call FuncoesExcel.CarregarDados(Plan, dicDados, Campos)
    If dicDados.Count = 2 Then Call Util.MsgAlerta("Não existem dados para Exportar.", "Sem Dados para exportar"): Exit Function
    
    For Each Chave In dicDados.Keys()
        arrTXT.Add VBA.Join(dicDados(Chave), "|")
    Next Chave
    
    If Caminho <> Arq Then
        
        Call Util.ExportarTxt(Caminho, VBA.Join(arrTXT.toArray, vbCrLf))
        Call Util.MsgInformativa("Dados exportados com sucesso!" & vbCrLf & "O arquivo foi salvo em: " & Caminho, "Exportação do Cadastro de Tributação", Inicio)
    
    End If
    
    Application.StatusBar = False
    
End Function
