Attribute VB_Name = "FuncoesAPI"
Option Explicit

Public Function ConsultarCEP(ByRef Plan As Worksheet)

Dim HotREST As New WinHttp.WinHttpRequest
Dim dicTitulos As New Dictionary
Dim item As Variant
Dim json As Object
Dim Vencimento As String, Consulta$, Msg$, Campo$
Dim CEP As String
    
    If Plan.name = "0150" Then Campo = "COD_PAIS" Else Campo = "CEP"
    
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    CEP = Plan.Cells(ActiveCell.Row, dicTitulos(Campo)).value
                            
    If Not IsNumeric(CEP) Or VBA.Len(CEP) <> 8 Then
        If VBA.Len(CEP) <> 8 Then Msg = "Informe o CEP com 8 dígitos numéricos."
        Call Util.MsgAlerta(Msg, "CEP inválido")
        Exit Function
    End If
    
    Set HotREST = CreateObject("WinHttp.WinHttpRequest.5.1")
    With HotREST
        
        .Open "GET", "https://viacep.com.br/ws/" & ActiveCell.value & "/json", False
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        
        Do
            If .Status <> 0 Then Exit Do
        Loop
        
        Set json = JsonConverter.ParseJson(.ResponseText)
        
        If Plan.name = "0005" Then
            
            With reg0005
            
                If json.Exists("erro") Then .Cells(ActiveCell.Row, dicTitulos("END")).value = "CEP não encontrado": Exit Function
                
                .Cells(ActiveCell.Row, dicTitulos("ENDERECO")).value = json("logradouro")
                .Cells(ActiveCell.Row, dicTitulos("COMPL")).value = json("complemento")
                .Cells(ActiveCell.Row, dicTitulos("BAIRRO")).value = json("bairro")
                
            End With
            
        ElseIf Plan.name = "0100" Then
            
            With reg0100
            
                If json.Exists("erro") Then .Cells(ActiveCell.Row, dicTitulos("ENDERECO")).value = "CEP não encontrado": Exit Function
                
                .Cells(ActiveCell.Row, dicTitulos("ENDERECO")).value = json("logradouro")
                .Cells(ActiveCell.Row, dicTitulos("COMPL")).value = json("complemento")
                .Cells(ActiveCell.Row, dicTitulos("BAIRRO")).value = json("bairro")
                .Cells(ActiveCell.Row, dicTitulos("COD_MUN")).value = json("ibge")
                
            End With
                        
        ElseIf Plan.name = "0150" Then
            
            With reg0150
            
                If json.Exists("erro") Then .Cells(ActiveCell.Row, dicTitulos("COD_PAIS")).value = "CEP não encontrado": Exit Function
                
                .Cells(ActiveCell.Row, dicTitulos("COD_PAIS")).value = "1058"
                .Cells(ActiveCell.Row, dicTitulos("COD_MUN")).value = json("ibge")
                .Cells(ActiveCell.Row, dicTitulos("END")).value = json("logradouro")
                .Cells(ActiveCell.Row, dicTitulos("COMPL")).value = json("complemento")
                .Cells(ActiveCell.Row, dicTitulos("BAIRRO")).value = json("bairro")
                
            End With
            
        End If
        
    End With
    
End Function

Public Function EnviarDadosParaCloudflareWorker()
    
Dim i As Long
Dim URL As String
Dim http As Object
Dim Arqs As Variant, Arq
Dim Arquivos As String, Dados$, Dado$
    
    ' URL do seu worker do Cloudflare
    URL = "https://api.escoladaautomacaofiscal.com.br/controldocs"
    
    Arqs = Util.SelecionarArquivos("txt", "Selecione os arquivos que deseja importar")
    If VarType(Arqs) = 11 Then Exit Function
    
    Arquivos = "["
    For i = LBound(Arqs) To UBound(Arqs)
        
        Open Arqs(i) For Input As #1
            Dado = Util.ImportarConteudo(Arqs(i))
            If i = UBound(Arqs) Then Dados = Dados & """" & Dado Else Dados = Dados & """" & Dado & """, """
        Close #1
        
    Next i
    
    Arquivos = Arquivos & Dados & "]"
    ' Inicializando a variável http
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Configurando e enviando a requisição POST
    http.Open "POST", URL, False
    
    ' Adicione headers se necessário, por exemplo: application/json
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send "{""funcao"":""PROCESSAR_SPED_FISCAL"", " & """arquivos"":" & Arquivos & "}"
    
    ' Verificando a resposta
    If http.Status = 200 Then
        MsgBox "Dados enviados com sucesso!", vbInformation
    Else
        MsgBox "Falha ao enviar dados. Status: " & http.Status, vbCritical
    End If
    
End Function

