Attribute VB_Name = "ResumoCFOP"
Option Explicit

Public Sub GerarResumoCFOP()

Dim Arqs, Arq, Registro
Dim Relatorio As New Dictionary
Dim Modelo As String

    'Carregar os endereços dos arquivos que eu quero trabalhar
    Arqs = Util.SelecionarArquivos("txt")
    
    If VarType(Arqs) <> 11 Then
    
        'Trabalhar com os arquivos individualmente
        For Each Arq In Arqs
            
            'Importar os dados do arquivo selecionado
            Arq = Util.ImportarTxt(Arq)
            
            'Percorre os registros do arquivo selecionado
            For Each Registro In Arq
                
                'Identificar o registro as ser trabalhado
                Select Case Mid(Registro, 2, 4) 'Trás o nome do registro. Exemplo: 0000, 0200, C100, C190
                    
                    Case "C100", "C500", "C800", "D100"
                        Modelo = ExtrairModelo(Registro)
                        
                    Case "C190", "C590", "C850", "D190"
                        'Extrair os dados do SPED para o dicionário
                        Call ExtrairDadosAnaliticosEFD(Registro, Modelo, Relatorio)

                End Select
                
            Next Registro
            
        Next Arq
        
        Call Util.ExportarDadosDicionario(relICMS, Relatorio)
        
    End If
    
    Relatorio.RemoveAll
    
End Sub

Private Sub ExtrairDadosAnaliticosEFD(ByVal Registro As String, ByVal Modelo As String, ByRef Dicionario As Dictionary)
    
Dim Campos

    Campos = Split(Registro, "|")
    
    With DadosDoce
    
        .CFOP = Campos(3)
        .CSTICMS = Campos(2)
        .pICMS = Util.FormatarValores(Campos(4)) / 100
        .vOperacao = Campos(5)
        .bcICMS = Campos(6)
        .vICMS = Campos(7)
        
        Select Case Campos(1)
        
            Case "C190", "C590"
                .vBCST = Campos(8)
                .vICMSST = Campos(9)
                .vRedBCICMS = Campos(10)
                .vIPI = Campos(11)
            
            Case "D190"
                .vRedBCICMS = Campos(8)
                
        End Select
        
        If Campos(1) = "C590" Then .vIPI = 0
        .Chave = Join(Array(.CFOP & .CSTICMS & .pICMS), "")
        
        If Dicionario.Exists(.Chave) Then
        
            .vOperacao = .vOperacao + Dicionario(.Chave)(3)
            .bcICMS = .bcICMS + Dicionario(.Chave)(4)
            .vICMS = .vICMS + Dicionario(.Chave)(5)
            .vBCST = .vBCST + Dicionario(.Chave)(6)
            .vICMSST = .vICMSST + Dicionario(.Chave)(7)
            .vRedBCICMS = .vRedBCICMS + Dicionario(.Chave)(8)
            .vIPI = .vIPI + Dicionario(.Chave)(9)
            
        End If
        
        Select Case VBA.Right(.CSTICMS, 2)
            
            Case "20", "30", "40", "41", "70"
                .vIsentas = Round(CDbl(.vOperacao) - CDbl(.bcICMS) - CDbl(.vICMSST) - CDbl(.vIPI), 2)
                .vOutras = 0
                
            Case Else
                .vOutras = Round(CDbl(.vOperacao) - CDbl(.bcICMS) - CDbl(.vICMSST) - CDbl(.vIPI), 2)
                .vIsentas = 0
                
        End Select
        
        Dicionario(.Chave) = Array(CInt(.CFOP), "'" & .CSTICMS, CDbl(.pICMS), CDbl(.vOperacao), CDbl(.bcICMS), CDbl(.vICMS), _
                                    CDbl(.vBCST), CDbl(.vICMSST), CDbl(.vRedBCICMS), CDbl(.vIPI), CDbl(.vIsentas), CDbl(.vOutras))
        
    End With
    
End Sub

Private Sub D190_ExtrairDados(ByVal Registro As String, ByVal Modelo As String, ByRef Dicionario As Dictionary)
    
Dim Campos

    Campos = Split(Registro, "|")
    
    With DadosDoce
    
        .CFOP = Campos(3)
        .CSTICMS = Campos(2)
        .pICMS = Util.FormatarValores(Campos(4)) / 100
        .vOperacao = Campos(5)
        .bcICMS = Campos(6)
        .vICMS = Campos(7)
        .vRedBCICMS = Campos(8)
        
        .Chave = Modelo & .CFOP & .CSTICMS & .pICMS
        
        If Dicionario.Exists(.Chave) Then
        
            .vOperacao = .vOperacao + Dicionario(.Chave)(3)
            .bcICMS = .bcICMS + Dicionario(.Chave)(4)
            .vICMS = .vICMS + Dicionario(.Chave)(5)
            .vRedBCICMS = .vRedBCICMS + Dicionario(.Chave)(8)
        
        End If
        
        Select Case VBA.Right(.CSTICMS, 2)
            
            Case "20", "30", "40", "41", "70"
                .vIsentas = Round(CDbl(.vOperacao) - CDbl(.bcICMS), 2)
                .vOutras = 0
                
            Case Else
                .vOutras = Round(CDbl(.vOperacao) - CDbl(.bcICMS), 2)
                .vIsentas = 0
                
        End Select
        
        Dicionario(.Chave) = Array(CInt(.CFOP), "'" & .CSTICMS, CDbl(.pICMS), CDbl(.vOperacao), CDbl(.bcICMS), _
                                    CDbl(.vICMS), 0, 0, CDbl(.vRedBCICMS), 0, CDbl(.vIsentas), CDbl(.vOutras))
        
    End With
    
End Sub

Public Function ExtrairModelo(ByVal Registro As String) As String

Dim Campos

    Campos = Split(Registro, "|")
    
    Select Case Campos(1)
    
        Case "C100", "C500", "D100"
            ExtrairModelo = Campos(5)
        
        Case "C800"
            ExtrairModelo = Campos(2)
        
    End Select

End Function
