Attribute VB_Name = "clsSimilaridadeJaccard"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Function CalcularSimilaridadeJaccard(ByVal Texto1 As String, ByVal Texto2 As String) As Double

Dim Palavras1 As New Dictionary
Dim Palavras2 As New Dictionary
Dim arrPalavras1 As Variant
Dim arrPalavras2 As Variant
Dim Intersecao As Long
Dim Palavra As Variant
Dim Chave As Variant
Dim Uniao As Long
    
    Texto1 = VBA.LCase(Texto1)
    Texto2 = VBA.LCase(Texto2)
    Texto1 = RemoverPontuacao(Texto1)
    Texto2 = RemoverPontuacao(Texto2)
    
    arrPalavras1 = VBA.Split(Texto1, " ")
    arrPalavras2 = VBA.Split(Texto2, " ")
    
    For Each Palavra In arrPalavras1
        
        If Len(Trim(Palavra)) > 0 Then
            
            If Not Palavras1.Exists(Palavra) Then
                
                Palavras1.Add Palavra, 1
                
            End If
            
        End If
        
    Next Palavra
    
    For Each Palavra In arrPalavras2
        
        If Len(Trim(Palavra)) > 0 Then
            
            If Not Palavras2.Exists(Palavra) Then
                
                Palavras2.Add Palavra, 1
                
            End If
            
        End If
        
    Next Palavra
    
    Intersecao = 0
    Uniao = 0
    
    For Each Chave In Palavras1.Keys
        
        If Palavras2.Exists(Chave) Then
            
            Intersecao = Intersecao + 1
            
        End If
        
        Uniao = Uniao + 1
        
    Next Chave
    
    Uniao = Uniao + (Palavras2.Count - Intersecao)
    
    If Uniao = 0 Then
        
        CalcularSimilaridadeJaccard = 0
        
    Else
        
        CalcularSimilaridadeJaccard = Intersecao / Uniao
        
    End If
    
End Function

Function RemoverPontuacao(ByVal Texto As String) As String

Dim TextoLimpo As String
Dim i As Integer
Dim CharCode As Integer
Dim Char As String
    
    TextoLimpo = ""
    
    Texto = RemoverAcentuacao(Texto)
    
    For i = 1 To Len(Texto)
        
        Char = Mid(Texto, i, 1)
        CharCode = Asc(Char)
        
        Select Case True
            
            Case (CharCode >= 48 And CharCode <= 57) Or (CharCode >= 65 And CharCode <= 90) _
                Or (CharCode >= 97 And CharCode <= 122) Or (CharCode = 32)
                TextoLimpo = TextoLimpo & Char
                
        End Select
        
    Next i
    
    RemoverPontuacao = TextoLimpo
    
End Function

Function RemoverAcentuacao(ByVal Texto As String) As String
    
Dim TextoSemAcento As String
Dim Char As String
Dim i As Integer
    
    TextoSemAcento = ""
    
    For i = 1 To Len(Texto)
        
        Char = Mid(Texto, i, 1)
        
        Select Case Char
            
            Case "á", "à", "ã", "â", "ä"
                TextoSemAcento = TextoSemAcento & "a"
                
            Case "é", "è", "ê", "ë"
                TextoSemAcento = TextoSemAcento & "e"
                
            Case "í", "ì", "î", "ï"
                TextoSemAcento = TextoSemAcento & "i"
                
            Case "ó", "ò", "õ", "ô", "ö"
                TextoSemAcento = TextoSemAcento & "o"
                
            Case "ú", "ù", "û", "ü"
                TextoSemAcento = TextoSemAcento & "u"
                
            Case "ç"
                TextoSemAcento = TextoSemAcento & "c"
                
            Case Else
                TextoSemAcento = TextoSemAcento & Char
                
        End Select
        
    Next i
    
    RemoverAcentuacao = TextoSemAcento
    
End Function
