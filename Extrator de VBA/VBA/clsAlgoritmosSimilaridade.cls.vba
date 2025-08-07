Attribute VB_Name = "clsAlgoritmosSimilaridade"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'------------------------------------------------------------------------------
' Classe utilitária centralizando algoritmos de similaridade textual
'------------------------------------------------------------------------------
' Todos os métodos retornam um valor em [0..1], onde 1 indica igualdade plena.
'------------------------------------------------------------------------------

' === JARO-WINKLER ===========================================================
Public Function CalcularSimilaridadeJaroWinkler(ByVal s1 As String, ByVal s2 As String) As Double
    Dim m As Long, t As Long, i As Long, j As Long
    Dim s1Len As Long, s2Len As Long, faixa As Long

    s1Len = Len(s1): s2Len = Len(s2)
    If s1Len = 0 Or s2Len = 0 Then Exit Function

    faixa = Application.WorksheetFunction.Max(s1Len, s2Len) \ 2 - 1
    Dim s1Matches() As Boolean: ReDim s1Matches(1 To s1Len)
    Dim s2Matches() As Boolean: ReDim s2Matches(1 To s2Len)

    ' Contar caracteres coincidentes dentro da faixa
    For i = 1 To s1Len
        Dim low As Long: low = Application.WorksheetFunction.Max(1, i - faixa)
        Dim high As Long: high = Application.WorksheetFunction.Min(i + faixa, s2Len)
        For j = low To high
            If Not s2Matches(j) And Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
                s1Matches(i) = True
                s2Matches(j) = True
                m = m + 1
                Exit For
            End If
        Next j
    Next i
    If m = 0 Then Exit Function

    ' Contar transposições
    Dim k As Long: k = 1
    For i = 1 To s1Len
        If s1Matches(i) Then
            Do While Not s2Matches(k): k = k + 1: Loop
            If Mid$(s1, i, 1) <> Mid$(s2, k, 1) Then t = t + 1
            k = k + 1
        End If
    Next i

    ' Fórmula Jaro
    Dim jaro As Double
    jaro = (m / s1Len + m / s2Len + (m - t \ 2) / m) / 3

    ' Bônus Winkler para prefixo
    Dim l As Long: l = 0
    For i = 1 To 4
        If Mid$(s1, i, 1) = Mid$(s2, i, 1) Then l = l + 1 Else Exit For
    Next i

    CalcularSimilaridadeJaroWinkler = jaro + 0.1 * l * (1 - jaro)
End Function

' === JACCARD (por palavra) ===================================================
Public Function CalcularSimilaridadeJaccard(ByVal Texto1 As String, ByVal Texto2 As String) As Double

Dim Palavras1 As Object: Set Palavras1 = CreateObject("Scripting.Dictionary")
Dim Palavras2 As Object: Set Palavras2 = CreateObject("Scripting.Dictionary")
Dim arrPalavras1 As Variant, arrPalavras2 As Variant
Dim Intersecao As Long, Uniao As Long
Dim Palavra As Variant, Chave As Variant

    Texto1 = LCase$(Texto1)
    Texto2 = LCase$(Texto2)
    Texto1 = RemoverPontuacao(Texto1)
    Texto2 = RemoverPontuacao(Texto2)

    arrPalavras1 = Split(Texto1, " ")
    arrPalavras2 = Split(Texto2, " ")

    For Each Palavra In arrPalavras1
        Palavra = Trim$(Palavra)
        If Len(Palavra) > 0 Then If Not Palavras1.Exists(Palavra) Then Palavras1.Add Palavra, 1
    Next Palavra

    For Each Palavra In arrPalavras2
        Palavra = Trim$(Palavra)
        If Len(Palavra) > 0 Then If Not Palavras2.Exists(Palavra) Then Palavras2.Add Palavra, 1
    Next Palavra

    Intersecao = 0: Uniao = 0
    For Each Chave In Palavras1.Keys
        If Palavras2.Exists(Chave) Then Intersecao = Intersecao + 1
        Uniao = Uniao + 1
    Next Chave
    Uniao = Uniao + (Palavras2.Count - Intersecao)

    If Uniao = 0 Then
        CalcularSimilaridadeJaccard = 0
    Else
        CalcularSimilaridadeJaccard = Intersecao / Uniao
    End If
    
End Function

' === HELPERS ================================================================
Private Function RemoverPontuacao(ByVal Texto As String) As String

Dim TextoLimpo As String, i As Long, CharCode As Integer, ch As String
    
    Texto = RemoverAcentuacao(Texto)
    
    For i = 1 To Len(Texto)
        
        ch = Mid$(Texto, i, 1)
        CharCode = Asc(ch)
        Select Case True
            
            Case (CharCode >= 48 And CharCode <= 57) _
                 Or (CharCode >= 65 And CharCode <= 90) _
                 Or (CharCode >= 97 And CharCode <= 122) _
                 Or (CharCode = 32)
                 TextoLimpo = TextoLimpo & ch
        
        End Select
    
    Next i
    
    RemoverPontuacao = TextoLimpo
    
End Function

Private Function RemoverAcentuacao(ByVal Texto As String) As String

Dim i As Long, ch As String, sb As String
    
    For i = 1 To Len(Texto)
        
        ch = Mid$(Texto, i, 1)
        Select Case ch
            
            Case "á", "à", "ã", "â", "ä": sb = sb & "a"
            Case "é", "è", "ê", "ë": sb = sb & "e"
            Case "í", "ì", "î", "ï": sb = sb & "i"
            Case "ó", "ò", "õ", "ô", "ö": sb = sb & "o"
            Case "ú", "ù", "û", "ü": sb = sb & "u"
            Case "ç": sb = sb & "c"
            Case Else: sb = sb & ch
        
        End Select
    
    Next i
    
    RemoverAcentuacao = sb
    
End Function


