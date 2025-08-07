Attribute VB_Name = "FuncoesAjusteSPED"
Option Explicit

Public Sub RemoverCadastrosNaoReferenciados(ByVal Arq As String, ByRef EFD As ArrayList, Optional ByRef SPEDContr As Boolean)

Dim nReg As String, cProd$, cNat$, cUnd$, cPart$, cBem$, cInf$, cObs$, CNPJ$
Dim Registro As Variant, Registros As Variant
Dim arrProdRef As New ArrayList
Dim arrBensRef As New ArrayList
Dim arrPartRef As New ArrayList
Dim arrNatRef As New ArrayList
Dim arrObsRef As New ArrayList
Dim arrInfRef As New ArrayList
Dim arrUndRef As New ArrayList
Dim a As Long

    Registros = VBA.Split(VBA.Join(EFD.toArray, vbCrLf), vbCrLf)
    EFD.Clear
    
    Call CarregarParticipantesReferenciados(Registros, arrPartRef, SPEDContr)
    Call CarregarProdutosReferenciados(Registros, arrProdRef, SPEDContr)
    Call CarregarBensReferenciados(Registros, arrBensRef, SPEDContr)
    Call CarregarNaturezasReferenciadas(Registros, arrNatRef, SPEDContr)
    Call CarregarInformacoesReferenciadas(Registros, arrInfRef, SPEDContr)
    Call CarregarObservacoesReferenciadas(Registros, arrObsRef, SPEDContr)
    Call CarregarUnidadesReferenciadas(Registros, arrProdRef, arrUndRef, SPEDContr)
    
    a = 0
    For Each Registro In Registros
        
        Call Util.AntiTravamento(a, 100, "Analisando dados do arquivo gerado, por favor aguarde...")
        nReg = VBA.Mid(Registro, 2, 4)
        Select Case True
                                
            Case nReg = "0140"
                CNPJ = fnSPED.ExtrairCampo(Registro, 4)
                EFD.Add Registro
                
            Case (nReg = "0150")
                cPart = DefinirChave(CNPJ, fnSPED.ExtrairCampo(Registro, 2), SPEDContr)
                If arrPartRef.contains(cPart) Then EFD.Add Registro
                
            Case (nReg = "0190")
                cUnd = DefinirChave(CNPJ, fnSPED.ExtrairCampo(Registro, 2), SPEDContr)
                If arrUndRef.contains(cUnd) Then EFD.Add Registro
                
            Case (nReg = "0200")
                cProd = DefinirChave(CNPJ, fnSPED.ExtrairCampo(Registro, 2), SPEDContr)
                If arrProdRef.contains(cProd) Then EFD.Add Registro
            
            Case (nReg = "0205") Or (nReg = "0206") Or (nReg = "0210") Or (nReg = "0220") Or (nReg = "0221")
                If arrProdRef.contains(cProd) Then EFD.Add Registro
            
            Case (nReg = "0300")
                cBem = DefinirChave(CNPJ, fnSPED.ExtrairCampo(Registro, 2), SPEDContr)
                If arrBensRef.contains(cBem) Then EFD.Add Registro
            
            Case (nReg = "0305")
                If arrBensRef.contains(cBem) Then EFD.Add Registro
                
            Case (nReg = "0400")
                cNat = DefinirChave(CNPJ, fnSPED.ExtrairCampo(Registro, 2), SPEDContr)
                If arrNatRef.contains(cNat) Then EFD.Add Registro
            
            Case (nReg = "0450")
                cInf = DefinirChave(CNPJ, fnSPED.ExtrairCampo(Registro, 2), SPEDContr)
                If arrInfRef.contains(cInf) Then EFD.Add Registro
                
            Case (nReg = "0460")
                cObs = DefinirChave(CNPJ, fnSPED.ExtrairCampo(Registro, 2), SPEDContr)
                If arrObsRef.contains(cObs) Then EFD.Add Registro
                
            Case (nReg <> "")
                EFD.Add Registro
                
        End Select
        
    Next Registro

    Call fnSPED.ExportarSPED(Arq, fnSPED.TotalizarRegistrosSPED(EFD, SPEDContr))
    
End Sub

Public Sub CarregarParticipantesReferenciados(ByRef Registros As Variant, ByRef arrPartRef As ArrayList, Optional ByRef SPEDContr As Boolean)

Dim Registro As Variant
Dim cPart As String, CNPJ$
    
    For Each Registro In Registros
        
        Select Case VBA.Mid(Registro, 2, 4)
            
            Case "A010", "C010", "D010", "F010", "I010", "P010"
                CNPJ = fnSPED.ExtrairCampo(Registro, 2)
            
            Case "A100"
                cPart = fnSPED.ExtrairCampo(Registro, 4)
                
            Case "F100"
                If SPEDContr Then
                    cPart = fnSPED.ExtrairCampo(Registro, 3)
                    arrPartRef.Add cPart
                End If
                
            Case "C160", "C165", "D140", "D400", "E113", "E240", "E313", "E531", "1110", "1600", "1923"
                cPart = fnSPED.ExtrairCampo(Registro, 2)
                
            Case "B440", "G130"
                cPart = fnSPED.ExtrairCampo(Registro, 3)
                
            Case "B020", "C100", "C113", "C500", "D100", "D500", "D700", "1500"
                If SPEDContr And VBA.Mid(Registro, 2, 4) = "C500" Then
                    cPart = fnSPED.ExtrairCampo(Registro, 2)
                Else
                    cPart = fnSPED.ExtrairCampo(Registro, 4)
                End If
                
            Case "K200"
                cPart = fnSPED.ExtrairCampo(Registro, 6)
                
            Case "K280"
                cPart = fnSPED.ExtrairCampo(Registro, 7)
                
            Case "H010"
                cPart = fnSPED.ExtrairCampo(Registro, 8)
                
            Case "D510", "1510"
                cPart = fnSPED.ExtrairCampo(Registro, 18)
                
            Case "C510"
                cPart = fnSPED.ExtrairCampo(Registro, 18)
                
            Case "C176"
                arrPartRef.Add fnSPED.ExtrairCampo(Registro, 6)
                arrPartRef.Add fnSPED.ExtrairCampo(Registro, 21)
                
            Case "D130"
                arrPartRef.Add fnSPED.ExtrairCampo(Registro, 2)
                arrPartRef.Add fnSPED.ExtrairCampo(Registro, 3)
                
            Case "D170"
                arrPartRef.Add fnSPED.ExtrairCampo(Registro, 2)
                arrPartRef.Add fnSPED.ExtrairCampo(Registro, 3)
                
            Case "1601"
                arrPartRef.Add fnSPED.ExtrairCampo(Registro, 2)
                arrPartRef.Add fnSPED.ExtrairCampo(Registro, 3)
                
        End Select
        
        cPart = DefinirChave(CNPJ, cPart, SPEDContr)
        If (cPart <> "") And (Not arrPartRef.contains(cPart)) Then arrPartRef.Add CStr(cPart)
        cPart = ""
        
    Next Registro
    
End Sub

Public Sub CarregarUnidadesReferenciadas(ByRef Registros As Variant, ByRef arrProdRef As ArrayList, ByRef arrUndRef As ArrayList, Optional ByRef SPEDContr As Boolean)

Dim nCampo As Byte
Dim Registro As Variant
Dim nReg As String, CNPJ$, cProd$, cUnd$, COD_VER$
    
    For Each Registro In Registros
        
        nReg = VBA.Mid(Registro, 2, 4)
        Select Case nReg
            
            Case "0000"
                COD_VER = fnSPED.ExtrairCampo(Registro, 2)
            
            Case "0140"
                CNPJ = fnSPED.ExtrairCampo(Registro, 4)
                
            Case "0200"
                nCampo = 6
                cProd = fnSPED.ExtrairCampo(Registro, 2)
                
            Case "0220"
                nCampo = 2
                
            Case "A010", "C010", "D010", "F010", "I010", "P010"
                CNPJ = fnSPED.ExtrairCampo(Registro, 2)
                
            Case "H010"
                nCampo = 3
                
            'Nenhum dos registros abaixo possui o campo unidade de medida no SPED Contribuições
            Case "C180", "C181", "C321", "C330", "C380", "C425", "C430", "C480", "C815", "C870", "C880"
                If Not SPEDContr Then nCampo = 4
                
            Case "C370", "C470", "C510", "C610", "C810", "D610", "G140"
                nCampo = 5
                If nReg = "G140" And (COD_VER = "005" Or COD_VER = "006") Then nCampo = 0
                
            Case "C170", "C495", "D510", "1510"
                nCampo = 6
                
            Case "C185", "C186"
                nCampo = 8
                
            Case "1110"
                nCampo = 10
                
        End Select
        
        If nReg = "0200" Or nReg = "0220" Then
            
            'Define a chave da unidade referenciada para SPED Contribuições ou não
            If nCampo > 0 Then cUnd = DefinirChave(CNPJ, fnSPED.ExtrairCampo(Registro, nCampo), SPEDContr)
            If SPEDContr Then cProd = CNPJ & cProd
            If cUnd <> "" And arrProdRef.contains(cProd) And Not arrUndRef.contains(cUnd) Then arrUndRef.Add CStr(cUnd)
            
        'Apaga referencias ao código de produto caso já tenham passado todos os registro 0200 e filhos
        ElseIf nReg >= "0300" Then
            
            cProd = ""
            If nCampo > 0 Then cUnd = DefinirChave(CNPJ, fnSPED.ExtrairCampo(Registro, nCampo), SPEDContr)
            
            'Verifica situações relacionadas aos registros 0200 e filhos
            If cUnd <> "" And Not arrUndRef.contains(cUnd) Then arrUndRef.Add CStr(cUnd)
            
        End If
        
        'Apaga dados da unidade referenciada e número do campo
        cUnd = ""
        nCampo = 0
        
    Next Registro
    
End Sub

Public Sub CarregarProdutosReferenciados(ByRef Registros As Variant, ByRef arrProdRef As ArrayList, Optional ByRef SPEDContr As Boolean)

Dim nCampo As Byte
Dim Registro As Variant
Dim cProd As String, nReg$, CNPJ$
    
    For Each Registro In Registros
        
        nReg = VBA.Mid(Registro, 2, 4)
        Select Case True
            
            Case nReg = "0140"
                CNPJ = fnSPED.ExtrairCampo(Registro, 4)
                
            Case nReg = "0200"
                cProd = fnSPED.ExtrairCampo(Registro, 2)
                
            Case nReg = "0220"
                arrProdRef.Add cProd
                
            Case nReg = "A010", nReg = "C010", nReg = "D010", nReg = "F010", nReg = "I010", nReg = "P010"
                CNPJ = fnSPED.ExtrairCampo(Registro, 2)
                
            Case nReg = "A170"
                nCampo = 3
                
            Case nReg = "F100"
                nCampo = 4
                
            Case nReg = "0210", nReg = "0221", nReg = "C186", nReg = "C321", nReg = "C425", nReg = "C470", _
                nReg = "C870", nReg = "H010", nReg = "K215", nReg = "K265", nReg = "K275", nReg = "K291", _
                nReg = "K292", nReg = "K301", nReg = "K302", nReg = "1300", nReg = "1400"
                nCampo = 2
                
            Case nReg = "C170", nReg = "C185", nReg = "C370", nReg = "C495", nReg = "C510", nReg = "C610", _
                nReg = "C810", nReg = "D110", nReg = "D510", nReg = "D610", nReg = "G140", nReg = "K200", _
                nReg = "K235", nReg = "K250", nReg = "K260", nReg = "K280", nReg = "1370", nReg = "1510"
                nCampo = 3
                
            Case nReg = "C197", nReg = "C597", nReg = "C857", nReg = "C897", nReg = "D197", nReg = "D737"
                nCampo = 4
                
            Case nReg = "K210", nReg = "K230", nReg = "K270", nReg = "C180" And SPEDContr, nReg = "C190" And SPEDContr
                nCampo = 5
                
            Case nReg = "1105"
                nCampo = 7
                
            Case nReg = "E113", nReg = "E240", nReg = "E531", nReg = "1923"
                nCampo = 8
                
            Case nReg = "E313"
                nCampo = 9
                
            Case nReg = "1391"
                nCampo = 18
                
            Case nReg = "K220"
                arrProdRef.Add fnSPED.ExtrairCampo(Registro, 3)
                arrProdRef.Add fnSPED.ExtrairCampo(Registro, 4)
                
            Case nReg = "K255"
                arrProdRef.Add fnSPED.ExtrairCampo(Registro, 3)
                arrProdRef.Add fnSPED.ExtrairCampo(Registro, 5)
                
        End Select
        
        'Reinicia variável cProd
        If nReg > "0220" Then cProd = ""
        
        'Define a chave do produto referenciado para SPED Contribuições ou não
        If nCampo > 0 Then cProd = DefinirChave(CNPJ, fnSPED.ExtrairCampo(Registro, nCampo), SPEDContr)
        If cProd <> "" And nReg <> "0200" And Not arrProdRef.contains(cProd) Then arrProdRef.Add CStr(cProd)
        
        nCampo = 0
        
    Next Registro
    
End Sub

Public Sub CarregarBensReferenciados(ByRef Registros As Variant, ByRef arrBensRef As ArrayList, Optional ByRef SPEDContr As Boolean)

Dim Registro As Variant
Dim cBem As String
    
    For Each Registro In Registros
        
        Select Case VBA.Mid(Registro, 2, 4)
                                            
            Case "G125"
                cBem = fnSPED.ExtrairCampo(Registro, 2)
                arrBensRef.Add cBem
                
            Case "0300"
                cBem = fnSPED.ExtrairCampo(Registro, 5)
                arrBensRef.Add cBem
                
        End Select
        
    Next Registro

End Sub

Public Sub CarregarNaturezasReferenciadas(ByRef Registros As Variant, ByRef arrNatRef As ArrayList, Optional ByRef SPEDContr As Boolean)

Dim nCampo As Byte
Dim Registro As Variant
Dim nReg As String, cNat$, CNPJ$
    
    For Each Registro In Registros
        
        nReg = VBA.Mid(Registro, 2, 4)
        Select Case True
            
            Case nReg = "C010"
                CNPJ = fnSPED.ExtrairCampo(Registro, 2)
                
            Case nReg = "C170"
                nCampo = 12
                
        End Select
        
        'Define a chave da natureza referenciada para SPED Contribuições ou não
        If nCampo > 0 Then cNat = DefinirChave(CNPJ, fnSPED.ExtrairCampo(Registro, nCampo), SPEDContr)
        If cNat <> "" And Not arrNatRef.contains(cNat) Then arrNatRef.Add CStr(cNat)
        
        nCampo = 0
        
    Next Registro
    
End Sub

Public Sub CarregarInformacoesReferenciadas(ByRef Registros As Variant, ByRef arrInfRef As ArrayList, Optional ByRef SPEDContr As Boolean)

Dim nCampo As Byte
Dim Registro As Variant
Dim nReg As String, cInf$, CNPJ$
    
    For Each Registro In Registros
        
        nReg = VBA.Mid(Registro, 2, 4)
        Select Case True
                                
            Case nReg = "A010", nReg = "C010", nReg = "D010", nReg = "F010", nReg = "I010", nReg = "P010"
                CNPJ = fnSPED.ExtrairCampo(Registro, 2)
                
            Case nReg = "C110"
                nCampo = 2
                
            Case nReg = "D700"
                nCampo = 19
                                
            Case nReg = "D500"
                nCampo = 20
                
            Case nReg = "D100"
                nCampo = 22
                
            Case nReg = "C500" And Not SPEDContr Or (nReg = "1500" And Not SPEDContr)
                nCampo = 23
                
        End Select
        
        'Define a chave da informação referenciada para SPED Contribuições ou não
        If nCampo > 0 Then cInf = DefinirChave(CNPJ, fnSPED.ExtrairCampo(Registro, nCampo), SPEDContr)
        If cInf <> "" And Not arrInfRef.contains(cInf) Then arrInfRef.Add CStr(cInf)
        
        nCampo = 0
        
    Next Registro

End Sub

Public Sub CarregarObservacoesReferenciadas(ByRef Registros As Variant, ByRef arrObsRef As ArrayList, Optional ByRef SPEDContr As Boolean)

Dim Registro As Variant
Dim nReg As String, cObs$
    
    For Each Registro In Registros
        
        nReg = VBA.Mid(Registro, 2, 4)
        Select Case True
            
            Case (nReg = "C195") Or (nReg = "C595") Or (nReg = "C855") Or (nReg = "C895") Or (nReg = "D195") Or (nReg = "D735")
                cObs = fnSPED.ExtrairCampo(Registro, 2)
                arrObsRef.Add cObs
                
            Case (nReg = "B460")
                cObs = fnSPED.ExtrairCampo(Registro, 7)
                arrObsRef.Add cObs
                
            Case (nReg = "C490") Or (nReg = "C850") Or (nReg = "C890")
                cObs = fnSPED.ExtrairCampo(Registro, 8)
                arrObsRef.Add cObs
                
            Case (nReg = "C320") Or (nReg = "C390") Or (nReg = "D190") Or (nReg = "D730") Or (nReg = "D760")
                cObs = fnSPED.ExtrairCampo(Registro, 9)
                arrObsRef.Add cObs
                
            Case (nReg = "B350") Or (nReg = "C500") Or (nReg = "C690") Or (nReg = "C700") Or (nReg = "D390") Or (nReg = "D590") Or (nReg = "D690") Or (nReg = "D696")
                cObs = fnSPED.ExtrairCampo(Registro, 11)
                arrObsRef.Add cObs
                
            Case (nReg = "B030") Or (nReg = "C190" And Not SPEDContr)
                cObs = fnSPED.ExtrairCampo(Registro, 12)
                arrObsRef.Add cObs
                
            Case (nReg = "D300")
                cObs = fnSPED.ExtrairCampo(Registro, 19)
                arrObsRef.Add cObs
                
            Case (nReg = "B020")
                cObs = fnSPED.ExtrairCampo(Registro, 21)
                arrObsRef.Add cObs
                
        End Select
        
    Next Registro
    
End Sub

Private Function DefinirChave(ByVal CNPJ As String, ByVal Chave As String, Optional ByRef SPEDContr As Boolean) As String
    
    'Verifica se o SPED é o Contribuições e define a chave com base nisso
    If SPEDContr Then DefinirChave = CNPJ & Chave Else DefinirChave = Chave

End Function
