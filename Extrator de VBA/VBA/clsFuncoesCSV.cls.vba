Attribute VB_Name = "clsFuncoesCSV"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Sub ImportarCSV(ByRef Arqs As Variant, ByRef arrChaves As ArrayList, ByRef dicEntradas As Dictionary, ByRef dicSaidas As Dictionary)

Dim Arq, Registros, Campos

On Error GoTo Tratar:

    For Each Arq In Arqs
    
        Open Arq For Input As #1
            Registros = Split(VBA.Input(LOF(1), 1), vbCrLf)
        Close #1
        
        Call ExtrairDadosCSV(Registros, arrChaves, dicEntradas, dicSaidas)
                
    Next Arq
            
Exit Sub
Tratar:
    
    If Err.Number <> 0 Then Call TratarErros(Err, "FuncoesCSV.ImportarCSV")

End Sub

Private Sub ExtrairDadosCSV(ByVal Registros, ByRef arrChaves As ArrayList, ByRef dicEntradas As Dictionary, ByRef dicSaidas As Dictionary)

Dim Registro, Campos
Dim dicTitulos As New Dictionary

On Error GoTo Tratar:
    
    For Each Registro In Registros
        
        If Registro <> "" Then
            
            Campos = Split(Registro, ";")
            
            If Not Util.ChecarCamposPreenchidos(Campos) Then GoTo Prx:
            
            If Registro Like "*Numero NF-e*" Then
            
                Set dicTitulos = MapearIndicesCabecalho(Registro)
            
            Else
                
                If Campos(2) Like "*&amp*" And Not VBA.IsDate(Campos(3)) Then Campos = UnificarCampos(Campos)
                                
                With DadosDoce
                    
                    Call ExtrairDadosRegistro(Registro, dicTitulos)
                    If Not arrChaves.contains(Replace(.chNFe, "'", "")) Then
                        
                        Call Util.GerarObservacoes(.Status, .CNPJEmit, .UF, .tpNF)
                        
                        Registro = Array(.nNF, .CNPJPart, .RazaoPart, .dtEmi, CDbl(.vNF), .chNFe, .UF, .Status, .tpNF, .StatusSPED, .DivergNF, .OBSERVACOES)
                        Call Util.ClassificarNotaFiscal(.CNPJEmit, .tpNF, .Modelo, Registro, arrChaves, dicEntradas, dicSaidas)
                        
                    End If
                    
                End With
                
            End If
            
        End If
Prx:
    Next Registro
    
Exit Sub
Tratar:
    Stop
    Resume
    Call Util.MsgAlerta("Houve um erro inesperado ao importar os dados do CSV. Por favor entre em contato com o suporte.")
    If Err.Number <> 0 Then Call TratarErros(Err, "FuncoesCSV.ExtrairDadosCSV")
    
End Sub

Private Function MapearIndicesCabecalho(ByVal Cabecalho As String) As Object

Dim dicIndices As New Dictionary
Dim Campos() As String
Dim i As Integer
    
    Campos = Split(Cabecalho, ";")
    
    For i = LBound(Campos) To UBound(Campos)
        dicIndices(Trim(Campos(i))) = i
    Next
    
    Set MapearIndicesCabecalho = dicIndices
    
End Function

Private Sub ExtrairDadosRegistro(ByVal Registro As String, ByVal dicIndices As Dictionary)

Dim Campos() As String
Dim chValor As String

    Campos = Split(Registro, ";")
    If Not dicIndices.Exists("Valor (R$)") Then chValor = "Valor" Else chValor = "Valor (R$)"
    
    With DadosDoce
        
        .Modelo = "55"
        .nNF = CLng(VBA.Trim(Replace(Campos(dicIndices("Numero NF-e")), ".", "")))
        .chNFe = VBA.Trim(Campos(dicIndices("Chave de Acesso")))
        .dtEmi = Util.FormatarData(VBA.Trim(Campos(dicIndices("Data de Emissao"))))
        .vNF = VBA.Trim(Campos(dicIndices(chValor)))
        .CNPJEmit = VBA.Mid(.chNFe, 8, 14)
        .Status = VBA.Trim(Campos(dicIndices("Situacao")))
        .tpNF = VBA.Trim(Campos(dicIndices("Tipo Operacao")))
        .StatusSPED = ""
        .DivergNF = ""
        .OBSERVACOES = ""
        
        Select Case True
            
            Case dicIndices.Exists("CNPJ/CPF Destinatario")
                .CNPJPart = VBA.Trim(Campos(dicIndices("CNPJ/CPF Destinatario")))
                .RazaoPart = VBA.Trim(Campos(dicIndices("Razao Social Destinatario")))
                .UF = VBA.Trim(Campos(dicIndices("UF Dest.")))
                
            Case dicIndices.Exists("CNPJ/CPF Emitente")
                .CNPJPart = VBA.Trim(Campos(dicIndices("CNPJ/CPF Emitente")))
                .RazaoPart = VBA.Trim(Campos(dicIndices("Razao Social Emitente")))
                .UF = VBA.Trim(Campos(dicIndices("UF Emit.")))
                
        End Select
        
    End With
    
End Sub

Public Function UnificarCampos(ByRef Campos As Variant) As Variant

Dim i As Long
Dim arrCampos As New ArrayList
    
    For i = 0 To UBound(Campos)
        
        Select Case i
            
            Case Is = 2
                arrCampos.Add VBA.Replace(Campos(2) & Campos(3), "&amp", "&")
                
            Case Is = 3
            Case Else
                arrCampos.Add Campos(i)
                
        End Select
        
        UnificarCampos = arrCampos.toArray()
        
    Next i
    
End Function
