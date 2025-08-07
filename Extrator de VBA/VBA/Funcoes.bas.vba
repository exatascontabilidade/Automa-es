Attribute VB_Name = "Funcoes"
Option Explicit
Option Base 1

Public Function IndentificarSubPastas(ByVal Caminho As String, ByRef ListaXMLS As ArrayList)
        
Dim fso As New FileSystemObject
Dim pasta As Folder
Dim ARQUIVO As file
Dim i As Long

    Set pasta = fso.GetFolder(Caminho)
    
    For Each pasta In pasta.SubFolders
    
        For Each ARQUIVO In pasta.Files

            If InStr(VBA.LCase(ARQUIVO.Path), ".xml") > 0 Then ListaXMLS.Add ARQUIVO.Path
            
        Next ARQUIVO
        
        Call IndentificarSubPastas(pasta.Path, ListaXMLS)
        
    Next pasta
    
End Function

Public Function FormatarPercentuais(ByVal Valor As String) As Double
    
    If Valor = "" Or Valor = "-" Then Valor = 0
    FormatarPercentuais = Replace(Valor, ".", ",") / 100
    
End Function

Public Function ValidarTag(ByRef Node As IXMLDOMNode, ByVal Tag As String) As String
    If Not Node.SelectSingleNode(Tag) Is Nothing Then ValidarTag = Node.SelectSingleNode(Tag).text
End Function

Public Function SelecionarArquivos(ByVal Extensao As String)
    SelecionarArquivos = Application.GetOpenFilename("Arquivos " & Extensao & " (*." & Extensao & "), *." & Extensao, , "Selecione os arquivos " & Extensao & " que deseja importar", , True)
End Function

Public Function CarregarDadosContribuinte() As Boolean
    
    CNPJContribuinte = Util.FormatarCNPJ(CadContrib.Range("CNPJContribuinte").value)
    
    If CNPJContribuinte <> "" Then
        
        CNPJBase = VBA.Left(CNPJContribuinte, 8)
        InscContribuinte = VBA.Replace(VBA.Replace(VBA.Replace(CadContrib.Range("InscContribuinte").value, ".", ""), "/", ""), "-", "")
        UFContribuinte = VBA.Trim(CadContrib.Range("UFContribuinte").value)
        RazaoContribuinte = CadContrib.Range("RazaoContribuinte").value
        CarregarDadosContribuinte = True
        
    Else
        
        MsgBox "O campo CNPJ da tela de cadastro do contribuinte não foi preenchido!" & vbCrLf & vbCrLf & _
               "Por favor faça o cadastro do contribuinte para continuar.", vbCritical, "Contribuinte não Cadastrado"
                
        CadContrib.Activate
        CadContrib.Range("CNPJContribuinte").Activate
        
    End If
    
End Function

Public Function ValidarUsuario() As Boolean
    If (VBA.Environ("COMPUTERNAME") <> "MVHAMORIM") And (InStr(1, UCase(EmailAssinante), "ESCOLADAAUTOMACAOFISCAL.COM.BR") = 0) Then ValidarUsuario = True
    If (VBA.Environ("COMPUTERNAME") = "MVHAMORIM") Then ValidarUsuario = True
End Function

Function ValidarEmail(EMAIL As String) As Boolean

Dim regex As Object
    
    EMAIL = Util.RemoverCaracteresNaoImprimiveis(EMAIL)
    Set regex = CreateObject("vbscript.regexp")
    
    With regex
        .Pattern = "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
        ValidarEmail = .Test(EMAIL)
    End With
    
End Function

Sub MontarCamposRegistros()

Dim Planilha As New ArrayList
Dim Linhas As Variant, Linha
Dim Texto As String, nCampo
Dim MyDataObj As Object
Dim Registro As String

    Set MyDataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

    MyDataObj.GetFromClipboard
    Texto = MyDataObj.getText
    
    Linhas = Split(Texto, vbLf)
    For Each Linha In Linhas
        
        If Registro = "" Then Registro = "(" & Trim(VBA.Left(Linha, InStr(1, Linha, " =") - 1)) Else Registro = Registro & ", " & Trim(VBA.Left(Linha, InStr(1, Linha, " =") - 1))
        
    Next Linha
    
    Registro = Registro & ")"
    
End Sub

Public Function InserirColunaCHV_REG()

Dim Plan As Worksheet
    
    For Each Plan In ThisWorkbook.Worksheets
        
        If VBA.Left(Plan.CodeName, 4) = "regK" Then
            
            If Plan.Range("C3").value <> "CHV_REG" Then
                
                Plan.Columns("C:C").Insert Shift:=xlToRight
                Plan.Range("C3").value = "CHV_REG"
                
            End If
            
        End If
        
    Next Plan
    
End Function

Public Function InserirColunaCHV_PAI()

Dim Plan As Worksheet
    
    For Each Plan In ThisWorkbook.Worksheets
        
        If VBA.Left(Plan.CodeName, 4) = "regK" Then
            
            If Plan.Range("D3").value <> "CHV_PAI_FISCAL" Then
                
                Plan.Columns("D:D").Insert Shift:=xlToRight
                Plan.Range("D3").value = "CHV_PAI_FISCAL"
                
            End If
            
        End If
        
    Next Plan
    
    Call Redimensionarcolunas
    
End Function

Public Function CongelarPaineis()

Dim Plan As Worksheet
    
    For Each Plan In ThisWorkbook.Worksheets
        
        If VBA.Left(Plan.CodeName, 3) = "reg" Then
            
            Plan.Activate
            Plan.Range("B4").Activate
            ActiveWindow.FreezePanes = False
            ActiveWindow.FreezePanes = True
            
        End If
        
    Next Plan
    
End Function

Public Function Redimensionarcolunas()

Dim Plan As Worksheet
    
    For Each Plan In ThisWorkbook.Worksheets
        
        If VBA.Left(Plan.CodeName, 3) = "reg" Then Plan.Columns.AutoFit
        
    Next Plan
    
End Function

Sub FormatarTipoColunas()

Dim Plan As Worksheet
Dim Cel As Range
    
    For Each Plan In ThisWorkbook.Worksheets
        
        If VBA.Left(Plan.CodeName, 3) = "reg" Then
            
            Set Cel = Plan.Range("A3")
            Do Until Cel.value = ""
                
                With Cel.EntireColumn
                    
                    Select Case True
                        
                        Case Cel.value Like "DT_*"
                            .NumberFormat = "m/d/yyyy"
                            
                        Case Cel.value Like "VL_*" Or Cel.value Like "QTD_*"
                            .Style = "Comma"
                            Cel.Offset(-1).Formula2R1C1 = "=SUBTOTAL(9,INDIRECT(SUBSTITUTE(ADDRESS(1, COLUMN(), 4), ""1"", """") & ""4:"" & SUBSTITUTE(ADDRESS(1, COLUMN(), 4), ""1"", """") & ""1048576""))"
                            
                        Case Cel.value Like "ALIQ_*"
                            .NumberFormat = "0.00%"
                            
                        Case Else
                            .NumberFormat = "@"
                            
                    End Select
                    
                End With
                Set Cel = Cel.Offset(, 1)
                
            Loop
            
        End If
        
    Next Plan
    
    MsgBox "Terminou!"
    
End Sub

Public Function ListarContribuintes(ByRef dicContribuintes As Dictionary)
    Set dicContribuintes = Util.CriarDicionarioRegistro(CadContrib, "CNPJ")
End Function
