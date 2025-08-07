Attribute VB_Name = "ImportadorPlanilhas"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function AbrirPastaTrabalho() As Workbook

Dim Caminho As Variant
Dim wb As Workbook
    
    Caminho = Util.SelecionarArquivo("xlsx")
    If VarType(Caminho) = vbBoolean And Caminho = False Then Exit Function
    
    On Error GoTo ErroAbrir
        
        Set wb = Workbooks.Open(Caminho, ReadOnly:=True)
        wb.Windows(1).visible = False
        Set AbrirPastaTrabalho = wb
        
    On Error GoTo 0
    
    Exit Function
    
ErroAbrir:
Dim Titulo As String
Dim Mensagem As String
    
    Titulo = "Erro ao Abrir Arquivo"
    Mensagem = "Não foi possível abrir a planilha arquivo:" & vbCrLf & Caminho & vbCrLf & "Erro: " & Err.Description
    
    Call Util.MsgAlerta(Mensagem, Titulo)
    Set AbrirPastaTrabalho = Nothing
    
End Function
