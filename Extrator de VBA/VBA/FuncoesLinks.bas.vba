Attribute VB_Name = "FuncoesLinks"
Option Explicit

#If VBA7 Then
    ' Para o Office 2010 ou posterior em sistemas de 64 bits
    Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As LongPtr, ByVal lpszOp As String, _
        ByVal lpszFile As String, ByVal lpszParams As String, _
        ByVal LpszDir As String, ByVal FsShowCmd As Long) As LongPtr
#Else
    ' Para o Office 2007 ou anterior em sistemas de 32 bits
    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, ByVal lpszOp As String, _
        ByVal lpszFile As String, ByVal lpszParams As String, _
        ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long
#End If

Sub AbrirUrl(URL As String)

    #If VBA7 Then
        ShellExecute 0, "open", URL, "", "", 1
    #Else
        ShellExecute 0, "open", URL, vbNullString, vbNullString, 1
    #End If

End Sub


