Attribute VB_Name = "clsCriptografia_MD5"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Function MD5(ByVal Texto As String) As String
        
Dim oT As Object, oMD5 As Object
Dim TextToHash() As Byte
Dim bytes() As Byte
     
     Set oT = CreateObject("System.Text.UTF8Encoding")
     Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
     
     TextToHash = oT.GetBytes_4(Texto)
     bytes = oMD5.ComputeHash_2((TextToHash))
     
     MD5 = VBA.UCase(Criptografar(bytes))
     
     Set oT = Nothing
     Set oMD5 = Nothing
     
End Function

Private Function Criptografar(Valor As Variant) As Variant

Dim oD As New DOMDocument60

      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = Valor
      End With
    Criptografar = Replace(oD.DocumentElement.text, vbLf, "")
    
    Set oD = Nothing

End Function

