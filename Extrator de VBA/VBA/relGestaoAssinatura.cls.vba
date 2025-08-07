Attribute VB_Name = "relGestaoAssinatura"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)
    
    If Target.AddressLocal = "$B$2" Then
        
        With Target
            .Hyperlinks.Delete
            .Font.name = "Calibri"
            .Interior.Color = RGB(255, 255, 204)
        End With
        
    End If
    
End Sub
