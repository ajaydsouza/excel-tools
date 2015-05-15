Option Explicit

Sub Auto_Open()
    Dim cControl As CommandBarControl
    
    Auto_Close    'Prevents duplicate entry of the menu item
    
    Set cControl = Application.CommandBars(1).FindControl(ID:=30007).Controls.Add _
    (Type:=msoControlPopup, temporary:=True)
    
    With cControl
        .Caption = "&WZ Macros"   'name the item
        
        With .Controls.Add(Type:=msoControlButton)  'adds a dropdown button'
            .DescriptionText = "Clears styles that are not built-in"
            .Caption = "&Clear styles not built-in"
            .OnAction = "StyleKill"
        End With
    
        With .Controls.Add(Type:=msoControlButton)  'adds a dropdown button'
            .DescriptionText = "Clears styles except Normal"
            .Caption = "&Clear all styles"
            .OnAction = "ClearStyles"
        End With
    
        With .Controls.Add(Type:=msoControlButton)  'adds a dropdown button'
            .Caption = "&Delete Unused Number Formats"
            .OnAction = "DeleteUnusedCustomNumberFormats"
        End With
    
        With .Controls.Add(Type:=msoControlButton)  'adds a dropdown button'
            .Caption = "&Delete Dead Names"
            .OnAction = "DeleteDeadNames"
        End With
    
        With .Controls.Add(Type:=msoControlButton)  'adds a dropdown button'
            .Caption = "&Delete External Names"
            .OnAction = "DeleteExtNames"
        End With
    
        With .Controls.Add(Type:=msoControlButton)  'adds a dropdown button'
            .Caption = "&Delete Chosen Names"
            .OnAction = "DeleteChosenNames"
        End With
    
        With .Controls.Add(Type:=msoControlButton)  'adds a dropdown button'
            .Caption = "&Create WorkSheet Index"
            .OnAction = "CreateIndex"
        End With
    
    End With
End Sub

Sub Auto_Close()
    On Error Resume Next
    Application.CommandBars(1).FindControl(ID:=30007).Controls("&WZ Macros").Delete
    On Error GoTo 0
End Sub

