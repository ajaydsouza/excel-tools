Option Explicit

Sub Auto_Open()
    Dim cControl As CommandBarControl
    
    Auto_Close    'Prevents duplicate entry of the menu item
    
'    Set cControl = Application.CommandBars(1).FindControl(ID:=30007).Controls.Add _
'    (Type:=msoControlButton, temporary:=True)
    
    Set cControl = Application.CommandBars(1).FindControl(ID:=30007)
    
    With cControl
        .Caption = "&WZ Macros"   'name the item
        
        With .Controls.Add(Type:=msoControlButton)  'adds a dropdown button'
            .DescriptionText = "Clears styles that are not built-in"
            .Caption = "&Clear styles not built-in"
            .FaceId = 417
            .OnAction = "StyleKill"
        End With
    
        With .Controls.Add(Type:=msoControlButton)  'adds a dropdown button'
            .DescriptionText = "Clears styles except Normal"
            .Caption = "&Clear all styles"
            .FaceId = 330
            .OnAction = "ClearStyles"
        End With
    
        With .Controls.Add(Type:=msoControlButton)  'adds a dropdown button'
            .Caption = "&Delete Unused Number Formats"
            .FaceId = 1555
            .OnAction = "DeleteUnusedCustomNumberFormats"
        End With
    
        With .Controls.Add(Type:=msoControlButton)  'adds a dropdown button'
            .Caption = "&Delete Dead Names"
            .FaceId = 478
            .OnAction = "DeleteDeadNames"
        End With
    
        With .Controls.Add(Type:=msoControlButton)  'adds a dropdown button'
            .Caption = "&Delete External Names"
            .FaceId = 323
            .OnAction = "DeleteExtNames"
        End With
    
        With .Controls.Add(Type:=msoControlButton)  'adds a dropdown button'
            .Caption = "&Delete Chosen Names"
            .FaceId = 536
            .OnAction = "DeleteChosenNames"
        End With
    
        With .Controls.Add(Type:=msoControlButton)  'adds a dropdown button'
            .Caption = "&Create WorkSheet Index"
            .FaceId = 637
            .OnAction = "CreateIndex"
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "&Copy Names Across WBs"
            .FaceId = 263
            .OnAction = "Copy_All_Defined_Names"
        End With
    
    End With
End Sub

Sub Auto_Close()
    On Error Resume Next
    Application.CommandBars(1).FindControl(ID:=30007).Controls("&WZ Macros").Delete
    Application.CommandBars(1).FindControl(ID:=30007).Controls("&Clear styles not built-in").Delete
    Application.CommandBars(1).FindControl(ID:=30007).Controls("&Clear all styles").Delete
    Application.CommandBars(1).FindControl(ID:=30007).Controls("&Delete Unused Number Formats").Delete
    Application.CommandBars(1).FindControl(ID:=30007).Controls("&Delete Dead Names").Delete
    Application.CommandBars(1).FindControl(ID:=30007).Controls("&Delete External Names").Delete
    Application.CommandBars(1).FindControl(ID:=30007).Controls("&Delete Chosen Names").Delete
    Application.CommandBars(1).FindControl(ID:=30007).Controls("&Create WorkSheet Index").Delete
    Application.CommandBars(1).FindControl(ID:=30007).Controls("&Copy Names Across WBs").Delete
    
    
    On Error GoTo 0
End Sub


