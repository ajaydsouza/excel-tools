Option Explicit

Sub Auto_Open()
    Dim cControl As CommandBarControl
    
    Auto_Close    'Prevents duplicate entry of the menu item
    
'    Set cControl = Application.CommandBars(2).FindControl(ID:=30007).Controls.Add _
'    (Type:=msoControlPopup, temporary:=True)
    
    Set cControl = Application.CommandBars(2).FindControl(ID:=30007)
    
    With cControl
        .Caption = "&WZ Macros"   'name the item
        
        With .Controls.Add(Type:=msoControlButton, temporary:=True)  'adds a dropdown button'
            .DescriptionText = "Clears styles that are not built-in"
            .Caption = "&Clear styles not built-in"
            .FaceId = 417
            .OnAction = "StyleKill"
        End With
    
        With .Controls.Add(Type:=msoControlButton, temporary:=True)  'adds a dropdown button'
            .DescriptionText = "Clears styles except Normal"
            .Caption = "&Clear all styles"
            .FaceId = 330
            .OnAction = "ClearStyles"
        End With
    
        With .Controls.Add(Type:=msoControlButton, temporary:=True)  'adds a dropdown button'
            .DescriptionText = "Delete all conditional formatting"
            .Caption = "&Clear Conditional Formatting"
            .FaceId = 417
            .OnAction = "RemoveConditionalFormatting"
        End With
    
        With .Controls.Add(Type:=msoControlButton, temporary:=True)  'adds a dropdown button'
            .Caption = "&Delete Unused Number Formats"
            .FaceId = 1555
            .OnAction = "DeleteUnusedCustomNumberFormats"
        End With
    
        With .Controls.Add(Type:=msoControlButton, temporary:=True)  'adds a dropdown button'
            .Caption = "&Delete Dead Names"
            .FaceId = 478
            .OnAction = "DeleteDeadNames"
        End With
    
        With .Controls.Add(Type:=msoControlButton, temporary:=True)  'adds a dropdown button'
            .Caption = "&Delete External Names"
            .FaceId = 323
            .OnAction = "DeleteExtNames"
        End With
    
        With .Controls.Add(Type:=msoControlButton, temporary:=True)  'adds a dropdown button'
            .Caption = "&Delete Chosen Names"
            .FaceId = 536
            .OnAction = "DeleteChosenNames"
        End With
    
        With .Controls.Add(Type:=msoControlButton, temporary:=True)  'adds a dropdown button'
            .Caption = "&Create WorkSheet Index"
            .FaceId = 637
            .OnAction = "CreateIndex"
        End With
        
        With .Controls.Add(Type:=msoControlButton, temporary:=True)
            .Caption = "&Copy Names Across WBs"
            .FaceId = 2045
            .OnAction = "Copy_All_Defined_Names"
        End With
    
        With .Controls.Add(Type:=msoControlButton, temporary:=True)
            .Caption = "&Combine workbooks"
            .FaceId = 263
            .OnAction = "CombineWorkbooks"
        End With
    
        With .Controls.Add(Type:=msoControlButton, temporary:=True)
            .Caption = "&Sort Sheets"
            .FaceId = 3650
            .OnAction = "Sort_Active_Book"
        End With
    
        With .Controls.Add(Type:=msoControlButton, temporary:=True)  'adds a dropdown button'
            .Caption = "&Reset Comment Position"
            .FaceId = 1546
            .OnAction = "ResetCommentsPosition"
        End With
    
        With .Controls.Add(Type:=msoControlButton, temporary:=True)  'adds a dropdown button'
            .Caption = "&Reset Comment Size"
            .FaceId = 1758
            .OnAction = "CommentsAutoSize"
        End With
    
End With
End Sub

Sub Auto_Close()
    On Error Resume Next
    Application.CommandBars(2).FindControl(ID:=30007).Controls("&WZ Macros").Delete
    Application.CommandBars(2).FindControl(ID:=30007).Controls("&Clear styles not built-in").Delete
    Application.CommandBars(2).FindControl(ID:=30007).Controls("&Clear all styles").Delete
    Application.CommandBars(2).FindControl(ID:=30007).Controls("&Clear Conditional Formatting").Delete
    Application.CommandBars(2).FindControl(ID:=30007).Controls("&Delete Unused Number Formats").Delete
    Application.CommandBars(2).FindControl(ID:=30007).Controls("&Delete Dead Names").Delete
    Application.CommandBars(2).FindControl(ID:=30007).Controls("&Delete External Names").Delete
    Application.CommandBars(2).FindControl(ID:=30007).Controls("&Delete Chosen Names").Delete
    Application.CommandBars(2).FindControl(ID:=30007).Controls("&Create WorkSheet Index").Delete
    Application.CommandBars(2).FindControl(ID:=30007).Controls("&Copy Names Across WBs").Delete
    Application.CommandBars(2).FindControl(ID:=30007).Controls("&Combine workbooks").Delete
    Application.CommandBars(2).FindControl(ID:=30007).Controls("&Sort Sheets").Delete
    Application.CommandBars(2).FindControl(ID:=30007).Controls("&Reset Comment Position").Delete
    Application.CommandBars(2).FindControl(ID:=30007).Controls("&Reset Comment Size").Delete
    
    On Error GoTo 0
End Sub


