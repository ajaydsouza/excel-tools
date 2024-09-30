Option Explicit

'Deletes All Styles (Except BuiltIn ones) From Active Workbook
Sub StyleKill()
    Dim styT As Style
    Dim i As Integer
    Dim oldStatusBar As String
    Dim intRet As Integer

    i = 0

    oldStatusBar = Application.DisplayStatusBar
    
    For Each styT In ActiveWorkbook.Styles
        If (Not styT.BuiltIn And Not (styT.Name = "")) Then
            On Error Resume Next
            styT.Locked = False
            styT.Delete
'           intRet = MsgBox("Delete style '" & styT.Name & "'?", vbYesNoCancel)
'           If intRet = vbYes Then styT.Delete
'           If intRet = vbCancel Then Exit Sub
            Application.StatusBar = "Deleting Dead Styles... #" & i
            i = i + 1
        End If
    Next styT
    
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    
    MsgBox (i & " dead styles have been cleaned")
End Sub

'-----------------------------------------------------------------

Sub DeleteDeadNames()

    Dim nName As Name
    Dim i As Integer
    Dim oldStatusBar As String
    i = 0

    oldStatusBar = Application.DisplayStatusBar
    
    For Each nName In ActiveWorkbook.Names

        If InStr(1, nName.RefersTo, "#REF!") > 0 Then
            nName.Delete
            Application.StatusBar = "Deleting Dead Names... #" & i
            i = i + 1
        End If

    Next nName

    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    
    MsgBox (i & " dead Names have been deleted")

End Sub
'-----------------------------------------------------------------
Sub DeleteChosenNames()

    Dim nName As Name
    Dim i As Integer
    Dim oldStatusBar As String
    i = 0
    
    Dim lReply As Long

    For Each nName In ActiveWorkbook.Names

        lReply = MsgBox("Delete the named range " & nName.Name & vbNewLine & "It refers to: " & nName.RefersTo, vbQuestion + vbYesNoCancel, "ajaydsouza.com")

        If lReply = vbCancel Then Exit Sub

        If lReply = vbYes Then
            nName.Delete
            i = i + 1
        End If

    Next nName

    MsgBox (i & " chosen Names have been deleted")

End Sub

'-----------------------------------------------------------------
Sub DeleteExtNames()

    Dim nName As Name
    Dim i As Integer
    Dim oldStatusBar As String
    i = 0

    oldStatusBar = Application.DisplayStatusBar
    
    For Each nName In Names
        If ((InStr(1, nName.RefersTo, "]") > 0) Or (InStr(1, nName.RefersTo, "\\") > 0) Or (InStr(1, nName.RefersTo, "#N/A") > 0)) Then
            nName.Delete
            Application.StatusBar = "Deleting External Names... #" & i
            i = i + 1
        End If
    Next nName
    
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    
    MsgBox (i & " External Names have been deleted")

End Sub
'-----------------------------------------------------------------
 
'Deletes All Styles (Except Normal) From Active Workbook
Sub ClearStyles()
    Dim i&, Cell As Range, RangeOfStyles As Range
    Dim j As Integer
    Dim oldStatusBar As String

    oldStatusBar = Application.DisplayStatusBar
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
     'Add a temporary sheet
    Sheets.Add Before:=Sheets(1)
     'List all the styles
    For i = 1 To ActiveWorkbook.Styles.Count
        [a65536].End(xlUp).Offset(1, 0) = ActiveWorkbook. _
        Styles(i).Name
    Next
    Set RangeOfStyles = Range(Columns(1).Rows(2), _
    Columns(1).Rows(65536).End(xlUp))
    For Each Cell In RangeOfStyles
        If Not Cell.Text Like "Normal" Then
            On Error Resume Next
            ActiveWorkbook.Styles(Cell.Text).Delete
            ActiveWorkbook.Styles(Cell.NumberFormat).Delete
            Application.StatusBar = "Deleting Style... #" & j
            j = j + 1
        End If
    Next Cell
     'delete the temp sheet
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    
    MsgBox (j & " styles have been deleted")

End Sub

'-----------------------------------------------------------------

Sub DeleteUnusedCustomNumberFormats()
    Dim Buffer As Object
    Dim Sh As Object
    Dim SaveFormat As Variant
    Dim fFormat As Variant
    Dim nFormat() As Variant
    Dim xFormat As Long
    Dim Counter As Long
    Dim Counter1 As Long
    Dim Counter2 As Long
    Dim StartRow As Long
    Dim EndRow As Long
    Dim Dummy As Variant
    Dim pPresent As Boolean
    Dim NumberOfFormats As Long
    Dim Answer
    Dim c As Object
    Dim DataStart As Long
    Dim DataEnd As Long
    Dim AnswerText As String

    NumberOfFormats = 1000
    ReDim nFormat(0 To NumberOfFormats)
    AnswerText = "Do you want to delete unused custom formats from the workbook?"
    AnswerText = AnswerText & Chr(10) & "To get a list of used and unused formats only, choose No."
    Answer = MsgBox(AnswerText, 259)
    If Answer = vbCancel Then GoTo Finito

    On Error GoTo Finito
    Worksheets.Add.Move After:=Worksheets(Worksheets.Count)
    Worksheets(Worksheets.Count).Name = "CustomFormats"
    Worksheets("CustomFormats").Activate
    Set Buffer = Range("A2")
    Buffer.Select
    nFormat(0) = Buffer.NumberFormatLocal
    Counter = 1
    Do
        SaveFormat = Buffer.NumberFormatLocal
        Dummy = Buffer.NumberFormatLocal
        DoEvents
        SendKeys "{tab 3}{down}{enter}"
        Application.Dialogs(xlDialogFormatNumber).Show Dummy
        nFormat(Counter) = Buffer.NumberFormatLocal
        Counter = Counter + 1
    Loop Until nFormat(Counter - 1) = SaveFormat

    ReDim Preserve nFormat(0 To Counter - 2)

    Range("A1").Value = "Custom formats"
    Range("B1").Value = "Formats used in workbook"
    Range("C1").Value = "Formats not used"
    Range("A1:C1").Font.Bold = True

    StartRow = 3
    EndRow = 16384

    For Counter = 0 To UBound(nFormat)
        Cells(StartRow, 1).Offset(Counter, 0).NumberFormatLocal = nFormat(Counter)
        Cells(StartRow, 1).Offset(Counter, 0).Value = nFormat(Counter)
    Next Counter

    Counter = 0
    For Each Sh In ActiveWorkbook.Worksheets
        If Sh.Name = "CustomFormats" Then Exit For
        For Each c In Sh.UsedRange.Cells
            fFormat = c.NumberFormatLocal
            If Application.WorksheetFunction.CountIf(Range(Cells(StartRow, 2), Cells(EndRow, 2)), fFormat) = 0 Then
                Cells(StartRow, 2).Offset(Counter, 0).NumberFormatLocal = fFormat
                Cells(StartRow, 2).Offset(Counter, 0).Value = fFormat
                Counter = Counter + 1
            End If
        Next c
    Next Sh

    xFormat = Range(Cells(StartRow, 2), Cells(EndRow, 2)).Find("").Row - 2
    Counter2 = 0
    For Counter = 0 To UBound(nFormat)
        pPresent = False
        For Counter1 = 1 To xFormat
            If nFormat(Counter) = Cells(StartRow, 2).Offset(Counter1, 0).NumberFormatLocal Then
                pPresent = True
            End If
        Next Counter1
        If pPresent = False Then
            Cells(StartRow, 3).Offset(Counter2, 0).NumberFormatLocal = nFormat(Counter)
            Cells(StartRow, 3).Offset(Counter2, 0).Value = nFormat(Counter)
            Counter2 = Counter2 + 1
        End If
    Next Counter
    With ActiveSheet.Columns("A:C")
        .AutoFit
        .HorizontalAlignment = xlLeft
    End With
    If Answer = vbYes Then
        DataStart = Range(Cells(1, 3), Cells(EndRow, 3)).Find("").Row + 1
        DataEnd = Cells(DataStart, 3).Resize(EndRow, 1).Find("").Row - 1
        On Error Resume Next
        For Each c In Range(Cells(DataStart, 3), Cells(DataEnd, 3)).Cells
            ActiveWorkbook.DeleteNumberFormat (c.NumberFormat)
        Next c
    End If
Finito:
    Set c = Nothing
    Set Sh = Nothing
    Set Buffer = Nothing


    MsgBox "Unused number formats have been cleaned"

End Sub


'-----------------------------------------------------------------

Sub CreateIndex()

    Dim wSheet As Worksheet
    Dim l As Long
    
    l = 1
    Worksheets.Add(Before:=Worksheets(1)).Name = "Index"
    
    With ActiveWorkbook.ActiveSheet

        .Columns(1).ClearContents

        .Cells(1, 1) = "INDEX"

        .Cells(1, 1).Name = "Index"

    End With

    For Each wSheet In Worksheets

        If wSheet.Name <> ActiveSheet.Name Then
                l = l + 1
'                With wSheet
'                    .Range("A1").Name = "Start_" & wSheet.Index
'                    .Hyperlinks.Add Anchor:=.Range("A1"), Address:="", SubAddress:="Index", TextToDisplay:="Back to Index"
'                End With

                ActiveSheet.Hyperlinks.Add Anchor:=ActiveSheet.Cells(l, 1), Address:="", SubAddress:="Start_" & wSheet.Index, TextToDisplay:=wSheet.Name
        End If

    Next wSheet

    MsgBox "Created the Index"

End Sub


'------------

Sub Copy_All_Defined_Names()
    Dim x As Name
    Dim SourceFile As String
    Dim SourceWb, ActiveWb As Workbook
    Dim SourceNames As Variant
    Dim Y As Variant
    Dim s As String
    
    Dim intChoice As Integer
    Dim strPath As String
    
       
    'On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Set ActiveWb = ActiveWorkbook

    'only allow the user to select one file
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    'Remove all other filters
    Call Application.FileDialog(msoFileDialogOpen).Filters.Clear
    'Add a custom filter
    Call Application.FileDialog(msoFileDialogOpen).Filters.Add( _
        "Excel Files Only", "*.xl*")
    
    'make the file dialog visible to the user
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    'determine what choice the user made
    If intChoice <> 0 Then
        'get the file path selected by the user
        strPath = Application.FileDialog( _
            msoFileDialogOpen).SelectedItems(1)
        Set SourceWb = Workbooks.Open(strPath)
    Else
        MsgBox "No File selected"
        GoTo ExitHandler
    End If
    
    
    ' Loop through all of the defined names in the active workbook.
    For Each x In SourceWb.Names
        On Error Resume Next
        ActiveWb.Names.Add Name:=x.Name, RefersTo:=x.Value
    Next x
    
    SourceWb.Close SaveChanges:=False

ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox Err.Description
    Resume ExitHandler

End Sub


Sub CombineWorkbooks()
    Dim FilesToOpen
    Dim x, i As Integer
    Dim ActiveWb, SourceWb As Workbook
    Dim Sh As Worksheet
    

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Set ActiveWb = ActiveWorkbook

    FilesToOpen = Application.GetOpenFilename _
      (FileFilter:="Excel Files (*.xl*), *.xl*", _
      MultiSelect:=True, Title:="Files to Merge")

    If TypeName(FilesToOpen) = "Boolean" Then
        MsgBox "No Files were selected"
        GoTo ExitHandler
    End If

    x = 1
    While x <= UBound(FilesToOpen)
        Set SourceWb = Workbooks.Open(Filename:=FilesToOpen(x), ReadOnly:=True)
        
        For Each Sh In SourceWb.Worksheets
            i = ActiveWb.Sheets.Count
            On Error Resume Next
            SourceWb.Worksheets(Sh.Name).Copy _
                After:=ActiveWb.Sheets(i)
        Next Sh
        
        SourceWb.Close SaveChanges:=False
        
        x = x + 1
    Wend

ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox Err.Description
    Resume ExitHandler
End Sub


Sub Sort_Active_Book()
    Dim i As Integer
    Dim j As Integer
    Dim iAnswer As VbMsgBoxResult

    iAnswer = MsgBox("Sort Sheets in Ascending Order?" & Chr(10) _
     & "Clicking No will sort in Descending Order", _
     vbYesNoCancel + vbQuestion + vbDefaultButton1, "Sort Worksheets")

    oldStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    Application.StatusBar = "Sorting..."

    iNumSheets = Sheets.Count
    ReDim sMySheets(1 to iNumSheets)

    For i = 1 to iNumSheets
        sMySheets(i) = Sheets(i).Name
    Next iAnswer
    
    QuickSort sMySheets

    Application.ScreenUpdating = False

    For i = 1 to iNumSheets
        If iAnswer = vbYes Then
            Sheets(sMySheets(i)).Move Before:=Sheets(i)

        ElseIf iAnswer = vbNo Then
            Sheets(sMySheets(i)).Move After:=Sheets(iNumSheets - i + 1)
        End If

        Application.StatusBar = "Sorting Sheet: " & Sheets(sMySheets(i)).Name
    Next i

    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar

    Application.ScreenUpdating = True
    
    If iAnswer <> vbCancel Then
        MsgBox ("Sheets have been sorted")
    End If


Public Sub QuickSort(arr() As String, Optional ByVal low As Long = -1, Optional ByVal high As Long = -1)
    Dim pivotIndex As Long
    Dim pivotValue As String
    Dim i As Long
    Dim j As Long
    Dim temp As String

    ' Initialize low and high on the first call
    If low = -1 Then low = LBound(arr)
    If high = -1 Then high = UBound(arr)

    ' If the low index is lower than the high index, continue
    If low < high Then
        pivotIndex = low
        pivotValue = arr(pivotIndex)

        i = low
        j = high

        ' Perform partitioning
        Do While i <= j
            Do While arr(i) < pivotValue
                i = i + 1
            Loop
            Do While arr(j) > pivotValue
                j = j - 1
            Loop
            If i <= j Then
                ' Swap values
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop

    ' Recursively sort left and right partitions
    If low < j Then QuickSort arr, low, j
    If i < high Then QuickSort arr, i, high
    End If
End Sub


Sub ResetCommentsPosition()
    Dim ws As Worksheet
    Dim cmt As Comment

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Loop through each comment in the worksheet
        For Each cmt In ws.Comments
            ' Reset comment position to the default
            With cmt.Shape
                .Top = cmt.Parent.Top + 5
                .Left = cmt.Parent.Left + 5
            End With
        Next cmt
    Next ws

    MsgBox "All comments have been reset to default positions!", vbInformation
End Sub

Sub CommentsAutoSize()
    Dim ws As Worksheet
    Dim cmt As Comment
    Dim maxWidth As Double
    Dim fixedWidth As Double

    ' Set the maximum allowed width and fixed width for comments
    maxWidth = 300
    fixedWidth = 200

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Loop through each comment in the worksheet
        For Each cmt In ws.Comments
            With cmt.Shape
                ' Auto size the comment to fit the text
                .TextFrame.AutoSize = True

                ' Check if the width exceeds the maxWidth
                If .Width > maxWidth Then
                    ' Set the width to fixedWidth and adjust the height proportionally
                    .Width = fixedWidth
                    .Height = .TextFrame.TextRange.BoundHeight
                End If
            End With
        Next cmt
    Next ws

    MsgBox "All comments have been auto-sized and adjusted.", vbInformation
End Sub


Sub RemoveConditionalFormatting()
    Dim ws As WorkSheet
    Set ws = ActiveSheet

    ws.Cells.FormatConditions.Delete
End Sub

