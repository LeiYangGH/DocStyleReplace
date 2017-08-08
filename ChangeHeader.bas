Sub ChangeHeader()
 Dim doc As Document, myFile As String
    Dim a As Range
    Dim DirName As String
    DirName = "C:\DocReplace\TestSrc\"
    FileName = Dir(DirName & "*.doc")
     
    Do While FileName <> ""
        'MsgBox myFile
        FullName = DirName & FileName
        Set doc = Documents.Open(FullName)
        If Not doc Is Nothing Then
            'MsgBox doc.Name
            If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
            ActiveWindow.Panes(2).Close
        End If
        If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
            ActiveWindow.ActivePane.View.Type = wdPrintView
        End If
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        Selection.MoveRight Unit:=wdCharacter, Count:=10, Extend:=wdExtend
        Selection.Paste
        'ActiveDocument.Save
        ActiveDocument.SaveAs2 FileName:="C:\DocReplace\TestDes\" & FileName
   
        'doc.Save
        doc.Close
        Set doc = Nothing
        End If

        FileName = Dir
    Loop
End Sub

