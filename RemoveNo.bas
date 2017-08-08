Sub RemoveNo()
'
' RemoveNo Macro
'
'
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Application.Keyboard (2052)
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.Paragraphs(1).SelectNumber
    Selection.TypeBackspace
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.Paragraphs(1).SelectNumber
    Selection.TypeBackspace
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeBackspace
    Selection.Paragraphs(1).SelectNumber
    Selection.TypeBackspace
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeBackspace
    Selection.Paragraphs(1).SelectNumber
    Selection.TypeBackspace
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeBackspace
    Selection.Paragraphs(1).SelectNumber
    Selection.TypeBackspace
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.Paragraphs(1).SelectNumber
    Selection.TypeBackspace
    Selection.MoveUp Unit:=wdLine, Count:=5
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.MoveRight Unit:=wdCell
    Selection.Paragraphs(1).SelectNumber
    Selection.TypeBackspace
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.Paragraphs(1).SelectNumber
    Selection.TypeBackspace
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.Paragraphs(1).SelectNumber
    Selection.TypeBackspace
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.Paragraphs(1).SelectNumber
    Selection.TypeBackspace
    ActiveDocument.Save
    ActiveWindow.Close
    Application.Quit
End Sub