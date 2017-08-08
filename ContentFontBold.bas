Sub ContentFontBold()
'
' ContentFontBold Macro
'
'
    'Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    'Selection.CopyFormat
    'Selection.PasteFormat
    With Selection.Font
    .Name = "ºÚÌå"
    .Size = 12
    .Bold = True
End With
    ActiveDocument.Save
End Sub