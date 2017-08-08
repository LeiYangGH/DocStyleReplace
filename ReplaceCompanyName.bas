Sub ReplaceCompanyName()
'
' ReplaceCompanyName Macro
'
'
 Dim doc As Document, myFile As String
    Dim a As Range
    Dim DirName As String
    'DirName = "C:\DocReplace\TestSrc\" ????????????(?????)?¡À??
    'DirName = "C:\DocReplace\ReplaceHeaderFooter\¡ä?¨¬?\?¨¬?¨¦????\¨°???\2???¡¤?¨º¡Á¨°3¨°3??o¨ª????¨°3¨°3??¦Ì?\"
    'DirName = "C:\DocReplace\ReplaceHeaderFooter\¡ä?¨¬?\?¨¬?¨¦????\¨°???\"
    'DirName = "C:\DocReplace\ReplaceHeaderFooter\¡ä?¨¬?\?¨¬?¨¦????\"
    'DirName = "C:\DocReplace\ReplaceHeaderFooter\¡ä?¨¬?\?¨º¨¢?¡À¨º¡Á?\"
    DirName = "C:\DocReplace\ReplaceHeaderFooter\¡ä?¨¬?\?¨º¨¢?¡À¨º¡Á?\¨°???\"
    FileName = Dir(DirName & "*.doc")
    HeaderTemplateFileName = "C:\DocReplace\Template\PageHeader.doc"
    Set docTemplate = Documents.Open(HeaderTemplateFileName)
    Set hdrTemplate = docTemplate.Sections(1).Headers(wdHeaderFooterPrimary)
    Set hdrTemplate1st = docTemplate.Sections(1).Headers(wdHeaderFooterFirstPage)
    On Error Resume Next
    Do While FileName <> ""

        FullName = DirName & FileName
        Set doc = Documents.Open(FullName)
      
            Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "3¨¦??¡À?¨¬?¨°?¨°¦Ì¨®D?T1???"
        .Replacement.Text = "??¡ä¡§¨¨¨º¡ã2¨°?¨°¦Ì¨®D?T?e¨¨?1???¡ê¡§?¨¤¡ã¡Á?-?¨´¦Ì?¡ê?"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
        doc.SaveAs2 FileName:="C:\DocReplace\ReplaceHeaderFooter\¨¬???1?????3?\?¨º¨¢?¡À¨º¡Á?\¨°???\" & FileName
    
        doc.Close
        Set doc = Nothing
        

        FileName = Dir
    Loop
    


    
End Sub