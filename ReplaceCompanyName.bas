Sub ReplaceCompanyName()
'
' ReplaceCompanyName Macro
'
'
 Dim doc As Document, myFile As String
    Dim a As Range
    Dim DirName As String
    'DirName = "C:\DocReplace\TestSrc\" ????????????(?????)?��??
    'DirName = "C:\DocReplace\ReplaceHeaderFooter\��?��?\?��?��????\��???\2???��?������3��3??o��????��3��3??��?\"
    'DirName = "C:\DocReplace\ReplaceHeaderFooter\��?��?\?��?��????\��???\"
    'DirName = "C:\DocReplace\ReplaceHeaderFooter\��?��?\?��?��????\"
    'DirName = "C:\DocReplace\ReplaceHeaderFooter\��?��?\?����?������?\"
    DirName = "C:\DocReplace\ReplaceHeaderFooter\��?��?\?����?������?\��???\"
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
        .Text = "3��??��?��?��?���̨�D?T1???"
        .Replacement.Text = "??�䡧������2��?���̨�D?T?e��?1???�ꡧ?�����?-?����?��?"
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
        
        doc.SaveAs2 FileName:="C:\DocReplace\ReplaceHeaderFooter\��???1?????3?\?����?������?\��???\" & FileName
    
        doc.Close
        Set doc = Nothing
        

        FileName = Dir
    Loop
    


    
End Sub