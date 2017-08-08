Sub ChangeHeader()
 Dim doc As Document, myFile As String
    Dim a As Range
    Dim DirName As String
    DirName = "C:\DocReplace\TestSrc\"
    FileName = Dir(DirName & "*.doc")
    HeaderTemplateFileName = "C:\DocReplace\Template\PageHeader.doc"
    Set docTemplate = Documents.Open(HeaderTemplateFileName)
     Set hdr1 = docTemplate.Sections(1).Headers(wdHeaderFooterPrimary)
    Do While FileName <> ""

        FullName = DirName & FileName
        Set doc = Documents.Open(FullName)
      

        Set hdr2 = doc.Sections(1).Headers(wdHeaderFooterPrimary)
        hdr1.Range.Copy
        hdr2.Range.Paste
        doc.SaveAs2 FileName:="C:\DocReplace\TestDes\" & FileName
    
        doc.Close
        Set doc = Nothing
        

        FileName = Dir
    Loop
    
End Sub

