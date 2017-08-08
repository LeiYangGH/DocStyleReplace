Sub ChangeHeader()
 Dim doc As Document, myFile As String
    Dim a As Range
    Dim DirName As String
    DirName = "C:\DocReplace\TestSrc\"
    FileName = Dir(DirName & "*.doc")
    HeaderTemplateFileName = "C:\DocReplace\Template\PageHeader.doc"
    Set docTemplate = Documents.Open(HeaderTemplateFileName)
    Set hdrTemplate = docTemplate.Sections(1).Headers(wdHeaderFooterPrimary)
    Set hdrTemplate1st = docTemplate.Sections(1).Headers(wdHeaderFooterFirstPage)
    Do While FileName <> ""

        FullName = DirName & FileName
        Set doc = Documents.Open(FullName)
      
        Set hdr1stDes = doc.Sections(1).Headers(wdHeaderFooterFirstPage)
        hdrTemplate1st.Range.Copy
        hdr1stDes.Range.Paste

        Set hdrDes = doc.Sections(1).Headers(wdHeaderFooterPrimary)
        hdrTemplate.Range.Copy
        hdrDes.Range.Paste
        
        doc.SaveAs2 FileName:="C:\DocReplace\TestDes\" & FileName
    
        doc.Close
        Set doc = Nothing
        

        FileName = Dir
    Loop
    
End Sub

