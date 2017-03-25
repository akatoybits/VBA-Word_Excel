Option Explicit

dim objWord, objDoc, objParagraph, paragraph_data, wordCount, strWord
dim paragraph_ctr, myString

On Error Resume Next
Set objWord = GetObject(, "Word.Application")

If objWord Is Nothing Then
  Set objWord = CreateObject("Word.Application")
End If

On Error GoTo 0
Set objDoc = objWord.Documents.Open("C:\Users\XXX-XXX\Desktop\xyz.docx")

objWord.Visible = True
objWord.Activate

paragraph_ctr = 1
wordCount = 1
          
For Each objParagraph in objDoc.Paragraphs
    For wordCount =  0 to objDoc.Words
        wordCount = wordCount + 1
    Next
Next  
MsgBox (wordCount)
