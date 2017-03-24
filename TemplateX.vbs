Option Explicit

dim objWord, objDoc, objParagraph, paragraph_data, wordCount, strWord
dim paragraph_ctr, myString

' we need to continue through errors since if Word isn't
' open the GetObject line will give an error
On Error Resume Next
Set objWord = GetObject(, "Word.Application")
' we've tried to get Word but if it's nothing then it isn't open
If objWord Is Nothing Then
  Set objWord = CreateObject("Word.Application")
End If
' it's good practice to reset error warnings
On Error GoTo 0
Set objDoc = objWord.Documents.Open("C:\Users\akatoybits\Desktop\xyz.docx")
' open your document and ensure its visible and activate after openning
objWord.Visible = True
objWord.Activate

paragraph_ctr = 1
wordCount = 1
          
For Each objParagraph in objDoc.Paragraphs
    For wordCount =  0 to objParagraph.Words
        wordCount = wordCount + 1
    Next
 Next  
MsgBox (wordCount)
