Option Explicit

dim objWord, objDoc, objParagraph, strWord
dim paragraph_ctr, word_ctr

On Error Resume Next
Set objWord = GetObject(, "Word.Application")

If objWord Is Nothing Then
  Set objWord = CreateObject("Word.Application")
End If

 On Error GoTo 0
Set objDoc = objWord.Documents.Open("C:\Users\xxx-xxx\Desktop\xyz.docx")
objWord.Visible = True
objWord.Activate

paragraph_ctr = 0
word_ctr = 0
          

For Each objParagraph in objDoc.Paragraphs
    objParagraph.Range.Select
	
		For each strWord in objParagraph.Range.Words
			
			word_ctr = word_ctr + 1
		Next
	paragraph_ctr = paragraph_ctr + 1

Next  

MsgBox ("Number of paragraph: " & paragraph_ctr & chr(13) & "Number of words: " & word_ctr)
