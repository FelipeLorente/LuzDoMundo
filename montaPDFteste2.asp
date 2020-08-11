<%


Set d = CreateObject("ABCpdf10.Doc")

d.FontSize = 96
d.AddText "Hello World!"
d.Save "mydoc.pdf"
MsgBox "Finished"

'
'Set theDoc = Server.CreateObject("ABCpdf11.Doc")
'
'theDoc.Rect.Inset 10, 10
'
'theDoc.Page = theDoc.AddPage()
'theURL = "https://www.google.com.br"
'theID = theDoc.AddImageUrl(theURL, True, 770, true)
'
'Do
'  theDoc.FrameRect ' add a black border
'  If Not theDoc.Chainable(theID) Then Exit Do
'  theDoc.Page = theDoc.AddPage()
'  theID = theDoc.AddImageToChain(theID)
'Loop
'
'For i = 1 To theDoc.PageCount
'  theDoc.PageNumber = i
'  theDoc.Flatten
'Next
'
'theDoc.Save Server.MapPath("arquivos\teste.pdf")
'
'response.End()
%>