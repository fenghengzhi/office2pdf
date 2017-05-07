On Error Resume Next
Set argv = WScript.Arguments '命令行参数 
if argv.Count < 1 then
	WScript.Quit
end if
Set docApp = CreateObject("Word.Application") '申明调用函数 
for i = 0 to (argv.Count - 1)
	for j = 1 to 1
		filename = argv(i)
		if Lcase(right(filename, 4)) = ".doc" then
			pdfname = left(filename, len(filename)-3)
		elseif Lcase(right(filename, 5)) = ".docx" then
			pdfname = left(filename, len(filename)-4)
		else
			exit for
		end if
		pdfname=pdfname+"pdf"
		Set MyPress = docApp.Documents.Open(filename)
		ppSaveAsPDF = 32 'to be fixed!
		MyPress.SaveAs pdfname, 17, false
		MyPress.Close
	next
next
docApp.Quit 
