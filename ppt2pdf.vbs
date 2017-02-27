On Error Resume Next
Set argv = WScript.Arguments
if argv.Count < 1 then
	WScript.Quit
end if
Set pptApp = CreateObject("PowerPoint.Application")
for i = 0 to (argv.Count - 1)
	for j = 1 to 1
		filename = argv(i)
		if Lcase(right(filename, 4)) = ".ppt" then
			pdfname = left(filename, len(filename)-3) + "pdf"
		elseif Lcase(right(filename, 5)) = ".pptx" then
			pdfname = left(filename, len(filename)-4) + "pdf"
		else
			exit for
		end if
		Set MyPress = pptApp.Presentations.Open(filename)
		ppSaveAsPDF = 32 'to be fixed!
		MyPress.SaveAs pdfname, ppSaveAsPDF, false
		MyPress.Close
	next
next
pptApp.Quit
