@echo off
reg add HKCR\Word.Document.12\shell\toPDF /ve /f /d "תΪPDF"
reg add HKCR\Word.Document.12\shell\toPDF\command /ve /d "C:\Windows\SysWOW64\wscript.exe %~dp0doc2pdf.vbs \"%%1\"" /f
reg add HKCR\Word.Document.8\shell\toPDF /ve /f /d "תΪPDF"
reg add HKCR\Word.Document.8\shell\toPDF\command /ve /d "C:\Windows\SysWOW64\wscript.exe %~dp0doc2pdf.vbs \"%%1\"" /f