#NoEnv
#SingleInstance force
Process, Exist, EXCEL.EXE
if !ErrorLevel
return
Loop
{
xl := ComObjActive("Excel.Application")
For book in xl.Workbooks
book.Close(1)
xl.Quit(), xl := ""
Process, Exist, EXCEL.EXE
} Until !ErrorLevel
Exit