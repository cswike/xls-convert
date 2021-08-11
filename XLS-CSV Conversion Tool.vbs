' Get current directory of this script
Set Fso = WScript.CreateObject("Scripting.FileSystemObject")
ParentFolder  = FSO.GetParentFolderName(WScript.ScriptFullName)
sParentFolder = ParentFolder & "\"

' Open new Excel workbook
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()

'Import and run .bas module, but .bas has to be in the same folder as this script
objExcel.VBE.ActiveVBProject.VBComponents.Import sParentFolder & "xls-convert-module.bas"
objExcel.Run "main"

' Close blank Excel workbook
objExcel.DisplayAlerts = False
objWorkbook.Close False
Set objWorkbook = Nothing
Set objExcel = Nothing
WScript.Quit