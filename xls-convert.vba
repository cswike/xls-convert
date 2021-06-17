Sub main()

Dim delRows As Boolean, f_Confirmation As Boolean, sameFolder As Boolean, sameName As Boolean
Dim delRowsNum As Integer

' *** CHANGE OPTIONS HERE AS NEEDED ***
' Option to delete first 3 rows (extraneous headers from original xls file)
' Can change # of rows to delete as well, i.e. if delRowsNum = 4 then rows 1-4 will be deleted.
delRows = True
delRowsNum = 3

' Make user confirm folder selection?
f_Confirmation = True

' Save to the same folder that contains the .xls files?
' (If False, will prompt user to select a Save folder as well as a source folder)
sameFolder = False

' Keep same file name for conversion?
' (If False, will use date from original filename INRPSGPF_yyyymmdd_hhmmss.xls and store number from cell A5)
' NOT RECOMMENDED to set this to False as some files may be empty
sameName = True

' *** END OPTIONS, CODE FOLLOWS ***





' initialize ConfirmFold and ConfirmSaveFold to vbNo for While loops
ConfirmFold = vbNo
ConfirmSaveFold = vbNo

' Choose source directory - while folder selection not confirmed, keep asking for folder selection
While ConfirmFold = vbNo
  folderPath = BrowseForFolder("Please select folder containing .xls files to be converted")
  If IsEmpty(folderPath) Then
    MsgBox "Operation canceled."
    Exit Sub
  ElseIf f_Confirmation = True Then
    ConfirmFold = MsgBox("Is " & folderPath & " the folder containing your Excel files?", vbYesNo)
  Else
    ConfirmFold = vbYes
  End If
Wend

' Choose directory to save to
If sameFolder = False Then
  While ConfirmSaveFold = vbNo
    savePath = BrowseForFolder("Please choose where to save converted .csv files")
    If IsEmpty(savePath) Then
      MsgBox "Operation canceled."
      Exit Sub
    ElseIf f_Confirmation = True Then
      ConfirmSaveFold = MsgBox("Save to " & savePath & "?", vbYesNo)
    Else
      ConfirmSaveFold = vbYes
    End If
  Wend
Else
  savePath = folderPath
End If


folderPath = folderPath & "\"
savePath = savePath & "\"
ChDir folderPath


'Optimize Macro Speed
Call LudicrousMode(True)

'Target File Extension (must include wildcard "*")
ext = "*.xls"
'Target Path with Ending Extention
oFile = Dir(folderPath & "\" & ext)

'Loop through each Excel file in folder
Do While oFile <> ""
  'Set variable equal to opened workbook
  Set wb = Workbooks.Open(Filename:=folderPath & oFile)
    
  'Ensure workbook has opened before moving on to next line of code
  DoEvents
  wb.Activate
  
  'Set file name if not the same
  If sameName = False Then
    custom_fName = wb.Sheets(1).Range("A5").Formula & " ASR Suggestion " & Mid(ActiveWorkbook.Name, 10, 8)
  End If
  
  'Optionally delete rows
  If delRows = True Then
    'delete rows 1 through (number)
    i = delRowsNum
    While i >= 1
      Rows(i).EntireRow.Delete
      i = i - 1
    Wend
  End If

  'Save as .csv :
  wb.Activate
  '...with same file name
  If sameName = True Then
    ActiveWorkbook.SaveAs Filename:=savePath & Left(ActiveWorkbook.Name, _
      (InStrRev(wb.Name, ".", -1, vbTextCompare) - 1)) & ".csv", _
      FileFormat:=xlCSV
  '...with customized file name
  Else
    ActiveWorkbook.SaveAs Filename:=savePath & custom_fName & ".csv", _
      FileFormat:=xlCSV
  End If

  'Close xls without saving
  wb.Close SaveChanges:=False

  'Ensure workbook has closed before moving on to next line of code
  DoEvents

  'Get next file
  oFile = Dir
Loop

Call LudicrousMode(False)

MsgBox "CSV export completed"

End Sub
'=========================================================================================
Function BrowseForFolder(Message As String)
  Dim oFolder
  Set oFolder = CreateObject("Shell.Application").BrowseForFolder(0, Message, 0, 0)
  If (oFolder Is Nothing) Then
    BrowseForFolder = Empty
  Else
    BrowseForFolder = oFolder.Self.Path
  End If
End Function
'=========================================================================================
Public Sub LudicrousMode(ByVal Toggle As Boolean)
'Adjusts Excel settings for faster VBA processing
    Application.ScreenUpdating = Not Toggle
    Application.EnableEvents = Not Toggle
    Application.DisplayAlerts = Not Toggle
    Application.EnableAnimations = Not Toggle
    Application.DisplayStatusBar = Not Toggle
    Application.PrintCommunication = Not Toggle
    Application.Calculation = IIf(Toggle, xlCalculationManual, xlCalculationAutomatic)
End Sub
