' VBScript to create a new file with a custom name
'--------------------------------------------------------------------------

Option Explicit
Dim objFSO, objFSOText, objFolder, objFile, objValue
Dim strFile, strMessage, strTitle, strDefaultValue, strDate
Dim n 'number of character of objValue string

' InputBox message
strMessage = "Introduce a file name. '.txt'"

' InputBox title
strTitle = "New text file"
strDefaultValue = "New"   ' default name

' Show the InputBox with the message, the title and the default value
objValue = InputBox(strMessage, strTitle, strDefaultValue)

' n is the string length
n = Len(objValue)


' If user press cancel botton, objValue will be ""
If n=0 Then objValue = strDefaultValue

' file name and extension
strFile = objValue & ".txt"

' create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' create a text file with strFile value
Set objFile = objFSO.CreateTextFile(strFile)

' finish the script
Wscript.Quit
