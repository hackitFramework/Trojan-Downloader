Set FSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")
Set WShell = CreateObject("WScript.Shell")
Set xHttp = CreateObject("Microsoft.XMLHTTP")
Set bStrm = CreateObject("Adodb.Stream")

temp = "C:\temp\"
ZipPath = temp & "rstandalone.zip"
ExtractPath = temp & "rstandalone"
rURL = "https://gitlab.com/RProgramming/R-standalone/repository/archive.zip?ref=master"
outBat = WShell.SpecialFolders("Startup") & "\runner.bat"

Set outFile = FSO.CreateTextFile(outBat, True)
outFile.Write "cd C:\" & vbCrLf & "C:\temp\rstandalone\R-standalone-master-2db7b18f6ef2967284a868b49ebc1319518e23e1\R-3.3.0\bin\Rscript.exe C:\temp\rstandalone\R-standalone-master-2db7b18f6ef2967284a868b49ebc1319518e23e1\R-3.3.0\
bi\rtt.R"
outFile.Close

' Create Directories
If NOT FSO.FolderExists(temp) Then
    FSO.CreateFolder(temp)
End If
If NOT FSO.FolderExists(ExtractPath) Then
    FSO.CreateFolder(ExtractPath)
End If

' Download R
xHttp.Open "GET", rURL, False
xHttp.Send

with bStrm
    .type = 1
    .open
    .write xHttp.responseBody
    .savetofile ZipPath, 1
End with

Set ZipFiles = objShell.NameSpace(ZipPath).Items()
objShell.NameSpace(ExtractPath).CopyHere ZipFiles, 20

Set FSO = Nothing
Set objShell = Nothing
