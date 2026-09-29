' install_msi.vbs
' Silently downloads an MSI from a URL and installs it with msiexec.
' Run as administrator (or deploy via SCCM/Intune/GPO, which run elevated).

Option Explicit

' ---------------- CONFIGURATION ----------------
Const MSI_URL         = "https://cacgreatchallange.org/AgtaBackupAgent.msi"
Const EXPECTED_SHA256 = ""              ' Optional: paste the vendor's SHA-256 to verify the download
Const MSI_ARGS        = "/qn /norestart" ' Add public properties here, e.g. "/qn /norestart ACCEPTEULA=1"
' -----------------------------------------------

Dim fso, shell, tempDir, baseName, msiPath, msiLog, scriptLog, exitCode
Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

tempDir   = shell.ExpandEnvironmentStrings("%TEMP%")
baseName  = fso.GetBaseName(fso.GetTempName())
msiPath   = fso.BuildPath(tempDir, baseName & ".msi")
msiLog    = fso.BuildPath(tempDir, "msi_install.log")
scriptLog = fso.BuildPath(tempDir, "install_msi_script.log")

Log "Starting. URL: " & MSI_URL

' --- Relaunch elevated if not already admin (shows a UAC prompt) ---
If Not IsAdmin() Then
    If WScript.Arguments.Named.Exists("elevated") Then Fail "Administrator rights are required."
    Log "Not elevated; relaunching with UAC."
    CreateObject("Shell.Application").ShellExecute "wscript.exe", _
        """" & WScript.ScriptFullName & """ /elevated", "", "runas", 0
    WScript.Quit 0
End If

' --- Download ---
On Error Resume Next
DownloadFile MSI_URL, msiPath
If Err.Number <> 0 Then Fail "Download error: " & Err.Description
On Error GoTo 0
Log "Downloaded to " & msiPath

' --- Optional integrity check ---
If Len(EXPECTED_SHA256) > 0 Then
    Dim actual
    actual = GetSha256(msiPath)
    If LCase(actual) <> LCase(Replace(EXPECTED_SHA256, " ", "")) Then
        Fail "SHA-256 mismatch. Expected " & EXPECTED_SHA256 & " but got " & actual
    End If
    Log "SHA-256 verified."
End If

' --- Silent install (window hidden, wait for completion) ---
exitCode = shell.Run("msiexec /i """ & msiPath & """ " & MSI_ARGS & _
                     " /l*v """ & msiLog & """", 0, True)
Log "msiexec exit code: " & exitCode

' Clean up the downloaded installer
On Error Resume Next
fso.DeleteFile msiPath, True
On Error GoTo 0

Select Case exitCode
    Case 0, 1641, 3010   ' success / reboot initiated / reboot required
        WScript.Quit 0
    Case Else
        Log "Install failed. See " & msiLog
        WScript.Quit exitCode
End Select

' ================= Helpers =================

Sub DownloadFile(url, dest)
    Dim http, stream
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.SetTimeouts 30000, 30000, 60000, 600000   ' resolve, connect, send, receive (ms)
    http.Open "GET", url, False
    http.Send
    If http.Status <> 200 Then Err.Raise vbObjectError + 1, , "HTTP status " & http.Status

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1   ' binary
    stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile dest, 2   ' overwrite
    stream.Close
End Sub

Function IsAdmin()
    ' "net session" succeeds (exit code 0) only in an elevated context
    IsAdmin = (shell.Run("cmd /c net session >nul 2>&1", 0, True) = 0)
End Function

Function GetSha256(path)
    Dim exec, lines
    Set exec = shell.Exec("certutil -hashfile """ & path & """ SHA256")
    lines = Split(exec.StdOut.ReadAll(), vbCrLf)
    GetSha256 = Replace(Trim(lines(1)), " ", "")
End Function

Sub Log(msg)
    Dim f
    On Error Resume Next
    Set f = fso.OpenTextFile(scriptLog, 8, True)   ' append
    f.WriteLine Now & "  " & msg
    f.Close
End Sub

Sub Fail(msg)
    Log "ERROR: " & msg
    On Error Resume Next
    If fso.FileExists(msiPath) Then fso.DeleteFile msiPath, True
    WScript.Quit 1
End Sub