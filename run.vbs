Option Explicit
' Shift JISで実行しようね
Dim w,ws,fs
Dim clip_data,objText
Dim sUrl

sUrl = "https://awsacademy.instructure.com/courses/11235/modules/items/8132144"

Set w = WScript
Set ws = w.CreateObject("WScript.Shell")
Set fs = w.CreateObject("Scripting.FileSystemObject")

DebugPrint "Chrome召喚"
ws.Run "chrome.exe --allow-file-access-from-files -url "& sUrl
ws.AppActivate "chrome.exe"

DebugPrint "クリップボード抹消"
clearClipboard

DebugPrint "クリップボードにやって来るまで待機"
Do
    clip_data = GetClipboard
    If clip_data = vbCrLf Then
        DebugPrint "wait..."
    ElseIf Left(clip_data, 9) = "[default]" Then
        DebugPrint "データ取得: " & Len(clip_data) & "文字"
        DebugPrint clip_data
        Exit Do
    Else
        DebugPrint "内容な違います"
    End If
    w.Sleep 100
Loop

Set objText = fs.OpenTextFile("C:\Users\" & CreateObject("WScript.Network").UserName & "\.aws\credentials", 2)
objtext.WriteLine(clip_data)
objText.Close

DebugPrint "クリップボード抹消"
clearClipboard

ws.Popup "完了"&vbCrLf&vbCrLf&clip_data, 2,"aws_sandbox_run"

Sub DebugPrint(text)
    Dim CSCRIPT_EXE
    CSCRIPT_EXE = "cscript.exe"
    If LCase(Right(w.FullName, Len(CSCRIPT_EXE))) = CSCRIPT_EXE Then
        w.StdOut.WriteLine text
    End If
End Sub

Sub clearClipboard
    ws.Run "cmd /c echo off | clip", 0
End Sub

Function GetClipboard
    On Error Resume Next

    Dim r,f_pth,cmd
    r = ""

    f_pth = fs.GetSpecialFolder(2) & "\" & fs.GetTempName

    ' クリップボードの内容を一時ファイルへ
    cmd = "cmd /c PowerShell Get-Clipboard -Format Text > " & f_pth
    ws.Run cmd, 0, True

    ' 一時ファイルの処理
    If fs.FileExists(f_pth) Then
        r = fs.OpenTextFile(f_pth, 1).ReadAll
        fs.DeleteFile f_pth, True
    End If
    GetClipboard = r

    On Error GoTo 0
End Function
