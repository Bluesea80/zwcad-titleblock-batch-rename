' ============================================================
'  _tgkm_openhelper.vbs —— 刷新助手（被 TKGPR 批量脚本调用）
'  作用：把工作包 XML 路径放进剪贴板 -> 等"打开"浏览窗出现 ->
'        Alt+N 切到文件名框 -> Ctrl+V 粘贴 -> 回车 ->
'        轮询"提示"框并回车关闭 -> 写完成标记
'  参数：由 TKGPR 通过同名 .param 文件传：第1行=XML路径 第2行=标记文件
' ============================================================
Option Explicit
On Error Resume Next

Dim ws, fso, lg, paramFile, xmlPath, marker, i, f, tf
Set ws  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

lg = ws.ExpandEnvironmentStrings("%TEMP%") & "\tkgm_vbs.log"
paramFile = fso.GetParentFolderName(WScript.ScriptFullName) & "\_tgkm_openhelper.param"

Sub L(s)
  Dim ff
  Set ff = fso.OpenTextFile(lg, 8, True)
  ff.WriteLine Now & "  " & s
  ff.Close
End Sub

L "helper start"
If Not fso.FileExists(paramFile) Then
  L "缺少参数文件 " & paramFile
  WScript.Quit
End If
Dim lines
lines = Split(fso.OpenTextFile(paramFile, 1).ReadAll, vbCrLf)
xmlPath = lines(0)
marker  = lines(1)
L "xml=" & xmlPath & " marker=" & marker

' 1) 路径写入剪贴板（PowerShell 按 ANSI(GBK) 读，中文无损）
tf = ws.ExpandEnvironmentStrings("%TEMP%") & "\tkgm_path.txt"
Set f = fso.CreateTextFile(tf, True)
f.Write xmlPath
f.Close
ws.Run "powershell -NoProfile -Command ""Set-Clipboard -Value (Get-Content -Raw -Encoding Default -LiteralPath '" & tf & "')""", 0, False
L "clipboard set"

' 2) 等"打开"浏览窗（最长 60s）
i = 0
Do While i < 300
  WScript.Sleep 200
  i = i + 1
  If ws.AppActivate("打开") Then Exit Do
Loop
L "open-dialog poll i=" & i
If i < 300 Then
  WScript.Sleep 400
  ws.SendKeys "%n"
  WScript.Sleep 150
  ws.SendKeys "^a"
  WScript.Sleep 80
  ws.SendKeys "^v"
  WScript.Sleep 200
  ws.SendKeys "~"
  L "path pasted and entered"
End If

' 3) 轮询"提示"框并关闭（最长 120s）
i = 0
Do While i < 120
  WScript.Sleep 1000
  i = i + 1
  If ws.AppActivate("提示") Then
    WScript.Sleep 300
    ws.SendKeys "~"
    L "tip box closed at " & i
    Exit Do
  End If
Loop
L "done, tip poll i=" & i

' 4) 写完成标记
Dim mf : Set mf = fso.CreateTextFile(marker, True) : mf.Close
L "marker written"
