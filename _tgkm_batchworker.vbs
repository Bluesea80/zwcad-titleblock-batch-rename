' ============================================================
'  _tgkm_batchworker.vbs —— 批处理引擎（被 TKGPR.hta 调用，隐藏运行）
'  配置：读 %TEMP%\tkgm_batch.ini（key=value：root/xml/pat/newname/scale/lsp/helper/doneflag/log）
'  流程：逐张 复用或打开 → 激活 → 载入LSP → TKGREN改名 → 刷新(EZ_SMEDI_OPENTASKWITHFILE+助手) → TKGFIX居中 → Save → Close
' ============================================================
Option Explicit
Dim fso, ws, app, docs, doc
Set fso = CreateObject("Scripting.FileSystemObject")
Set ws  = CreateObject("WScript.Shell")

' ---------- 读配置 ----------
Dim iniPath, iniLines, cfg
Set cfg = CreateObject("Scripting.Dictionary")
iniPath = ws.ExpandEnvironmentStrings("%TEMP%") & "\tkgm_batch.ini"
If Not fso.FileExists(iniPath) Then WScript.Quit
Dim raw, kv, i2
raw = fso.OpenTextFile(iniPath, 1).ReadAll
For Each kv In Split(raw, vbCrLf)
  If InStr(kv, "=") > 0 Then
    i2 = InStr(kv, "=")
    cfg(Trim(Left(kv, i2 - 1))) = Trim(Mid(kv, i2 + 1))
  End If
Next

Dim root, xmlPath, pat, newName, scaleS, lspFile, helper, doneFlag, logfile
Dim doPurge, purgeFlag
root    = cfg("root")
xmlPath = cfg("xml")
pat     = cfg("pat")
newName = cfg("newname")
scaleS  = cfg("scale")
lspFile = cfg("lsp")
helper  = cfg("helper")
doneFlag= cfg("doneflag")
logfile = cfg("log")
doPurge = LCase(Trim(cfg("purge")))
If doPurge = "" Then doPurge = "0"

Sub Log(s)
  Dim ff
  On Error Resume Next
  Set ff = fso.OpenTextFile(logfile, 8, True)
  ff.WriteLine Time & "  " & s
  ff.Close
End Sub

' ---------- 连接 ZWCAD ----------
On Error Resume Next
Set app = GetObject(, "ZWCAD.Application")
If Err.Number <> 0 Then
  Log "连接 ZWCAD 失败：" & Err.Description
  Log "===== 失败结束 ====="
  WScript.Quit
End If
' 全程容错（对齐 v8.7 已实跑通过的写法）：任何一步 COM 报错都不能让引擎
' 静默退出或弹错误框卡死；单张失败只记日志，继续下一张
On Error Resume Next
Set docs = app.Documents

' ---------- 工具 ----------
Function FindOpenDoc(path)
  Dim d2, nm
  Set FindOpenDoc = Nothing
  On Error Resume Next
  For Each d2 In docs
    Err.Clear
    nm = d2.FullName
    If Err.Number = 0 Then
      If LCase(nm) = LCase(path) Then
        Set FindOpenDoc = d2
        Exit Function
      End If
    End If
  Next
  On Error GoTo 0
End Function

Sub MkDirs(p)
  On Error Resume Next
  If p = "" Then Exit Sub
  If fso.FolderExists(p) Then Exit Sub
  MkDirs fso.GetParentFolderName(p)
  If Not fso.FolderExists(p) Then fso.CreateFolder p
  On Error GoTo 0
End Sub

Function WaitFlag(f, secs)
  Dim t1
  WaitFlag = False
  t1 = Timer
  Do While (Timer - t1 < secs)
    WScript.Sleep 400
    If fso.FileExists(f) Then
      WaitFlag = True
      Exit Function
    End If
  Loop
End Function

Function ReadText(f)
  On Error Resume Next
  ReadText = ""
  If fso.FileExists(f) Then ReadText = Trim(fso.OpenTextFile(f, 1).ReadAll)
  On Error GoTo 0
End Function

Sub CollectDwg(folder, dict)
  Dim f2, s2
  On Error Resume Next
  For Each f2 In fso.GetFolder(folder).Files
    If LCase(fso.GetExtensionName(f2.Name)) = "dwg" Then
      If Left(f2.Name, 2) <> "~$" Then dict.Add f2.Path, f2.Path
    End If
  Next
  For Each s2 In fso.GetFolder(folder).SubFolders
    If Left(s2.Name, 2) <> "备份" Then CollectDwg s2.Path, dict
  Next
  On Error GoTo 0
End Sub

' ---------- 开始 ----------
Log "===== 批处理引擎启动 ====="
Log "根文件夹：" & root
Log "工作包XML：" & xmlPath
Log "改名模式：" & pat & "   新名：" & newName & "   优先关键字：" & cfg("prefer") & "   比例：" & scaleS
Dim files, cnt, p, rel, dst, bakDir
Set files = CreateObject("Scripting.Dictionary")
CollectDwg root, files
cnt = files.Count
Log "收集图纸：" & cnt & " 张"
If cnt = 0 Then
  Log "！！没有可处理的 DWG，结束。"
  Log "===== 失败结束 ====="
  WScript.Quit
End If

' 引擎自定位 LISP / 助手（不依赖 ini 路径，防止中文路径解析偏差）
Dim myDir
myDir = fso.GetParentFolderName(WScript.ScriptFullName)
If (Not fso.FileExists(lspFile)) Then lspFile = myDir & "\图框块批量改名.lsp"
If (Not fso.FileExists(helper))  Then helper  = myDir & "\_tgkm_openhelper.vbs"
Log "引擎目录：" & myDir
Log "LISP：" & lspFile & "   存在=" & CStr(fso.FileExists(lspFile))
Log "助手：" & helper & "   存在=" & CStr(fso.FileExists(helper))
If Not (fso.FileExists(lspFile) And fso.FileExists(helper)) Then
  Log "！！LISP 或助手脚本缺失，结束。"
  Log "===== 失败结束 ====="
  WScript.Quit
End If
Log "（备份已由主页面完成，本引擎不再重复备份）"

Dim i, od, renFlag, fixFlag
Dim f3
renFlag = ws.ExpandEnvironmentStrings("%TEMP%") & "\tkgm_ren.flag"
fixFlag = ws.ExpandEnvironmentStrings("%TEMP%") & "\tkgm_fix.flag"
purgeFlag = ws.ExpandEnvironmentStrings("%TEMP%") & "\tkgm_purge.flag"

i = 0
For Each p In files
  i = i + 1
  Log "[" & i & "/" & cnt & "] " & p

  Set od = FindOpenDoc(p)
  If Not (od Is Nothing) Then
    Set doc = od
    Log "   原已打开：复用现有文档"
  Else
    Err.Clear
    Set doc = docs.Open(p)
    If Err.Number <> 0 Then
      Log "   打开失败(" & Err.Number & ")：" & Err.Description
      Err.Clear
      Set doc = Nothing
    End If
  End If

  If Not (doc Is Nothing) Then
    doc.Activate
    WScript.Sleep 400
    Log "   已激活"

    ' 载入 LISP（改名/居中入口在里面）
    Err.Clear
    doc.SendCommand "(load """ & Replace(lspFile, "\", "/") & """)" & vbCr
    WScript.Sleep 1500

    ' ① 改名
    If pat <> "" Then
      If fso.FileExists(renFlag) Then fso.DeleteFile renFlag, True
      Err.Clear
      doc.SendCommand "TKGREN" & vbCr
      If Not WaitFlag(renFlag, 45) Then
        Log "   ①改名无响应，重发一次"
        Err.Clear
        doc.SendCommand "TKGREN" & vbCr
        If Not WaitFlag(renFlag, 60) Then Log "   ①改名超时"
      End If
      Log "   ①改名结果：" & Trim(ReadText(renFlag))
    Else
      Log "   ①改名：跳过"
    End If

    ' ①b 清理无用图框块定义（可选，仅删 0 实例的块，图面引用不动）
    If doPurge = "1" Then
      If fso.FileExists(purgeFlag) Then fso.DeleteFile purgeFlag, True
      Err.Clear
      doc.SendCommand "TKGPURGE" & vbCr
      If Not WaitFlag(purgeFlag, 45) Then
        Log "   ①b清理无响应，重发一次"
        Err.Clear
        doc.SendCommand "TKGPURGE" & vbCr
        If Not WaitFlag(purgeFlag, 60) Then Log "   ①b清理超时"
      End If
      Log "   ①b清理：" & Trim(ReadText(purgeFlag))
    End If

    ' ② 刷新
    If fso.FileExists(doneFlag) Then fso.DeleteFile doneFlag, True
    Dim pf, pfh
    pf = fso.GetParentFolderName(helper) & "\_tgkm_openhelper.param"
    Set pfh = fso.CreateTextFile(pf, True)
    pfh.WriteLine xmlPath
    pfh.WriteLine doneFlag
    pfh.Close
    ws.Run "wscript.exe """ & helper & """", 0, False
    Err.Clear
    doc.SendCommand "EZ_SMEDI_OPENTASKWITHFILE" & vbCr
    If WaitFlag(doneFlag, 180) Then
      Log "   ②刷新完成"
    Else
      Log "   ②刷新超时(180s)"
    End If

    ' ③ 居中
    If fso.FileExists(fixFlag) Then fso.DeleteFile fixFlag, True
    Err.Clear
    doc.SendCommand "TKGFIX" & vbCr
    If Not WaitFlag(fixFlag, 45) Then
      Log "   ③居中无响应，重发一次"
      Err.Clear
      doc.SendCommand "TKGFIX" & vbCr
      If Not WaitFlag(fixFlag, 60) Then Log "   ③居中超时"
    End If
    Log "   ③居中结果：" & Trim(ReadText(fixFlag))

    ' ④ 存盘（无改动时 Save 可能报错，重试一次）
    Dim saved
    saved = False
    For i2 = 1 To 2
      If Not saved Then
        Err.Clear
        doc.Save
        If Err.Number = 0 Then
          saved = True
          Log "   ④存盘完成"
        Else
          Log "   ④存盘失败(第" & i2 & "次) 号" & Err.Number & "：" & Err.Description
        End If
      End If
    Next

    ' ⑤ 关闭
    Err.Clear
    doc.Close
    If Err.Number <> 0 Then
      Log "   ⑤关闭失败(" & Err.Number & ")：" & Err.Description
      Err.Clear
    Else
      Log "   ⑤已关闭"
    End If
  End If
  Set doc = Nothing
Next

Log "===== 全部完成 ====="
