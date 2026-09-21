' ============================================================'  TKGPR 图框批量处理 v9.1 —— 一体化驱动（备用入口，只依赖 wscript.exe）
'
'  为什么保留这个文件：主交付 TKGPR批量处理.hta 依赖 mshta.exe，
'  而部分终端管控 / EDR 策略会禁用 mshta。本脚本不需要 mshta，
'  当 HTA 打不开时用它作退路；功能已与 HTA 主交付对齐
'  （含「优先关键字 Prefer」与「清理无用块定义 TKGPURGE」）。
'
'  一条龙：逐张 打开 → ①改名(可选) → ①b清理垃圾块(可选) → ②刷新图框
'          → ③折行居中 → ④存盘 → ⑤关闭
'
'  用法：先打开 ZWCAD（打开任意一张图纸），然后双击本脚本
'  依赖：同目录下的 _tgkm_openhelper.vbs 与 图框块批量改名.lsp
'  参数：写注册表 HKCU\Software\TKGM（与 HTA 共用同一套键）
' ============================================================
Option Explicit

Dim fso, ws, app, docs, doc
Dim root, xmlPath, files, i, cnt, p, helper, lspFile
Dim pf, pfh, rel, dst, bakDir, stamp
Dim pat, newName, scale, prefer, doPurge, doPurgeS
Dim renFlag, fixFlag, purgeFlag, doneFlag, LOGF
Dim msg, lastRoot, found, attempt, od, reused

Set fso = CreateObject("Scripting.FileSystemObject")
Set ws  = CreateObject("WScript.Shell")

' 脚本同目录资源
helper  = fso.GetParentFolderName(WScript.ScriptFullName) & "\_tgkm_openhelper.vbs"
lspFile = fso.GetParentFolderName(WScript.ScriptFullName) & "\图框块批量改名.lsp"
If Not fso.FileExists(helper) Then
  MsgBox "找不到助手脚本 _tgkm_openhelper.vbs（应与本脚本在同一文件夹）。"
  WScript.Quit
End If
If Not fso.FileExists(lspFile) Then
  MsgBox "找不到 图框块批量改名.lsp（应与本脚本在同一文件夹）。"
  WScript.Quit
End If

doneFlag  = ws.ExpandEnvironmentStrings("%TEMP%") & "\tkgm_done.flag"
renFlag   = ws.ExpandEnvironmentStrings("%TEMP%") & "\tkgm_ren.flag"
fixFlag   = ws.ExpandEnvironmentStrings("%TEMP%") & "\tkgm_fix.flag"
purgeFlag = ws.ExpandEnvironmentStrings("%TEMP%") & "\tkgm_purge.flag"

' ---------- 0. 连接 ZWCAD ----------
On Error Resume Next
Set app = GetObject(, "ZWCAD.Application")
If Err.Number <> 0 Then
  MsgBox "未找到正在运行的 ZWCAD。请先打开 ZWCAD（并打开任意一张图纸）再运行本脚本。"
  WScript.Quit
End If
On Error GoTo 0
Set docs = app.Documents

' ---------- 1. 选根文件夹（四入口，注册表记忆）----------
Do
  lastRoot = RegGet("TKGM", "RootDir")
  ' 上次目录无效或为程序目录时不作为默认值
  If lastRoot = "" Or IsSystemPath(lastRoot) Or Not fso.FolderExists(lastRoot) Then lastRoot = ""
  If lastRoot = "" Then
    ' 直接弹浏览框
    Do
      root = PickFolder("选择要处理的根文件夹", "")
      If root = "" Then Exit Do
      If IsSystemPath(root) Then
        MsgBox "选中的是 CAD/系统程序目录，请重新选择项目文件夹。"
        root = ""
      End If
    Loop Until root <> ""
  Else
    msg = InputBox("根文件夹 [回车=上次目录 浏览(B) 当前图纸目录(D) 选文件定位(F)]" & vbCrLf & vbCrLf & _
         "上次目录：" & lastRoot & vbCrLf & vbCrLf & _
         "回车=用上次目录；B=浏览；D=当前图纸目录；F=选文件定位；也可直接粘贴路径。", _
         "TKGPR v9.1 —— 选择根文件夹", lastRoot)
    root = ""
    If msg = "" Then
      root = lastRoot
    ElseIf UCase(Left(msg, 1)) = "B" Then
      root = PickFolder("选择要处理的根文件夹", lastRoot)
    ElseIf UCase(Left(msg, 1)) = "D" Then
      root = ActiveDocFolder()
      If root = "" Then MsgBox "取不到当前图纸目录（可能未保存，或位于 CAD 程序目录）。"
    ElseIf UCase(Left(msg, 1)) = "F" Then
      root = LocateByFile()
    Else
      If fso.FolderExists(msg) Then root = msg
    End If
    If IsSystemPath(root) Then
      MsgBox "选中的是 CAD/系统程序目录，请重新选择。"
      root = ""
    ElseIf root <> "" And Not fso.FolderExists(root) Then
      MsgBox "该文件夹不存在：" & root
      root = ""
    End If
  End If
  If root = "" Then
    If MsgBox("未选定有效根文件夹，是否重试？", 4 + 48, "TKGPR") = 7 Then WScript.Quit
  End If
Loop Until root <> ""
RegSet "TKGM", "RootDir", root

' ---------- 2. 工作包 XML ----------
xmlPath = ResolveXml(root)
If xmlPath = "" Then
  MsgBox "未选定工作包 XML，退出。"
  WScript.Quit
End If
RegSet "TKGM", "Xml", xmlPath

' ---------- 3. 改名设置 / 优先关键字 / 清理开关 / 比例 ----------
pat = InputBox("要改名的图框块名（支持通配符 * ?）" & vbCrLf & vbCrLf & _
               "回车=跳过改名；上次：" & RegGet("TKGM", "Pattern"), "TKGPR v9.1 —— 改名", RegGet("TKGM", "Pattern"))
If IsEmpty(pat) Then WScript.Quit
If pat = "" Then
  If RegGet("TKGM", "Pattern") <> "" Then
    ' 用户回车但上次有值 -> 询问是否沿用
    If MsgBox("沿用上次匹配模式 " & RegGet("TKGM", "Pattern") & " ？" & vbCrLf & "选“否”=跳过改名", 4 + 32, "TKGPR") = 7 Then
      pat = ""
    Else
      pat = RegGet("TKGM", "Pattern")
    End If
  End If
End If

newName = ""
If pat <> "" Then
  newName = InputBox("新图框块名" & vbCrLf & vbCrLf & "回车=按工作包项目编号自动生成（推荐）", "TKGPR v9.1 —— 新块名", "")
  If IsEmpty(newName) Then WScript.Quit
End If

' 优先关键字：一张图里多个图框都被图面用到时，用它定出唯一一个（与 HTA 对齐）
prefer = InputBox("优先关键字（一张图里多个图框都被用到时，按此关键字选定唯一一个）" & vbCrLf & vbCrLf & _
                  "例：20XXJS000CS-乙 或 乙   （子串匹配，不分大小写）" & vbCrLf & _
                  "留空 = 不干预：多个并存时跳过并报 MULTI_USED" & vbCrLf & vbCrLf & _
                  "上次：" & RegGet("TKGM", "Prefer"), "TKGPR v9.1 —— 优先关键字", RegGet("TKGM", "Prefer"))
If IsEmpty(prefer) Then WScript.Quit
prefer = Trim(prefer)
RegSet "TKGM", "Prefer", prefer

' 清理开关：只删「图面上 0 个实例」的匹配块定义，图面引用绝不触碰
doPurgeS = InputBox("是否清理「无用图框块定义」？" & vbCrLf & vbCrLf & _
                    "只删图面上一个实例都没有的匹配块定义（历次复制图纸残留的垃圾），" & vbCrLf & _
                    "图面正在引用的块绝不触碰。默认关闭。" & vbCrLf & vbCrLf & _
                    "Y = 清理      N 或留空 = 不清理", "TKGPR v9.1 —— 清理垃圾块", RegGet("TKGM", "DoPurge"))
If IsEmpty(doPurgeS) Then WScript.Quit
If UCase(Left(Trim(doPurgeS), 1)) = "Y" Then doPurge = "1" Else doPurge = "0"
RegSet "TKGM", "DoPurge", doPurge

scale = InputBox("图框比例（居中距离按此缩放）", "TKGPR v9.1 —— 比例", RegGet("TKGM", "Scale"))
If IsEmpty(scale) Then WScript.Quit
If scale = "" Then scale = "1"
RegSet "TKGM", "Pattern", pat
RegSet "TKGM", "NewName", newName
RegSet "TKGM", "Scale", scale

' ---------- 4. 收集 / 备份 ----------
Set files = CreateObject("Scripting.Dictionary")
CollectDwg root, files
cnt = files.Count
If cnt = 0 Then
  MsgBox "该文件夹下没有 DWG。"
  WScript.Quit
End If

stamp = Year(Now) & Right("0" & Month(Now),2) & Right("0" & Day(Now),2) & _
        Right("0" & Hour(Now),2) & Right("0" & Minute(Now),2) & Right("0" & Second(Now),2)
LOGF = root & "\TKGPR日志_" & stamp & ".txt"
LogLine "===== TKGPR v9.1 开始（wscript 备用入口）====="
LogLine "根文件夹：" & root
LogLine "工作包XML：" & xmlPath
LogLine "改名模式：" & pat & "   新名：" & newName
LogLine "优先关键字：" & prefer & "   清理垃圾块：" & doPurge & "   比例：" & scale

bakDir = root & "\备份_" & stamp
MkDirs bakDir
LogLine "备份目录：" & bakDir
For Each p In files
  rel = Mid(p, Len(root) + 2)
  dst = bakDir & "\" & rel
  If Not fso.FolderExists(fso.GetParentFolderName(dst)) Then
    MkDirs fso.GetParentFolderName(dst)
  End If
  fso.CopyFile p, dst, True
Next
LogLine "备份完成：" & cnt & " 张"

' ---------- 5. 逐张处理 ----------
i = 0
For Each p In files
  i = i + 1
  LogLine "[" & i & "/" & cnt & "] " & p
  On Error Resume Next
  Set od = FindOpenDoc(p)
  reused = False
  If Not (od Is Nothing) Then
    Set doc = od
    reused = True
    Err.Clear
  Else
    Set doc = docs.Open(p)
  End If
  If Err.Number <> 0 Then
    LogLine "   打开失败(" & Err.Number & ")：" & Err.Description
    Err.Clear
  Else
    doc.Activate
    WScript.Sleep 400
    If reused Then
      LogLine "   原已打开：复用现有文档并激活"
    Else
      LogLine "   已打开并激活"
    End If

    ' 载入 LISP 函数（供改名/清理/居中调用）
    Err.Clear
    doc.SendCommand "(load """ & Replace(lspFile, "\", "/") & """)" & vbCr
    WScript.Sleep 1500
    LogLine "   已载入 LISP(等待1.5s)"

    ' ① 改名
    If pat <> "" Then
      If fso.FileExists(renFlag) Then fso.DeleteFile renFlag, True
      Err.Clear
      doc.SendCommand "TKGREN" & vbCr
      found = WaitFlag(renFlag, 45)
      If Not found Then
        LogLine "   ①改名无响应，重发一次"
        Err.Clear
        doc.SendCommand "TKGREN" & vbCr
        found = WaitFlag(renFlag, 60)
      End If
      If found Then
        LogLine "   ①改名完成：" & Trim(ReadText(renFlag))
      Else
        LogLine "   ①改名超时"
      End If
    Else
      LogLine "   ①改名：跳过"
    End If

    ' ①b 清理无用图框块定义（可选，仅删 0 实例的块，图面引用不动）
    If doPurge = "1" Then
      If fso.FileExists(purgeFlag) Then fso.DeleteFile purgeFlag, True
      Err.Clear
      doc.SendCommand "TKGPURGE" & vbCr
      found = WaitFlag(purgeFlag, 45)
      If Not found Then
        LogLine "   ①b清理无响应，重发一次"
        Err.Clear
        doc.SendCommand "TKGPURGE" & vbCr
        found = WaitFlag(purgeFlag, 60)
      End If
      If found Then
        LogLine "   ①b清理完成：" & Trim(ReadText(purgeFlag))
      Else
        LogLine "   ①b清理超时"
      End If
    Else
      LogLine "   ①b清理：跳过（未开启）"
    End If

    ' ② 刷新图框（助手填工作包 + EZ_SMEDI_OPENTASKWITHFILE）
    If fso.FileExists(doneFlag) Then fso.DeleteFile doneFlag, True
    pf = fso.GetParentFolderName(helper) & "\_tgkm_openhelper.param"
    Set pfh = fso.CreateTextFile(pf, True)
    pfh.WriteLine xmlPath
    pfh.WriteLine doneFlag
    pfh.Close
    ws.Run "wscript.exe """ & helper & """", 0, False
    Err.Clear
    doc.SendCommand "EZ_SMEDI_OPENTASKWITHFILE" & vbCr
    found = WaitFlag(doneFlag, 180)
    If found Then
      LogLine "   ②刷新完成"
    Else
      LogLine "   ②刷新超时(180s)"
    End If

    ' ③ 折行居中
    If fso.FileExists(fixFlag) Then fso.DeleteFile fixFlag, True
    Err.Clear
    doc.SendCommand "TKGFIX" & vbCr
    found = WaitFlag(fixFlag, 45)
    If Not found Then
      LogLine "   ③居中无响应，重发一次"
      Err.Clear
      doc.SendCommand "TKGFIX" & vbCr
      found = WaitFlag(fixFlag, 60)
    End If
    If found Then
      LogLine "   ③居中完成（调整/跳过 = " & Trim(ReadText(fixFlag)) & "）"
    Else
      LogLine "   ③居中超时"
    End If

    ' ④ 存盘（失败重试一次）
    found = False
    For attempt = 1 To 2
      If Not found Then
        Err.Clear
        doc.Save
        If Err.Number = 0 Then
          found = True
          LogLine "   ④存盘完成"
        Else
          LogLine "   ④存盘失败(第" & attempt & "次) 号" & Err.Number & "：" & Err.Description
        End If
      End If
    Next

    ' ⑤ 关闭
    Err.Clear
    doc.Close
    If Err.Number <> 0 Then
      LogLine "   ⑤关闭失败(" & Err.Number & ")：" & Err.Description
      Err.Clear
    Else
      LogLine "   ⑤已关闭"
    End If
  End If
  On Error GoTo 0
Next

LogLine "===== 全部完成 ====="
MsgBox "批量处理完成：共 " & cnt & " 张。" & vbCrLf & _
       "改名模式：" & pat & "   优先关键字：" & prefer & "   清理垃圾块：" & doPurge & vbCrLf & vbCrLf & _
       "日志：" & LOGF, 0, "TKGPR v9.1"

' ==================== 工具 ====================

Sub MkDirs(p)
  Dim parent
  On Error Resume Next
  If p = "" Then Exit Sub
  If fso.FolderExists(p) Then Exit Sub
  MkDirs fso.GetParentFolderName(p)   ' 先递归建上级
  If Not fso.FolderExists(p) Then fso.CreateFolder p
  On Error GoTo 0
End Sub

Sub LogLine(s)
  Dim ff
  On Error Resume Next
  Set ff = fso.OpenTextFile(LOGF, 8, True)
  ff.WriteLine Time & "  " & s
  ff.Close
End Sub

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

Function WaitFlag(f, secs)
  Dim t1
  WaitFlag = False
  t1 = Timer
  Do While (Timer - t1 < secs)
    WScript.Sleep 500
    If fso.FileExists(f) Then
      WaitFlag = True
      Exit Function
    End If
  Loop
End Function

Function ReadText(f)
  Dim tx
  On Error Resume Next
  ReadText = ""
  If fso.FileExists(f) Then
    tx = fso.OpenTextFile(f, 1).ReadAll
    tx = Replace(Replace(Trim(tx), vbCr, ""), vbLf, "")
    ReadText = tx
  End If
  On Error GoTo 0
End Function

Function RegGet(key_, name)
  On Error Resume Next
  RegGet = ws.RegRead("HKCU\Software\" & key_ & "\" & name)
  If Err.Number <> 0 Then RegGet = ""
  On Error GoTo 0
End Function

Sub RegSet(key_, name, v)
  On Error Resume Next
  ws.RegWrite "HKCU\Software\" & key_ & "\" & name, v
  On Error GoTo 0
End Sub

Function PickFolder(title, startDir)
  Dim sh, fo2
  Set sh = CreateObject("Shell.Application")
  If startDir = "" Then startDir = 0
  Set fo2 = sh.BrowseForFolder(0, title, 0, startDir)
  If fo2 Is Nothing Then
    PickFolder = ""
  Else
    PickFolder = fo2.Self.Path
  End If
End Function

Function ActiveDocFolder()
  Dim d
  ActiveDocFolder = ""
  On Error Resume Next
  Set d = app.ActiveDocument
  If Err.Number = 0 Then
    If IsObject(d) Then ActiveDocFolder = d.Path
  End If
  On Error GoTo 0
  ' 未保存的新图会落在 CAD 程序目录，直接拒掉
  If IsSystemPath(ActiveDocFolder) Then
    MsgBox "当前图纸目录在 CAD 程序目录下（多半是未保存的新图）。" & vbCrLf & _
           "请改用 B=浏览 或 F=选文件定位。"
    ActiveDocFolder = ""
  End If
End Function

' 系统/程序目录判定（这类目录里的 xml 不是工作包）
Function IsSystemPath(p)
  Dim low
  IsSystemPath = False
  If p = "" Then
    IsSystemPath = True
    Exit Function
  End If
  low = LCase(p)
  If InStr(low, "program files") > 0 _
     Or InStr(low, "windows\") > 0 _
     Or InStr(low, "zwsoft") > 0 Then
    IsSystemPath = True
  End If
End Function

Function LocateByFile()
  Dim d, fs, fo3, lst, n2, k2
  d = PickFolder("选择包含目标图纸的文件夹", RegGet("TKGM", "RootDir"))
  If IsSystemPath(d) Then
    MsgBox "选中的是 CAD/系统程序目录，请重新选择项目文件夹。"
    LocateByFile = ""
    Exit Function
  End If
  If d = "" Then
    LocateByFile = ""
    Exit Function
  End If
  Set fs = fso.GetFolder(d)
  lst = ""
  n2 = 0
  For Each fo3 In fs.Files
    If LCase(fso.GetExtensionName(fo3.Name)) = "dwg" Then
      n2 = n2 + 1
      lst = lst & n2 & ") " & fo3.Name & vbCrLf
    End If
  Next
  If n2 = 0 Then
    LocateByFile = d
    Exit Function
  End If
  k2 = InputBox("该文件夹内的 DWG（取所选图纸所在文件夹作为根目录）：" & vbCrLf & vbCrLf & lst, "TKGPR", "1")
  LocateByFile = d
End Function

Sub AddXmls(dir_, dict)
  Dim fo4
  On Error Resume Next
  If dir_ = "" Then Exit Sub
  If Not fso.FolderExists(dir_) Then Exit Sub
  If IsSystemPath(dir_) Then Exit Sub   ' 跳过 CAD/系统程序目录里的无关 xml
  For Each fo4 In fso.GetFolder(dir_).Files
    If LCase(fso.GetExtensionName(fo4.Name)) = "xml" Then
      If Not dict.Exists(LCase(fo4.Path)) Then dict.Add LCase(fo4.Path), fo4.Path
    End If
  Next
  On Error GoTo 0
End Sub

Function ResolveXml(rootDir)
  Dim cands, k3, lst, idx, d, x, arr, i2
  Set cands = CreateObject("Scripting.Dictionary")
  AddXmls rootDir & "\工作包", cands
  AddXmls fso.GetParentFolderName(rootDir) & "\工作包", cands
  AddXmls rootDir, cands

  If cands.Count = 1 Then
    ResolveXml = cands.Items()(0)
    Exit Function
  End If
  If cands.Count > 1 Then
    lst = ""
    arr = cands.Items()
    For i2 = 0 To cands.Count - 1
      lst = lst & (i2 + 1) & ") " & arr(i2) & vbCrLf
    Next
    k3 = InputBox("找到多个工作包 XML，请输入序号：" & vbCrLf & vbCrLf & lst, "TKGPR", "1")
    If IsNumeric(k3) Then
      idx = CInt(k3)
      If idx >= 1 And idx <= cands.Count Then
        ResolveXml = arr(idx - 1)
        Exit Function
      End If
    End If
  End If

  d = PickFolder("选择「工作包」文件夹", rootDir)
  If d <> "" Then
    AddXmls d, cands
    If cands.Count >= 1 Then
      ResolveXml = cands.Items()(cands.Count - 1)
      Exit Function
    End If
  End If

  x = InputBox("请输入工作包 XML 的完整路径（文件夹 + 文件名）", "TKGPR", rootDir & "\工作包\")
  If x <> "" And fso.FileExists(x) Then
    ResolveXml = x
  Else
    ResolveXml = ""
  End If
End Function

Sub CollectDwg(folder, dict)
  Dim f2, s2
  For Each f2 In fso.GetFolder(folder).Files
    If LCase(fso.GetExtensionName(f2.Name)) = "dwg" Then
      If Left(f2.Name, 2) <> "~$" Then dict.Add f2.Path, f2.Path
    End If
  Next
  For Each s2 In fso.GetFolder(folder).SubFolders
    If Left(s2.Name, 2) <> "备份" Then CollectDwg s2.Path, dict
  Next
End Sub
