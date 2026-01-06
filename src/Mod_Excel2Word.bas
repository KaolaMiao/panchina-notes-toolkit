Attribute VB_Name = "Mod_Excel2Word"
' =================================================================
' Copyright (C) 2025 Dota (PCCPA). All Rights Reserved.
' 本程序受著作权法和国际条约保护。
' 未经授权的复制或分发本程序（或其中任何部分），将导致严厉的民事和刑事处罚。
' =================================================================
' =================================================================
' 模块名称: 附注EXCEL2WORD工具 V3.3
' 作者：Dota | 天健会计师事务所 (PCCPA)
' 更新日志:
'   1211: 修复引用库报错的问题
'   1120: 增加可视化分屏、取消焦点强制会聚、表格粘贴反馈及未成功标记
' 功能描述: 实现Excel数据到Word报告的自动化生成、更新及校验
' =================================================================

Option Explicit

' 在模块最顶部 (Option Explicit 下方) 添加这两个全局变量
Private g_ExcelAppHandler As ExcelAppEventHandler
Private g_WordAppHandler As WordEventHandler
' 放在 Option Explicit 下方
' =================================================================
' Windows API 统一声明
' 将此代码块放在模块的最顶部
' =================================================================

#If VBA7 Then
    ' 64位 Office (VBA7) 兼容声明
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
    ' 32位 Office (旧版) 兼容声明
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
#End If

Private Const SW_RESTORE = 9

Private Const VISUAL_DELAY As Long = 200
Private Const SHEET_TEMPLATE As String = "附注模板"

' =================================================================
' 【第二部分】 全局变量
' 作用：在不同函数间共享状态（如表格缓存、统计数据、取消标记）
' =================================================================
Public g_tableRanges As Object
Public g_tableCellMap As Object
Public g_tablesToUpdate As Object
Public g_UserCancelled As Boolean
Public g_VisualMode As Boolean
' 在 Mod_Excel2Word 模块顶部声明
Public g_Ribbon As IRibbonUI

' 用于存储窗体返回的结果
Public g_ConfirmResult As VbMsgBoxResult

' 统计变量
Public g_totalTablesAttempted As Long
Public g_totalTablesSucceeded As Long
' 统计变量 (新增 - 用于单元格级统计)
Public g_cellChangeCount As Long ' 成功更新到Word的单元格数
Public g_cellSkipCount As Long   ' 有差异但因格式限制未更新(已标色)的单元格数

' =================================================================
' 【第三部分】 Ribbon 接口层 (功能入口)
' 作用：直接绑定 Excel 功能区按钮，负责接收点击事件，不含复杂逻辑
' =================================================================

' [入口 1] 核心：一键生成 Word 文档
Public Sub GenerateWordDocumentFromExcel(control As Object)
    Call GenerateWordDocumentLogic ' 调用下方的核心逻辑
End Sub

' [入口 2] 替换/更新类
Public Sub ReplaceSingleTable(control As Object)
    Call GenericTableProcessor("REPLACE_SINGLE") ' 弹窗单选替换
End Sub

Public Sub BatchReplaceTables(control As Object)
    Call GenericTableProcessor("REPLACE_BATCH")  ' 弹窗多选批量替换
End Sub

Public Sub UpdateTableDataOnly(control As Object)
    Call GenericTableProcessor("UPDATE_FULL")    ' 仅更新数据
End Sub

Public Sub UpdateSelectedAreaInWord(control As Object)
    Call GenericTableProcessor("UPDATE_PARTIAL") ' 选区局部更新
End Sub

' [入口 3] 辅助工具类
Public Sub QuickAutoTagTables(control As Object)
    Dim ws As Worksheet
    Set ws = GetWorksheetSafe(SHEET_TEMPLATE)
    If ws Is Nothing Then Exit Sub
    Call PreCheckAndTagTables(ws)
End Sub

Public Sub AutoTagByBorders(control As Object)
    Call AutoTagByBordersLogic ' 智能边框识别
End Sub

Public Sub VerifyWordTableConsistency(control As Object)
    Call VerifyWordTableConsistencyLogic ' 一致性校验
End Sub

Public Sub LocateAndValidateTable(control As Object)
    Call LocateTableLogic ' 双向定位
End Sub

Public Sub CleanSelectedWordTable(control As Object)
    Call CleanWordTableLogic ' 清洗Word表格空行
End Sub

Public Sub ResetWindowLayout(control As Object)
    Call ResetWindowLayoutLogic ' 重置窗口
End Sub

Public Sub CheckWordNumbering(control As Object)
    Call AutoFixNumberingLogic
End Sub



' [入口 4] 信息展示类
Public Sub ShowAboutInfo(control As Object)
    Dim msg As String
    msg = "--------------------------------------------------" & vbCrLf & _
          "       附注 EXCEL 2 WORD 自动化工具箱        " & vbCrLf & _
          "--------------------------------------------------" & vbCrLf & vbCrLf & _
          "版本：Ver 2.1 (试用版)" & vbCrLf & _
          "作者：Dota | 天健会计师事务所 (PCCPA) | 二总" & vbCrLf & _
          "日期：2025-11" & vbCrLf & vbCrLf & _
          "作者寄语：" & vbCrLf & _
          "报错是正常的，不报错是幸运的。 " & vbCrLf & _
          "运行前请深呼吸，心中默念‘借贷必相等’三次， 程序成功率可提升 50%。" & vbCrLf & vbCrLf & _
          "格言：" & vbCrLf & _
          "“天行健，君子以自强不息，”" & vbCrLf & _
          "“地势坤，审计以工具护体。”" & vbCrLf & vbCrLf & _
          "声明：" & vbCrLf & _
          "本工具仅供内部业务交流与效能提升使用。"
    MsgBox msg, vbInformation, "关于本工具"
End Sub

Public Sub ShowUserGuide(control As Object)
    Dim msg As String
    msg = "【标准作业流程 (SOP)】" & vbCrLf & vbCrLf & _
          "1. 初稿生成 (前期)" & vbCrLf & _
          "   点击 [生成文档]。程序将根据Excel模板，一键生成带有基础格式的Word附注初稿。" & vbCrLf & vbCrLf & _
          "2. 核对与补漏 (中期)" & vbCrLf & _
          "   A. 点击 [一致性校验]。系统会自动比对Excel与Word表格数量。" & vbCrLf & _
          "   B. 若提示缺失，或Word中表格未正确生成（如书签丢失）：" & vbCrLf & _
          "      请在Word正确位置手动插入一个空表格并拆入标签（或定位到旧表格），" & vbCrLf & _
          "      然后在Excel选中对应表格，点击 [替换单表] 或 [批量替换] 完成修复。" & vbCrLf & vbCrLf & _
          "3. 数据刷新 (后期·强烈推荐)" & vbCrLf & _
          "   当Word格式（边框、字体、行高）已手工精调完毕，仅需更新数字时：" & vbCrLf & _
          "   请务必使用 [全表刷新] 或 [选区更新]。" & vbCrLf & _
          "   ★ 优势：这会保留您在Word中辛苦调整的所有样式，仅替换单元格内的文字/数值。" & vbCrLf & vbCrLf & _
          "4. 辅助工具" & vbCrLf & _
          "   - 表格清洗：去除Excel带入的顽固空行和回车。" & vbCrLf & _
          "   - 定位校验：双向跳转，快速核对底稿。"
    MsgBox msg, vbInformation, "操作指南"
End Sub

' =================================================================
' 【第四部分】 核心业务控制器 (Controller)
' 作用：负责调度整个业务流程，初始化环境，处理用户交互，调用底层Worker
' =================================================================

' [业务 A] 一键生成文档主逻辑
Private Sub GenerateWordDocumentLogic()
    Dim wdApp As Object, wdDoc As Object, ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim identifier As String, content As String
    Dim docPath As String, baseName As String
    
        ' --- [新增变量] ---
    Dim startRow As Long     ' 起始行号
    Dim userResp As String   ' 用户输入
    Dim saveInterval As Long ' 自动保存间隔
    Dim promptMsg As String  ' 提示信息
    
    saveInterval = 100        ' 设定：每处理 100 行自动保存一次
    
    'On Error GoTo ErrorHandler
    
    ' 初始化
    g_totalTablesAttempted = 0: g_totalTablesSucceeded = 0
    ForceAppFocus Application.hwnd
    
    ' 【核心修改】加载配置单例 (支持多模板)
    Call Mod_Financial_Core.Init_Global_Config
    
    ' 环境准备
    Set ws = GetWorksheetSafe(SHEET_TEMPLATE)
    If ws Is Nothing Then Exit Sub
    Call ClearPreviousHighlights(ws)
    Call PreCheckAndTagTables(ws)
    If Not PreCacheTableRanges(ws) Then Exit Sub
    
    
    ' =============================================================
    ' [修改点] 断点续传询问：明确告知需要手工拼接
    ' =============================================================
    promptMsg = "请输入开始处理的 Excel 行号：" & vbCrLf & vbCrLf & _
                "【模式选择】" & vbCrLf & _
                "全量生成：请输入 1 （默认）" & vbCrLf & _
                "断点续传：请输入上次中断的行号" & vbCrLf & vbCrLf & _
                "重要提示 (必读) " & vbCrLf & _
                "断点续传将生成一个【新的独立文档】。" & vbCrLf & _
                "程序不会自动追加到旧文档末尾，任务完成后：" & vbCrLf & _
                "请您【手工】将新文档的内容拼接到旧文档中。"

    startRow = 1
    userResp = InputBox(promptMsg, "生成设置 (含拼接警告)", "1")
    
    ' 校验输入
    If StrPtr(userResp) = 0 Then Exit Sub ' 点击取消则退出
    If IsNumeric(userResp) Then
        startRow = CLng(userResp)
        If startRow < 1 Then startRow = 1
    Else
        MsgBox "输入无效，将从第 1 行开始。", vbExclamation
        startRow = 1
    End If
    
    
    Call AskForVisualMode
    
    ' 【关键修改】加载全局配置 (包含 User_Config 中的样式参数)
    ' 确保这里调用了 Init 和 Load，将 Excel 中的数据读入内存 g_TableConfig
    Call Mod_Financial_Core.Init_Global_Config
    
    ' 启动 Word
    If Not InitWordApp(wdApp, wdDoc, CreateNew:=True) Then Exit Sub
    If g_VisualMode Then ArrangeWindowsSplitScreen wdApp Else wdApp.WindowState = 1
    ToggleSystemUpdates False, wdApp
    
    ' =============================================================
    ' [修复] 强制显示状态栏
    ' =============================================================
    Application.DisplayStatusBar = True  ' 确保状态栏本身是可见的
    Application.StatusBar = "正在初始化..." ' 先显示一句测试文本
    DoEvents ' 强制重绘 UI
    
    ' 无论 Word 模板里有没有，这里都会根据 Excel 配置重写样式
    Application.StatusBar = "正在同步文档样式..."
    Call Ensure_Word_Style(wdDoc, "H")
    Call Ensure_Word_Style(wdDoc, "H0")
    Call Ensure_Word_Style(wdDoc, "H1")
    Call Ensure_Word_Style(wdDoc, "H2")
    Call Ensure_Word_Style(wdDoc, "H3")
    Call Ensure_Word_Style(wdDoc, "P")
    Call Ensure_Word_Style(wdDoc, "D")
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    ' 计算保存路径 (提前计算，用于自动保存)
    baseName = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
    ' 文件名加上 起始行号 标记，方便用户区分
    docPath = ActiveWorkbook.path & "\" & baseName & "_StartRow" & startRow & "_" & Format(Now, "MMDD_HHMMSS") & ".docx"
    ' 先保存一次空文档，确保路径存在，后续只需 .Save
    wdDoc.SaveAs2 docPath
    
    ' =============================================================
    ' 循环处理每一行
    ' =============================================================
    Dim percent As Double ' 新增变量
    
    For i = 1 To lastRow
        
        ' [断点续传] 跳过起始行
        If i < startRow Then GoTo NextIteration
        
        identifier = UCase(Trim(ws.Cells(i, 1).Value))
        If identifier <> "" Then
            content = CStr(ws.Cells(i, 2).Value)
            Select Case identifier
                Case "H", "H0", "H1", "H2", "H3", "P", "B", "D"
                    ProcessTextParagraph wdDoc, identifier, content
                Case "TABLE"
                    ProcessTableInsertion wdDoc, ws.Cells(i, 1)
            End Select
            
            ' =========================================================
            ' [修复 2] 强力进度更新逻辑
            ' =========================================================
            ' 1. 计算进度 (强制转为 Double 防止整数除法精度丢失)
            percent = (CDbl(i) - startRow + 1) / (CDbl(lastRow) - startRow + 1)
            If percent > 1 Then percent = 1
            If percent < 0 Then percent = 0
            
            ' 2. 调试：请按 Ctrl+G 打开立即窗口，看看这里有没有数字在跳动
            ' 如果这里在动但界面不动，说明是界面卡死；如果这里不动，是逻辑错。
            Debug.Print "Row: " & i & " | Pct: " & Format(percent, "0%")
            
            ' 3. 更新界面
            If g_VisualMode Then
                wdDoc.Application.ScreenRefresh
                DoEvents
            Else
                ' 每次都更新，不再跳过
                UpdateVisualStatusBar i, lastRow, percent
                
                ' 【关键】给予 Excel 喘息时间重绘界面
                DoEvents
            End If
            ' =========================================================
            
            ' [自动保存]
            If i Mod saveInterval = 0 Then
                Application.StatusBar = "正在自动保存..." ' 加个图标更显眼
                wdDoc.Save
                wdDoc.UndoClear
                ' 保存完后立即恢复进度条显示
                UpdateVisualStatusBar i, lastRow, percent
            End If
            
        End If

NextIteration:
    Next i
    
    ' 最后再保存一次
    wdDoc.Save
    
    ForceAppFocus Application.hwnd
    
    ' 生成报告
    Dim reportMsg As String, msgIcon As VbMsgBoxStyle
    reportMsg = "文档生成完毕！" & vbCrLf & _
                "保存路径：" & docPath & vbCrLf & vbCrLf & _
                "------ 任务统计 ------" & vbCrLf & _
                "处理范围: 第 " & startRow & " 行 至 第 " & lastRow & " 行" & vbCrLf & _
                "计划表格: " & g_totalTablesAttempted & " 个" & vbCrLf & _
                "成功表格: " & g_totalTablesSucceeded & " 个"
    
    ' 如果是断点续传，再次提醒拼接
    If startRow > 1 Then
        reportMsg = reportMsg & vbCrLf & vbCrLf & _
                    "【拼接提醒】" & vbCrLf & _
                    "您使用了断点续传，请记得将此文档内容" & vbCrLf & _
                    "手工复制到旧文档的末尾。"
    End If
    
    msgIcon = vbInformation
    If g_totalTablesSucceeded < g_totalTablesAttempted Then
        reportMsg = reportMsg & vbCrLf & vbCrLf & "【注意】" & vbCrLf & (g_totalTablesAttempted - g_totalTablesSucceeded) & " 个表格粘贴失败！" & vbCrLf & "请返回Excel工作表，查看红色高亮的单元格。"
        msgIcon = vbExclamation
    End If
    MsgBox reportMsg, msgIcon, "处理完成"

SafeExit:
    Cleanup wdApp
    Exit Sub
ErrorHandler:
    HandleRuntimeError "GenerateWordDocumentFromExcel"
    Resume SafeExit
End Sub

' [业务 B] 通用表格处理器 (替换/批量/更新)
Private Sub GenericTableProcessor(ByVal mode As String)
    Dim wdApp As Object, wdDoc As Object
    Dim ws As Worksheet, intersectDict As Object
    Dim updatedCount As Long, skippedCount As Long
    Dim k As Variant, tName As String
    Dim userConfirm As Long, batchConfirmMode As Boolean
    
    'On Error GoTo ErrorHandler
    
    ' 初始化
    g_totalTablesAttempted = 0: g_totalTablesSucceeded = 0
    g_cellChangeCount = 0: g_cellSkipCount = 0 ' <--- 新增初始化
    
    ForceAppFocus Application.hwnd
    
    ' 【核心修改】加载配置单例
    Call Mod_Financial_Core.Init_Global_Config
    
    Set ws = GetWorksheetSafe(SHEET_TEMPLATE)
    If ws Is Nothing Then Exit Sub
    If Not PreCacheTableRanges(ws) Then Exit Sub
    Call ClearPreviousHighlights(ws)
    
    ' 模式选择与交互
    Select Case mode
        Case "REPLACE_SINGLE"
            If Not GetUserSelection(AllowMultiSelect:=False) Then Exit Sub
        Case "REPLACE_BATCH"
            If Not GetUserSelection(AllowMultiSelect:=True) Then Exit Sub
            ' 批量模式确认逻辑
            Dim response As VbMsgBoxResult
            response = MsgBox("批量替换模式选择：" & vbCrLf & vbCrLf & "【是(Yes)】：逐个确认模式 (推荐)" & vbCrLf & "【否(No)】：极速模式" & vbCrLf & "【取消(Cancel)】：退出操作", vbYesNoCancel + vbQuestion, "操作确认")
            If response = vbCancel Then Exit Sub
            batchConfirmMode = (response = vbYes)
            If batchConfirmMode Then g_VisualMode = True
        Case "UPDATE_FULL"
            If Not GetUserSelection(AllowMultiSelect:=True) Then Exit Sub
            batchConfirmMode = True
        Case "UPDATE_PARTIAL"
            If TypeName(Selection) <> "Range" Then MsgBox "请先在 Excel 中选择区域", vbExclamation: Exit Sub
            Set intersectDict = GetIntersectDictionary(Selection)
            If intersectDict.count = 0 Then MsgBox "未选中任何有效表格区域。", vbExclamation: Exit Sub
            If MsgBox("选区涉及 " & intersectDict.count & " 个表格，是否继续？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End Select
    
    If mode <> "REPLACE_BATCH" Or (mode = "REPLACE_BATCH" And batchConfirmMode = False) Then Call AskForVisualMode
    
    ' 启动 Word
    If Not InitWordApp(wdApp, wdDoc, CreateNew:=False) Then Exit Sub
    If g_VisualMode Then ArrangeWindowsSplitScreen wdApp Else wdApp.WindowState = 1
    ToggleSystemUpdates False, wdApp
    
    ' 核心执行循环
    Select Case mode
        ' --- 场景: 单表替换 ---
        Case "REPLACE_SINGLE"
            For Each k In g_tablesToUpdate: tName = CStr(k): Exit For: Next k
            Application.StatusBar = "正在处理: " & tName
            If ValidateTableExists(wdDoc, tName) Then
                If g_VisualMode Then FocusAndScrollTo wdDoc, wdDoc.Bookmarks(tName).Range
                userConfirm = MsgBox("已定位表格: [" & tName & "]" & vbCrLf & "模式: 整表替换" & vbCrLf & "确认执行？", vbYesNo + vbQuestion, "确认")
                If userConfirm = vbNo Then MsgBox "操作已取消", vbInformation: GoTo SafeExit
                If g_VisualMode Then ForceAppFocus wdApp.ActiveWindow.hwnd
                
                g_totalTablesAttempted = 1
                If ReplaceTableCore(wdDoc, g_tableRanges(tName), tName) Then g_totalTablesSucceeded = 1 Else g_tableCellMap(tName).Interior.Color = vbRed
            End If
            
        ' --- 场景: 批量替换 ---
        Case "REPLACE_BATCH"
            For Each k In g_tablesToUpdate
                tName = CStr(k)
                Application.StatusBar = "批量处理: " & tName
                If ValidateTableExists(wdDoc, tName, Silent:=True) Then
                    If batchConfirmMode Then
                        FocusAndScrollTo wdDoc, wdDoc.Bookmarks(tName).Range
                        ForceAppFocus Application.hwnd
                        ' ================= 修改点开始 =================
                        ' 1. 确保 Excel 可见且处于激活状态，方便用户操作
                        Application.ScreenUpdating = True
                        If g_VisualMode Then ForceAppFocus Application.hwnd
                        
                        ' 2. 使用非模态自定义函数替代 MsgBox
                        ' 此时代码会暂停在这里，但你可以去滚动 Excel
                        userConfirm = AskUserModeless("即将替换表格: [" & tName & "]" & vbCrLf & vbCrLf & _
                                                      "提示：您现在可以滚动 Excel 查看内容，" & vbCrLf & _
                                                      "确认无误后点击下方按钮。", "逐项确认中...")
                        
                        ' 3. 恢复可能的屏幕冻结 (如果后续代码需要)
                        If Not g_VisualMode Then Application.ScreenUpdating = False
                        ' ================= 修改点结束 =================
                        If userConfirm = vbCancel Then GoTo SafeExit
                        If userConfirm = vbNo Then GoTo NextBatchItem
                        If g_VisualMode Then ForceAppFocus wdApp.ActiveWindow.hwnd
                    End If
                    
                    g_totalTablesAttempted = g_totalTablesAttempted + 1
                    If ReplaceTableCore(wdDoc, g_tableRanges(tName), tName) Then
                        g_totalTablesSucceeded = g_totalTablesSucceeded + 1: updatedCount = updatedCount + 1
                    Else
                        g_tableCellMap(tName).Interior.Color = vbRed
                    End If
                Else
                    skippedCount = skippedCount + 1
                End If
NextBatchItem:
            Next k
            
        ' --- 场景: 全量数据更新 ---
        Case "UPDATE_FULL"
            For Each k In g_tablesToUpdate
                tName = CStr(k)
                Application.StatusBar = "数据更新: " & tName
                If ValidateTableExists(wdDoc, tName) Then
                    If g_VisualMode Then FocusAndScrollTo wdDoc, wdDoc.Bookmarks(tName).Range
                    userConfirm = MsgBox("即将更新表格: [" & tName & "]" & vbCrLf & "模式: 仅更新数值" & vbCrLf & "是否执行？", vbYesNo + vbQuestion, "逐表确认")
                    If userConfirm = vbNo Then MsgBox "操作终止", vbInformation: GoTo SafeExit
                    If g_VisualMode Then ForceAppFocus wdApp.ActiveWindow.hwnd
                    
                    PerformFullDataUpdate wdDoc, g_tableRanges(tName), tName
                    updatedCount = updatedCount + 1
                Else
                    skippedCount = skippedCount + 1
                End If
            Next k
            
        ' --- 场景: 局部更新 ---
        Case "UPDATE_PARTIAL"
            For Each k In intersectDict
                tName = CStr(k)
                Application.StatusBar = "局部更新: " & tName
                If ValidateTableExists(wdDoc, tName) Then
                    If g_VisualMode Then FocusAndScrollTo wdDoc, wdDoc.Bookmarks(tName).Range
                    userConfirm = MsgBox("即将更新表格: [" & tName & "]" & vbCrLf & "确认更新区域？", vbYesNo + vbQuestion, "确认")
                    If userConfirm = vbNo Then MsgBox "操作终止", vbInformation: GoTo SafeExit
                    If g_VisualMode Then ForceAppFocus wdApp.ActiveWindow.hwnd
                    
                    PerformPartialDataUpdate wdDoc, g_tableRanges(tName), intersectDict(tName), tName
                    updatedCount = updatedCount + 1
                End If
            Next k
    End Select
    
    ' 结束报告
    ForceAppFocus Application.hwnd
    Dim reportMsg As String, msgIcon As VbMsgBoxStyle
    msgIcon = vbInformation
    
If mode = "UPDATE_FULL" Or mode = "UPDATE_PARTIAL" Then
        ' 数据更新模式的详细报告
        reportMsg = "数据刷新完成！" & vbCrLf & vbCrLf & _
                    "表格统计：" & vbCrLf & _
                    "-覆盖表格: " & updatedCount & " 个" & vbCrLf & vbCrLf & _
                    "单元格明细：" & vbCrLf & _
                    "成功更新: " & g_cellChangeCount & " 处"
        
        ' 如果有跳过的差异，重点提示
        If g_cellSkipCount > 0 Then
            reportMsg = reportMsg & vbCrLf & _
                        " 发现差异但未更新: " & g_cellSkipCount & " 处" & vbCrLf & vbCrLf & _
                        "【注意】" & vbCrLf & _
                        "上述未更新的单元格（通常为文本/表头修改）已在 Excel 中" & _
                        "标为【橙色】。请检查 Excel 并手动修改 Word 对应位置。"
            msgIcon = vbExclamation
        Else
            reportMsg = reportMsg & vbCrLf & " - 所有数值均已同步。"
        End If
        
        MsgBox reportMsg, msgIcon, "更新报告"
    Else
        ' 替换模式的报告 (保持原有逻辑)
        reportMsg = "处理完成！" & vbCrLf & "------ 任务统计 ------" & vbCrLf & "计划处理: " & g_totalTablesAttempted & " 个" & vbCrLf & "成功处理: " & g_totalTablesSucceeded & " 个"
        If g_totalTablesSucceeded < g_totalTablesAttempted Then
            reportMsg = reportMsg & vbCrLf & vbCrLf & "【注意】" & (g_totalTablesAttempted - g_totalTablesSucceeded) & " 个表格处理失败！请查看Excel标红单元格。"
            msgIcon = vbExclamation
        End If
        MsgBox reportMsg, msgIcon, "处理完成"
    End If
    
SafeExit:
    Cleanup wdApp
    Exit Sub
ErrorHandler:
    HandleRuntimeError "GenericTableProcessor"
    Resume SafeExit
End Sub

' =================================================================
' 【第五部分】 业务逻辑实现 - 表格操作 (Worker Functions)
' 作用：执行具体的表格复制、粘贴、数据更新、格式清洗
' =================================================================

' [Logic] 替换表格核心函数
Private Function ReplaceTableCore(doc As Object, dataRange As Range, tableName As String, Optional targetRange As Object = Nothing) As Boolean
    ReplaceTableCore = False
    Dim insertRange As Object
    
    If g_VisualMode Then FocusAndScrollTo doc, insertRange
    
    ' 确定位置
    If targetRange Is Nothing Then
        If doc.Bookmarks.exists(tableName) Then
            Set insertRange = doc.Bookmarks(tableName).Range
            If insertRange.Tables.count > 0 Then insertRange.Tables(1).Delete
        Else
            Exit Function
        End If
    Else
        Set insertRange = targetRange
    End If
    
    ' 1. 调用安全复制
    If Not CopyRangeSafe(dataRange) Then
        MsgBox "复制失败: " & tableName & vbCrLf & "剪贴板可能被其他程序(如迅雷/旺旺)占用。", vbCritical
        Exit Function
    End If
    
    ' 2. 调用安全粘贴
    Dim isSuccess As Boolean
    isSuccess = PasteTableSafe(insertRange)
    
    If Not isSuccess Then
        MsgBox "表格 [" & tableName & "] 粘贴失败！" & vbCrLf & "Word 无法接收数据，请手动检查。", vbExclamation
        ' 可以在这里加个断点或 Stop，方便调试
        Exit Function
    End If
    
        ' 3. 粘贴成功后的处理（格式化、设书签等）
    Application.CutCopyMode = False ' 清空剪贴板，释放内存
    
    
    ' 后处理
    Dim tbl As Object
    On Error Resume Next
    If insertRange.Tables.count > 0 Then Set tbl = insertRange.Tables(1) Else Set tbl = doc.Range(insertRange.Start, insertRange.Start + 300).Tables(1)
    On Error GoTo 0
    
    If Not tbl Is Nothing Then
        doc.Bookmarks.Add Name:=tableName, Range:=tbl.Range
        FormatWordTable tbl ' <--- 这里会应用新样式
        If g_VisualMode Then
            On Error Resume Next
            tbl.Select: doc.Application.ScreenRefresh: DoEvents: Sleep VISUAL_DELAY * 3
            On Error GoTo 0
        End If
        ReplaceTableCore = True
    End If
End Function

' [Logic] 表格插入占位处理 (文档生成时调用)
Private Sub ProcessTableInsertion(doc As Object, cell As Range)
    If cell.Comment Is Nothing Then Exit Sub
    Dim txt As String, tName As String
    txt = Replace(cell.Comment.Text, "：", ":")
    
    If InStr(1, txt, "TableName:", vbTextCompare) > 0 Then
        tName = Trim(Split(Split(txt, "TableName:")(1), vbCrLf)(0))
        If g_tableRanges.exists(tName) Then
            g_totalTablesAttempted = g_totalTablesAttempted + 1
            Dim para As Object
            Set para = doc.content.Paragraphs.Add
            If ReplaceTableCore(doc, g_tableRanges(tName), tName, para.Range) Then
                g_totalTablesSucceeded = g_totalTablesSucceeded + 1
            Else
                cell.Interior.Color = vbRed
            End If
        End If
    End If
End Sub

' [更新] 文本段落处理
' 功能：将 Excel 内容写入 Word 并应用对应样式
Private Sub ProcessTextParagraph(doc As Object, id As String, content As String)
    Dim Sel As Object
    Dim styleName As String
    Dim cleanContent As String
    
    Set Sel = doc.Application.Selection
    
    ' 1. 映射样式名 (必须与 Ensure_Word_Style 中的名称一致)
    Select Case id
        Case "H":  styleName = "附注-标题"
        Case "H0": styleName = "附注-小标题"
        Case "H1": styleName = "附注-一级标题"
        Case "H2": styleName = "附注-二级标题"
        Case "H3": styleName = "附注-三级标题"
        Case "P", "B": styleName = "附注-正文"
        Case "D":  styleName = "附注-单位说明"
        Case Else: styleName = "Normal"
    End Select
    
    Sel.EndKey Unit:=6 ' wdStory
    
    ' 空内容处理
    If Trim(content) = "" Then
        Sel.TypeParagraph
        Exit Sub
    End If
    
    cleanContent = Replace(Replace(content, vbCr, ""), vbLf, "")
    Sel.TypeText Text:=cleanContent
    
    ' 2. 应用样式
    On Error Resume Next
    Sel.Paragraphs(1).Style = doc.Styles(styleName)
    
    ' =======================================================
    ' 【新增双重保险】
    ' 无论样式是否正确，都强制为当前段落直接设置一次大纲级别。
    ' 这可以覆盖任何来自样式的错误继承。
    ' =======================================================
    Dim finalOutlineLevel As Long
    Select Case id
        Case "H", "H1": finalOutlineLevel = 1
        Case "H2": finalOutlineLevel = 2
        Case "H3": finalOutlineLevel = 3
        Case Else: finalOutlineLevel = 10 ' 正文级别
    End Select
    On Error Resume Next
    Sel.Paragraphs(1).Format.OutlineLevel = finalOutlineLevel
    On Error GoTo 0
    ' =======================================================
    
    ' 3. 特殊处理：Tag "B" 为加粗正文
    ' 虽然 B 使用了 P 的样式，但需要额外加粗
    If id = "B" Then
        Sel.Paragraphs(1).Range.Font.Bold = True
    End If
    On Error GoTo 0
    
    ' 4. 辅助识别 (如果 Excel 是 P 但内容像标题，自动升级大纲级别，但不改样式)
    ' 仅调整 OutlineLevel 以便于导航，不改变外观
    If id = "P" Or id = "B" Then
        Dim outlineLvl As Long: outlineLvl = 10
        Dim reg As Object
        Set reg = CreateObject("VBScript.RegExp")
        reg.Global = False: reg.IgnoreCase = True
        
        ' (一) -> Level 2
        reg.Pattern = "^[（\(][一二三四五六七八九十]+[）\)]"
        If reg.Test(cleanContent) Then outlineLvl = 2
        
        ' 1. -> Level 3
        reg.Pattern = "^\d+\."
        If reg.Test(cleanContent) Then outlineLvl = 3
        
        ' (1) -> Level 4
        reg.Pattern = "^[（\(]\d+[）\)]"
        If reg.Test(cleanContent) Then outlineLvl = 4
        
        If outlineLvl <> 10 Then
            On Error Resume Next
            Sel.Paragraphs(1).Format.OutlineLevel = outlineLvl
            On Error GoTo 0
        End If
        Set reg = Nothing
    End If
    
    Sel.TypeParagraph
    
    ' 重置下一段为正文默认 (防止样式连带)
    On Error Resume Next
    Sel.ParagraphFormat.OutlineLevel = 10
    Sel.Style = doc.Styles("Normal")
    On Error GoTo 0
End Sub

' [Logic] 全表数据更新
' [Logic] 全表数据更新 (优化版：智能跳过合并单元格的附属区域)
Private Sub PerformFullDataUpdate(doc As Object, dataRange As Range, tableName As String)
    If Not ValidateTableDimensions(doc, dataRange, tableName) Then Exit Sub
    
    Dim wdTbl As Object
    Dim r As Long, c As Long
    Dim arrData As Variant
    Dim curCell As Range
    Dim isMerged As Boolean
    Dim isTopLeft As Boolean
    
    Set wdTbl = doc.Bookmarks(tableName).Range.Tables(1)
    
    ' 1. 将数值读入数组 (为了速度，仍然保留数组读取)
    If dataRange.Cells.count = 1 Then
        ReDim arrData(1 To 1, 1 To 1): arrData(1, 1) = dataRange.Value
    Else
        arrData = dataRange.Value
    End If
    
    On Error Resume Next
    
    ' 2. 遍历数据
    For r = 1 To UBound(arrData, 1)
        For c = 1 To UBound(arrData, 2)
            
            ' 获取当前单元格对象用于判断合并状态
            ' 注意：dataRange.Cells(r, c) 虽然比直接读数组慢一点点，
            ' 但为了判断 MergeCells 属性，这是必须的。
            ' 相比于减少了 Word 的无效操作，这点开销非常值得。
            Set curCell = dataRange.Cells(r, c)
            
            isMerged = curCell.MergeCells
            isTopLeft = False
            
            If isMerged Then
                ' 如果是合并单元格，判断是否为左上角第一个
                ' 原理：只有左上角的单元格地址 = 整个合并区域第一个单元格的地址
                If curCell.Address = curCell.MergeArea.Cells(1, 1).Address Then
                    isTopLeft = True
                End If
            End If
            
            ' =========================================================
            ' 【核心优化逻辑】
            ' 1. 如果不是合并单元格 -> 正常更新
            ' 2. 如果是合并单元格，且是左上角 -> 正常更新
            ' 3. 如果是合并单元格，但不是左上角 (是附属空单元格) -> 直接跳过！
            ' =========================================================
            If (Not isMerged) Or (isMerged And isTopLeft) Then
                ' 只有有效数据才去骚扰 Word
                UpdateSingleCell wdTbl, arrData(r, c), r, c, curCell
            End If
            
        Next c
    Next r
    On Error GoTo 0
End Sub

' [Logic] 局部数据更新
Private Sub PerformPartialDataUpdate(doc As Object, fullTableRange As Range, overlapRange As Range, tableName As String)
    Dim wdTbl As Object, cell As Range, startR As Long, startC As Long
    Set wdTbl = doc.Bookmarks(tableName).Range.Tables(1)
    startR = fullTableRange.Row: startC = fullTableRange.Column
    
    On Error Resume Next
    For Each cell In overlapRange
        UpdateSingleCell wdTbl, cell.Value, cell.Row - startR + 1, cell.Column - startC + 1, cell
    Next cell
    On Error GoTo 0
End Sub

' [Logic] 单单元格更新 (智能比对 + 差异标记版)
Private Sub UpdateSingleCell(ByVal wdTbl As Object, ByVal newVal As Variant, ByVal r As Long, ByVal c As Long, Optional srcCell As Range = Nothing)
    If isError(newVal) Then Exit Sub
    
    ' 1. 准备 Excel 端的比对字符串
    Dim excelStr As String
    Dim isPercentage As Boolean
    Dim isNumberUpdate As Boolean
    
    If Not srcCell Is Nothing Then
        If InStr(srcCell.NumberFormat, "%") > 0 Then isPercentage = True
        ' 清除之前的颜色标记 (重置状态)
        If srcCell.Interior.Color = 49407 Then srcCell.Interior.ColorIndex = xlNone ' 49407 是橙色
    End If
    
    ' 计算 excelStr (这里我们要同时获取文本和数值的字符串形式)
    If IsEmpty(newVal) Or Trim(CStr(newVal)) = "" Or (IsNumeric(newVal) And VBA.val(newVal) = 0) Then
        excelStr = ""
        isNumberUpdate = True ' 空值视为数值更新的一种（清空）
    ElseIf IsNumeric(newVal) Then
        excelStr = IIf(isPercentage, Format(newVal, "0.00%"), Format(newVal, "#,##0.00"))
        isNumberUpdate = True
    Else
        ' 如果是文本，直接取字符串，但在后续逻辑中会限制更新
        excelStr = Trim(CStr(newVal))
        isNumberUpdate = False
    End If
    
    ' 2. 获取 Word 端的当前值
    Dim wdCell As Object
    On Error Resume Next: Set wdCell = wdTbl.cell(r, c): On Error GoTo 0
    If wdCell Is Nothing Then Exit Sub
    
    Dim wdStr As String
    ' 清洗 Word 字符 (去除回车、响铃符等)
    wdStr = Replace(Replace(wdCell.Range.Text, vbCr, ""), Chr(7), "")
    wdStr = Trim(wdStr)
    
    ' 3. 核心比对与分流处理
    If wdStr <> excelStr Then
        
        ' 场景 A: 这是一个合法的数值更新 (或清空操作)
        ' 且满足二级保护：不能用空值去覆盖 Word 里的纯文本表头
        Dim allowUpdate As Boolean
        allowUpdate = False
        
        If isNumberUpdate Then
            allowUpdate = True
            ' 二级防误删保护: Excel为空, Word为非数字文本 -> 禁止更新
            If excelStr = "" And Len(wdStr) > 0 And Not IsNumeric(wdStr) Then allowUpdate = False
        End If
        
        If allowUpdate Then
            ' >>> 执行更新 <<<
            wdCell.Range.Text = excelStr
            g_cellChangeCount = g_cellChangeCount + 1 ' 计数+1
            
            If g_VisualMode Then
                On Error Resume Next: wdCell.Select: wdTbl.Application.ScreenRefresh: DoEvents: Sleep VISUAL_DELAY: On Error GoTo 0
            End If
            
        Else
            ' 场景 B: 内容有差异，但由于是文本格式或被保护，程序决定不自动更新 Word
            ' 动作: 在 Excel 中标记颜色，提示用户手动检查
            If Not srcCell Is Nothing Then
                ' 只有当 Excel 真的是文本变化时（例如改了表头文字），才标记
                ' 排除掉 "Excel是空但Word是表头" 这种正常的跳过情况
                Dim isValidDiff As Boolean
                isValidDiff = True
                
                ' 如果 Excel 是空值/0，而 Word 是表头，这通常是正常的（合并单元格或无需填写的行），不视为"差异"
                If excelStr = "" And Len(wdStr) > 0 And Not IsNumeric(wdStr) Then isValidDiff = False
                
                If isValidDiff Then
                    srcCell.Interior.Color = 49407 ' 标为橙色 (RGB: 255, 192, 0)
                    g_cellSkipCount = g_cellSkipCount + 1 ' 计数+1
                End If
            End If
        End If
    End If
End Sub

' =================================================================
' [新增] Word 样式工厂：检查并创建样式 (核心)
' =================================================================
Public Sub Ensure_Word_Style(doc As Object, tag As String)
    Dim styleName As String
    Dim prefix As String
    Dim dict As Object
    Dim sty As Object
    
    ' 1. 映射标签到 Word 样式名
    ' prefix 用于在 Config 字典中查找 Key (例如 H1_FontSize)
    Select Case tag
        Case "H":  styleName = "附注-标题":   prefix = "H"
        Case "H0": styleName = "附注-小标题": prefix = "H0"
        Case "H1": styleName = "附注-一级标题": prefix = "H1"
        Case "H2": styleName = "附注-二级标题": prefix = "H2"
        Case "H3": styleName = "附注-三级标题": prefix = "H3"
        Case "P", "B": styleName = "附注-正文": prefix = "P" ' B 共享 P 的基础样式
        Case "D":  styleName = "附注-单位说明": prefix = "D"
        Case Else: Exit Sub
    End Select
    
    ' 2. 获取配置字典
    ' 确保 g_TableConfig 已加载
    If Mod_Financial_Core.g_TableConfig Is Nothing Then
        Call Mod_Financial_Core.Init_Global_Config
    End If
    Set dict = Mod_Financial_Core.g_TableConfig
    
    ' 3. 检查样式是否存在，不存在则创建
    Dim exists As Boolean: exists = False
    On Error Resume Next
    Set sty = doc.Styles(styleName)
    If Err.Number = 0 Then exists = True
    On Error GoTo 0
    
    If Not exists Then
        Set sty = doc.Styles.Add(styleName, 1) ' 1 = wdStyleTypeParagraph
        On Error Resume Next
        sty.BaseStyle = doc.Styles("Normal")
        On Error GoTo 0
    End If
    
    ' 4. 【核心逻辑】强制使用 Excel 配置覆盖样式属性
    ' 无论是否新建，都执行此步骤，确保 Word 样式与 Excel 设置一致
    With sty
        ' A. 字体设置
        If dict.exists(prefix & "_FontName") Then .Font.Name = dict(prefix & "_FontName")
        If dict.exists(prefix & "_FontSize") Then
            Dim fSize As Single: fSize = VBA.val(dict(prefix & "_FontSize"))
            If fSize > 0 Then .Font.Size = fSize
        End If
        If dict.exists(prefix & "_Bold") Then
            .Font.Bold = (VBA.val(dict(prefix & "_Bold")) = 1)
        End If
        
        ' B. 段落格式
        With .ParagraphFormat
            
            ' 大纲级别 (强制逻辑，防止 Excel 配置错误导致结构混乱)
            Select Case tag
                Case "H":  .OutlineLevel = 1 ' Level 1
                Case "H0": .OutlineLevel = 10 ' Body Text
                Case "H1": .OutlineLevel = 1
                Case "H2": .OutlineLevel = 2
                Case "H3": .OutlineLevel = 3
                Case Else: .OutlineLevel = 10 ' Body Text
            End Select
            
            ' 对齐方式 (0=Left, 1=Center, 2=Right, 3=Justify)
            If dict.exists(prefix & "_Alignment") Then
                .Alignment = VBA.val(dict(prefix & "_Alignment"))
            End If
            
            ' 间距 (段前/段后，单位：行)
            ' Word VBA 的 SpaceBefore 默认单位是磅，这里我们需要处理一下
            ' 如果配置的是 0.5 (行)，需要计算。但为简单起见，这里假设配置的是磅值，或者使用 Auto
            ' 更好的做法是使用 LineUnitBefore，但为了兼容性，我们做个简单判断
            If dict.exists(prefix & "_SpaceBefore") Then .SpaceBefore = VBA.val(dict(prefix & "_SpaceBefore")) * 12 ' 假设输入的是行，粗略转为磅
            If dict.exists(prefix & "_SpaceAfter") Then .SpaceAfter = VBA.val(dict(prefix & "_SpaceAfter")) * 12
            
            ' 行距
            If dict.exists(prefix & "_LineSpacing") Then
                Dim lSpace As Single: lSpace = VBA.val(dict(prefix & "_LineSpacing"))
                If lSpace = 1 Then
                    .LineSpacingRule = 0 ' 单倍
                ElseIf lSpace = 1.5 Then
                    .LineSpacingRule = 1 ' 1.5倍
                ElseIf lSpace = 2 Then
                    .LineSpacingRule = 2 ' 双倍
                Else
                    .LineSpacingRule = 5 ' 多倍
                    .LineSpacing = 12 * lSpace
                End If
            End If
            
            ' 缩进 (首行缩进，单位：字符)
            If dict.exists(prefix & "_FirstLineIndent") Then
                Dim indentVal As Single: indentVal = VBA.val(dict(prefix & "_FirstLineIndent"))
                .CharacterUnitFirstLineIndent = indentVal
                .FirstLineIndent = 0 ' 强制清除磅值缩进，优先字符
            End If

        End With
    End With
End Sub

' [重构] Word表格美化 (支持多模板 + 智能对齐)
Private Sub FormatWordTable(tbl As Object)
    On Error Resume Next
    
    Call Mod_Financial_Core.Init_Global_Config
    Dim dict As Object
    Set dict = Mod_Financial_Core.g_TableConfig
    
    Dim prefix As String: prefix = "Table"
    Dim sFont As String, sSize As Single, sHeight As Single, sSpace As Single
    
    ' 默认值
    sFont = "宋体": sSize = 9: sHeight = 22: sSpace = 1
    
    If Not dict Is Nothing Then
        If dict.exists(prefix & "_FontName") Then sFont = dict(prefix & "_FontName")
        If dict.exists(prefix & "_FontSize") Then sSize = VBA.val(dict(prefix & "_FontSize"))
        If dict.exists(prefix & "_RowHeight") Then sHeight = VBA.val(dict(prefix & "_RowHeight"))
        If dict.exists(prefix & "_LineSpacing") Then sSpace = VBA.val(dict(prefix & "_LineSpacing"))
    End If
    
    With tbl
        .AutoFitBehavior 2 ' wdAutoFitWindow
        
        ' 1. 行高
        .Rows.HeightRule = 1
        .Rows.Height = sHeight
        
        ' 2. 基础段落
        .Range.ParagraphFormat.SpaceBefore = 0
        .Range.ParagraphFormat.SpaceAfter = 0
        .Range.ParagraphFormat.LineSpacingRule = 5
        .Range.ParagraphFormat.LineSpacing = 12 * sSpace
        
        .Rows(1).HeadingFormat = True
        .Rows.AllowBreakAcrossPages = False
        
        ' 3. 字体
        .Range.Font.Name = sFont
        .Range.Font.Size = sSize
        
        ' 4. [修改] 对齐方式 (0=不调整)
        If Not dict Is Nothing Then
            If dict.exists(prefix & "_Alignment") Then
                Dim uiAlign As Integer
                uiAlign = VBA.val(dict(prefix & "_Alignment"))
                If uiAlign > 0 Then
                    ' 映射: UI(1)->Word(0), UI(2)->Word(1), ...
                    .Range.ParagraphFormat.Alignment = uiAlign - 1
                    .Range.Cells.VerticalAlignment = 1
                End If
            End If
            
            ' 5. [新增] 表格内缩进
            If dict.exists(prefix & "_FirstLineIndent") Then
                Dim indentVal As Single
                indentVal = VBA.val(dict(prefix & "_FirstLineIndent"))
                If indentVal >= 0 Then
                     .Range.ParagraphFormat.CharacterUnitFirstLineIndent = indentVal
                End If
            End If
        End If
        
        ' 6. 边框 (保持原样)
        .Borders.Enable = True
        .Borders.InsideLineStyle = 1
        .Borders.OutsideLineStyle = 1
        .Borders.InsideLineWidth = 2
        .Borders.OutsideLineWidth = 2
        .Borders.InsideColor = 0
        .Borders.OutsideColor = 0
        .Borders(-2).LineStyle = 0
        .Borders(-4).LineStyle = 0
    End With
    On Error GoTo 0
End Sub


' [Logic] Word表格清洗 (抗 5991 错误增强版 - 双重扫描)
Private Sub CleanWordTableLogic()
    Dim wdApp As Object, Sel As Object, tbl As Object
    Dim delCount As Long
    Dim cell As Object
    Dim isEffective As Boolean
    Dim txt As String, cleanTxt As String
    Dim rangeStart As Long, prevRangeStart As Long
    Dim cycle As Integer ' 新增循环变量

    On Error Resume Next
    Set wdApp = GetObject(, "Kwps.Application")
    If wdApp Is Nothing Then Set wdApp = GetObject(, "Word.Application")
    On Error GoTo 0
    If wdApp Is Nothing Then MsgBox "未检测到 Word。", vbExclamation: Exit Sub
    If wdApp.Documents.count = 0 Then Exit Sub

    Set Sel = wdApp.Selection
    If Sel.Tables.count = 0 Then ForceAppFocus wdApp.ActiveWindow.hwnd: MsgBox "请先点击进入表格内部！", vbExclamation: Exit Sub

    ForceAppFocus wdApp.ActiveWindow.hwnd
    Set tbl = Sel.Tables(1)
    
    wdApp.ScreenUpdating = False
    Application.StatusBar = "正在清洗空行..."
    
    ' =======================================================
    ' 【双重扫描结构】
    ' 循环执行 2 次，确保因光标跳跃而遗漏的空行在第二轮被清除
    ' =======================================================
    For cycle = 1 To 2
        
        ' 每一轮开始前，强制重置光标到表格底部
        tbl.Range.Cells(tbl.Range.Cells.count).Select
        prevRangeStart = -1
    
        Do
            ' 1. 尝试选中当前行
            Sel.SelectRow
            
            ' 防止死循环：如果光标位置不再变化，说明到顶了或卡住了
            rangeStart = Sel.Range.Start
            If rangeStart = prevRangeStart Then Exit Do
            prevRangeStart = rangeStart
            
            If Not Sel.Information(12) Then Exit Do ' 12 = wdWithInTable
            
            isEffective = False
            
            ' =======================================================
            ' 判空逻辑 (直接使用 .Cells 避开 5991)
            ' =======================================================
            If Sel.Cells.count > 0 Then
                For Each cell In Sel.Cells
                    txt = cell.Range.Text
                    ' 清洗：去除 回车(13)、响铃(7)、换行(10)
                    cleanTxt = Replace(txt, Chr(13), "")
                    cleanTxt = Replace(cleanTxt, Chr(7), "")
                    cleanTxt = Replace(cleanTxt, Chr(10), "")
                    cleanTxt = Trim(cleanTxt)
                    
                    ' 只要发现一个格子里有字，就视为有效行
                    If Len(cleanTxt) > 0 Then
                        isEffective = True
                        Exit For
                    End If
                Next cell
            Else
                isEffective = True
            End If
            
            ' =======================================================
            ' 删除时的容错处理
            ' =======================================================
            If Not isEffective Then
                ' 这是一个空行，尝试删除
                On Error Resume Next
                Err.Clear
                Sel.Rows.Delete ' 尝试删除
                
                If Err.Number = 0 Then
                    ' 删除成功
                    ' 仅在第一轮计数，或者累计计数（此处选择累计）
                    delCount = delCount + 1
                    
                    ' 删除后，光标通常会自动移动，但为了保险，稍微往上修正一下位置
                    ' (注意：这里的 MoveUp 是导致可能需要第二轮扫描的原因之一，但在复杂合并表中这是必要的保护措施)
                     wdApp.Selection.MoveUp Unit:=5, count:=1 ' 5=wdLine
                Else
                    ' 报错了 (通常是 5991)，说明这行涉及复杂的合并单元格
                    ' 策略：惹不起，躲得起。不删了，直接往上走。
                    Err.Clear
                    wdApp.Selection.MoveUp Unit:=5, count:=1
                End If
                On Error GoTo 0
            Else
                ' 有效行，向上移动继续检查
                wdApp.Selection.MoveUp Unit:=5, count:=1
            End If
    
            ' 边界检查
            If Not wdApp.Selection.Information(12) Then Exit Do
            ' 到了表格头部停止
            If wdApp.Selection.Range.Start <= tbl.Range.Start Then Exit Do
        Loop
        
    Next cycle
    
    Application.StatusBar = "修复边框..."
    tbl.Select
    With tbl
        ' ==========================================
        ' 5. 边框设置 (0.25磅 + 无左右边框)
        ' ==========================================
        ' A. 先画全套细线 (0.25磅 = 代码2)
        .Borders.Enable = True
        .Borders.InsideLineStyle = 1
        .Borders.OutsideLineStyle = 1
        .Borders.InsideLineWidth = 2 ' 0.25磅
        .Borders.OutsideLineWidth = 2 ' 0.25磅
        .Borders.InsideColor = 0
        .Borders.OutsideColor = 0
        
        ' B. 擦除左右两边的线
        .Borders(-2).LineStyle = 0 ' 左边框无
        .Borders(-4).LineStyle = 0 ' 右边框无
    End With
    
    wdApp.ScreenUpdating = True: Application.StatusBar = False
    tbl.Range.Collapse 0: tbl.Range.Select
    ForceAppFocus wdApp.ActiveWindow.hwnd
    
    If delCount > 0 Then MsgBox "清洗完成！移除 " & delCount & " 个空行，并重置了边框。", vbInformation Else MsgBox "表格格式已修复。", vbInformation
End Sub

' =================================================================
' 【第六部分】 业务逻辑实现 - 校验与辅助 (Processing & Validation)
' =================================================================

' [Check] 自动检查与打标
Public Sub PreCheckAndTagTables(ByVal ws As Worksheet)
    Dim lastRow As Long, i As Long, countUntagged As Long, countRed As Long
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    For i = 1 To lastRow
        If UCase(Trim(ws.Cells(i, 1).Value)) = "TABLE" Then
            If ws.Cells(i, 1).Comment Is Nothing Then countUntagged = countUntagged + 1 Else If ws.Cells(i, 1).Interior.Color = vbRed Then countRed = countRed + 1
        End If
    Next i
    
    If countUntagged = 0 And countRed = 0 Then MsgBox "未发现需要处理的表格标签。", vbInformation: Exit Sub
    If MsgBox("检测到待处理项：" & vbCrLf & "- 未打标: " & countUntagged & vbCrLf & "- 需修复: " & countRed & vbCrLf & "是否执行智能打标？", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    Dim nextIndex As Long: nextIndex = GetNextTableIndex(ws)
    Dim startCell As Range, endCell As Range, idCell As Range
    Dim tableName As String, rStart As Long, rEnd As Long, cEnd As Long, maxDataRow As Long
    Dim countSuccess As Long, countFail As Long
    
    maxDataRow = ws.UsedRange.Rows.count + ws.UsedRange.Row
    
    For i = 1 To lastRow
        Set idCell = ws.Cells(i, 1)
        If UCase(Trim(idCell.Value)) = "TABLE" Then
            If idCell.Comment Is Nothing Or idCell.Interior.Color = vbRed Then
                idCell.ClearComments
                tableName = "Table" & nextIndex
                Set startCell = ws.Cells(i, 2): rStart = startCell.Row
                
                ' 查找结束行 (基于边框)
                rEnd = maxDataRow
                Dim checkRow As Long
                For checkRow = rStart + 1 To maxDataRow
                    If UCase(Trim(ws.Cells(checkRow, 1).Value)) <> "" Then rEnd = checkRow - 1: Exit For
                Next checkRow
                If rEnd < rStart Then rEnd = rStart
                
                Do While rEnd > rStart
                    Dim rngCheck As Range, hasBorder As Boolean
                    Set rngCheck = ws.Range(ws.Cells(rEnd, 2), ws.Cells(rEnd, 11))
                    If IsNull(rngCheck.Borders.LineStyle) Or rngCheck.Borders.LineStyle <> xlNone Then hasBorder = True Else hasBorder = False
                    If hasBorder Then Exit Do Else rEnd = rEnd - 1
                Loop
                
                ' 查找结束列
                cEnd = 2
                Dim lastDataCell As Range
                Set lastDataCell = ws.Range(ws.Cells(rStart, 2), ws.Cells(rEnd, ws.Columns.count)).Find("*", , xlValues, xlPart, xlByColumns, xlPrevious)
                If Not lastDataCell Is Nothing Then cEnd = lastDataCell.Column
                
                If rEnd >= rStart And cEnd >= 2 Then
                    Set endCell = ws.Cells(rEnd, cEnd)
                    If endCell.MergeCells Then Set endCell = endCell.MergeArea.Cells(1, 1)
                    startCell.ClearComments: endCell.ClearComments
                    AddCommentSafe idCell, "TableName:" & tableName
                    AddCommentSafe startCell, "TableStart:" & tableName
                    AddCommentSafe endCell, "TableEnd:" & tableName
                    idCell.Interior.ColorIndex = xlNone
                    countSuccess = countSuccess + 1: nextIndex = nextIndex + 1
                Else
                    countFail = countFail + 1
                End If
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "打标完成！" & vbCrLf & "成功: " & countSuccess & ", 失败: " & countFail, vbInformation
End Sub

' [Check] 智能边框识别 + 智能标题层级识别 (完整保留版)
Private Sub AutoTagByBordersLogic()
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim rngCheck As Range, hasRightBorder As Boolean, isTablePrev As Boolean, tableCount As Long
    Dim regH1 As Object, regH2 As Object, regH3 As Object
    Dim txt As String, tag As String
    Dim isBold As Boolean, fontSize As Single, alignVal As Long
    
    Set ws = ActiveSheet
    ' 【提示保留】
    If MsgBox("将执行【智能边框识别 + 标题分级】。" & vbCrLf & "A列将被覆盖，是否继续？", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Application.ScreenUpdating = False
    
    ' --- 新增功能：初始化正则 ---
    Set regH1 = CreateObject("VBScript.RegExp")
    regH1.Global = False: regH1.Pattern = "^[一二三四五六七八九十]+[、\s]?" ' 匹配 "一、" 或 "一 "
    
    Set regH2 = CreateObject("VBScript.RegExp")
    regH2.Global = False: regH2.Pattern = "^[（\(][一二三四五六七八九十]+[）\)]" ' 匹配 "(一)"
    ' -------------------------
    ' 【新增】 H3 正则：匹配阿拉伯数字加点开头 (如 "1.", "2.", "10.")
    Set regH3 = CreateObject("VBScript.RegExp")
    regH3.Global = False: regH3.Pattern = "^\d+\."
    
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    ws.Range("A2:A" & lastRow).Value = "P"
    ws.Range("A2:A" & lastRow).Interior.ColorIndex = xlNone
    
    isTablePrev = False
    
    For i = 2 To lastRow
        Set rngCheck = ws.Cells(i, 2)
        
        ' =========================================================
        ' 【核心保留 1】合并单元格处理
        ' 作用：如果当前行是合并单元格，则基于整个合并区域判断边框，
        ' 防止因只看左上角而误判边框缺失。
        ' =========================================================
        If rngCheck.MergeCells Then Set rngCheck = rngCheck.MergeArea
        
        ' 【核心保留 2】右边框判断逻辑
        If IsNull(rngCheck.Borders(10).LineStyle) Or rngCheck.Borders(10).LineStyle <> xlNone Then
            hasRightBorder = True
        Else
            hasRightBorder = False
        End If
        
        If hasRightBorder Then
            ' --- 进入表格判断逻辑 ---
            If Not isTablePrev Then
                Dim isSingleCellTable As Boolean: isSingleCellTable = False
                
                ' 【核心保留 3】单列宽表格的特殊排除逻辑
                ' (防止把带边框的标题误判为表格)
                If rngCheck.Columns.count = 1 Then
                    Dim nextRowIdx As Long: nextRowIdx = i + rngCheck.Rows.count
                    Dim nextRowHasBorder As Boolean: nextRowHasBorder = False
                    If nextRowIdx <= lastRow Then
                        Dim nextRng As Range: Set nextRng = ws.Cells(nextRowIdx, 2)
                        If nextRng.MergeCells Then Set nextRng = nextRng.MergeArea
                        If IsNull(nextRng.Borders(10).LineStyle) Or nextRng.Borders(10).LineStyle <> xlNone Then nextRowHasBorder = True
                    End If
                    If nextRowHasBorder = False Then isSingleCellTable = True
                End If
                
                If isSingleCellTable Then
                    isTablePrev = False
                    ' 之前这里直接不管了，现在让它跳转去检查一下是不是标题
                    GoTo CheckTextLogic
                Else
                    ' =========================================================
                    ' 【核心保留 4】确认为表格：标记 TABLE + 标黄
                    ' =========================================================
                    ws.Cells(i, 1).Value = "TABLE"
                    ws.Cells(i, 1).Interior.Color = 65535 ' 标黄 (颜色代码未动)
                    tableCount = tableCount + 1
                    isTablePrev = True
                End If
            Else
                ' 连续表格区域，清除 A 列内容
                ws.Cells(i, 1).ClearContents
                isTablePrev = True
            End If
            
        Else
            ' =========================================================
            ' 非表格区域：此处进行了增强，替换了原有的简单 "P"
            ' =========================================================
            isTablePrev = False
            
CheckTextLogic: ' 标签跳转点
            
            ' 获取单元格特征
            txt = Trim(ws.Cells(i, 2).Value)
            isBold = ws.Cells(i, 2).Font.Bold
            fontSize = ws.Cells(i, 2).Font.Size
            alignVal = ws.Cells(i, 2).HorizontalAlignment
            tag = "P" ' 默认为正文
            
            If Len(txt) > 0 Then
                ' 规则1: H (大标题) - 字号大+居中+加粗
                If fontSize >= 16 And alignVal = -4108 And isBold Then
                    tag = "H"
                ' 规则2: H1 (一级标题) - 正则匹配汉字序号+加粗
                ElseIf isBold And regH1.Test(txt) Then
                    tag = "H1"
                ' 规则3: H2 (二级标题) - 正则匹配括号序号
                ElseIf regH2.Test(txt) Then
                    tag = "H2"
                ' 规则 4: H3 (三级标题) - "1."
                ElseIf regH3.Test(txt) Then
                    tag = "H3"
                End If
            End If
            
            ' 写入识别结果 (覆盖之前的 P)
            ws.Cells(i, 1).Value = tag
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "识别完成！" & vbCrLf & "已标记表格: " & tableCount & " 个" & vbCrLf & "已尝试自动区分 H/H1/H2/P。", vbInformation
End Sub

' [Check] Word/Excel 一致性校验
Private Sub VerifyWordTableConsistencyLogic()
    Dim wdApp As Object, wdDoc As Object, ws As Worksheet
    Dim lastRow As Long, i As Long, cell As Range, tName As String
    Dim countExcelAll As Long, countExcelValid As Long, countWordExist As Long
    Dim missingList As String, isError As Boolean
    
    'On Error GoTo ErrorHandler
    ForceAppFocus Application.hwnd
    Set ws = GetWorksheetSafe(SHEET_TEMPLATE)
    If ws Is Nothing Then Exit Sub
    Call ClearPreviousHighlights(ws)
    Call PreCacheTableRanges(ws)
    
    If Not InitWordApp(wdApp, wdDoc, CreateNew:=False) Then MsgBox "未检测到Word文档。", vbExclamation: Exit Sub
    Application.StatusBar = "正在执行双向核对..."
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    For i = 1 To lastRow
        Set cell = ws.Cells(i, 1)
        If UCase(Trim(cell.Value)) = "TABLE" Then
            countExcelAll = countExcelAll + 1
            isError = False: tName = ""
            
            If Not cell.Comment Is Nothing Then
                If InStr(1, cell.Comment.Text, "TableName:", vbTextCompare) > 0 Then
                    tName = Trim(Split(Split(Replace(cell.Comment.Text, "：", ":"), "TableName:")(1), vbCrLf)(0))
                End If
            End If
            
            If tName = "" Then
                isError = True: tName = "[无名表格-行" & i & "]"
            Else
                If g_tableRanges.exists(tName) Then
                    countExcelValid = countExcelValid + 1
                    If wdDoc.Bookmarks.exists(tName) Then countWordExist = countWordExist + 1 Else isError = True
                Else
                    isError = True: tName = tName & " (定义无效)"
                End If
            End If
            
            If isError Then
                cell.Interior.Color = vbRed
                If Len(missingList) < 500 Then missingList = missingList & tName & vbCrLf
            End If
        End If
    Next i
    
    Application.StatusBar = False
    Dim report As String, iconStyle As VbMsgBoxStyle
    report = "校验完成！" & vbCrLf & "Excel标签: " & countExcelAll & vbCrLf & "有效识别: " & countExcelValid & vbCrLf & "Word书签: " & countWordExist
    If countExcelAll > countWordExist Then
        report = report & vbCrLf & vbCrLf & "【异常】(已标红):" & vbCrLf & missingList
        iconStyle = vbExclamation
    Else
        report = report & vbCrLf & vbCrLf & "数据完美匹配。"
        iconStyle = vbInformation
    End If
    MsgBox report, iconStyle, "报告"
    Exit Sub
ErrorHandler:
    HandleRuntimeError "VerifyConsistency"
End Sub

' [Logic] 表格双向定位
Private Sub LocateTableLogic()
    Dim wdApp As Object, wdDoc As Object, wdTbl As Object, ws As Worksheet
    Dim intersectDict As Object, tName As String, excelRng As Range, isMismatch As Boolean
    
    'On Error GoTo ErrorHandler
    ForceAppFocus Application.hwnd
    Set ws = GetWorksheetSafe(SHEET_TEMPLATE)
    If ws Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then MsgBox "请选择单元格", vbExclamation: Exit Sub
    
    If Not PreCacheTableRanges(ws) Then Exit Sub
    Set intersectDict = GetIntersectDictionary(Selection)
    If intersectDict.count = 0 Then MsgBox "当前选区不在表格范围内。", vbExclamation: Exit Sub
    
    If Not InitWordApp(wdApp, wdDoc, CreateNew:=False) Then Exit Sub
    
    tName = intersectDict.Keys()(0)
    Set excelRng = g_tableRanges(tName)
    
    If Not wdDoc.Bookmarks.exists(tName) Then MsgBox "Word未找到书签: " & tName, vbCritical: Exit Sub
    
    Set wdTbl = wdDoc.Bookmarks(tName).Range.Tables(1)
    If excelRng.Rows.count <> wdTbl.Rows.count Or excelRng.Columns.count <> wdTbl.Columns.count Then isMismatch = True
    
    Dim oldVisual As Boolean: oldVisual = g_VisualMode
    g_VisualMode = True
    FocusAndScrollTo wdDoc, wdDoc.Bookmarks(tName).Range
    g_VisualMode = oldVisual
    
    ForceAppFocus wdApp.ActiveWindow.hwnd
    If isMismatch Then MsgBox "警告：表格结构不一致！" & vbCrLf & "Excel: " & excelRng.Rows.count & "x" & excelRng.Columns.count & vbCrLf & "Word: " & wdTbl.Rows.count & "x" & wdTbl.Columns.count, vbExclamation
    Exit Sub
ErrorHandler:
    HandleRuntimeError "LocateTable"
End Sub

' =================================================================
' 【第七部分】 窗口与应用管理 (Window & App Management)
' 作用：Word初始化、分屏布局、焦点强制控制
' =================================================================

' =============================================================
'  2. 初始化 Word/WPS 函数 (修复了变量名不一致的问题)
' =============================================================
Public Function InitWordApp(app As Object, doc As Object, CreateNew As Boolean) As Boolean
    On Error Resume Next
    
    ' 优先尝试 WPS
    Set app = GetObject(, "Kwps.Application")
    If app Is Nothing Then Set app = CreateObject("Kwps.Application")
    
    ' 失败则尝试 Office Word
    If app Is Nothing Then Set app = GetObject(, "Word.Application")
    If app Is Nothing Then Set app = CreateObject("Word.Application")
    On Error GoTo 0
    
    If app Is Nothing Then MsgBox "无法启动 WPS 或 Word。", vbCritical: Exit Function
    
    app.Visible = True
    
    ' --- [核心修复点] 挂接 Word 事件监听 ---
    If g_WordAppHandler Is Nothing Then
        Set g_WordAppHandler = New WordEventHandler
    End If
    
    ' [关键] 这里的 .appInstance 必须与类模块中的变量名完全一致！
    ' 即使 app 是 WPS 对象，VBA 也能通过接口兼容性将其赋值给 Word.Application 变量
    Set g_WordAppHandler = New WordEventHandler
    Set g_WordAppHandler.appInstance = app
    ' ------------------------------------
    
    ' 获取文档对象 (保持原有逻辑)
    If CreateNew Then
        Set doc = app.Documents.Add
    Else
        If app.Documents.count > 0 Then
            Set doc = app.ActiveDocument
        Else
            Dim f As Variant
            f = Application.GetOpenFilename("Word文件 (*.docx; *.doc),*.docx;*.doc")
            If f <> False Then Set doc = app.Documents.Open(f)
        End If
    End If
    
    InitWordApp = Not (doc Is Nothing)
End Function

' 分屏布局逻辑
Private Sub ArrangeWindowsSplitScreen(wdApp As Object)
    On Error Resume Next
    Dim needWait As Boolean
    needWait = (Application.WindowState <> xlMaximized)
    
    Application.WindowState = xlMaximized
    
    Dim k As Long
    If needWait Then
        For k = 1 To 3000: DoEvents: Next k
    End If
    
    Dim baseW As Double, baseH As Double
    baseW = Application.UsableWidth: baseH = Application.UsableHeight
    
    Dim targetW As Double, targetH As Double, targetLeft As Double, targetTop As Double
    targetW = baseW * 0.35: targetH = baseH * 0.9
    targetLeft = baseW - targetW - 150: targetTop = (baseH - targetH) / 2
    
    With wdApp
        .Visible = True
        If .WindowState <> 2 Then .WindowState = 2: DoEvents
        .WindowState = 0: .Activate
        
        Dim i As Long
        For i = 1 To 2000: DoEvents: Next i
        
        .Left = targetLeft: .Top = targetTop: .Width = targetW: .Height = targetH
        If Abs(.Left - targetLeft) > 10 Then .Left = targetLeft: .Width = targetW
        If .Windows.count > 0 Then .ActiveWindow.WindowState = 1
        .Activate
    End With
    On Error GoTo 0
End Sub

' 强制置顶窗口
#If VBA7 Then
    Private Sub ForceAppFocus(ByVal hwnd As LongPtr)
#Else
    Private Sub ForceAppFocus(ByVal hwnd As Long)
#End If
    
    On Error Resume Next
    If hwnd <> 0 Then
        If IsIconic(hwnd) <> 0 Then
            ShowWindow hwnd, 9
        Else
            ShowWindow hwnd, 5
        End If
        
        SetForegroundWindow hwnd
    End If
    On Error GoTo 0
End Sub

' 镜头跟随
Private Sub FocusAndScrollTo(wdDoc As Object, rng As Object)
    If Not g_VisualMode Or rng Is Nothing Or wdDoc Is Nothing Then Exit Sub
    On Error Resume Next
    ForceAppFocus wdDoc.Application.ActiveWindow.hwnd
    rng.Select
    wdDoc.Application.ScreenRefresh
    Sleep VISUAL_DELAY
    On Error GoTo 0
End Sub

' 重置窗口
Private Sub ResetWindowLayoutLogic()
    On Error Resume Next
    Application.ScreenUpdating = True: Application.WindowState = xlMaximized
    Dim wdApp As Object
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = GetObject(, "Kwps.Application")
    If Not wdApp Is Nothing Then
        ArrangeWindowsSplitScreen wdApp
        wdApp.Visible = True: wdApp.Activate
    Else
        MsgBox "未检测到运行中的 Word。", vbExclamation
    End If
    On Error GoTo 0
End Sub

' =================================================================
' 【第八部分】 底层工具与缓存 (Helpers & Cache)
' 作用：表格范围计算、字典管理、错误处理、清理垃圾
' =================================================================

' 缓存表格位置到全局变量
Public Function PreCacheTableRanges(ws As Worksheet) As Boolean
    Set g_tableRanges = CreateObject("Scripting.Dictionary"): g_tableRanges.CompareMode = 1
    Set g_tableCellMap = CreateObject("Scripting.Dictionary"): g_tableCellMap.CompareMode = 1
    
    Dim cmt As Comment, txt As String, tName As String
    Dim startD As Object: Set startD = CreateObject("Scripting.Dictionary")
    Dim endD As Object: Set endD = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    For Each cmt In ws.Comments
        txt = Replace(cmt.Text, "：", ":")
        If InStr(1, txt, "TableStart:", vbTextCompare) > 0 Then
            Set startD(Trim(Split(txt, ":")(1))) = cmt.Parent
        ElseIf InStr(1, txt, "TableEnd:", vbTextCompare) > 0 Then
            Set endD(Trim(Split(txt, ":")(1))) = cmt.Parent
        ElseIf InStr(1, txt, "TableName:", vbTextCompare) > 0 Then
            tName = Trim(Split(Split(txt, "TableName:")(1), vbCrLf)(0))
            If Not g_tableCellMap.exists(tName) Then g_tableCellMap.Add tName, cmt.Parent
        End If
    Next cmt
    On Error GoTo 0
    
    Dim k As Variant, r1 As Range, r2 As Range
    For Each k In startD.Keys
        If endD.exists(k) Then
            Set r1 = startD(k): Set r2 = endD(k)
            If r2.Row > r1.Row And WorksheetFunction.CountA(r2.EntireRow) = 0 Then Set r2 = ws.Cells(r2.Row - 1, r2.Column)
            If r2.Row >= r1.Row Then g_tableRanges.Add k, ws.Range(r1, r2)
        End If
    Next k
    
    PreCacheTableRanges = (g_tableRanges.count > 0)
    If Not PreCacheTableRanges Then MsgBox "未找到有效表格区域 (TableStart/TableEnd)", vbExclamation
End Function

' 其他小工具
' [核心逻辑] 序号智能修正 (最终完整版：一级至五级 + 性能预警)
Private Sub AutoFixNumberingLogic()
    Dim wdApp As Object, doc As Object
    Dim para As Object
    Dim txt As String
    
    Dim regL1 As Object, regL2 As Object, regL3 As Object, regL4 As Object, regL5 As Object
    Dim matches As Object
    
    Dim lastL1 As Long, lastL2 As Long, lastL3 As Long, lastL4 As Long, lastL5 As Long
    Dim curL1 As Long, curL2 As Long, curL3 As Long, curL4 As Long, curL5 As Long
    
    Dim fixCount As Long
    Dim fixRng As Object
    
    If Not InitWordApp(wdApp, doc, CreateNew:=False) Then Exit Sub
    
    Dim response As VbMsgBoxResult
    response = MsgBox("即将开始全文档序号扫描与修正。" & vbCrLf & vbCrLf & _
                      "检查范围：一级(一、) 至 五级 1) " & vbCrLf & _
                      "当前操作会开启【修订模式】，对于长文档可能需要几十秒至数分钟。" & vbCrLf & vbCrLf & _
                      "建议先保存文档。是否继续？", vbYesNo + vbExclamation, "性能预警")
    If response = vbNo Then Exit Sub
    
    wdApp.ScreenUpdating = False
    wdApp.StatusBar = "正在初始化..."
    
    If doc.TrackRevisions = False Then doc.TrackRevisions = True
    
    Set regL1 = CreateObject("VBScript.RegExp"): regL1.Pattern = "^([一二三四五六七八九十]+)、"
    Set regL2 = CreateObject("VBScript.RegExp"): regL2.Pattern = "^[（\(]([一二三四五六七八九十]+)[）\)]"
    Set regL3 = CreateObject("VBScript.RegExp"): regL3.Pattern = "^(\d+)\."
    Set regL4 = CreateObject("VBScript.RegExp"): regL4.Pattern = "^[（\(](\d+)[）\)]"
    Set regL5 = CreateObject("VBScript.RegExp"): regL5.Pattern = "^(\d+)\)"
    
    lastL1 = 0: lastL2 = 0: lastL3 = 0: lastL4 = 0: lastL5 = 0
    fixCount = 0
    
    Dim totalParas As Long, pIndex As Long
    totalParas = doc.Paragraphs.count
    pIndex = 0
    
    For Each para In doc.Paragraphs
        pIndex = pIndex + 1
        If pIndex Mod 50 = 0 Then
            wdApp.StatusBar = "正在扫描段落: " & pIndex & " / " & totalParas & " (已修复: " & fixCount & ")"
            DoEvents
        End If
    
        If para.Range.Information(12) = True Then GoTo NextPara
        
        txt = Trim(para.Range.Text)
        txt = Replace(txt, vbCr, ""): txt = Replace(txt, vbLf, "")
        txt = Replace(txt, Chr(11), "")
        If Len(txt) = 0 Then GoTo NextPara
        
        If regL1.Test(txt) Then
            Set matches = regL1.Execute(txt)
            curL1 = ChineseToNumber(matches(0).SubMatches(0))
            
            If curL1 <> 0 Then
                If curL1 <> 1 And curL1 <> lastL1 + 1 Then
                    Call ExecuteFix(doc, para, matches(0), NumberToChinese(lastL1 + 1) & "、")
                    fixCount = fixCount + 1
                    lastL1 = lastL1 + 1
                Else
                    lastL1 = curL1
                    If curL1 = 1 Then lastL1 = 1
                End If
                lastL2 = 0: lastL3 = 0: lastL4 = 0: lastL5 = 0
            End If
            GoTo NextPara
        End If
        
        If regL2.Test(txt) Then
            Set matches = regL2.Execute(txt)
            curL2 = ChineseToNumber(matches(0).SubMatches(0))
            
            If curL2 <> 0 Then
                If curL2 <> 1 And curL2 <> lastL2 + 1 Then
                    Call ExecuteFix(doc, para, matches(0), "(" & NumberToChinese(lastL2 + 1) & ")")
                    fixCount = fixCount + 1
                    lastL2 = lastL2 + 1
                Else
                    lastL2 = curL2
                    If curL2 = 1 Then lastL2 = 1
                End If
                lastL3 = 0: lastL4 = 0: lastL5 = 0
            End If
            GoTo NextPara
        End If
        
        If regL3.Test(txt) Then
            Set matches = regL3.Execute(txt)
            curL3 = CLng(matches(0).SubMatches(0))
            
            If curL3 <> 1 And curL3 <> lastL3 + 1 Then
                Call ExecuteFix(doc, para, matches(0), (lastL3 + 1) & ".")
                fixCount = fixCount + 1
                lastL3 = lastL3 + 1
            Else
                lastL3 = curL3
                If curL3 = 1 Then lastL3 = 1
            End If
            lastL4 = 0: lastL5 = 0
            GoTo NextPara
        End If
        
        If regL4.Test(txt) Then
            Set matches = regL4.Execute(txt)
            curL4 = CLng(matches(0).SubMatches(0))
            
            If curL4 <> 1 And curL4 <> lastL4 + 1 Then
                Call ExecuteFix(doc, para, matches(0), "(" & (lastL4 + 1) & ")")
                fixCount = fixCount + 1
                lastL4 = lastL4 + 1
            Else
                lastL4 = curL4
                If curL4 = 1 Then lastL4 = 1
            End If
            lastL5 = 0
            GoTo NextPara
        End If
        
        If regL5.Test(txt) Then
            Set matches = regL5.Execute(txt)
            curL5 = CLng(matches(0).SubMatches(0))
            
            If curL5 <> 1 And curL5 <> lastL5 + 1 Then
                Call ExecuteFix(doc, para, matches(0), (lastL5 + 1) & ")")
                fixCount = fixCount + 1
                lastL5 = lastL5 + 1
            Else
                lastL5 = curL5
                If curL5 = 1 Then lastL5 = 1
            End If
            GoTo NextPara
        End If
        
NextPara:
    Next para
    
    ' 恢复环境
    wdApp.ScreenUpdating = True
    wdApp.StatusBar = False
    
    If fixCount = 0 Then
        MsgBox "检查完成：全文序号连续，无需修正。", vbInformation
    Else
        MsgBox "修正完成！共自动修复 " & fixCount & " 处断号。" & vbCrLf & vbCrLf & _
               "★ 已启用【修订模式】并【标黄】修改处，请在 Word 中复核。", vbInformation
        wdApp.Activate
    End If
End Sub

' [逻辑] 执行具体的替换操作 (修订+标黄)
Private Sub ExecuteFix(doc As Object, para As Object, match As Object, newStr As String)
    Dim rng As Object
    Set rng = doc.Range(para.Range.Start + match.FirstIndex, para.Range.Start + match.FirstIndex + match.Length)
    
    rng.Text = newStr
    
    rng.HighlightColorIndex = 7 ' wdYellow
End Sub

' [辅助] 中文数字转阿拉伯 (支持1-99)
Private Function ChineseToNumber(s As String) As Long
    Dim nums As String: nums = "一二三四五六七八九十"
    If Len(s) = 1 Then
        ChineseToNumber = InStr(nums, s)
    ElseIf InStr(s, "十") > 0 Then
        If Left(s, 1) = "十" Then
            ChineseToNumber = 10 + InStr(nums, Right(s, 1))
        ElseIf Right(s, 1) = "十" Then
            ChineseToNumber = InStr(nums, Left(s, 1)) * 10
        Else
            ChineseToNumber = InStr(nums, Left(s, 1)) * 10 + InStr(nums, Right(s, 1))
        End If
    Else
        ChineseToNumber = 0
    End If
End Function

' [辅助] 阿拉伯转中文 (支持1-99)
Private Function NumberToChinese(n As Long) As String
    Dim nums As String: nums = "一二三四五六七八九十"
    If n <= 10 Then
        NumberToChinese = Mid(nums, n, 1)
    ElseIf n < 20 Then
        NumberToChinese = "十" & Mid(nums, n - 10, 1)
    ElseIf n < 100 Then
        NumberToChinese = Mid(nums, n \ 10, 1) & "十"
        If n Mod 10 <> 0 Then NumberToChinese = NumberToChinese & Mid(nums, n Mod 10, 1)
    Else
        NumberToChinese = CStr(n)
    End If
End Function

Private Function GetNextTableIndex(ws As Worksheet) As Long
    Dim maxIdx As Long, cmt As Comment, regex As Object, matches As Object
    On Error Resume Next
    Set regex = CreateObject("VBScript.RegExp"): regex.Global = True: regex.Pattern = "Table\s*(\d+)"
    For Each cmt In ws.Comments
        If regex.Test(cmt.Text) Then
            Dim num As Long: num = CLng(regex.Execute(cmt.Text)(0).SubMatches(0))
            If num > maxIdx Then maxIdx = num
        End If
    Next cmt
    GetNextTableIndex = maxIdx + 1
End Function

Private Sub AddCommentSafe(rng As Range, txt As String)
    On Error Resume Next
    If rng.Comment Is Nothing Then rng.AddComment
    Dim curTxt As String: curTxt = rng.Comment.Text
    If InStr(curTxt, txt) = 0 Then rng.Comment.Text Text:=IIf(curTxt = "", "", curTxt & vbCrLf) & txt
    rng.Comment.Shape.TextFrame.AutoSize = True: rng.Comment.Visible = False
    On Error GoTo 0
End Sub

Private Function GetUserSelection(AllowMultiSelect As Boolean) As Boolean
    Set g_tablesToUpdate = CreateObject("Scripting.Dictionary")
    g_UserCancelled = False
    On Error Resume Next
    Load UserForm1
    If Err.Number <> 0 Then MsgBox "缺失 UserForm1 窗体", vbCritical: Exit Function
    On Error GoTo 0
    With UserForm1
        .ListBox1.MultiSelect = IIf(AllowMultiSelect, 1, 0)
        .Show
    End With
    If Not g_UserCancelled And g_tablesToUpdate.count > 0 Then GetUserSelection = True
End Function

Public Function GetWorksheetSafe(n As String) As Worksheet
    On Error Resume Next: Set GetWorksheetSafe = ActiveWorkbook.Sheets(n): On Error GoTo 0
    If GetWorksheetSafe Is Nothing Then MsgBox "找不到工作表: " & n, vbCritical
End Function

Public Function GetIntersectDictionary(rng As Range) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    If g_tableRanges Is Nothing Then PreCacheTableRanges rng.Worksheet
    For Each k In g_tableRanges.Keys
        If Not Intersect(rng, g_tableRanges(k)) Is Nothing Then d.Add k, Intersect(rng, g_tableRanges(k))
    Next k
    Set GetIntersectDictionary = d
End Function

Private Sub ToggleSystemUpdates(state As Boolean, Optional wdApp As Object = Nothing)
    Application.ScreenUpdating = state: Application.EnableEvents = state
    If state Then Application.StatusBar = False
    If Not wdApp Is Nothing Then
        If g_VisualMode Then wdApp.ScreenUpdating = True Else wdApp.ScreenUpdating = state
    End If
End Sub

' [工具] 询问是否开启可视化模式 (含焦点警告)
Private Sub AskForVisualMode()
    Dim msg As String
    
    msg = "是否开启【可视化演示模式】？" & vbCrLf & _
          "(开启后，您可以直观看到 Word 的自动化操作过程)" & vbCrLf & vbCrLf & _
          "【警告：请勿动鼠标】" & vbCrLf & _
          "程序运行期间，系统将接管鼠标和键盘焦点。" & vbCrLf & _
          "请切勿随意【切换窗口】或【点击鼠标】，否则会导致程序失去控制权从而报错！" & vbCrLf & vbCrLf & _
          "建议：如果不确定，请选择“否”使用后台静默模式（更稳定）。"
    
    ' vbDefaultButton2 表示默认选中“否”，引导用户倾向于更稳定的后台模式
    g_VisualMode = (MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2, "模式选择与安全警告") = vbYes)
End Sub

Private Function ValidateTableExists(doc As Object, n As String, Optional Silent As Boolean = False) As Boolean
    ValidateTableExists = doc.Bookmarks.exists(n)
    If Not ValidateTableExists And Not Silent Then MsgBox "Word中找不到书签: " & n, vbExclamation
End Function

Private Function ValidateTableDimensions(doc As Object, rng As Range, n As String) As Boolean
    Dim tbl As Object: Set tbl = doc.Bookmarks(n).Range.Tables(1)
    ValidateTableDimensions = (rng.Rows.count = tbl.Rows.count And rng.Columns.count = tbl.Columns.count)
    If Not ValidateTableDimensions Then MsgBox "表格维度不匹配: " & n, vbExclamation
End Function

' [工具] 安全复制函数 (增强稳定性版)
' =================================================================
' 函数: CopyRangeSafe
' 功能: 增强型复制，包含清空剪贴板、重试机制、成功验证
' =================================================================
Private Function CopyRangeSafe(rng As Range) As Boolean
    Dim i As Integer
    Dim success As Boolean
    success = False
    
    ' 1. 尝试清空之前的剪贴板状态 (防止粘贴了上一次的内容)
    On Error Resume Next
    Application.CutCopyMode = False
    Err.Clear
    On Error GoTo 0
    
    ' 2. 循环尝试复制 (最多 10 次)
    For i = 1 To 10
        On Error Resume Next
        rng.Copy ' 执行复制
        
        ' 检查点：如果没有报错，且 Excel 处于“复制模式”(有蚂蚁线)，说明成功进入内存
        If Err.Number = 0 And Application.CutCopyMode <> 0 Then
            success = True
            Exit For
        End If
        
        ' 失败处理
        Err.Clear
        On Error GoTo 0
        
        ' 关键：交出 CPU 控制权，等待系统缓存 I/O
        DoEvents
        Sleep 200 ' 等待 200毫秒 (指数退避策略可进一步优化，但线性够用了)
    Next i
    
    ' 3. 最后的挽救：大延迟重试
    If Not success Then
        DoEvents: Sleep 500
        On Error Resume Next
        rng.Copy
        If Application.CutCopyMode <> 0 Then success = True
        On Error GoTo 0
    End If
    
    CopyRangeSafe = success
End Function

' =================================================================
' 函数: PasteTableSafe
' 功能: 验证型粘贴，确保表格成功进入 Word，否则报错并重试
' 参数:
'   wdRange: 粘贴的目标位置
'   PasteDataType: 粘贴类型 (默认 wdPasteRTF 或 wdPasteDefault)
' =================================================================
Private Function PasteTableSafe(wdRange As Object) As Boolean
    Dim i As Integer
    Dim tableCountBefore As Long
    Dim tableCountAfter As Long
    Dim pasteSuccess As Boolean
    
    pasteSuccess = False
    
    ' 获取粘贴前的表格数量（作为基准线）
    ' 注意：这里统计的是整个文档或当前 Range 所在容器的表格数
    ' 建议检查 wdRange.Document.Tables.Count，因为 Range 范围粘贴后会变
    tableCountBefore = wdRange.Document.Tables.count
    
    ' 循环重试机制
    For i = 1 To 5
        On Error Resume Next
        Err.Clear
        
        ' 执行粘贴
        ' 建议使用 PasteExcelTable 保持格式，或者 Paste
        ' False, False, True 分别代表: Link to Excel, Word Formatting, RTF
        wdRange.PasteExcelTable False, False, True
        ' 如果上方报错，尝试通用粘贴: wdRange.Paste
        
        If Err.Number = 0 Then
            ' 仅仅没有报错是不够的，我们需要验证
            DoEvents
            Sleep 200 ' 等待 Word 处理 DOM
            
            tableCountAfter = wdRange.Document.Tables.count
            
            ' 验证：表格数量增加了，说明粘贴进去了
            If tableCountAfter > tableCountBefore Then
                pasteSuccess = True
                Exit For
            Else
                ' 虽然没报错，但表格没增加（可能粘贴成了图片或文本，或者完全没反应）
                ' 此时不需要 Exit For，继续重试
                Debug.Print "粘贴尝试 " & i & " 失败：对象数量未增加"
            End If
        Else
            Debug.Print "粘贴尝试 " & i & " 报错: " & Err.Description
        End If
        On Error GoTo 0
        
        ' 失败等待
        DoEvents
        Sleep 500 * i ' 递增等待：0.5s, 1.0s, 1.5s...
    Next i
    
    PasteTableSafe = pasteSuccess
End Function

Private Sub ClearPreviousHighlights(ws As Worksheet)
    On Error Resume Next
    Dim cell As Range
    For Each cell In ws.Columns("A").SpecialCells(xlCellTypeConstants)
        If UCase(Trim(cell.Value)) = "TABLE" Then cell.Interior.ColorIndex = xlNone
    Next cell
    On Error GoTo 0
End Sub

Private Sub HandleRuntimeError(src As String)
    MsgBox "运行错误 [" & src & "]: " & vbCrLf & Err.Description, vbCritical
    Cleanup
End Sub

Private Sub Cleanup(Optional wdApp As Object = Nothing)
    On Error Resume Next
    Application.ScreenUpdating = True: Application.StatusBar = False
    Application.EnableEvents = True: Application.CutCopyMode = False
    If Not wdApp Is Nothing Then wdApp.ScreenUpdating = True
    Set g_tableRanges = Nothing: Set g_tablesToUpdate = Nothing
    On Error GoTo 0
End Sub
' [新增] 辅助：加载表格样式配置到内存
Private Sub LoadTableConfigFromSheet()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim k As String, v As String
    
    Set g_TableConfig = CreateObject("Scripting.Dictionary")
    
    ' 默认兜底值 (防止读取失败)
    g_TableConfig("FontName") = "宋体"
    g_TableConfig("FontSize") = 9
    g_TableConfig("RowHeight") = 22
    g_TableConfig("LineSpacing") = 1
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("User_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ' 读取 E/F 列
    For i = 2 To lastRow
        k = Trim(CStr(ws.Cells(i, "E").Value))
        v = Trim(CStr(ws.Cells(i, "F").Value))
        
        If k <> "" Then
            ' 如果已经存在则更新，不存在则添加
            If g_TableConfig.exists(k) Then
                g_TableConfig(k) = v
            Else
                g_TableConfig.Add k, v
            End If
        End If
    Next i
End Sub

' =================================================================
' 函数: AskUserModeless
' 功能: 弹出非模态确认框，允许用户操作Excel，同时等待用户点击按钮
' =================================================================
Private Function AskUserModeless(msg As String, title As String) As VbMsgBoxResult
    Dim frm As New FrmConfirm
    
    ' 1. 初始化状态
    g_ConfirmResult = 0 ' 0 表示未选择
    
    ' 2. 设置窗体内容
    frm.Caption = title
    frm.lblMessage.Caption = msg
    
    ' 3. 显示窗体 (非模态)
    ' 关键：用户此时可以去滚动 Excel 界面
    frm.Show vbModeless
    
    ' 4. 进入等待循环
    ' 只要 g_ConfirmResult 还是 0，就一直循环
    Do While g_ConfirmResult = 0
        DoEvents ' 让出控制权给操作系统，允许用户点击 Excel 单元格
        Sleep 50 ' 休息 50毫秒，防止 CPU 占用率 100%
        
        ' 检查窗体是否还在（防止意外关闭报错）
        If Not IsFormLoaded("FrmConfirm") Then
            g_ConfirmResult = vbCancel
            Exit Do
        End If
    Loop
    
    ' 5. 用户点击了按钮，循环结束，卸载窗体
    Unload frm
    Set frm = Nothing
    
    ' 6. 返回结果
    AskUserModeless = g_ConfirmResult
End Function

' 辅助函数：检查窗体是否加载
Private Function IsFormLoaded(formName As String) As Boolean
    Dim frm As Object
    IsFormLoaded = False
    For Each frm In UserForms
        If StrComp(frm.Name, formName, vbTextCompare) = 0 Then
            IsFormLoaded = True
            Exit For
        End If
    Next frm
End Function
' =================================================================
' [新增工具] 在状态栏绘制视觉进度条
' 效果：[■■■■■□□□□□] 50% (正在处理 50/100)
' =================================================================
Private Sub UpdateVisualStatusBar(cur As Long, total As Long, pct As Double)
    Dim barLen As Integer
    Dim filled As Integer
    Dim barStr As String
    
    barLen = 20 ' 进度条总长度（20个方块）
    filled = Int(pct * barLen)
    
    ' 生成方块字符串
    ' String函数生成指定数量的字符：ChrW(9632)是实心方块，ChrW(9633)是空心方块
    barStr = String(filled, ChrW(9632)) & String(barLen - filled, ChrW(9633))
    
    ' 更新状态栏
    Application.StatusBar = "进度: [" & barStr & "] " & Format(pct, "0%") & _
                            " (行: " & cur & "/" & total & ")"
End Sub


' =============================================================
'  1. 监听器开关 (供 ThisWorkbook 调用)
' =============================================================
Public Sub StartAppEventListeners()
    On Error GoTo ErrHandle
    
    ' 如果已经存在，先不做处理（或者先销毁再重建）
    If Not g_ExcelAppHandler Is Nothing Then Exit Sub

    ' 尝试实例化
    Set g_ExcelAppHandler = New ExcelAppEventHandler
    Set g_ExcelAppHandler.xlApp = Application
    
    Exit Sub

ErrHandle:
    ' 【关键】一旦出错，必须把对象清空，否则外层判断会失效
    Set g_ExcelAppHandler = Nothing
    ' 可以选择在这里 debug.print，或者在外层弹窗
End Sub

Public Sub StopAppEventListeners()
    ' 停止所有监听
    Set g_ExcelAppHandler = Nothing
    Set g_WordAppHandler = Nothing
End Sub

' 强力激活窗口函数
Public Sub ForceActivateWindow(appObj As Object)
    On Error Resume Next
    Dim hwnd As Long ' 如果是32位Office，改成 Long
    
    ' 尝试获取窗口句柄 (WPS和Word通常都有 Hwnd 属性，如果没有则忽略)
    ' WPS 有时 Hwnd 属性隐藏较深，我们尝试直接 Activate
    appObj.Visible = True
    
    ' 1. 标准激活
    appObj.Activate
    
    ' 2. 尝试窗口还原 (防止最小化)
    #If VBA7 Then
        Dim appHwnd As LongPtr
    #Else
        Dim appHwnd As Long
    #End If
    
    ' 尝试读取句柄，如果对象不支持可能会报错，忽略即可
    appHwnd = appObj.ActiveWindow.hwnd
    
    If appHwnd <> 0 Then
        ShowWindow appHwnd, SW_RESTORE
        SetForegroundWindow appHwnd
        BringWindowToTop appHwnd
    End If
    On Error GoTo 0
End Sub

' =================================================================
' 4. 功能开关核心逻辑 (业务层)
' =================================================================
Sub Toggle_Link_Feature_Action()
    Dim currentStatus As Boolean
    
    ' 读取当前状态
    currentStatus = Mod_Financial_Core.Get_Sys_Switch_Status("Link_Word_Excel")
    
    If currentStatus Then
        ' --- 关闭 ---
        Call StopAppEventListeners
        Call Mod_Financial_Core.Set_Sys_Switch_Status("Link_Word_Excel", False)
        MsgBox "双击互转功能已【关闭】。", vbInformation
    Else
        ' --- 开启 ---
        ' 【修改点】不再调用 EnsureWordReference，直接尝试启动
        ' 如果引用丢失，StartAppEventListeners 内部会捕获错误或实例化失败
        Call StartAppEventListeners
        
        If Not g_ExcelAppHandler Is Nothing Then
            Call Mod_Financial_Core.Set_Sys_Switch_Status("Link_Word_Excel", True)
            MsgBox "双击互转功能已【开启】！" & vbCrLf & "现在支持 Excel 和 Word/WPS 表格双击互跳。", vbInformation
        Else
            ' 如果启动失败，说明引用可能有问题，或者Word没启动
            MsgBox "功能启动失败！" & vbCrLf & vbCrLf & _
                   "可能原因：" & vbCrLf & _
                   "1. Word/WPS 未运行。" & vbCrLf & _
                   "2. 程序版本兼容性问题 (引用丢失)。", vbExclamation
            
            ' 确保状态为关闭
            Call Mod_Financial_Core.Set_Sys_Switch_Status("Link_Word_Excel", False)
        End If
    End If
End Sub

' 对应 XML 中的 onLoad="Ribbon_OnLoad"
Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set g_Ribbon = ribbon
    ' 可以在这里强制刷新一下状态
    ribbon.InvalidateControl "tglLinkFeature"
End Sub
' 对应 XML 中的 getPressed="Rx_GetLinkPressed"
' 参数 returnedVal: 返回 True 表示按钮按下(开启状态)，False 表示弹起(关闭状态)
Public Sub Rx_GetLinkPressed(control As IRibbonControl, ByRef returnedVal As Boolean)
    ' 从配置表中读取状态
    returnedVal = Mod_Financial_Core.Get_Sys_Switch_Status("Link_Word_Excel")
End Sub
' 对应 XML 中的 onAction="Rx_ToggleLinkFeature"
' 参数 pressed: 用户点击后的新状态 (True/False)
Public Sub Rx_ToggleLinkFeature(control As IRibbonControl, pressed As Boolean)
    If pressed Then
        ' --- 开启 ---
        ' 【修改点】直接启动，移除修复逻辑
        Call StartAppEventListeners
        
        If Not g_ExcelAppHandler Is Nothing Then
            MsgBox "双击互转功能已【开启】！", vbInformation
            Call Mod_Financial_Core.Set_Sys_Switch_Status("Link_Word_Excel", True)
        Else
            MsgBox "启动失败：未检测到 Word 或环境不支持。", vbExclamation
            ' 失败回滚：更新状态并强制刷新UI
            Call Mod_Financial_Core.Set_Sys_Switch_Status("Link_Word_Excel", False)
            ' 【安全修正】检查 g_Ribbon 是否还活着，防止报错
            If Not g_Ribbon Is Nothing Then
                g_Ribbon.InvalidateControl "tglLinkFeature"
            End If
        End If
        
    Else
        ' --- 关闭 ---
        Call StopAppEventListeners
        Call Mod_Financial_Core.Set_Sys_Switch_Status("Link_Word_Excel", False)
        MsgBox "双击互转功能已【关闭】。", vbInformation
    End If
End Sub
' =================================================================
' 8. [新增] 智能查找目标工作簿
' 作用：遍历所有打开的工作簿，找到包含"附注模板"的那个文件
'       排除掉插件本身(ThisWorkbook)，防止跳错到代码文件
' =================================================================
Public Function FindTargetWorkbook() As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' 遍历当前 Excel 实例中所有打开的文件
    For Each wb In Application.Workbooks
        ' 1. 排除插件本身 (.xlam 或 .xlsm)
        ' 注意：这里判断名字是否与当前代码所在的工作簿相同
        If UCase(wb.Name) <> UCase(ThisWorkbook.Name) Then
            
            ' 2. 尝试在这个 wb 里找 "附注模板"
            On Error Resume Next
            Set ws = Nothing
            Set ws = wb.Sheets("附注模板")
            On Error GoTo 0
            
            ' 3. 如果找到了，说明这就是我们要的目标数据文件
            If Not ws Is Nothing Then
                Set FindTargetWorkbook = wb
                Exit Function
            End If
        End If
    Next wb
    
    ' 如果循环完了都没找到
    Set FindTargetWorkbook = Nothing
End Function
