Attribute VB_Name = "Mod_Financial_Core"
' =================================================================
' Copyright (C) 2025 Dota (PCCPA). All Rights Reserved.
' 本程序受著作权法和国际条约保护。
' 未经授权的复制或分发本程序（或其中任何部分），将导致严厉的民事和刑事处罚。
' =================================================================
' =================================================================
' 模块名称: 财务附注ETL工具
' 作者：Dota | 天健会计师事务所 (PCCPA)
' 更新日志:
'   1204: 增加单体和汇总注释科目的功能1.0版本
' =================================================================
Option Explicit

' ==========================================================================================
' MODULE:       Mod_Financial_Core
' DESCRIPTION:  财务报表自动化核心系统 (含复杂表检测与拆分)
' VERSION:      3.4 (Manual Split Enhanced)
' AUTHOR:       Senior VBA Architect
' DATE:         2025-12-03
' UPDATE:       升级手动分段功能，复用复杂表拆分核心逻辑
' ==========================================================================================

' ==========================================================================================
' 1. CONFIG & GLOBALS (全局变量与配置)
' ==========================================================================================
Private Const SHEET_CONFIG As String = "Sys_Config"
Private Const SHEET_USER_CONFIG As String = "User_Config" ' [新增] 用户配置表名
Private Const SHEET_MASTER As String = "Master"
Private Const CODE_SENTINEL As String = "9999999"
Private Const COL_BORDER_CHK As Integer = 2
Private Const DIR_DATA As String = "Data"
Private ConfigDict As Object
Private BlacklistDict As Object ' [新增] 存放黑名单关键词
Private Const COL_CODE As Variant = "A"
' [新增] 锁定与解锁配置
Private Const LOCK_PASSWORD As String = "888888" ' 用于锁定和解锁的密码
Private Const LOCK_COLOR As Long = 12632256          ' 锁定时的背景色：浅灰色 RGB(192, 192, 192)
' [新增] 公共变量：用于在窗体和模块间共享配置
Public Pub_ConfigDict As Object
Public Pub_BlacklistDict As Object
Public Pub_AccountDict As Object ' [新增] 存放特征码与会计科目的映射关系
' [新增] 全局配置单例控制
Public g_IsConfigLoaded As Boolean ' 标志位：True表示已加载，False表示未加载
Public g_TableConfig As Object      ' [新增] User_Config (表格样式，移到这里统一管理)

' 【核心配置区】在此处调整样式的初始默认值 (方便维护)
' =================================================================
' [Logic] 确保 Word 样式存在并更新 (修复对齐逻辑)
' 修复日志：
'   - 增加了对 Alignment 的值校正。
'   - Excel配置: 1=左, 2=居中, 3=右, 4=两端
'   - Word标准:  0=左, 1=居中, 2=右, 3=两端
'   - 逻辑: Word值 = Excel值 - 1
' =================================================================
Private Sub Add_Default_Style_Configs(ws As Worksheet)
    ' 辅助函数：写入默认配置 (Key, Value)
    ' 只有当该 Key 不存在时才写入，避免覆盖用户已调整的参数
    
    ' --- H: 报告主标题 (如：审计报告) ---
    WriteDefault ws, "H_FontName", "黑体"
    WriteDefault ws, "H_FontSize", "18"       ' 小二
    WriteDefault ws, "H_Alignment", "1"       ' 1=居中, 0=左, 2=右
    WriteDefault ws, "H_Bold", "0"            ' 1=加粗
    WriteDefault ws, "H_LineSpacing", "1.5"   ' 1.5倍行距
    WriteDefault ws, "H_SpaceBefore", "0"
    WriteDefault ws, "H_SpaceAfter", "0"
    
    ' --- H0: 报告副标题/小标题 ---
    WriteDefault ws, "H0_FontName", "宋体"
    WriteDefault ws, "H0_FontSize", "10.5"    ' 五号
    WriteDefault ws, "H0_Alignment", "1"      ' 居中
    WriteDefault ws, "H0_Bold", "0"
    WriteDefault ws, "H0_LineSpacing", "1.5"
    
    ' --- H1: 一级标题 (如：一、公司基本情况) ---
    WriteDefault ws, "H1_FontName", "黑体"
    WriteDefault ws, "H1_FontSize", "10.5"      ' 五号
    WriteDefault ws, "H1_Alignment", "0"      ' 左对齐
    WriteDefault ws, "H1_FirstLineIndent", "2" ' 首行缩进2字符
    WriteDefault ws, "H1_LineSpacing", "1.5"    ' 1.5倍行间距
    WriteDefault ws, "H1_SpaceBefore", "0"  ' 段前0.5行
    WriteDefault ws, "H1_SpaceAfter", "0"   ' 段后0.5行
    
    ' --- H2: 二级标题 (如：(一) 历史沿革) ---
    WriteDefault ws, "H2_FontName", "宋体"
    WriteDefault ws, "H2_FontSize", "10.5"      ' 小四
    WriteDefault ws, "H2_Alignment", "0"        ' 左对齐
    WriteDefault ws, "H2_FirstLineIndent", "2" ' 首行缩进2字符
    WriteDefault ws, "H2_LineSpacing", "1.5"      ' 1.5倍行间距
    
    ' --- H3: 三级标题 (如：1. 设立情况) ---
    WriteDefault ws, "H3_FontName", "宋体"
    WriteDefault ws, "H3_FontSize", "10.5"
    WriteDefault ws, "H3_Bold", "0"           ' 加粗
    WriteDefault ws, "H3_Alignment", "0"
    WriteDefault ws, "H3_FirstLineIndent", "2"
    WriteDefault ws, "H3_LineSpacing", "1.5"
    
    ' --- P: 正文 ---
    WriteDefault ws, "P_FontName", "宋体"
    WriteDefault ws, "P_FontSize", "10.5"     ' 五号
    WriteDefault ws, "P_Alignment", "0"       ' 3=两端对齐
    WriteDefault ws, "P_FirstLineIndent", "2" ' 首行缩进2字符
    WriteDefault ws, "P_LineSpacing", "1.5"
    
    ' --- D: 单位说明 (如：单位：元) ---
    WriteDefault ws, "D_FontName", "宋体"
    WriteDefault ws, "D_FontSize", "10.5"
    WriteDefault ws, "D_Alignment", "2"       ' 右对齐
    WriteDefault ws, "D_LineSpacing", "1.5"     ' 单倍
    
    ' --- Table: 表格通用样式 ---
    WriteDefault ws, "Table_FontName", "宋体"
    WriteDefault ws, "Table_FontSize", "9"
    WriteDefault ws, "Table_RowHeight", "18"   ' 0=自动, 固定值则填数字
    WriteDefault ws, "Table_LineSpacing", "1"   ' 单倍行间距
    WriteDefault ws, "Table_Alignment", "0"   ' 左对齐(指表格内文字)
End Sub
' ==========================================================================================
' [修正版] 供 UserForm 调用的公共接口
' ==========================================================================================

' [修正] 批量处理入口
Sub P2_Run_Batch_From_Form(filesCollection As Collection, formObj As Object)
    Dim wsMaster As Worksheet
    Dim wbMain As Workbook ' <--- 新增：专门记录用户的主文件
    Dim wbSource As Workbook
    Dim fPath As Variant
    Dim filename As String, compName As String
    Dim i As Integer, total As Integer

    
    Set wbMain = Application.ActiveWorkbook
    If wbMain Is Nothing Then Exit Sub
    ' 1. 初始化
    wbMain.Activate
    Set wsMaster = Lib_GetMasterSheet()
    
    Call Init_Global_Config
    
    If Pub_ConfigDict Is Nothing Then Exit Sub
    
    ' 绑定配置 (因为现在变量是 Public 的，其实可以直接用 Pub_ConfigDict)
    Set ConfigDict = Pub_ConfigDict
    

    ' [新增调试代码]
    If ConfigDict.count = 0 Then MsgBox "严重错误：配置规则字典为空！请检查 Sys_Config 表是否有数据。", vbCritical: Exit Sub
    Set BlacklistDict = Pub_BlacklistDict
    
    total = filesCollection.count
    i = 0
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With
    
    ' 2. 循环处理
    For Each fPath In filesCollection
        i = i + 1
        filename = Dir(fPath)
        
        If filename <> "" Then
            compName = Left(filename, InStrRev(filename, ".") - 1)
            
            ' A. 更新进度条
            formObj.Update_Progress i, total, compName
            
            ' ==========================================================
            ' B. 【核心修复】 同步公司名到 User_Config
            '    这步是把公司名写入下拉菜单数据源的关键
            ' ==========================================================
            Call Tool_UpdateCompanyList(compName)
            
            ' C. 删除旧数据 & 提取新数据
            Call P2_DeleteOldData(wsMaster, compName)
            
            ' ==================================================================
            ' [修改] 打开文件并指定读取 "科目注释" 工作表
            ' ==================================================================
            On Error Resume Next
            
            If IsFileOpen(CStr(fPath)) Then
            ' 策略 A: 跳过并记录日志
            Debug.Print "文件被占用，跳过: " & fPath
            GoTo NextFile
            DoEvents
            End If
            
            Set wbSource = Workbooks.Open(fPath, ReadOnly:=True, UpdateLinks:=False)
            
            If Not wbSource Is Nothing Then
                Dim wsTarget As Worksheet
                Set wsTarget = Nothing
                
                ' 1. 尝试获取名为 "科目注释" 的表
                On Error Resume Next
                Set wsTarget = wbSource.Sheets("科目注释")
                On Error GoTo 0
                
                ' 2. (可选兜底) 如果找不到 "科目注释"，尝试找 "附注模板" 或直接跳过
                'If wsTarget Is Nothing Then
                '    On Error Resume Next
                '    Set wsTarget = wbSource.Sheets("附注模板") ' 尝试备用名
                '    On Error GoTo 0
                'End If
                
                ' 3. 执行提取
                If Not wsTarget Is Nothing Then
                    ' 只有找到了正确的表，才执行提取
                    Call P2_ExtractData_Hybrid(wsTarget, filename, compName, wsMaster)
                Else
                    ' 如果找不到目标表，记录错误或忽略
                    ' Debug.Print "跳过文件 [" & fileName & "]: 未找到 [科目注释] 表"
                End If
                
                wbSource.Close SaveChanges:=False
            End If
            On Error GoTo 0
            ' ==================================================================
        End If
NextFile:  ' <--- 加上这一行标签（注意后面有个冒号）
    Next fPath
    
    ' 3. 恢复环境
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With
    

    
    ' 强制刷新一次 C1 单元格的下拉菜单缓存（可选）
    On Error Resume Next
    ActiveWorkbook.ActiveSheet.Range("C1").Validation.Delete
    ActiveWorkbook.ActiveSheet.Range("C1").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
             xlBetween, Formula1:="=OFFSET('User_Config'!$B$2,0,0,COUNTA('User_Config'!$B:$B)-1,1)"
    On Error GoTo 0
    
    
    MsgBox "汇总完成！共处理 " & total & " 个文件。", vbInformation
    
    If Not wsMaster Is Nothing Then
        wsMaster.Visible = xlSheetHidden
    End If
    
End Sub
' ==========================================================================================
' 2. CONTROLLERS (用户交互入口)
' ==========================================================================================

' [入口 1] 模板生成模式
Sub Main_GenerateTemplate()
    On Error GoTo ErrorHandler
    
    Dim answer As VbMsgBoxResult
    answer = MsgBox("【生成模板模式】" & vbCrLf & vbCrLf & _
                    "1. 常规扫描：生成基础代码与配置。" & vbCrLf & _
                    "2. 智能检测：识别含有重复项目的复杂表格。" & vbCrLf & _
                    "3. 交互拆分：协助您将复杂表按板块拆分(如_01, _02)。" & vbCrLf & vbCrLf & _
                    "是否继续？", vbQuestion + vbYesNo, "系统初始化")
    
    If answer = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
    
    ' 1. 执行基础生成 (P1 Core)
    Call P1_Core_Generate(ActiveWorkbook.ActiveSheet)
    
    ' 2. 执行复杂表后处理
    Application.ScreenUpdating = True
    Call P1_PostProcess_ComplexTables(ActiveWorkbook.ActiveSheet)
    
    MsgBox "模板生成与处理完毕！", vbInformation
           
ExitHandler:
    Application.ScreenUpdating = True
    'Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "运行时错误 [Generate]: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

' [入口 2] 数据聚合模式 (更新：改为文件选择模式)
Sub Main_RunAggregation()
    On Error GoTo ErrorHandler
    
    Dim answer As VbMsgBoxResult
    answer = MsgBox("【数据聚合模式】" & vbCrLf & vbCrLf & _
                    "即将打开文件选择窗口。" & vbCrLf & _
                    "您可以选择一个或多个财务报表文件进行汇总。" & vbCrLf & vbCrLf & _
                    "逻辑说明：" & vbCrLf & _
                    "1. 增量模式：如果选择了从未汇总过的新文件，将直接追加数据。" & vbCrLf & _
                    "2. 更新模式：如果选择了已存在的公司文件，系统将清除旧数据并重新汇总。" & vbCrLf & vbCrLf & _
                    "是否继续？", vbQuestion + vbYesNo, "开始聚合")
    
    If answer = vbNo Then Exit Sub
    
    ' 调用 Phase 2 核心
    Call P2_Core_Aggregate
    
ExitHandler:
    ' 清理内存
    Set ConfigDict = Nothing
    Set BlacklistDict = Nothing ' [新增清理]
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    MsgBox "聚合过程中断: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub
' [入口 3] 手动拆分工具 (已升级)
' 功能：允许用户手动选中表格中的任意一行，强制触发“复杂表拆分”流程
Sub Main_ManualSplit()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim selRow As Long
    Dim baseCode As String
    Dim rngTable As Range
    
    Set ws = ActiveWorkbook.ActiveSheet
    selRow = ActiveCell.Row
    
    ' 1. 获取基础代码
    baseCode = Trim(CStr(ws.Cells(selRow, 1).Value))
    
    ' 2. 校验选中区域
    If baseCode = "" Or baseCode = CODE_SENTINEL Then
        MsgBox "请选中表格内部的数据行 (A列必须有有效代码)。", vbExclamation
        Exit Sub
    End If
    
    ' 3. 确定整个表格范围
    Set rngTable = P1_FindTableRange(ws, baseCode)
    If rngTable Is Nothing Then
        MsgBox "无法定位该代码 [" & baseCode & "] 的完整表格范围。", vbCritical
        Exit Sub
    End If
    
    ' 4. 选中范围并提示确认
    rngTable.Select
    ws.Cells(selRow, 1).Select
    
    If MsgBox("即将在表格 [" & baseCode & "] 上执行手动拆分。" & vbCrLf & vbCrLf & _
              "操作流程：" & vbCrLf & _
              "1. 接下来请选择作为【分段点】的标题单元格。" & vbCrLf & _
              "2. 系统将依据选择，把代码拆分为 " & baseCode & "_01, _02 等。" & vbCrLf & _
              "3. 自动在配置表中复制对应的取数规则。" & vbCrLf & vbCrLf & _
              "是否继续？", vbQuestion + vbYesNo, "手动拆分") = vbNo Then Exit Sub
    
    ' 5. 调用核心拆分逻辑 (复用 Phase 1 的逻辑)
    Call P1_ExecuteSplit(ws, rngTable, baseCode)
    
    MsgBox "拆分完成！请检查 A 列及 Sys_Config 配置表。", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "操作出错: " & Err.Description, vbCritical
End Sub

' [入口 4] 公式注入工具 (逻辑更新)
' 功能：遍历表格，将【显式白色背景】的单元格填充 SUMIFS+XLOOKUP 公式
Sub Main_InjectFormulas()
    On Error GoTo ErrorHandler
    
    Dim answer As VbMsgBoxResult
    answer = MsgBox("【一键注入公式】" & vbCrLf & vbCrLf & _
                    "即将执行以下操作：" & vbCrLf & _
                    "1. 扫描当前工作表中所有带有代码(A列)的行。" & vbCrLf & _
                    "2. 仅针对【显式白色背景(vbWhite)】的单元格写入公式。" & vbCrLf & _
                    "3. 【无填充(No Fill)】或其他颜色的单元格将被跳过。" & vbCrLf & vbCrLf & _
                    "是否继续？", vbQuestion + vbYesNo, "公式注入")
    
    If answer = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Call Tool_Inject_Formula_Logic(ActiveWorkbook.ActiveSheet)
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "公式注入完成！", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "注入失败: " & Err.Description, vbCritical
End Sub
' [入口 5] 打开设置面板 (新增)
' 功能：当用户需要修改黑名单时，调用此功能将 User_Config 表显示出来
Sub Main_OpenSettings()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(SHEET_USER_CONFIG)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "尚未初始化设置表，请先运行【生成模板】。", vbExclamation
        Exit Sub
    End If
    
    ' 取消隐藏并激活
    ws.Visible = xlSheetVisible '.Visible = xlSheetVisible/xlSheetVisible 隐身/不隐身
    ws.Activate
    
    MsgBox "已打开设置表。" & vbCrLf & _
           "1. 您可以在 A 列修改【黑名单】关键词。" & vbCrLf & _
           "2. 修改完成后，您可以右键点击下方工作表标签手动隐藏，或保持原状。", vbInformation, "系统设置"
End Sub

' [入口 6] 启动综合控制台
Sub Main_LaunchDashboard()
    ' 非模态显示，允许用户在打开窗口的同时操作 Excel 单元格（可选）
    Frm_Main.Show vbModeless

    ' 模态显示（推荐），用户必须先完成窗口操作
    'Frm_Main.Show
End Sub

' [入口 7] 数据保护工具 (新增)
' 功能：锁定所有非公式且无底色/非白色底色的已编辑单元格
Sub Main_ProtectNonFormulaData()
    On Error GoTo ErrorHandler

    Dim answer As VbMsgBoxResult
    Const sPassword As String = "888888" ' 默认密码

    answer = MsgBox("【锁定数据模式】" & vbCrLf & vbCrLf & _
                    "警告：此操作将使用密码 [" & sPassword & "] 保护当前活动工作表！" & vbCrLf & vbCrLf & _
                    "操作逻辑：" & vbCrLf & _
                    "1. 解锁所有单元格。" & vbCrLf & _
                    "2. 锁定所有【非空】且【无底色或非白色底色】的单元格（被认为是手工录入数据）。" & vbCrLf & _
                    "3. 使用密码保护工作表。", vbQuestion + vbYesNo, "锁定数据")

    If answer = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 调用核心保护逻辑
    Call Tool_Protect_Data_By_Color(ActiveWorkbook.ActiveSheet)

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "数据锁定与工作表保护完成！", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "数据锁定失败: " & Err.Description, vbCritical
End Sub

' [入口 8] 一键解锁工具 (新增)
' 功能：取消保护，并将锁定的单元格颜色恢复为无填充
Sub Main_Unprotect_Data()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim enteredPassword As String
    Dim cell As Range
    Dim lastRow As Long, lastCol As Long

    Set ws = ActiveWorkbook.ActiveSheet
    
    ' 1. 检查工作表是否已保护
    If ws.ProtectContents = False Then
        MsgBox "当前工作表未受保护。", vbInformation
        Exit Sub
    End If
    
    ' 2. 输入密码
    enteredPassword = Application.InputBox("请输入解锁密码:", "工作表解锁", Type:=2)
    
    ' 检查是否取消
    If enteredPassword = "False" Or enteredPassword = "" Then Exit Sub
    
    ' 3. 尝试取消保护
    If enteredPassword <> LOCK_PASSWORD Then
        MsgBox "密码错误，无法解锁。", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ws.Unprotect enteredPassword
    
    ' 4. 恢复颜色
    On Error Resume Next
    If ws.UsedRange.Cells.count < 2 Then GoTo ColorRestoreEnd
    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.count).Row
    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.count).Column
    On Error GoTo 0
    
    For Each cell In ws.Range("A1", ws.Cells(lastRow, lastCol)).Cells
        ' 仅恢复被锁定颜色（浅灰色）的单元格
        If cell.Interior.Color = LOCK_COLOR Then
            cell.Interior.ColorIndex = xlNone ' 恢复为无填充色
        End If
    Next cell

ColorRestoreEnd:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "工作表已成功解锁，数据区域颜色已恢复。", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "解锁失败: " & Err.Description, vbCritical
End Sub
' [入口 9] 格式化白色区域工具 (新增)
' 功能：将填充色为白色的区域设置为右对齐、数值格式、千分位
Sub Main_Format_WhiteCells()
    On Error GoTo ErrorHandler

    Dim answer As VbMsgBoxResult
    
    answer = MsgBox("【格式化白色区域模式】" & vbCrLf & vbCrLf & _
                    "即将对当前工作表执行以下操作：" & vbCrLf & _
                    "1. 查找所有显式设置为【纯白色填充】的单元格。" & vbCrLf & _
                    "2. 将这些单元格设置为【右对齐】。" & vbCrLf & _
                    "3. 将这些单元格格式设置为【#,##0.00】(数值，千分位，两位小数)。" & vbCrLf & vbCrLf & _
                    "是否继续？", vbQuestion + vbYesNo, "格式化数据")

    If answer = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    
    ' 调用核心格式化逻辑
    Call Tool_Format_WhiteCells(ActiveWorkbook.ActiveSheet)
    
    Application.ScreenUpdating = True

    MsgBox "白色区域格式化完成！", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "格式化失败: " & Err.Description, vbCritical
End Sub

' ==========================================================================================
' 3. PHASE 1 ENGINE (模板生成核心)
' ==========================================================================================

' P1 核心：扫描表格、打码、写配置
' P1 核心：扫描表格、打码、写配置 (含 User_Config 初始化和 C1 下拉生成)
Private Sub P1_Core_Generate(ws As Worksheet)
    Dim wsConfig As Worksheet, wsUserCfg As Worksheet
    Dim lastRow As Long, i As Long
    Dim nextCode As Long
    Dim currentCode As String
    Dim currHasBorder As Boolean, prevHasBorder As Boolean
    
    ' 1. 初始化两张配置表
    Set wsConfig = Lib_InitConfigSheet()
    Set wsUserCfg = Lib_InitUserConfigSheet() ' [新增] 初始化用户配置表(黑名单+公司列)
    
    ' 2. 准备主表环境
    lastRow = ws.Cells(ws.Rows.count, COL_BORDER_CHK).End(xlUp).Row
    If lastRow < 10 Then lastRow = 100
    
    With ws.Columns("A")
        .ClearContents
        .NumberFormat = "@"
    End With
    
    ' [新增] 设置 C1 单元格下拉菜单 (引用 User_Config 的 B 列)
    ' 使用 OFFSET 函数创建动态引用，这样 B 列增加公司时，下拉菜单自动更新
    With ws.Range("C1")
        .Value = "合并数" ' 默认值
        .Interior.Color = RGB(255, 255, 204) ' 浅黄色提示可输入
        With .Validation
            .Delete
            ' 公式含义：从 User_Config!B2 开始，高度为 B列非空行数-1 (减去标题)
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
             xlBetween, Formula1:="=OFFSET('" & SHEET_USER_CONFIG & "'!$B$2,0,0,COUNTA('" & SHEET_USER_CONFIG & "'!$B:$B)-1,1)"
        End With
    End With
    
    ' 3. 状态机扫描 (保持原有逻辑)
    nextCode = 100000
    prevHasBorder = False
    
    For i = 1 To lastRow + 5
        If i <= lastRow + 2 Then
            currHasBorder = Lib_CheckRightBorder(ws.Cells(i, COL_BORDER_CHK))
        Else
            currHasBorder = False
        End If
        
        If (Not prevHasBorder) And currHasBorder Then
            nextCode = nextCode + 1
            currentCode = CStr(nextCode)
            ws.Cells(i, 1).Value = currentCode
            Call P1_AnalyzeHeader(ws, i, currentCode, wsConfig)
        ElseIf prevHasBorder And currHasBorder Then
            ws.Cells(i, 1).Value = currentCode
        ElseIf prevHasBorder And (Not currHasBorder) Then
            ws.Cells(i - 1, 1).Value = CODE_SENTINEL
        End If
        prevHasBorder = currHasBorder
    Next i
    
    wsConfig.Columns.AutoFit
    wsUserCfg.Columns.AutoFit
End Sub

' P1 后处理：复杂表检测与交互拆分
Sub P1_PostProcess_ComplexTables(ws As Worksheet)
    Dim complexCodes As Collection
    Dim code As Variant
    Dim rngTable As Range
    Dim userResp As VbMsgBoxResult
    
    ' 1. 扫描所有表格
    Set complexCodes = P1_ScanForComplexTables(ws)
    
    If complexCodes.count = 0 Then Exit Sub
    
    MsgBox "系统检测到 " & complexCodes.count & " 个复杂表格（含重复项目）。" & vbCrLf & _
           "即将进入逐个确认模式。", vbExclamation, "复杂表检测"
    
    ' 2. 遍历每一个复杂表
    For Each code In complexCodes
        Set rngTable = P1_FindTableRange(ws, CStr(code))
        
        If Not rngTable Is Nothing Then
            rngTable.Select
            ws.Cells(rngTable.Row, 1).Select
            
            userResp = MsgBox("检测到表格 [" & code & "] 存在重复项目。" & vbCrLf & vbCrLf & _
                              "是否需要拆分？" & vbCrLf & _
                              "- 是：手动选择分段标题，系统自动加后缀(_01, _02)。" & vbCrLf & _
                              "- 否：保持原样。", vbQuestion + vbYesNo, "处理复杂表")
            
            If userResp = vbYes Then
                Call P1_ExecuteSplit(ws, rngTable, CStr(code))
            End If
        End If
    Next code
End Sub

' P1 辅助：执行拆分逻辑 (确保参数传递正确)
Private Sub P1_ExecuteSplit(ws As Worksheet, rngTable As Range, baseCode As String)
    Dim splitRanges As Range
    Dim cell As Range
    Dim startRow As Long, endRow As Long
    Dim segmentCount As Integer
    Dim rowsArr() As Long, count As Integer
    Dim cleanBaseCode As String
    
    ' 1. 清洗基础代码 (去除 _01 等后缀，如果已经是拆分过的代码)
    cleanBaseCode = Trim(CStr(baseCode))
    If InStr(cleanBaseCode, "_") > 0 Then
        cleanBaseCode = Split(cleanBaseCode, "_")(0)
    End If
    
    ' 2. 用户交互
    On Error Resume Next
    Set splitRanges = Application.InputBox("请用鼠标选择【分段标题所在的单元格】（可按住 Ctrl 多选）。", _
                                           "选择分段点", Type:=8)
    On Error GoTo 0
    If splitRanges Is Nothing Then Exit Sub
    
    ' 3. 排序选区
    count = 0
    ReDim rowsArr(1 To splitRanges.Cells.count)
    For Each cell In splitRanges
        count = count + 1
        rowsArr(count) = cell.Row
    Next cell
    
    Dim i As Integer, j As Integer, temp As Long
    For i = 1 To count - 1
        For j = i + 1 To count
            If rowsArr(i) > rowsArr(j) Then
                temp = rowsArr(i): rowsArr(i) = rowsArr(j): rowsArr(j) = temp
            End If
        Next j
    Next i
    
    ' 4. 执行拆分与配置复制
    startRow = rngTable.Row
    endRow = rngTable.Row + rngTable.Rows.count - 1
    
    Application.ScreenUpdating = False
    
    Dim currentRow As Long
    Dim activeSuffix As String
    Dim ptr As Integer
    
    ptr = 1
    segmentCount = 0
    
    For currentRow = startRow To endRow
        ' 检查分段点
        If ptr <= count Then
            If currentRow = rowsArr(ptr) Then
                segmentCount = segmentCount + 1
                activeSuffix = "_" & Format(segmentCount, "00")
                ptr = ptr + 1
                
                ' 【关键】调用复制配置: 源代码 -> 新代码
                ' 例如：把 10001 的规则复制一份给 10001_01
                Call Tool_DuplicateConfig(cleanBaseCode, cleanBaseCode & activeSuffix)
            End If
        End If
        
        ' 更新 A 列代码 (仅当行内代码包含基础代码时才更新，避免覆盖 9999999)
        If segmentCount > 0 Then
            Dim curRowCode As String
            curRowCode = Trim(CStr(ws.Cells(currentRow, 1).Value))
            ' 只要包含基础代码（比如当前是 10001，或者是旧的 10001_Old），都更新
            If InStr(curRowCode, cleanBaseCode) > 0 And curRowCode <> "9999999" Then
                ws.Cells(currentRow, 1).Value = cleanBaseCode & activeSuffix
            End If
        End If
    Next currentRow
    
    Application.ScreenUpdating = True
End Sub

' P1 辅助：扫描复杂表
Private Function P1_ScanForComplexTables(ws As Worksheet) As Collection
    Dim result As New Collection
    Dim dictCheck As Object, dictCodes As Object
    Dim lastRow As Long, r As Long
    Dim code As String, val As String
    Dim currCode As String, prevCode As String
    
    Set dictCodes = CreateObject("Scripting.Dictionary")
    Set dictCheck = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    For r = 1 To lastRow
        currCode = Trim(CStr(ws.Cells(r, 1).Value))
        If currCode = CODE_SENTINEL Then currCode = ""
        
        If currCode <> prevCode Then Set dictCheck = CreateObject("Scripting.Dictionary")
        
        If currCode <> "" Then
            val = Trim(CStr(ws.Cells(r, COL_BORDER_CHK).Value))
            If val <> "" And val <> "合计" And Not IsNumeric(val) Then
                If dictCheck.exists(val) Then
                    If Not dictCodes.exists(currCode) Then
                        dictCodes.Add currCode, True
                        result.Add currCode
                    End If
                Else
                    dictCheck.Add val, True
                End If
            End If
        End If
        prevCode = currCode
    Next r
    Set P1_ScanForComplexTables = result
End Function

' P1 辅助：获取 Code 范围
Private Function P1_FindTableRange(ws As Worksheet, code As String) As Range
    Dim fst As Range, lst As Range
    On Error Resume Next
    Set fst = ws.Columns(1).Find(What:=code, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext)
    If fst Is Nothing Then Exit Function
    Set lst = ws.Columns(1).Find(What:=code, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)
    Set P1_FindTableRange = ws.Range(fst, lst)
    On Error GoTo 0
End Function

' 辅助：分析表头
' 辅助：分析表头结构并写入配置 (完全还原 V3.0)
Private Sub P1_AnalyzeHeader(ws As Worksheet, startRow As Long, code As String, wsCfg As Worksheet)
    Dim lastCol As Long, c As Long
    Dim headerDepth As Integer
    Dim fullMetric As String
    Dim tableTitle As String
    
    ' 1. 智能计算表头深度 (遇数字停止)
    headerDepth = P1_CalcSmartHeaderDepth(ws, startRow)
    
    ' 2. 查找表格标题
    tableTitle = Lib_FindTableTitle(ws, startRow)
    
    ' 3. 计算表格最大宽度 (全表头扫描)
    Dim maxCol As Long, r_chk As Long, curCol As Long
    maxCol = 0
    For r_chk = startRow To startRow + headerDepth - 1
        curCol = ws.Cells(r_chk, ws.Columns.count).End(xlToLeft).Column
        If curCol > maxCol Then maxCol = curCol
    Next r_chk
    If maxCol < 5 Then maxCol = 10
    lastCol = maxCol
    
    ' 4. 写入 Config
    Dim nextRow As Long
    For c = 3 To lastCol
        fullMetric = P1_GetMergedHeader(ws, c, startRow, headerDepth)
        
        If fullMetric <> "" Then
            nextRow = wsCfg.Cells(wsCfg.Rows.count, "A").End(xlUp).Row + 1
            wsCfg.Cells(nextRow, 1).Value = code
            wsCfg.Cells(nextRow, 2).Value = c
            wsCfg.Cells(nextRow, 3).Value = fullMetric
            wsCfg.Cells(nextRow, 4).Value = tableTitle
        End If
    Next c
End Sub

' 辅助：计算表头深度
Private Function P1_CalcSmartHeaderDepth(ws As Worksheet, startRow As Long) As Integer
    Dim r As Integer, checkRow As Long
    Dim valC As Variant, strVal As String
    Dim valB As String
    
    P1_CalcSmartHeaderDepth = 1
    
    ' 向下探测最多 5 行
    For r = 1 To 4
        checkRow = startRow + r
        
        ' 1. 边框检查
        If Not Lib_CheckRightBorder(ws.Cells(checkRow, 2)) Then Exit Function
        
        ' 2. 获取 C 列的值
        valC = ws.Cells(checkRow, 3).Value
        
        ' 【核心修复】如果遇到错误值 (如 #REF!, #N/A)，立即停止，视为数据行
        If isError(valC) Then Exit Function
        
        ' 3. 数字检测
        strVal = Trim(CStr(valC))
        strVal = Replace(strVal, ",", "")
        If IsNumeric(strVal) And strVal <> "" Then Exit Function
        
        ' 4. 关键词辅助 (B列)
        valB = Replace(ws.Cells(checkRow, 2).Value, " ", "")
        If InStr(valB, "金额") > 0 Or InStr(valB, "比例") > 0 Or InStr(valB, "准备") > 0 Or _
           InStr(valB, "价值") > 0 Or InStr(valB, "计提") > 0 Then
            P1_CalcSmartHeaderDepth = r + 1
        Else
            ' 默认延伸
            P1_CalcSmartHeaderDepth = r + 1
        End If
    Next r
End Function

' 辅助：拼接表头
' 辅助：纵向拼接表头字符串 (完全还原 V3.0)
Private Function P1_GetMergedHeader(ws As Worksheet, col As Long, startRow As Long, depth As Integer) As String
    Dim r As Long, txt As String, res As String, cleanTxt As String
    Dim cellVal As Variant
    
    res = ""
    For r = startRow To startRow + depth - 1
        ' 获取单元格值（兼容合并单元格）
        If ws.Cells(r, col).MergeCells Then
            cellVal = ws.Cells(r, col).MergeArea.Cells(1, 1).Value
        Else
            cellVal = ws.Cells(r, col).Value
        End If
        
        ' 【核心修复】检查是否为错误值
        If isError(cellVal) Then
            txt = "" ' 如果是错误值，视为空，不拼接
        Else
            txt = Trim(CStr(cellVal))
        End If
        
        ' 清洗
        txt = Replace(txt, Chr(10), ""): txt = Replace(txt, Chr(13), "")
        txt = Replace(txt, " ", ""): txt = Replace(txt, Chr(160), "")
        
        ' 数字过滤
        cleanTxt = Replace(txt, ",", ""): cleanTxt = Replace(cleanTxt, ".", "")
        
        ' 拼接逻辑
        If Not IsNumeric(cleanTxt) And txt <> "" Then
            If res = "" Then
                res = txt
            ElseIf InStr(res, txt) = 0 Then
                ' 只有不重复时才拼接
                res = res & "_" & txt
            End If
        End If
    Next r
    P1_GetMergedHeader = res
End Function


' ==========================================================================================
' 4. PHASE 2 ENGINE (数据聚合核心 - 保持原样)
' ==========================================================================================

' P2 主逻辑：文件选择与聚合 (从自动遍历改为手动多选)
' P2 主逻辑：文件选择与聚合 (自动同步公司名到 User_Config)
Private Sub P2_Core_Aggregate()
    Dim wbSource As Workbook, wsMaster As Worksheet
    Dim t As Single, fileCount As Integer
    Dim fd As FileDialog
    Dim vFile As Variant
    Dim fullPath As String, filename As String, compName As String
    
    t = Timer
    Set wsMaster = Lib_GetMasterSheet()
    Call P2_LoadConfigToMemory ' 这里会加载 User_Config 的黑名单
    If ConfigDict Is Nothing Then Exit Sub
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "请选择需要汇总的财务报表文件"
        .Filters.Clear: .Filters.Add "Excel Files", "*.xlsx; *.xls; *.xlsm"
        .AllowMultiSelect = True
        .InitialFileName = ActiveWorkbook.path & "\" & DIR_DATA & "\"
        
        If .Show = -1 Then
            With Application
                .ScreenUpdating = False: .DisplayAlerts = False: .Calculation = xlCalculationManual
            End With
            
            fileCount = 0
            For Each vFile In .SelectedItems
                fileCount = fileCount + 1
                fullPath = vFile
                filename = Dir(fullPath)
                compName = Left(filename, InStrRev(filename, ".") - 1)
                
                Application.StatusBar = "处理中: " & compName
                
                ' [新增] 同步公司名称到 User_Config 表
                Call Tool_UpdateCompanyList(compName)
                Call P2_DeleteOldData(wsMaster, compName)
                
                ' ==========================================================
                ' [修复] 指定读取 "科目注释" 表 (原为 Sheets(1))
                ' ==========================================================
                On Error Resume Next
                Set wbSource = Workbooks.Open(fullPath, ReadOnly:=True, UpdateLinks:=False)
                
                If Not wbSource Is Nothing Then
                    Dim wsTarget As Worksheet
                    Set wsTarget = Nothing
                    
                    ' 优先找 "科目注释"
                    On Error Resume Next
                    Set wsTarget = wbSource.Sheets("科目注释")
                    On Error GoTo 0
                    
                   
                    ' 执行提取
                    If Not wsTarget Is Nothing Then
                        Call P2_ExtractData_Hybrid(wsTarget, filename, compName, wsMaster)
                    End If
                    
                    wbSource.Close SaveChanges:=False
                End If
                On Error GoTo 0
                ' ==========================================================
                
            Next vFile
            
            Application.StatusBar = False
            Application.Calculation = xlCalculationAutomatic ' 恢复计算以便 C1 下拉公式生效
            Application.ScreenUpdating = True
            
            MsgBox "聚合完成！" & vbCrLf & "已同步公司列表至 " & SHEET_USER_CONFIG & " 表。", vbInformation
        Else
            MsgBox "取消操作。", vbExclamation
        End If
    End With
End Sub

Private Sub P2_ExtractData_Hybrid(wsIn As Worksheet, fName As String, comp As String, wsOut As Worksheet)
    Dim arrSrc As Variant, arrOut() As Variant
    Dim lastRow As Long, lastCol As Long, r As Long, outIdx As Long
    Dim code As String, currItemName As String, lastItemName As String
    Dim colRules As Collection, rule As Variant, parts() As String
    Dim targetCol As Long, metric As String, val As Variant
    Dim accName As String ' [新增]
    
    lastRow = wsIn.Cells(wsIn.Rows.count, "A").End(xlUp).Row
    On Error Resume Next
    lastCol = wsIn.Cells.Find(What:="*", After:=wsIn.Cells(1, 1), LookIn:=xlFormulas, _
              LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    On Error GoTo 0
    If lastCol < 10 Then lastCol = 20
    If lastRow < 2 Then Exit Sub
    
    arrSrc = wsIn.Range(wsIn.Cells(1, 1), wsIn.Cells(lastRow, lastCol)).Value
    ReDim arrOut(1 To lastRow * 10, 1 To 8)
    outIdx = 0: lastItemName = ""
    
    For r = 1 To UBound(arrSrc, 1)
        code = Trim(CStr(wsIn.Cells(r, COL_CODE).Value))
        
        ' 项目名称填充逻辑 (不变)
        If UBound(arrSrc, 2) >= COL_BORDER_CHK Then
            currItemName = Trim(CStr(arrSrc(r, COL_BORDER_CHK)))
            If currItemName <> "" Then lastItemName = currItemName Else If Len(code) > 0 And code <> CODE_SENTINEL Then currItemName = lastItemName
        Else
            currItemName = ""
        End If
        
        If Len(code) > 0 And code <> CODE_SENTINEL Then
            ' [新增] 查找科目名称
            ' 优先使用 Public 变量，如果没有(单文件调试)则默认为空
            accName = ""
            If Not Pub_AccountDict Is Nothing Then
                If Pub_AccountDict.exists(code) Then accName = Pub_AccountDict(code)
            End If
            
            ' 如果当前代码有配置规则
            ' 注意：这里假设 ConfigDict 已经被 Main 或 P2_Run_Batch 赋值了
            ' 如果是单步调试，ConfigDict 可能为空，需要注意
            If Not ConfigDict Is Nothing Then
                If ConfigDict.exists(code) Then
                    Set colRules = ConfigDict(code)
                    For Each rule In colRules
                        parts = Split(rule, "|")
                        targetCol = CLng(parts(0))
                        metric = parts(1)
                        
                        If targetCol <= UBound(arrSrc, 2) Then
                            val = arrSrc(r, targetCol)
                            If Len(CStr(val)) = 0 Then
                                If wsIn.Cells(r, targetCol).MergeCells Then val = wsIn.Cells(r, targetCol).MergeArea.Cells(1, 1).Value
                            End If
                            
                            If Len(CStr(val)) > 0 And Not P2_IsHeaderNoise(val) Then
                                outIdx = outIdx + 1
                                arrOut(outIdx, 1) = fName
                                arrOut(outIdx, 2) = comp
                                arrOut(outIdx, 3) = code
                                arrOut(outIdx, 4) = currItemName
                                arrOut(outIdx, 5) = metric
                                arrOut(outIdx, 6) = val
                                arrOut(outIdx, 7) = Now
                                arrOut(outIdx, 8) = accName ' [新增] 写入科目
                            End If
                        End If
                    Next rule
                End If
            End If
        End If
    Next r
    
' [修改] 输出时 Resize 改为 8 列
' [修复] 输出时 Resize 改为 8 列 (安全写入版)
    If outIdx > 0 Then
        ' === 新增防御性检查 ===
        If wsOut Is Nothing Then
            MsgBox "写入失败：目标工作表对象丢失 (wsOut is Nothing)。", vbCritical
            Exit Sub
        End If
        
        ' === 核心修复：只写入数组的有效部分 ===
        ' 这里的技巧是：Resize 指定了目标区域大小，Excel 会自动截取 arrOut 左上角对应的数据写入
        ' 只要确保 arrOut 足够大即可。
        ' 但为了防止 arrOut 定义得太大导致内存溢出或写入缓慢，我们可以做一个小的优化：
        
        ' 1. 确定写入起点
        Dim rngDest As Range
        Set rngDest = wsOut.Cells(wsOut.Rows.count, "A").End(xlUp).Offset(1, 0)
        
        ' 2. 执行写入 (加上错误捕捉以防万一)
        On Error Resume Next
        rngDest.Resize(outIdx, 8).Value = arrOut
        
        If Err.Number <> 0 Then
            ' 如果直接写入报错，尝试逐行写入 (虽然慢但能救命，通常不会走到这里)
            ' 或者提示用户数据量过大
            MsgBox "写入数据时发生错误 (Err=" & Err.Number & ")，可能数据量过大或包含特殊字符。", vbExclamation
            Err.Clear
        End If
        On Error GoTo 0
        
    End If
End Sub
' 辅助：表头噪音过滤器 (改为查字典模式)
Private Function P2_IsHeaderNoise(val As Variant) As Boolean
    Dim s As String
    ' 清洗数据：去空格、去换行
    s = Replace(Trim(CStr(val)), " ", "")
    s = Replace(s, Chr(10), "")
    s = Replace(s, Chr(13), "")
    s = Replace(s, Chr(160), "") ' 去除不间断空格
    
    If s = "" Then
        P2_IsHeaderNoise = True
        Exit Function
    End If
    
    ' 直接查字典，速度极快
    If Not BlacklistDict Is Nothing Then
        If BlacklistDict.exists(s) Then
            P2_IsHeaderNoise = True
            Exit Function
        End If
    End If
    
    P2_IsHeaderNoise = False
End Function

' 辅助：加载配置到内存 (同时加载取数规则 + 黑名单)
' 辅助：加载配置 (取数规则来自 Sys_Config，黑名单来自 User_Config)
Private Sub P2_LoadConfigToMemory()
    Dim wsCfg As Worksheet, wsUser As Worksheet
    Dim arrCfg As Variant, arrBlack As Variant
    Dim i As Long, lastRow As Long
    
    ' 1. 加载 Sys_Config (规则)
    On Error Resume Next
    Set wsCfg = ActiveWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0
    If wsCfg Is Nothing Then MsgBox "缺少 " & SHEET_CONFIG, vbCritical: Set ConfigDict = Nothing: Exit Sub
    
    Set ConfigDict = CreateObject("Scripting.Dictionary")
    lastRow = wsCfg.Cells(wsCfg.Rows.count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        arrCfg = wsCfg.Range("A2:C" & lastRow).Value
        For i = 1 To UBound(arrCfg, 1)
            Dim c As String: c = Trim(CStr(arrCfg(i, 1)))
            If c <> "" Then
                If Not ConfigDict.exists(c) Then ConfigDict.Add c, New Collection
                ConfigDict(c).Add Trim(CStr(arrCfg(i, 2))) & "|" & Trim(CStr(arrCfg(i, 3)))
            End If
        Next i
    End If
    
    ' 2. 加载 User_Config (黑名单 - A列)
    Set BlacklistDict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set wsUser = ActiveWorkbook.Sheets(SHEET_USER_CONFIG)
    On Error GoTo 0
    
    ' 如果 User_Config 存在，加载 A 列数据
    If Not wsUser Is Nothing Then
        lastRow = wsUser.Cells(wsUser.Rows.count, "A").End(xlUp).Row
        If lastRow >= 2 Then
            arrBlack = wsUser.Range("A2:A" & lastRow).Value
            For i = 1 To UBound(arrBlack, 1)
                Dim bWord As String
                bWord = Replace(Trim(CStr(arrBlack(i, 1))), " ", "")
                If bWord <> "" Then
                    If Not BlacklistDict.exists(bWord) Then BlacklistDict.Add bWord, True
                End If
            Next i
        End If
    End If
    
    ' 如果字典为空(表不存在或被清空)，加载内存默认值防止报错，但通常 Init 时已写入
    If BlacklistDict.count = 0 Then BlacklistDict.Add "合计", True
End Sub

' [修正] 加载配置 (Public 版本 - 增加错误值清洗)
' ==========================================================================================
' [修正版] 加载配置 (Public 版本 - 修复声明重复错误 & 逻辑整理)
' ==========================================================================================
Sub P2_LoadConfigToMemory_Public()
    Dim wsCfg As Worksheet, wsUser As Worksheet
    Dim arrCfg As Variant, arrBlack As Variant, arrMap As Variant
    Dim i As Long, lastRow As Long
    Dim rawCol As Variant, rawMetric As Variant
    Dim strCol As String, strMetric As String
    ' 【修复点1】将变量 c 的声明移到这里，整个过程只声明一次
    Dim c As String
    Dim bWord As String
    Dim mapCode As String, mapAcc As String
    
    ' 初始化字典
    Set Pub_ConfigDict = CreateObject("Scripting.Dictionary")
    Set Pub_BlacklistDict = CreateObject("Scripting.Dictionary")
    Set Pub_AccountDict = CreateObject("Scripting.Dictionary")
    
    ' ==========================================================================
    ' 1. 加载 Sys_Config (取数规则)
    ' ==========================================================================
    On Error Resume Next
    Set wsCfg = ActiveWorkbook.Sheets("Sys_Config")
    On Error GoTo 0
    
    If Not wsCfg Is Nothing Then
        lastRow = wsCfg.Cells(wsCfg.Rows.count, "A").End(xlUp).Row
        If lastRow >= 2 Then
            arrCfg = wsCfg.Range("A2:C" & lastRow).Value
            
            For i = 1 To UBound(arrCfg, 1)
                ' 错误值清洗逻辑
                rawCol = arrCfg(i, 2)
                rawMetric = arrCfg(i, 3)
                
                If isError(rawCol) Then strCol = "0" Else strCol = Trim(CStr(rawCol))
                If isError(rawMetric) Then strMetric = "配置错误" Else strMetric = Trim(CStr(rawMetric))
                
                c = Trim(CStr(arrCfg(i, 1))) ' 使用顶部声明的变量 c
                
                If c <> "" And c <> "9999999" Then
                    If Not Pub_ConfigDict.exists(c) Then Pub_ConfigDict.Add c, New Collection
                    Pub_ConfigDict(c).Add strCol & "|" & strMetric
                End If
            Next i
        End If
    Else
        MsgBox "未找到 Sys_Config 表，无法加载取数规则。", vbExclamation
    End If
    
    ' ==========================================================================
    ' 2. 加载 User_Config (黑名单 + 科目映射)
    ' ==========================================================================
    On Error Resume Next
    Set wsUser = ActiveWorkbook.Sheets("User_Config")
    On Error GoTo 0
    
    If Not wsUser Is Nothing Then
        ' ------------------------------------------------
        ' A. 加载黑名单 (A列)
        ' ------------------------------------------------
        lastRow = wsUser.Cells(wsUser.Rows.count, "A").End(xlUp).Row
        If lastRow >= 2 Then
            arrBlack = wsUser.Range("A2:A" & lastRow).Value
            For i = 1 To UBound(arrBlack, 1)
                If Not isError(arrBlack(i, 1)) Then
                    bWord = Replace(Trim(CStr(arrBlack(i, 1))), " ", "")
                    If bWord <> "" Then
                        If Not Pub_BlacklistDict.exists(bWord) Then Pub_BlacklistDict.Add bWord, True
                    End If
                End If
            Next i
        End If
        
        ' ------------------------------------------------
        ' B. 加载科目映射 (C列=Code, D列=Account)
        ' ------------------------------------------------
        lastRow = wsUser.Cells(wsUser.Rows.count, "C").End(xlUp).Row
        If lastRow >= 2 Then
            arrMap = wsUser.Range("C2:D" & lastRow).Value
            For i = 1 To UBound(arrMap, 1)
                mapCode = "": mapAcc = ""
                If Not isError(arrMap(i, 1)) Then mapCode = Trim(CStr(arrMap(i, 1)))
                If Not isError(arrMap(i, 2)) Then mapAcc = Trim(CStr(arrMap(i, 2)))
                
                If mapCode <> "" Then
                    If Not Pub_AccountDict.exists(mapCode) Then Pub_AccountDict.Add mapCode, mapAcc
                End If
            Next i
        End If
    End If
    
    ' 默认黑名单兜底
    If Pub_BlacklistDict.count = 0 Then Pub_BlacklistDict.Add "合计", True
End Sub
' [修改] 删除旧数据 (扩大范围到 H 列)
Private Sub P2_DeleteOldData(ws As Worksheet, compName As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ' 范围改为 A1:H
    With ws.Range("A1:H" & lastRow)
        .AutoFilter Field:=2, Criteria1:=compName
        On Error Resume Next
        .Offset(1, 0).Resize(lastRow - 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        .AutoFilter
    End With
End Sub


' ==========================================================================================
' 5. TOOLS & COMMON LIB (新增 User_Config 相关逻辑)
' ==========================================================================================

' 辅助：公式注入核心逻辑 (严格颜色匹配版)
Private Sub Tool_Inject_Formula_Logic(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim code As String
    Dim formulaTemplate As String
    Dim currentFormula As String
    Dim checkWs As Worksheet
    
    ' ==============================================================================
    ' [新增] 安全检查：确保 Master 和 Sys_Config 表存在
    ' ==============================================================================
    On Error Resume Next
    Set checkWs = ActiveWorkbook.Sheets(SHEET_MASTER)
    On Error GoTo 0
    If checkWs Is Nothing Then
        MsgBox "错误：未找到 [" & SHEET_MASTER & "] 表！" & vbCrLf & vbCrLf & _
               "公式依赖该表进行取数。" & vbCrLf & _
               "请先运行【数据聚合】功能生成数据仓库，然后再注入公式。", vbCritical, "缺少依赖"
        Exit Sub
    End If
    
    On Error Resume Next
    Set checkWs = ActiveWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0
    If checkWs Is Nothing Then
        MsgBox "错误：未找到 [" & SHEET_CONFIG & "] 表！" & vbCrLf & vbCrLf & _
               "公式依赖该表进行指标匹配。" & vbCrLf & _
               "请先运行【生成模板】功能。", vbCritical, "缺少依赖"
        Exit Sub
    End If
    ' ==============================================================================
    
    ' 1. 确定范围
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastRow < 10 Then lastRow = 100
    lastCol = ws.Cells(10, ws.Columns.count).End(xlToLeft).Column
    If lastCol < 5 Then lastCol = 20
    
    ' 2. 定义公式模板
    formulaTemplate = "=SUMIFS(" & SHEET_MASTER & "!$F:$F," & _
                      SHEET_MASTER & "!$B:$B,IF($C$1=""合并数"",""*"",$C$1)," & _
                      SHEET_MASTER & "!$C:$C,$A@R," & _
                      SHEET_MASTER & "!$D:$D,$B@R," & _
                      SHEET_MASTER & "!$E:$E,XLOOKUP($A@R&COLUMN()," & SHEET_CONFIG & "!$A:$A&" & SHEET_CONFIG & "!$B:$B," & SHEET_CONFIG & "!$C:$C,0))"
    
    ' 3. 遍历并注入
    For r = 1 To lastRow
        code = Trim(CStr(ws.Cells(r, 1).Value))
        
        ' 仅处理有效的数据行
        If code <> "" And code <> CODE_SENTINEL Then
            currentFormula = Replace(formulaTemplate, "@R", r)
            
            For c = 3 To lastCol
                ' 严格颜色判断：非无填充 且 为白色
                If ws.Cells(r, c).Interior.ColorIndex <> xlNone And ws.Cells(r, c).Interior.Color = vbWhite Then
                    ' 使用 Formula2 以支持动态数组公式
                    ws.Cells(r, c).Formula2 = currentFormula
                End If
            Next c
        End If
    Next r
End Sub

' [修改] 初始化 User_Config 表 (增强健壮性：独立检查 E 列)
' [修改] 初始化 User_Config 表 (调用上面的默认值配置)
Private Function Lib_InitUserConfigSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("User_Config")
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count))
        ws.Name = "User_Config"
    End If
    On Error GoTo 0
    
    ' 1. 初始化基础列 (A-D列)
    If ws.Range("A1").Value = "" Then
        ws.Columns("A:F").NumberFormat = "@"
        ws.Range("A1").Value = "Blacklist (黑名单)"
        ws.Range("B1").Value = "Company_List (公司列表)"
        ws.Range("C1").Value = "Map_Code (特征码)"
        ws.Range("D1").Value = "Map_Account (会计科目)"
        ws.Range("A1:D1").Font.Bold = True
        ws.Range("A1:D1").Interior.Color = RGB(220, 240, 220)
        
        ' 写入默认黑名单... (此处省略，保持原有大量数组写入代码即可)
        Dim defaults As Variant
        defaults = Array("项目", "名称", "种类", "单位名称", "行次", "合计", "期末数", "期初数", "收回或转回", _
                         "成本", "累计确认的信用减值准备", "上期数", "本期数", "期末余额", "期初余额", _
                         "账面余额", "坏账准备", "计提比例（%）", "比例（%）", "账面价值", "金额", "比例", _
                         "本期变动金额", "计提", "收回", "转回", "核销", "其他")
        ws.Range("A2").Resize(UBound(defaults) + 1, 1).Value = Application.Transpose(defaults)
        ws.Range("B2").Value = "合并数"
        ws.Range("C2").Value = "100001"
        ws.Range("D2").Value = "货币资金"
    End If
    
    ' === 第二部分：【核心修改】独立检查并初始化样式配置 (E-F列) ===
    ' 2. 初始化样式配置区 (E列 Key, F列 默认模板)
    If ws.Range("E1").Value = "" Then
        ws.Range("E1").Value = "Style_Key (样式参数)"
        ws.Range("F1").Value = "通用模板"
        ws.Range("E1:F1").Font.Bold = True
        ws.Range("E1:F1").Interior.Color = RGB(220, 240, 220)
    End If
    
    ' 【关键】调用默认样式填充
    Call Add_Default_Style_Configs(ws)
    
    ws.Visible = xlSheetHidden
    Set Lib_InitUserConfigSheet = ws
End Function

' ==========================================================================================
' [重构] 样式加载器 (支持指定模板)
' ==========================================================================================
' 参数 templateName: 如果为空，则加载 F 列（默认/第一个模板），否则查找对应列
Public Sub Load_Table_Style_Config(Optional templateName As String = "")
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim colIndex As Long
    Dim k As String, v As String
    Dim targetCol As Range
    
    Set g_TableConfig = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("User_Config")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    ' 1. 确定要读取哪一列 (模板列)
    colIndex = 6 ' 默认为 F 列 (第6列)
    If templateName <> "" Then
        ' 在第一行查找模板名称
        Dim found As Range
        Set found = ws.Rows(1).Find(What:=templateName, LookIn:=xlValues, LookAt:=xlWhole)
        If Not found Is Nothing Then
            colIndex = found.Column
        End If
    End If
    
    ' 2. 读取数据 (E列为Key, colIndex列为Value)
    lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ' 读取所有配置到字典
    For i = 2 To lastRow
        k = Trim(CStr(ws.Cells(i, "E").Value))
        v = Trim(CStr(ws.Cells(i, colIndex).Value))
        If v = "" Then v = Trim(CStr(ws.Cells(i, 6).Value)) ' 兜底
        
        If k <> "" Then
            If g_TableConfig.exists(k) Then g_TableConfig(k) = v Else g_TableConfig.Add k, v
        End If
    Next i
End Sub

' [新增] 获取所有可用模板名称
Public Function Get_All_Template_Names() As Collection
    Dim ws As Worksheet
    Dim res As New Collection
    Dim col As Long
    Dim val As String
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("User_Config")
    On Error GoTo 0
    If ws Is Nothing Then Set Get_All_Template_Names = res: Exit Function
    
    ' 从 F 列开始向右遍历，直到标题为空
    col = 6
    Do
        val = Trim(CStr(ws.Cells(1, col).Value))
        If val = "" Then Exit Do
        res.Add val
        col = col + 1
    Loop
    
    Set Get_All_Template_Names = res
End Function

' [新增] 保存配置到指定模板列
Public Sub Save_Config_To_Template(templateName As String, k As String, v As String)
    Dim ws As Worksheet
    Dim colIndex As Long
    Dim keyRow As Long
    
    Set ws = ActiveWorkbook.Sheets("User_Config")
    
    ' 1. 找列 (模板)
    Dim foundCol As Range
    Set foundCol = ws.Rows(1).Find(What:=templateName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If foundCol Is Nothing Then
        ' 如果没找到(这是新模板)，则追加新列
        colIndex = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
        ws.Cells(1, colIndex).Value = templateName
        ws.Cells(1, colIndex).Interior.Color = RGB(240, 240, 240) ' 标记颜色
    Else
        colIndex = foundCol.Column
    End If
    
    ' 2. 找行 (Key)
    Dim foundRow As Range
    Set foundRow = ws.Columns("E").Find(What:=k, LookIn:=xlValues, LookAt:=xlWhole)
    
    If foundRow Is Nothing Then
        ' 没找到Key，追加新行
        keyRow = ws.Cells(ws.Rows.count, "E").End(xlUp).Row + 1
        ws.Cells(keyRow, "E").Value = k
    Else
        keyRow = foundRow.Row
    End If
    
    ' 3. 写入值
    ws.Cells(keyRow, colIndex).Value = v
End Sub

' [新增] 工具：同步公司名称到 User_Config 的 B 列
Private Sub Tool_UpdateCompanyList(newComp As String)
    Dim wsUser As Worksheet
    Dim rngFound As Range
    Dim lastRow As Long
    
    If Trim(newComp) = "" Then Exit Sub
    
    On Error Resume Next
    Set wsUser = ActiveWorkbook.Sheets("User_Config")
    On Error GoTo 0
    If wsUser Is Nothing Then Exit Sub
    
    ' 检查 B 列是否已存在该公司
    Set rngFound = wsUser.Columns("B").Find(What:=newComp, LookIn:=xlValues, LookAt:=xlWhole)
    
    If rngFound Is Nothing Then
        lastRow = wsUser.Cells(wsUser.Rows.count, "B").End(xlUp).Row
        wsUser.Cells(lastRow + 1, "B").Value = newComp
    End If
End Sub

' 辅助：复制 Config 规则 (保持 TableTitle 原样)
' ==========================================================================================
' [修正版] 工具：在 Sys_Config 中复制规则 (修复不更新问题)
' ==========================================================================================
Private Sub Tool_DuplicateConfig(srcCode As String, destCode As String)
    Dim wsCfg As Worksheet
    Dim i As Long, lastRow As Long, outRow As Long
    Dim newRows As Collection, item As Variant, rowData As Variant
    Dim cellVal As String, searchVal As String
    Dim copyCount As Integer
    
    ' 1. 获取当前活动工作簿的配置表
    On Error Resume Next
    Set wsCfg = ActiveWorkbook.Sheets("Sys_Config")
    On Error GoTo 0
    
    If wsCfg Is Nothing Then
        MsgBox "错误：在当前文件中找不到 [Sys_Config] 表，无法更新配置。", vbCritical
        Exit Sub
    End If
    
    Set newRows = New Collection
    lastRow = wsCfg.Cells(wsCfg.Rows.count, "A").End(xlUp).Row
    
    ' 2. 预处理查找值 (去除空格，强制转文本)
    searchVal = Trim(CStr(srcCode))
    copyCount = 0
    
    ' 3. 遍历查找源规则
    ' 注意：这里使用 Trim(CStr(...)) 确保数字 10001 能匹配文本 "10001"
    For i = 2 To lastRow
        cellVal = Trim(CStr(wsCfg.Cells(i, 1).Value))
        
        If cellVal = searchVal Then
            ReDim rowData(1 To 4)
            rowData(1) = destCode                ' 新代码 (例如 10001_01)
            rowData(2) = wsCfg.Cells(i, 2).Value ' 列号
            rowData(3) = wsCfg.Cells(i, 3).Value ' 指标名
            rowData(4) = wsCfg.Cells(i, 4).Value ' 表格标题
            newRows.Add rowData
            copyCount = copyCount + 1
        End If
    Next i
    
    ' 4. 关键调试：如果没找到源规则，说明匹配失败
    If copyCount = 0 Then
        ' 可以在这里取消注释进行调试，看看到底是哪个代码没找到
        ' Debug.Print "未找到源规则: " & searchVal
        ' MsgBox "警告：无法在配置表中找到代码 [" & searchVal & "] 的规则，因此无法生成 [" & destCode & "] 的配置。", vbExclamation
        Exit Sub
    End If
    
    ' 5. 写入新规则到表底
    outRow = lastRow + 1
    For Each item In newRows
        wsCfg.Cells(outRow, 1).Value = item(1)
        wsCfg.Cells(outRow, 2).Value = item(2)
        wsCfg.Cells(outRow, 3).Value = item(3)
        wsCfg.Cells(outRow, 4).Value = item(4)
        outRow = outRow + 1
    Next item
    
    ' 强制保存一下状态，防止内存未刷新
    ' DoEvents
End Sub

' =================================================================
' 1. 初始化 Sys_Config 表 (增加开关配置项)
' =================================================================
Public Function Lib_InitConfigSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(SHEET_CONFIG)
    
    ' 如果不存在则新建
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count))
        ws.Name = SHEET_CONFIG
    End If
    On Error GoTo 0
    
    ws.Columns("A:F").Clear
        
    ' --- A-D 列：原有的取数规则区域 (保持不变) ---
    If ws.Range("A1").Value = "" Then
        ws.Columns("A").NumberFormat = "@"
        ws.Range("A1:D1").Value = Array("ModuleCode", "ColIndex", "MetricName", "TableTitle")
    End If
    
    ' --- [新增] G-H 列：系统功能开关区域 ---
    ' 使用 G/H 列是为了避开左侧的数据规则，便于扩展
    If ws.Range("G1").Value = "" Then
        ws.Range("G1").Value = "System_Switch (系统开关)"
        ws.Range("H1").Value = "Status"
        ws.Range("G1:H1").Font.Bold = True
        ws.Range("G1:H1").Interior.Color = RGB(220, 220, 240)
        
        ' 写入默认值：默认关闭 (FALSE)
        ws.Range("G2").Value = "Link_Word_Excel"
        ws.Range("H2").Value = "FALSE"
    End If
    
    ' 设置为深度隐藏，避免用户误删
    ws.Visible = xlSheetHidden
    
    Set Lib_InitConfigSheet = ws
End Function

' =================================================================
' 2. [新增] 读取功能开关状态
' =================================================================
Public Function Get_Sys_Switch_Status(switchKey As String) As Boolean
    Dim ws As Worksheet
    Dim rngFound As Range
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0
    
    ' 如果没有配置表，默认返回 FALSE (安全策略)
    If ws Is Nothing Then
        Get_Sys_Switch_Status = False
        Exit Function
    End If
    
    ' 在 G 列查找开关名称
    Set rngFound = ws.Columns("G").Find(What:=switchKey, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rngFound Is Nothing Then
        ' 读取 H 列对应的值
        Dim val As String
        val = UCase(Trim(CStr(rngFound.Offset(0, 1).Value)))
        If val = "TRUE" Then
            Get_Sys_Switch_Status = True
        Else
            Get_Sys_Switch_Status = False
        End If
    Else
        ' 找不到该键值，默认为 False
        Get_Sys_Switch_Status = False
    End If
End Function

' =================================================================
' 3. [修改] 写入功能开关状态 (安全写入版)
' =================================================================
Public Sub Set_Sys_Switch_Status(switchKey As String, status As Boolean)
    Dim ws As Worksheet
    Dim rngFound As Range
    Dim nextRow As Long
    
    ' 1. 【核心修改】不要调用 Lib_InitConfigSheet，而是直接获取
    '    防止每次写入开关都触发清空 A-F 列的逻辑
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("Sys_Config")
    On Error GoTo 0
    
    ' 如果表不存在，说明用户还没生成过模板，此时可以调用初始化或者提示错误
    If ws Is Nothing Then
        ' 方案 A: 自动初始化 (如果此时初始化，A-F列本来就是空的，也无所谓)
        ' Set ws = Lib_InitConfigSheet()
        
        ' 方案 B: 提示用户 (推荐，逻辑更严谨)
        MsgBox "未找到配置表 [Sys_Config]，请先运行【生成模板】。", vbExclamation
        Exit Sub
    End If
    
    ' 2. 确保表头存在 (防止手动删除了 G/H 列表头)
    If ws.Range("G1").Value <> "System_Switch (系统开关)" Then
        ws.Range("G1").Value = "System_Switch (系统开关)"
        ws.Range("H1").Value = "Status"
        ws.Range("G1:H1").Font.Bold = True
        ws.Range("G1:H1").Interior.Color = RGB(220, 230, 241)
    End If
    
    ' 3. 查找键值 (switchKey)
    '    注意：增加 LookIn:=xlValues, LookAt:=xlWhole 确保精确匹配
    Set rngFound = ws.Columns("G").Find(What:=switchKey, LookIn:=xlValues, LookAt:=xlWhole)
    
    If rngFound Is Nothing Then
        ' 如果没找到，追加到 G 列末尾
        ' 注意：如果表是空的（只有标题），End(xlUp) 会返回 1，所以 nextRow = 2
        nextRow = ws.Cells(ws.Rows.count, "G").End(xlUp).Row + 1
        If nextRow < 2 Then nextRow = 2 ' 双重保险
        
        ws.Cells(nextRow, "G").Value = switchKey
        Set rngFound = ws.Cells(nextRow, "G")
    End If
    
    ' 4. 写入状态到 H 列
    '    使用 UCase 统一转为大写 TRUE/FALSE 字符串，方便后续读取
    rngFound.Offset(0, 1).Value = UCase(CStr(status))
    
    ' 5. 简单美化
    ws.Columns("G:H").AutoFit
End Sub

' 必须在 ActiveWorkbook (用户文件) 中创建/查找，而不是 ThisWorkbook (插件)
Private Function Lib_GetMasterSheet() As Worksheet
    Dim ws As Worksheet
    Dim wbTarget As Workbook
    
    ' 1. 锁定目标工作簿：必须是用户当前激活的窗口
    Set wbTarget = Application.ActiveWorkbook
    
    ' 2. 防御性检查：如果没有打开任何工作簿
    If wbTarget Is Nothing Then
        MsgBox "错误：未检测到活动工作簿！" & vbCrLf & "请先打开或新建一个 Excel 文件用于存放汇总数据。", vbCritical
        End ' 强制终止
    End If
    
    ' 3. 在用户文件中查找 Sheet
    On Error Resume Next
    Set ws = wbTarget.Sheets(SHEET_MASTER)
    On Error GoTo 0
    
    ' 4. 如果不存在，则新建
    If ws Is Nothing Then
        On Error Resume Next
        ' 尝试新建
        Set ws = wbTarget.Sheets.Add
        ws.Name = SHEET_MASTER
        ' 初始化表头
        ws.Range("A1:H1").Value = Array("SourceFile", "Company", "Code", "ItemName", "Metric", "Value", "Time", "Account")
        On Error GoTo 0
    End If
    
    ' 5. 最终检查 (防止工作簿保护导致新建失败)
    If ws Is Nothing Then
        MsgBox "无法初始化 [" & SHEET_MASTER & "] 表！" & vbCrLf & _
               "可能原因：工作簿结构被保护，无法新建工作表。", vbCritical
        End
    End If
    
    ' 确保表头完善
    If ws.Cells(1, 8).Value = "" Then ws.Cells(1, 8).Value = "Account"
    
    ws.Visible = xlSheetHidden
    Set Lib_GetMasterSheet = ws
End Function
' [修改] 工具：基于颜色/内容状态锁定单元格
' [修改] 工具：基于颜色/内容状态锁定单元格
Private Sub Tool_Protect_Data_By_Color(ws As Worksheet)
    
    Dim cell As Range
    Dim rngToLock As Range      ' 收集需要锁定的区域
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' 1. 取消保护，确保可以修改 Locked 属性
    ' 使用常量密码
    On Error Resume Next
    ws.Unprotect LOCK_PASSWORD
    On Error GoTo 0
    
    ' 2. 初始化：解锁整个工作表
    ' 这一步必须在工作表处于非保护状态下才能执行
    ws.Cells.Locked = False
    
    ' 3. 确定遍历范围 (使用已用区域)
    On Error Resume Next
    ' 如果单元格数量小于2（通常只有表头或为空），则退出
    If ws.UsedRange.Cells.count < 2 Then Exit Sub
    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.count).Row
    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.count).Column
    On Error GoTo 0

    ' 4. 遍历并锁定满足条件的单元格
    For Each cell In ws.Range("A1", ws.Cells(lastRow, lastCol)).Cells
        
        ' 【核心修改逻辑】
        ' 检查颜色条件：如果单元格的背景色【不等于】纯白色 (vbWhite)
        ' 这包含了三种情况：
        ' 1. 单元格无填充色 (ColorIndex = xlNone)
        ' 2. 单元格有其他颜色 (例如灰色、蓝色等)
        ' 3. 单元格有白色填充，但不是 vbWhite（极少见）
        
        ' 注意：如果单元格无填充，它的 .Interior.Color 属性返回 16777215 (即 vbWhite 的颜色值)，
        ' 但它的 .Interior.ColorIndex 返回 xlNone (-4142)。
        ' 必须使用组合判断来区分 "无填充" 和 "纯白色填充"。
        
        If cell.Interior.ColorIndex = xlNone Or cell.Interior.Color <> vbWhite Then
            
            ' 收集需要锁定的单元格
            If rngToLock Is Nothing Then
                Set rngToLock = cell
            Else
                Set rngToLock = Union(rngToLock, cell)
            End If
            
        End If
    Next cell
    
    ' 5. 锁定收集到的区域
    If Not rngToLock Is Nothing Then
        rngToLock.Locked = True
        
        ' 设置锁定区域的背景色为浅灰色，以示区分
        With rngToLock.Interior
            .Pattern = xlSolid
            .Color = LOCK_COLOR ' LOCK_COLOR = 12632256 (浅灰色)
        End With
    End If
    
    ' 6. 使用密码保护工作表
    ws.Protect Password:=LOCK_PASSWORD, _
               Contents:=True, _
               Scenarios:=True, _
               AllowFormattingCells:=True ' 允许用户格式化已解锁的单元格

End Sub
' [新增] 工具：格式化填充色为白色的单元格
Private Sub Tool_Format_WhiteCells(ws As Worksheet)
    
    Dim cell As Range
    Dim rngToFormat As Range    ' 收集需要格式化的区域
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' 1. 确定遍历范围 (使用已用区域)
    On Error Resume Next
    ' 如果单元格数量小于2（通常只有表头或为空），则退出
    If ws.UsedRange.Cells.count < 2 Then Exit Sub
    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.count).Row
    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.count).Column
    On Error GoTo 0

    ' 2. 遍历并收集满足条件的单元格
    For Each cell In ws.Range("A1", ws.Cells(lastRow, lastCol)).Cells
        
        ' 【核心逻辑】只选择显式设置为纯白色填充的单元格
        ' 注意：必须排除 xlNone (无填充)，因为无填充的 .Color 属性也可能返回 vbWhite
        If cell.Interior.ColorIndex <> xlNone And cell.Interior.Color = vbWhite Then
            
            ' 收集需要格式化的单元格
            If rngToFormat Is Nothing Then
                Set rngToFormat = cell
            Else
                Set rngToFormat = Union(rngToFormat, cell)
            End If
            
        End If
    Next cell
    
    ' 3. 应用格式化
    If Not rngToFormat Is Nothing Then
        
        With rngToFormat
            ' 1. 对齐方式：右对齐
            .HorizontalAlignment = xlRight
            
            ' 2. 数值格式与千分位
            ' "#,##0.00" 代表两位小数，带千分位分隔符，零显示为 0.00
            .NumberFormat = "#,##0.00"
            
            ' 3. 可选：避免小数位被格式化
            ' .NumberFormat = "#,##0"
        End With
    End If

End Sub

Private Function Lib_CheckRightBorder(rng As Range) As Boolean
    On Error Resume Next
    If rng.MergeCells Then Lib_CheckRightBorder = (rng.MergeArea.Borders(xlEdgeRight).LineStyle <> xlNone) Else Lib_CheckRightBorder = (rng.Borders(xlEdgeRight).LineStyle <> xlNone)
    On Error GoTo 0
End Function

Private Function Lib_FindTableTitle(ws As Worksheet, currRow As Long) As String
    Dim r As Long, val As String
    For r = currRow - 1 To currRow - 10 Step -1
        If r < 1 Then Exit For
        val = Trim(ws.Cells(r, COL_BORDER_CHK).Value)
        If val <> "" And Len(val) < 50 Then Lib_FindTableTitle = val: Exit Function
    Next r
    Lib_FindTableTitle = "Table_" & currRow
End Function

' ==========================================================================================
' [新增] 配置管理器 (单例模式核心)
' ==========================================================================================

' 1. 统一初始化入口 (只读一次)
' 说明：在任何需要用到配置的地方调用此函数。它会自动判断是否需要读取 Excel。
Sub Init_Global_Config()
    ' 如果已经加载过，先强制清空，防止叠加
    ' (这是为了解决您遇到的诡异重复问题而做的防御性编程)
    If g_IsConfigLoaded Then
        Set Pub_ConfigDict = Nothing
        Set Pub_BlacklistDict = Nothing
        Set Pub_AccountDict = Nothing
    End If
    
    ' --- 开始读取 (只有第一次或重置后会执行到这里) ---
    
    ' 1. 加载 Sys_Config 和 User_Config (黑名单/科目)
    ' 直接调用原有的加载逻辑 (复用现有代码)
    Call P2_LoadConfigToMemory_Public
    
    ' 2. 加载表格样式 (User_Config 的 E/F 列)
    ' 这是一个新函数，下面会定义
    Call Load_Table_Style_Config
    
    ' 3. 设置标志位
    g_IsConfigLoaded = True
End Sub

' 2. 强制重置 (用于保存设置后刷新)
' 说明：当用户在窗体修改了配置并保存时，调用此函数清空缓存。
Sub Force_Reload_Config()
    Set Pub_ConfigDict = Nothing
    Set Pub_BlacklistDict = Nothing
    Set Pub_AccountDict = Nothing
    Set g_TableConfig = Nothing
    
    g_IsConfigLoaded = False
    
    ' 立即重新加载
    Call Init_Global_Config
End Sub

' [修改] 试用期检测 (仅提醒，不拦截)
Public Function CheckTrialVersion() As Boolean
    Dim expireDate As Date
    Dim currentDate As Date
    
    ' 1. 设定截止日期
    expireDate = DateSerial(2026, 1, 31)
    currentDate = Date
    
    ' 2. 判断是否过期
    If currentDate > expireDate Then
        ' === 变化点 A：修改提示措辞和图标 ===
        ' 使用 vbExclamation (黄色感叹号) 代替 vbCritical (红色错误叉)
        MsgBox "【温馨提示：试用期提醒】" & vbCrLf & vbCrLf & _
               "当前日期已超过试用截止日 (" & Format(expireDate, "yyyy-mm-dd") & ")。" & vbCrLf & vbCrLf & _
               "您可以继续使用本工具的所有功能。" & vbCrLf & _
               "为了获得更好的技术支持，建议您联系开发人员获取正式版本。", _
               vbExclamation, "版本提示"
        
        ' === 变化点 B：关键修改 ===
        ' 即使过期了，也返回 True (代表“放行”)
        ' 这样 Ribbon 按钮处的代码就不会执行 Exit Sub 了
        CheckTrialVersion = True
    Else
        ' 未过期，通过
        CheckTrialVersion = True
    End If
End Function
' 检查文件是否被占用
Function IsFileOpen(filename As String) As Boolean
    Dim filenum As Integer, errnum As Integer
    On Error Resume Next
    filenum = FreeFile()
    ' 尝试以独占方式打开文件
    Open filename For Input Lock Read As #filenum
    Close filenum
    errnum = Err.Number
    On Error GoTo 0
    
    ' 错误号 70 表示权限被拒绝（通常是因为文件已打开）
    If errnum = 70 Then IsFileOpen = True Else IsFileOpen = False
End Function


' ==========================================================================================
' [新增] 数据管理模块 (Data Management)
' ==========================================================================================

' 1. 获取 Master 表中所有已导入的源文件名 (去重)
Public Function DM_GetImportedFileNames() As Collection
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim dict As Object
    Dim val As String
    Dim res As New Collection
    
    ' 获取 Master 表
    Set ws = Lib_GetMasterSheet()
    If ws Is Nothing Then Set DM_GetImportedFileNames = res: Exit Function
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 遍历 A 列 (SourceFile)
    If lastRow >= 2 Then
        Dim arr As Variant
        ' 数组读取加速
        arr = ws.Range("A2:A" & lastRow).Value
        
        For i = 1 To UBound(arr, 1)
            val = Trim(CStr(arr(i, 1)))
            If val <> "" Then
                If Not dict.exists(val) Then
                    dict.Add val, True
                    res.Add val
                End If
            End If
        Next i
    End If
    
    Set DM_GetImportedFileNames = res
End Function

' 2. 根据文件名列表批量删除数据
' [修改版] 根据文件名列表批量删除数据 (同步清理 Master 和 User_Config)
' 位置：Mod_Financial_Core 模块
Public Sub DM_DeleteDataByFiles(filesToDelete As Collection)
    Dim wsMaster As Worksheet, wsUser As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim fName As Variant
    Dim arrCriteria() As String
    Dim i As Long
    
    ' 新增变量：用于处理 User_Config
    Dim compName As String
    Dim rngFound As Range
    
    If filesToDelete.count = 0 Then Exit Sub
    
    Set wsMaster = Lib_GetMasterSheet()
    If wsMaster Is Nothing Then Exit Sub
    
    ' 获取 User_Config 表
    On Error Resume Next
    Set wsUser = ActiveWorkbook.Sheets("User_Config")
    On Error GoTo 0
    
    ' ==============================================================================
    ' 第一步：同步删除 User_Config 中的公司列表 (B列)
    ' ==============================================================================
    If Not wsUser Is Nothing Then
        Application.StatusBar = "正在清理公司列表..."
        
        For Each fName In filesToDelete
            ' 1. 根据文件名还原公司名称 (逻辑与导入时保持一致)
            '    假设文件名是 "ABC公司.xlsx"，则公司名是 "ABC公司"
            If InStr(fName, ".") > 0 Then
                compName = Left(fName, InStrRev(fName, ".") - 1)
            Else
                compName = fName
            End If
            
            ' 2. 在 B 列查找该公司
            '    注意：LookAt:=xlWhole 确保全字匹配，防止误删类似名称
            Set rngFound = wsUser.Columns("B").Find(What:=compName, LookIn:=xlValues, LookAt:=xlWhole)
            
            ' 3. 如果找到且不是表头，则删除
            If Not rngFound Is Nothing Then
                If rngFound.Row > 1 Then ' 保护表头
                    ' 删除单元格并让下方数据上移，保持列表连续性
                    rngFound.Delete Shift:=xlUp
                End If
            End If
        Next fName
    End If
    
    ' ==============================================================================
    ' 第二步：删除 Master 表中的数据 (原有逻辑保持不变)
    ' ==============================================================================
    Application.StatusBar = "正在删除 Master 数据..."
    Application.ScreenUpdating = False
    
    ' 转换 Collection 为数组，用于 AutoFilter
    ReDim arrCriteria(0 To filesToDelete.count - 1)
    i = 0
    For Each fName In filesToDelete
        arrCriteria(i) = CStr(fName)
        i = i + 1
    Next fName
    
    ' 移除旧筛选
    If wsMaster.AutoFilterMode Then wsMaster.AutoFilterMode = False
    
    lastRow = wsMaster.Cells(wsMaster.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then GoTo ExitHandler
    
    ' 执行筛选删除 (Column 1 = SourceFile)
    Set rng = wsMaster.Range("A1:H" & lastRow)
    
    ' 筛选出要删除的文件名
    rng.AutoFilter Field:=1, Criteria1:=arrCriteria, Operator:=xlFilterValues
    
    ' 删除可见行 (跳过标题行)
    On Error Resume Next
    Dim rngToDelete As Range
    Set rngToDelete = rng.Offset(1, 0).Resize(lastRow - 1).SpecialCells(xlCellTypeVisible)
    
    If Not rngToDelete Is Nothing Then
        rngToDelete.EntireRow.Delete
    End If
    On Error GoTo 0
    
    ' 清除筛选
    If wsMaster.AutoFilterMode Then wsMaster.ShowAllData
    
ExitHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub
' user-config辅助写入函数
Private Sub WriteDefault(ws As Worksheet, k As String, v As String)
    Dim rngFound As Range
    ' 在E列查找Key
    Set rngFound = ws.Columns("E").Find(What:=k, LookIn:=xlValues, LookAt:=xlWhole)
    If rngFound Is Nothing Then
        Dim nextRow As Long
        nextRow = ws.Cells(ws.Rows.count, "E").End(xlUp).Row + 1
        ws.Cells(nextRow, "E").Value = k
        ws.Cells(nextRow, "F").Value = v ' 默认写入到 F 列(通用模板)
    End If
End Sub
