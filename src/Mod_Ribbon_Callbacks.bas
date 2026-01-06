Attribute VB_Name = "Mod_Ribbon_Callbacks"
' =================================================================
' Copyright (C) 2025 Dota (PCCPA). All Rights Reserved.
' 本程序受著作权法和国际条约保护。
' 未经授权的复制或分发本程序（或其中任何部分），将导致严厉的民事和刑事处罚。
' =================================================================
' =================================================================
' 模块名称: 财务附注ETL工具
' 作者：Dota | 天健会计师事务所 (PCCPA)
' =================================================================
Option Explicit

' ==========================================================================================
' MODULE:       Mod_Ribbon_Callbacks
' DESCRIPTION:  Ribbon 功能区按钮的回调函数汇总 (接口层)
' ==========================================================================================

' 1. [大按钮] 启动汇总控制台
' XML: onAction="Ribbon_LaunchDashboard"
Sub Ribbon_LaunchDashboard(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ' 呼叫窗体
    If Not Mod_Financial_Core.CheckTrialVersion() Then Exit Sub
    'Frm_Main.Show vbModeless
    Frm_Main.Show
    Exit Sub
ErrorHandler:
    MsgBox "启动控制台失败：" & Err.Description, vbCritical
End Sub

' 2. [小按钮] 快速生成模板
' XML: onAction="Ribbon_GenerateTemplate"
Sub Ribbon_GenerateTemplate(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ' 调用核心模块的生成逻辑
    Call Mod_Financial_Core.Main_GenerateTemplate
    Exit Sub
ErrorHandler:
    MsgBox "执行失败：" & Err.Description, vbCritical
End Sub

' 3. [小按钮] 快速注入公式
' XML: onAction="Ribbon_InjectFormula"
Sub Ribbon_InjectFormula(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ' 调用核心模块的公式注入逻辑
    Call Mod_Financial_Core.Main_InjectFormulas
    Exit Sub
ErrorHandler:
    MsgBox "执行失败：" & Err.Description, vbCritical
End Sub

' 4. [小按钮] 手动拆分复杂表
' XML: onAction="Ribbon_ManualSplit"
Sub Ribbon_ManualSplit(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ' 调用核心模块的手动拆分逻辑
    Call Mod_Financial_Core.Main_ManualSplit
    Exit Sub
ErrorHandler:
    MsgBox "执行失败：" & Err.Description, vbCritical
End Sub

' ==========================================================================================
' 5. [小按钮] 保护非公式数据
' XML: onAction="Ribbon_ProtectData"
Sub Ribbon_ProtectData(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ' 调用核心模块的数据保护逻辑
    Call Mod_Financial_Core.Main_ProtectNonFormulaData
    Exit Sub
ErrorHandler:
    MsgBox "执行失败：" & Err.Description, vbCritical
End Sub

' 6. [小按钮] 一键解锁
' XML: onAction="Ribbon_UnprotectData"
Sub Ribbon_UnprotectData(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ' 调用核心模块的解锁逻辑
    Call Mod_Financial_Core.Main_Unprotect_Data
    Exit Sub
ErrorHandler:
    MsgBox "执行失败：" & Err.Description, vbCritical
End Sub
' 7. [小按钮] 格式化白色区域
' XML: onAction="Ribbon_FormatWhiteCells"
Sub Ribbon_FormatWhiteCells(control As IRibbonControl)
    On Error GoTo ErrorHandler
    ' 调用核心模块的格式化逻辑
    Call Mod_Financial_Core.Main_Format_WhiteCells
    Exit Sub
ErrorHandler:
    MsgBox "执行失败：" & Err.Description, vbCritical
End Sub
' ==========================================================================================
' [新增] 8. 启动动态数据核对
' XML: onAction="Ribbon_RunReconciliation"
' ==========================================================================================
Sub Ribbon_RunReconciliation(control As IRibbonControl)
    On Error GoTo ErrorHandler

    ' 确保你的模块中有 Run_Data_Reconciliation 这个 Sub
    Call Run_Data_Reconciliation
    
    Exit Sub
ErrorHandler:
    MsgBox "启动校验失败：" & vbCrLf & "错误信息: " & Err.Description & vbCrLf & _
           "请检查是否已将 'Mod_Data_Validation_V19' 模块导入工程。", vbCritical, "调用错误"
End Sub

