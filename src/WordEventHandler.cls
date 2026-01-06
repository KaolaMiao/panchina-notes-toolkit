VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =================================================================
' 类模块: WordEventHandler
' 注意：由于使用了 WithEvents，本模块必须保留对 Word Object Library 的引用
' 请务必在安装了旧版 Office (2007/2010) 的电脑上保存本文件以确保兼容性
' =================================================================
Option Explicit

' 【必须保留 Early Binding】 否则无法监听事件
Public WithEvents appInstance As Word.Application
Attribute appInstance.VB_VarHelpID = -1

' 事件参数类型 (ByVal Sel As Word.Selection) 也是由接口定义的，不能改为 Object，否则事件不触发
Private Sub appInstance_WindowBeforeDoubleClick(ByVal Sel As Word.Selection, Cancel As Boolean)
    ' 内部变量改为通用 Object，增加一点点兼容弹性
    Dim bm As Object
    Dim tableName As String
    
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    
    On Error GoTo ExitHandler
    
    ' 1. 判断表格 (使用数字常量 12 代替 wdWithInTable，防止常量丢失)
    ' Sel.Information(12) 是通用的
    If Sel.Information(12) Then
        
        ' 2. 查找书签
        ' 遍历集合改为通用方式
        For Each bm In Sel.Document.Bookmarks
            If Sel.InRange(bm.Range) Then
                tableName = bm.Name
                Exit For
            End If
        Next bm
        
        ' 3. 跳转回 Excel
        If tableName <> "" Then
            Cancel = True
            
            ' A. 激活 Excel
            Call Mod_Excel2Word.ForceActivateWindow(Application)
            
            ' B. 查找工作簿
            Set targetWb = Mod_Excel2Word.FindTargetWorkbook()
            
            If targetWb Is Nothing Then
                MsgBox "跳转失败！" & vbCrLf & _
                       "未找到包含 '附注模板' 的数据文件。", vbExclamation
                Exit Sub
            End If
            
            ' C. 激活工作簿
            targetWb.Activate
            On Error Resume Next
            Set targetWs = targetWb.Sheets("附注模板")
            On Error GoTo ExitHandler
            
            If targetWs Is Nothing Then Exit Sub
            
            ' D. 缓存处理
            If Mod_Excel2Word.g_tableRanges Is Nothing Then
                 Call Mod_Excel2Word.PreCacheTableRanges(targetWs)
            End If
            
            ' E. 执行跳转
            If Mod_Excel2Word.g_tableRanges.exists(tableName) Then
                Application.GoTo Mod_Excel2Word.g_tableRanges(tableName), True
            Else
                ' 二次尝试
                Call Mod_Excel2Word.PreCacheTableRanges(targetWs)
                If Mod_Excel2Word.g_tableRanges.exists(tableName) Then
                    Application.GoTo Mod_Excel2Word.g_tableRanges(tableName), True
                Else
                    MsgBox "未找到关联表格: " & tableName, vbExclamation
                End If
            End If
            
        End If
    End If

ExitHandler:
    Set bm = Nothing
    Set targetWb = Nothing
    Set targetWs = Nothing
End Sub

