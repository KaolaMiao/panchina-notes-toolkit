VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' 1. 打开加载项时
Private Sub Workbook_Open()
    ' 直接运行业务逻辑
    ' 如果引用丢失，早在编译阶段就报错了，这里的检查代码也救不回来
    If Mod_Financial_Core.Get_Sys_Switch_Status("Link_Word_Excel") Then
        Call Mod_Excel2Word.StartAppEventListeners
    End If
End Sub

' 2. 安装加载项时
Private Sub Workbook_AddinInstall()
    ' 保持静默，无操作
End Sub

' 3. 关闭/卸载时
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Call Mod_Excel2Word.StopAppEventListeners
End Sub

Private Sub Workbook_AddinUninstall()
    ' Call Mod_Excel2Word.StopAppEventListeners
End Sub

' 【已删除 EnsureWordReference 函数】

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' 错误处理：防止导出代码出错导致 Excel 无法保存
    On Error GoTo ErrorHandler
    
    ' 调用刚才写的导出模块
    Call Mod_GitExporter.ExportSourceFiles
    
    Exit Sub

ErrorHandler:
    ' 如果导出出错（比如文件被占用），不阻拦保存，但提示错误
    MsgBox "代码导出失败，但文件将继续保存。" & vbCrLf & _
           "错误: " & Err.Description, vbExclamation, "Git Export Error"
    Resume Next
End Sub
