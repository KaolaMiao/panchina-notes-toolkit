Attribute VB_Name = "Mod_GitExporter"
' 模块: GitExporter
Option Explicit

Sub ExportSourceFiles()
    Dim component As Object
    Dim path As String
    Dim extension As String
    Dim fso As Object
    
    ' 1. 检查是否已保存
    ' 如果文件从未保存过，就没有路径，无法导出
    If ThisWorkbook.path = "" Then Exit Sub
    
    ' 2. 定义导出路径
    ' 默认导出到 Excel 文件同级目录下的 src 文件夹中
    path = ThisWorkbook.path & "\src\"
    
    ' 3. 创建文件夹（如果不存在）
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
    
    ' 4. 遍历并导出所有代码模块
    For Each component In ThisWorkbook.VBProject.VBComponents
        Select Case component.Type
            Case 1 ' 标准模块 (.bas)
                extension = ".bas"
            Case 2 ' 类模块 (.cls)
                extension = ".cls"
            Case 3 ' 窗体 (.frm)
                extension = ".frm"
            Case 100 ' 工作表和 ThisWorkbook 代码 (.cls)
                extension = ".cls"
            Case Else
                extension = ".txt"
        End Select
        
        ' 导出组件
        ' 注意：如果有同名文件，Export 方法会自动覆盖
        On Error Resume Next ' 忽略导出空模块可能产生的错误
        component.Export path & component.Name & extension
        On Error GoTo 0
    Next component
    
    ' (可选) 在状态栏提示导出成功，而不是弹出烦人的对话框
    Application.StatusBar = "Git同步准备：VBA代码已成功导出至 src 文件夹 - " & Now
End Sub
