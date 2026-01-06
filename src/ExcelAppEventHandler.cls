VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelAppEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =================================================================
' 类模块: ExcelAppEventHandler
' 作用: 监听 Excel 全局双击事件，实现 Excel -> Word 跳转
' =================================================================
Option Explicit

' 声明带事件的 Excel 应用程序对象
Public WithEvents xlApp As Excel.Application
Attribute xlApp.VB_VarHelpID = -1

' Excel 双击事件 (全局监听)
Private Sub xlApp_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim intersectDict As Object
    Dim tableName As String
    Dim wdApp As Object, wdDoc As Object
    
    On Error GoTo ExitHandler
    
    ' [过滤] 只在 "附注模板" 工作表中生效
    If Sh.Name <> "附注模板" Then Exit Sub
    
    ' 1. 判断双击位置是否在已知表格内
    Set intersectDict = Mod_Excel2Word.GetIntersectDictionary(Target)
    If intersectDict.count = 0 Then Exit Sub
    
    ' 2. 阻止 Excel 进入编辑模式
    Cancel = True
    
    ' 3. 获取表格名称
    tableName = intersectDict.Keys()(0)
    
    ' 4. 连接 Word/WPS (复用模块中的函数)
    If Not Mod_Excel2Word.InitWordApp(wdApp, wdDoc, CreateNew:=False) Then
        MsgBox "未找到活动的 Word/WPS 文档。", vbExclamation
        Exit Sub
    End If
    
    ' 5. 检查 Word 中是否存在该书签
    If wdDoc.Bookmarks.exists(tableName) Then
        ' 6. 强力激活Word/WPS窗口 (修复最小化或无法弹出的问题)
        ' 传入 wdApp (Word应用程序对象)
        Call Mod_Excel2Word.ForceActivateWindow(wdApp)
        wdDoc.Bookmarks(tableName).Range.Select
    Else
        MsgBox "Word文档中未找到书签: " & tableName, vbExclamation
    End If

ExitHandler:
    Set wdApp = Nothing
    Set wdDoc = Nothing
    Set intersectDict = Nothing
End Sub
