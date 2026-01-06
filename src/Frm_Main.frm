VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Main 
   Caption         =   "财务自动化ETL智能控制台Ver 2.1(Beta 20251209)"
   ClientHeight    =   8580
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11928
   OleObjectBlob   =   "Frm_Main.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Label8_Click()

End Sub

' ==========================================================================================
' USERFORM:     Frm_Main
' DESCRIPTION:  财务ETL系统综合控制台
' ==========================================================================================
' ==========================================================================================
' 1. 初始化与窗口事件
' ==========================================================================================

Private Sub UserForm_Initialize()
    ' 1. 界面初始化
    Me.MultiPage1.Value = 0 ' 默认显示第1页
    Me.lst_Target.ColumnWidths = "500" ' 设置一个足够大的宽度
    
    ' 2. 进度条归零
    Me.lbl_Prog_Bar.Width = 0
    Me.lbl_Status.Caption = "等待添加文件..."
    
    ' 3. 加载黑名单
    LoadBlacklistToUI
    
    Me.cboStyleAlign.List = Array("左对齐", "居中", "右对齐", "两端对齐", "不调整")
    
    Load_Style_To_UI
    
    
    ' 初始化样式对象下拉框
    Me.cbo_Style_Target.Clear
    Me.cbo_Style_Target.AddItem "表格 (Table)"
    Me.cbo_Style_Target.AddItem "一级标题 (H1)"
    Me.cbo_Style_Target.AddItem "二级标题 (H2)"
    Me.cbo_Style_Target.AddItem "正文 (P)"
    Me.cbo_Style_Target.ListIndex = 1
    

    
    ' 【新增】加载模板列表
    Call Refresh_Template_List
    ' 初始化数据管理列表（可选，如果用户可能直接看到该页）
    Call RefreshDataManageList

End Sub

' 1. 选项卡切换事件：点到“数据管理”时自动刷新
Private Sub MultiPage1_Change()
    ' 假设数据管理是第3页 (索引为2)，请根据实际情况调整 Index
    If Me.MultiPage1.SelectedItem.Caption = "数据管理" Then
        Call RefreshDataManageList
    End If
End Sub

' ==========================================================================================
' 2. [Tab 1] 文件选择与穿梭框逻辑
' ==========================================================================================

' [修复版] 按钮：选择源文件夹 (使用 Shell 方法)
Private Sub btn_BrowseFolder_Click()
    Dim folderPath As String
    Dim filename As String
    Dim count As Long
    
    ' 1. 调用系统选择器获取路径
    folderPath = GetFolder_Shell()
    
    ' 2. 判断用户是否取消
    If folderPath = "" Then
        ' 用户点了取消，或者没选文件夹，直接退出，不弹窗打扰
        Exit Sub
    End If
    
    ' 3. 显示路径
    Me.txtFolderPath.Text = folderPath
    
    ' 4. 清空左侧列表
    Me.lst_Source.Clear
    
    ' 5. 遍历文件
    ' 使用通配符 *.xls* 匹配所有 Excel 版本
    filename = Dir(folderPath & "*.xls*")
    
    If filename = "" Then
        MsgBox "该文件夹下没有找到 Excel 文件。", vbExclamation
        Exit Sub
    End If
    
    count = 0
    Do While filename <> ""
        ' 防重复添加逻辑
        If Not IsInList(Me.lst_Target, folderPath & filename) Then
            Me.lst_Source.AddItem filename
            count = count + 1
        End If
        filename = Dir
    Loop
    
    UpdateCounts
    
    ' 可选：提示一下添加了多少个
    ' MsgBox "已加载 " & count & " 个文件。", vbInformation
End Sub

' 2. 按钮：添加选中 ( > )
Private Sub btn_Add_Click()
    MoveItems Me.lst_Source, Me.lst_Target, True
End Sub

' 3. 按钮：添加全部 ( >> )
Private Sub btn_AddAll_Click()
    MoveAllItems Me.lst_Source, Me.lst_Target, True
End Sub

' 4. 按钮：移除选中 ( < )
Private Sub btn_Remove_Click()
    ' 从右侧移除，如果是当前源文件夹的文件，则退回到左侧
    MoveItems Me.lst_Target, Me.lst_Source, False
End Sub

' 5. 按钮：移除全部 ( << )
Private Sub btn_RemoveAll_Click()
    MoveAllItems Me.lst_Target, Me.lst_Source, False
End Sub

' 6. 快捷操作：双击列表项自动移动
Private Sub lst_Source_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call btn_Add_Click
End Sub

Private Sub lst_Target_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call btn_Remove_Click
End Sub

' ==========================================================================================
' 核心逻辑：移动列表项
' ==========================================================================================
' 参数 isAdd: True表示从左到右(拼接路径)，False表示从右到左(仅看文件名)
Private Sub MoveItems(lstFrom As MSForms.ListBox, lstTo As MSForms.ListBox, isAdd As Boolean)
    Dim i As Long
    Dim currentPath As String, fullPath As String, fName As String
    Dim arrRemove() As Integer
    Dim removeCount As Integer
    
    currentPath = Me.txtFolderPath.Text
    If isAdd And currentPath = "" Then MsgBox "请先选择源文件夹！", vbExclamation: Exit Sub
    
    ' 倒序遍历（因为要删除项目）
    For i = lstFrom.ListCount - 1 To 0 Step -1
        If lstFrom.Selected(i) Then
            If isAdd Then
                ' [添加模式] 左 -> 右
                ' 左侧只显示文件名，右侧存储完整路径（为了后端稳定）
                fName = lstFrom.List(i)
                fullPath = currentPath & "\" & fName
                
                ' 防重复检查
                If Not IsInList(lstTo, fullPath) Then
                    lstTo.AddItem fullPath
                End If
            Else
                ' [移除模式] 右 -> 左
                ' 右侧是完整路径，我们需要判断这个文件是否属于当前选中的左侧文件夹
                fullPath = lstFrom.List(i)
                ' 提取文件名
                fName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
                Dim fileDir As String
                fileDir = Left(fullPath, InStrRev(fullPath, "\") - 1)
                
                ' 如果该文件属于当前选中的文件夹，则把它放回左侧列表
                If fileDir = currentPath Then
                     If Not IsInList(lstTo, fName) Then lstTo.AddItem fName
                End If
            End If
            
            ' 从源列表移除
            lstFrom.RemoveItem i
        End If
    Next i
    
    UpdateCounts
End Sub

Private Sub MoveAllItems(lstFrom As MSForms.ListBox, lstTo As MSForms.ListBox, isAdd As Boolean)
    Dim i As Long
    ' 暂时全选
    For i = 0 To lstFrom.ListCount - 1
        lstFrom.Selected(i) = True
    Next i
    ' 调用移动逻辑
    MoveItems lstFrom, lstTo, isAdd
End Sub



' 辅助：更新计数标签
Private Sub UpdateCounts()
    Me.lbl_LeftCount.Caption = "待选文件 (" & Me.lst_Source.ListCount & ")"
    Me.lbl_RightCount.Caption = "已选文件 (" & Me.lst_Target.ListCount & ")"
End Sub
' ==========================================================================================
' 3. [Tab 1] 底部执行按钮
' ==========================================================================================
' 1. 按钮：开始汇总 (对接核心引擎)
Private Sub btn_RunImport_Click()
    Dim filesColl As New Collection
    Dim i As Integer
    
    ' 检查是否已选
    If Me.lst_Target.ListCount = 0 Then
        MsgBox "请至少将一个文件移动到右侧【已选文件】列表中！", vbExclamation
        Exit Sub
    End If
    
    ' 收集文件路径 (从 lst_Target 读取)
    ' 注意：我们在移动逻辑中已经确保 lst_Target 存的是【完整路径】，所以直接传给后端即可
    For i = 0 To Me.lst_Target.ListCount - 1
        filesColl.Add Me.lst_Target.List(i)
    Next i
    
    ' --- 以下代码完全保持原样 ---
    ' 禁用界面
    Me.Enabled = False
    
    ' 调用模块中的批量处理逻辑
    Call Mod_Financial_Core.P2_Run_Batch_From_Form(filesColl, Me)
    
    ' 恢复界面
    Me.Enabled = True
    Me.lbl_Status.Caption = "处理完成！"
    Me.lbl_Prog_Bar.Width = Me.lbl_Prog_Bg.Width ' 满格
End Sub

' 2. 按钮：注入公式
Private Sub btn_InjectFormula_Click()
    Call Mod_Financial_Core.Main_InjectFormulas
End Sub

' 3. 按钮：生成模板 (修复版 - 去掉模块名前缀)
Private Sub btn_Generate_Click()
    If MsgBox("确定要清空 A 列并重新生成模板特征码吗？", vbQuestion + vbYesNo) = vbYes Then
        ' 直接调用 Sub，VBA 会自动在所有标准模块中查找
        Call Mod_Financial_Core.Main_GenerateTemplate
        MsgBox "模板初始化完成！请检查 Sys_Config 和 User_Config。", vbInformation
    End If
End Sub


' 4. 按钮：检查复杂表 (修复版)
Private Sub btn_CheckComplex_Click()
    ' 直接调用，去掉 Mod_Financial_Core 前缀
    ' 注意：P1_PostProcess_ComplexTables 必须在模块中声明为 Public (或者去掉 Private)
    Call Mod_Financial_Core.P1_PostProcess_ComplexTables(ActiveWorkbook.ActiveSheet)
    MsgBox "复杂表检查完毕。", vbInformation
End Sub

' ==========================================================================================
' 4. [Tab 2] 数据管理逻辑
' ==========================================================================================

' 2. 按钮：刷新列表
Private Sub btn_Refresh_Data_Click()
    Call RefreshDataManageList
    MsgBox "列表已刷新。", vbInformation
End Sub

' 3. 按钮：全选/反选
Private Sub btn_SelectAll_Data_Click()
    Dim i As Long
    Dim isSelect As Boolean
    
    If Me.lst_Manage_Files.ListCount = 0 Then Exit Sub
    
    ' 根据第一项的状态决定是全选还是全不选
    isSelect = Not Me.lst_Manage_Files.Selected(0)
    
    For i = 0 To Me.lst_Manage_Files.ListCount - 1
        Me.lst_Manage_Files.Selected(i) = isSelect
    Next i
End Sub

' 4. 按钮：执行删除
Private Sub btn_Delete_Data_Click()
    Dim i As Long
    Dim collDelete As New Collection
    Dim count As Long
    Dim msg As String
    
    ' 1. 收集要删除的文件名
    For i = 0 To Me.lst_Manage_Files.ListCount - 1
        If Me.lst_Manage_Files.Selected(i) Then
            collDelete.Add Me.lst_Manage_Files.List(i)
        End If
    Next i
    
    count = collDelete.count
    If count = 0 Then
        MsgBox "请先勾选需要删除数据的文件名！", vbExclamation
        Exit Sub
    End If
    
    ' 2. 二次确认
    msg = "【警告】" & vbCrLf & vbCrLf & _
          "您选中了 " & count & " 个文件的数据。" & vbCrLf & _
          "执行删除后，Master 表中对应的数据将永久丢失且无法撤销！" & vbCrLf & vbCrLf & _
          "确定要删除吗？"
          
    If MsgBox(msg, vbCritical + vbYesNo, "危险操作确认") = vbNo Then Exit Sub
    
    ' 3. 调用后端删除
    Call Mod_Financial_Core.DM_DeleteDataByFiles(collDelete)
    
    ' 4. 刷新列表并提示
    Call RefreshDataManageList
    MsgBox "成功删除了 " & count & " 个文件的数据。", vbInformation
End Sub

' [内部辅助] 刷新已导入文件列表
Private Sub RefreshDataManageList()
    Dim coll As Collection
    Dim item As Variant
    
    Me.lst_Manage_Files.Clear
    
    ' 调用 Core 获取 Master 表中的真实文件名
    Set coll = Mod_Financial_Core.DM_GetImportedFileNames()
    
    If coll.count > 0 Then
        For Each item In coll
            Me.lst_Manage_Files.AddItem item
        Next item
    Else
        Me.lst_Manage_Files.AddItem "（Master表中暂无数据）"
    End If
    
    '更新标签状态
    '假设你有一个 label 叫 lbl_DataStatus，如果没有可以忽略这行
    Me.lbl_DataStatus.Caption = "当前共有 " & coll.count & " 个已导入文件"
End Sub

' ==========================================================================================
' 5. [Tab 3] 样式配置逻辑
' ==========================================================================================

' 1. 模板切换事件：加载整套配置
Private Sub cbo_Template_Select_Change()
    Dim tName As String
    tName = Me.cbo_Template_Select.Value
    If tName = "" Then Exit Sub
    
    ' 调用 Core 加载指定列到内存字典
    Call Mod_Financial_Core.Load_Table_Style_Config(tName)
    
    ' 刷新下方 UI 显示 (复用之前的显示逻辑)
    Call Load_Style_To_UI
End Sub


' 2.[补全] 样式对象切换事件

Private Sub cbo_Style_Target_Change()
    ' 当下拉框改变时，重新加载对应的数据到下方文本框
    Call Load_Style_To_UI
End Sub

' 3. [修改] 保存按钮：保存到【当前选中的模板】
Private Sub btn_SaveStyle_Click()
    Dim tName As String
    Dim prefix As String
    
    tName = Me.cbo_Template_Select.Value
    prefix = Get_Style_Prefix()
    
    If tName = "" Or prefix = "" Then Exit Sub
    
    ' 校验
    If Not IsNumeric(Me.txtStyleSize.Value) Then MsgBox "字号必须为数字", vbExclamation: Exit Sub
    If Not IsNumeric(Me.txtStyleIndex.Value) Then MsgBox "缩进必须为数字", vbExclamation: Exit Sub
    
    ' 保存字段
    Call Mod_Financial_Core.Save_Config_To_Template(tName, prefix & "_FontName", Me.txtStyleFont.Value)
    Call Mod_Financial_Core.Save_Config_To_Template(tName, prefix & "_FontSize", Me.txtStyleSize.Value)
    Call Mod_Financial_Core.Save_Config_To_Template(tName, prefix & "_SpaceBefore", Me.txtStyleBefore.Value)
    Call Mod_Financial_Core.Save_Config_To_Template(tName, prefix & "_SpaceAfter", Me.txtStyleAfter.Value)
    Call Mod_Financial_Core.Save_Config_To_Template(tName, prefix & "_LineSpacing", Me.txtStyleSpace.Value)
    Call Mod_Financial_Core.Save_Config_To_Template(tName, prefix & "_Alignment", CStr(Me.cboStyleAlign.ListIndex))
    
    ' [新增] 保存缩进
    Call Mod_Financial_Core.Save_Config_To_Template(tName, prefix & "_FirstLineIndent", Me.txtStyleIndex.Value)
    
    If prefix = "Table" Then
        Call Mod_Financial_Core.Save_Config_To_Template(tName, prefix & "_RowHeight", Me.txtStyleHeight.Value)
    End If
    
    Call Mod_Financial_Core.Load_Table_Style_Config(tName)
    
    MsgBox "已更新模板 [" & tName & "] 中的 [" & Me.cbo_Style_Target.Value & "] 样式配置。", vbInformation
End Sub

' 4. [新增] 另存为新模板按钮
Private Sub btn_New_Template_Click()
    Dim newName As String
    Dim sourceName As String
    
    sourceName = Me.cbo_Template_Select.Value
    newName = InputBox("请输入新模板名称：" & vbCrLf & "(将基于当前模板 [" & sourceName & "] 创建)", "新建模板")
    
    If Trim(newName) = "" Then Exit Sub
    
    ' 检查是否重名
    If IsInCombo(Me.cbo_Template_Select, newName) Then
        MsgBox "模板名称已存在！", vbExclamation
        Exit Sub
    End If
    
    ' 实现原理：
    ' 1. 读取当前内存中的所有配置 (g_TableConfig)
    ' 2. 遍历字典，全部写入到新的一列
    
    Application.ScreenUpdating = False
    Dim dict As Object
    Dim k As Variant
    
    ' 确保内存是最新的
    Call Mod_Financial_Core.Load_Table_Style_Config(sourceName)
    Set dict = Mod_Financial_Core.g_TableConfig
    
    If Not dict Is Nothing Then
        For Each k In dict.Keys
            Call Mod_Financial_Core.Save_Config_To_Template(newName, CStr(k), dict(k))
        Next k
    End If
    Application.ScreenUpdating = True
    
    ' 刷新列表并选中新模板
    Call Refresh_Template_List
    Me.cbo_Template_Select.Value = newName
    
    MsgBox "新模板 [" & newName & "] 创建成功！", vbInformation
End Sub

' 5. 按钮：保存黑名单
Private Sub btn_SaveBlacklist_Click()
    Dim wsUser As Worksheet
    Dim arrList As Variant
    Dim i As Long
    
    On Error Resume Next
    Set wsUser = ActiveWorkbook.Sheets("User_Config")
    On Error GoTo 0
    
    If wsUser Is Nothing Then
        MsgBox "找不到 User_Config 表，请先运行【生成模板】。", vbCritical
        Exit Sub
    End If
    
    ' 获取文本框内容，按换行符分割
    arrList = Split(Me.txt_Blacklist.Text, vbCrLf)
    
    ' 写入 User_Config A列
    wsUser.Range("A2:A1000").ClearContents
    For i = 0 To UBound(arrList)
        If Trim(arrList(i)) <> "" Then
            wsUser.Cells(i + 2, 1).Value = Trim(arrList(i))
        End If
    Next i
    
    ' [新增] 强制刷新内存缓存，确保下次操作使用新配置
    Call Mod_Financial_Core.Force_Reload_Config
    
    MsgBox "黑名单已保存更新！", vbInformation
End Sub

' 6.内部方法：加载黑名单到 UI
Private Sub LoadBlacklistToUI()
    Dim wsUser As Worksheet
    Dim lastRow As Long, i As Long
    Dim txt As String
    
    On Error Resume Next
    Set wsUser = ActiveWorkbook.Sheets("User_Config")
    On Error GoTo 0
    
    If Not wsUser Is Nothing Then
        lastRow = wsUser.Cells(wsUser.Rows.count, "A").End(xlUp).Row
        If lastRow >= 2 Then
            For i = 2 To lastRow
                txt = txt & wsUser.Cells(i, 1).Value & vbCrLf
            Next i
        End If
        Me.txt_Blacklist.Text = txt
    Else
        Me.txt_Blacklist.Text = "（未找到配置文件）"
    End If
End Sub

' ==========================================================================================
' 7.[修复] 核心逻辑：加载样式数据到 UI 控件
' ==========================================================================================
Private Sub Load_Style_To_UI()
    Dim Target As String
    Dim prefix As String
    
    prefix = Get_Style_Prefix()
    If prefix = "" Then Exit Sub
    
    ' 定义默认值
    Dim defFont As String, defSize As String, defBefore As String, defAfter As String
    Dim defLine As String, defAlign As String, defHeight As String
    Dim defIndent As String ' [新增]
    
    ' 默认配置逻辑调整
    ' Table: 默认对齐设为 "0" (不调整)
    ' P(正文): 默认缩进 "2" (字符)
    Select Case prefix
        ' -------------------------------------------------------------------------
        ' [修改] 表格配置
        Case "Table":
            defFont = "宋体": defSize = "9": defHeight = "22": defAlign = "0": defIndent = "0"
            
        ' -------------------------------------------------------------------------
        ' [原有] H1 (一级标题)
        Case "H1":
            defFont = "黑体": defSize = "10.5": defLine = "1": defBefore = "0": defAfter = "0": defAlign = "1": defIndent = "2"
            
        ' -------------------------------------------------------------------------
        ' [原有] P (正文)
        Case "P":
            defFont = "宋体": defSize = "10.5": defLine = "1": defBefore = "0": defAfter = "0": defAlign = "1": defIndent = "2"
            
        Case Else:
            defFont = "宋体": defSize = "10.5": defAlign = "0": defIndent = "0"
    End Select
    
    Call Mod_Financial_Core.Init_Global_Config
    Dim dict As Object
    Set dict = Mod_Financial_Core.g_TableConfig
    
    ' 填入原有控件
    Me.txtStyleFont.Value = GetDictVal(dict, prefix & "_FontName", defFont)
    Me.txtStyleSize.Value = GetDictVal(dict, prefix & "_FontSize", defSize)
    Me.txtStyleBefore.Value = GetDictVal(dict, prefix & "_SpaceBefore", defBefore)
    Me.txtStyleAfter.Value = GetDictVal(dict, prefix & "_SpaceAfter", defAfter)
    Me.txtStyleSpace.Value = GetDictVal(dict, prefix & "_LineSpacing", IIf(defLine = "", "1", defLine))
    
    ' [新增] 填入缩进控件
    Me.txtStyleIndex.Value = GetDictVal(dict, prefix & "_FirstLineIndent", defIndent)
    
    ' 处理对齐方式 (0=不调整, 1=左, 2=中, 3=右, 4=两端)
    Dim alignVal As Integer
    alignVal = val(GetDictVal(dict, prefix & "_Alignment", defAlign))
    If alignVal >= 0 And alignVal < Me.cboStyleAlign.ListCount Then
        Me.cboStyleAlign.ListIndex = alignVal
    Else
        Me.cboStyleAlign.ListIndex = 3 ' 默认不调整
    End If
    
    ' 处理表格行高
    If prefix = "Table" Then
        Me.txtStyleHeight.Enabled = True
        Me.txtStyleHeight.Value = GetDictVal(dict, prefix & "_RowHeight", defHeight)
        Me.txtStyleHeight.BackColor = vbWhite
    Else
        Me.txtStyleHeight.Value = ""
        Me.txtStyleHeight.Enabled = False
        Me.txtStyleHeight.BackColor = &H8000000F
    End If
End Sub


' 8.[新增] 刷新模板下拉框
Private Sub Refresh_Template_List()
    Dim coll As Collection
    Dim item As Variant
    Dim currentSel As String
    
    ' 记录当前选中的，刷新后尝试恢复
    currentSel = Me.cbo_Template_Select.Value
    
    Me.cbo_Template_Select.Clear
    Set coll = Mod_Financial_Core.Get_All_Template_Names()
    
    For Each item In coll
        Me.cbo_Template_Select.AddItem item
    Next item
    
    ' 默认选中第一个(通用模板)，或者恢复之前的选择
    If Me.cbo_Template_Select.ListCount > 0 Then
        If currentSel <> "" And IsInCombo(Me.cbo_Template_Select, currentSel) Then
            Me.cbo_Template_Select.Value = currentSel
        Else
            Me.cbo_Template_Select.ListIndex = 0
        End If
    End If
    
    ' 触发一次加载
    Call cbo_Template_Select_Change
End Sub
' ==========================================================================================
' 6. 通用辅助函数 (Helpers)
' ==========================================================================================


' 1.系统级文件夹选择器 (兼容 WPS/Excel)
Private Function GetFolder_Shell() As String
    Dim shellApp As Object
    Dim folder As Object
    
    ' 调用 Windows 底层 Shell
    Set shellApp = CreateObject("Shell.Application")
    ' 参数说明: 0=桌面为根, 字符串=标题, 0=默认选项
    On Error Resume Next
    Set folder = shellApp.BrowseForFolder(0, "请选择包含Excel文件的文件夹", 0)
    On Error GoTo 0
    
    If Not folder Is Nothing Then
        ' 获取路径
        GetFolder_Shell = folder.Self.path
        ' 补全路径末尾的斜杠 (防止 C:Windows 变成 C:Windows\)
        If Right(GetFolder_Shell, 1) <> "\" Then GetFolder_Shell = GetFolder_Shell & "\"
    Else
        GetFolder_Shell = ""
    End If
    
    Set folder = Nothing
    Set shellApp = Nothing
End Function


' 2.  辅助函数：获取当前选中的前缀
Private Function Get_Style_Prefix() As String
    Dim txt As String
    ' 确保 cbo_Style_Target 有值
    If Me.cbo_Style_Target.Value = "" Then Exit Function
    
    txt = Me.cbo_Style_Target.Value
    
    ' 根据选中项返回对应前缀
    If InStr(txt, "Table") > 0 Or InStr(txt, "表格") > 0 Then Get_Style_Prefix = "Table"
    If InStr(txt, "H1") > 0 Or InStr(txt, "一级") > 0 Then Get_Style_Prefix = "H1"
    If InStr(txt, "H2") > 0 Or InStr(txt, "二级") > 0 Then Get_Style_Prefix = "H2"
    If InStr(txt, "P") > 0 Or InStr(txt, "正文") > 0 Then Get_Style_Prefix = "P"
End Function


' ==========================================================================================
' 3. 辅助函数：字典取值兜底 (防止字典为空报错)
' ==========================================================================================
Private Function GetDictVal(d As Object, k As String, def As String) As String
    If d Is Nothing Then
        GetDictVal = def
        Exit Function
    End If
    
    If d.exists(k) Then
        GetDictVal = d(k)
    Else
        GetDictVal = def
    End If
End Function

' 4.辅助：检查重复
Private Function IsInList(lst As MSForms.ListBox, val As String) As Boolean
    Dim i As Long
    IsInList = False
    For i = 0 To lst.ListCount - 1
        ' 使用 StrComp 进行文本比较 (vbTextCompare)
        If StrComp(lst.List(i), val, vbTextCompare) = 0 Then
            IsInList = True
            Exit Function
        End If
    Next i
End Function


' 5.辅助：检查 ComboBox 是否包含某项
Private Function IsInCombo(cbo As MSForms.ComboBox, val As String) As Boolean
    Dim i As Long
    IsInCombo = False
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = val Then IsInCombo = True: Exit Function
    Next i
End Function


Public Sub Update_Progress(curr As Integer, total As Integer, currentTarget As String)
    Dim pct As Single
    Dim maxWidth As Single
    
    ' 1. 计算百分比
    If total > 0 Then pct = curr / total Else pct = 0
    
    ' 2. 计算进度条宽度
    maxWidth = Me.lbl_Prog_Bg.Width
    Me.lbl_Prog_Bar.Width = maxWidth * pct
    
    ' 3. 更新文字信息 (使用标准符号)
    ' 效果：>> 正在处理: XX公司 ... ( 3 / 10 )
    Me.lbl_Status.Caption = ">> 正在处理: " & currentTarget & " ...  ( 已完成 " & curr & " / 共 " & total & " )"
    
    ' 4. 强制刷新界面
    Me.Repaint
    DoEvents
End Sub
