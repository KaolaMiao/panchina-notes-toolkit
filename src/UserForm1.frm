VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4428
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   2556
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 缓存所有表格名称的字典，用于搜索过滤
Private m_AllTableNames As Object

' =======================================================
' 1. 窗体初始化：加载数据并备份
' =======================================================
Private Sub UserForm_Initialize()
    Dim key As Variant
    
    ' 初始化字典
    Set m_AllTableNames = CreateObject("Scripting.Dictionary")
    
    ' 清空列表框
    Me.ListBox1.Clear
    
    ' 检查全局变量是否有数据
    If Not g_tableRanges Is Nothing Then
        For Each key In g_tableRanges.Keys
            ' 1. 添加到界面列表
            Me.ListBox1.AddItem key
            ' 2. 添加到后台备份用于搜索
            m_AllTableNames.Add key, key
        Next key
    End If
    
    ' 设置焦点到搜索框，方便直接打字
    Me.txtSearch.SetFocus
End Sub

' =======================================================
' 2. 搜索框输入事件：实时过滤列表
' =======================================================
Private Sub txtSearch_Change()
    Dim searchText As String
    Dim key As Variant
    
    searchText = UCase(Trim(Me.txtSearch.Text))
    
    ' 暂时关闭重绘以防闪烁
    Me.ListBox1.Visible = False
    Me.ListBox1.Clear
    
    ' 循环备份的完整名单
    For Each key In m_AllTableNames.Keys
        ' 如果搜索框为空，或者名称包含搜索词
        If searchText = "" Or InStr(1, UCase(key), searchText) > 0 Then
            Me.ListBox1.AddItem key
        End If
    Next key
    
    Me.ListBox1.Visible = True
End Sub

' =======================================================
' 3. 确认按钮 (你需要确保窗体上有个按钮对应这个事件，通常是CommandButton1)
' =======================================================
Private Sub CommandButton1_Click()
    Dim i As Long
    Dim hasSelection As Boolean
    
    ' 填充全局选择字典
    Set g_tablesToUpdate = CreateObject("Scripting.Dictionary")
    
    For i = 0 To Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) Then
            g_tablesToUpdate.Add Me.ListBox1.List(i), Me.ListBox1.List(i)
            hasSelection = True
        End If
    Next i
    
    If Not hasSelection Then
        MsgBox "请至少选择一个表格！", vbExclamation
        Exit Sub
    End If
    
    g_UserCancelled = False
    Me.Hide
End Sub

' =======================================================
' 4. 取消/关闭按钮
' =======================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' 点击右上角X关闭
        Cancel = 1
        g_UserCancelled = True
        Me.Hide
    End If
End Sub

' 如果你有取消按钮 (CommandButton2)
Private Sub CommandButton2_Click()
    g_UserCancelled = True
    Me.Hide
End Sub
