VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmConfirm 
   Caption         =   "UserForm2"
   ClientHeight    =   3576
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5928
   OleObjectBlob   =   "FrmConfirm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "FrmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =================================================================
' 窗体名称: FrmConfirm
' 属性设置: ShowModal = False (必须!)
' =================================================================
Option Explicit

' 防止用户直接点右上角X关闭导致逻辑混乱
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnCancel_Click ' 视为点击取消
    End If
End Sub

Private Sub btnYes_Click()
    g_ConfirmResult = vbYes ' 全局变量记录结果
    Me.Hide
End Sub

Private Sub btnNo_Click()
    g_ConfirmResult = vbNo
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    g_ConfirmResult = vbCancel
    Me.Hide
End Sub
