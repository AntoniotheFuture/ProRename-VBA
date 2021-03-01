VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_cut 
   Caption         =   "截去和替换"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5025
   OleObjectBlob   =   "F_cut.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "F_cut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub L_AddLetters1_Click()
Me.T_find.Value = "[A-Za-z]"
End Sub

Private Sub L_AddNum_Click()
Me.T_find.Value = "[0-9]"
End Sub

Private Sub L_AddSymbols_Click()
Me.T_find.Value = "[`~!@#$%^&*()_+-=;',./{}:<>?*]"
End Sub

Private Sub L_Cancel_Click()
Me.hide
End Sub

Private Sub L_Ok_Click()

If Not Me.T_find.Value = "" Then
    CutFind = Me.T_find.Value
Else
    MsgBox ("正则表达式不能为空")
    Exit Sub
End If
CutReplace = Me.T_replace.Value
Me.hide
End Sub

Private Sub T_replace_Change()
If Not checkAVB(Me.T_replace) Then
    MsgBox ("不能使用文件系统不支持的字符")
    Me.T_replace = ""
End If
End Sub

Private Sub UserForm_Activate()
CutFind = ""
CutReplace = ""
End Sub
