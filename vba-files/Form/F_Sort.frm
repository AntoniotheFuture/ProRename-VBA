VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Sort 
   Caption         =   "添加筛选"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4950
   OleObjectBlob   =   "F_Sort.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "F_Sort"
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
    SortFind = Me.T_find.Value
Else
    MsgBox ("正则表达式不能为空")
    Exit Sub
End If
If O_meet Then
    SortCondition = "满足"
Else
    SortCondition = "排除"
End If
Me.hide
End Sub

Private Sub UserForm_Activate()
SortFind = ""
SortCondition = ""
End Sub

