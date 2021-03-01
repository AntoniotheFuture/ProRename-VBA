VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "批量重命名-查重定制版"
   ClientHeight    =   9450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17610
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'两个数组保存处理动作
Dim ActionsType As Collection
Dim ActionContent As Collection
'是否已经预览过
Dim Previewed As Boolean

'UI操作
'更新界面显示
Sub UpdateView()
With Me
    .Button_DeleteSource.Visible = False
    If .List_Source.ListItems.Count > 0 Then
        .Button_DeleteSource.Visible = .List_Source.SelectedItem.Index > -1
    End If
    .Button_DeleteType.Visible = False
    If .List_Type.ListItems.Count > 0 Then
        .Button_DeleteType.Visible = .List_Type.SelectedItem.Index > -1
    End If
    .L_DelCut.Visible = False
    If .List_Cut.ListItems.Count > 0 Then
        .L_DelCut.Visible = .List_Cut.SelectedItem.Index > -1
    End If
    
    '显示当前选择的Action
    .L_DelAction1.Visible = False
    If .L_Actions.ListItems.Count > 0 Then
        Dim selected As Boolean
        selected = .L_Actions.SelectedItem.Index > -1
       .L_DelAction1.Visible = selected
       .L_CurrentAction.Caption = .L_Actions.SelectedItem.Index
       '更新各页的动作按钮
        .BT_ActionUpdate1.Visible = False
        .BT_ActionUpdate2.Visible = selected
        .BT_ActionUpdate3.Visible = selected
        .BT_ActionUpdate4.Visible = selected
    End If
    
    .Button_Run.Visible = Previewed
End With
End Sub

'更新状态语
Sub UpdateNotice(Key As Integer)
'Dim NoticeKeys
NoticeTexts = Split("双击打开文件,2222", ",")
Me.L_Notice.Caption = NoticeTexts(Key)
End Sub
'更新进度条
Sub UpdateProgress(progress As Integer)
Me.ProgressBar1.Value = progress
DoEvents
End Sub

Function getActionText(page As Integer) As String
Dim text As String
With Me
    Select Case page
    Case 1
        If .T_Letters.Value = "" Then
            MsgBox ("字符不能为空")
            Exit Function
        Else
            text = .T_Letters.Value
        End If
    Case 2
        If .CB_ReverseNumbering.Value Then
            text = .Text_Len.Value & ";" & .Text_StartNum.Value & ";" & 1
        Else
            text = .Text_Len.Value & ";" & .Text_StartNum.Value & ";" & 0
        End If
    Case 3
        If .T_SelectedTimeFormat.Value = "" Then
             MsgBox ("日期格式不能为空")
             Exit Function
        End If
        TimeType = 1
        If .O_CurrentTime.Value Then TimeType = 1
        If .O_EditTime.Value Then TimeType = 2
        If .O_CreateTime.Value Then TimeType = 3
        text = TimeType & ";" & .T_SelectedTimeFormat.Value
    Case 4
        letters = 0
        nums = 0
        diff = 0
        If .CB_RamdomLetters.Value Then letters = 1
        If .CB_RamdomNums.Value Then nums = 1
        If .CB_RamdomDiff.Value Then diff = 1
        text = letters & ";" & nums & ";" & diff & ";" & .T_RamdomD.Value
    End Select
End With
getActionText = text
End Function

Function getActionType(page As Integer) As String
Dim atype As String
Select Case page
    Case 1
        atype = "字符"
    Case 2
        atype = "序号"
    Case 3
        atype = "时间"
    Case 4
        atype = "随机"
End Select
getActionType = atype
End Function

Sub AddAction(page As Integer)
Dim newitem As ListItem
With Me
    Set newitem = .L_Actions.ListItems.Add
    newitem.text = .L_Actions.ListItems.Count
    newitem.SubItems(1) = getActionType(page)
    newitem.SubItems(2) = getActionText(page)
End With
UpdateView
End Sub

'加载某个动作的设定
Sub LoadAction()
Dim atype As String
Dim acontent As String
Dim Li As ListItem
For Each Li In Me.L_Actions.ListItems
    If Li.selected Then
        atype = Li.SubItems(1)
        acontent = Li.SubItems(2)
        Select Case atype
            Case "字符"
            MultiPage1.Value = 0
            Me.T_Letters.Value = acontent
            Case "序号"
            MultiPage1.Value = 1
            Me.Text_Len.Value = Split(acontent, ";")(0)
            Me.Text_StartNum.Value = Split(acontent, ";")(1)
            Me.CB_ReverseNumbering.Value = (Split(acontent, ";")(2) = "1")
            Case "时间"
            MultiPage1.Value = 2
            Select Case Split(acontent, ";")(0)
                Case "1"
                Me.O_CurrentTime.Value = True
                Case "2"
                Me.O_EditTime.Value = True
                Case "3"
                Me.O_CreateTime.Value = True
            End Select
            Me.T_SelectedTimeFormat.Value = Split(acontent, ";")(1)
            Case "随机"
            MultiPage1.Value = 3
            Me.CB_RamdomLetters.Value = (Split(acontent, ";")(0) = "1")
            Me.CB_RamdomNums.Value = (Split(acontent, ";")(1) = "1")
            Me.CB_RamdomDiff.Value = (Split(acontent, ";")(2) = "1")
            Me.T_RamdomD.Value = Split(acontent, ";")(3)
        End Select
    End If
Next
End Sub

Sub UpdateAction(page As Integer)
Dim Li As ListItem
With Me
    For Each Li In .L_Actions.ListItems
        If Li.selected Then
            Li.SubItems(1) = getActionType(page)
            Li.SubItems(2) = getActionText(page)
        End If
    Next
End With
End Sub

'更新Action列表显示
Sub UpdateActionList()
Dim c As New Collection
Dim c2 As New Collection
Dim Li As ListItem
Dim id As Integer
Dim total As Integer
Dim newitem As ListItem
total = 1
With Me
    For Each Li In .L_Actions.ListItems
        c.Add Li.SubItems(1)
        c2.Add Li.SubItems(2)
        id = Li.text
        If Not id = total Then
            ReplaceExpTextNo id, total
        End If
        total = total + 1
    Next
    .L_Actions.ListItems.Clear
    total = 1
    For Each s In c
        Set newitem = .L_Actions.ListItems.Add
        newitem.text = total
        newitem.SubItems(1) = s
        newitem.SubItems(2) = c2(total)
        total = total + 1
    Next
End With
End Sub

'替换表达式序号
Sub ReplaceExpTextNo(fromno As Integer, tono As Integer)
Dim arr
Dim s As String
Dim newText As String
arr = Split(Me.T_Formula.Value, "#")
For i = 0 To UBound(arr)
    s = arr(i)
    If s = (fromno & "") Then s = (tono & "")
    newText = newText & "#" & s
Next
Me.T_Formula.Value = newText
End Sub

Private Sub BT_ActionAdd1_Click()
AddAction (1)
End Sub

Private Sub BT_ActionAdd2_Click()
AddAction (2)
End Sub

Private Sub BT_ActionAdd3_Click()
AddAction (3)
End Sub

Private Sub BT_ActionAdd4_Click()
AddAction (4)
End Sub

Private Sub BT_ActionUpdate1_Click()
UpdateAction (1)
End Sub

Private Sub BT_ActionUpdate2_Click()
UpdateAction (2)
End Sub

Private Sub BT_ActionUpdate3_Click()
UpdateAction (3)
End Sub

Private Sub BT_ActionUpdate4_Click()
UpdateAction (4)
End Sub


Private Sub BT_AddD_Click()
Dim D As Integer
D = 0
With Me
    If Not .Text_Len.Value = "" And IsNumeric(.Text_Len.Value) Then
        D = .Text_Len.Value * 1
    End If
    D = D + 1
    .Text_Len.Value = D
End With
End Sub

Private Sub BT_AddRamdomD_Click()
Dim D As Integer
D = 0
With Me
    If Not .T_RamdomD.Value = "" And IsNumeric(.T_RamdomD.Value) Then
        D = .T_RamdomD.Value * 1
    End If
    D = D + 1
    .T_RamdomD.Value = D
End With
End Sub

Private Sub BT_ReduceD_Click()
Dim D As Integer
D = 0
With Me
    If Not .Text_Len.Value = "" And IsNumeric(.Text_Len.Value) Then
        D = .Text_Len.Value * 1
    End If
    If D = 0 Then Exit Sub
    D = D - 1
    If D = 0 Then
        .Text_Len.Value = ""
    Else
        .Text_Len.Value = D
    End If
End With
End Sub

Private Sub BT_ReduceRamdomD_Click()
Dim D As Integer
D = 0
With Me
    If Not .T_RamdomD.Value = "" And IsNumeric(.T_RamdomD.Value) Then
        D = .T_RamdomD.Value * 1
    End If
    If D = 1 Then Exit Sub
    D = D - 1
End With
End Sub

Private Sub Button_About_Click()
MsgBox "批量重命名工具" & Chr(10) & "版本:V1.0" & Chr(10) & "作者:AntoniotheFuture" & Chr(10) & "QQ：547052212", vbOKOnly, "关于"
End Sub

Private Sub Button_AddSource_Click()
Set Folder = Application.FileDialog(msoFileDialogFolderPicker)
With Me
    If Folder.Show = -1 Then
        Set newitem = .List_Source.ListItems.Add()
        newitem.text = Folder.SelectedItems(1)
    End If
End With
Previewed = False
UpdateView

End Sub

Private Sub Button_AddType_Click()
F_Sort.Show (1)
With Me
    If Not SortFind = "" Then
        Set newitem = .List_Type.ListItems.Add()
        newitem.text = SortCondition
        newitem.SubItems(1) = SortFind
    End If
    Previewed = False
    UpdateView
End With
End Sub

Private Sub Button_DeleteSource_Click()
With Me.List_Source
    For i = .ListItems.Count To 1 Step -1
        If .ListItems(i).selected = True Then
            .ListItems.Remove (i)
        End If
    Next
End With
Previewed = False
UpdateView
End Sub

Private Sub Button_DeleteType_Click()
With Me.List_Type
    For i = .ListItems.Count To 1 Step -1
        If .ListItems(i).selected = True Then
            .ListItems.Remove (i)
        End If
    Next
End With
Previewed = False
UpdateView
End Sub
Private Sub Button_PreView_Click()
PreView
End Sub

Private Sub Button_Run_Click()
Dim OldName1
Dim NewName1
Dim Li As ListItem
Dim OldPath As String, OldName As String, NewName As String
Dim NewPath As String
'检查表单
With Me
    If .List_Files.ListItems.Count = 0 Then
        MsgBox ("没有发现满足条件的文件")
        Exit Sub
    End If
    
    For Each Li In .List_Files.ListItems
        If Li.text = "" Then
            OldName = Li.SubItems(1)
            NewName = Li.SubItems(2)
            OldPath = Li.SubItems(3)
            NewPath = Left(OldPath, Len(OldPath) - Len(OldName)) & NewName
            Name OldPath As NewPath
            Li.text = "OK"
        End If
    Next
End With

'如果有新增的随机符号则加入到数据库
If NewDiffs.Count > 0 Then
    AddToDiffDB NewDiffs
    ThisWorkbook.Save
    'UpdateNotice ("新的随机字符已保存在本工作簿")
End If

End Sub


Private Sub CB_TimeFormat_Change()
    Me.T_SelectedTimeFormat.Value = Me.CB_TimeFormat.Value
End Sub


Private Sub L_Actions_Click()
LoadAction
UpdateView
End Sub

Private Sub L_Actions_ItemClick(ByVal Item As MSComctlLib.ListItem)
UpdateView
End Sub

Private Sub L_AddCut_Click()
F_cut.Show (1)
With Me
    If Not CutFind = "" Then
        Set newitem = .List_Cut.ListItems.Add()
        newitem.text = CutFind
        newitem.SubItems(1) = CutReplace
    End If
    Previewed = False
    UpdateView
End With
End Sub

Private Sub L_DelAction1_Click()
With Me.L_Actions
    For i = .ListItems.Count To 1 Step -1
        If .ListItems(i).selected Then
            .ListItems.Remove (i)
        End If
    Next
End With
UpdateView
UpdateActionList
End Sub

Private Sub L_DelCut_Click()
With Me.List_Cut
    For i = .ListItems.Count To 1 Step -1
        If .ListItems(i).selected = True Then
            .ListItems.Remove (i)
        End If
    Next
End With
Previewed = False
UpdateView
End Sub

Private Sub L_Setting_Click()
    MsgBox ("无可用设置")
End Sub

Private Sub List_Files_DblClick()
With Me.List_Files
    For i = 1 To .ListItems.Count
        If .ListItems(i).selected = True And .ListItems(i).SubItems(2) = "就绪" Then
            mOpen = Shell("Explorer.exe " & .ListItems(i).SubItems(3) & "\" & .ListItems(i), vbNormalFocus)
        ElseIf .ListItems(i).selected = True And .ListItems(i).SubItems(2) = "玩逼" Then
            mOpen = Shell("Explorer.exe " & .ListItems(i).SubItems(3) & "\" & .ListItems(i).SubItems(1), vbNormalFocus)
        End If
    Next
End With
End Sub

Private Sub List_Source_Click()
UpdateView
End Sub

Private Sub List_Source_DblClick()
With Me.List_Source
For i = 1 To .ListItems.Count
    If .ListItems(i).selected = True Then
        mOpen = Shell("Explorer.exe " & .ListItems(i), vbNormalFocus)
    End If
Next
End With
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub List_Source_ItemClick(ByVal Item As MSComctlLib.ListItem)
UpdateView
End Sub

Private Sub O_SortByCreateTime_Click()
Previewed = False
UpdateView
End Sub

Private Sub O_SortByEditTime_Click()
Previewed = False
UpdateView
End Sub

Private Sub O_SortByFileSize_Click()
Previewed = False
UpdateView
End Sub

Private Sub O_SortByLetters_Click()
Previewed = False
UpdateView
End Sub

Private Sub O_SortByOldNum_Click()
Previewed = False
UpdateView
End Sub

Private Sub O_Unsort_Click()
Previewed = False
UpdateView
End Sub

Private Sub T_FilterET_AfterUpdate()
Previewed = False
UpdateView
End Sub

Private Sub T_FilterET_Enter()
With ALDTPicker
    .OldValue = Me.T_FilterET
    .Show (1)
    Me.T_FilterET = .DateTime
End With
End Sub

Private Sub T_FilterSMax_AfterUpdate()
Previewed = False
UpdateView
End Sub

Private Sub T_FilterSMin_AfterUpdate()
Previewed = False
UpdateView
End Sub

Private Sub T_FilterST_AfterUpdate()
Previewed = False
UpdateView
End Sub

Private Sub T_FilterST_Enter()
With ALDTPicker
    .OldValue = Me.T_FilterST
    .Show (1)
    Me.T_FilterST = .DateTime
End With
End Sub

Private Sub T_Formula_AfterUpdate()
Previewed = False
UpdateView
End Sub

Private Sub T_Letters_Change()
If Not checkAVB(Me.T_Letters) Then
    MsgBox ("不能使用文件系统不支持的字符")
    Me.T_Letters = ""
End If
End Sub

Private Sub T_SelectedTimeFormat_Change()
If Not checkAVB(Me.T_SelectedTimeFormat) Then
    MsgBox ("不能使用文件系统不支持的字符")
    Me.T_SelectedTimeFormat = ""
End If
End Sub

Private Sub Text_StartNum_AfterUpdate()
With Me
    If Not Me.Text_StartNum.Value = "" And Not IsNumeric(.Text_StartNum.Value) Then
        MsgBox ("请输入数字")
        Me.Text_StartNum.Value = ""
        Exit Sub
    End If
End With
End Sub

Private Sub Text_StartNum_Change()

End Sub

Private Sub UserForm_Activate()
UpdateView
End Sub

Private Sub UserForm_Initialize()
With Me
    .List_Source.ColumnHeaders.Add , , "路径", .List_Source.Width
    .List_Type.ColumnHeaders.Add , , "条件", .List_Type.Width * 0.2
    .List_Type.ColumnHeaders.Add , , "正则表达式", .List_Type.Width * 0.8
    .List_Cut.ColumnHeaders.Add , , "替换", .List_Cut.Width * 0.5
    .List_Cut.ColumnHeaders.Add , , "替换为", .List_Cut.Width * 0.5
    .L_Actions.ColumnHeaders.Add , , "序号", .L_Actions.Width * 0.2
    .L_Actions.ColumnHeaders.Add , , "类型", .L_Actions.Width * 0.8
    .L_Actions.ColumnHeaders.Add , , "内容", .L_Actions.Width * 2
    .List_Files.ColumnHeaders.Add , , "状态", .List_Files.Width * 0.1
    .List_Files.ColumnHeaders.Add , , "原文件名", .List_Files.Width * 0.3
    .List_Files.ColumnHeaders.Add , , "新文件名", .List_Files.Width * 0.3
    .List_Files.ColumnHeaders.Add , , "路径", .List_Files.Width
    
    For x = 2 To Sheet1.UsedRange.Rows.Count
        .Combo_MissionNames.AddItem Sheet1.Cells(x, 2)
    Next
    
    '加载日期时间格式
    Dim c As New Collection
    Set c = getTimeFormats
    For Each s In c
        CB_TimeFormat.AddItem s
    Next
    
End With
End Sub
'检查表单
Function CheckForm() As Boolean

With Me
    If Not .T_FilterST = "" And Not IsDate(.T_FilterST) Then
        MsgBox ("修改时间范围格式错误")
        Exit Function
    End If
    If Not .T_FilterET = "" And Not IsDate(.T_FilterET) Then
        MsgBox ("修改时间范围格式错误")
        Exit Function
    End If
    If Not .T_FilterSMin = "" And Not IsNumeric(.T_FilterSMin) Then
        MsgBox ("文件大小范围格式错误")
        Exit Function
    End If
    If Not .T_FilterSMax = "" And Not IsNumeric(.T_FilterSMax) Then
        MsgBox ("文件大小范围格式错误")
        Exit Function
    End If
    If InStr(.T_Formula, "#@") = 0 Then
        MsgBox ("表达式格式错误")
        Exit Function
    End If
End With
CheckForm = True
End Function


Sub PreView()
If Not CheckForm Then Exit Sub
Dim Li As ListItem

Dim sourceCollection As New Collection
Dim FilterCollection As New Collection
Dim FilterTypeCollertion As New Collection
Dim SortType As String
Dim FilterST As Date
Dim FilterET As Date
Dim MinSize As Double
Dim MaxSize As Double
Dim CutCollection As New Collection
Dim ReplaceCollection As New Collection
Dim ActionTypes As New Collection
Dim ActionContents As New Collection
Dim Formula As String

Dim OldFiles As New Collection
Dim newFiles As New Collection

Dim Exclude As Boolean

With Me
    .List_Files.ListItems.Clear
    
    For Each Li In .List_Source.ListItems
        sourceCollection.Add Li.text
    Next
    
    For Each Li In .List_Type.ListItems
        FilterTypeCollertion.Add Li.text
        FilterCollection.Add Li.SubItems(1)
    Next
    
    If .O_Unsort Then SortType = ""
    If .O_SortByEditTime Then SortType = "EditTime"
    If .O_SortByCreateTime Then SortType = "CreateTime"
    If .O_SortByFileSize Then SortType = "FileSize"
    If .O_SortByOldNum Then SortType = "OldNum"
    If .O_SortByLetters Then SortType = "Letters"

    
    For Each Li In .List_Cut.ListItems
        CutCollection.Add Li.text
        ReplaceCollection.Add Li.SubItems(1)
    Next
    
    For Each Li In .L_Actions.ListItems
        ActionTypes.Add Li.SubItems(1)
        ActionContents.Add Li.SubItems(2)
    Next
    
    Formula = .T_Formula
    
    Set OldFiles = getherFiles(sourceCollection)
    For i = 1 To FilterCollection.Count
        Exclude = False
        If FilterTypeCollertion.Item(i) = "满足" Then Exclude = True
        Set OldFiles = FilterByReg(OldFiles, FilterCollection.Item(i), Exclude)
    Next
    
    If Not (.T_FilterST = "" And .T_FilterET = "") Then
        If .T_FilterST = "" Then
            FilterST = #1/1/1900#
        Else
            FilterST = CDate(.T_FilterST)
        End If
        If .T_FilterET = "" Then
            FilterET = #12/31/9999#
        Else
            FilterET = CDate(.T_FilterET)
        End If
        Set OldFiles = FilterByEditTime(OldFiles, FilterST, FilterET)
    End If
    
    If Not (.T_FilterSMin = "" And .T_FilterSMax = "") Then
        If .T_FilterSMin = "" Then
            MinSize = 0
        Else
            MinSize = CDbl(.T_FilterSMin)
        End If
        If .T_FilterSMax = "" Then
            MaxSize = 0
        Else
            MaxSize = CDbl(.T_FilterSMax)
        End If
        Set OldFiles = FilterBySize(OldFiles, MinSize, MaxSize)
    End If
    
    'TODO:实现排序文件
    
    Set newFiles = OldFiles
    For i = 1 To CutCollection.Count
        Set newFiles = FindAndReplaceFile(newFiles, CutCollection.Item(i), ReplaceCollection.Item(i))
    Next
    
    Formula = makeFormula(Formula, ActionTypes, ActionContents)
    
    Set newFiles = RenameAdd(newFiles, Formula)
    
    Dim hasDouble As Boolean
    hasDouble = False
    '检查重命名是否会出现重名
    For i = 1 To newFiles.Count
        For ii = 1 To newFiles.Count
            If Not i = ii And newFiles.Item(i) = newFiles.Item(ii) Then
                hasDouble = True
                Exit For
            End If
        Next
        If hasDouble Then Exit For
    Next
    
    If hasDouble Then
        MsgBox ("将产生重名文件，请检查设置")
        Previewed = False
    Else
        Previewed = True
    End If
    
    '放在列表中
    For i = 1 To newFiles.Count
        Set Li = .List_Files.ListItems.Add
        Li.text = ""
        Li.SubItems(1) = getFileName(OldFiles.Item(i))
        Li.SubItems(2) = getFileName(newFiles.Item(i))
        Li.SubItems(3) = OldFiles.Item(i)
    Next
    
    UpdateView

End With
End Sub

