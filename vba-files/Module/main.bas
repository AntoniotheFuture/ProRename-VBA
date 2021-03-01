Attribute VB_Name = "main"
'判断字符串内是否含有不能用的
Function checkAVB(s As String) As Boolean
    arr = Array("\", "/", ":", "*", "<", ">", "|", Chr(34))
    For Each a In arr
        If InStr(s, a) > 0 Then
            checkAVB = False
            Exit Function
        End If
    Next
    checkAVB = True
End Function

'从文件夹中获取文件
Function getherFiles(Folders As Collection) As Collection
Dim path1 As String
Dim Files As New Collection
For Each path In Folders
    path1 = path
    arr = GetFiles(path1)
    For i = 0 To UBound(arr)
        Files.Add arr(i)
    Next
Next
Set getherFiles = Files
End Function

'按正则表达式筛选文件,Exclude 代表是白名单模式
Function FilterByReg(Files As Collection, Reg As String, Exclude As Boolean) As Collection
Dim RegExp As Object
Dim FilePath As String, FileName As String
Dim newFiles As New Collection
Set RegExp = CreateObject("vbscript.regexp")
RegExp.Global = True
RegExp.ignorecase = False
RegExp.Pattern = Reg
For Each FilePath1 In Files
    FilePath = FilePath1
    FileName = getFileName(FilePath, True)
    If Exclude Then
        If RegExp.test(FileName) Then
            newFiles.Add FilePath
        End If
    Else
        If Not RegExp.test(FileName) Then
            newFiles.Add FilePath
        End If
    End If
Next

Set FilterByReg = newFiles
End Function

'按文件修改时间筛选
Function FilterByEditTime(Files As Collection, ST As Date, ET As Date) As Collection
Dim FilePath As String
Dim FileDate As Date
Dim newFiles As New Collection
For Each FilePath1 In Files
    FilePath = FilePath1
    FileDate = FileDateTime(FilePath)
    If FileDate >= ST And FileDate <= ET Then newFiles.Add FilePath
Next
Set FilterByEditTime = newFiles
End Function

'按文件大小筛选(KB)，0 max 为无限大
Function FilterBySize(Files As Collection, MinL As Double, MaxL As Double) As Collection
Dim FilePath As String
Dim FileLenKB As Double
Dim newFiles As New Collection
For Each FilePath1 In Files
    FilePath = FilePath1
    FileLenKB = FileLen(FilePath) / 1024
    If FileLenKB >= MinL Then
        If MaxL = 0 Then
            newFiles.Add FilePath
        ElseIf MaxL > 0 And FileLenKB <= MaxL Then
            newFiles.Add FilePath
        End If
    End If
Next
Set FilterBySize = newFiles
End Function

'排序文件
Function SortFile(Files As Collection, SortType As String) As Collection
'TODO
Select Case SortType
    Case "EditTime"
    
    Case "CreateTime"
    Case "FileSize"
    Case "OldNum"
    Case "Letters"
End Select

End Function

'按正则表达式截去和替换
Function FindAndReplaceFile(Files As Collection, Find As String, Replace As String) As Collection
Dim RegExp As Object
Dim FilePath As String, FileName As String, FileName2 As String, Tail As String
Dim newFiles As New Collection
Dim NewPath As String
Set RegExp = CreateObject("vbscript.regexp")
RegExp.Global = True
RegExp.ignorecase = False
RegExp.Pattern = Find
For Each FilePath1 In Files
    FilePath = FilePath1
    FileName = getFileName(FilePath, False)
    FileName2 = getFileName(FilePath, True)
    Tail = Right(FileName2, Len(FileName2) - Len(FileName) - 1)
    FileName = RegExp.Replace(FileName, Replace)
    NewPath = Left(FilePath, Len(FilePath) - Len(FileName2)) & FileName & "." & Tail
    newFiles.Add NewPath
Next
Set FindAndReplaceFile = newFiles
End Function

'生成最终表达式
Function makeFormula(Formula As String, types As Collection, Contents As Collection) As String
Dim s1 As String
Dim newFormula As String
Dim id As Integer
arr = Split(Formula, "#")
For Each s In arr
    s1 = s
    If Not s1 = "" Then
        If Not s1 = "@" Then
            id = Int(s1)
            newFormula = newFormula & "|" & types.Item(id) & ";" & Contents.Item(id)
        Else
            newFormula = newFormula & "|?"
        End If
    End If
Next
makeFormula = newFormula
End Function

'按增加模式执行重命名 ,l:字符,i:序号,t:时间,r:随机
Function RenameAdd(Files As Collection, Formula As String) As Collection
Dim id As Integer '初始ID
Dim FilePath As String
Dim FileName As String
Dim newFiles As New Collection
Dim t As String, s As String
Dim newFileName As String
Dim newFilePath As String
Dim Content As String
Dim DiffDB As New Collection
Dim Folder As String

arr = Split(Formula, "|")
id = 0
Set DiffDB = getDiffDB
Set NewDiffs = New Collection
Folder = ""

For Each FP In Files
    FilePath = FP
    FileName = getFileName(FilePath, False)
    FileName2 = getFileName(FilePath, True)
    Tail = Right(FileName2, Len(FileName2) - Len(FileName))
    newFolder = Left(FilePath, Len(FilePath) - Len(FileName2))
    '按文件夹内部序号
    If Not Folder = newFolder Then
        Folder = newFolder
        id = 0
    End If
    newFileName = ""
    '计算最终文件名
    For Each s1 In arr
        s = s1
        If Not s = "" Then
        If s = "?" Then
            newFileName = newFileName & FileName
        Else
            actionArr = Split(s, ";")
            Select Case actionArr(0)
                Case "字符"
                    newFileName = newFileName & actionArr(1)
                Case "序号"
                    '是否固定位数
                    If actionArr(1) = "" Then
                        formatStr = ""
                    Else
                        formatStr = String(Int(actionArr(1)), 0)
                    End If
                    r = 1
                    If actionArr(3) = "1" Then r = -1
                    StartN = 1
                    If Not actionArr(2) = "" Then StartN = Int(actionArr(2))
                    newFileName = newFileName & Format(id * r + StartN, formatStr)
                Case "时间"
                    'TODO:实现文件修改日期
                    CurrentTime = Time()
                    newFileName = newFileName & Format(CurrentTime, actionArr(2))
                Case "随机"
                    If actionArr(3) = "1" Then
                        '重试100次
                        Dim trys As Integer
                        trys = 0
                        Do
                            ramdomText = getRamdom(actionArr(1) = "1", actionArr(2) = "1", Int(actionArr(4)))
                            hit = False
                            For Each str1 In DiffDB
                                If str1 = ramdomText Then
                                    hit = True
                                'hit = (str = ramdomText)
                                    Exit For
                                End If
                            Next
                            If Not hit Then
                                DiffDB.Add ramdomText
                                NewDiffs.Add ramdomText
                            End If
                            trys = trys + 1
                        Loop While trys < 100 And hit
                        newFileName = newFileName & ramdomText
                    Else
                        newFileName = newFileName & getRamdom(actionArr(1) = "1", actionArr(2) = "1", Int(actionArr(4)))
                    End If
            End Select
        End If
        End If
    Next
    newFilePath = Left(FilePath, Len(FilePath) - Len(FileName2)) & newFileName & Tail
    newFiles.Add newFilePath
    id = id + 1
Next
Set RenameAdd = newFiles

End Function

'获取已经使用的查重库
Function getDiffDB() As Collection
Dim WS As Worksheet
Dim maxRow As Integer
Dim str As String
Dim c As New Collection
Set WS = ThisWorkbook.Sheets("DiffDB")
maxRow = WS.UsedRange.Rows.Count
For x = 1 To maxRow
    str = WS.Cells(x, 1).text
    arr = Split(str, ",")
    For Each s In arr
        c.Add s
    Next
Next
Set getDiffDB = c
End Function

'向查重库中添加新的
Function AddToDiffDB(newDiff As Collection) As Boolean
Dim WS As Worksheet
Dim maxRow As Integer
Dim str As String
Dim cb As String
Dim c As New Collection
Set WS = ThisWorkbook.Sheets("DiffDB")
maxRow = WS.UsedRange.Rows.Count
For Each s In newDiff
    cb = cb + s + ","
Next
cb = Left(cb, Len(cb) - 1)
WS.Cells(maxRow + 1, 1).Value = cb
AddToDiffDB = True
End Function

'检查是否会生成重复，无重复返回true
Function CheckDouble(newFiles As Collection) As Boolean
For Each s In newFiles
    For Each s2 In newFiles
        If s = s2 Then
            CheckDouble = False
            Exit Function
        End If
    Next
Next
CheckDouble = True
End Function

'获取某文件夹下的文件
Function GetFiles(path$, Optional Fullname As Boolean = True, Optional SubFolders As Boolean = False)
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    Dim Fso As Object, Folder As Object
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set Folder = Fso.GetFolder(path)
    i = 1
        Call aGetFile(Folder, dic, SubFolders)
    If Fullname Then
        GetFiles = dic.keys '返回文件名
    Else
        GetFiles = dic.items '返回完整路径带文件名
    End If
    Set Folder = Nothing
    Set Fso = Nothing
End Function

Sub aGetFile(ByVal Folder As Object, dic, Optional SubFolders As Boolean)
    Dim SubFolder As Object
    Dim File As Object
    For Each File In Folder.Files '遍历文件
        dic.Add File.path & "" & FileName, File.Name
    Next
    If SubFolders Then
        For Each SubFolder In Folder.SubFolders
            Call aGetFile(SubFolder, dic) '递归调用子文件夹
        Next
    End If
End Sub

'截取文件名称,是否包括扩展名
Function getFileName(FullPath As String, Optional Tail As Boolean = True) As String
    If InStr(FullPath, ".") = 0 Then Tail = True
    Arr1 = Split(FullPath, "\")
    FileName = Arr1(UBound(Arr1))
    If Tail Then
        getFileName = FileName
        Exit Function
    End If
    Arr2 = Split(FileName, ".")
    FileType = Arr2(UBound(Arr2))
    FFileName = Left(FileName, Len(FileName) - Len(FileType) - 1)
    getFileName = FFileName
End Function





