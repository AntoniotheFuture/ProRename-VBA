Attribute VB_Name = "m1"

Global CutFind As String
Global CutReplace As String

Global SortFind As String
Global SortCondition As String

Global NewDiffs As Collection

'格式化当前时间
Function getCurrentTime(timeFormat As String) As String
    Dim DT As Date
    DT = Time()
    getCurrentTime = Format(DT, timeFormat)
End Function


'获取时间格式化
Function getTimeFormats() As Collection
    Dim all As New Collection
    all.Add "YYYYMMDD"
    all.Add "YYYY-MM-DD"
    all.Add "YMD"
    all.Add "YYYY年MM月DD日"
    all.Add "YYYY年M月D日"
    all.Add "YYYYMMDD hhmmss"
    all.Add "YYYY-MM-DD  hhmmss"
    all.Add "hh:mm:ss"
    all.Add "MMDD"
    all.Add "MM月DD日"
    all.Add "MD"
    all.Add "M月D日"
    Set getTimeFormats = all
End Function


'按位数获取随机数
Function getRamdom(letters As Boolean, nums As Boolean, D As Integer) As String
allletters = "0 1 2 3 4 5 6 7 8 9 A B C D E F G H I J K L M N O P Q R S T U V W X Y Z a b c d e f g h i j k l m n o p q r s t u v w x y z"
arr = Split(allletters)
Dim text As String
Dim d2 As Integer, upperbound As Integer, lowerbound As Integer
If letters And nums Then
    lowerbound = 0
    upperbound = 61
ElseIf Not letters And nums Then
    lowerbound = 0
    upperbound = 9
Else
    lowerbound = 10
    upperbound = 61
End If
For i = 1 To D
    d2 = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
    text = text & arr(d2)
Next
getRamdom = text
End Function




