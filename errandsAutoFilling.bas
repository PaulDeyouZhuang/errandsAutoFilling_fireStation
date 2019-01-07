Attribute VB_Name = "Module1"

Dim exeTime, timeInterval As Integer
Dim flag_A, flag_B, flag_C As Boolean


Sub autoGenerateJS()
exeTime = 0
flag_A = True
flag_B = True
flag_C = True

Dim strDate_ori, strMonth, strYear As String
Dim exampleDate As String
Dim monthPath, yearPath As String
  Do
  strDate_ori = InputBox("輸入要填勤務表的日期" & vbNewLine & "格式：年/月/日", "User date", Format(Year(DateAdd("yyyy", -1911, DateAdd("d", 1, Now()))), "000") & "/" & Format(DateAdd("yyyy", -1911, DateAdd("d", 1, Now())), "mm") & "/" & Format(DateAdd("yyyy", -1911, DateAdd("d", 1, Now())), "dd"))
  If strDate_ori = vbNullString Then
  Exit Sub
  
  ElseIf (Not IsDate(strDate_ori)) Then
  MsgBox "錯誤的日期格式"
  
 Else
  monthPath = Month(strDate_ori)
  yearPath = Year(strDate_ori)
  exampleDate = Replace(strDate_ori, "/", "-")

    strDate = Replace(strDate_ori, "/", "")




Dim SOME_PATH As String
SOME_PATH = "\\Dpc32071-101002\汐止資料(無個人資料)\3.勤務表\" & yearPath & "年勤務表\" & monthPath & "月\"



Dim file As String
file = Dir$(SOME_PATH & strDate & "*" & ".xls")
Dim myValue As Variant

If (Len(file) > 0) Then

  Else
  MsgBox "沒有找到這個檔案"

End If

End If
Loop While (Not IsDate(strDate_ori)) Or (Len(file) = 0)

Dim wb1 As Excel.Workbook
Set wb1 = Workbooks.Open(SOME_PATH & file)

wb1.Sheets("Sheet1").Activate
ThisWorkbook.Sheets("工作表1").Activate

Dim NumRows As Integer
NumRows = ThisWorkbook.Sheets("工作表1").Range("A1", Range("A1").End(xlDown)).Rows.Count

Dim arr As New Collection

For y = 1 To NumRows
 arr.Add ThisWorkbook.Sheets("工作表1").Range("A" & y).Text
 ActiveCell.Offset(1, 0).Select
Next


Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
Set oFile = fso.CreateTextFile(ThisWorkbook.Path & "\" & exampleDate & ".txt")



Set dict = CreateObject("Scripting.Dictionary")
Set dict_vacation = CreateObject("Scripting.Dictionary")

For x = 1 To NumRows
If Not dict.Exists(ThisWorkbook.Sheets("工作表1").Range("A" & x).Text) Then
dict.Add ThisWorkbook.Sheets("工作表1").Range("A" & x).Text, ThisWorkbook.Sheets("工作表1").Range("B" & x).Text
End If
ActiveCell.Offset(1, 0).Select
Next


Dim NumRows2 As Integer
NumRows2 = ThisWorkbook.Sheets("工作表1").Range("C1", Range("C1").End(xlDown)).Rows.Count

For x2 = 1 To NumRows2
If Not dict_vacation.Exists(ThisWorkbook.Sheets("工作表1").Range("A" & x2).Text) Then
dict_vacation.Add ThisWorkbook.Sheets("工作表1").Range("A" & x2).Text, ThisWorkbook.Sheets("工作表1").Range("C" & x2).Text
End If
ActiveCell.Offset(1, 0).Select
Next





wb1.Sheets("Sheet1").Activate

Dim errandName As String
errandName = ""

Dim i As Integer
i = 5

Dim errandOrder, errandFightOrder As Integer
errandOrder = 0
errandFightOrder = 0

Dim office_flag As Boolean
office_flag = False


Do While InStr(errandName, "第二備勤") = 0

If errandName = "" Then
GoTo Continuel
End If

If InStr(errandName, "環境清潔值日生") Or InStr(errandName, "環境區域值日生") Or InStr(errandName, "11車司機") Or InStr(errandName, "洗碗值日生") Or InStr(errandName, "後勤盤點") Or InStr(errandName, "常年訓練") Or InStr(errandName, "義消協勤") Then
GoTo Continuel
End If

setTime (4)
If InStr(errandName, "91救護勤務") Or InStr(errandName, "92救護勤務") Or InStr(errandName, "MER支援") Then
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & "listGroupType" & Chr(34) & ").value=2;', " & exeTime & ");"
setTime (4)
Else
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & "listGroupType" & Chr(34) & ").value=1;', " & exeTime & ");"
setTime (9)
oFile.Write "setTimeout('__doPostBack(\'listGroupType\',\'\')', " & exeTime & ");"
setTime (4)
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & "listItemType" & Chr(34) & ").value=" & Chr(34) & "不能派遣" & Chr(34) & ";', " & exeTime & ");"
End If

setTime (4)
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & "txtItemName" & Chr(34) & ").value=" & Chr(34) & errandName & Chr(34) & ";', " & exeTime & ");"
setTime (15)
If office_flag = False Then
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & "controloffice_checkbox_0" & Chr(34) & ").click();', " & exeTime & ");"
office_flag = True
End If
setTime (4)

oFile.Write "setTimeout('document.getElementById(" & Chr(34) & "btnAddItem" & Chr(34) & ").click();', " & exeTime & ");"
setTime (15)

If InStr(errandName, "91救護勤務") Or InStr(errandName, "92救護勤務") Or InStr(errandName, "MER支援") Then
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & "gridGroupFightMan_rdoItemName_" & errandFightOrder & Chr(34) & ").click();', " & exeTime & ");"
errandFightOrder = errandFightOrder + 1
Else
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & "gridGroupWorkMan_rdoItemName_" & errandOrder & Chr(34) & ").click();', " & exeTime & ");"
errandOrder = errandOrder + 1
End If
setTime (4)


Dim flagFirst As Boolean
flagFirst = False
Dim rng As Range
Set rng = Range("E" & i, "AB" & i)

Dim tempSign As String
Dim preSign As String

For Each cell In rng

If flagFirst = False And cell <> vbNullString Then
flagFirst = True
tempSign = cell.Text
preSign = cell.Text

ElseIf flagFirst = True And cell = vbNullString And cell.MergeCells Then
tempSign = preSign

ElseIf flagFirst = True And cell <> vbNullString Then
tempSign = cell.Text


Else
tempSign = ""

End If

If tempSign <> vbNullString Then

For Each s In arr
If InStr(tempSign, s) Then


oFile.Write "setTimeout('document.getElementById(" & Chr(34) & dict(s) & Chr(34) & ").click();', " & exeTime & ");"
Dim timeZone As Integer
timeZone = getTimeZone(Split(cell.Address(columnAbsolute:=False), "$")(0))

setTime (4)
If InStr(errandName, "91救護勤務") Or InStr(errandName, "92救護勤務") Or InStr(errandName, "MER支援") Then
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & "gridGroupFightMan_Button" & timeZone & Chr(34) & ").click();', " & exeTime & ");"
Else
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & "gridGroupWorkMan_Button" & timeZone & Chr(34) & ").click();', " & exeTime & ");"
End If
setTime (4)
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & dict(s) & Chr(34) & ").click();', " & exeTime & ");"
setTime (4)


End If
Next s



End If


preSign = tempSign
Next cell


Continuel:
i = i + 1
errandName = Range("A" & i).Text
Loop


Dim errandName2, errandPre, errandDoublePre As String
errandName2 = ""
errandPre = ""
errandDoublePre = ""

Dim j As Integer
j = 5


Do While InStr(errandName2, "勤    務    輪    流    順") = 0

errandDoublePre = errandPre
errandPre = errandName2


j = j + 1
errandName2 = Range("A" & j).Text
Loop

Dim flag_First As Boolean
flag_First = False
Dim rng1 As Range
Set rng1 = Range("A" & (j - 2), "AB" & (j - 1))

Dim temp_Sign As String
Dim pre_Sign As String

For Each cell In rng1

If flag_First = False And cell <> vbNullString Then
flag_First = True
temp_Sign = cell.Text
pre_Sign = cell.Text

ElseIf flag_First = True And cell = vbNullString Then
GoTo Continuel2


Else
pre_Sign = temp_Sign
temp_Sign = cell.Text
End If



setTime (4)
If InStr(pre_Sign, "輪休") Then
For Each s In arr
If InStr(temp_Sign, s) And InStr(s, "60") <> 1 Then
checkDirectors (s)
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & dict_vacation(s) & Chr(34) & ").value=1;', " & exeTime & ");"
End If
Next s


ElseIf InStr(pre_Sign, "請休") Then
For Each s In arr
If InStr(temp_Sign, s) And InStr(s, "60") <> 1 Then
checkDirectors (s)
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & dict_vacation(s) & Chr(34) & ").value=2;', " & exeTime & ");"
End If
Next s

ElseIf InStr(pre_Sign, "超休") Then
For Each s In arr
If InStr(temp_Sign, s) And InStr(s, "60") <> 1 Then
checkDirectors (s)
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & dict_vacation(s) & Chr(34) & ").value=3;', " & exeTime & ");"
End If
Next s

ElseIf InStr(pre_Sign, "傷") Then
For Each s In arr
If InStr(temp_Sign, s) And InStr(s, "60") <> 1 Then
checkDirectors (s)
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & dict_vacation(s) & Chr(34) & ").value=6;', " & exeTime & ");"
End If
Next s

ElseIf InStr(pre_Sign, "婚") Then
For Each s In arr
If InStr(temp_Sign, s) And InStr(s, "60") <> 1 Then
checkDirectors (s)
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & dict_vacation(s) & Chr(34) & ").value=7;', " & exeTime & ");"
End If
Next s

ElseIf InStr(pre_Sign, "產") Or InStr(pre_Sign, "胎") Then
For Each s In arr
If InStr(temp_Sign, s) And InStr(s, "60") <> 1 Then
checkDirectors (s)
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & dict_vacation(s) & Chr(34) & ").value=8;', " & exeTime & ");"
End If
Next s

ElseIf InStr(pre_Sign, "喪") Then
For Each s In arr
If InStr(temp_Sign, s) And InStr(s, "60") <> 1 Then
checkDirectors (s)
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & dict_vacation(s) & Chr(34) & ").value=9;', " & exeTime & ");"
End If
Next s

ElseIf InStr(pre_Sign, "照顧") Then
For Each s In arr
If InStr(temp_Sign, s) And InStr(s, "60") <> 1 Then
checkDirectors (s)
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & dict_vacation(s) & Chr(34) & ").value=11;', " & exeTime & ");"
End If
Next s

ElseIf InStr(pre_Sign, "訓") Or InStr(pre_Sign, "支援") Then
For Each s In arr
If InStr(temp_Sign, s) And InStr(s, "60") <> 1 Then
checkDirectors (s)
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & dict_vacation(s) & Chr(34) & ").value=12;', " & exeTime & ");"
End If
Next s

ElseIf InStr(pre_Sign, "外宿") Then
For Each s In arr
If InStr(temp_Sign, s) And InStr(s, "60") <> 1 Then
checkDirectors (s)
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & dict_vacation(s) & Chr(34) & ").value=13;', " & exeTime & ");"
End If
Next s

ElseIf InStr(pre_Sign, "日休") Then
For Each s In arr
If InStr(temp_Sign, s) And InStr(s, "60") <> 1 Then
checkDirectors (s)
oFile.Write "setTimeout('document.getElementById(" & Chr(34) & dict_vacation(s) & Chr(34) & ").value=14;', " & exeTime & ");"
End If
Next s

End If
setTime (4)

Continuel2:
Next cell

setTime (4)


oFile.Write "setTimeout('document.getElementById(" & Chr(34) & "btnVacationSave" & Chr(34) & ").click();', " & exeTime & ");"

setTime (6)



oFile.Write "setTimeout('document.getElementById(" & Chr(34) & "Button26" & Chr(34) & ").click();', " & exeTime & ");"




oFile.Close
Set fso = Nothing
Set oFile = Nothing


ThisWorkbook.Sheets("說明").Activate
End Sub

Public Function getTimeZone(ByVal s As Variant) As Integer

Select Case s
     Case "E"
       getTimeZone = 8
     Case "F"
       getTimeZone = 9
     Case "G"
       getTimeZone = 10
     Case "H"
       getTimeZone = 11
    Case "I"
       getTimeZone = 12
     Case "J"
       getTimeZone = 13
     Case "K"
       getTimeZone = 14
     Case "L"
       getTimeZone = 15
     Case "M"
       getTimeZone = 16
     Case "N"
       getTimeZone = 17
     Case "O"
       getTimeZone = 18
     Case "P"
       getTimeZone = 19
     Case "Q"
       getTimeZone = 20
     Case "R"
       getTimeZone = 21
     Case "S"
       getTimeZone = 22
     Case "T"
       getTimeZone = 23
     Case "U"
       getTimeZone = 0
     Case "V"
       getTimeZone = 1
     Case "W"
       getTimeZone = 2
     Case "X"
       getTimeZone = 3
     Case "Y"
       getTimeZone = 4
     Case "Z"
       getTimeZone = 5
     Case "AA"
       getTimeZone = 6
     Case "AB"
       getTimeZone = 7
     Case Else
       getTimeZone = 9999
  End Select

End Function

Public Function setTime(x As Integer)

exeTime = exeTime + x * 1000

End Function

Public Function checkDirectors(s As String)

If s = "A" Then
flag_A = False
End If

If s = "B" Then
flag_B = False
End If

If s = "C" Then
flag_C = False
End If

End Function


