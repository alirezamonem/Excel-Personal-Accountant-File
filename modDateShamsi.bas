Attribute VB_Name = "modDateShamsi"

'Çíä ãÊÛííÑ ãŞÏÇÑ ßáíß ÔÏå ÏÑ İÑã ÊŞæíã Ñæ ÈÕæÑÊ ÓÑÇÓÑí ÏÑ ÎæÏÔ ĞÎíÑå ãíßäå
Public STRDATE As String
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'                            Ç'   ãÇæá ÇÕáÇÍ ÔÏå ÌäÇÈ ÂÒÇÏí ÊæÓØ ÑÓæá ÛáÇãí Èå ÊÇÑíÎ 1390/10/15
'                       1- ÊÚÑíİ ßäíÏ Number(Long) ÇÓÊ ÑÇ ÈÕæÑÊ Date Ç'   İíáÏåÇíí ßå äæÚ ÂäåÇ
'                            2- Çíä İíáÏåÇ ÑÇ ÈÕæÑÊ 0000/00/00 ÊäÙíã ßäíÏ InputMask Ç'   ÎÇÕíÊ
'                Ç'   ÈÏáíá 8 ÑŞãí ÏÑ äÙÑ ÑİÊä İíáÏ ÊÇÑíÎ ¡ Çíä ÊæÇÈÚ ÊÇ ÓÇá 9999 ßÇÑÇíí ÏÇÑÏ
'                             ÊÇÑíÎ ÌÇÑí ÓíÓÊã ÑÇ Èå åÌÑí ÔãÓí ÊÈÏíá ãí ßäÏ Shamsi() Ç'   ÊÇÈÚ
'                           ÈßÇÑ ÈÈÑíÏ Now() ÑÇ ãí ÊæÇäíÏ ÏÑ ÒÇÑÔÇÊ ÈÌÇí ÊÇÈÚ Dat() Ç'   ÊÇÈÚ
'             Èå ßÇÑ ÈÈÑíÏ / ÑÇ ãíÊæÇäíÏ ÌåÊ ÏÑÌ ÊÇÑíÎ ÏÑ ÌÏÇæá Èå åãÑÇå ShamsiDat() Ç'   ÊÇÈÚ
'                    Ç'   ÈÑÇí ÌáæíÑí ÇÒ æÑæÏ ÊÇÑíÎ ÛáØ Èå ÏÑæä íß İíáÏ ÈÊÑÊíÈ ÒíÑ Úãá ãíßäíÏ
'   İíáÏ ãæÑÏ äÙÑ ÈßÇÑ ÈÈÑíÏ ValidationRule ÑÇ ÏÑ ÎÇÕíÊ ValidDate([Field Name])=True Ç'   ÊÇÈÚ
'/////////////////////////////////////////////////////////////////////////////////////////////

Public Static Function Shamsi() As Long
'Çíä ÊÇÈÚ ÊÇÑíÎ ÌÇÑí ÓíÓÊã ÑÇ Èå ÊÇÑíÎ åÌÑí ÔãÓí ÊÈÏíá ãí ßäÏ
Dim Shamsi_Mabna As Long
Dim Miladi_mabna As Date
Dim Dif As Long
'ÏÑ ÇíäÌÇ 78/10/11 ÈÇ 2000/01/01 ãÚÇÏá ŞÑÇÑÏÇÏå ÔÏå
Shamsi_Mabna = 13781011
Miladi_mabna = #1/1/2000#
Dif = DateDiff("d", Miladi_mabna, Date)
If Dif < 0 Then
MsgBox "ÊÇÑíÎ ÌÇÑí ÓíÓÊã ÔãÇ äÇÏÑÓÊ ÇÓÊ , ÂäÑÇ ÇÕáÇÍ ßäíÏ."
Else
Shamsi = AddDay(Shamsi_Mabna, Dif)
End If
End Function

Public Function dat() As String
' ÈßÇÑ ÈÈÑíÏ Now() ÑÇ ãí ÊæÇäíÏ ÏÑ ÒÇÑÔÇÊ ÈÌÇí ÊÇÈÚ Dat() ÊÇÈÚ
dat = DayWeek(Shamsi) & " - " & Slash(Shamsi)
End Function

Public Function Slash(F_Date As Variant) As String
' Çíä ÊÇÈÚ íß ÊÇÑíÎ ÑÇ ÏÑíÇİÊ æ ÈÕæÑÊ íß ÑÔÊå 10 ÑŞãí ÔÇãá / æ åÇÑ ÑŞã ÈÑÇí ÓÇá ÈÇÒãíÑÏÇäÏ
F_Date = Replace(F_Date, "/", "")
Dim a As Long
a = CLng(F_Date)
Slash = Format(IL(a), "0000") & "/" & Format(ay(a), "00") & "/" & Format(Guon(a), "00")
End Function

Function ValidDate(F_Date As Variant) As Boolean
' Çíä ÊÇÈÚ ÇÚÊÈÇÑ íß ÚÏÏ æÑæÏí ÑÇ ÇÒ äÙÑ ÊÇÑíÎ åÌÑí ÔãÓí ÈÑÑÓí ãí ßäÏ
' ÑÇ ÈÑãí ÑÏÇäÏ False æÇÑ äÇãÚÊÈÑ ÈÇÔÏ True ÇÑ ÊÇÑíÎ ãÚÊÈÑ ÈÇÔÏ
On Error GoTo Err_ValidDate
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Dim M, s, R As Byte
F_Date = Replace(F_Date, "/", "")
R = Guon(CLng(F_Date))
M = ay(CLng(F_Date))
s = IL(CLng(F_Date))
If F_Date < 10000101 Then Exit Function
If M > 12 Or M = 0 Or R = 0 Then Exit Function
If R > ayDays(s, M) Then Exit Function
ValidDate = True

Exit_ValidDate:
    On Error Resume Next
    Exit Function
Err_ValidDate:
    Select Case err.Number
        Case 0
            Resume Exit_ValidDate:
        Case 94
            ValidDate = True
        Case Else
            MsgBox err.Number & " " & err.Description, vbExclamation, "Error in module Module2 - function ValidDate"
            Resume Exit_ValidDate:
    End Select
End Function

Public Function AddDay(ByVal F_Date As Variant, ByVal add As Long) As Long
'Çíä ÊÇÈÚ ÊÚÏÇÏ ÑæÒ ÏáÎæÇå ÑÇ Èå ÊÇÑíÎ ÑæÒ ÇÖÇİå ãíßäÏ
On Error GoTo Err_AddDay
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
F_Date = Replace(F_Date, "/", "")
Dim k, M, R, Days As Byte
Dim s As Integer
R = Guon(CLng(F_Date))
M = ay(CLng(F_Date))
s = IL(CLng(F_Date))
k = Kabiseh(s)
'ÊÈÏíá ÑæÒ Èå ÚÏÏ 1 ÌåÊ ÇÏÇãå ãÍÇÓÈÇÊ æ íÇ ÇÊãÇã ãÍÇÓÈå
Days = ayDays(s, M)
If add > Days - R Then
add = add - (Days - R + 1)
R = 1
If M < 12 Then
M = M + 1
Else
M = 1
s = s + 1
End If
Else
R = R + add
add = 0
End If
While add > 0
k = Kabiseh(s) 'ßÈíÓå: 1 æ ÛíÑ ßÈíÓå: 0
Days = ayDays(s, M) 'ÊÚÏÇÏ ÑæÒåÇí ãÇå İÚáí
Select Case add
Case Is < Days
'ÇÑ ÊÚÏÇÏ ÑæÒåÇí ÇİÒæÏäí ßãÊÑ ÇÒ íß ãÇå ÈÇÔÏ
R = R + add
add = 0
Case Days To IIf(k = 0, 365, 366) - 1
'ÇÑ ÊÚÏÇÏ ÑæÒåÇí ÇİÒæÏäí ÈíÔÊÑ ÇÒ íß ãÇå æ ßãÊÑ ÇÒ íß ÓÇá ÈÇÔÏ
add = add - Days
If M < 12 Then
M = M + 1
Else
s = s + 1
M = 1
End If
Case Else
'ÇÑ ÊÚÏÇÏ ÑæÒåÇí ÇİÒæÏäí ÈíÔÊÑ ÇÒ íß ÓÇá ÈÇÔÏ
s = s + 1
add = add - IIf(k = 0, 365, 366)
End Select
Wend
AddDay = CLng(s & Format(M, "00") & Format(R, "00"))

Exit_AddDay:
    On Error Resume Next
    Exit Function
Err_AddDay:
    Select Case err.Number
        Case 0
            Resume Exit_AddDay:
        Case 94
            AddDay = 0
        Case Else
            MsgBox err.Number & " " & err.Description, vbExclamation, "Error in module Module2 - function AddDay"
            Resume Exit_AddDay:
    End Select
End Function

Function SubDay(ByVal F_Date As Variant, ByVal Subtract As Long) As Long
'Èå ÊÚÏÇÏ ÑæÒ ãÚíäí ÇÒ íß ÊÇÑíÎ ßã ßÑÏå æ ÊÇÑíÎ ÍÇÕáå ÑÇ ÇÑÇÆå ãíßäÏ
On Error GoTo Err_SubDay
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
F_Date = Replace(F_Date, "/", "")
Dim k, M, s, R, Days As Byte
R = Guon(CLng(F_Date))
M = ay((CLng(F_Date)))
s = IL((CLng(F_Date)))
k = Kabiseh(s)
'ÊÈÏíá ÑæÒ Èå ÚÏÏ 1 ÌåÊ ÇÏÇãå ãÍÇÓÈÇÊ æ íÇ ÇÊãÇã ãÍÇÓÈå
If Subtract >= R - 1 Then
Subtract = Subtract - (R - 1)
R = 1
Else
R = R - Subtract
Subtract = 0
End If
While Subtract > 0
k = Kabiseh(s - 1) 'ßÈíÓå: 1 æ ÛíÑ ßÈíÓå: 0
Days = ayDays(IIf(M >= 2, s, s - 1), IIf(M >= 2, M - 1, 12)) 'ÊÚÏÇÏ ÑæÒåÇí ãÇå ŞÈáí
Select Case Subtract
Case Is < Days
'ÇÑ ÊÚÏÇÏ ÑæÒåÇí ßÇåÔ ßãÊÑ ÇÒ íß ãÇå ÈÇÔÏ
R = Days - Subtract + 1
Subtract = 0
If M >= 2 Then
M = M - 1
Else
s = s - 1
M = 12
End If
Case Days To IIf(k = 0, 365, 366) - 1
'ÇÑ ÊÚÏÇÏ ÑæÒåÇí ßÇåÔ ÈíÔÊÑ ÇÒ íß ãÇå æ ßãÊÑ ÇÒ íß ÓÇá ÈÇÔÏ
Subtract = Subtract - Days
If M >= 2 Then
M = M - 1
Else
s = s - 1
M = 12
End If
Case Else
'ÇÑ ÊÚÏÇÏ ÑæÒåÇí ßÇåÔ ÈíÔÊÑ ÇÒ íß ÓÇá ÈÇÔÏ
s = s - 1
Subtract = Subtract - IIf(k = 0, 365, 366)
End Select
Wend
SubDay = (s * 10000) + (M * 100) + (R)

Exit_SubDay:
    On Error Resume Next
    Exit Function
Err_SubDay:
    Select Case err.Number
        Case 0
            Resume Exit_SubDay:
        Case 94
            SubDay = 0
        Case Else
            MsgBox err.Number & " " & err.Description, vbExclamation, "Error in module Module2 - function SubDay"
            Resume Exit_SubDay:
    End Select
End Function

Public Function DayWeekNo(F_Date As Variant) As Byte
'Çíä ÊÇÈÚ íß ÊÇÑíÎ ÑÇ ÏÑíÇİÊ ßÑÏå æ ÔãÇÑå ÑæÒ åİÊå ÑÇ ãÔÎÕ ãí ßäÏ
'ÇÑ ÔäÈå ÈÇÔÏ ÚÏÏ 0
'ÇÑ 1ÔäÈå ÈÇÔÏ ÚÏÏ 1
'......
'ÇÑ ÌãÚå ÈÇÔÏ ÚÏÏ 6
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
F_Date = Replace(F_Date, "/", "")
Dim day As String
Dim Shmsi_Mabna As Long
Dim Dif As Long
'ãÈäÇ 80/10/11
Shmsi_Mabna = 13801011
Dif = Diff(Shmsi_Mabna, CLng(F_Date))
If Shmsi_Mabna > CLng(F_Date) Then
Dif = -Dif
End If
'ÈÇ ÊæÌå Èå Çíäßå 80/10/11 3ÔäÈå ÇÓÊ ãÍÇÓÈå ãíÔæÏ day ãÊÛíÑ
day = (Dif + 3) Mod 7
If day < 0 Then
DayWeekNo = day + 7
Else
DayWeekNo = day
End If
End Function

Public Function DayWeek(F_Date As Variant) As String
'Çíä ÊÇÈÚ íß ÊÇÑíÎ ÑÇ ÏÑíÇİÊ ßÑÏå æ ãÔÎÕ ãí ßäÏ å ÑæÒí ÇÒ åİÊå ÇÓÊ
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Dim a As String
Dim n As Byte
n = DayWeekNo(F_Date)
Select Case n
Case 0
a = "ÔäÈå"
Case 1
a = "íß ÔäÈå"
Case 2
a = "Ïæ ÔäÈå"
Case 3
a = "Óå ÔäÈå"
Case 4
a = "åÇÑ ÔäÈå"
Case 5
a = "äÌ ÔäÈå"
Case 6
a = "ÌãÚå"
End Select
DayWeek = a
End Function

Public Function Diff(ByVal date1 As Variant, ByVal Date2 As Variant) As Long
'Çíä ÊÇÈÚ ÊÚÏÇÏ ÑæÒåÇí Èíä Ïæ ÊÇÑíÎ ÑÇ ÇÑÇÆå ãí ßäÏ
On Error GoTo Err_Diff
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
date1 = Replace(date1, "/", "")
Date2 = Replace(Date2, "/", "")
Dim Tmp As Long
Dim s1, M1, R1, s2, M2, R2 As Integer
Dim Sumation As Single
Dim Flag As Boolean
Flag = False
If CLng(date1) = 0 Or IsNull(CLng(date1)) = True Or CLng(Date2) = 0 Or IsNull(CLng(Date2)) = True Then
Diff = 0
Exit Function
End If
'ÇÑ ÊÇÑíÎ ÔÑæÚ ÇÒ ÊÇÑíÎ ÇíÇä ÈÒÑÊÑ ÈÇÔÏ ÂäåÇ ãæŞÊÇ ÌÇÈÌÇ ãí ÔæäÏ
If CLng(date1) > CLng(Date2) Then
Flag = True
Tmp = CLng(date1)
date1 = CLng(Date2)
Date2 = Tmp
End If
R1 = Guon(CLng(date1))
M1 = ay(CLng(date1))
s1 = IL(CLng(date1))
R2 = Guon(CLng(Date2))
M2 = ay(CLng(Date2))
s2 = IL(CLng(Date2))
Sumation = 0
Do While s1 < s2 - 1 Or (s1 = s2 - 1 And (M1 < M2 Or (M1 = M2 And R1 <= R2)))
'ÇÑ íß ÓÇá íÇ ÈíÔÊÑ ÇÎÊáÇİ ÈæÏ
If Kabiseh((s1)) = 1 Then
If M1 = 12 And R1 = 30 Then
Sumation = Sumation + 365
R1 = 29
Else
Sumation = Sumation + 366
End If
Else
Sumation = Sumation + 365
End If
s1 = s1 + 1
Loop
Do While s1 < s2 Or M1 < M2 - 1 Or (M1 = M2 - 1 And R1 < R2)
'ÇÑ íß ãÇå íÇ ÈíÔÊÑ ÇÎÊáÇİ ÈæÏ
Select Case M1
Case 1 To 6
If M1 = 6 And R1 = 31 Then
Sumation = Sumation + 30
R1 = 30
Else
Sumation = Sumation + 31
End If
M1 = M1 + 1
Case 7 To 11
If M1 = 11 And R1 = 30 And Kabiseh(s1) = 0 Then
Sumation = Sumation + 29
R1 = 29
Else
Sumation = Sumation + 30
End If
M1 = M1 + 1
Case 12
If Kabiseh(s1) = 1 Then
Sumation = Sumation + 30
Else
Sumation = Sumation + 29
End If
s1 = s1 + 1
M1 = 1
End Select
Loop
If M1 = M2 Then
Sumation = Sumation + (R2 - R1)
Else
Select Case M1
Case 1 To 6
Sumation = Sumation + (31 - R1) + R2
Case 7 To 11
Sumation = Sumation + (30 - R1) + R2
Case 12
If Kabiseh(s1) = 1 Then
Sumation = Sumation + (30 - R1) + R2
Else
Sumation = Sumation + (29 - R1) + R2
End If
End Select
End If
If Flag = True Then
Sumation = -Sumation
End If
Diff = Sumation

Exit_Diff:
    On Error Resume Next
    Exit Function
Err_Diff:
    Select Case err.Number
        Case 0
            Resume Exit_Diff:
        Case 94
            Diff = 0
        Case Else
            MsgBox err.Number & " " & err.Description, vbExclamation, "Error in module Module2 - function Diff"
            Resume Exit_Diff:
    End Select
End Function

Function ayName(ByVal ay_no As Byte) As String
'Çíä ÊÇÈÚ íß ÊÇÑíÎ ÑÇ ÏÑíÇİÊ ßÑÏå æ ãÔÎÕ ãí ßäÏ å ãÇåí ÇÒ ÓÇá ÇÓÊ
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Select Case ay_no
Case 1
ayName = "İÑæÑÏíä"
Case 2
ayName = "ÇÑÏíÈåÔÊ"
Case 3
ayName = "ÎÑÏÇÏ"
Case 4
ayName = "ÊíÑ"
Case 5
ayName = "ãÑÏÇÏ"
Case 6
ayName = "ÔåÑíæÑ"
Case 7
ayName = "ãåÑ"
Case 8
ayName = "ÂÈÇä"
Case 9
ayName = "ÂĞÑ"
Case 10
ayName = "Ïí"
Case 11
ayName = "Èåãä"
Case 12
ayName = "ÇÓİäÏ"
End Select
End Function

Function ayDays(ByVal IL As Integer, ByVal ay As Byte) As Byte
'Çíä ÊÇÈÚ ÊÚÏÇÏ ÑæÒåÇí íß ãÇå ÑÇ ÈÑãí ÑÏÇäÏ
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Select Case ay
Case 1 To 6
ayDays = 31
Case 7 To 11
ayDays = 30
Case 12
If Kabiseh(IL) = 1 Then
ayDays = 30
Else
ayDays = 29
End If
End Select
End Function

Function Make_Date(ByVal F_Date As Long) As String
'íß ÊÇÑíÎ ÑÇ ÈÕæÑÊ íß ÑÔÊå 10 ÑŞãí ÈÇ ĞßÑ åÇÑ
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>ÑŞã ÈÑÇí ÓÇá ÇÑÇÆå ãí ßäÏ
Dim d As String
d = Trim(Str(F_Date))
If IsNull(F_Date) = True Or F_Date = 0 Then
Make_Date = ""
Else
Make_Date = Mid(d, 1, 4) & "/" & Mid(d, 5, 2) & "/" & Mid(d, 7, 2)
End If
End Function

Function ILay(ByVal F_Date As Long) As Long
'ÔÔ ÑŞã Çæá ÊÇÑíÎ ßå ãÚÑİ ÓÇá æ ãÇå ÇÓÊ ÑÇ ÈÑãí ÑÏÇäÏ
ILay = Val(Left$(F_Date, 6))
End Function

Public Function Guon(F_Date As Long) As Byte
'Çíä ÊÇÈÚ ÚÏÏ ãÑÈæØ Èå ÑæÒ íß ÊÇÑíÎ ÑÇ ÈÑãÑÏÇäÏ
Guon = F_Date Mod 100
End Function

Function ay(F_Date As Long) As Byte
'Çíä ÊÇÈÚ ÚÏÏ ãÑÈæØ Èå ãÇå íß ÊÇÑíÎ ÑÇ ÈÑãÑÏÇäÏ
ay = Int((F_Date Mod 10000) / 100)
End Function

Public Function IL(F_Date As Long) As Integer
'Çíä ÊÇÈÚ ÚÏÏ ãÑÈæØ Èå ÓÇá íß ÊÇÑíÎ ÑÇ ÈÑãÑÏÇäÏ
IL = Int(F_Date / 10000)
End Function

Public Function Kabiseh(ByVal OnlyIL As Variant) As Byte
'æÑæÏí ÊÇÈÚ ÚÏÏ ÏæÑŞãí ÇÓÊ
'Çíä ÊÇÈÚ ßÈíÓå ÈæÏä ÓÇá ÑÇ ÈÑãíÑÏÇäÏ
'ÇÑ ÓÇá ßÈíÓå ÈÇÔÏ ÚÏÏ íß æ ÏÑÛíÑ ÇíäÕæÑÊ ÕİÑ ÑÇ ÈÑ ãíÑÏÇäÏ
Kabiseh = 0
If OnlyIL >= 1375 Then
If (OnlyIL - 1375) Mod 4 = 0 Then
Kabiseh = 1
Exit Function
End If
ElseIf OnlyIL <= 1370 Then
If (1370 - OnlyIL) Mod 4 = 0 Then
Kabiseh = 1
Exit Function
End If
End If
End Function

Function Nextay(ByVal IL_ay As Long) As Long
If (IL_ay Mod 100) = 12 Then
Nextay = (Int(IL_ay / 100) + 1) * 100 + 1
Else
Nextay = IL_ay + 1
End If
End Function

Function Previousay(ByVal IL_ay As Long) As Long
If (IL_ay Mod 100) = 1 Then
Previousay = (Int(IL_ay / 100) - 1) * 100 + 12
Else
Previousay = IL_ay - 1
End If
End Function

Public Function Firstday(IL As Integer, ay As Integer) As Long
'ÔãÇÑå Çæáíä ÑæÒ ãÇå
Dim strfd As Long
strfd = IL & Format(ay, "00") & Format(1, "00")
Firstday = DayWeekNo(strfd)
End Function


