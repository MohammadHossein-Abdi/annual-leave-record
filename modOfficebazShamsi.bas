Option Explicit
Public STRDATE As String

Private Month_Name, Spring_Fall
Private Time_Difference, Time_Client
Private Base_Year
Public a, B, c As String

Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_DLGMODALFRAME As Long = &H1


'\\\\\\\\\\\\\\\\
    Dim DateFormats()
    Dim CurrentYear()
    Dim MonthNames1(), MonthNames2(), MonthNames3()
    Dim DaysOfMonth1(), DaysOfMonth2(), DaysOfMonth3()
    Dim EmsalKabiseAst As Boolean
    Dim LastKabiseG As Integer 'ÊÇÑíÎ ÂÎÑíä ÓÇá ÞãÑí ßå ßÈíÓå ÔÏå ÇÓÊ'
    Dim TarikheMabna()
    Dim YearNames()
    Dim WeekDays1(), WeekDays2(), WeekDays3()

'--- Farsi Date Convertor --------------------'
 
Private Sub Get_Date(ByVal Days, sal, Mah, Rooz)
   Dim Years, Year_Length
   Do While Days >= 0
     If Kabiseh1(Years) Then
        Year_Length = 366
     Else
        Year_Length = 365
     End If
     If Days - Year_Length >= 0 Then
        Years = Years + 1
        Days = Days - Year_Length
     Else
        sal = Base_Year + Years
        If Days <= 185 Then
           Mah = 1 + (Days \ 31)
           Rooz = 1 + (Days Mod 31)
        Else
           Days = Days - 186
           Mah = 7 + (Days \ 30)
           Rooz = 1 + (Days Mod 30)
        End If
        Exit Sub
     End If
   Loop
End Sub
Private Function Kabiseh1(ByVal Years)
   Dim Temp
   Kabiseh1 = False
   Temp = (Base_Year + Years) - 1309
   If (((Temp Mod 32) - (Temp \ 32)) Mod 4) = 0 Then Kabiseh1 = True
End Function
Public Property Let SFhour(x)
   Spring_Fall = x
End Property
Public Property Let Time_Diff(ByVal t)
   Time_Difference = t
End Property
Public Property Let state(ByVal s)
   Month_Name = s
End Property
Public Function To_Hejri(ByVal what_date, Optional Month_Name)
   Dim Days, Day_Name, Day_Number, Temp_Days, Months
   Spring_Fall = False
   If IsMissing(Month_Name) Then Month_Name = 0
 
   Time_Difference = #12:00:00 AM#
   Base_Year = 1332
 
Months = Array(ChrW(1601) & ChrW(1585) & ChrW(1608) & ChrW(1585) & ChrW(1583) & ChrW(1740) & ChrW(1606), ChrW(1575) & ChrW(1585) & ChrW(1583) & ChrW(1740) & ChrW(1576) & ChrW(1588) & ChrW(1607) & ChrW(1578), ChrW(1582) & ChrW(1585) & ChrW(1583) & ChrW(1575) & ChrW(1583), ChrW(1578) & ChrW(1740) & ChrW(1585), ChrW(1605) & ChrW(1585) & ChrW(1583) & ChrW(1575) & ChrW(1583), ChrW(1588) & ChrW(1607) & ChrW(1585) & ChrW(1740) & ChrW(1608) & ChrW(1585), ChrW(1605) & ChrW(1607) & ChrW(1585), ChrW(1570) & ChrW(1576) & ChrW(1575) & ChrW(1606), ChrW(1570) & ChrW(1584) & ChrW(1585), ChrW(1583) & ChrW(1740), ChrW(1576) & ChrW(1607) & ChrW(1605) & ChrW(1606), ChrW(1570) & ChrW(1587) & ChrW(1601) & ChrW(1606) & ChrW(1583))
 
Day_Name = Array(ChrW(1740) & ChrW(1705) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607), ChrW(1583) & ChrW(1608) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607), ChrW(1587) & ChrW(1607) & " " & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607), ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607), ChrW(1662) & ChrW(1606) & ChrW(1580) & " " & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607), ChrW(1580) & ChrW(1605) & ChrW(1593) & ChrW(1607), ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607))
   
   Days = DateDiff("d", #3/21/1953#, what_date)
   Day_Number = Weekday(what_date)
   Dim Year_Length, sal, Mah, Rooz, temp_date
   If FormatDateTime(what_date + Time_Difference, vbShortDate) <> FormatDateTime(what_date, vbShortDate) Then
      Days = Days + 1
      Day_Number = (Day_Number + 1)
      If Day_Number = 8 Then Day_Number = 1
   End If
   Time_Client = FormatDateTime(what_date + Time_Difference, vbLongTime)
   Call Get_Date(Days, sal, Mah, Rooz)
   If ((Mah >= 1 And Mah <= 6) And Not ((Mah = 1 And Rooz = 1) Or (Mah = 6 And Rooz = 31))) And Spring_Fall = True Then
      If FormatDateTime(what_date + Time_Difference + #1:00:00 AM#, vbShortDate) <> FormatDateTime(what_date + Time_Difference, vbShortDate) Then
         Temp_Days = Days + 1
         Day_Number = (Day_Number + 1)
         If Day_Number = 8 Then Day_Number = 1
      Else
         Temp_Days = Days
      End If
      Time_Client = FormatDateTime(what_date + Time_Difference + #1:00:00 AM#, vbLongTime)
      If Temp_Days <> Days Then
         Days = Temp_Days
         If Rooz = 30 And Mah = 6 Then
            If DateDiff("n", Time_Client, #1:00:00 AM#) <= 60 And DateDiff("n", Time_Client, #1:00:00 AM#) >= 0 Then
               Time_Client = FormatDateTime(what_date + Time_Difference, vbLongTime)
               Days = Days - 1
               If Day_Number = 1 Then
                  Day_Number = 7
               Else
                  Day_Number = Day_Number - 1
               End If
            End If
         End If
         Call Get_Date(Days, sal, Mah, Rooz)
      End If
   End If
   If Month_Name = 0 Then
      If Rooz < 10 Then Rooz = "0" & Rooz
      If Mah < 10 Then Mah = "0" & Mah
      To_Hejri = sal & "/" & Mah & "/" & Rooz
   ElseIf Month_Name = 1 Then
      To_Hejri = Rooz & " " & Months(Mah - 1) & " " & sal
   ElseIf Month_Name = 2 Then
      To_Hejri = Day_Name(Day_Number - 1) & " " & sal & "/" & Mah & "/" & Rooz
   ElseIf Month_Name = 3 Then
      To_Hejri = Day_Name(Day_Number - 1) & "  " & Rooz & "  " & Months(Mah - 1) & "  " & sal
   End If
End Function
Public Function To_Time(what_date)
   Call To_Hejri(what_date)
   To_Time = Time_Client
End Function
Private Sub Class_Initialize()
   Spring_Fall = False
   Month_Name = 0
   Time_Difference = #12:00:00 AM#
   Base_Year = 1332
End Sub



Public Static Function Shamsi() As Long
'Çíä ÊÇÈÚ ÊÇÑíÎ ÌÇÑí ÓíÓÊã ÑÇ Èå ÊÇÑíÎ åÌÑí ÔãÓí ÊÈÏíá ãí ßäÏ
Dim Shamsi_Mabna As Long
Dim Miladi_mabna As Date
Dim Dif As Long
'ÏÑ ÇíäÌÇ 78/10/11 ÈÇ 2000/01/01 ãÚÇÏá ÞÑÇÑÏÇÏå ÔÏå
Shamsi_Mabna = 13411011
Miladi_mabna = #1/1/1963#
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

Public Function Slash(f_date As Variant) As String
' Çíä ÊÇÈÚ íß ÊÇÑíÎ ÑÇ ÏÑíÇÝÊ æ ÈÕæÑÊ íß ÑÔÊå 10 ÑÞãí ÔÇãá / æ åÇÑ ÑÞã ÈÑÇí ÓÇá ÈÇÒãíÑÏÇäÏ
f_date = Replace(f_date, "/", "")
Dim a As Long
a = CLng(f_date)
Slash = Format(sal(a), "0000") & "/" & Format(Mah(a), "00") & "/" & Format(Rooz(a), "00")
End Function
Public Function NoSlash(f_date As Variant) As String
' Çíä ÊÇÈÚ íß ÊÇÑíÎ ÑÇ ÏÑíÇÝÊ æ ÈÕæÑÊ íß ÑÔÊå 10 ÑÞãí ÔÇãá / æ åÇÑ ÑÞã ÈÑÇí ÓÇá ÈÇÒãíÑÏÇäÏ
NoSlash = Replace(f_date, "/", "")
End Function

Function ValidDate(f_date As Variant) As Boolean
' Çíä ÊÇÈÚ ÇÚÊÈÇÑ íß ÚÏÏ æÑæÏí ÑÇ ÇÒ äÙÑ ÊÇÑíÎ åÌÑí ÔãÓí ÈÑÑÓí ãí ßäÏ
' ÑÇ ÈÑãí ÑÏÇäÏ False æÇÑ äÇãÚÊÈÑ ÈÇÔÏ True ÇÑ ÊÇÑíÎ ãÚÊÈÑ ÈÇÔÏ
On Error GoTo Err_ValidDate
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Dim m, s, R As Byte
f_date = Replace(f_date, "/", "")
R = Rooz(CLng(f_date))
m = Mah(CLng(f_date))
s = sal(CLng(f_date))
If f_date < 10000101 Then Exit Function
If m > 12 Or m = 0 Or R = 0 Then Exit Function
If R > MahDays(s, m) Then Exit Function
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

Public Function AddDay(ByVal f_date As Variant, ByVal add As Long) As Long
    'Çíä ÊÇÈÚ ÊÚÏÇÏ ÑæÒ ÏáÎæÇå ÑÇ Èå ÊÇÑíÎ ÑæÒ ÇÖÇÝå ãíßäÏ
    On Error GoTo Err_AddDay
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    f_date = Replace(f_date, "/", "")
    Dim K, m, R, Days As Byte
    Dim s As Integer
    R = Rooz(CLng(f_date))
    m = Mah(CLng(f_date))
    s = sal(CLng(f_date))
    K = Kabiseh(s)
    'ÊÈÏíá ÑæÒ Èå ÚÏÏ 1 ÌåÊ ÇÏÇãå ãÍÇÓÈÇÊ æ íÇ ÇÊãÇã ãÍÇÓÈå
    Days = MahDays(s, m)
    If add > Days - R Then
        add = add - (Days - R + 1)
        R = 1
        If m < 12 Then
            m = m + 1
        Else
            m = 1
            s = s + 1
        End If
    Else
        R = R + add
        add = 0
    End If
    While add > 0
        K = Kabiseh(s) 'ßÈíÓå: 1 æ ÛíÑ ßÈíÓå: 0
        Days = MahDays(s, m) 'ÊÚÏÇÏ ÑæÒåÇí ãÇå ÝÚáí
        Select Case add
            Case Is < Days
                'ÇÑ ÊÚÏÇÏ ÑæÒåÇí ÇÝÒæÏäí ßãÊÑ ÇÒ íß ãÇå ÈÇÔÏ
                R = R + add
                add = 0
            Case Days To IIf(K = 0, 365, 366) - 1
                'ÇÑ ÊÚÏÇÏ ÑæÒåÇí ÇÝÒæÏäí ÈíÔÊÑ ÇÒ íß ãÇå æ ßãÊÑ ÇÒ íß ÓÇá ÈÇÔÏ
                add = add - Days
                If m < 12 Then
                    m = m + 1
                Else
                    s = s + 1
                    m = 1
                End If
            Case Else
                'ÇÑ ÊÚÏÇÏ ÑæÒåÇí ÇÝÒæÏäí ÈíÔÊÑ ÇÒ íß ÓÇá ÈÇÔÏ
                s = s + 1
                add = add - IIf(K = 0, 365, 366)
        End Select
    Wend
    AddDay = CLng(s & Format(m, "00") & Format(R, "00"))

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


Public Function AddWeek(ByVal f_date As Variant, ByVal addW As Long) As Long
    'Çíä ÊÇÈÚ ÊÚÏÇÏ åÝÊå ÏáÎæÇå ÑÇ Èå ÊÇÑíÎ ÑæÒ ÇÖÇÝå ãíßäÏ
    On Error GoTo Err_AddDay
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    addW = addW * 7
    f_date = Replace(f_date, "/", "")
    Dim K, m, R, Days As Byte
    Dim s As Integer
    R = Rooz(CLng(f_date))
    m = Mah(CLng(f_date))
    s = sal(CLng(f_date))
    K = Kabiseh(s)
    'ÊÈÏíá ÑæÒ Èå ÚÏÏ 1 ÌåÊ ÇÏÇãå ãÍÇÓÈÇÊ æ íÇ ÇÊãÇã ãÍÇÓÈå
    Days = MahDays(s, m)
    If addW > Days - R Then
        addW = addW - (Days - R + 1)
        R = 1
        If m < 12 Then
            m = m + 1
        Else
            m = 1
            s = s + 1
        End If
    Else
        R = R + addW
        addW = 0
    End If
    While addW > 0
        K = Kabiseh(s) 'ßÈíÓå: 1 æ ÛíÑ ßÈíÓå: 0
        Days = MahDays(s, m) 'ÊÚÏÇÏ ÑæÒåÇí ãÇå ÝÚáí
        Select Case addW
            Case Is < Days
                'ÇÑ ÊÚÏÇÏ åÝÊå åÇí ÇÝÒæÏäí ßãÊÑ ÇÒ íß ãÇå ÈÇÔÏ
                R = R + addW
                addW = 0
            Case Days To IIf(K = 0, 365, 366) - 1
                'ÇÑ ÊÚÏÇÏ åÝÊå åÇí ÇÝÒæÏäí ÈíÔÊÑ ÇÒ íß ãÇå æ ßãÊÑ ÇÒ íß ÓÇá ÈÇÔÏ
                addW = addW - Days
                If m < 12 Then
                    m = m + 1
                Else
                    s = s + 1
                    m = 1
                End If
            Case Else
                'ÇÑ ÊÚÏÇÏ åÝÊå åÇí ÇÝÒæÏäí ÈíÔÊÑ ÇÒ íß ÓÇá ÈÇÔÏ
                s = s + 1
                addW = addW - IIf(K = 0, 365, 366)
        End Select
    Wend
    AddWeek = CLng(s & Format(m, "00") & Format(R, "00"))

Exit_AddDay:
    On Error Resume Next
    Exit Function
Err_AddDay:
    Select Case err.Number
        Case 0
            Resume Exit_AddDay:
        Case 94
            AddWeek = 0
        Case Else
            MsgBox err.Number & " " & err.Description, vbExclamation, "Error in module Module2 - function AddDay"
            Resume Exit_AddDay:
    End Select
End Function

Public Function AddMonth(ByVal f_date As Variant, ByVal addM As Long, ByVal add As Long, ByVal Subtract As Long) As Long
    'Çíä ÊÇÈÚ ÊÚÏÇÏ ãÇå ÏáÎæÇå ÑÇ Èå ÊÇÑíÎ ÑæÒ ÇÖÇÝå ãíßäÏ
    On Error GoTo Err_AddDay
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    f_date = Replace(f_date, "/", "")
    Dim K, m, R, Days As Byte
    Dim s As Integer
    R = Rooz(CLng(f_date))
    m = Mah(CLng(f_date))
    s = sal(CLng(f_date))
    K = Kabiseh(s)
    
    
 
 If Subtract >= R - 1 Then
Subtract = Subtract - (R - 1)
R = 1
Else
R = R - Subtract
Subtract = 0
End If
While Subtract > 0
K = Kabiseh(s - 1)
Days = MahDays(IIf(m >= 2, s, s - 1), IIf(m >= 2, m - 1, 12))
Select Case Subtract
Case Is < Days

R = Days - Subtract + 1
Subtract = 0
If m >= 2 Then
m = m - 1
Else
s = s - 1
m = 12
End If
Case Days To IIf(K = 0, 365, 366) - 1

Subtract = Subtract - Days
If m >= 2 Then
m = m - 1
Else
s = s - 1
m = 12
End If
Case Else

s = s - 1
Subtract = Subtract - IIf(K = 0, 365, 366)

End Select
Wend


 
Days = MahDays(s, m)
    If add > Days - R Then
        add = add - (Days - R + 1)
        R = 1
        If m < 12 Then
            m = m + 1
        Else
            m = 1
            s = s + 1
        End If
    Else
        R = R + add
        add = 0
    End If
    While add > 0
        K = Kabiseh(s)
        Days = MahDays(s, m)
        Select Case add
            Case Is < Days
               R = R + add
                add = 0
            Case Days To IIf(K = 0, 365, 366) - 1
                add = add - Days
                If m < 12 Then
                    m = m + 1
                Else
                    s = s + 1
                    m = 1
                End If
            Case Else
                s = s + 1
                add = add - IIf(K = 0, 365, 366)
        End Select
    Wend


    AddMonth = CLng(s & Format(m + addM, "00") & Format(R + add, "00"))
    
    If Mah(AddMonth) > 12 Then
            AddMonth = CLng(s + Int(Mah(AddMonth) / 12) & Format((((Mah(AddMonth) / 12)) - Int(Mah(AddMonth) / 12)) * 12, "00")) & Format(R + add, "00")

    End If
    If Mah(AddMonth) >= 7 And Mah(AddMonth) <= 11 And Rooz(AddMonth) >= 31 Then
            R = 30
            AddMonth = CLng(s + Int(addM / 12) & Format(m + addM - Int(addM / 12) * 12, "00")) & Format(R + add, "00")
    End If
    If Mah(AddMonth) = 12 And Rooz(AddMonth) >= 29 Then
            R = 29
            AddMonth = CLng(s + Int(addM / 12) & Format(m + addM - Int(addM / 12) * 12, "00")) & Format(R + add, "00")
    End If

    If Mah(AddMonth) = 12 And Rooz(AddMonth) >= 29 And Kabiseh(sal(AddMonth)) = 1 Then
            R = 30
            AddMonth = CLng(s + Int(addM / 12) & Format(m + addM - Int(addM / 12) * 12, "00")) & Format(R + add, "00")

    End If
 

 
 
 
Exit_AddDay:
    On Error Resume Next
    Exit Function
Err_AddDay:
    Select Case err.Number
        Case 0
            Resume Exit_AddDay:
        Case 94
            AddMonth = 0
                   Case 13
            AddMonth = 0
 
        Case Else
            MsgBox err.Number & " " & err.Description, vbExclamation, "Error in module Module2 - function AddDay"
            Resume Exit_AddDay:
    End Select
End Function







Function SubDay(ByVal f_date As Variant, ByVal Subtract As Long) As Long
'Èå ÊÚÏÇÏ ÑæÒ ãÚíäí ÇÒ íß ÊÇÑíÎ ßã ßÑÏå æ ÊÇÑíÎ ÍÇÕáå ÑÇ ÇÑÇÆå ãíßäÏ
On Error GoTo Err_SubDay
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
f_date = Replace(f_date, "/", "")
Dim K, m, s, R, Days As Byte
R = Rooz(CLng(f_date))
m = Mah((CLng(f_date)))
s = sal((CLng(f_date)))
K = Kabiseh(s)
'ÊÈÏíá ÑæÒ Èå ÚÏÏ 1 ÌåÊ ÇÏÇãå ãÍÇÓÈÇÊ æ íÇ ÇÊãÇã ãÍÇÓÈå
If Subtract >= R - 1 Then
Subtract = Subtract - (R - 1)
R = 1
Else
R = R - Subtract
Subtract = 0
End If
While Subtract > 0
K = Kabiseh(s - 1) 'ßÈíÓå: 1 æ ÛíÑ ßÈíÓå: 0
Days = MahDays(IIf(m >= 2, s, s - 1), IIf(m >= 2, m - 1, 12)) 'ÊÚÏÇÏ ÑæÒåÇí ãÇå ÞÈáí
Select Case Subtract
Case Is < Days
'ÇÑ ÊÚÏÇÏ ÑæÒåÇí ßÇåÔ ßãÊÑ ÇÒ íß ãÇå ÈÇÔÏ
R = Days - Subtract + 1
Subtract = 0
If m >= 2 Then
m = m - 1
Else
s = s - 1
m = 12
End If
Case Days To IIf(K = 0, 365, 366) - 1
'ÇÑ ÊÚÏÇÏ ÑæÒåÇí ßÇåÔ ÈíÔÊÑ ÇÒ íß ãÇå æ ßãÊÑ ÇÒ íß ÓÇá ÈÇÔÏ
Subtract = Subtract - Days
If m >= 2 Then
m = m - 1
Else
s = s - 1
m = 12
End If
Case Else
'ÇÑ ÊÚÏÇÏ ÑæÒåÇí ßÇåÔ ÈíÔÊÑ ÇÒ íß ÓÇá ÈÇÔÏ
s = s - 1
Subtract = Subtract - IIf(K = 0, 365, 366)
End Select
Wend
SubDay = (s * 10000) + (m * 100) + (R)

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

Public Function DayWeekNo(f_date As Variant) As Byte
'Çíä ÊÇÈÚ íß ÊÇÑíÎ ÑÇ ÏÑíÇÝÊ ßÑÏå æ ÔãÇÑå ÑæÒ åÝÊå ÑÇ ãÔÎÕ ãí ßäÏ
'ÇÑ ÔäÈå ÈÇÔÏ ÚÏÏ 0
'ÇÑ 1ÔäÈå ÈÇÔÏ ÚÏÏ 1
'......
'ÇÑ ÌãÚå ÈÇÔÏ ÚÏÏ 6
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
f_date = Replace(f_date, "/", "")
Dim day As String
Dim Shmsi_Mabna As Long
Dim Dif As Long
'ãÈäÇ 80/10/11
Shmsi_Mabna = 13411011
Dif = Diff(Shmsi_Mabna, CLng(f_date))
If Shmsi_Mabna > CLng(f_date) Then
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

Public Function DayWeek(f_date As Variant) As String
'Çíä ÊÇÈÚ íß ÊÇÑíÎ ÑÇ ÏÑíÇÝÊ ßÑÏå æ ãÔÎÕ ãí ßäÏ å ÑæÒí ÇÒ åÝÊå ÇÓÊ
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Dim a As String
Dim n As Byte
n = DayWeekNo(f_date)
Select Case n
Case 0
a = ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607)
Case 1
a = ChrW(1740) & ChrW(1705) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607)
Case 2
a = ChrW(1583) & ChrW(1608) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607)
Case 3
a = ChrW(1587) & ChrW(1607) & " " & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607)
Case 4
a = ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607)
Case 5
a = ChrW(1662) & ChrW(1606) & ChrW(1580) & " " & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607)
Case 6
a = ChrW(1580) & ChrW(1605) & ChrW(1593) & ChrW(1607)
End Select
DayWeek = a
End Function

Public Function Diff(ByVal date1 As Variant, ByVal Date2 As Variant) As Long
'Çíä ÊÇÈÚ ÊÚÏÇÏ ÑæÒåÇí Èíä Ïæ ÊÇÑíÎ ÑÇ ÇÑÇÆå ãí ßäÏ
On Error GoTo Err_Diff
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
date1 = Replace(date1, "/", "")
Date2 = Replace(Date2, "/", "")
Dim tmp As Long
Dim s1, M1, R1, S2, M2, R2 As Integer
Dim Sumation As Single
Dim Flag As Boolean
Flag = False
If CLng(date1) = 0 Or IsNull(CLng(date1)) = True Or CLng(Date2) = 0 Or IsNull(CLng(Date2)) = True Then
Diff = 0
Exit Function
End If
'ÇÑ ÊÇÑíÎ ÔÑæÚ ÇÒ ÊÇÑíÎ ÇíÇä ÈÒÑÊÑ ÈÇÔÏ ÂäåÇ ãæÞÊÇ ÌÇÈÌÇ ãí ÔæäÏ
If CLng(date1) > CLng(Date2) Then
Flag = True
tmp = CLng(date1)
date1 = CLng(Date2)
Date2 = tmp
End If
R1 = Rooz(CLng(date1))
M1 = Mah(CLng(date1))
s1 = sal(CLng(date1))
R2 = Rooz(CLng(Date2))
M2 = Mah(CLng(Date2))
S2 = sal(CLng(Date2))
Sumation = 0
Do While s1 < S2 - 1 Or (s1 = S2 - 1 And (M1 < M2 Or (M1 = M2 And R1 <= R2)))
'ÇÑ íß ÓÇá íÇ ÈíÔÊÑ ÇÎÊáÇÝ ÈæÏ
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
Do While s1 < S2 Or M1 < M2 - 1 Or (M1 = M2 - 1 And R1 < R2)
'ÇÑ íß ãÇå íÇ ÈíÔÊÑ ÇÎÊáÇÝ ÈæÏ
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
                Case 13
            Diff = 0
        Case Else
            MsgBox err.Number & " " & err.Description, vbExclamation, "Error in module Module2 - function Diff"
            Resume Exit_Diff:
    End Select
End Function

Function MahName(ByVal MahNo As Byte) As String
'Çíä ÊÇÈÚ íß ÊÇÑíÎ ÑÇ ÏÑíÇÝÊ ßÑÏå æ ãÔÎÕ ãí ßäÏ å ãÇåí ÇÒ ÓÇá ÇÓÊ
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Select Case MahNo
Case 1
MahName = ChrW(1601) & ChrW(1585) & ChrW(1608) & ChrW(1585) & ChrW(1583) & ChrW(1740) & ChrW(1606)
Case 2
MahName = ChrW(1575) & ChrW(1585) & ChrW(1583) & ChrW(1740) & ChrW(1576) & ChrW(1607) & ChrW(1588) & ChrW(1578)
Case 3
MahName = ChrW(1582) & ChrW(1585) & ChrW(1583) & ChrW(1575) & ChrW(1583)
Case 4
MahName = ChrW(1578) & ChrW(1740) & ChrW(1585)
Case 5
MahName = ChrW(1605) & ChrW(1585) & ChrW(1583) & ChrW(1575) & ChrW(1583)
Case 6
MahName = ChrW(1588) & ChrW(1607) & ChrW(1585) & ChrW(1740) & ChrW(1608) & ChrW(1585)
Case 7
MahName = ChrW(1605) & ChrW(1607) & ChrW(1585)
Case 8
MahName = ChrW(1570) & ChrW(1576) & ChrW(1575) & ChrW(1606)
Case 9
MahName = ChrW(1570) & ChrW(1584) & ChrW(1585)
Case 10
MahName = ChrW(1583) & ChrW(1740)
Case 11
MahName = ChrW(1576) & ChrW(1607) & ChrW(1605) & ChrW(1606)
Case 12
MahName = ChrW(1575) & ChrW(1587) & ChrW(1601) & ChrW(1606) & ChrW(1583)
End Select
End Function

Function MahDays(ByVal sal As Integer, ByVal Mah As Byte) As Byte
'Çíä ÊÇÈÚ ÊÚÏÇÏ ÑæÒåÇí íß ãÇå ÑÇ ÈÑãí ÑÏÇäÏ
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Select Case Mah
Case 1 To 6
MahDays = 31
Case 7 To 11
MahDays = 30
Case 12
If Kabiseh(sal) = 1 Then
MahDays = 30
Else
MahDays = 29
End If
End Select
End Function

Function SalMah(ByVal f_date As Long) As Long
'ÔÔ ÑÞã Çæá ÊÇÑíÎ ßå ãÚÑÝ ÓÇá æ ãÇå ÇÓÊ ÑÇ ÈÑãí ÑÏÇäÏ
SalMah = Val(Left$(f_date, 6))
End Function

Public Function Rooz(f_date As Long) As Byte
'Çíä ÊÇÈÚ ÚÏÏ ãÑÈæØ Èå ÑæÒ íß ÊÇÑíÎ ÑÇ ÈÑãÑÏÇäÏ
Rooz = f_date Mod 100
End Function

Function Mah(f_date As Long) As Byte
'Çíä ÊÇÈÚ ÚÏÏ ãÑÈæØ Èå ãÇå íß ÊÇÑíÎ ÑÇ ÈÑãÑÏÇäÏ
Mah = Int((f_date Mod 10000) / 100)
End Function

Public Function sal(f_date As Long) As Integer
'Çíä ÊÇÈÚ 4 ÑÞã ãÑÈæØ Èå ÓÇá íß ÊÇÑíÎ ÑÇ ÈÑãÑÏÇäÏ
sal = Int(f_date / 10000)
End Function
Public Function Sal1(f_date As Long) As Integer
'Çíä ÊÇÈÚ Ïæ ÑÞã ãÑÈæØ Èå ÓÇá íß ÊÇÑíÎ ÑÇ ÈÑãÑÏÇäÏ
Sal1 = Right(Int(f_date / 10000), 2)

End Function

Public Function Kabiseh(ByVal OnlySal As Variant) As Byte
'æÑæÏí ÊÇÈÚ ÚÏÏ ÏæÑÞãí ÇÓÊ
'Çíä ÊÇÈÚ ßÈíÓå ÈæÏä ÓÇá ÑÇ ÈÑãíÑÏÇäÏ
'ÇÑ ÓÇá ßÈíÓå ÈÇÔÏ ÚÏÏ íß æ ÏÑÛíÑ ÇíäÕæÑÊ ÕÝÑ ÑÇ ÈÑ ãíÑÏÇäÏ
Kabiseh = 0
If OnlySal >= 1375 Then
If (OnlySal - 1375) Mod 4 = 0 Then
Kabiseh = 1
Exit Function
End If
ElseIf OnlySal <= 1370 Then
If (1370 - OnlySal) Mod 4 = 0 Then
Kabiseh = 1
Exit Function
End If
End If
End Function

Function NextMah(ByVal Sal_Mah As Long) As Long
If (Sal_Mah Mod 100) = 12 Then
NextMah = (Int(Sal_Mah / 100) + 1) * 100 + 1
Else
NextMah = Sal_Mah + 1
End If
End Function

Function PreviousMah(ByVal Sal_Mah As Long) As Long
If (Sal_Mah Mod 100) = 1 Then
PreviousMah = (Int(Sal_Mah / 100) - 1) * 100 + 12
Else
PreviousMah = Sal_Mah - 1
End If
End Function

Public Function Firstday(sal As Integer, Mah As Integer) As Long
'ÔãÇÑå Çæáíä ÑæÒ ãÇå
Dim strfd As Long
strfd = sal & Format(Mah, "00") & Format(1, "00")
Firstday = DayWeekNo(strfd)
End Function

Function SalMah1(ByVal f_date As Long) As Long
'ÊÇÑíÎ Çæáíä ÑæÒ ãÇå ÌÇÑí '
SalMah1 = Val(Left$(f_date, 6) & "01")
End Function
Function SalMah2(ByVal f_date As Long) As Long
'ÊÇÑíÎ ÂÎÑíä ÑæÒ ãÇå ÌÇÑí
If Mah(f_date) < 7 Then
SalMah2 = Val(Left$(f_date, 6) & "31")
End If
If Mah(f_date) > 6 And Mah(f_date) < 12 Then
SalMah2 = Val(Left$(f_date, 6) & "30")
End If
If Mah(f_date) = 12 Then
SalMah2 = Val(Left$(f_date, 6) & "29") + Kabiseh(sal(f_date))
End If
End Function

Public Function Salh(f_date As String) As String
'Çíä ÊÇÈÚ 4 ÑÞã ãÑÈæØ Èå ÓÇá íß ÊÇÑíÎ ÑÇ Èå ÍÑæÝ ÈÑ ãí ÑÏÇäÏ
Salh = Horof(Int(f_date / 10000))
End Function

Public Function roozH(f_date As String) As String
'Çíä ÊÇÈÚ ÚÏÏ ãÑÈæØ Èå ÑæÒ íß ÊÇÑíÎ ÑÇ Èå ÍÑæÝ ÈÑãÑÏÇäÏ
roozH = Horof(f_date Mod 100)
End Function
Function MahH(f_date As String) As String
'Çíä ÊÇÈÚ ÚÏÏ ãÑÈæØ Èå ãÇå íß ÊÇÑíÎ ÑÇ Èå ÍÑæÝ ÈÑãÑÏÇäÏ
MahH = MahName(Int((f_date Mod 10000) / 100))
End Function

' Çíä ÊÇÈÚ í˜ ÊÇÑíÎ ÔãÓí ÑÇ Èå ÕæÑÊ ÊãÇãÇ ãÊäí äãÇíÔ ãí ÏåÏ
Public Function Matni(f_date As String, Optional mode As Integer) As String
    Dim strSal As String
    Dim strRooz As String
    Dim strMah As String
    strSal = Salh(f_date)
    strRooz = roozH(f_date)
    strMah = MahH(f_date)
    
If Right(f_date, 2) <> 3 And Right(f_date, 2) <> 30 And Right(f_date, 2) <> 23 Then
Matni = strRooz & ChrW(1605) & " " & strMah & " " & ChrW(1605) & ChrW(1575) & ChrW(1607) & " " & strSal
Else
Matni = strRooz & " " & strMah & " " & ChrW(1605) & ChrW(1575) & ChrW(1607) & " " & strSal
End If

If mode = 1 Then

If Right(f_date, 2) <> 3 And Right(f_date, 2) <> 30 And Right(f_date, 2) <> 23 Then
Matni = ChrW(1575) & ChrW(1605) & ChrW(1585) & ChrW(1608) & ChrW(1586) & " " & strRooz & ChrW(1605) & " " & strMah & " " & ChrW(1605) & ChrW(1575) & ChrW(1607) & " " & strSal
Else
Matni = ChrW(1575) & ChrW(1605) & ChrW(1585) & ChrW(1608) & ChrW(1586) & " " & strRooz & " " & strMah & " " & ChrW(1605) & ChrW(1575) & ChrW(1607) & " " & strSal
End If
End If


If mode = 2 Then

If Right(f_date, 2) <> 3 And Right(f_date, 2) <> 30 And Right(f_date, 2) <> 23 Then
    Matni = DayWeek(f_date) & " " & strRooz & ChrW(1605) & " " & strMah & " " & ChrW(1605) & ChrW(1575) & ChrW(1607) & " " & strSal
Else
    Matni = DayWeek(f_date) & " " & strRooz & " " & ChrW(1575) & ChrW(1605) & " " & strMah & " " & ChrW(1605) & ChrW(1575) & ChrW(1607) & " " & strSal
End If
End If

End Function

'Public Function ghamari(f_date As String) As Double
'
'    Dim part1 As Double
'    Dim part2 As Double
'
'    Dim sGhamari As String
'    Dim mGhamari As String
'    Dim rGhamari As Double
'
'    part1 = ((Left(f_date, 4) - 1) * 365.2422 + Diff(Left(f_date, 4) & "0101", f_date) - 118) / 354.367
'    sGhamari = Int(part1) + 1
'
'    part2 = Int((part1 - Int(part1)) * 12) + 1
'    mGhamari = Format(part2, "00")
'
'    rGhamari = Round((((part1 - Int(part1)) * 12) - Int(((part1 - Int(part1)) * 12))) * 29.53, 1)
'
''    ghamari = sGhamari & mGhamari & Format(rGhamari, "00")
'    ghamari = rGhamari
'
'End Function



' ÊæÇÈÚ ãÑÈæØ Èå ÊÈÏíá ÚÏÏ Èå ÍÑæÝ naderweb.ir


Function Horof(eNo As Double) As String
    Dim eFixed As String, eDecimal As String
    Dim sResult As String
    Dim LTZ As Boolean

   If eNo < 0 Then
        LTZ = True
        eNo = -eNo
    Else
        LTZ = False
    End If


If (eNo < 1 And eNo > 0 And Len(CStr(eNo)) > 8) Or InStr(1, Trim(Str(eNo)), "E") > 0 Then
            If LTZ Then
            Horof = "-##########"
            Else
            Horof = "##########"
            End If
Exit Function
End If

    'return fixed value of given number as string
    eFixed = Fix(eNo)
    
    'if this number has some decimals
    If (Len(CStr(eNo)) - Len(eFixed)) Then
        'get it as a string, Example: return `125` for `12.125`
       If eFixed = 0 Then
        eDecimal = Mid(CStr(eNo), Len(eFixed) + 2)
       Else
        eDecimal = Left(Mid(CStr(eNo), Len(eFixed) + 2), 6)
       End If
        'return fixed part as text
        sResult = Horof_fix(eFixed) + " " & ChrW(1605) & ChrW(1605) & ChrW(1740) & ChrW(1586) & " "
        'if decimal section is `5` then use `äíã` Instead of `äÌ Ïåã`
        'this is optional, u can remove it if u like
        If eDecimal = 5 Then
            sResult = sResult + " " + ChrW(1606) & ChrW(1740) & ChrW(1605)
        Else
            'convert the decimal part of number to text
            sResult = sResult + " " + Horof_fix(eDecimal)
            'add extra suffix at end of string, depending to number of decimal places
            sResult = sResult + " " + Choose(Len(eDecimal), ChrW(1583) & ChrW(1607) & ChrW(1605), ChrW(1589) & ChrW(1583) & ChrW(1605), _
                                            ChrW(1607) & ChrW(1586) & ChrW(1575) & ChrW(1585) & ChrW(1605), ChrW(1583) & ChrW(1607) & " " & ChrW(1607) & ChrW(1586) & ChrW(1575) & ChrW(1585) & ChrW(1605), _
                                           ChrW(1589) & ChrW(1583) & " " & ChrW(1607) & ChrW(1586) & ChrW(1575) & ChrW(1585) & ChrW(1605), ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1608) & ChrW(1606) & ChrW(1740) & ChrW(1608) & ChrW(1605)) ', _
                                             ....
        End If
            
    Else
        'if this number is originally an integer then convert it using normal method
        sResult = Horof_fix(eNo)
    End If
    'return the result. ;)
    Horof = sResult

If LTZ Then Horof = " " & ChrW(1605) & ChrW(1606) & ChrW(1601) & ChrW(1740) & Horof
If Horof = ChrW(1605) & ChrW(1606) & ChrW(1601) & ChrW(1740) & "##########" Then Horof = "##########"
If eNo = 0 Then Horof = ChrW(1589) & ChrW(1601) & ChrW(1585)
End Function


Function Horof_fix(ByVal Number As Double) As String
If Number = 0 Then
Horof_fix = ChrW(1589) & ChrW(1601) & ChrW(1585)
End If
Dim Flag As Boolean
Dim s As String
Dim i, L As Byte
Dim K(1 To 5) As Double

s = Trim(Str(Number))
L = Len(s)
If L > 15 Then
Horof_fix = "##########"
Exit Function
End If
For i = 1 To 15 - L
s = "0" & s
Next i
For i = 1 To Int((L / 3) + 0.99)
K(5 - i + 1) = Val(Mid(s, 3 * (5 - i) + 1, 3))
Next i
Flag = False
s = ""
For i = 1 To 5
If K(i) <> 0 Then
Select Case i
Case 1
s = s & Three(K(i)) & " " & ChrW(1578) & ChrW(1585) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1608) & ChrW(1606) & " "
Flag = True
Case 2
s = s & IIf(Flag = True, ChrW(1608), "") & Three(K(i)) & " " & ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1575) & ChrW(1585) & ChrW(1583) & " "
Flag = True
Case 3
s = s & IIf(Flag = True, ChrW(1608), "") & Three(K(i)) & " " & ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1608) & ChrW(1606) & " "
Flag = True
Case 4
s = s & IIf(Flag = True, ChrW(1608), "") & Three(K(i)) & " " & ChrW(1607) & ChrW(1586) & ChrW(1575) & ChrW(1585) & " "
Flag = True
Case 5
s = s & IIf(Flag = True, ChrW(1608), "") & Three(K(i))
End Select
End If
Next i
Horof_fix = s
End Function


Function Three(ByVal Number As Integer) As String
Dim s As String
Dim i, L As Long
Dim h(1 To 3) As Byte
Dim Flag As Boolean
L = Len(Trim(Str(Abs(Number))))
If Number = 0 Then
Three = ""
Exit Function
End If
If Number = 100 Then
Three = ChrW(1740) & ChrW(1705) & ChrW(1589) & ChrW(1583)
Exit Function
End If

If L = 2 Then h(1) = 0
If L = 1 Then
h(1) = 0
h(2) = 0
End If

For i = 1 To L
h(3 - i + 1) = Mid(Trim(Str(Abs(Number))), L - i + 1, 1)
Next i

Select Case h(1)
Case 1
s = " " & ChrW(1740) & ChrW(1705) & ChrW(1589) & ChrW(1583)
Case 2
s = " " & ChrW(1583) & ChrW(1608) & ChrW(1740) & ChrW(1587) & ChrW(1578)
Case 3
s = " " & ChrW(1587) & ChrW(1740) & ChrW(1589) & ChrW(1583)
Case 4
s = " " & ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585) & ChrW(1589) & ChrW(1583)
Case 5
s = " " & ChrW(1662) & ChrW(1575) & ChrW(1606) & ChrW(1589) & ChrW(1583)
Case 6
s = " " & ChrW(1588) & ChrW(1588) & ChrW(1589) & ChrW(1583)
Case 7
s = " " & ChrW(1607) & ChrW(1601) & ChrW(1578) & ChrW(1589) & ChrW(1583)
Case 8
s = " " & ChrW(1607) & ChrW(1588) & ChrW(1578) & ChrW(1589) & ChrW(1583)
Case 9
s = " " & ChrW(1606) & ChrW(1607) & ChrW(1589) & ChrW(1583)
End Select

Select Case h(2)
Case 1
Select Case h(3)
Case 0
s = s & " " & ChrW(1608) & " " & ChrW(1583) & ChrW(1607)
Case 1
s = s & " " & ChrW(1608) & " " & ChrW(1740) & ChrW(1575) & ChrW(1586) & ChrW(1583) & ChrW(1607)
Case 2
s = s & " " & ChrW(1608) & " " & ChrW(1583) & ChrW(1608) & ChrW(1575) & ChrW(1586) & ChrW(1583) & ChrW(1607)
Case 3
s = s & " " & ChrW(1608) & " " & ChrW(1587) & ChrW(1740) & ChrW(1586) & ChrW(1583) & ChrW(1607)
Case 4
s = s & " " & ChrW(1608) & " " & ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585) & ChrW(1583) & ChrW(1607)
Case 5
s = s & " " & ChrW(1608) & " " & ChrW(1662) & ChrW(1575) & ChrW(1606) & ChrW(1586) & ChrW(1583) & ChrW(1607)
Case 6
s = s & " " & ChrW(1608) & " " & ChrW(1588) & ChrW(1575) & ChrW(1606) & ChrW(1586) & ChrW(1583) & ChrW(1607)
Case 7
s = s & " " & ChrW(1608) & " " & ChrW(1607) & ChrW(1601) & ChrW(1583) & ChrW(1607)
Case 8
s = s & " " & ChrW(1608) & " " & ChrW(1607) & ChrW(1580) & ChrW(1583) & ChrW(1607)
Case 9
s = s & " " & ChrW(1608) & " " & ChrW(1606) & ChrW(1608) & ChrW(1586) & ChrW(1583) & ChrW(1607)
End Select

Case 2
s = s & " " & ChrW(1608) & " " & ChrW(1576) & ChrW(1740) & ChrW(1587) & ChrW(1578)
Case 3
s = s & " " & ChrW(1608) & " " & ChrW(1587) & ChrW(1740)
Case 4
s = s & " " & ChrW(1608) & " " & ChrW(1670) & ChrW(1607) & ChrW(1604)
Case 5
s = s & " " & ChrW(1608) & " " & ChrW(1662) & ChrW(1606) & ChrW(1580) & ChrW(1575) & ChrW(1607)
Case 6
s = s & " " & ChrW(1608) & " " & ChrW(1588) & ChrW(1589) & ChrW(1578)
Case 7
s = s & " " & ChrW(1608) & " " & ChrW(1607) & ChrW(1601) & ChrW(1578) & ChrW(1575) & ChrW(1583)
Case 8
s = s & " " & ChrW(1608) & " " & ChrW(1607) & ChrW(1588) & ChrW(1578) & ChrW(1575) & ChrW(1583)
Case 9
s = s & " " & ChrW(1608) & " " & ChrW(1606) & ChrW(1608) & ChrW(1583)
End Select

If h(2) <> 1 Then
Select Case h(3)
Case 1
s = s & " " & ChrW(1608) & " " & ChrW(1740) & ChrW(1705)
Case 2
s = s & " " & ChrW(1608) & " " & ChrW(1583) & ChrW(1608)
Case 3
s = s & " " & ChrW(1608) & " " & ChrW(1587) & ChrW(1607)
Case 4
s = s & " " & ChrW(1608) & " " & ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585)
Case 5
s = s & " " & ChrW(1608) & " " & ChrW(1662) & ChrW(1606) & ChrW(1580)
Case 6
s = s & " " & ChrW(1608) & " " & ChrW(1588) & ChrW(1588)
Case 7
s = s & " " & ChrW(1608) & " " & ChrW(1607) & ChrW(1601) & ChrW(1578)
Case 8
s = s & " " & ChrW(1608) & " " & ChrW(1607) & ChrW(1588) & ChrW(1578)
Case 9
s = s & " " & ChrW(1608) & " " & ChrW(1606) & ChrW(1607)
End Select
End If
s = IIf(L < 3, Right(s, Len(s) - 3), s)
Three = s
End Function







'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\



'----------------------ÊæÇÈÚ ÊÈÏíá ÊÇÑíÎ----------------------------
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'ÝåÑÓÊ ãÊÛíÑåÇí Úãæãí ßå Èå ÂäåÇ äíÇÒ ÏÇÑíã:

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


   ' Public Const seperator = "/"
Public Function DateDifShamsi(dsYear1, dsMonth1, dsD1, dsYear2, dsMonth2, dsD2) As Long
    'Çíä ÊÇÈÚ¡ ÝÇÕáå ÒãÇäí Èíä Ïæ ÊÇÑíÎ ÔãÓí ÑÇ ÈÑ ÍÓÈ ÑæÒ ãÍÇÓÈå ãí ßäÏ
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim dsDays1, dsDays2 As Long
    Dim dsKabiseYears1, dsKabiseYears2 As Integer
 
 
    Call SetVariables
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'---> ãÍÇÓÈå ÝÇÕáå ÒãÇäí ÊÇÑíÎ Çæá ÊÇ 1/1/1 ÈÑ ÍÓÈ ÑæÒ
'
Dim i As Integer
    dsDays1 = dsD1 - 1
    For i = 1 To dsMonth1 - 1
        dsDays1 = dsDays1 + DaysOfMonth2(i - 1)
        
    Next
    
    For i = 1 To dsYear1 - 1
        If IsKabiseShamsi(i) Then
            dsDays1 = dsDays1 + 366
        Else
            dsDays1 = dsDays1 + 365
            
        End If
    Next
   'Debug.Print dsDays1
   
'--> ãÍÇÓÈå ÝÇÕáå ÒãÇäí ÊÇÑíÎ Ïæã ÊÇ ãÈÏÇ ÊÇÑíÎ ÔãÓí ÈÑ ÍÓÈ ÑæÒ
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    dsDays2 = dsD2 - 1
    For i = 1 To dsMonth2 - 1
        dsDays2 = dsDays2 + DaysOfMonth2(i - 1)
        
    Next
    
    For i = 1 To dsYear2 - 1
        If IsKabiseShamsi(i) Then
            dsDays2 = dsDays2 + 366
        Else
            dsDays2 = dsDays2 + 365
            
        End If
    Next
    
    DateDifShamsi = dsDays2 - dsDays1
    
    
    
End Function

Public Function DateDifGhamari(dgYear1, dgMonth1, dgD1, dgYear2, dgMonth2, dgD2) As Long
 
    'Çíä ÊÇÈÚ¡ ÝÇÕáå ÒãÇäí Èíä Ïæ ÊÇÑíÎ ÞãÑí ÑÇ ÈÑ ÍÓÈ ÑæÒ ãÍÇÓÈå ãí ßäÏ
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim dgDays1, dgDays2 As Long
    'Dim dgKabiseYears1, dgKabiseYears2 As Integer
    
    Call SetVariables
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'---> ãÍÇÓÈå ÝÇÕáå ÒãÇäí ÊÇÑíÎ Çæá ÊÇ 1/1/1 ÈÑ ÍÓÈ ÑæÒ
'
     Dim i As Integer
    dgDays1 = dgD1 - 1
    For i = 1 To dgMonth1 - 1
        dgDays1 = dgDays1 + DaysOfMonth3(i - 1)
        
    Next
    
    For i = 1 To dgYear1 - 1
        If IsKabiseGhamari(i) Then
            dgDays1 = dgDays1 + 355
        Else
            dgDays1 = dgDays1 + 354
            
        End If
    Next
   'Debug.Print dgDays1 '505658
   
'--> ãÍÇÓÈå ÝÇÕáå ÒãÇäí ÊÇÑíÎ Ïæã ÊÇ ãÈÏÇ ÊÇÑíÎ ÞãÑí ÈÑ ÍÓÈ ÑæÒ
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    dgDays2 = dgD2 - 1
    For i = 1 To dgMonth2 - 1
        dgDays2 = dgDays2 + DaysOfMonth3(i - 1)
        
    Next
    '
    
    For i = 1 To dgYear2 - 1
        'i:11 350823,i:10 351178 350941, 352595,i=4:3553304, 353658, i=2 354012,
        If i <= 11 Then
            Debug.Print i & ":" & dgDays2
        End If
        If IsKabiseGhamari(Abs(i)) Then
            dgDays2 = dgDays2 + 355
        Else
            dgDays2 = dgDays2 + 354
            
        End If
    Next
    '354366
    DateDifGhamari = dgDays2 - dgDays1
    
    
    
End Function

Public Function IsKabiseGhamari(igYear) As Boolean

' Çíä ÊÇÈÚ¡ íß ÓÇá ÞãÑí ÑÇ ÑÝÊå æ ÊÚííä ãí ßäÏ ßå ÂíÇ ÓÇá ßÈíÓå ÇÓÊ íÇ äå
    B = igYear Mod 30
    
    Select Case B
        Case 2, 5, 7, 10, 13, 16, 18, 21, 24, 26, 29
            IsKabiseGhamari = True
            Exit Function
    End Select
    IsKabiseGhamari = False
    
End Function

Public Function IsKabiseMiladi(imYear) As Boolean
' Çíä ÊÇÈÚ¡ ÔãÇÑå íß ÓÇá ãíáÇÏí ÑÇ ÑÝÊå æ ÊÚííä ãí ßäÏ ßå ÂíÇ Çíä ÓÇá ßÈíÓå ÇÓÊ íÇ äå¿
    
    If (imYear Mod 4) = 0 And (imYear Mod 100) <> 0 Then
        IsKabiseMiladi = True
    ElseIf (imYear Mod 400) = 0 Then
        IsKabiseMiladi = True
    Else
        IsKabiseMiladi = False
        
    End If
    
End Function

Public Function IsKabiseShamsi(isYear) As Boolean
' Çíä ÊÇÈÚ ÔãÇÑå íß ÓÇá ÔãÓí ÑÇ ÑÝÊå æ ÊÚííä ãí ßäÏ ßå ÂíÇ ÓÇá ßÈíÓå ÇÓÊ íÇ äå

    '1, 5, 9, 13, 17, 22, 26, 30
    Dim n As Integer
   n = isYear Mod 33
   Select Case n
    Case 1, 5, 9, 13, 17, 22, 26, 30
        IsKabiseShamsi = True
        Exit Function
    End Select
    IsKabiseShamsi = False
    

End Function

Public Function AddToShamsi(atsYear, atsMonth, atsDay, atsAdd) As String
' Çíä ÊÇÈÚ¡ íß ÊÇÑíÎ ÔãÓí ÑÇ ÑÝÊå æ ÊÚÏÇÏ ÑæÒ ãÚíäí ÑÇ Èå Âä ÇÖÇÝå íÇ ßã ãí ßäÏ
    Dim atsDate, atsAdd2, atsKabiseNum, atsYear2
    
    atsYear2 = atsYear
    atsAdd2 = atsAdd
    
    
    
    If atsAdd2 > 0 Then
        Do Until atsAdd2 < 366
            If IsKabiseShamsi(atsYear2) Then atsKabiseNum = atsKabiseNum + 1
            atsYear2 = atsYear2 + 1
            atsAdd2 = atsAdd2 - 365
        Loop
        
        atsDate = atsYear2 & "/" & atsMonth & "/" & atsDay
        Dim i As Integer
        For i = 1 To atsKabiseNum
            atsDate = PrevDayShamsi(atsDate)
        Next
        For i = 1 To atsAdd2
            atsDate = NextDayShamsi(atsDate)
        Next
    
    ''''''''''''''''''''''''''''
    ElseIf atsAdd2 < 0 Then
        Do Until atsAdd2 > -366
            atsYear2 = atsYear2 - 1
            If IsKabiseShamsi(atsYear2) Then atsKabiseNum = atsKabiseNum + 1
            atsAdd2 = atsAdd2 + 365
            
        Loop
        atsDate = atsYear2 & "/" & atsMonth & "/" & atsDay
        For i = 1 To atsKabiseNum
            atsDate = NextDayShamsi(atsDate)
        Next
        For i = 1 To Abs(atsAdd2)
            atsDate = PrevDayShamsi(atsDate)
            
        Next
    Else
        atsDate = atsYear2 & "/" & atsMonth & "/" & atsDay
    End If
    
    AddToShamsi = atsDate
    AddToShamsi = Format(AddToShamsi, "yyyy/mm/dd")
    

    
End Function

Public Function AddToGhamari(atgYear, atgMonth, atgDay, atgAdd) As String
' Çíä ÊÇÈÚ¡ íß ÊÇÑíÎ ÞãÑí ÑÇ ÑÝÊå æ ÊÚÏÇÏ ÑæÒ ãÚíäí ÑÇ Èå Âä ÇÖÇÝå íÇ ßã ãí ßäÏ
    Dim atgDate, atgAdd2, atgKabiseNum, atgYear2
    
    atgYear2 = atgYear
    atgAdd2 = atgAdd
    
    
    
    If atgAdd2 > 0 Then
        Do Until atgAdd2 < 355
            If IsKabiseGhamari(atgYear2) Then atgKabiseNum = atgKabiseNum + 1
            atgYear2 = atgYear2 + 1
            atgAdd2 = atgAdd2 - 354
        Loop
        
        atgDate = atgYear2 & "/" & atgMonth & "/" & atgDay
        Dim i As Integer
        For i = 1 To atgKabiseNum
            atgDate = PrevDayGhamari(atgDate)
        Next
        For i = 1 To atgAdd2
            atgDate = NextDayGhamari(atgDate)
        Next
    
    ''''''''''''''''''''''''''''
    ElseIf atgAdd2 < 0 Then
        Do Until atgAdd2 > -355
            atgYear2 = atgYear2 - 1
            If IsKabiseGhamari(atgYear2) Then atgKabiseNum = atgKabiseNum + 1
            atgAdd2 = atgAdd2 + 354
            
        Loop
        atgDate = atgYear2 & "/" & atgMonth & "/" & atgDay
        For i = 1 To atgKabiseNum
            atgDate = NextDayGhamari(atgDate)
        Next
        For i = 1 To Abs(atgAdd2)
            atgDate = PrevDayGhamari(atgDate)
            
        Next
    Else
        atgDate = atgYear2 & "/" & atgMonth & "/" & atgDay
    End If
    
    AddToGhamari = atgDate
    AddToGhamari = Format(AddToGhamari, "yyyy/mm/dd")
    

End Function
'''Public Function MiladiToShamsi(ParamArray DatePart()) As Variant
'''
'''' Çíä ÊÇÈÚ¡ íß ÊÇÑíÎ ãíáÇÏí ÑÇ ÑÝÊå Èå ÔãÓí ÊÈÏíá ãí ßäÏ
'''
'''    Call SetVariables
'''
'''    Dim mtsDateDiff As Long
'''    Dim mtsDate
'''
'''    mtsYear = 0
'''    mtsMonth = 0
'''    mtsDay = 0
'''
'''    argnum = UBound(DatePart()) - LBound(DatePart())
'''    Select Case argnum
'''        Case -1
'''            mtsYear = Year(Now())
'''            mtsMonth = Month(Now())
'''            mtsDay = day(Now())
'''
'''        Case 0
'''
'''                mtsYear = Year(DatePart(0))
'''                mtsMonth = Month(DatePart(0))
'''                mtsDay = day(DatePart(0))
'''
'''
'''        Case 2
''''            If VarType(DatePart(0)) = vbInteger _
''''                And VarType(DatePart(1)) = vbInteger _
''''                And VarType(DatePart(2)) = vbInteger Then
'''                mtsYear = DatePart(0)
'''                mtsMonth = DatePart(1)
'''                mtsDay = DatePart(2)
''''            End If
'''
'''        Case Else
'''                MsgBox "Count of Arguments Error"
'''                Exit Function
'''
'''    End Select
'''
'''   If mtsYear = 0 Then
'''    MsgBox "Type of Arguments Error"
'''    Exit Function
'''   End If
'''
'''
'''
'''    If mtsYear < 620 Then
'''        MiladiToShamsi = ""
'''        Exit Function
'''
'''    End If
'''
'''    mtsDate = mtsMonth & "/" & mtsDay & "/" & mtsYear
'''
'''
'''    mtsDateDiff = DateDiff("d", TarikheMabna(0), mtsDate)
'''    '285615
'''    MiladiToShamsi = AddToShamsi(1385, 10, 7, mtsDateDiff)
'''    MiladiToShamsi = Format(MiladiToShamsi, "yyyy/mm/dd")
'''
'''
'''End Function

Public Function MiladiToShamsi(mtsYear, mtsMonth, mtsDay) As String
' Çíä ÊÇÈÚ¡ íß ÊÇÑíÎ ãíáÇÏí ÑÇ ÑÝÊå Èå ÔãÓí ÊÈÏíá ãí ßäÏ

    Call SetVariables
    Dim mtsDateDiff As Long
    Dim mtsDate
    If mtsYear < 620 Then
        MiladiToShamsi = ""
        Exit Function

    End If
    mtsDate = mtsMonth & "/" & mtsDay & "/" & mtsYear


    mtsDateDiff = DateDiff("d", TarikheMabna(0), mtsDate)
    '285615
    MiladiToShamsi = AddToShamsi(1385, 10, 7, mtsDateDiff)





End Function

Public Function MiladiToGhamari(mtgYear, mtgMonth, mtgDay) As String
' Çíä ÊÇÈÚ¡ íß ÊÇÑíÎ ãíáÇÏí ÑÇ ÑÝÊå Èå ÞãÑí ÊÈÏíá ãí ßäÏ
    
    Call SetVariables
    Dim mtgDateDiff As Long
    Dim mtgDate
    
    If mtgYear < 570 Then
        MiladiToGhamari = ""
        Exit Function
    End If
    mtgDate = mtgMonth & "/" & mtgDay & "/" & mtgYear
    
    
    mtgDateDiff = DateDiff("d", TarikheMabna(0), mtgDate)
    
    MiladiToGhamari = AddToGhamari(1427, 12, 7, mtgDateDiff)
    MiladiToGhamari = Format(MiladiToGhamari, "yyyy/mm/dd")
    Debug.Print ""
    
    
    
    
    
End Function

Public Function NextDayShamsi(ndDate) As String
    'Çíä ÊÇÈÚ¡ íß ÊÇÑíÎ ÔãÓí ÑÇ ÑÝÊå æ ÊÇÑíÎ íß ÑæÒ ÈÚÏ ÑÇ Èå ãÇ ãí ÏåÏ
   ' ÝÑãÊ ÊÇÑíÎí ßå Èå Çíä ÊÇÈÚ ÏÇÏå ãí ÔæÏ ÈÇíÏ Èå ÕæÑÊ ÒíÑ ÈÇÔÏ
   ' Y/M/D (Ex: 1385/9/23)
   
   Dim ndDaysOfMonths()
   Dim ndP() As String
   Dim ndP2()
   
   ndP = Split(ndDate, "/")
   'Debug.Print ndP(0) & "," & ndP(1) & "," & ndP(2)
   ndP2 = Array(Val(ndP(0)), Val(ndP(1)), Val(ndP(2)))
   
   
   
   
   ndDaysOfMonths = DaysOfMonth2
   
   'Debug.Print ndDate
   
   If IsKabiseShamsi(ndP2(0)) Then
        ndDaysOfMonths(11) = 30
    Else
        ndDaysOfMonths(11) = 29
    
   End If
   
   If ndP2(1) = 12 And ndP2(2) = ndDaysOfMonths(11) Then
        NextDayShamsi = (ndP2(0) + 1) & "/" & 1 & "/" & 1
    ElseIf ndP2(1) < 12 And ndP2(2) = (ndDaysOfMonths(ndP2(1) - 1)) Then
        NextDayShamsi = ndP2(0) & "/" & (ndP2(1) + 1) & "/" & 1
    Else
       NextDayShamsi = ndP2(0) & "/" & ndP2(1) & "/" & (ndP2(2) + 1)
       
    
    
   End If
   NextDayShamsi = Format(NextDayShamsi, "yyyy/mm/dd")
    
End Function

Public Function NextDayGhamari(ndgDate) As String
    'Çíä ÊÇÈÚ¡ íß ÊÇÑíÎ ÞãÑí ÑÇ ÑÝÊå æ ÊÇÑíÎ íß ÑæÒ ÈÚÏ ÑÇ Èå ãÇ ãí ÏåÏ
   ' ÝÑãÊ ÊÇÑíÎí ßå Èå Çíä ÊÇÈÚ ÏÇÏå ãí ÔæÏ ÈÇíÏ Èå ÕæÑÊ ÒíÑ ÈÇÔÏ
   ' Y-M-D (Ex: 1385-9-23)
   
   Dim ndgDaysOfMonths()
   Dim ndgP() As String
   Dim ndgP2()
   
   ndgP = Split(ndgDate, "/")
   'Debug.Print ndP(0) & "," & ndP(1) & "," & ndP(2)
   ndgP2 = Array(Val(ndgP(0)), Val(ndgP(1)), Val(ndgP(2)))
   
   
   
   
   ndgDaysOfMonths = DaysOfMonth3
   
   'Debug.Print ndDate
   
   If IsKabiseGhamari(ndgP2(0)) Then
        ndgDaysOfMonths(11) = 30
    Else
        ndgDaysOfMonths(11) = 29
    
   End If
   
   If ndgP2(1) = 12 And ndgP2(2) = ndgDaysOfMonths(11) Then
        NextDayGhamari = (ndgP2(0) + 1) & "/" & 1 & "/" & 1
    ElseIf ndgP2(1) < 12 And ndgP2(2) = (ndgDaysOfMonths(ndgP2(1) - 1)) Then
        NextDayGhamari = ndgP2(0) & "/" & (ndgP2(1) + 1) & "/" & 1
    Else
       NextDayGhamari = ndgP2(0) & "/" & ndgP2(1) & "/" & (ndgP2(2) + 1)
       
    
    
   End If
   NextDayGhamari = Format(NextDayGhamari, "yyyy/mm/dd")
    
End Function


Public Function PrevDayShamsi(pdDate) As String
    'Çíä ÊÇÈÚ¡ íß ÊÇÑíÎ ÔãÓí ÑÇ ÑÝÊå æ ÊÇÑíÎ íß ÑæÒ ÞÈá ÑÇ Èå ãÇ ãí ÏåÏ
   ' ÝÑãÊ ÊÇÑíÎí ßå Èå Çíä ÊÇÈÚ ÏÇÏå ãí ÔæÏ ÈÇíÏ Èå ÕæÑÊ ÒíÑ ÈÇÔÏ
   ' Y-M-D (Ex: 1385-9-23)
   
   Dim pdDaysOfMonths()
   Dim pdP() As String
   Dim pdP2()
   
   pdP = Split(pdDate, "/")
   'Debug.Print ndP(0) & "," & ndP(1) & "," & ndP(2)
   pdP2 = Array(Val(pdP(0)), Val(pdP(1)), Val(pdP(2)))
   
   
   
   
   pdDaysOfMonths = DaysOfMonth2
   
   'Debug.Print ndDate
   
   If IsKabiseShamsi(pdP2(0)) Then
        pdDaysOfMonths(11) = 30
    Else
        pdDaysOfMonths(11) = 29
    
   End If
   
   If pdP2(1) = 1 And pdP2(2) = 1 Then
   Dim n As Integer
        n = pdP2(0) - 1
        If IsKabiseShamsi(n) Then
            PrevDayShamsi = (pdP2(0) - 1) & "/" & 12 & "/" & 30
        Else
            PrevDayShamsi = (pdP2(0) - 1) & "/" & 12 & "/" & 29
        End If
    ElseIf pdP2(1) > 1 And pdP2(2) = 1 Then
        PrevDayShamsi = pdP2(0) & "/" & (pdP2(1) - 1) & "/" & (pdDaysOfMonths((pdP2(1) - 2)))
    Else
       PrevDayShamsi = pdP2(0) & "/" & pdP2(1) & "/" & (pdP2(2) - 1)
       
    
    
   End If
   
    PrevDayShamsi = Format(PrevDayShamsi, "yyyy/mm/dd")
End Function



Public Function PrevDayGhamari(pdgDate) As String
    'Çíä ÊÇÈÚ¡ íß ÊÇÑíÎ ÞãÑí ÑÇ ÑÝÊå æ ÊÇÑíÎ íß ÑæÒ ÞÈá ÑÇ Èå ãÇ ãí ÏåÏ
   ' ÝÑãÊ ÊÇÑíÎí ßå Èå Çíä ÊÇÈÚ ÏÇÏå ãí ÔæÏ ÈÇíÏ Èå ÕæÑÊ ÒíÑ ÈÇÔÏ
   ' Y-M-D (Ex: 1385-9-23)
   
   Dim pdgDaysOfMonths()
   Dim pdgP() As String
   Dim pdgP2()
   
   pdgP = Split(pdgDate, "/")
   'Debug.Print ndP(0) & "," & ndP(1) & "," & ndP(2)
   pdgP2 = Array(Val(pdgP(0)), Val(pdgP(1)), Val(pdgP(2)))
   
   
   
   
   pdgDaysOfMonths = DaysOfMonth3
   
   'Debug.Print ndDate
   
   If IsKabiseGhamari(pdgP2(0)) Then
        pdgDaysOfMonths(11) = 30
    Else
        pdgDaysOfMonths(11) = 29
    
   End If
   
   If pdgP2(1) = 1 And pdgP2(2) = 1 Then
   Dim n As Integer
        n = pdgP2(0) - 1
        If IsKabiseGhamari(n) Then
            PrevDayGhamari = (pdgP2(0) - 1) & "/" & 12 & "/" & 30
        Else
            PrevDayGhamari = (pdgP2(0) - 1) & "/" & 12 & "/" & 29
        End If
    ElseIf pdgP2(1) > 1 And pdgP2(2) = 1 Then
        PrevDayGhamari = pdgP2(0) & "/" & (pdgP2(1) - 1) & "/" & (pdgDaysOfMonths((pdgP2(1) - 2)))
    Else
       PrevDayGhamari = pdgP2(0) & "/" & pdgP2(1) & "/" & (pdgP2(2) - 1)
       
    
    
   End If
   PrevDayGhamari = Format(PrevDayGhamari, "yyyy/mm/dd")
    
End Function
Private Sub SetVariables()
    DateFormats = Array("ãíáÇÏí", "ÔãÓí", "ÞãÑí")
    CurrentYear = Array(2006, 1385, 1427)
    TarikheMabna = Array("12/28/2006", "1385-10-7", "1427-12-7")
    YearNames = Array("ãÇÑ", "ÇÓÈ", "æÓÝäÏ", "ãíãæä", "ãÑÛ", "Ó", "Îæß", "ãæÔ", "Çæ", "áä", "ÎÑæÔ", "äåä")
    WeekDays1 = Array("Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    WeekDays2 = Array("ÔäÈå", "íßÔäÈå", "ÏæÔäÈå", "Óå ÔäÈå", "åÇÑÔäÈå", "äÌ ÔäÈå", "ÌãÚå")
    WeekDays3 = Array("ÇáÓÈÊ", "ÇáÇÍÏ", "ÇáÇËäíä", "ÇáËáÇËÇÁ", "ÇáÇÑÈÚÇÁ", "ÇáÎãíÓ", "ÇáÌãÚå")
    
    'Taghvim-e Miladi
    MonthNames1 = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    DaysOfMonth1 = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)

    'Taghvim-e Shamsi
    MonthNames2 = Array("ÝÑæÑÏíä", "ÇÑÏíÈåÔÊ", "ÎÑÏÇÏ", "ÊíÑ", "ãÑÏÇÏ", "ÔåÑíæÑ", "ãåÑ", "ÂÈÇä", "ÂÐÑ", "Ïí", "Èåãä", "ÇÓÝäÏ")
    DaysOfMonth2 = Array(31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29)
    
    'Taghvim-e Ghamari
    MonthNames3 = Array("ãÍÑã", "ÕÝÑ", "ÑÈíÚ ÇáÇæá", "ÑÈíÚ ÇáËÇäí", "ÌãÇÏí ÇáÇæá", "ÌãÇÏí ÇáËÇäí", "ÑÌÈ", "ÔÚÈÇä", "ÑãÖÇä", "ÔæÇá", "ÐæÇáÞÚÏå", "ÐæÇáÍÌå")
    DaysOfMonth3 = Array(30, 29, 30, 29, 30, 29, 30, 29, 30, 29, 30, 29)
End Sub

Public Function ShamsiToGhamari(stgYear, stgMonth, stgDay) As String
' Çíä ÊÇÈÚ¡ íß ÊÇÑíÎ ÔãÓí ÑÇ Èå ÞãÑí ÊÈÏíá ãí äãÇíÏ

    Call SetVariables
    Dim stgDateDiff As Long
    Dim stgDate2, stgYear2, stgMonth2, stgDay2
    
    
    stgDateDiff = DateDifShamsi(1385, 10, 7, stgYear, stgMonth, stgDay)
    stgDate2 = AddToGhamari(1427, 12, 7, stgDateDiff)
     
    ShamsiToGhamari = stgDate2
    
    ShamsiToGhamari = Format(ShamsiToGhamari, "yyyy/mm/dd")
    
    
End Function

Public Function ShamsiToMiladi(stmYear, stmMonth, stmDay) As String
' Çíä ÊÇÈÚ¡ ÊÇÑíÎ åÇí ÔãÓí ÑÇ Èå ãíáÇÏí ÊÈÏíá ãí ßäÏ
    Call SetVariables
   
    Dim stmDateDiff As Long
    Dim stmDate2, stmYear2, stmMonth2, stmDay2
    
    
    stmDateDiff = DateDifShamsi(1385, 10, 7, stmYear, stmMonth, stmDay)
    stmDate2 = DateAdd("d", stmDateDiff, "12/28/2006") 'DateTime.DateSerial(2006, 12, 28 + stmDateDiff)
    stmYear2 = Year(stmDate2)
    stmMonth2 = Month(stmDate2)
    stmDay2 = day(stmDate2)
    
    ShamsiToMiladi = stmYear2 & "/" & stmMonth2 & "/" & stmDay2
    
    ShamsiToMiladi = Format(ShamsiToMiladi, "yyyy/mm/dd")
    
End Function

Public Function GhamariToMiladi(gtmYear, gtmMonth, gtmDay) As String
' Çíä ÊÇÈÚ¡ ÊÇÑíÎ åÇí ÞãÑí ÑÇ Èå ãíáÇÏí ÊÈÏíá ãí ßäÏ
   If gtmYear = "" Or gtmMonth = "" Or gtmDay = "" Then
        GhamariToMiladi = ""
        Exit Function
    End If
    Call SetVariables
    Dim gtmDateDiff As Long
    Dim gtmDate2, gtmYear2, gtmMonth2, gtmDay2
    
    
    gtmDateDiff = DateDifGhamari(1427, 12, 7, gtmYear, gtmMonth, gtmDay)
    gtmDate2 = DateAdd("d", gtmDateDiff, "12/28/2006") 'DateTime.DateSerial(2006, 12, 28 + stmDateDiff)
    gtmYear2 = Year(gtmDate2)
    gtmMonth2 = Month(gtmDate2)
    gtmDay2 = day(gtmDate2)
    
    GhamariToMiladi = gtmYear2 & "/" & gtmMonth2 & "/" & gtmDay2
    GhamariToMiladi = Format(GhamariToMiladi, "yyyy/mm/dd")
    
    
End Function


Public Function GhamariToShamsi(gtsYear, gtsMonth, gtsDay) As String
' Çíä ÊÇÈÚ¡ ÊÇÑíÎ åÇí ÞãÑí ÑÇ Èå ÔãÓí ÊÈÏíá ãí ßäÏ
   If gtsYear = "" Or gtsMonth = "" Or gtsDay = "" Then
        GhamariToShamsi = ""
        Exit Function
    End If
    
    
    Call SetVariables
    Dim gtsDateDiff As Long
    Dim gtsDate2, gtsYear2, gtsMonth2, gtsDay2
    
    
    gtsDateDiff = DateDifGhamari(1427, 12, 7, gtsYear, gtsMonth, gtsDay)
    '151292
    
    gtsDate2 = AddToShamsi(1385, 10, 7, gtsDateDiff)
    GhamariToShamsi = gtsDate2
    GhamariToShamsi = Format(GhamariToShamsi, "yyyy/mm/dd")
    
End Function


