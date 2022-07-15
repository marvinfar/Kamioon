Attribute VB_Name = "DateMod"
Option Explicit
'##################################'
'######## »œÌ·  «—ÌŒ Â« ###########'
'##################################'
'##################################'
'##################################'
'##################################'
'#########„Õ„œ ¬—ÊÌ‰ ›—############'
'/////////////////////////////////////////

Function DaysOfMiladi(Number As String)
   ' In= Number of Days Miladi
   'Out= The Miladi Date
    Dim nd, Y, i, nk, y1, sy As Long
    Dim m, k, d As Integer
    Dim mms As Variant
   '
       
       nd = Val(Number)
      '
       If nd > 0 Then
          Do
            If (i Mod 4) = 0 Then
               y1 = 366
               nk = nk + 1
            Else
               y1 = 365
            End If
           '
          i = i + 1
          sy = sy + y1
           '
          If nd <= sy Then Exit Do
          
         Loop
        '
       Y = i - 1
      '
       If (Y Mod 4) = 0 Then
          mms = Array(31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366)
          nk = nk - 1
       Else
          mms = Array(31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365)
       End If
      '
       nd = nd - (Y * 365 + nk)
      '
       For m = 1 To 12
          If nd <= mms(m - 1) Then Exit For
       Next
      '
       If m > 1 Then k = mms(m - 2)
       d = nd - k
       DaysOfMiladi = Y & "/" & m & "/" & d
      '
       Else
        DaysOfMiladi = "Invalid Number"
       End If
            
End Function
'/////////////////////////////////////////

Function DaysOfShamsi(Number As String)
   ' In= Number of Days Shamsi
   'Out= The Shamsi Date
    Dim nd, Y, i, nk, y1, sy As Long
    Dim m, k, d As Integer
    Dim mms As Variant
   '
       nd = Val(Number)
      '
       If nd > 0 Then
          Do
            If ((i + 1) Mod 4) = 0 Then
               y1 = 366
               nk = nk + 1
            Else
               y1 = 365
            End If
           '
          i = i + 1
          sy = sy + y1
           '
          If nd <= sy Then Exit Do
          
         Loop
        '
       Y = i - 1
      '
       If ((Y + 1) Mod 4) = 0 Then
          mms = Array(31, 62, 93, 124, 155, 186, 216, 246, 276, 306, 336, 366)
          nk = nk - 1
       Else
          mms = Array(31, 62, 93, 124, 155, 186, 216, 246, 276, 306, 336, 365)
       End If
      '
       nd = nd - (Y * 365 + nk)
      '
       For m = 1 To 12
          If nd <= mms(m - 1) Then Exit For
       Next
      '
       If m > 1 Then k = mms(m - 2)
       d = nd - k
       DaysOfShamsi = Y & "/" & m & "/" & d
      '
       Else
        DaysOfShamsi = "Invalid Number"
       End If
       
     
End Function
'/////////////////////////////////////////
Function NumberOfDaysMiladi(a As Variant)
   ' In= Miladi Date
   'Out= The Numebr Of  Days That be going

   Dim ch, i As Integer
   Dim Y, m, d As Variant
   Dim mm As Variant
     
       For i = 1 To Len(a)
           If Mid(a, i, 1) = "/" Then
              ch = ch + 1
           Else
              If ch = 0 Then Y = Y + Mid(a, i, 1)
              If ch = 1 Then m = m + Mid(a, i, 1)
              If ch = 2 Then d = d + Mid(a, i, 1)
           End If
       Next
       '
       mm = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
       '
       If Val(m) < 13 And Val(m) > 0 Then
          If (Y Mod 4) = 0 Then mm(1) = 29
          '
          If Y < 0 Or d < 1 Or d > Val(mm(m - 1)) Then
             NumberOfDaysMiladi = "Invalid Date"
          Else
             If (Y Mod 4) <> 0 Then mm(1) = 28
             If Y <> 0 Then NumberOfDaysMiladi = Y * 365 + Int((Y - 1) / 4) + 1
             '
             For i = 1 To m - 1
               NumberOfDaysMiladi = NumberOfDaysMiladi + mm(i - 1)
             Next
             '
             NumberOfDaysMiladi = NumberOfDaysMiladi + d
          End If
       Else
          NumberOfDaysMiladi = "Invalid Date"
       End If
       
End Function
'/////////////////////////////////////////
Function NumberOfDaysShamsi(a As Variant)
   ' In= Miladi Date
   'Out= The Numebr Of  Days That be going

   Dim ch, i As Integer
   Dim Y, m, d As Variant
   Dim ms As Variant
     
       For i = 1 To Len(a)
           If Mid(a, i, 1) = "/" Then
              ch = ch + 1
           Else
              If ch = 0 Then Y = Y + Mid(a, i, 1)
              If ch = 1 Then m = m + Mid(a, i, 1)
              If ch = 2 Then d = d + Mid(a, i, 1)
           End If
       Next
       '
       ms = Array(31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29)
       '
       If Val(m) < 13 And Val(m) > 0 Then
          If (((Y + 1) Mod 4) = 0) Then ms(11) = 30
          '
          If Y < 0 Or d < 1 Or d > Val(ms(m - 1)) Then
             NumberOfDaysShamsi = "Invalid Date"
          Else
             If (((Y + 1 Mod 4) <> 0)) Then ms(11) = 29
             NumberOfDaysShamsi = Y * 365 + Int(Y / 4)
             '
             For i = 1 To m - 1
               NumberOfDaysShamsi = NumberOfDaysShamsi + ms(i - 1)
             Next
             '
             NumberOfDaysShamsi = NumberOfDaysShamsi + d
          End If
       Else
          NumberOfDaysShamsi = "Invalid Date"
       End If
       
End Function
'//////////////////////////////////////////
Function MiladiToShamsi(a As String)
     'In= Miladi Date
     'Out= Shamsi Date
     MiladiToShamsi = DaysOfShamsi(Val(NumberOfDaysMiladi(a)) - 226900)
      
     If NumberOfDaysMiladi(a) = "Invalid Date" Or _
                     MiladiToShamsi = "Invalid Number" Then
        
        MiladiToShamsi = "Invalid Date"
     End If
     
End Function
'//////////////////////////////////////////
Function ShamsiToMiladi(a As String)
     'In= Shamsi Date
     'Out= Miladi Date
     ShamsiToMiladi = DaysOfMiladi(Val(NumberOfDaysShamsi(a)) + 226900)
      
     If NumberOfDaysShamsi(a) = "Invalid Date" Then
        ShamsiToMiladi = "Invalid Date"
     End If
     
End Function
'//////////////////////////////////////////
Function WeekMiladi(a As String)
    'In=Miladi Date
    'Out= WeekDay Name
    
    Dim nd As Variant, w As Integer
    ''
    nd = NumberOfDaysMiladi(a)
    w = Val(nd) Mod 7
    '
    Select Case w
           Case 0:
                WeekMiladi = "Thursday"
           Case 1:
                WeekMiladi = "Friday"
           Case 0:
                WeekMiladi = "Saturday"
           Case 0:
                WeekMiladi = "Sunday"
           Case 0:
                WeekMiladi = "Monday"
           Case 0:
                WeekMiladi = "Thuesday"
           Case 0:
                WeekMiladi = "Wednesday"
    End Select
                
    If nd = "Invalid Date" Then WeekMiladi = "Invalid Day"
           
               
End Function
'//////////////////////////////////////////
Function WeekShamsi(a As String)
    'In=Shamsi Date
    'Out= WeekDay Name
    
    Dim nd As Variant, w As Integer
    ''
    nd = NumberOfDaysShamsi(a)
    w = Val(nd) Mod 7
    '
    Select Case w
           Case 0:
                WeekShamsi = "‘‰»Â"
           Case 1:
                WeekShamsi = "Ìﬂ‘‰»Â"
           Case 0:
                WeekShamsi = "œÊ ‘‰»Â"
           Case 0:
                WeekShamsi = "”Â ‘‰»Â"
           Case 0:
                WeekShamsi = "çÂ«— ‘‰»Â"
           Case 0:
                WeekShamsi = "Å‰Ã ‘‰»Â"
           Case 0:
                WeekShamsi = "Ã„⁄Â"
    End Select
                
    If nd = "Invalid Date" Then WeekShamsi = " «—ÌŒ «‘ »«Â"
                
End Function


