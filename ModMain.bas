Attribute VB_Name = "Modmain"
Option Explicit

Public rs As New Recordset
Public CNS As String
'
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal Flags As Long) As Long
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public Const HKL_NEXT = 1
Public Const HKL_PREV = 0


Sub Main()
  If Trim(GetSetting("HKEY_CURRENT_USER", "xMehrvarzan", "PASSWORD", "x2x")) <> Trim("x2x") Then
     FrmGetPass.LblPass = Trim(GetSetting("HKEY_CURRENT_USER", "xMehrvarzan", "PASSWORD", "x2x"))
     FrmGetPass.Show 1
  Else 'No Stored PassWord
     FrmStart.Show
  End If
    
    CNS = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbShayan.mdb;Persist Security Info=False"
    FrmStart.Show
End Sub
 
Public Function CompactDB(pFileName As String) As Boolean
On Error GoTo ErrH
Dim CONN As New JRO.JetEngine
Dim ConnstringSorg As String, ConnstringDest As String

' Ensure file is not read only
SetAttr pFileName, vbNormal
ConnstringSorg = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
pFileName & ";User ID=;Password=;"
ConnstringDest = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
App.Path & "\Temp.mdb" & ";Jet OLEDB:Engine Type=5;"

Screen.MousePointer = vbHourglass
CONN.CompactDatabase ConnstringSorg, ConnstringDest
Screen.MousePointer = vbDefault

'Copia il file compattato.
Kill pFileName
FileCopy App.Path & "\Temp.mdb", pFileName
Kill App.Path & "\Temp.mdb"

Set CONN = Nothing
CompactDB = True
Exit Function
ErrH:
Screen.MousePointer = vbDefault
Debug.Print Err.Description
End Function

Public Function DigitGrouping(Number As Currency) As String
 Dim S$, T$
 Dim i%, L%
 
 S = CStr(Number)
 L = Len(S)

 For i = 1 To L \ 3
     T = "," & Right(S, 3) & T
     S = Left(S, Len(S) - 3)
 Next
 If L Mod 3 <> 0 Then T = S & T ' if Not 3,6,9,12,... digit
 If Left(T, 1) = "," Then T = Mid(T, 2) ' ,150,000
 DigitGrouping = T
End Function
