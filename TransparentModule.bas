Attribute VB_Name = "TransparentModule"
Option Explicit

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Const RGN_OR = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Global Const CB_ERR = -1
    Global Const CB_FINDSTRING = &H14C

 Public fso As New FileSystemObject
 Public SpFol_ As SpecialFolderConst
 
  Private rs As ADODB.Recordset
  Private RS1 As ADODB.Connection

  Public MainDate As String
  Dim dd As String
  Public GridFont As String
  Public VarCOstan As String * 3
  Public VarFYear As String * 4
  Public SelectCityName, SelectOfficeName As String * 30
  Public SelectCityCode, SelectCityCodeHlp, SelectOfficeCode, SelectOfficeCodeHlp, SelectOfficeType As String * 3
  Public WichForm As Form
 '
  Public UserName As String * 30
  Public PassWord As String * 20
  Public SuperV, SG, endyearhlp As Boolean  'SuperViser and SarGrooh , endofyear
 '
  Public UsCode As String * 3
  Public ProgCode As String * 3

'IP And Computer Name And Port'''''''''''''''
  
  Public usernamehlp As String
  Public ComputerNameHlp As String
  Public IPhlp As String
 '
  Public IPName As String
  Public ComputerName As String
  Public PortNumber As Integer
  
''''''''''''''''
  
  Public Enum MsgSubject
        [Save] = 1
        [Delete] = 2
        [Info] = 3
        [Warning] = 4
        [Crit] = 5
  End Enum
  
  Public Enum AccessEnum
        [ReadOnly] = 1
        [ReadWrite] = 2
  End Enum
  Public ChkUser  As AccessEnum

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public ArrTKeyHlp(1 To 100) As String


Public cns, DataBaseName, DBPathPishtaz, ServerName, us, ps As String

Private Type POSTConfigRecord 'A Record For hold The PostConfig Values
        dbName As String * 50
        UsName As String * 50
        PassW As String * 50
        ComputerName As String * 50
End Type

'Dim InsRec As InstallRecord
Dim PConfigRec As POSTConfigRecord
Dim F_Ins, F_Config As Integer 'File Variable
'''
Global FirstRes As TRes
'
Public pfn0 As String

 'Public Const CNS = "MSDASQL.1; persist security info=False; driver=Sql server; DataBase=Post; uid=; pwd=sa; server=sqlserver"



  Public Function MakeRegion(picSkin As PictureBox) As Long
       
    Dim X As Long, Y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean
    Dim hdc As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hdc = picSkin.hdc
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    X = Y = StartLineX = 0
    
    
    
    TransparentColor = GetPixel(hdc, 0, 0)
    
    For Y = 0 To PicHeight - 1
        For X = 0 To PicWidth - 1
            
            If GetPixel(hdc, X, Y) = TransparentColor Or X = PicWidth Then
                
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, Y, X, Y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                      
                        DeleteObject LineRegion
                    End If
                End If
            Else
                
                If Not InLine Then
                    InLine = True
                    StartLineX = X
                End If
            End If
        Next
    Next
    
    MakeRegion = FullRegion
End Function

Sub Main()
    On Error Resume Next
    Dim Connect0 As New ADODB.Connection
    Dim DBPathName  As String
    
   '
    ' enum DeskTop Setting first
      FirstRes = FirstSetting
      ChangeResolution 1024, 768
   '
    Set Connect0 = New ADODB.Connection

    'open dabase config
    SpFol_ = WindowsFolder
    F_Config = FreeFile
    'MsgBox fso.GetSpecialFolder(SpFol_)
  '  Open fso.GetSpecialFolder(SpFol_) + "\PostConfig.dat" For Random As #F_Config Len = LenB(PConfigRec)
    'Get #F_Config, 1, PConfigRec
  '  DataBaseName = Trim(PConfigRec.dbName)
  '  ServerName = Trim(PConfigRec.ComputerName)
    'MsgBox DataBaseName
    'Us = Trim(PConfigRec.UsName)
    'Ps = Trim(PConfigRec.PassW)
  
   '  Open fso.GetSpecialFolder(SpFol_) + "\PostConfig.txt" For Input As #1
  '   Input #1, DataBaseName, UserName, PassWord, ServerName
   '  Close #1
           
     Open fso.GetSpecialFolder(SpFol_) + "\PostConfig.dat" For Input As 1
  '  Open "c:\windows\PostConfig.dat" For Input As 1
    Input #1, DataBaseName, ServerName, us, ps ', GridDateHlp ' , PishtazAddress
    Close #1
    
    DataBaseName = Trim(Left(DataBaseName, 30))
    ServerName = Trim(Left(ServerName, 30))
    us = Trim(Left(us, 30))
    ps = Trim(Left(ps, 30))
    'PishtazAddress = Trim(PishtazAddress)
' DataBaseName = "postins0": ServerName = "node1"
      If DataBaseName = Empty Or ServerName = Empty Then
         MsgBox "«‘ò«· œ—  ‰ŸÌ„«  »«‰ò «ÿ·«⁄« Ì", vbCritical, Space(20) + "Œÿ«"
         Exit Sub
    End If
 '''''''''''''Get IP and Computer Name''''''''''
   With MDIForm1.Winsock1
        IPName = .LocalIP
        ComputerName = .LocalHostName
        PortNumber = .LocalPort
   End With

'''''''''''''''''''''''''''''''''''''''''''''''
'ServerName = "sqlserver"
'ps = "sa"
'us = "sa"
'DataBaseName = "post100101101138411"
   'cns = "MSDASQL.1; persist security info=False; driver=Sql server; ; ; ; server=" + ServerName
 ' ServerName = "169.254.129.125"
    cns = "MSDASQL.1; persist security info=False; driver=Sql server; ; uid=" + us + "; pwd=" + ps + "; server=" + ServerName
      DBPathName = App.Path + "\DataBase\" + DataBaseName
      DBPathPishtaz = App.Path + "\DataBase"
      Connect0.ConnectionString = cns
      Connect0.CommandTimeout = 50
      Connect0.CursorLocation = adUseClient
      Connect0.IsolationLevel = adXactReadUncommitted
      Connect0.Open
      Connect0.Execute ("sp_attach_db @dbname = N'" & DataBaseName & _
               "',@filename1 = N'" & DBPathName & ".mdf' , @filename2 = N'" & DBPathName & ".ldf'")
      Connect0.Close
   
'
     cns = "MSDASQL.1; persist security info=False; driver=Sql server; DataBase=" + DataBaseName + "; uid=" + us + "; pwd=" + ps + "; server=" + ServerName
  
    'Open The Bank
    Call SetRecordSet(rs, RS1)
    
    rs.Open "SELECT * from StartPost where flag=1", cns, adOpenStatic
    If Not rs.EOF And rs("CodeCity") <> Empty Then
          SelectCityCodeHlp = rs("CodeCity")
          SelectOfficeCodeHlp = rs("CodeOffice")
       End If
    rs.MoveFirst

     'Set Super Viser if SuperViser Was Not Exist
          If IsNull(Trim(rs("SpVsPass"))) Or Trim(rs("SpVsPass")) = Empty Then
             FrmSetSV.Show 1
          End If
          
      VarCOstan = rs("CodeOstan") 'Get Fix CodeOstan From StartPost Table
      VarFYear = rs("FYear") 'Get Fix Year From StartPost Table
  '    GridDateHlp = Trim(RS("GridDate"))
      GridFont = Trim(rs("GridFont"))
      endyearhlp = rs("endyear")
     MDIForm1.Show
     MDIForm1.Caption = Space(5) + " ”Ì” „  —«›Ìﬂ Ê œ—¬„œ (TIS)" + Space(50) + " «” «‰ : êÌ·«‰ " + Space(15) + " ”«· „«·Ì : " + Trim(VarFYear)
'''''''''''''''''''''''Change Date To Shamsi''''''''''''''
     MainDate = Format(Date, "yyyy/mm/dd")
    MainDate = MiladiToShamsi(MainDate)
     
     If Mid(MainDate, 7, 1) = "/" Then
        dd = Left(MainDate, 5) + "0" + Mid(MainDate, 6, 1)
        MainDate = dd + "/" + Mid(MainDate, 8)
     Else
        dd = Left(MainDate, 7)
        MainDate = Left(MainDate, 8) + Mid(MainDate, 9)
     End If
     If Val(Mid(MainDate, 9)) < 10 Then
        MainDate = dd + "/0" + Right(MainDate, 1)
     Else
        MainDate = dd + "/" + Right(MainDate, 2)
     End If
  
     
    
     rs.Close
     rs.Open "SELECT * from Office WHERE Fyear='" + Trim(VarFYear) + "'", cns, adOpenStatic
      If Not rs.EOF Then
         rs.Close
         FrmSelectOffice.Show 1
      Else
         rs.Close
         FrmGetPassWord.Show 1
      End If

      'If Trim(LCase(ServerName)) <> (LCase(ComputerName)) Then ' remove Copy menu if computer is not server!!
         'If Trim(ServerName) <> IPName Then
          '  MDIForm1.MnuBaseDataOut.Enabled = False
          '  MDIForm1.MnuBaseDataInp.Enabled = False
          '  MDIForm1.MnuTransout.Enabled = False
          '  MDIForm1.MnuTransInp.Enabled = False
          '  MDIForm1.MnuTransoutCity.Enabled = False
          '  MDIForm1.MnuTransInpCity.Enabled = False
      '   End If
     ' End If
      
      rs.Open "SELECT * from Styleobject ", cns, adOpenStatic
      If rs.EOF Then
         FrmOption.Show
      End If
      rs.Close
End Sub
'======================================================================================================

'Public Sub ShowMessage(Msg As String, Title As String, Sbjct As MsgSubject, X, Y As Integer)
'   Dim r, m As String
'
'      With FrmMsg
'           .ImgCrit.Visible = False
'           .ImgWar.Visible = False
'           .ImgQ.Visible = False
'           .ImgInfo.Visible = False
'      End With
    '
'      FrmMsg.Caption = Title
'    '
'      If Len(Msg) > 26 Then 'Split Text in 2 Line
'         r = Right(Msg, 20)
'         m = Mid(Msg, 20, Len(Msg))
'         Msg = r + vbCrLf + m
'      End If
'    '
'      FrmMsg.LblMsg.Caption = Msg 'Show Message in Form
'     'Set Button
'       With FrmMsg
'        SELECT Case Sbjct
'               Case 1 'Save
'                   .CmdOk.Visible = False
'                   .CmdYes.Visible = True
'                   .CmdNo.Visible = True
'                   .ImgQ.Visible = True
'               Case 2 'delete
'                   .CmdOk.Visible = False
'                   .CmdYes.Visible = True
'                   .CmdNo.Visible = True
'                   .ImgQ.Visible = True
'               Case 3 'Info
'                   .CmdOk.Visible = True
'                   .CmdYes.Visible = False
'                   .CmdNo.Visible = False
'                   .ImgInfo.Visible = True
'               Case 4 'Warning
'                   .CmdOk.Visible = True
'                   .CmdYes.Visible = False
'                   .CmdNo.Visible = False
'                   .ImgWar.Visible = True
'               Case 5 'Crit
'                   .CmdOk.Visible = True
'                   .CmdYes.Visible = False
'                   .CmdNo.Visible = False
'                   .ImgCrit.Visible = True
'               Case Else
'                   MsgBox "Œÿ«"
'        End SELECT
'       End With
'      If X = 0 And Y = 0 Then ' Set Form In The Left And Top  of Screen
'         FrmMsg.Left = (MDIform1.Width - FrmMsg.Width) / 2
'         FrmMsg.Top = (MDIform1.Height - FrmMsg.Height) / 2
'      Else
'         FrmMsg.Left = X
'         FrmMsg.Top = Y
'      End If
'      FrmMsg.Show 1
'End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FdateA(Frm As Form) As String
  Dim X, Y, Pos As Integer
  Dim r, m, TxtD, TxtM As String
    'Define Date

 ''''''''''''''Convert TxtDate To Split Text'''''''''''''''''''
  
  Pos = InStr(1, Frm.TxtDate, "/")
  m = Mid(Frm.TxtDate, Pos + 1, 2) 'Get Mounth Number
  If Right(m, 1) = "/" Then 'Find "/" Char in Mounth
     TxtM = Left(m, 1)
  Else
     TxtM = m
  End If
  
  
  r = Right(Frm.TxtDate, 2)
  
  If Left(r, 1) = "/" Then 'Find "/" Char in Day
     TxtD = Right(r, 1)
  Else
     TxtD = r
  End If
''''''''''''''''''
  Y = Val(TxtM)
 ''
  X = Val(TxtD)
  If Y <= 6 Then
     If X = 31 Then
        If Y < 12 Then
           Y = Y + 1
           X = 1
        End If
     Else
        X = X + 1
     End If
  Else
     If X = 30 Then
        If Y < 12 Then
           Y = Y + 1
           X = 1
        End If
     Else
        X = X + 1
     End If
 End If
''
  If Y < 10 Then
     TxtM = "0" + Trim(Str(Y))
  Else
     TxtM = Trim(Str(Y))
  End If
''
  If X < 10 Then
     TxtD = "0" + Trim(Str(X))
  Else
     TxtD = Trim(Str(X))
  End If
  
  Frm.TxtDate = Left(Frm.TxtDate, 5) + TxtM + "/" + TxtD
  
  FdateA = Frm.TxtDate
 
End Function
'''''''''''''''''''''''
Public Function FdateB(Frm As Form) As String
  Dim X, Y, Pos As Integer
  Dim r, m, TxtD, TxtM As String
    'Define Date

 ''''''''''''''Convert TxtDate To Split Text'''''''''''''''''''
  
  Pos = InStr(1, Frm.TxtDate, "/")
  m = Mid(Frm.TxtDate, Pos + 1, 2) 'Get Mounth Number
  If Right(m, 1) = "/" Then 'Find "/" Char in Mounth
     TxtM = Left(m, 1)
  Else
     TxtM = m
  End If
  
  
  r = Right(Frm.TxtDate, 2)
  
  If Left(r, 1) = "/" Then 'Find "/" Char in Day
     TxtD = Right(r, 1)
  Else
     TxtD = r
  End If
''''''''''''''''''
  Y = Val(TxtM)
 ''
  X = Val(TxtD)
  If Y <= 6 Then
     If X = 1 Then
        If Y > 1 Then
           Y = Y - 1
           X = 31
        End If
     Else
        X = X - 1
     End If
  Else
     If X = 1 Then
        If Y > 1 Then
           Y = Y - 1
           If Y = 6 Then
              X = 31
           Else
              X = 30
           End If
        End If
     Else
        X = X - 1
     End If
 End If
''
  If Y < 10 Then
     TxtM = "0" + Trim(Str(Y))
  Else
     TxtM = Trim(Str(Y))
  End If
''
  If X < 10 Then
     TxtD = "0" + Trim(Str(X))
  Else
     TxtD = Trim(Str(X))
  End If
  
  Frm.TxtDate = Left(Frm.TxtDate, 5) + TxtM + "/" + TxtD
  
  FdateB = Frm.TxtDate
 
End Function

Public Sub SetObjectStyle(Frm As Form)
 ' On Error Resume Next
  'Change All Of Buttons And TextBoxs Style To Values Of Bank
    Dim oCtrl As Control
' Open From Bank To Set Style Of Buttons!
  '  RS.Close
    rs.Open "SELECT * from StyleObject WHERE UN='" _
           + Trim(UserName) + "'", cns, adOpenStatic
    
    If Not rs.EOF Then
   '''''''''''''''Button Style''''''''
       For Each oCtrl In Frm.Controls
           If TypeOf oCtrl Is TypeButton Then 'Set Style Button
              oCtrl.ButtonType = Val(Trim(rs("BtnCode")))
              oCtrl.ColorScheme = Val(Trim(rs("BtnColor")))
              oCtrl.Font.Name = Trim(rs("BtnFont"))
              oCtrl.Font.Size = Trim(rs("BtnFSize"))
              oCtrl.Refresh
           End If
       Next
   '''''''''''''''Text Style''''''''
       For Each oCtrl In Frm.Controls
           If TypeOf oCtrl Is TextBox Then
              oCtrl.Font.Name = Trim(rs("TxtFont"))
              oCtrl.Font.Size = rs("TxtFSize") 'Int
              
              oCtrl.BackColor = Trim(rs("TxtBColor"))
              oCtrl.ForeColor = Trim(rs("TxtFColor"))
       
           End If
       Next
    Else
   '''''''''''''''Button Style''''''''
       For Each oCtrl In Frm.Controls
           If TypeOf oCtrl Is TypeButton Then 'Set Style Button
              oCtrl.ButtonType = 3 ' Val(Trim(Rs("BtnCode")))
              oCtrl.ColorScheme = 3 ' Val(Trim(Rs("BtnColor")))
              oCtrl.Font.Name = "Traditional Arabic" ' Trim(Rs("BtnFont"))
              oCtrl.Font.Size = 12 'Trim(Rs("BtnFSize"))
              oCtrl.Refresh
           End If
       Next
   '''''''''''''''Text Style''''''''
       For Each oCtrl In Frm.Controls
           If TypeOf oCtrl Is TextBox Then
              oCtrl.Font.Name = "Traditional Arabic" ' Trim(Rs("TxtFont"))
              oCtrl.Font.Size = 14 ' Rs("TxtFSize") 'Int
              'oCtrl.BackColor = Trim(Rs("TxtBColor"))
              'oCtrl.ForeColor = Trim(Rs("TxtFColor"))
       
           End If
       Next
    
    End If
    rs.Close
  
End Sub

Public Sub HelpKey(Index As Byte, ParamArray HlpText() As Variant)
  Dim i As Integer
   'Set The HelpKey Text In Status bar With ParamArray

   With MDIForm1.StatusBar1
        
        .Panels(2).Text = Empty
        
        For i = 1 To Index
           .Panels(2).Text = .Panels(2).Text + HlpText(i - 1)
        Next
   End With
End Sub

Public Sub MyHelpKey()
   'Call Help Key For This Subject
   
   Call HelpKey(6, "  F11 «‰ Œ«»  «—ÌŒ  ", "  F4 Œ—ÊÃ  ", _
        "  Esc «‰’—«›  ", "  F9 ÊÌ—«Ì‘  ", _
        "  Del Õ–›  ", "  Ins À»   ")
End Sub

Public Sub MYLoadImage(Img As Image, StrAdrs As String)
   
 
'    If Fso.FolderExists(App.Path + "\Image") Then
       If fso.FileExists(StrAdrs) Then
          Img.Picture = LoadPicture(StrAdrs)
       Else
          MsgBox "›«Ì· ⁄ò” „ÊÃÊœ ‰Ì” ", vbExclamation, ""
          Exit Sub
       End If
'    Else
'       MsgBox "ÅÊ‘Â ⁄ò”Â« „ÊÃÊœ ‰Ì” ", vbExclamation, ""
'       Exit Sub
'    End If
End Sub


Public Function Cancate(Index As Byte, ParamArray StrN() As Variant)
   Dim i As Byte
     ' Sumation n String
     ' Example "Ali"+str(VarT)+"Mahdi"
     
       For i = 1 To Index
           If VarType(StrN(i - 1)) = vbString Then
              Cancate = Cancate + StrN(i - 1)
           Else
              Cancate = Cancate + Str(StrN(i - 1))
           End If
       Next 'for
End Function

Public Sub SetRecordSet(RecSet As ADODB.Recordset, RSConnection As ADODB.Connection)
'   On Error Resume Next
   
    Set RecSet = New ADODB.Recordset
    Set RSConnection = New ADODB.Connection
    RecSet.LockType = adLockOptimistic
    
End Sub

Public Function ShowPageSetup(Grid As FlexCell.Grid)
    'Load frmPageSetup
    'frmPageSetup.SetGrid Grid
    'frmPageSetup.Show vbModal
End Function

Public Function CheckUsers(PW As String, Frm As Form) As AccessEnum
Dim rs5 As ADODB.Recordset
Set rs5 = New ADODB.Recordset
  ' Dim UsCode As String * 3
   Dim ProgCode As String * 3
   Dim TypAccess As String * 1
  ' MsgBox PW
   'Get Data From Banks For Access To PostDefAccess Bank
   If Not SuperV Then ' Other Users
'
      If Trim(VarFYear) <> Left(Trim(MainDate), 4) Or endyearhlp = True Then
         CheckUsers = ReadOnly
         Exit Function
      End If
'
      rs.Open "SELECT * From PostDefUsers WHERE UserPass='" + Trim(PW) + "'", cns, adOpenStatic
       
      If Not rs.EOF Then
         UsCode = Trim(rs("CodeUser"))
      End If
      
      rs.Close
      rs.Open "SELECT * From PostProgNames WHERE " _
             + "ProgramName='" + Trim(Frm.Name) + "'", cns, adOpenStatic
      If Not rs.EOF Then
         ProgCode = Trim(rs("ProgramCode"))
      Else
         MsgBox "Œÿ«..."
         rs.Close

         'Unload Me
         Exit Function
      End If
      
      rs.Close
      rs.Open "SELECT * From PostDefAccess WHERE FYear='" _
             + Trim(VarFYear) + "' And CodeOstan='" + Trim(VarCOstan) _
             + "' And CodeCity='" + Trim(SelectCityCode) _
             + "' And CodeOffice='" + Trim(SelectOfficeCode) _
             + "' And CodeUser='" + Trim(UsCode) _
             + "' And ProgramCode='" + Trim(ProgCode) + "'", cns, adOpenStatic
             
      If Not rs.EOF Then
         TypAccess = Trim(rs("TypeAccess"))
      End If
     rs.Close
      
      Select Case TypAccess
             Case "1": CheckUsers = ReadOnly
             Case "2": CheckUsers = ReadWrite
      End Select
      
   End If
End Function
