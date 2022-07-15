VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmSaderatDetail 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ê—Êœ «ÿ·«⁄«  Ã«‰»Ì ò«‰ Ì‰— ’«œ—« Ì"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11235
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "B Zar"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDetail 
      BackColor       =   &H00C0C0FF&
      Caption         =   "«ÿ·«⁄«  «’·Ì ò«‰ Ì‰— ’«œ—« Ì"
      Height          =   4455
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1200
      Width           =   11055
      Begin VB.TextBox TxtBarnameDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   0
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox TxtBarnameNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox TxtKantiner 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   7080
         MaxLength       =   12
         TabIndex        =   4
         Top             =   3360
         Width           =   2415
      End
      Begin VB.TextBox TxtKamioon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   8280
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox TxtSerial 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   7080
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox TxtMobile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   480
         MaxLength       =   11
         TabIndex        =   6
         Top             =   1680
         Width           =   2535
      End
      Begin VB.ComboBox Combsize 
         Height          =   510
         ItemData        =   "FrmSaderatDetail.frx":0000
         Left            =   480
         List            =   "FrmSaderatDetail.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2520
         Width           =   2535
      End
      Begin PrjShayan.TypeButton CmdOK 
         Height          =   495
         Left            =   2760
         TabIndex        =   8
         Top             =   3840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   " «ÌÌœ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Zar"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmSaderatDetail.frx":0004
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrjShayan.TypeButton CmdCancel 
         Height          =   495
         Left            =   480
         TabIndex        =   28
         Top             =   3840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "«‰’—«›"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Zar"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmSaderatDetail.frx":0020
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox TxtTotal 
         Height          =   510
         Left            =   480
         TabIndex        =   5
         Top             =   780
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   900
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin PrjShayan.TypeButton CmdSave 
         Height          =   495
         Left            =   2760
         TabIndex        =   29
         Top             =   3840
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "À»   €ÌÌ—« "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Zar"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmSaderatDetail.frx":003C
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox TxtEditKamioon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   7080
         TabIndex        =   30
         Top             =   2520
         Visible         =   0   'False
         Width           =   2415
      End
      Begin PrjShayan.TypeButton CmdEditCancel 
         Height          =   495
         Left            =   480
         TabIndex        =   31
         Top             =   3840
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "«‰’—«›"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Zar"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmSaderatDetail.frx":0058
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ »«—‰«„Â "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   9720
         TabIndex        =   27
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â »«—‰«„Â "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   9720
         TabIndex        =   26
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ò«‰ Ì‰— "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   9720
         TabIndex        =   25
         Top             =   3360
         Width           =   1140
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ò«„ÌÊ‰ "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   9720
         TabIndex        =   24
         Top             =   2520
         Width           =   1245
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â „Ê»«Ì·"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   3360
         TabIndex        =   23
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ò· ò—«ÌÂ  »Â —Ì«·"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   3360
         TabIndex        =   22
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "”«Ì“ ò«‰ Ì‰—"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   3360
         TabIndex        =   21
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label LblRadif 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   5700
         TabIndex        =   20
         Top             =   0
         Width           =   180
      End
      Begin VB.Label LblRadif1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "‘„«—Â —œÌ› "
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   6180
         TabIndex        =   19
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Frame FrameMaster 
      BackColor       =   &H00C0C0FF&
      Caption         =   "«ÿ·«⁄«  «’·Ì ò«‰ Ì‰— ’«œ—« Ì"
      Height          =   1095
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   11055
      Begin VB.TextBox TxtPart 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800080&
         Height          =   510
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtRadifMarzi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800080&
         Height          =   510
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox TxtKootaj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800080&
         Height          =   510
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "12345678"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox TxtTransitNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800080&
         Height          =   510
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "125547899"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Å«— "
         ForeColor       =   &H00800080&
         Height          =   390
         Left            =   1560
         TabIndex        =   17
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â  —«‰“Ì "
         ForeColor       =   &H00800080&
         Height          =   495
         Left            =   9600
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â òÊ «é"
         ForeColor       =   &H00800080&
         Height          =   495
         Left            =   6600
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—œÌ› „—“Ì"
         ForeColor       =   &H00800080&
         Height          =   495
         Left            =   3720
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmSaderatDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_Click()
 Dim msg As Integer
    
   msg = MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ Œ«—Ã ‘ÊÌœø", vbQuestion + vbYesNo, "")
   If msg = vbNo Then
      TxtBarnameDate.SetFocus
   Else
      Unload Me
      FrmStart.Show
   End If

End Sub

Private Sub CmdEditCancel_Click()
  Unload Me
End Sub

Private Sub CmdOk_Click()
  If Not CheckValidate Then Exit Sub
  
  Dim strSQL As String
  Dim strKamioon As String
  Dim rsT As New Recordset
  
  strKamioon = TxtKamioon & TxtSerial
'''
  With rsT
       .Open "SELECT * FROM TabSaderat_Detail", CNS, adOpenStatic, adLockOptimistic
       .AddNew
       
       .Fields("TransitNo") = TxtTransitNo
       .Fields("Count0") = Val(LblRadif)
       .Fields("BarnameDate") = TxtBarnameDate
       .Fields("BarnameNo") = TxtBarnameNo
       .Fields("Kamioon") = strKamioon
       .Fields("Kantiner") = Trim(TxtKantiner)
       .Fields("Takhlie") = 0
       .Fields("Size") = Val(Combsize.Text)
       .Fields("Total") = CCur(TxtTotal)
       .Fields("Mobile") = TxtMobile
       
       .Update
         
       .Close
  End With
  Set rsT = Nothing
  
  '
  Call ClearField
  LblRadif = Val(LblRadif) + 1
  If Val(LblRadif) > Val(TxtPart) Then
     MsgBox " ⁄œ«œ —œÌ› «“  ⁄œ«œ Å«—  »Ì‘ — «” " & vbCrLf & _
            "‘„««Ã«“Â «œ«„Â Ê—Êœ «ÿ·«⁄«  —« ‰œ«—Ìœ", vbInformation
     
     Unload Me
  End If
  TxtBarnameDate.SetFocus
End Sub

Private Sub CmdSave_Click()
  If Not EDITCheckValidate Then Exit Sub
  
  Dim rsT As New Recordset
  Dim strSQL As String
    
  strSQL = "SELECT * FROM TabSaderat_Detail "
  strSQL = strSQL & "WHERE TransitNo='" & TxtTransitNo & "' AND "
  strSQL = strSQL & "Count0=" & Val(LblRadif)
  '
  rsT.Open strSQL, CNS, adOpenStatic, adLockOptimistic
  rsT("BarnameDate") = TxtBarnameDate
  rsT("BarnameNo") = TxtBarnameNo
  rsT("Kamioon") = TxtEditKamioon
  rsT("Kantiner") = TxtKantiner
  rsT("Size") = Val(Combsize.Text)
  rsT("Total") = CCur(TxtTotal)
  rsT("Mobile") = TxtMobile
  '
  rsT.Update
  '
  rsT.Close
  Set rsT = Nothing
  MsgBox " €ÌÌ—«  «⁄„«· ‘œÂ À»  ‘œ", vbInformation
  '
  With FrmTakhlieKantiner
       .Grid1.Cell(.Grid1.ActiveCell.Row, 8).Text = TxtBarnameDate
       .Grid1.Cell(.Grid1.ActiveCell.Row, 7).Text = TxtBarnameNo
       .Grid1.Cell(.Grid1.ActiveCell.Row, 6).Text = TxtEditKamioon
       .Grid1.Cell(.Grid1.ActiveCell.Row, 5).Text = TxtKantiner
       .Grid1.Cell(.Grid1.ActiveCell.Row, 3).Text = Combsize.Text
       .Grid1.Cell(.Grid1.ActiveCell.Row, 2).Text = Format(TxtTotal, "#,##0")
       .Grid1.Cell(.Grid1.ActiveCell.Row, 1).Text = TxtMobile
       
       Unload Me
  End With
  
End Sub

Private Sub Combsize_Click()
  If Combsize.ListIndex = 2 Then Combsize.ListIndex = 0
End Sub

Private Sub Combsize_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
   Me.BackColor = RGB(250, 160, 160)
   FrameMaster.BackColor = Me.BackColor
   FrameDetail.BackColor = Me.BackColor
   
   LblRadif.BackColor = RGB(250, 100, 150)
   LblRadif1.BackColor = RGB(250, 100, 150)
   '
   With Combsize
        .AddItem "20ft"
        .AddItem "40ft"
        .AddItem String(20, "-")
        .AddItem "30ft"
        .AddItem "35ft"
        
   End With

   ClearField

End Sub

Private Sub TxtBarnameDate_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtBarNameDate_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"

End Sub

Private Sub TxtBarnameDate_LostFocus()
   TxtBarnameDate = Format(TxtBarnameDate, "yy/mm/dd")
End Sub

Private Sub TxtBarnameNo_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtBarnameNo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"

 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Private Sub TxtKamioon_Change()
  If Len(TxtKamioon) = 3 Then
     TxtKamioon = TxtKamioon & "⁄"
     SendKeys "{End}"
  End If
  '
  If Len(TxtKamioon) = 6 Then TxtSerial.SetFocus
End Sub

Private Sub TxtKamioon_GotFocus()
  SendKeys "{Home}+{End}"
 
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard Engish
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If
  
End Sub

Private Sub TxtKamioon_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"

End Sub

Private Sub TxtKantiner_Change()
  If Len(TxtKantiner) = 10 Then
     TxtKantiner = TxtKantiner & "-"
     SendKeys "{End}"
  End If
End Sub

Private Sub TxtKantiner_GotFocus()
  SendKeys "{End}"
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard Engish
  If oldKB = 67699721 Then 'keyboard is English
  Else
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If
End Sub

Private Sub TxtKantiner_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
  If KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub TxtMobile_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtMobile_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"

 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Private Sub TxtSerial_GotFocus()
  SendKeys "{End}"
End Sub

Private Sub TxtSerial_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub TxtTotal_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtTotal_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"

 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Sub ClearField()

   'TxtBarnameDate = Empty
   TxtBarnameNo = IIf(TxtBarnameNo = Empty, Empty, Val(TxtBarnameNo) + 1)
   TxtKamioon = Empty
   TxtSerial = "«Ì—«‰"
   TxtKantiner = "irsu"
   Combsize.ListIndex = 0
   'TxtTotal = Empty
   TxtMobile = Empty
End Sub


Function CheckValidate() As Boolean
   CheckValidate = True
   If Len(Trim(TxtBarnameDate)) <> 8 Then
      MsgBox "  «—ÌŒ »«—‰«„Â «‘ »«Â «” ", vbExclamation
      TxtBarnameDate.SetFocus
      CheckValidate = False
      Exit Function
   End If
   '
   If Trim(TxtBarnameNo) = Empty Then
      MsgBox "‘„«—Â »«—‰«„Â Œ«·Ì «” ", vbExclamation
      TxtBarnameNo.SetFocus
      CheckValidate = False
      Exit Function
   End If
   '
   If TxtKamioon & TxtSerial = Empty Then
      MsgBox "«ÿ·«⁄«  ‘„«—Â ò«„ÌÊ‰ Œ«·Ì «” ", vbExclamation
      TxtKamioon.SetFocus
      CheckValidate = False
      Exit Function
   End If
   '
   If Len(Trim(TxtKantiner)) < 7 Then
      MsgBox "‘„«—Â ò«‰ Ì‰— «‘ »«Â «” ", vbExclamation
      TxtKantiner.SetFocus
      CheckValidate = False
      Exit Function
   End If
   '
   If Trim(TxtTotal) = Empty Then
      MsgBox "«ÿ·«⁄«  ò—«ÌÂ Œ«·Ì «” ", vbExclamation
      TxtTotal.SetFocus
      CheckValidate = False
      Exit Function
   End If
   
End Function


Function EDITCheckValidate() As Boolean
   EDITCheckValidate = True
   If Len(Trim(TxtBarnameDate)) <> 8 Then
      MsgBox "  «—ÌŒ »«—‰«„Â «‘ »«Â «” ", vbExclamation
      TxtBarnameDate.SetFocus
      EDITCheckValidate = False
      Exit Function
   End If
   '
   If Trim(TxtBarnameNo) = Empty Then
      MsgBox "‘„«—Â »«—‰«„Â Œ«·Ì «” ", vbExclamation
      TxtBarnameNo.SetFocus
      EDITCheckValidate = False
      Exit Function
   End If
   '
   If TxtEditKamioon = Empty Then
      MsgBox "«ÿ·«⁄«  ‘„«—Â ò«„ÌÊ‰ Œ«·Ì «” ", vbExclamation
      TxtEditKamioon.SetFocus
      EDITCheckValidate = False
      Exit Function
   End If
   '
   If Len(Trim(TxtKantiner)) < 7 Then
      MsgBox "‘„«—Â ò«‰ Ì‰— «‘ »«Â «” ", vbExclamation
      TxtKantiner.SetFocus
      EDITCheckValidate = False
      Exit Function
   End If
   '
   If Trim(TxtTotal) = Empty Then
      MsgBox "«ÿ·«⁄«  ò—«ÌÂ Œ«·Ì «” ", vbExclamation
      TxtTotal.SetFocus
      EDITCheckValidate = False
      Exit Function
   End If
   
End Function

