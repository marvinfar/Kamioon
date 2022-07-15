VERSION 5.00
Object = "{9DBDC544-49CA-11D7-B1ED-C2237039C523}#1.1#0"; "FarDate.Ocx"
Begin VB.Form FrmSaderatMaster 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10485
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtSaheb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   240
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox TxtTarkhiskar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   240
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2760
      Width           =   3135
   End
   Begin VB.ComboBox CombPackage 
      Height          =   510
      ItemData        =   "FrmSaderatMaster.frx":0000
      Left            =   240
      List            =   "FrmSaderatMaster.frx":000D
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox TxtPart 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   240
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox TxtTypeProduct 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   240
      Width           =   3135
   End
   Begin VB.TextBox TxtBarnameDarya 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   5640
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3593
      Width           =   3135
   End
   Begin VB.TextBox TxtRadifMarzi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800080&
      Height          =   510
      Left            =   5640
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2751
      Width           =   3135
   End
   Begin VB.TextBox TxtKootaj 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800080&
      Height          =   510
      Left            =   5640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1082
      Width           =   3135
   End
   Begin VB.TextBox TxtTransitNo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800080&
      Height          =   510
      Left            =   5640
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   3135
   End
   Begin FarDate1.FarDate TxtTransitDate 
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   1920
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Zar"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PrjShayan.TypeButton CmdOk 
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   4800
      Width           =   1455
      _ExtentX        =   2566
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSaderatMaster.frx":0025
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
      Cancel          =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   4800
      Width           =   1455
      _ExtentX        =   2566
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSaderatMaster.frx":0041
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "’«Õ» ò«·«"
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   3600
      TabIndex        =   21
      Top             =   3600
      Width           =   960
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„  —ŒÌ’ ò«—"
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   3600
      TabIndex        =   20
      Top             =   2760
      Width           =   1275
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ »” Â »‰œÌ"
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   3600
      TabIndex        =   19
      Top             =   1920
      Width           =   1230
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‹‹«— "
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3600
      TabIndex        =   18
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ ò‹‹«·«"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3600
      TabIndex        =   17
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "»«—‰«„Â œ—Ì«ÌÌ"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   9000
      TabIndex        =   16
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â  —«‰“Ì "
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   9000
      TabIndex        =   15
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â òÊ «é"
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   9000
      TabIndex        =   14
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ  —«‰“Ì‹ "
      Height          =   495
      Left            =   9000
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "—œÌ› „—“Ì"
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   9000
      TabIndex        =   12
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "FrmSaderatMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_Click()
 Dim msg As Integer
    
   msg = MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ Œ«—Ã ‘ÊÌœø", vbQuestion + vbYesNo, "")
   If msg = vbNo Then
      TxtTransitNo.SetFocus
   Else
      Unload Me
      FrmStart.Show
   End If

End Sub

Private Sub CmdOk_Click()
  If Not CheckValidate Then Exit Sub
  '
  Dim rsT As New Recordset
  Dim strSQL As String
  
  strSQL = "INSERT INTO TabSaderat_Master "
  strSQL = strSQL & "(TransitNo,KootajNo,TransitDate,RadifMarzi,"
  strSQL = strSQL & "BarnameDarya,TypeProduct,Part,"
  strSQL = strSQL & "TypePackage,Tarkhiskar,Saheb) "
  strSQL = strSQL & "VALUES('" & TxtTransitNo & "','"
  strSQL = strSQL & TxtKootaj & "','" & Mid(TxtTransitDate.Text, 3) & "',"
  strSQL = strSQL & Val(TxtRadifMarzi) & ",'" & TxtBarnameDarya & "','"
  strSQL = strSQL & TxtTypeProduct & "'," & Val(TxtPart) & ",'"
  strSQL = strSQL & CombPackage.Text & "','"
  strSQL = strSQL & TxtTarkhiskar & "','" & TxtSaheb & "')"
  '
  rsT.Open strSQL, CNS
  Set rsT = Nothing
  '
  MsgBox "«ÿ·«⁄«  À»  ‘œ", vbInformation, "MehrVarzan"
  '
  FrmSaderatDetail.TxtTransitNo = TxtTransitNo
  FrmSaderatDetail.TxtKootaj = TxtKootaj
  FrmSaderatDetail.TxtRadifMarzi = TxtRadifMarzi
  FrmSaderatDetail.TxtPart = TxtPart
  FrmSaderatDetail.LblRadif = 1
  '
  Unload Me
  FrmSaderatDetail.Show
End Sub

Private Sub CombPackage_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"

End Sub

Private Sub Form_Load()
   Me.BackColor = RGB(250, 160, 160)
   '
   ClearField
End Sub

Sub ClearField()
   TxtTransitNo = Empty
   TxtKootaj = Empty
   TxtTransitDate.Text = Empty
   TxtRadifMarzi = Empty
   TxtBarnameDarya = Empty
   TxtTypeProduct = Empty
   TxtPart = Empty
   CombPackage.ListIndex = 0
   TxtTarkhiskar = "‘«Ì«‰ „Â—"
   TxtSaheb = "ò‘ Ì—«‰Ì Ã‰Ê»-Œÿ «Ì—«‰"
End Sub

Private Sub TxtBarnameDarya_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtBarnameDarya_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"
 
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtKootaj_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtKootaj_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"
 
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtPart_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtPart_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"
 
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtRadifMarzi_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtRadifMarzi_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"

 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtSaheb_GotFocus()
  SendKeys "{Home}+{End}"
 
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard Engish
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

End Sub

Private Sub TxtSaheb_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"

End Sub

Private Sub TxtTarkhiskar_GotFocus()
  SendKeys "{Home}+{End}"
 
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard Engish
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

End Sub

Private Sub TxtTarkhiskar_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"

End Sub

Private Sub TxtTransitNo_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtTransitNo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"

 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub


Private Sub TxtTypeProduct_GotFocus()
  SendKeys "{Home}+{End}"
 
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard Engish
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

  
End Sub

Private Sub TxtTypeProduct_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"

End Sub

Function CheckValidate() As Boolean
   CheckValidate = True
   If Trim(TxtTransitNo) = Empty Then
      MsgBox "«ÿ·«⁄«  ‘„«—Â  —«‰“Ì  Œ«·Ì «” ", vbExclamation
      TxtTransitNo.SetFocus
      CheckValidate = False
      Exit Function
   End If
   '
   If Trim(TxtKootaj) = Empty Then
      MsgBox "«ÿ·«⁄«  òÊ ‹«é Œ«·Ì «” ", vbExclamation
      TxtKootaj.SetFocus
      CheckValidate = False
      Exit Function
   End If
   '
   If TxtTransitDate.Text = Empty Then
      MsgBox "«ÿ·«⁄«   «—ÌŒ  —«‰“Ì  Œ«·Ì «” ", vbExclamation
      TxtTransitDate.SetFocus
      CheckValidate = False
      Exit Function
   End If
   '
   If Trim(TxtRadifMarzi) = Empty Then
      MsgBox "«ÿ·«⁄«  —œÌ› „—“Ì Œ«·Ì «” ", vbExclamation
      TxtRadifMarzi.SetFocus
      CheckValidate = False
      Exit Function
   End If
   '
   If Trim(TxtPart) = Empty Then
      MsgBox "«ÿ·«⁄«  Å«—  Œ«·Ì «” ", vbExclamation
      TxtPart.SetFocus
      CheckValidate = False
      Exit Function
   End If
   '
   Dim rsT As New Recordset
   rsT.Open "SELECT RadifMarzi FROM TabSaderat_Master " & _
            "WHERE RadifMarzi=" & TxtRadifMarzi, CNS
   If Not rsT.EOF Then
      MsgBox " ‘„«—Â —œÌ› „—“Ì  ò—«—Ì «” ", vbExclamation
      TxtRadifMarzi.SetFocus
      CheckValidate = False
      rsT.Close
      Set rsT = Nothing
      Exit Function
   Else
      rsT.Close
      Set rsT = Nothing
   End If
   
End Function

