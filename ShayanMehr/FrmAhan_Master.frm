VERSION 5.00
Object = "{9DBDC544-49CA-11D7-B1ED-C2237039C523}#1.1#0"; "FarDate.Ocx"
Begin VB.Form FrmAhan_Kantiner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10380
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
   Icon            =   "FrmAhan_Master.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " ﬁ«»·  ÊÃÂ ”—Ê— ê—«„Ì :"
      ForeColor       =   &H00000000&
      Height          =   5775
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   960
      Width           =   10095
      Begin VB.TextBox TxtShakhe 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   510
         Left            =   240
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   3360
         Width           =   2535
      End
      Begin VB.TextBox TxtGhabz 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   510
         Left            =   5040
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3360
         Width           =   2535
      End
      Begin VB.TextBox TxtTarkhiskar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   510
         Left            =   240
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   4080
         Width           =   2535
      End
      Begin VB.TextBox TxtBandel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   510
         Left            =   240
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox TxtKootaj 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   510
         Left            =   5040
         MaxLength       =   11
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox TxtWeight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   510
         Left            =   240
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox TxtParvane 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   510
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox TxtNWeight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   510
         Left            =   240
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox TxtTypeKala 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   510
         Left            =   5040
         TabIndex        =   2
         Text            =   "Tex1"
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox TxtKSize1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   480
         Width           =   2535
      End
      Begin FarDate1.FarDate TxtKootajDate 
         Height          =   495
         Left            =   5040
         TabIndex        =   5
         Top             =   2640
         Width           =   2535
         _ExtentX        =   4471
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
      Begin FarDate1.FarDate TxtGhabzDate 
         Height          =   495
         Left            =   5040
         TabIndex        =   7
         Top             =   4080
         Width           =   2535
         _ExtentX        =   4471
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
      Begin FarDate1.FarDate TxtParDate 
         Height          =   495
         Left            =   5040
         TabIndex        =   8
         Top             =   4800
         Width           =   2535
         _ExtentX        =   4471
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
      Begin VB.Label LblKan 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ ò«‰ Ì‰—"
         Height          =   390
         Left            =   3360
         TabIndex        =   33
         Top             =   2640
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ ò· ‘«ŒÂ"
         Height          =   390
         Left            =   3330
         TabIndex        =   32
         Top             =   3360
         Width           =   1320
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ ‘„«—Â ﬁ»÷"
         Height          =   390
         Left            =   8280
         TabIndex        =   31
         Top             =   4080
         Width           =   1500
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ﬁ»÷ «‰»«—"
         Height          =   390
         Left            =   8280
         TabIndex        =   30
         Top             =   3360
         Width           =   1365
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„  —ŒÌ’ ò«—"
         Height          =   390
         Left            =   3375
         TabIndex        =   29
         Top             =   4080
         Width           =   1275
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ Å—Ê«‰Â"
         Height          =   390
         Left            =   8280
         TabIndex        =   28
         Top             =   4800
         Width           =   1005
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ òÊ «é"
         Height          =   390
         Left            =   8280
         TabIndex        =   27
         Top             =   2640
         Width           =   990
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â òÊ «é"
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   8280
         TabIndex        =   26
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ê“‰ Œ«·’"
         Height          =   390
         Left            =   3360
         TabIndex        =   25
         Top             =   1920
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â Å—Ê«‰Â"
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   8280
         TabIndex        =   24
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ê“‰ ‰«Œ«·’"
         Height          =   390
         Left            =   3360
         TabIndex        =   23
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ ò«·«"
         Height          =   390
         Left            =   8280
         TabIndex        =   22
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "”«Ì“ ò«·«"
         Height          =   390
         Left            =   3480
         TabIndex        =   21
         Top             =   480
         Width           =   720
      End
      Begin VB.Label LblBan 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ »‰œ·°»” Â"
         Height          =   390
         Left            =   3360
         TabIndex        =   20
         Top             =   2640
         Width           =   1350
      End
   End
   Begin VB.TextBox TxtKeshti 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   510
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox TxtBArname 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   510
      Left            =   6840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   2535
   End
   Begin PrjShayan.TypeButton CmdOk 
      Height          =   495
      Left            =   8760
      TabIndex        =   15
      Top             =   6840
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
      COLTYPE         =   4
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAhan_Master.frx":169B2
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
      Left            =   6960
      TabIndex        =   16
      Top             =   6840
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
      COLTYPE         =   4
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAhan_Master.frx":169CE
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdDefSize 
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Top             =   6840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   " ⁄ÌÌ‰ „‘Œ’«  ”«Ì“ Â«"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648447
      BCOLO           =   32896
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAhan_Master.frx":169EA
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ò‘ Ì"
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   4155
      TabIndex        =   18
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " »«— ‰«„Â"
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   9540
      TabIndex        =   17
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "FrmAhan_Kantiner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ahan As Byte ' 1=AHAN   2= Kantiner    3=A.E.L or Soroosh

Private Sub CmdCancel_Click()
 Unload Me
 FrmStart.Show
End Sub

Private Sub CmdDefSize_Click()
 FrmDefSize.Ahan = Ahan
 FrmDefSize.Show
End Sub

Private Sub CmdOk_Click()
 Dim Table As String
' On Error Resume Next
   If Not CheckValidate Then Exit Sub ' If Information not Complete
   '
   Dim msg As Integer
    
    msg = MsgBox("«“ ’Õ  Ê—Êœ «ÿ·«⁄«  „ÿ„∆‰ Â” Ìœø", vbQuestion + vbYesNo, "")
    If msg = vbYes Then
       With FrmAhan_Detail
            If Ahan = 1 Then
               .FrameAhan.Visible = True: .FrameKantiner.Visible = False
               Table = "TabAhan_Master"
            ElseIf Ahan = 2 Then
               Table = "TabKantiner_Master"
               .FrameAhan.Visible = False: .FrameKantiner.Visible = True
            ElseIf Ahan = 3 Then
               FrmAhan_Detail.AEL = True
               .FrameAhan.Caption = "´ Ê—Êœ «ÿ·«⁄«  ¬Â‰ ¬·« ª   —ŒÌ’ Å«—”Â"
               .FrameAhan.Visible = True: .FrameKantiner.Visible = False
               Table = "TabAEL_Master"
            End If
            
            ''''Add To Tab Tonaj For Report
            If Ahan = 1 Then
               rs.Open "INSERT INTO TabAhan_Tonaj " & _
                       "(Parvane,TonajPar) " & _
                       "VALUES('" & Trim(TxtParvane) & "'," & Val(TxtNWeight) & ")", CNS
            ElseIf Ahan = 2 Then
               rs.Open "INSERT INTO TabKantiner_Tonaj " & _
                       "(Parvane,TonajPar) " & _
                       "VALUES('" & Trim(TxtParvane) & "'," & Val(TxtNWeight) & ")", CNS
            ElseIf Ahan = 3 Then
               rs.Open "INSERT INTO TabAEL_Tonaj " & _
                       "(Parvane,TonajPar) " & _
                       "VALUES('" & Trim(TxtParvane) & "'," & Val(TxtNWeight) & ")", CNS
            End If
            ''''''''''''''''''
            
            If Not SaveData(Table) Then Exit Sub
            .TxtBArname = Trim(TxtBArname)
            .TxtKeshti = Trim(TxtKeshti)
            .TxtParvane = Trim(TxtParvane)
            .TxtKootaj = Trim(TxtKootaj)
            '
       End With
       
       Unload Me
       FrmAhan_Detail.Show
    Else
       TxtBArname.SetFocus
    End If
End Sub

Private Sub Form_Activate()
   
   If Caption = "Ê—Êœ «ÿ·«⁄«  ¬Â‰ ¬·« " Then
      Ahan = 1
   ElseIf Caption = "Ê—Êœ «ÿ·«⁄«  ò«‰ Ì‰—" Then
      Ahan = 2
      LblKan.Visible = True
      LblBan.Visible = False
   ElseIf Caption = "Ê—Êœ «ÿ·«⁄«    —ŒÌ’ Å«—”Â" Then
      Ahan = 3
   End If
   
End Sub

Private Sub Form_Load()
   Call ClearField
   '
   BackColor = RGB(83, 132, 178)
   Frame1.BackColor = BackColor
End Sub
Sub ClearField()
   TxtBArname = Empty
   TxtKeshti = Empty
   TxtTypeKala = Empty
   TxtParvane = Empty
   TxtKootaj = Empty
   TxtKootajDate.Text = Empty
   TxtGhabz = Empty
   TxtGhabzDate.Text = Empty
   TxtParDate.Text = Empty
   TxtKSize1 = Empty
   TxtNWeight = 0
   TxtWeight = 0
   TxtBandel = 0
   TxtTarkhiskar = Empty
   TxtShakhe = 0
End Sub

Private Sub TxtBandel_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtBandel_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TxtShakhe.SetFocus
'
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtBarname_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TxtKeshti.SetFocus
End Sub

Private Sub TxtGhabz_Change()
'If Len(TxtGhabz) = 6 Then
'   TxtGhabz.SelStart = 6
'   TxtGhabz = "_" & TxtGhabz
' End If
End Sub

Private Sub TxtGhabz_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TxtGhabzDate.SetFocus
End Sub

Private Sub TxtGhabzDate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TxtParDate.SetFocus
End Sub

Private Sub TxtKeshti_GotFocus()
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard Engish
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

End Sub

Private Sub TxtKeshti_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TxtTypeKala.SetFocus
End Sub

Private Sub TxtKootaj_GotFocus()
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard english
  If oldKB = 67699721 Then 'keyboard is farsi
     'do nothing
  Else
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

End Sub

Private Sub TxtKootaj_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TxtKootajDate.SetFocus
End Sub

Private Sub TxtKootajDate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TxtGhabz.SetFocus
End Sub

Private Sub TxtKSize1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtNWeight_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtNWeight_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TxtWeight.SetFocus
'
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtParDate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TxtKSize1.SetFocus
End Sub

Private Sub TxtParvane_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TxtKootaj.SetFocus
End Sub

Private Sub TxtShakhe_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtShakhe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TxtTarkhiskar.SetFocus
End Sub

Private Sub TxtTarkhiskar_GotFocus()
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard Engish
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

End Sub

Private Sub TxtTarkhiskar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then CmdOk.SetFocus
End Sub

Private Sub TxtTypeKala_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TxtParvane.SetFocus
End Sub

Private Sub TxtWeight_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtWeight_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TxtBandel.SetFocus
'
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Function CheckValidate() As Boolean
   CheckValidate = True
   If Trim(TxtBArname) = Empty Then
      MsgBox "«ÿ·«⁄«  »«— ‰«„Â Œ«·Ì «” ", vbExclamation
      TxtBArname.SetFocus
      CheckValidate = False
      Exit Function
   End If
   '
   If Trim(TxtKeshti) = Empty Then
      MsgBox "«ÿ·«⁄«  ‰«„ ò‘ Ì Œ«·Ì «” ", vbExclamation
      TxtKeshti.SetFocus
      CheckValidate = False
      Exit Function
   End If
   '
   If Trim(TxtParvane) = Empty Then
      MsgBox "«ÿ·«⁄«  ‘„«—Â Å—Ê«‰Â Œ«·Ì «” ", vbExclamation
      TxtParvane.SetFocus
      CheckValidate = False
      Exit Function
   End If
   '
   If Trim(TxtKootaj) = Empty Then
      MsgBox "«ÿ·«⁄«  ‘„«—Â òÊ «é Œ«·Ì «” ", vbExclamation
      TxtKootaj.SetFocus
      CheckValidate = False
      Exit Function
   End If

End Function

Function SaveData(ByVal Table As String) As Boolean
'On Error Resume Next
   Dim strSQL As String
   Dim strKSize As String
         '
   SaveData = True
   strKSize = TxtKSize1
''''''''''''''
   strSQL = "INSERT INTO  " & Table
   strSQL = strSQL & "(Barname,Keshti,TypeKala,Parvane,Kootaj,DKootaj,"
   strSQL = strSQL & "Ghabz,DGhabz,Dparvane,SizeKala,NWeight,Weight,Bandel,Shakhe,"
   strSQL = strSQL & "Tarkhiskar) "
   '
   strSQL = strSQL & "VALUES('" & Trim(TxtBArname) & "',"
   strSQL = strSQL & "'" & Trim(TxtKeshti) & "',"
   strSQL = strSQL & "'" & Trim(TxtTypeKala) & "',"
   strSQL = strSQL & "'" & Trim(TxtParvane) & "',"
   strSQL = strSQL & "'" & Trim(TxtKootaj) & "',"
   strSQL = strSQL & "'" & Mid(TxtKootajDate.Text, 3) & "',"
   strSQL = strSQL & "'" & Trim(TxtGhabz) & "',"
   strSQL = strSQL & "'" & Mid(TxtGhabzDate.Text, 3) & "',"
   strSQL = strSQL & "'" & Mid(TxtParDate.Text, 3) & "',"
   strSQL = strSQL & "'" & Trim(strKSize) & "',"
   strSQL = strSQL & Val(TxtNWeight) & ","
   strSQL = strSQL & Val(TxtWeight) & ","
   strSQL = strSQL & Val(TxtBandel) & ","
   strSQL = strSQL & Val(TxtShakhe) & ","
   strSQL = strSQL & "'" & Trim(TxtTarkhiskar) & "')"
   '
   On Error GoTo L:
   rs.Open strSQL, CNS
   

L:    If InStr(1, Err.Description, "duplicate") > 0 Then
         MsgBox "‘„«—Â Å—Ê«‰Â Ì« ‘„«—Â òÊ «é  ò—«—Ì «” ", vbCritical, ""
         SaveData = False
         TxtParvane.SetFocus
         Exit Function
      End If
End Function
