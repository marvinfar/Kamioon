VERSION 5.00
Begin VB.Form FrmTools 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ‰ŸÌ„« "
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "B Zar"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTools.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4965
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramePass 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   7455
      Begin VB.TextBox TxtPass2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   510
         IMEMode         =   3  'DISABLE
         Left            =   3360
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox TxtPass1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   510
         IMEMode         =   3  'DISABLE
         Left            =   3360
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin PrjShayan.TypeButton CmdOkPass 
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
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
         MICON           =   "FrmTools.frx":29C12
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrjShayan.TypeButton CmdCancelPass 
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
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
         MICON           =   "FrmTools.frx":29C2E
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrjShayan.TypeButton CmdRemove 
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "»—œ«‘ ‰ ò·„Â ⁄»Ê—"
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
         BCOL            =   12640511
         BCOLO           =   16448
         FCOL            =   255
         FCOLO           =   4210816
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmTools.frx":29C4A
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
         Caption         =   "œ— Â‰ê«„ Ê—Êœ ò·„Â ⁄»Ê— »Â —«”  «“ Õ—Ê› ·« Ì‰ «” ›«œÂ ‰„«ÌÌœ"
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   1740
         TabIndex        =   10
         Top             =   2640
         Width           =   5475
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «ÌÌœ ò·„Â ⁄»Ê— "
         Height          =   390
         Left            =   5940
         TabIndex        =   4
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ò·„Â ⁄»Ê—"
         Height          =   390
         Left            =   6450
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option1"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "FrmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOkPass_Click()
  If TxtPass1 = Empty Or TxtPass2 = Empty Then
     MsgBox "—„“ ⁄»Ê— Ê«—œ ‰‘œÂ «” ", vbCritical, ""
     TxtPass1.SetFocus
     Exit Sub
  End If
  '
  If StrComp(TxtPass1, TxtPass2, vbBinaryCompare) = 0 Then 'Equal
     SaveSetting "HKEY_CURRENT_USER", "xMehrvarzan", "PASSWORD", TxtPass1
     MsgBox "—„“ ⁄»Ê— À»  ‘œ", vbInformation, ""
     Unload Me
     Exit Sub
  Else
     MsgBox "Â— œÊ —„“ Ìò”«‰ Ê«—œ ‘Êœ", vbCritical, ""
     TxtPass1.SetFocus
     Exit Sub
  End If

End Sub

Private Sub Form_Load()
   RightToLeft = True
   
   BackColor = RGB(83, 132, 178)
   FramePass.BackColor = BackColor
   
End Sub

Private Sub CmdRemove_Click()
  Call DeleteSetting("HKEY_CURRENT_USER", "xMehrvarzan", "PASSWORD")
  MsgBox "—„“ ⁄»Ê— »—œ«‘ Â ‘œ", vbInformation, ""
  Unload Me
End Sub

Private Sub TypeButton2_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  FrmStart.Show
End Sub

Private Sub TxtPass1_GotFocus()
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard english
  If oldKB = 67699721 Then 'keyboard is farsi
     'do nothing
  Else
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

End Sub

Private Sub TxtPass2_GotFocus()
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard english
  If oldKB = 67699721 Then 'keyboard is farsi
     'do nothing
  Else
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

End Sub
