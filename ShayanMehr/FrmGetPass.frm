VERSION 5.00
Begin VB.Form FrmGetPass 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
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
   Picture         =   "FrmGetPass.frx":0000
   ScaleHeight     =   2970
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPass 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   510
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1440
      Width           =   3015
   End
   Begin PrjShayan.TypeButton CmdOk 
      Default         =   -1  'True
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   2280
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
      MICON           =   "FrmGetPass.frx":42B7
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
      Left            =   1920
      TabIndex        =   3
      Top             =   2280
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
      MICON           =   "FrmGetPass.frx":42D3
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LblPass 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "·ÿ›« ò·„Â ⁄»Ê— —« Ê«—œ ‰„«ÌÌœ "
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   3900
      TabIndex        =   2
      Top             =   960
      Width           =   2520
   End
End
Attribute VB_Name = "FrmGetPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOk_Click()
  If TxtPass.Text = LblPass Then
     Unload Me
     FrmStart.Show
  Else
     MsgBox "‘„« ò«—»— „Ã«“ ‰Ì” Ìœ", vbCritical, ""
     TxtPass.SetFocus
     SendKeys "{Home}+{End}"
  End If

End Sub

Private Sub Form_Load()
   RightToLeft = True
  
End Sub

Private Sub Label1_Click()

End Sub

Private Sub CmdCancel_Click()
   End
End Sub

Private Sub TxtPass_GotFocus()
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard english
  If oldKB = 67699721 Then 'keyboard is farsi
     'do nothing
  Else
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

End Sub
