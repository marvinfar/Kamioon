VERSION 5.00
Begin VB.Form FrmGetPrintDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7755
   BeginProperty Font 
      Name            =   "B Zar"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmGetPrintDate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2265
   ScaleWidth      =   7755
   StartUpPosition =   1  'CenterOwner
   Begin PrjShayan.TypeButton CmdOk 
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "ÑíäÊ"
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
      MICON           =   "FrmGetPrintDate.frx":169B2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtDate2 
      Alignment       =   2  'Center
      Height          =   510
      Left            =   960
      MaxLength       =   8
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox TxtDate1 
      Alignment       =   2  'Center
      Height          =   510
      Left            =   3840
      MaxLength       =   8
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin PrjShayan.TypeButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "ÇäÕÑÇÝ"
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
      MICON           =   "FrmGetPrintDate.frx":169CE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdDateTakhlie 
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   1680
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "ÊÇííÏ"
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
      MICON           =   "FrmGetPrintDate.frx":169EA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Çáí"
      Height          =   390
      Left            =   3435
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÇÒ ÊÇÑíÎ "
      Height          =   390
      Left            =   6345
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   750
   End
End
Attribute VB_Name = "FrmGetPrintDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ParCode As String
Public WichF As Byte 'Ahan=1,Kantiner=2,AEL=3

Private Sub CmdCancel_Click()
  Me.BackColor = vbBlack
  Unload Me
End Sub

Private Sub CmdDateTakhlie_Click()
  With FrmTakhlieKantiner
       .strDate1 = TxtDate1
       .strDate2 = TxtDate2
       Unload Me
  End With
End Sub

Private Sub CmdOk_Click()
  If WichF = 1 Then
     FrmAhanRep.rAEL = False
     FrmAhanRep.ParvaneCode = ParCode
     FrmAhanRep.GetPrintDate = " DBarname>='" & TxtDate1 & "' AND DBarname<='" & TxtDate2 & "'"
     '
     Unload Me
     FrmAhanRep.Show
  ElseIf WichF = 3 Then
     FrmAhanRep.rAEL = True
     FrmAhanRep.ParvaneCode = ParCode
     FrmAhanRep.GetPrintDate = " DBarname>='" & TxtDate1 & "' AND DBarname<='" & TxtDate2 & "'"
     '
     Unload Me
     FrmAhanRep.Show
  ElseIf WichF = 2 Then
     FrmKantinerRep.ParvaneCode = ParCode
     FrmKantinerRep.GetPrintDate = " BarNameDate>='" & TxtDate1 & "' AND BarNameDate<='" & TxtDate2 & "'"
     '
     Unload Me
     FrmKantinerRep.Show
  ElseIf WichF = 4 Then
     FrmDefSize.ReportDate1 = TxtDate1
     FrmDefSize.ReportDate2 = TxtDate2
     '
     Unload Me
  End If
  '
End Sub

Private Sub TxtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub TxtDate1_LostFocus()
 TxtDate1 = Format(TxtDate1, "yy/mm/dd")
 TxtDate2 = TxtDate1
End Sub

Private Sub TxtDate2_GotFocus()
 SendKeys "{home}+{end}"
End Sub

Private Sub TxtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub TxtDate2_LostFocus()
 TxtDate2 = Format(TxtDate2, "yy/mm/dd")
End Sub
