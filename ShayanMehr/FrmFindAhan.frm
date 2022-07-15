VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form FrmFindAhan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "B Zar"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmFindAhan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "«ÿ·«⁄«  «’·Ì »«— ‰«„Â"
      Height          =   2295
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1560
      Width           =   8655
      Begin VB.TextBox TxtShakhe 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   510
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox TxtKootajDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   510
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox TxtKeshti 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   510
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox TxtBArname 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   510
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox TxtParvane 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   510
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox TxtKootaj 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   510
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   8
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ ‘«ŒÂ"
         Height          =   390
         Left            =   3360
         TabIndex        =   24
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ òÊ «é"
         Height          =   390
         Left            =   7560
         TabIndex        =   22
         Top             =   1680
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ò‘ Ì"
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   3360
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " »«— ‰«„Â"
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   7560
         TabIndex        =   20
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â Å—Ê«‰Â"
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   7560
         TabIndex        =   19
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â òÊ «é"
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   3360
         TabIndex        =   18
         Top             =   1080
         Width           =   1035
      End
   End
   Begin VB.TextBox TxtFind 
      Height          =   510
      Left            =   3600
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.OptionButton OptKootaj 
      Alignment       =   1  'Right Justify
      Caption         =   "Ã” ÃÊ »« ‘„«—Â òÊ «é"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.OptionButton OptParvane 
      Alignment       =   1  'Right Justify
      Caption         =   "Ã” ÃÊ »« ‘„«—Â Å—Ê«‰Â"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin FlexCell.Grid Grid1 
      Height          =   3015
      Left            =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5318
      Appearance      =   0
      BackColorBkg    =   -2147483633
      BorderColor     =   -2147483633
      Cols            =   6
      DefaultFontName =   "Traditional Arabic"
      DefaultFontSize =   14.25
      DefaultFontBold =   -1  'True
      DefaultRowHeight=   32
      GridColor       =   0
      ReadOnly        =   -1  'True
      Rows            =   7
      SelectionMode   =   3
      EnterKeyMoveTo  =   1
   End
   Begin PrjShayan.TypeButton CmdOkFind 
      Default         =   -1  'True
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   " «ÌÌœ »—«Ì Ã” ÃÊ"
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
      MICON           =   "FrmFindAhan.frx":169B2
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdBack 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "»«“ê‘ "
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
      MICON           =   "FrmFindAhan.frx":169CE
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdContinue 
      Height          =   495
      Left            =   6720
      TabIndex        =   10
      Top             =   7320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "«œ«„Â Ê—Êœ «ÿ·«⁄« "
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
      MICON           =   "FrmFindAhan.frx":169EA
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdEditMaster 
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   7320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "«’·«Õ «ÿ·«⁄«  «’·Ì"
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
      MICON           =   "FrmFindAhan.frx":16A06
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdEditDetail 
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   7320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "«’·«Õ «ÿ·«⁄«  Ã«‰»Ì"
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
      MICON           =   "FrmFindAhan.frx":16A22
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdPrint 
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   7320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "Å—Ì‰  «“ ò· Å—Ê«‰Â"
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
      MICON           =   "FrmFindAhan.frx":16A3E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdPrintDate 
      Height          =   495
      Left            =   6480
      TabIndex        =   14
      Top             =   8040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "Å—Ì‰  Å—Ê«‰Â »Â  ›òÌò  «—ÌŒ"
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
      BCOL            =   33023
      BCOLO           =   12640511
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmFindAhan.frx":16A5A
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdDelParvane 
      Height          =   495
      Left            =   3480
      TabIndex        =   26
      Top             =   8040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      BTYPE           =   1
      TX              =   "Õ–› ò«„· Å—Ê«‰Â"
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
      BCOL            =   8421631
      BCOLO           =   255
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmFindAhan.frx":16A76
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LblWait 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "...·ÿ›« ç‰œ ·ÕŸÂ ’»— ò‰Ìœ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   7920
      Width           =   2745
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00400040&
      BorderWidth     =   3
      X1              =   240
      X2              =   8880
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label LblFind 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ã” ÃÊ »«"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   450
      Left            =   4800
      TabIndex        =   16
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "FrmFindAhan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ahan As Byte ' 1=AHAN   2= Kantiner    3=A.E.L

Private Sub Option1_Click()

End Sub

Private Sub CmdBack_Click()
  FrmStart.Show
  Unload Me
End Sub

Private Sub CmdContinue_Click()
 Dim Table As String
   
  On Error Resume Next
  
   If MsgBox("¬Ì« „ÿ„∆‰ Â” Ìœø", vbQuestion + vbYesNo) = vbNo Then Exit Sub
   
   If Ahan = 1 Then
      Table = "TabAhan_Detail"
   ElseIf Ahan = 2 Then
      Table = "TabKantiner_Detail"
   ElseIf Ahan = 3 Then
      Table = "TabAEL_Detail"
   End If
   '
   
   With FrmAhan_Detail
        .TxtBArname = TxtBArname
        .TxtKeshti = TxtKeshti
        .TxtParvane = TxtParvane
        .TxtKootaj = TxtKootaj
        .TxtSaheb = Grid1.Cell(4, 1).Text
        .TxtEtebar = Grid1.Cell(5, 1).Text
        .TxtGharardad = Grid1.Cell(6, 1).Text
        .TxtNWeight = Trim(Grid1.Cell(6, 4).Text)
        '
        rs.Open "SELECT COUNT(Parvane) FROM " & Table & " " & _
                "WHERE Parvane='" & Trim(TxtParvane) & "'", CNS
        If Ahan = 1 Or Ahan = 3 Then
        
           .LblRadif = rs(0) + 1
            rs.Close
           .FrameAhan.Visible = True
           .FrameKantiner.Visible = False
           .FrameAhan.Enabled = True
           .LblRadif.ForeColor = vbBlack
           .LblRadif1.ForeColor = vbBlack '
           '
           If Ahan = 3 Then
              .Caption = "«œ«„Â Ê—Êœ «ÿ·«⁄«    —ŒÌ’ Å«—”Â"
              .FrameAhan.Caption = "´ Ê—Êœ «ÿ·«⁄«  ¬Â‰ ¬·« ª   —ŒÌ’ Å«—”Â"
           Else
              .Caption = "«œ«„Â Ê—Êœ «ÿ·«⁄«  ¬Â‰ ¬·« "
           End If
           .Frame2.Enabled = False
           .CmdOk0.Visible = False
        ElseIf Ahan = 2 Then
           .kLblRadif = rs(0) + 1
            rs.Close
           .FrameAhan.Visible = False
           .FrameKantiner.Visible = True
           .FrameKantiner.Enabled = True
           .kLblRadif.ForeColor = vbBlack
           .kLblRadif1.ForeColor = vbBlack '
           '
           .Caption = "«œ«„Â Ê—Êœ «ÿ·«⁄«  ò«‰ Ì‰—"
           .Frame2.Enabled = False
           .CmdOk0.Visible = False
        End If
        
        Unload Me
        .Show
   End With
End Sub

Private Sub CmdDelParvane_Click()
 Dim msg As Integer
 Dim strMessage As String
 Dim Master$, Detail$, Tonaj$
 strMessage = "œ— ’Ê—  Õ–› ò· «ÿ·«⁄«  „—»Êÿ »Â «Ì‰ Å—Ê«‰Â Õ–›"
 strMessage = strMessage & vbNewLine & " ŒÊ«Âœ ‘œ Ê ﬁ«»· »—ê‘  ‰ŒÊ«Âœ »Êœ"
 strMessage = strMessage & vbNewLine & "„«Ì· »Â Õ–› Â” Ìœø"
 
 If TxtParvane = Empty Then
    MsgBox "Å—Ê«‰Â —« «‰ Œ«» ò‰Ìœ", vbExclamation, ""
    Exit Sub
 End If
 '
 msg = MsgBox(strMessage, vbCritical + vbYesNo, "")
 If msg = vbYes Then
 
   Select Case Ahan
      Case 1:
        Master = "TabAhan_Master"
        Detail = "TabAhan_Detail"
        Tonaj = "TabAhan_Tonaj"
      Case 2:
        Master = "TabKantiner_Master"
        Detail = "TabKantiner_Detail"
        Tonaj = "TabKantiner_Tonaj"
      Case 3:
        Master = "TabAEL_Master"
        Detail = "TabAEL_Detail"
        Tonaj = "TabAEL_Tonaj"
   End Select
   '
   Dim Myrs As New Recordset
   
   Myrs.Open "DELETE FROM " & Master & " " & _
             "WHERE Parvane='" & TxtParvane & "'", CNS
   '
   Myrs.Open "DELETE FROM " & Detail & " " & _
             "WHERE Parvane='" & TxtParvane & "'", CNS
   '
   Myrs.Open "DELETE FROM " & Tonaj & " " & _
             "WHERE Parvane='" & TxtParvane & "'", CNS
   '
   Set Myrs = Nothing
   MsgBox "Å—Ê«‰Â Õ–› ‘œ", vbInformation, ""
   Unload Me
 End If

End Sub

Private Sub CmdEditDetail_Click()
  If TxtParvane = Empty Then Exit Sub
  
  '
  LblWait = "...·ÿ›« ç‰œ ·ÕŸÂ ’»— ò‰Ìœ"
  
  Select Case Ahan
         Case 1 'AHAN
              FrmEditDetailA.Ahan = True
              FrmEditDetailA.ParvaneCode = TxtParvane
              FrmEditDetailA.Show 1
              '
         Case 2 'Kantiner
              FrmEditDetailKan.ParvaneCode = TxtParvane
              FrmEditDetailKan.Show 1
              '
         Case 3 ' AEL
              FrmEditDetailA.Ahan = False
              FrmEditDetailA.ParvaneCode = TxtParvane
              FrmEditDetailA.Show 1
  End Select
  
End Sub

Private Sub CmdEditMaster_Click()
 Dim Table, strSQL As String
On Error Resume Next
  If TxtParvane = Empty Then Exit Sub
  
  If Ahan = 1 Then
     Table = "TabAhan_Master"
  ElseIf Ahan = 2 Then
     Table = "TabKantiner_Master"
  ElseIf Ahan = 3 Then
     Table = "TabAEL_Master"
  End If
  ''
  
  If CmdEditMaster.Caption = "«’·«Õ «ÿ·«⁄«  «’·Ì" Then
     TxtBArname.Locked = False
     TxtKeshti.Locked = False
     TxtParvane.Locked = True
     TxtKootaj.Locked = False
     TxtKootajDate.Locked = False
     TxtShakhe.Locked = False
     '
     Grid1.ReadOnly = False
     Grid1.Column(3).Locked = True
     '
     CmdEditMaster.Caption = "À»   €ÌÌ—« "
     '
     TxtBArname.SetFocus
     '
     CmdOkFind.Enabled = False
     TxtFind.Enabled = False
     OptKootaj.Enabled = False
     OptParvane.Enabled = False
     '
     CmdEditDetail.Enabled = False
  ElseIf CmdEditMaster.Caption = "À»   €ÌÌ—« " Then
     
     strSQL = "UPDATE " & Table
     strSQL = strSQL & " SET Barname='" & Trim(TxtBArname) & "', "
     strSQL = strSQL & "Keshti='" & Trim(TxtKeshti) & "', "
     strSQL = strSQL & "TypeKala='" & Trim(Grid1.Cell(1, 4).Text) & "', "
     strSQL = strSQL & "Parvane='" & Trim(TxtParvane) & "', "
     strSQL = strSQL & "Kootaj='" & Trim(TxtKootaj) & "', "
     strSQL = strSQL & "DKootaj='" & Format(TxtKootajDate, "yy/mm/dd") & "', "
     strSQL = strSQL & "Ghabz='" & Trim(Grid1.Cell(2, 4).Text) & "', "
     strSQL = strSQL & "DGhabz='" & Format(Trim(Grid1.Cell(3, 4).Text), "yy/mm/dd") & "', "
     strSQL = strSQL & "Dparvane='" & Format(Trim(Grid1.Cell(4, 4).Text), "yy/mm/dd") & "', "
     strSQL = strSQL & "SizeKala='" & Trim(Grid1.Cell(5, 4).Text) & "', "
     strSQL = strSQL & "NWeight=" & Val(Trim(Grid1.Cell(6, 4).Text)) & ", "
     strSQL = strSQL & "Weight=" & Val(Trim(Grid1.Cell(1, 1).Text)) & ", "
     strSQL = strSQL & "Bandel=" & Val(Trim(Grid1.Cell(2, 1).Text)) & ", "
     strSQL = strSQL & "Shakhe=" & Val(Trim(TxtShakhe.Text)) & ", "
     strSQL = strSQL & "Tarkhiskar='" & Trim(Grid1.Cell(3, 1).Text) & "', "
     strSQL = strSQL & "Saheb='" & Trim(Grid1.Cell(4, 1).Text) & "', "
     strSQL = strSQL & "Etebar='" & Trim(Grid1.Cell(5, 1).Text) & "', "
     strSQL = strSQL & "Gharardad='" & Trim(Grid1.Cell(6, 1).Text) & "' "
     strSQL = strSQL & "WHERE Parvane='" & Trim(TxtParvane) & "'"
     ''''''''
     rs.Open strSQL, CNS
     ''''''''
     TxtBArname.Locked = True
     TxtKeshti.Locked = True
     TxtParvane.Locked = True
     TxtKootaj.Locked = True
     TxtKootajDate.Locked = True
     TxtShakhe.Locked = True
     '
     Grid1.ReadOnly = True
     'Grid1.Column(3).Locked = True
     '
     CmdEditMaster.Caption = "«’·«Õ «ÿ·«⁄«  «’·Ì"
     CmdEditDetail.Enabled = True
     '
     CmdOkFind.Enabled = True
     TxtFind.Enabled = True
     OptKootaj.Enabled = True
     OptParvane.Enabled = True
     '
     CmdOkFind.SetFocus
     '
     MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  À»  ‘œ", vbInformation
  End If
 ''
   
End Sub

Private Sub CmdOk_Click()

End Sub

Private Sub CmdOkFind_Click()
 On Error Resume Next
 
 Dim xField, Table As String
   If OptKootaj Then
      xField = "Kootaj"
   ElseIf OptParvane Then
      xField = "Parvane"
   End If
   '''
   If Ahan = 1 Then
      Table = "TabAhan_Master"
   ElseIf Ahan = 2 Then
      Table = "TabKantiner_Master"
   ElseIf Ahan = 3 Then
      Table = "TabAEL_Master"
   End If

   '
   rs.Open "SELECT * FROM " & Table & " " & _
           "WHERE " & xField & "='" & TxtFind & "'", CNS
   If Not rs.EOF Then
       TxtBArname = Trim(rs("Barname"))
       TxtKeshti = Trim(rs("Keshti"))
       TxtParvane = Trim(rs("Parvane"))
       TxtKootaj = Trim(rs("Kootaj"))
       TxtKootajDate = Trim(rs("DKootaj"))
       TxtShakhe = Trim(rs("Shakhe"))
       '''''
       With Grid1
           On Error Resume Next
            .Cell(1, 1).Text = rs("Weight")
            .Cell(2, 1).Text = rs("Bandel")
            .Cell(3, 1).Text = Trim(rs("Tarkhiskar"))
            .Cell(4, 1).Text = Trim(rs("Saheb"))
            .Cell(5, 1).Text = Trim(rs("Etebar"))
            .Cell(6, 1).Text = Trim(rs("Gharardad"))
      ''''''''''''''''''''''''''''''''''''''''''''''
            .Cell(1, 4).Text = Trim(rs("TypeKala"))
            .Cell(2, 4).Text = Trim(rs("Ghabz"))
            .Cell(3, 4).Text = Trim(rs("DGhabz"))
            .Cell(4, 4).Text = Trim(rs("DParvane"))
            .Cell(5, 4).Text = Trim(rs("SizeKala"))
            .Cell(6, 4).Text = rs("NWeight")
       End With
       '''
       CmdContinue.Enabled = True
       CmdEditDetail.Enabled = True
       CmdEditMaster.Enabled = True
       CmdPrint.Enabled = True
       CmdContinue.SetFocus
       ''''
   Else
       MsgBox "«ÿ·«⁄«  „Ê—œ ‰Ÿ— ÅÌœ« ‰‘œ", vbExclamation, ""
       TxtFind.SetFocus
       SendKeys "{home}+{end}"
   End If
   rs.Close
End Sub

Private Sub CmdPrint_Click()
  If Ahan = 1 Then
     FrmFinishAhan.fParvaneCode = TxtParvane.Text
     FrmFinishAhan.rAEL = False
     Unload Me
     FrmFinishAhan.Show
  ElseIf Ahan = 2 Then
     FrmFinishKantiner.fParvaneCode = TxtParvane.Text
     Unload Me
     FrmFinishKantiner.Show
  ElseIf Ahan = 3 Then
     FrmFinishAhan.fParvaneCode = TxtParvane.Text
     FrmFinishAhan.rAEL = True
     Unload Me
     FrmFinishAhan.Show
  End If
End Sub

Private Sub CmdPrintDate_Click()
  If Ahan = 1 Then
     FrmGetPrintDate.WichF = 1
     FrmGetPrintDate.ParCode = TxtParvane.Text
     FrmGetPrintDate.Caption = "Å—Ì‰  «“ ¬Â‰ ¬·«  »Â ‘„«—Â Å—Ê«‰Â  " & TxtParvane
  ElseIf Ahan = 2 Then
     FrmGetPrintDate.WichF = 2
     FrmGetPrintDate.ParCode = TxtParvane.Text
     FrmGetPrintDate.Caption = "Å—Ì‰  «“ ò«‰ Ì‰— »Â ‘„«—Â Å—Ê«‰Â  " & TxtParvane
  ElseIf Ahan = 3 Then
     FrmGetPrintDate.WichF = 3
     FrmGetPrintDate.ParCode = TxtParvane.Text
     FrmGetPrintDate.Caption = "Å—Ì‰  «“ AEL »Â ‘„«—Â Å—Ê«‰Â  " & TxtParvane
  End If
  
  FrmGetPrintDate.Show 1, FrmFindAhan
End Sub

Private Sub Form_Activate()
   If Caption = "Ã” ÃÊÌ «ÿ·«⁄«  ¬Â‰ ¬·« " Then
      Ahan = 1
   ElseIf Caption = "Ã” ÃÊÌ «ÿ·«⁄«  ò«‰ Ì‰—" Then
      Ahan = 2
   ElseIf Caption = "Ã” ÃÊÌ «ÿ·«⁄«    —ŒÌ’ Å«—”Â" Then
      Ahan = 3
   End If
   '
   LblWait = ""
   
End Sub

Private Sub Form_Load()
   Grid1.Column(0).Width = 0
   '
   Grid1.RowHeight(0) = 0
   '
   ''''
   Grid1.Range(1, 2, Grid1.Rows - 1, 2).BackColor = RGB(205, 194, 177)
   Grid1.Range(1, 5, Grid1.Rows - 1, 5).BackColor = RGB(205, 194, 177)

 '
   Grid1.Range(1, 1, Grid1.Rows - 1, 1).BackColor = RGB(106, 173, 90)
   Grid1.Range(1, 4, Grid1.Rows - 1, 4).BackColor = RGB(106, 173, 90)

   ''''
   With Grid1
        .Column(1).Width = 150
        .Column(2).Width = 110
        '
        .Cell(1, 2).Text = "Ê“‰ Œ«·’"
        .Cell(2, 2).Text = " ⁄œ«œ »‰œ·"
        .Cell(3, 2).Text = "‰«„  —ŒÌ’ ò«—"
        .Cell(4, 2).Text = "‰«„ ’«Õ» ò«·«"
        .Cell(5, 2).Text = "‘„«—Â «⁄ »«—"
        .Cell(6, 2).Text = "‘„«—Â ﬁ—«— œ«œ"
        '
        .Column(1).Alignment = cellRightCenter
        .Column(2).Alignment = cellRightCenter
        '.Column(2).AutoFit
        ''''''''''''''''''''''''''
        .Column(3).Width = 55
        .Range(1, 3, .Rows - 1, 3).BackColor = vbBlack
        .ReadOnly = False
        .Range(1, 3, .Rows - 1, 3).Merge
        .ReadOnly = True
        ''''''''''''''''''''''''''
        .Column(4).Width = 150
        .Column(5).Width = 110
        '
        .Cell(1, 5).Text = "‰Ê⁄ ò«·«"
        .Cell(2, 5).Text = "‘„«—Â ﬁ»÷ «‰»«—"
        .Cell(3, 5).Text = " «—ÌŒ ‘„«—Â ﬁ»÷"
        .Cell(4, 5).Text = " «—ÌŒ Å—Ê«‰Â"
        .Cell(5, 5).Text = "”«Ì“"
        .Cell(6, 5).Text = "Ê“‰ ‰« Œ«·’"
        '
        .Column(4).Alignment = cellRightCenter
        .Column(5).Alignment = cellRightCenter
        '.Column(2).AutoFit
   End With
''''''''''''''''''''''''
   CmdContinue.Enabled = False
   CmdEditMaster.Enabled = False
   CmdEditDetail.Enabled = False
   CmdPrint.Enabled = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  CmdBack_Click
End Sub

Private Sub OptKootaj_Click()
   LblFind = OptKootaj.Caption
   TxtFind.SetFocus
   TxtFind = Empty
End Sub

Private Sub OptParvane_Click()
   LblFind = OptParvane.Caption
   TxtFind.SetFocus
   TxtFind = Empty
End Sub

Private Sub TxtFind_GotFocus()
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard english
  If oldKB = 67699721 Then 'keyboard is farsi
     'do nothing
  Else
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

End Sub

Private Sub TxtFind_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
     If OptKootaj Then
         OptParvane.SetFocus
     Else
        OptKootaj.SetFocus
     End If
  End If
End Sub

Private Sub TxtKootajDate_LostFocus()
   TxtKootajDate = Format(TxtKootajDate, "yy/mm/dd")
End Sub

Private Sub TypeButton1_Click()

End Sub
