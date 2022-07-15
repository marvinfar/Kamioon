VERSION 5.00
Object = "{9DBDC544-49CA-11D7-B1ED-C2237039C523}#1.1#0"; "FarDate.Ocx"
Begin VB.Form FrmAhan_Detail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘—ò  Õ„· Ê ‰ﬁ· „Â—Ê—“«‰  —«»— Œ“—"
   ClientHeight    =   8970
   ClientLeft      =   -375
   ClientTop       =   540
   ClientWidth     =   12735
   BeginProperty Font 
      Name            =   "B Zar"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAhan_Detail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   12735
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNWeight 
      Alignment       =   1  'Right Justify
      Height          =   510
      Left            =   360
      TabIndex        =   67
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "´„‘Œ’«  »«— ‰«„Â Ê ’«Õ» ò«·«ª"
      Height          =   1575
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   1680
      Width           =   12495
      Begin VB.TextBox TxtEtebar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox TxtSaheb 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   8040
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox TxtGharardad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin PrjShayan.TypeButton CmdOk0 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         MICON           =   "FrmAhan_Detail.frx":169B2
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â «⁄ »«— "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   6780
         TabIndex        =   44
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ ’«Õ» ò«·«"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   10725
         TabIndex        =   43
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ﬁ—«— œ«œ "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   2700
         TabIndex        =   42
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " «ÿ·«⁄«  „— »Êÿ »Â :"
      Height          =   1335
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   120
      Width           =   12495
      Begin VB.TextBox TxtKootaj 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00404040&
         Height          =   510
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox TxtParvane 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00404040&
         Height          =   510
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox TxtBArname 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00404040&
         Height          =   510
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TxtKeshti 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00404040&
         Height          =   510
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â òÊ «é "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   2460
         TabIndex        =   40
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â Å—Ê«‰Â "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   5700
         TabIndex        =   39
         Top             =   480
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " »«— ‰«„Â"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   11460
         TabIndex        =   38
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ò‘ Ì "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   8820
         TabIndex        =   37
         Top             =   480
         Width           =   555
      End
   End
   Begin PrjShayan.TypeButton CmdFinishDay 
      Height          =   495
      Left            =   9240
      TabIndex        =   30
      Top             =   8280
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "« „«„ Ê—Êœ «ÿ·«⁄«  —Ê“«‰Â"
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
      MICON           =   "FrmAhan_Detail.frx":169CE
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdFinish 
      Height          =   495
      Left            =   1800
      TabIndex        =   31
      Top             =   8280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "« „«„ Å—Ê«‰Â"
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
      MICON           =   "FrmAhan_Detail.frx":169EA
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame FrameAhan 
      Caption         =   "´ Ê—Êœ «ÿ·«⁄«  ¬Â‰ ¬·« ª"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Traditional Arabic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4095
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   3480
      Visible         =   0   'False
      Width           =   12495
      Begin VB.TextBox TxtMobile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox TxtTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Text            =   "0"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox TxtSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox TxtShakhe 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox TxtTedad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox TxtWeight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox TxtSerial 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   7680
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox TxtKamioon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   9240
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox TxtAnbar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox TxtBarNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1440
         Width           =   2415
      End
      Begin FarDate1.FarDate TxtBarNameDate 
         Height          =   495
         Left            =   8400
         TabIndex        =   4
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
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
         BackColor       =   16777152
      End
      Begin PrjShayan.TypeButton CmdOk 
         Height          =   495
         Left            =   1560
         TabIndex        =   15
         Top             =   3360
         Width           =   1215
         _ExtentX        =   2143
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
         MICON           =   "FrmAhan_Detail.frx":16A06
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
         Left            =   360
         TabIndex        =   16
         Top             =   3360
         Width           =   1215
         _ExtentX        =   2143
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
         MICON           =   "FrmAhan_Detail.frx":16A22
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblRadif 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1"
         ForeColor       =   &H00404040&
         Height          =   390
         Left            =   5760
         TabIndex        =   57
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â „Ê»«Ì·"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   3030
         TabIndex        =   56
         Top             =   2280
         Width           =   1125
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ò· ò—«ÌÂ"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   3030
         TabIndex        =   55
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "”«Ì“"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   3030
         TabIndex        =   54
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘«ŒÂ"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   7080
         TabIndex        =   53
         Top             =   2280
         Width           =   435
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   7170
         TabIndex        =   52
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ê“‰"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   7170
         TabIndex        =   51
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ò«„ÌÊ‰"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   11115
         TabIndex        =   50
         Top             =   3120
         Width           =   1185
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«‰»«—  Œ·ÌÂ"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   11115
         TabIndex        =   49
         Top             =   2280
         Width           =   825
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ »«—‰«„Â"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   11115
         TabIndex        =   48
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â »«—‰«„Â "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   11055
         TabIndex        =   47
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label LblRadif1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â —œÌ›  :"
         ForeColor       =   &H00404040&
         Height          =   390
         Left            =   6120
         TabIndex        =   46
         Top             =   0
         Width           =   1260
      End
   End
   Begin VB.Frame FrameKantiner 
      Caption         =   "´Ê—Êœ «ÿ·«⁄«  ò«‰ Ì‰—ª"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   4095
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   3480
      Visible         =   0   'False
      Width           =   12495
      Begin VB.TextBox KTxtTedad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   4200
         TabIndex        =   23
         Text            =   "0"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox kTxtAnbar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox kTxtWeight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   4200
         TabIndex        =   22
         Text            =   "0"
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox kCombsize 
         Height          =   510
         ItemData        =   "FrmAhan_Detail.frx":16A3E
         Left            =   4200
         List            =   "FrmAhan_Detail.frx":16A40
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox kTxtKantiner 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   4200
         TabIndex        =   25
         Top             =   3360
         Width           =   2535
      End
      Begin VB.TextBox kTxtBarNumber 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   8400
         TabIndex        =   18
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox kTxtKamioon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   9600
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox kTxtSerial 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   8400
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox kTxtTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   240
         TabIndex        =   26
         Text            =   "0"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox kTxtMobile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   240
         TabIndex        =   27
         Top             =   1680
         Width           =   2295
      End
      Begin FarDate1.FarDate kTxtBarnameDate 
         Height          =   495
         Left            =   8400
         TabIndex        =   17
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
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
         BackColor       =   16777152
      End
      Begin PrjShayan.TypeButton kCmdOK 
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   2760
         Width           =   2295
         _ExtentX        =   4048
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
         MICON           =   "FrmAhan_Detail.frx":16A42
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrjShayan.TypeButton kCmdCancel 
         Height          =   495
         Left            =   240
         TabIndex        =   29
         Top             =   3360
         Width           =   2295
         _ExtentX        =   4048
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
         MICON           =   "FrmAhan_Detail.frx":16A5E
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ ò«‰ Ì‰—"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   6960
         TabIndex        =   71
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«‰»«—  Œ·ÌÂ"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   11115
         TabIndex        =   70
         Top             =   2400
         Width           =   825
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ê“‰ ò«‰ Ì‰— "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   7020
         TabIndex        =   69
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "”«Ì“ ò«‰ Ì‰—"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   6960
         TabIndex        =   68
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ò«‰ Ì‰— "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   6960
         TabIndex        =   66
         Top             =   3360
         Width           =   1140
      End
      Begin VB.Label kLblRadif1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "‘„«—Â —œÌ›  :"
         ForeColor       =   &H00404040&
         Height          =   390
         Left            =   9000
         TabIndex        =   65
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â »«—‰«„Â "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   11055
         TabIndex        =   64
         Top             =   1560
         Width           =   1140
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ »«—‰«„Â"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   11115
         TabIndex        =   63
         Top             =   525
         Width           =   1035
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ò«„ÌÊ‰ "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   11160
         TabIndex        =   62
         Top             =   3360
         Width           =   1245
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ò· ò—«ÌÂ "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   2640
         TabIndex        =   61
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â „Ê»«Ì·"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   2640
         TabIndex        =   60
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label kLblRadif 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1"
         ForeColor       =   &H00404040&
         Height          =   390
         Left            =   8640
         TabIndex        =   59
         Top             =   0
         Width           =   120
      End
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   3720
      Picture         =   "FrmAhan_Detail.frx":16A7A
      Stretch         =   -1  'True
      ToolTipText     =   "ò·Ìò ò‰Ìœ"
      Top             =   8280
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   12000
      Picture         =   "FrmAhan_Detail.frx":16C50
      Stretch         =   -1  'True
      ToolTipText     =   "ò·Ìò ò‰Ìœ"
      Top             =   8280
      Width           =   345
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      X1              =   360
      X2              =   12360
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004000&
      BorderWidth     =   3
      X1              =   360
      X2              =   12360
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "FrmAhan_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public AEL As Boolean

Private Sub CmdCancel_Click()
   Call ClearField
   'TxtBarNumber.SetFocus
End Sub

Private Sub CmdFinishDay_Click()
 Dim msg As Integer
 ''
 If FrameKantiner.Visible Then
    Call RepKantiner
 ElseIf FrameAhan.Visible And Not AEL Then
    Call RepAhan("TabAhan_Detail", "TabAhan_Tonaj")
 ElseIf FrameAhan.Visible And AEL Then
     Call RepAhan("TabAEL_Detail", "TabAEL_Tonaj")
 End If
 
 '
    msg = MsgBox("¬Ì« „«Ì· »Â „‘«ÂœÂ ê“«—‘ —Ê“«‰Â „Ì »«‘Ìœø", vbExclamation + vbYesNo + vbDefaultButton3, "„Â—Ê—“«‰")
    If msg = vbYes Then
       If FrameAhan.Visible Then
          If AEL Then
             FrmAhanRep.rAEL = True
          Else
             FrmAhanRep.rAEL = False
          End If
          ''
          FrmAhanRep.Finish = False
          FrmAhanRep.ParvaneCode = Trim(TxtParvane)
          FrmAhanRep.DayDate = Mid(TxtBarnameDate.Text, 3)
          FrmAhanRep.Show
       ElseIf FrameKantiner.Visible Then
          FrmKantinerRep.Finish = False
          FrmKantinerRep.ParvaneCode = Trim(TxtParvane)
          FrmKantinerRep.DayDate = Mid(kTxtBarnameDate.Text, 3)
          FrmKantinerRep.Show
       End If
       '
       Unload Me
    ElseIf msg = vbNo Then
       Unload Me
    End If
    '
End Sub

Private Sub CmdOk_Click()
 Dim msg As Integer
    
  msg = MsgBox("«“ ’Õ  Ê—Êœ «ÿ·«⁄«  „ÿ„∆‰ Â” Ìœø", vbQuestion + vbYesNo, "")
  If msg = vbYes Then
     If AEL Then
        If Not SaveData("TabAEL_Detail") Then Exit Sub
     Else
        If Not SaveData("TabAhan_Detail") Then Exit Sub
     End If
     '
     TxtBarNumber = Val(TxtBarNumber) + 1
     'TxtBarNameDate.Text = TxtBarNameDate.today
     'TxtAnbar = Empty
     TxtKamioon = Empty
     TxtSerial = "«Ì—«‰"
     TxtWeight = Empty
     'TxtTedad = Empty
     'TxtShakhe = Empty
     'TxtSize = Empty
     'TxtTotal = Empty
     TxtMobile = Empty
     
     LblRadif = Val(LblRadif) + 1
     '
     TxtBarNumber.SetFocus
  End If
   
End Sub

Private Sub CmdOk0_Click()
 Dim strSQL As String
 Dim Table As String
 
  If Trim(TxtSaheb) <> Empty Then

     If FrameAhan.Visible And Not AEL Then
        Table = "TabAhan_Master"
     ElseIf FrameAhan.Visible And AEL Then
        Table = "TabAEL_Master"
     ElseIf FrameKantiner.Visible Then
        Table = "TabKantiner_Master"
     End If
     '''''''''''
     strSQL = "UPDATE " & Table
     strSQL = strSQL & " SET Saheb='" & Trim(TxtSaheb) & "',"
     strSQL = strSQL & " Etebar='" & Trim(TxtEtebar) & "',"
     strSQL = strSQL & " Gharardad='" & Trim(TxtGharardad) & "'"
     strSQL = strSQL & " WHERE Parvane='" & Trim(TxtParvane) & "'"
     
     rs.Open strSQL, CNS
     '
     If FrameAhan.Visible Then
        FrameAhan.Enabled = True
        LblRadif.ForeColor = vbBlack
        LblRadif1.ForeColor = vbBlack '
        '
        Frame2.Enabled = False
        CmdOk0.Visible = False
        '
        TxtBarNumber.SetFocus
     ElseIf FrameKantiner.Visible Then
        FrameKantiner.Enabled = True
        kLblRadif.ForeColor = vbBlack
        kLblRadif1.ForeColor = vbBlack '
        '
        Frame2.Enabled = False
        CmdOk0.Visible = False
        '
        kTxtBarNumber.SetFocus
     End If
  Else
     MsgBox "«ÿ·«⁄«  ò«„· ‰Ì” ", vbExclamation, ""
     TxtSaheb.SetFocus
  End If

End Sub

Private Sub CmdFinish_Click()
 Dim MBandel, DBandel As Long
 If FrameKantiner.Visible Then
    Call RepKantiner
 ElseIf FrameAhan.Visible And Not AEL Then
    Call RepAhan("TabAhan_Detail", "TabAhan_Tonaj")
 ElseIf FrameAhan.Visible And AEL Then
    Call RepAhan("TabAEL_Detail", "TabAEL_Tonaj")
 End If
 '
 Dim msg As Integer
    msg = MsgBox("¬Ì« „«Ì· »Â „‘«ÂœÂ ê“«—‘ —Ê“ ¬Œ— „Ì »«‘Ìœø", vbExclamation + vbYesNoCancel + vbDefaultButton3, "„Â—Ê—“«‰")
    If msg = vbYes Then
       If FrameAhan.Visible Then
          Dim b As Boolean '''
          If Not AEL Then
             b = CheckBandel("TabAhan_Master", "TabAhan_Detail", MBandel, DBandel)
          ElseIf AEL Then
             b = CheckBandel("TabAEL_Master", "TabAEL_Detail", MBandel, DBandel)
          End If
          '
          If b Then
             FrmAhanRep.Finish = True
             If AEL Then
                FrmAhanRep.rAEL = True
             Else
                FrmAhanRep.rAEL = False
             End If
             FrmAhanRep.ParvaneCode = Trim(TxtParvane)
             FrmAhanRep.DayDate = Mid(TxtBarnameDate.Text, 3)
             FrmAhanRep.Show
          Else
             msg = MsgBox(" Ã„⁄ ò·  ⁄œ«œ »‰œ· Â«, »Ì‘ — «“  ⁄œ«œ »‰œ· »«—‰«„Â „Ì »«‘œ" & vbCrLf & _
                                         "¬Ì« „«Ì· »Â  ’ÕÌÕ Œÿ« œ—  ⁄œ«œ „Ì »«‘Ìœø", vbCritical + vbYesNo, "¬Â‰")
               If msg = vbYes Then ' edit Data For Bandel
                   MsgBox " ⁄œ«œ »‰œ· »«— ‰«„Â ===" & MBandel, vbInformation, ""
                   MsgBox " ⁄œ«œ »‰œ· Ê«—œ ‘œÂ ===" & DBandel, vbInformation, ""
                 
                   FrmEditDetailA.Ahan = Not AEL
                   FrmEditDetailA.ParvaneCode = TxtParvane
                   FrmEditDetailA.Show
                   Unload Me
               Else
                   Unload Me
                   Exit Sub
               End If
          End If
       ElseIf FrameKantiner.Visible Then
          If CheckBandel("TabKantiner_Master", "TabKantiner_Detail", MBandel, DBandel) Then
              FrmKantinerRep.Finish = True
              FrmKantinerRep.ParvaneCode = Trim(TxtParvane)
              FrmKantinerRep.DayDate = Mid(kTxtBarnameDate.Text, 3)
              FrmKantinerRep.Show
          Else
             msg = MsgBox(" Ã„⁄ ò·  ⁄œ«œ »‰œ· Â«, »Ì‘ — «“  ⁄œ«œ »‰œ· »«—‰«„Â „Ì »«‘œ" & vbCrLf & _
                                         "¬Ì« „«Ì· »Â  ’ÕÌÕ Œÿ« œ—  ⁄œ«œ „Ì »«‘Ìœø", vbCritical + vbYesNo, "ò«‰ Ì‰—")
             If msg = vbYes Then
                MsgBox " ⁄œ«œ ò«‰ Ì‰— »«— ‰«„Â ===" & MBandel, vbInformation, ""
                MsgBox " ⁄œ«œ ò«‰ Ì‰— Ê«—œ ‘œÂ ===" & DBandel, vbInformation, ""
                
                FrmEditDetailKan.ParvaneCode = TxtParvane
                FrmEditDetailKan.Show
                Unload Me
             Else
                Unload Me
                Exit Sub
             End If
          End If
       End If
       Unload Me
    ElseIf msg = vbNo Then
       Unload Me
    End If
     
End Sub

Private Sub Form_Activate()
   
   If FrameAhan.Caption = "´ Ê—Êœ «ÿ·«⁄«  ¬Â‰ ¬·« ª   —ŒÌ’ Å«—”Â" Or _
                Caption = "«œ«„Â Ê—Êœ «ÿ·«⁄«    —ŒÌ’ Å«—”Â" Then
      AEL = True
   Else
      AEL = False
   End If

End Sub

Private Sub Form_Load()
   Call ClearField
   '
   RightToLeft = True
   '
   BackColor = RGB(158, 179, 215)
   Frame1.BackColor = BackColor
   Frame2.BackColor = BackColor
   FrameAhan.BackColor = BackColor
   FrameKantiner.BackColor = BackColor
   LblRadif.BackColor = BackColor
   LblRadif1.BackColor = BackColor
   kLblRadif.BackColor = BackColor
   kLblRadif1.BackColor = BackColor
   Width = 13000
   '
   With kCombsize
        .AddItem "20ft"
        .AddItem "40ft"
        .AddItem String(20, "-")
        .AddItem "30ft"
        .AddItem "35ft"
        
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FrmStart.Show
End Sub

Sub ClearField()
    '
    TxtBarNumber = Empty
    TxtBarnameDate.Text = TxtBarnameDate.Today
    TxtAnbar = Empty
    TxtKamioon = Empty
    TxtSerial = "«Ì—«‰"
    TxtWeight = 0
    TxtTedad = 0
    TxtShakhe = 0
    TxtSize = 0
    TxtTotal = 0
    TxtMobile = Empty
    '''''
    kTxtBarNumber = Empty
    kTxtBarnameDate.Text = kTxtBarnameDate.Today
    kTxtKamioon = Empty
    kTxtSerial = "«Ì—«‰"
    kTxtKantiner = "IRSU "
    kTxtTotal = 0
    kTxtMobile = Empty
    kTxtAnbar = Empty
    KTxtTedad = 0
    
End Sub

Private Sub Image1_Click()
   Dim strHlp As String
   
   strHlp = " »« ò·Ìò »— —ÊÌ «Ì‰ œò„Â  »—‰«„Â «“ " + vbCrLf
   strHlp = strHlp + "‘„« ”Ê«·Ì „Ì Å—”œ òÂ »«  «ÌÌœ ¬‰ «ÿ·«⁄«  " + vbCrLf
   strHlp = strHlp + "Ìò —Ê“ ò«—Ì –ŒÌ—Â ‘œÂ Ê ê“«—‘ ¬‰ " + vbCrLf
   strHlp = strHlp + " »—«Ì ‘„« ‰„«Ì‘ œ«œÂ „Ì ‘Êœ " & vbCrLf
   strHlp = strHlp + "œ«œ‰ Å«”Œ „‰›Ì »Â «Ì‰ ”Ê«· »Â „‰“·Â «Ì‰ «”  òÂ " + vbCrLf
   strHlp = strHlp + "—Ê“ ò«—Ì  „«„ ‰‘œÂ Ê Å” «“ „œ  òÊ «ÂÌ  " + vbCrLf
   strHlp = strHlp + "ﬁ’œ «œ«„Â Ê—Êœ «ÿ·«⁄«  —Ê“ Ã«—Ì —« œ«—Ìœ " + vbCrLf + vbCrLf
   strHlp = strHlp + Space(10) + "„Ê›ﬁ »«‘Ìœ"
   
   MsgBox strHlp, vbInformation, "„Â—Ê—“«‰"
End Sub

Private Sub Image2_Click()
   Dim strHlp As String
   
   strHlp = " ò·Ìò »— —ÊÌ «Ì‰ œò„Â »Â „‰“·Â " + vbCrLf
   strHlp = strHlp + "Å«Ì«‰ «ÿ·«⁄«  Ìò Å—Ê«‰Â «”  Ê œ— ’Ê—  " + vbCrLf
   strHlp = strHlp + " «ÌÌœ ”Ê«·Ì òÂ Å—”ÌœÂ „Ì ‘Êœ " + vbCrLf
   strHlp = strHlp + "ê“«—‘ —Ê“ ¬Œ— »«—êÌ—Ì »Â Â„—«Â «ÿ·«⁄«  Ã“ÌÌ " & vbCrLf
   strHlp = strHlp + "»—«Ì ‘„« Ÿ«Â— „Ì ‘Êœ òÂ „Ì  Ê«‰Ìœ " + vbCrLf
   strHlp = strHlp + "«ÿ·«⁄«  —« ç«Å ê—› Â Ê Ì« ¬‰ —« –ŒÌ—Â ‰„«ÌÌœ" + vbCrLf + vbCrLf
   
   strHlp = strHlp + Space(10) + "„Ê›ﬁ »«‘Ìœ"
   
   MsgBox strHlp, vbInformation, "„Â—Ê—“«‰"

End Sub

Private Sub kCmdCancel_Click()
   Call ClearField
   kTxtBarNumber.SetFocus
End Sub

Private Sub kCmdOK_Click()
 Dim msg As Integer
    
  msg = MsgBox("«“ ’Õ  Ê—Êœ «ÿ·«⁄«  „ÿ„∆‰ Â” Ìœø", vbQuestion + vbYesNo, "")
  If msg = vbYes Then
     Call SaveDataKantiner
     kTxtBarNumber = Val(kTxtBarNumber) + 1
     'kTxtBarnameDate.Text = kTxtBarnameDate.today
     kTxtKamioon = Empty
     kTxtSerial = "«Ì—«‰"
     kTxtKantiner = "IRSU"
     'kTxtTotal = 0
     kTxtMobile = Empty
     
     kLblRadif = Val(kLblRadif) + 1
     '
     kTxtBarNumber.SetFocus
  End If

End Sub

Private Sub kCombsize_Click()
  If kCombsize.ListIndex = 2 Then kCombsize.ListIndex = 0
End Sub

Private Sub kCombsize_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub kTxtAnbar_GotFocus()
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard Engish
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

End Sub

Private Sub kTxtAnbar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub kTxtBarnameDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then kTxtBarNumber.SetFocus
  If KeyCode = vbKeyUp Then kTxtMobile.SetFocus
  
End Sub

Private Sub kTxtBarNumber_GotFocus()
   SendKeys "{Home}+{End}"
End Sub

Private Sub kTxtBarNumber_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
'
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Private Sub kTxtKamioon_Change()
  If Len(kTxtKamioon) = 3 Then
     kTxtKamioon = kTxtKamioon & "⁄"
     SendKeys "{End}"
  End If
  '
  If Len(kTxtKamioon) = 6 Then kTxtSerial.SetFocus
End Sub

Private Sub kTxtKamioon_GotFocus()
   SendKeys "{Home}+{End}"
End Sub

Private Sub kTxtKamioon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub kTxtKantiner_GotFocus()
   kTxtKantiner.SelStart = Len(kTxtKantiner)
End Sub

Private Sub kTxtKantiner_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub kTxtMobile_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
'
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Private Sub kTxtSerial_GotFocus()
  SendKeys "{End}"
End Sub

Private Sub kTxtSerial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub KTxtTedad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub kTxtTotal_GotFocus()
   SendKeys "{Home}+{End}"
End Sub

Private Sub kTxtTotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
'
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Private Sub kTxtWeight_GotFocus()
   SendKeys "{Home}+{End}"
End Sub

Private Sub Text2_Change()

End Sub

Private Sub kTxtWeight_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
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
End Sub

Private Sub TxtSerial_GotFocus()
  SendKeys "{End}"
End Sub

Private Sub TxtTedad_GotFocus()
   SendKeys "{Home}+{End}"
End Sub

Private Sub TxtAnbar_GotFocus()
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard Engish
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If
End Sub

Private Sub TxtAnbar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub TxtBarNameDate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub TxtBarNumber_GotFocus()
   SendKeys "{Home}+{End}"
End Sub

Private Sub TxtBarNumber_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"
'
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtEtebar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
'
End Sub

Private Sub TxtGharardad_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub TxtKamioon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub TxtMobile_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
'
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtSaheb_GotFocus()
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

Private Sub TxtSerial_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub TxtShakhe_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtShakhe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
'
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtSize_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtSize_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
'
' Dim strValid As String
'   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
'   If InStr(strValid, Chr(KeyAscii)) = 0 Then
'      KeyAscii = 0
'   End If
End Sub

Private Sub kTxtTedad_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtTedad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
'
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtTotal_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtTotal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
'
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If KeyAscii = 46 Then KeyAscii = 0
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtWeight_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtWeight_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
'
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Function SaveData(Table As String) As Boolean
On Error GoTo L:
   Dim strSQL As String
   Dim strKamioon As String
   
   SaveData = True
   strKamioon = TxtKamioon & TxtSerial

''''''''''''''
   strSQL = "INSERT INTO " & Table & " "
   strSQL = strSQL & "(Parvane,Count0,BarNumber,DBarname,Anbar,"
   strSQL = strSQL & "Kamioon,Weight,Tedad,Shakhe,Size0,"
   strSQL = strSQL & "Total,Mobile) "
   '
   strSQL = strSQL & "VALUES('" & Trim(TxtParvane) & "',"
   strSQL = strSQL & Val(LblRadif) & ","
   strSQL = strSQL & "'" & Trim(TxtBarNumber) & "',"
   strSQL = strSQL & "'" & Mid(TxtBarnameDate.Text, 3) & "',"
   strSQL = strSQL & "'" & Trim(TxtAnbar) & "',"
   strSQL = strSQL & "'" & Trim(strKamioon) & "',"
   strSQL = strSQL & Val(TxtWeight) & ","
   strSQL = strSQL & Val(TxtTedad) & ","
   strSQL = strSQL & Val(TxtShakhe) & ",'"
   strSQL = strSQL & TxtSize & "',"
   strSQL = strSQL & CCur(TxtTotal) & ","
   strSQL = strSQL & "'" & Trim(TxtMobile) & "')"
   '
   rs.Open strSQL, CNS
L:  If Err.Number <> 0 Then
       MsgBox "‘„«—Â —œÌ› œ— »«‰ò  ò—«—Ì «” ", vbCritical, ""
       SaveData = False
    End If
End Function


Private Sub SaveDataKantiner()
On Error Resume Next
   Dim strSQL As String
   Dim strKamioon As String
      
   strKamioon = kTxtKamioon & kTxtSerial

''''''''''''''
   strSQL = "INSERT INTO TabKantiner_Detail "
   strSQL = strSQL & "(Parvane,Count0,BarNumber,"
   strSQL = strSQL & "BarNameDate,Anbar,Kamioon,Kantiner,Weight,"
   strSQL = strSQL & "Tedad,Shakhe,Size0,Total,Mobile) "
   '
   strSQL = strSQL & "VALUES('" & Trim(TxtParvane) & "',"
   strSQL = strSQL & Val(kLblRadif) & ","
   strSQL = strSQL & "'" & Trim(kTxtBarNumber) & "',"
   strSQL = strSQL & "'" & Mid(kTxtBarnameDate.Text, 3) & "',"
   strSQL = strSQL & "'" & Trim(kTxtAnbar) & "',"
   strSQL = strSQL & "'" & Trim(strKamioon) & "',"
   strSQL = strSQL & "'" & Trim(UCase(kTxtKantiner)) & "',"
   strSQL = strSQL & Val(kTxtWeight) & ","
   strSQL = strSQL & Val(KTxtTedad) & ",0,"
   strSQL = strSQL & Val(kCombsize.Text) & ","
   strSQL = strSQL & CCur(kTxtTotal) & ","
   strSQL = strSQL & "'" & Trim(kTxtMobile) & "')"
   '
  
   rs.Open strSQL, CNS

End Sub

Sub RepKantiner()
On Error Resume Next
  Dim Weight As Long
  Dim Tedad As Long
  Dim Price As Currency
  '
  rs.Open "SELECT SUM(Weight),SUM(Tedad) ," & _
          "SUM(Total) " & _
          "FROM TabKantiner_Detail " & _
          "WHERE Parvane='" & TxtParvane & "' " & _
          "AND BarNameDate='" & Mid(kTxtBarnameDate.Text, 3) & "'", CNS
 
 If Not rs.EOF Then
    Weight = rs(0)
    Tedad = rs(1)
    Price = rs(2)
    '''
 End If
 rs.Close
 
 '''Date Is Empty (Avalin Bar)
 
 rs.Open "SELECT * FROM TabKantiner_Tonaj " & _
         "WHERE Parvane='" & TxtParvane & "' " & _
         "AND BarDate='n'", CNS
 '''
 If Not rs.EOF Then 'Date Is Empty (Avalin Bar)
    rs.Close
    rs.Open "UPDATE TabKantiner_Tonaj " & _
            "SET BarDate='" & Mid(kTxtBarnameDate.Text, 3) & "'," & _
            "TonajEx=" & Weight & ",TotalBandel=" & Tedad & "," & _
            "TonajMod=TonajPar-" & Weight & "," & _
            "TotalPrice=" & Price & _
            " WHERE Parvane='" & TxtParvane & "'", CNS
 Else
    '''''''''''DafaAte Badi dar rooze jari
     rs.Close
     rs.Open "SELECT * FROM TabKantiner_Tonaj " & _
             "WHERE Parvane='" & TxtParvane & "' " & _
             "AND BarDate='" & Mid(kTxtBarnameDate.Text, 3) & "'", CNS
     If Not rs.EOF Then
        rs.Close
        rs.Open "UPDATE TabKantiner_Tonaj " & _
                "SET TonajEx=" & Weight & ",TotalBandel=" & Tedad & "," & _
                "TonajMod=TonajPar-" & Weight & "," & _
                "TotalPrice=" & Price & _
                " WHERE Parvane='" & TxtParvane & "' " & _
                "AND BarDate='" & Mid(kTxtBarnameDate.Text, 3) & "'", CNS
     Else
        '''''''Roozhaye BaAdi
        Dim Yesterday_ As String
        Dim yTonajMod As Long
        
        rs.Close
        rs.Open "SELECT MAX(BarDate) FROM TabKantiner_Tonaj " & _
                "WHERE Parvane='" & TxtParvane & "' ", CNS
        Yesterday_ = rs(0)
        If CDate(Mid(kTxtBarnameDate.Text, 3)) > CDate(rs(0)) Then
           rs.Close
           rs.Open "SELECT TonajMod FROM TabKantiner_Tonaj " & _
                   "WHERE Parvane='" & TxtParvane & "' " & _
                   "AND BarDate='" & Yesterday_ & "'", CNS
           yTonajMod = rs(0)
           '''
           rs.Close
           rs.Open "INSERT INTO TabKantiner_Tonaj " & _
                   "(Parvane,BarDate,TonajPar,TonajEx,TonajMod," & _
                   "TotalBandel,TotalPrice) " & _
                   "VALUES('" & TxtParvane & "','" & _
                   Mid(kTxtBarnameDate.Text, 3) & "'," & _
                   yTonajMod & "," & Weight & "," & yTonajMod - Weight & "," & _
                   Tedad & "," & Price & ")", CNS
           'Val(TxtNWeight)
        End If
     End If
 End If

End Sub

Sub RepAhan(Table As String, Tonaj As String)
 On Error Resume Next
 
  Dim Weight As Long
  Dim Tedad As Long
  Dim Shakhe As Long
  Dim Price As Currency
  '
  rs.Open "SELECT SUM(Weight)," & _
          "SUM(Shakhe),SUM(Total) " & _
          "FROM " & Table & " " & _
          "WHERE Parvane='" & TxtParvane & "' " & _
          "AND DBarname='" & Mid(TxtBarnameDate.Text, 3) & "'", CNS
 
 If Not rs.EOF Then
    Weight = rs(0)
    Shakhe = rs(1)
    Price = rs(2)
    '''
 End If
 rs.Close
 '
 rs.Open "SELECT * FROM " & Tonaj & " WHERE Parvane='" & TxtParvane & "' " & _
          "AND BarDate='" & Mid(TxtBarnameDate.Text, 3) & "'", CNS
 
 If rs.EOF Then ' if not found --> jame tamame tedade KOLE Parvane
    rs.Close
    rs.Open "SELECT SUM(Tedad) FROM " & Table & _
             " WHERE Parvane='" & TxtParvane & "' ", CNS
    Tedad = rs(0)
 Else ' jem tedad az tarikhe aval ta haman tarikh
    rs.Close
    ' avalin tarikhe parvane
    Dim FirstDate As String
    
    rs.Open "SELECT MIN(BarDate) FROM " & Tonaj & " " & _
                "WHERE Parvane='" & TxtParvane & "' ", CNS
    FirstDate = rs(0)
    rs.Close
    ''
    rs.Open "SELECT SUM(Tedad) FROM " & Table & _
            " WHERE Parvane='" & TxtParvane & "' " & _
            "AND (DBarname BETWEEN '" & FirstDate & "' AND '" & Mid(TxtBarnameDate.Text, 3) & "')", CNS
    Tedad = rs(0)
    rs.Close
 End If
 rs.Close
 '
 FrmAhanRep.DateBarname = Mid(TxtBarnameDate.Text, 3)
 '
 '''Date Is Empty (Avalin Bar)
 ''''For Bandel Calculate'''''''''''
     Dim MainBandel As Integer
     Dim rsBandel As New Recordset
     Dim MasTable As String
     
     If Table = "TabAhan_Detail" Then
        MasTable = "TabAhan_Master"
     ElseIf Table = "TabAEL_Detail" Then
        MasTable = "TabAEL_Master"
     End If
     
     rsBandel.Open "SELECT Bandel FROM " & MasTable & _
                   " WHERE Parvane='" & TxtParvane & "'", CNS
     MainBandel = rsBandel(0)
     rsBandel.Close
     Set rsBandel = Nothing
 ''''''''''''''''''''''''''''''''''''
 rs.Open "SELECT * FROM " & Tonaj & " " & _
         "WHERE Parvane='" & TxtParvane & "' " & _
         "AND BarDate='n'", CNS
 '''
 If Not rs.EOF Then 'Date Is Empty (Avalin Bar)
    rs.Close
    rs.Open "UPDATE " & Tonaj & " " & _
            "SET BarDate='" & Mid(TxtBarnameDate.Text, 3) & "'," & _
            "TonajEx=" & Weight & ",TotalBandel=" & MainBandel - Tedad & "," & _
            "TonajMod=TonajPar-" & Weight & "," & _
            "TotalShakhe=" & Shakhe & ",TotalPrice=" & Price & _
            " WHERE Parvane='" & TxtParvane & "'", CNS
 Else
    '''''''''''DafaAte Badi dar rooze jari
     rs.Close
     Dim TBand As Long
     'rs.Open "SELECT SUM(Tedad) FROM " & Table & " " & _
     '        "WHERE Parvane='" & TxtParvane & "' " & _
     '        "AND DBarname='" & Mid(TxtBarnameDate.Text, 3) & "'", CNS
     'TBand = rs(0)
     'rs.Close
     '''
     rs.Open "SELECT * FROM " & Tonaj & " WHERE Parvane='" & TxtParvane & "' " & _
             "AND BarDate='" & Mid(TxtBarnameDate.Text, 3) & "'", CNS
     If Not rs.EOF Then
        TBand = Tedad
        rs.Close
        rs.Open "UPDATE " & Tonaj & " " & _
                "SET TonajEx=" & Weight & ",TotalBandel=" & MainBandel - TBand & "," & _
                "TonajMod=TonajPar-" & Weight & "," & _
                "TotalShakhe=" & Shakhe & ",TotalPrice=" & Price & _
                " WHERE Parvane='" & TxtParvane & "' " & _
                "AND BarDate='" & Mid(TxtBarnameDate.Text, 3) & "'", CNS
     Else
        '''''''Roozhaye BaAdi
        Dim Yesterday_ As String
        Dim yTonajMod As Long
        Dim yBandel As Integer
        rs.Close
        rs.Open "SELECT MAX(BarDate) FROM " & Tonaj & " " & _
                "WHERE Parvane='" & TxtParvane & "' ", CNS
        Yesterday_ = rs(0)
        If CDate(Mid(TxtBarnameDate.Text, 3)) >= CDate(rs(0)) Then
           rs.Close
           rs.Open "SELECT TonajMod,TotalBandel FROM " & Tonaj & " " & _
                   "WHERE Parvane='" & TxtParvane & "' " & _
                   "AND BarDate='" & Yesterday_ & "'", CNS
           yTonajMod = rs(0)
           yBandel = rs(1)
           '''
           rs.Close
           rs.Open "INSERT INTO " & Tonaj & " " & _
                   "(Parvane,BarDate,TonajPar,TonajEx,TonajMod," & _
                   "TotalBandel,TotalShakhe,TotalPrice) " & _
                   "VALUES('" & TxtParvane & "','" & _
                   Mid(TxtBarnameDate.Text, 3) & "'," & _
                   yTonajMod & "," & Weight & "," & yTonajMod - Weight & "," & _
                   MainBandel - Tedad & "," & Shakhe & "," & Price & ")", CNS
           'Val(TxtNWeight)
        Else
           rs.Close
        End If
     End If
 End If

End Sub


Function CheckBandel(MTable, DTable As String, ByRef MBandel, DBandel As Long) As Boolean
    CheckBandel = True
    '
    rs.Open "SELECT * FROM " & MTable & _
            " WHERE Parvane='" & TxtParvane & "'", CNS
    MBandel = rs("Bandel")
    rs.Close
    '''
    rs.Open "SELECT SUM(Tedad) FROM " & DTable & _
            " WHERE Parvane='" & TxtParvane & "'", CNS
    DBandel = rs(0)
    rs.Close
    
    If MBandel < DBandel Then
       CheckBandel = False
    End If
            
End Function
