VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmFindSaderat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ã” ÃÊÌ «ÿ·«⁄«  ò«‰ Ì‰— Â«Ì ’«œ—« Ì"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9585
   BeginProperty Font 
      Name            =   "B Zar"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmFindSaderat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin PrjShayan.TypeButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   7560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      BTYPE           =   1
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
      MICON           =   "FrmFindSaderat.frx":169B2
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TabDlg.SSTab TabBox1 
      Height          =   7335
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   1411
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483640
      MouseIcon       =   "FrmFindSaderat.frx":169CE
      TabCaption(0)   =   "Ã” ÃÊÌ ò«‰ Ì‰— Ê „‘«ÂœÂ ê“«—‘"
      TabPicture(0)   =   "FrmFindSaderat.frx":169EA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TxtFindKantiner"
      Tab(0).Control(1)=   "CmdFindKantiner"
      Tab(0).Control(2)=   "Grid1"
      Tab(0).Control(3)=   "CmdPrint"
      Tab(0).Control(4)=   "Label2"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "(Ã” ÃÊÌ —œÌ› „—“Ì ( Œ·ÌÂ ò«‰ Ì‰—"
      TabPicture(1)   =   "FrmFindSaderat.frx":16A06
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Line4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "LabelX(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "LabelX(3)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "LabelX(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "LabelX(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "LabelX(5)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "LabelX(6)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "LabelX(7)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "LabelX(8)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "LabelX(9)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "LabelX(10)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label1"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "CmdTakhlie"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "CmdEditMaster"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "CmdContinue"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "CmdFindRadifMarzi"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "TxtTransitDate"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "TxtTransitNo"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "TxtKootaj"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "TxtRadifMarzi"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "TxtBarnameDarya"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "TxtTypeProduct"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "TxtPart"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "CombPackage"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "TxtTarkhiskar"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "TxtSaheb"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "TxtFindRadifMarzi"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).ControlCount=   30
      Begin VB.TextBox TxtFindRadifMarzi 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   2640
         TabIndex        =   0
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox TxtSaheb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   360
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   5760
         Width           =   3135
      End
      Begin VB.TextBox TxtTarkhiskar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   360
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   4920
         Width           =   3135
      End
      Begin VB.ComboBox CombPackage 
         Height          =   510
         ItemData        =   "FrmFindSaderat.frx":2D3C8
         Left            =   360
         List            =   "FrmFindSaderat.frx":2D3D5
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4080
         Width           =   3135
      End
      Begin VB.TextBox TxtPart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   360
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   3240
         Width           =   3135
      End
      Begin VB.TextBox TxtTypeProduct 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   360
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2400
         Width           =   3135
      End
      Begin VB.TextBox TxtBarnameDarya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   5160
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   5760
         Width           =   2535
      End
      Begin VB.TextBox TxtRadifMarzi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   5160
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   4920
         Width           =   2535
      End
      Begin VB.TextBox TxtKootaj 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   5160
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox TxtTransitNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800080&
         Height          =   510
         Left            =   5160
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox TxtTransitDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   7
         Top             =   4080
         Width           =   2535
      End
      Begin VB.TextBox TxtFindKantiner 
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   -71760
         MaxLength       =   12
         TabIndex        =   15
         Text            =   "IRSU"
         Top             =   1440
         Width           =   3015
      End
      Begin PrjShayan.TypeButton CmdFindRadifMarzi 
         Default         =   -1  'True
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   1
         TX              =   " «ÌÌœ Ê Ã” ÃÊ"
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
         MICON           =   "FrmFindSaderat.frx":2D3ED
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
         Left            =   6840
         TabIndex        =   2
         Top             =   6720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BTYPE           =   1
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
         COLTYPE         =   3
         FOCUSR          =   -1  'True
         BCOL            =   12640511
         BCOLO           =   12640511
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmFindSaderat.frx":2D409
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
         Left            =   3600
         TabIndex        =   3
         Top             =   6720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BTYPE           =   1
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
         COLTYPE         =   3
         FOCUSR          =   -1  'True
         BCOL            =   12640511
         BCOLO           =   12640511
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmFindSaderat.frx":2D425
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrjShayan.TypeButton CmdTakhlie 
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   6720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BTYPE           =   1
         TX              =   " Œ·ÌÂ ò«‰ Ì‰— Â«"
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
         COLTYPE         =   3
         FOCUSR          =   -1  'True
         BCOL            =   12640511
         BCOLO           =   12640511
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmFindSaderat.frx":2D441
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PrjShayan.TypeButton CmdFindKantiner 
         Height          =   495
         Left            =   -73920
         TabIndex        =   16
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   1
         TX              =   " «ÌÌœ Ê Ã” ÃÊ"
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
         MICON           =   "FrmFindSaderat.frx":2D45D
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin FlexCell.Grid Grid1 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   20
         Top             =   2160
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   7646
         Cols            =   12
         DefaultFontName =   "B Zar"
         DefaultFontSize =   12
         DefaultFontBold =   -1  'True
         DefaultRowHeight=   32
         GridColor       =   -2147483630
         GridLiness      =   -1  'True
         Rows            =   15
      End
      Begin PrjShayan.TypeButton CmdPrint 
         Height          =   495
         Left            =   -68520
         TabIndex        =   17
         Top             =   6600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         BTYPE           =   1
         TX              =   "ÅÌ‘ ‰„«Ì‘ Ê ç«Å ê“«—‘"
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
         MICON           =   "FrmFindSaderat.frx":2D479
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—œÌ› „—“Ì —« Ê«—œ ‰„«ÌÌœ "
         Height          =   390
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1440
         Width           =   2355
      End
      Begin VB.Label LabelX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "’«Õ» ò«·«"
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   10
         Left            =   3600
         TabIndex        =   31
         Top             =   5760
         Width           =   960
      End
      Begin VB.Label LabelX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„  —ŒÌ’ ò«—"
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   9
         Left            =   3600
         TabIndex        =   30
         Top             =   5040
         Width           =   1275
      End
      Begin VB.Label LabelX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ »” Â »‰œÌ"
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   8
         Left            =   3600
         TabIndex        =   29
         Top             =   4080
         Width           =   1230
      End
      Begin VB.Label LabelX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Å‹‹«— "
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   7
         Left            =   3600
         TabIndex        =   28
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label LabelX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ ò‹‹«·«"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   6
         Left            =   3600
         TabIndex        =   27
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label LabelX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "»«—‰«„Â œ—Ì«ÌÌ"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   5
         Left            =   7920
         TabIndex        =   26
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label LabelX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â  —«‰“Ì "
         Enabled         =   0   'False
         ForeColor       =   &H00800080&
         Height          =   495
         Index           =   1
         Left            =   7920
         TabIndex        =   25
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label LabelX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â òÊ «é"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   2
         Left            =   7920
         TabIndex        =   24
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label LabelX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ  —«‰“Ì‹ "
         Height          =   495
         Index           =   3
         Left            =   7920
         TabIndex        =   23
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label LabelX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—œÌ› „—“Ì"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   4
         Left            =   7920
         TabIndex        =   22
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         X1              =   360
         X2              =   8880
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         X1              =   360
         X2              =   8880
         Y1              =   2150
         Y2              =   2150
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         X1              =   360
         X2              =   8880
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000E&
         X1              =   360
         X2              =   8880
         Y1              =   6495
         Y2              =   6495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ò«‰ Ì‰— —« Ê«—œ ò‰Ìœ"
         Height          =   390
         Left            =   -68370
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1440
         Width           =   2205
      End
   End
End
Attribute VB_Name = "FrmFindSaderat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_Click()
  Unload Me
End Sub

Private Sub CmdContinue_Click()
  If TxtTransitNo = Empty Then Exit Sub
  '
  Dim rsT As New Recordset
  Dim Max As Long
  
  rsT.Open "SELECT COUNT(Count0),MAX(Count0) FROM TabSaderat_Detail " & _
           "WHERE TransitNo='" & TxtTransitNo & "'", CNS
  Max = IIf(rsT(0) = 0, 1, rsT(0) + 1)
  rsT.Close
  Set rsT = Nothing
  
  If Val(TxtPart) < Max Then
     MsgBox " ⁄œ«œ —œÌ› »—«»— »«  ⁄œ«œ Å«—  «” " & vbCrLf & _
            "‘„« «Ã«“Â «œ«„Â Ê—Êœ «ÿ·«⁄«  —« ‰œ«—Ìœ", vbExclamation
     
     Exit Sub
  End If
  
  With FrmSaderatDetail
       .LblRadif = Max
       .TxtTransitNo = TxtTransitNo
       .TxtKootaj = TxtKootaj
       .TxtRadifMarzi = TxtRadifMarzi
       .TxtPart = TxtPart
       Unload Me
       .Show
  End With
  
End Sub

Private Sub CmdEditMaster_Click()
  If CmdEditMaster.Caption = "«’·«Õ «ÿ·«⁄«  «’·Ì" Then
     If TxtTransitNo = Empty Then Exit Sub
     '
     CmdContinue.Enabled = False
     CmdTakhlie.Enabled = False
     '
     EnableField False
     '
     CmdEditMaster.Caption = "À»   €ÌÌ—« "
     '
     TxtKootaj.SetFocus
 ElseIf CmdEditMaster.Caption = "À»   €ÌÌ—« " Then
     Dim rsT As New Recordset
     Dim strSQL As String
     '
     strSQL = "UPDATE TabSaderat_Master "
     strSQL = strSQL & "SET KootajNo='" & TxtKootaj & "',"
     strSQL = strSQL & "TransitDate='" & TxtTransitDate & "',"
     strSQL = strSQL & "RadifMarzi=" & Val(TxtRadifMarzi) & ","
     strSQL = strSQL & "BarnameDarya='" & TxtBarnameDarya & "',"
     strSQL = strSQL & "TypeProduct='" & TxtTypeProduct & "',"
     strSQL = strSQL & "Part=" & Val(TxtPart) & ","
     strSQL = strSQL & "TypePackage='" & CombPackage.Text & "',"
     strSQL = strSQL & "Tarkhiskar='" & TxtTarkhiskar & "',"
     strSQL = strSQL & "Saheb='" & TxtSaheb & "' "
     strSQL = strSQL & "WHERE TransitNo='" & TxtTransitNo & "'"
     '
     rsT.Open strSQL, CNS
     Set rsT = Nothing
     '
     CmdContinue.Enabled = True
     CmdTakhlie.Enabled = True
     '
     EnableField True
     '
     CmdEditMaster.Caption = "«’·«Õ «ÿ·«⁄«  «’·Ì"
  End If
  
End Sub

Private Sub CmdFindKantiner_Click()
  If TxtFindKantiner = Empty Then Exit Sub
  Grid1.Rows = 1
  
  Dim rsT As New Recordset
  Dim strSQL As String
  Dim i As Integer
  
  strSQL = "SELECT TabSaderat_Master.TransitNo,TabSaderat_Master.Part,"
  strSQL = strSQL & "TabSaderat_Master.RadifMarzi,TabSaderat_Master.BarnameDarya,"
  strSQL = strSQL & "TabSaderat_Master.TransitDate "
  strSQL = strSQL & "FROM TabSaderat_Master INNER JOIN TabSaderat_Detail ON "
  strSQL = strSQL & "TabSaderat_Master.TransitNo = TabSaderat_Detail.TransitNo "
  strSQL = strSQL & "WHERE (((TabSaderat_Detail.Kantiner)='" & TxtFindKantiner & "')) "
  strSQL = strSQL & "ORDER BY TabSaderat_Master.TransitDate,"
  strSQL = strSQL & "TabSaderat_Master.TransitNo"
  
  rsT.Open strSQL, CNS
  '
  Dim b As Boolean
  
  b = False
  With Grid1
       While Not rsT.EOF
            .AddItem ""
            For i = 1 To 5
                .Cell(.Rows - 1, i).Text = rsT(5 - i)
            Next
            '
            rsT.MoveNext
            b = True
       Wend
       rsT.Close
       Set rsT = Nothing
  End With
  If Not b Then MsgBox "ò«‰ Ì‰— „Ê—œ ‰Ÿ— ÅÌœ« ‰‘œ", vbExclamation, ""
  '
  TxtFindKantiner.SetFocus
  TxtFindKantiner.SelStart = 4
  TxtFindKantiner.SelLength = 9
End Sub

Private Sub CmdFindRadifMarzi_Click()
  If TxtFindRadifMarzi = Empty Then Exit Sub
  
  Dim rsT As New Recordset
  
  rsT.Open "SELECT * FROM TabSaderat_Master " & _
           "WHERE RadifMarzi=" & Val(TxtFindRadifMarzi), CNS
  If rsT.EOF Then
     rsT.Close
     Set rsT = Nothing
     MsgBox "«ÿ·«⁄«  „Ê—œ ‰Ÿ— ÅÌœ« ‰‘œ", vbExclamation, "MehrVarzan"
     TxtFindRadifMarzi.SetFocus
     Exit Sub
  End If
  '
  TxtTransitNo = rsT(0)
  TxtKootaj = rsT(1)
  TxtTransitDate = rsT(2)
  TxtRadifMarzi = rsT(3)
  TxtBarnameDarya = IIf(IsNull(rsT(4)), "", rsT(4))
  TxtTypeProduct = IIf(IsNull(rsT(5)), "", rsT(5))
  TxtPart = rsT(6)
  If Not IsNull(rsT(7)) Then
     CombPackage.Locked = False
     If rsT(7) = "Å«· " Then CombPackage.ListIndex = 0
     If rsT(7) = "‰ê·Â" Then CombPackage.ListIndex = 1
     If rsT(7) = "œ” ê«Â" Then CombPackage.ListIndex = 2
     CombPackage.Locked = True
 End If
 '
 TxtTarkhiskar = IIf(IsNull(rsT(8)), "", rsT(8))
 TxtSaheb = IIf(IsNull(rsT(9)), "", rsT(9))
 '
 rsT.Close
 Set rsT = Nothing
End Sub

Private Sub CmdPrint_Click()
  With Grid1.PageSetup
     
     .PaperSize = cellPaperA4  'A4 paper
     .Orientation = cellPortrait  'Portrait
     '.LeftMargin = 1
     '.RightMargin = 1
     '.BottomMargin = 2.5
     '.TopMargin = 1
     .CenterHorizontally = True  'Center on page horizontally
     .BlackAndWhite = True
     .PrintFixedRow = True
     .PrintGridlines = True
     .FooterFont.Name = "Traditional Arabic"
     .FooterFont.Bold = True
     .FooterFont.Size = 13
     .FooterMargin = 0.5
     .Footer = " »‰œ— «‰“·Ì°€«“Ì«‰° «» œ«Ì ŒÌ«»«‰ —„÷«‰Ì°Ã‰» »«‰ò ’«œ—«  ‘⁄»Â »‰«œ— Ê ò‘ Ì—«‰Ì°òÊçÂ ‘ÂÌœ ”Ì—Ì°”«Œ „«‰  Ã«—Ì »—«œ—«‰ „Ãœ ÅÊ—° ÿ»ﬁÂ œÊ„ " & " E-Mail: mehrvarzantarabar@yahoo.com" & Space(0) & vbCrLf & _
               "òœ Å” Ì 73337-43156" & Space(10) & "›«ò” 3239400-0181" & Space(10) & " ·›‰ 4-3239880-0181" & Space(15) & "’›ÕÂ &P"

    '
  End With
  
  Grid1.PrintPreview

End Sub

Private Sub CmdTakhlie_Click()
 Dim i As Integer
 
   If TxtTransitNo = Empty Then Exit Sub
   '
   With FrmTakhlieKantiner
        .TxtTransitNo = TxtTransitNo
        .TxtTransitDate = TxtTransitDate
        .TxtKootaj = TxtKootaj
        .TxtPart = TxtPart
        .TxtRadifMarzi = TxtRadifMarzi
        .TxtBarnameDarya = TxtBarnameDarya
        '
        Dim rsT As New Recordset
        
        rsT.Open "SELECT * FROM TabSaderat_Detail " & _
                 "WHERE TransitNo='" & TxtTransitNo & "' " & _
                 "ORDER BY Count0 ", CNS
        If rsT.EOF Then
           MsgBox "«ÿ·«⁄« Ì »—«Ì «Ì‰ ‘„«—Â  —«‰“Ì  „ÊÃÊœ ‰„Ì »«‘œ", vbExclamation
           TxtFindRadifMarzi.SetFocus
           rsT.Close
           Set rsT = Nothing
           Exit Sub
        End If
        '
        With .Grid1
             While Not rsT.EOF
                  .AddItem ""
                  For i = 1 To 9
                      If i = 3 Then
                         .Cell(.Rows - 1, i).Text = rsT(7) & "ft"
                      ElseIf i = 2 Then
                         .Cell(.Rows - 1, i).Text = Format(rsT(8), "#,##0")
                      Else
                         .Cell(.Rows - 1, i).Text = rsT(10 - i)
                      End If
                  Next
                  '
                  rsT.MoveNext
             Wend
             rsT.Close
             Set rsT = Nothing
        End With
   End With
   '
   Unload Me
   FrmTakhlieKantiner.Show
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
   'Me.BackColor = RGB(250, 160, 160)
   '
   TabBox1.BackColor = Me.BackColor
   TabBox1.Tab = 1
   '
   ClearField
   EnableField True
   '
   SetGridProp
End Sub

Private Sub Label4_Click(Index As Integer)

End Sub

Private Sub TabBox1_Click(PreviousTab As Integer)
 On Error Resume Next
  If PreviousTab = 0 Then
     TabBox1.TabPicture(0) = LoadPicture("")
     TabBox1.TabPicture(1) = LoadPicture(App.Path & "\ICON\arvin icon\binoculars.ico")
     TxtFindRadifMarzi.SetFocus
     '
     CmdFindRadifMarzi.Default = True
  Else
     TabBox1.TabPicture(1) = LoadPicture("")
     TabBox1.TabPicture(0) = LoadPicture(App.Path & "\ICON\arvin icon\binoculars.ico")
     TxtFindKantiner.SetFocus
     '
     CmdFindKantiner.Default = True
  End If
End Sub

Private Sub TxtBarnameDarya_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtBarnameDarya_KeyPress(KeyAscii As Integer)

 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Private Sub TxtFindKantiner_Change()
  If Len(TxtFindKantiner) = 10 Then
     TxtFindKantiner = TxtFindKantiner & "-"
     SendKeys "{END}"
  End If
End Sub

Private Sub TxtFindKantiner_GotFocus()
   SendKeys "{End}"
End Sub

Private Sub TxtFindRadifMarzi_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtFindRadifMarzi_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CmdFindRadifMarzi_Click
 
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Sub ClearField()
   TxtTransitNo = Empty
   TxtKootaj = Empty
   TxtTransitDate.Text = Empty
   TxtRadifMarzi = Empty
   TxtBarnameDarya = Empty
   TxtTypeProduct = Empty
   TxtPart = Empty
   CombPackage.ListIndex = -1
   TxtTarkhiskar = Empty
   TxtSaheb = Empty
End Sub

Sub EnableField(Optional e As Boolean = True)
Dim i As Integer
    
    TxtTransitNo.Enabled = False
    TxtKootaj.Locked = e
    TxtTransitDate.Locked = e
    TxtRadifMarzi.Locked = e
    TxtBarnameDarya.Locked = e
    TxtTypeProduct.Locked = e
    TxtPart.Locked = e
    CombPackage.Locked = e
    TxtTarkhiskar.Locked = e
    TxtSaheb.Locked = e

End Sub

Private Sub TxtKootaj_KeyPress(KeyAscii As Integer)

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

 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Private Sub TxtSaheb_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtTarkhiskar_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtTransitDate_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub TxtTransitDate_LostFocus()
   TxtTransitDate = Format(TxtTransitDate, "YY/MM/DD")
End Sub

Private Sub TxtTransitNo_KeyPress(KeyAscii As Integer)

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

Sub SetGridProp()
 Dim i As Integer
  With Grid1
       .Cols = 6
       .Rows = 1
       ''
       .DefaultFont.Name = "Traditional Arabic"
       .DefaultFont.Size = 14
       .DefaultFont.Bold = True
       '
       .DisplayRowIndex = True
       .AllowUserResizing = False
       .MultiSelect = True
       '
       .Appearance = Flat
       .ScrollBarStyle = Flat
       ''
       .BackColorFixed = RGB(138, 200, 100)
       .BackColorFixedSel = RGB(188, 250, 100)
       .BackColorBkg = vbButtonFace 'RGB(90, 158, 214)
       .BackColorScrollBar = RGB(200, 135, 200)
       .BackColor1 = RGB(231, 247, 235)
       .BackColor2 = RGB(239, 255, 243)
       .GridColor = RGB(148, 231, 190)
       ''
       .Cell(0, 1).Text = " «—ÌŒ  —«‰“Ì "
       .Cell(0, 2).Text = "»«—‰«„Â œ—Ì«ÌÌ"
       .Cell(0, 3).Text = "—œÌ› „—“Ì"
       .Cell(0, 4).Text = "Å«— "
       .Cell(0, 5).Text = "‘„«—Â  —«‰“Ì "
       '
       For i = 1 To 5
           .Column(i).Alignment = cellCenterCenter
           .Column(i).Locked = True
       Next
       
       .Column(0).Width = 30
       .Column(1).Width = 100 ' Transit Date
       .Column(2).Width = 105 ' Barname DAryaEE
       .Column(3).Width = 110 ' Radif Marzi
       .Column(4).Width = 40 ' Part
       .Column(5).Width = 180 ' Transit NO
       ''
  End With
End Sub

Private Sub TypeButton1_Click()

End Sub
