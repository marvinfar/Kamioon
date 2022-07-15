VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form FrmTakhlieKantiner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›—„  Œ·ÌÂ ò«‰ Ì‰— Â«"
   ClientHeight    =   10515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11520
   BeginProperty Font 
      Name            =   "B Zar"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTakhlieKantiner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10515
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin FlexCell.Grid GrdVirtual 
      Height          =   495
      Left            =   2760
      TabIndex        =   29
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   1
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "ç«Å"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox TxtFind 
      Height          =   510
      Left            =   5880
      MaxLength       =   12
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "IRSU"
      Top             =   2160
      Width           =   3255
   End
   Begin VB.ListBox LstDelayRow 
      Height          =   840
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton CmdCalcDelay 
      Caption         =   "Delay"
      Height          =   390
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   9960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox LstDelay 
      Height          =   840
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   9600
      Width           =   9735
   End
   Begin FlexCell.Grid Grid1 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   9340
      Cols            =   12
      DefaultFontName =   "B Zar"
      DefaultFontSize =   12
      DefaultFontBold =   -1  'True
      DefaultRowHeight=   32
      GridColor       =   -2147483630
      GridLiness      =   -1  'True
      Rows            =   15
   End
   Begin VB.Frame FrameMaster 
      Caption         =   "«ÿ·«⁄«  «’·Ì ò«‰ Ì‰— ’«œ—« Ì"
      Height          =   1935
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   11295
      Begin VB.TextBox TxtBarnameDarya 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox TxtTransitDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox TxtTransitNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "125547899"
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox TxtKootaj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "12345678"
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox TxtRadifMarzi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox TxtPart 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   375
         Left            =   6000
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label LabelX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "»«—‰«„Â œ—Ì«ÌÌ"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   5
         Left            =   6000
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label LabelX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ  —«‰“Ì‹ "
         Height          =   495
         Index           =   3
         Left            =   2280
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—œÌ› „—“Ì"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2280
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â òÊ «é"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6000
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â  —«‰“Ì "
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   9720
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Å«— "
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   9840
         TabIndex        =   8
         Top             =   1200
         Width           =   435
      End
   End
   Begin PrjShayan.TypeButton CmdEdit 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   9600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BTYPE           =   1
      TX              =   "ÊÌ—«Ì‘"
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
      BCOL            =   16711680
      BCOLO           =   16777152
      FCOL            =   16777215
      FCOLO           =   33023
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTakhlieKantiner.frx":169B2
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LabelX 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ã” ÃÊÌ ‘„«—Â ò«‰ Ì‰—"
      Height          =   390
      Index           =   0
      Left            =   9240
      TabIndex        =   27
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label LblBaghiPrice 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2010
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   9000
      Width           =   150
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "»«ﬁÌ„«‰œÂ ò—«ÌÂ :"
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   9000
      Width           =   1350
   End
   Begin VB.Label LblTakhlie 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   8880
      Width           =   150
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ  Œ·ÌÂ : "
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   9915
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   9000
      Width           =   1140
   End
   Begin VB.Label LblTotalPrice 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2010
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   8400
      Width           =   150
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ò· ò—«Ì‹Â :"
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   8400
      Width           =   1020
   End
   Begin VB.Label LblTotalTedad 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   8280
      Width           =   150
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ Ê«—œÂ : "
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   9855
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   8400
      Width           =   1200
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   11175
   End
   Begin VB.Menu mnuFind 
      Caption         =   "Find"
      Visible         =   0   'False
      Begin VB.Menu mnuTakhlieRange 
         Caption         =   "„‘«ÂœÂ ò«‰ Ì‰— Â«Ì  Œ·ÌÂ ‘œÂ œ—  «—ÌŒ Œ«’"
         Index           =   0
      End
      Begin VB.Menu ln1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTakhlieToday 
         Caption         =   "„‘«ÂœÂ ò«‰ Ì‰— Â«Ì  Œ·ÌÂ ‘œÂ œ— «„—Ê“"
      End
      Begin VB.Menu ln2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTakhlieDefault 
         Caption         =   "»—ê‘  »Â Õ«· «’·Ì"
      End
   End
End
Attribute VB_Name = "FrmTakhlieKantiner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strDate1 As String
Public strDate2 As String

Private Sub CmdCalcDelay_Click()
  Dim i As Integer
  Dim Today As String
  Dim delay As Integer, Kasr As Currency
 'EDIT
  Dim RowDate As String  ' Each Kantiner Date
  
  LstDelay.Clear
  LstDelayRow.Clear
  'EDIT
  Today = Format(Date, "yyyy/mm/dd")
  'Today = Format(MiladiToShamsi(Format(#8/5/2008#, "yyyy/MM/dd")), "yy/MM/dd")
  'Today = Mid(FrmAhan_Kantiner.TxtKootajDate.Today, 3)
  With Grid1
       For i = 1 To .Rows - 1
           'EDIT
           RowDate = Format(ShamsiToMiladi("13" & .Cell(i, 8).Text), "yyyy/mm/dd")
           If CDate(Today) > CDate(RowDate) + 5 Then ''EDIT
              .Cell(i, 8).BackColor = vbRed
              delay = CDate(Today) - (CDate(RowDate) + 5) ''Edit
              Kasr = 125000 * delay
              LstDelayRow.AddItem i
              LstDelay.AddItem Chr(26) & " ò«‰ Ì‰— —œÌ› " & .Cell(i, 9).Text & " »Â „œ  " & _
                               delay & " —Ê“  «ŒÌ— œ«‘ Â «”  „»·€ " & Kasr & _
                               " —Ì«· «“ ò—«ÌÂ ÊÌ ò”— ŒÊ«Âœ ‘œ."
           End If
       Next
  End With
  ''

  Dim Sum, Total As Currency
       
     LblTakhlie = 0
     For i = 1 To Grid1.Rows - 1
         If Grid1.Cell(i, 4).Text = True Then
            LblTakhlie = Val(LblTakhlie) + 1
            Sum = Sum + CCur(Grid1.Cell(i, 2).Text)
         End If
         Total = Total + CCur(Grid1.Cell(i, 2).Text)
     Next
     '
     LblTotalPrice = Total
     LblBaghiPrice = Total - Sum
    
End Sub

Private Sub CmdEdit_Click()
  If Grid1.ActiveCell.Row > 0 Then
     With FrmSaderatDetail
          .TxtTransitNo = TxtTransitNo
          .TxtKootaj = TxtKootaj
          .TxtRadifMarzi = TxtRadifMarzi
          .TxtPart = TxtPart
          '
          .Caption = "ÊÌ—«Ì‘ «ÿ·«⁄«  ò«‰ Ì‰— ’«œ—« Ì"
          .FrameDetail.Caption = "ÊÌ—«Ì‘ «ÿ·«⁄«  ò«‰ Ì‰— ’«œ—« Ì"
          .CmdOk.Visible = False
          .CmdSave.Visible = True
          .CmdCancel.Visible = False
          .CmdEditCancel.Visible = True
          .TxtKamioon.Visible = False: .TxtSerial.Visible = False
          .TxtEditKamioon.Visible = True
          '
          .LblRadif = Grid1.Cell(Grid1.ActiveCell.Row, 9).Text
          .TxtBarnameDate = Grid1.Cell(Grid1.ActiveCell.Row, 8).Text
          .TxtBarnameNo = Grid1.Cell(Grid1.ActiveCell.Row, 7).Text
          .TxtEditKamioon = Grid1.Cell(Grid1.ActiveCell.Row, 6).Text
          .TxtKantiner = Grid1.Cell(Grid1.ActiveCell.Row, 5).Text
          
          Dim i As Integer
          For i = 0 To .Combsize.ListCount - 1
              If Grid1.Cell(Grid1.ActiveCell.Row, 3).Text = .Combsize.List(i) Then
                 .Combsize.ListIndex = i
                 Exit For
              End If
          Next
          '
          .TxtTotal = Format(Grid1.Cell(Grid1.ActiveCell.Row, 2).Text, "")
          .TxtMobile = Grid1.Cell(Grid1.ActiveCell.Row, 1).Text
          '
           .BackColor = RGB(160, 200, 100)
           .FrameMaster.BackColor = .BackColor
           .FrameDetail.BackColor = .BackColor
           
           .LblRadif.BackColor = RGB(250, 100, 150)
           .LblRadif1.BackColor = RGB(250, 100, 150)
          
          .Show 1
     End With
     
  End If
End Sub

Private Sub CmdPrint_Click()
  Dim i As Integer
  Dim j As Integer

  With GrdVirtual
       .OpenFile App.Path & "\takhlie.cel"
       ''Make Titr
       
       .Cell(2, 4).Text = InputBox("ﬁ«»·  ÊÃÂ :", "", "")
       .Cell(3, 3).Text = InputBox("«“  «—ÌŒ :", "", strDate1)
       .Cell(3, 1).Text = InputBox(" «  «—ÌŒ :", "", strDate2)
       .Cell(5, 7).Text = TxtTransitNo
       .Cell(5, 5).Text = TxtKootaj
       .Cell(5, 2).Text = TxtRadifMarzi
       .Cell(6, 7).Text = TxtPart
       .Cell(6, 5).Text = TxtBarnameDarya
       .Cell(6, 1).Text = TxtTransitDate
       
       'Add item
       For i = 1 To Grid1.Rows - 1 ' Read info from Grid1
           ' + 7 because GRDvirtual start at 8
           .AddItem ""
           For j = 1 To Grid1.Cols - 1 ' For Column reading
               .Cell(i + 7, j).Text = Grid1.Cell(i, j).Text
           Next
           If Grid1.Cell(i, 4).Text = True Then .Cell(i + 7, 4).Text = True
           
       Next
       .Range(8, 4, .Rows - 1, 4).CellType = cellCheckBox
       'make Border
       .Range(8, 1, .Rows - 1, .Cols - 1).Borders(cellInsideHorizontal) = cellThin
       .Range(8, 1, .Rows - 1, .Cols - 1).Borders(cellInsideVertical) = cellThin
       .Range(8, 1, .Rows - 1, .Cols - 1).Borders(cellEdgeRight) = cellThick
       .Range(8, 1, .Rows - 1, .Cols - 1).Borders(cellEdgeLeft) = cellThick
       .Range(8, 1, .Rows - 1, .Cols - 1).Borders(cellEdgeBottom) = cellThick
       'Alignment
       .Range(8, 1, .Rows - 1, .Cols - 1).Alignment = cellCenterCenter
       '
       .PageSetup.FooterFont.Name = "B Nazanin"
       .PageSetup.FooterFont.Size = 12
       .PageSetup.FooterFont.Bold = True
       .PageSetup.FooterMargin = 0.5
       .PageSetup.Footer = "&P"
       .PrintPreview
  End With
  
End Sub

Private Sub Form_Activate()
  strDate1 = Format(MiladiToShamsi(Format(Date, "yyyy/mm/dd")), "yy/mm/dd")
  strDate2 = strDate1
  FrameMaster.Caption = "«ÿ·«⁄«  «’·Ì ò«‰ Ì‰— ’«œ—« Ì" & _
        "    «—ÌŒ «„—Ê“ --" & strDate1 & "--"
  
  '
  Grid1.Range(1, 4, Grid1.Rows - 1, 4).BackColor = RGB(200, 250, 100)
  '
  Grid1.Column(4).Locked = False
  '
  CmdCalcDelay_Click
  '
  Dim i As Long
  Dim Sum As Currency
  
  LblTotalTedad = Grid1.Rows - 1
  For i = 1 To Grid1.Rows - 1
      Sum = Sum + CCur(Grid1.Cell(i, 2).Text)
  Next
  LblTotalPrice = Sum
  '
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
 Dim i As Integer
  With Grid1
       .Cols = 10
       .Rows = 1
       ''
       .DefaultFont.Name = "Traditional Arabic"
       .DefaultFont.Size = 14
       .DefaultFont.Bold = True
       '
       .AllowUserResizing = False
       .MultiSelect = True
       '.ReadOnly = True
       '
       .Appearance = Flat
       .ScrollBarStyle = Flat
       ''
       .BackColorFixed = RGB(100, 150, 200)
       .BackColorFixedSel = RGB(100, 250, 230)
       .BackColorBkg = vbButtonFace 'RGB(90, 158, 214)
       .BackColorScrollBar = RGB(200, 135, 200)
       .BackColor1 = RGB(231, 235, 247)
       .BackColor2 = RGB(239, 243, 255)
       .GridColor = RGB(148, 190, 231)
       ''
       .Cell(0, 1).Text = "‘„«—Â  „«”"
       .Cell(0, 2).Text = "ò—«ÌÂ »Â —Ì«·"
       .Cell(0, 3).Text = "”«Ì‹“"
       .Cell(0, 4).Text = " Œ·ÌÂ"
       .Cell(0, 5).Text = "‘„«—Â ò«‰ Ì‰—"
       .Cell(0, 6).Text = "‘„«—Â ò«„ÌÊ‰"
       .Cell(0, 7).Text = "‘„«—Â »«—‰«„Â"
       .Cell(0, 8).Text = " «—ÌŒ »«—‰«„Â"
       .Cell(0, 9).Text = "—œÌ›"
       '
       For i = 1 To 9
           .Column(i).Alignment = cellCenterCenter
           .Column(i).Locked = True
       Next
       '
       .Column(4).CellType = cellCheckBox 'Takhlie
       '
       .Column(0).Width = 10
       .Column(1).Width = 120 ' Mobile
       .Column(2).Width = 85 ' Total
       .Column(3).Width = 40 ' Size
       .Column(4).Width = 37 ' Takhlie
       .Column(5).Width = 127 ' Kantiner
       .Column(6).Width = 110 ' Kamioon
       .Column(7).Width = 80 ' Barname NO
       .Column(8).Width = 80 ' DAte Barname
       .Column(9).Width = 40 'radif
       ''
  End With
  
End Sub

Private Sub Grid1_CellChange(ByVal Row As Long, ByVal Col As Long)
  Dim i As Integer
  Dim Sum As Currency
  If Col = 4 Then
     LblTakhlie = 0
     For i = 1 To Grid1.Rows - 1
         If Grid1.Cell(i, 4).Text = True Then
            LblTakhlie = Val(LblTakhlie) + 1
            Sum = Sum + CCur(Grid1.Cell(i, 2).Text)
         End If
     Next
     '
     LblBaghiPrice = LblTotalPrice - Sum
  End If
End Sub

Private Sub Grid1_Click()
    If Grid1.ActiveCell.Col = 4 Then
       Dim rsT As New Recordset
       Dim sD As String
       If Grid1.Cell(Grid1.ActiveCell.Row, 4).Text = True Then
          sD = strDate1
       Else ' for change or remove from today takhlie
          sD = Empty
       End If
       '
       rsT.Open "UPDATE TabSaderat_Detail SET Takhlie=" & _
                Grid1.Cell(Grid1.ActiveCell.Row, 4).Text & _
                " , TakhlieDate='" & sD & "' " & _
                "WHERE TransitNo='" & TxtTransitNo & "' AND " & _
                "Count0=" & Val(Grid1.Cell(Grid1.ActiveCell.Row, 9).Text), CNS
      Set rsT = Nothing
    End If
End Sub

Private Sub Timer1_Timer()
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuFind
End Sub

Private Sub LstDelay_DblClick()
  LstDelayRow.ListIndex = LstDelay.ListIndex
  Grid1.Range(Val(LstDelayRow.Text), 1, Val(LstDelayRow.Text), 9).Selected
End Sub

Private Sub mnuFindKantiner_Click()

End Sub

Private Sub mnuTakhlieDefault_Click()
  Call TakhlieDate("", "", True)
  Me.Height = 10965
End Sub

Private Sub mnuTakhlieRange_Click(Index As Integer)
  strDate1 = Format(MiladiToShamsi(Format(Date, "yyyy/mm/dd")), "yy/mm/dd")
  With FrmGetPrintDate
      .CmdOk.Visible = False
      .Caption = ""
      'Send today Date to other Form
      .TxtDate1 = strDate1
      .TxtDate2 = strDate1
      .Show 1
      If .BackColor <> -2147483633 Then
         Call TakhlieDate(strDate1, strDate2)
      End If
  End With

End Sub

Private Sub mnuTakhlieToday_Click()
  strDate1 = Format(MiladiToShamsi(Format(Date, "yyyy/mm/dd")), "yy/mm/dd")
  strDate2 = strDate1
  Call TakhlieDate(strDate1, strDate1)
End Sub

Private Sub TxtFind_Change()
  If Len(TxtFind) = 10 Then
     TxtFind = TxtFind & "-"
     SendKeys "{End}"
  End If

End Sub

Private Sub TxtFind_GotFocus()
SendKeys "{End}"
End Sub

Private Sub TxtFind_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Dim i As Integer
     Dim b As Boolean
     
      '
      b = False
      For i = 1 To Grid1.Rows - 1
          If LCase(Grid1.Cell(i, 5).Text) = LCase(TxtFind) Then
             b = True
             Exit For
          End If
      Next
      '
      If b Then
         Grid1.Cell(i, 4).SetFocus
         Grid1.Cell(i, 4).Text = 1
         Grid1_Click
         Grid1.Range(i, 0, i, 9).Selected
      Else
         MsgBox "«Ì‰ ò«‰ Ì‰— „ÊÃÊœ ‰„Ì »«‘œ", vbExclamation
      End If
      '
      TxtFind.SetFocus
      TxtFind.SelStart = 4
      TxtFind.SelLength = 9
  End If
End Sub

Sub TakhlieDate(Date1 As String, Date2 As String, Optional Default As Boolean = False)
    '  „«„Ì ò«‰ Ì‰— Â«ÌÌ òÂ œ—  «—ÌŒ Œ«’Ì  Œ·ÌÂ ‘œÂ «”  ‰‘«‰ „Ì œÂœ
    'if defualt is TRUE  means rollback to first status
 Dim rs As New Recordset
 Dim i As Integer
 Dim Sum As Currency
 Dim RemainSum  As Currency ' baghi mande
    If Not Default Then
       rs.Open "SELECT * FROM TabSaderat_Detail " & _
                "WHERE (((TabSaderat_Detail.TakhlieDate) BETWEEN '" & Date1 & "' " & _
                "And '" & Date2 & "') " & _
                "AND ((TabSaderat_Detail.TransitNo)='" & TxtTransitNo & "')) " & _
                "ORDER BY Count0 ", CNS
    Else
       rs.Open "SELECT * FROM TabSaderat_Detail " & _
                "WHERE TransitNo='" & TxtTransitNo & "' " & _
                "ORDER BY Count0 ", CNS
      
    End If

    If rs.EOF Then
       MsgBox "ÂÌç ò«‰ Ì‰—Ì «„—Ê“  Œ·ÌÂ ‰‘œÂ «” ", vbExclamation
       rs.Close
       Set rs = Nothing
       Exit Sub
    End If
    '
    With Grid1
         .Rows = 1
         While Not rs.EOF
              .AddItem ""
              For i = 1 To 9
                  If i = 3 Then
                     .Cell(.Rows - 1, i).Text = rs(7) & "ft"
                  ElseIf i = 2 Then
                     .Cell(.Rows - 1, i).Text = Format(rs(8), "#,##0")
                  Else
                     .Cell(.Rows - 1, i).Text = rs(10 - i)
                  End If
                  If Not Default Then .Cell(.Rows - 1, .Cols - 1).Text = .Rows - 1
              Next
              '
              rs.MoveNext
              DoEvents
         Wend
         rs.Close
         Set rs = Nothing
     End With
     '
     'Me.Height = 8730
     '
     LblTotalTedad = Grid1.Rows - 1
     For i = 1 To Grid1.Rows - 1
         Sum = Sum + CCur(Grid1.Cell(i, 2).Text)
         If Grid1.Cell(i, 4).Text = True Then
            RemainSum = RemainSum + CCur(Grid1.Cell(i, 2).Text)
         End If
     Next
     LblTotalPrice = Sum
     LblBaghiPrice = Sum - RemainSum
     '
     Grid1.ReadOnly = Not Default
     Grid1.Range(1, 4, Grid1.Rows - 1, 4).BackColor = RGB(200, 250, 100)
     '
     CmdCalcDelay_Click
End Sub
