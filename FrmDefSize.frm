VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form FrmDefSize 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›—„  ⁄ÌÌ‰ „‘Œ’«  ”«Ì“"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7020
   BeginProperty Font 
      Name            =   "B Zar"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDefSize.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin FlexCell.Grid Grid1 
      Height          =   3495
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6165
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin VB.TextBox TxtSizeTedad 
      Alignment       =   2  'Center
      Height          =   510
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox TxtSizeKala 
      Alignment       =   2  'Center
      Height          =   510
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox TxtParvane 
      Alignment       =   2  'Center
      Height          =   510
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2655
   End
   Begin PrjShayan.TypeButton CmdAdd 
      Height          =   495
      Left            =   4080
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "«÷«›Â ò—œ‰"
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
      MICON           =   "FrmDefSize.frx":29C12
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "»” ‰"
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
      MICON           =   "FrmDefSize.frx":29C2E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdDelete 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "Õ–› «“ «‰ Â«Ì ÃœÊ·"
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
      MICON           =   "FrmDefSize.frx":29C4A
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdFind 
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "Ã” ÃÊÌ ‘„«—Â Å—Ê«‰Â"
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
      MICON           =   "FrmDefSize.frx":29C66
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdEdit 
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "ÊÌ—«Ì‘ ê“Ì‰Â Â«"
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
      MICON           =   "FrmDefSize.frx":29C82
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdReport 
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "„‘«ÂœÂ ê“«—‘ ò· Å—Ê«‰Â"
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
      BCOL            =   8438015
      BCOLO           =   16744703
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDefSize.frx":29C9E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdReportDate 
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "„‘«ÂœÂ ê“«—‘ »Â  ›òÌò  «—ÌŒ"
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
      BCOL            =   8438015
      BCOLO           =   16744703
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDefSize.frx":29CBA
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
      Left            =   4800
      TabIndex        =   14
      Top             =   7440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "ç«Å ÃœÊ· »«·«"
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
      BCOL            =   8438015
      BCOLO           =   16744703
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDefSize.frx":29CD6
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ ò· ”«Ì“"
      Height          =   390
      Left            =   5700
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   1230
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "”‹‹‹«Ì‹“ ò‹«·«"
      Height          =   390
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â Å—Ê«‰Â  "
      Height          =   390
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1170
   End
End
Attribute VB_Name = "FrmDefSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Ahan As Byte ' 1 Ahan 2 Kant 3 AEL
Public ReportDate1 As String, ReportDate2 As String

Private Sub CmdAdd_Click()
 Dim Myrs  As New Recordset
 Dim strSQL As String
 '
 If Trim(TxtParvane) = Empty Or Val(TxtSizeKala) = 0 Or Val(TxtSizeTedad) = 0 Then
    MsgBox "ÌòÌ «“ ò«œ— Â«Ì »«·« Œ«·Ì «” ", vbExclamation, ""
    TxtSizeKala.SetFocus
    Exit Sub
 End If
 '
 If Grid1.Rows > 1 Then
   If TxtParvane <> Grid1.Cell(1, 4).Text Then
      MsgBox "‘„«—Â Å—Ê«‰Â «‘ »«Â «”  «“ ò·Ìœ »” ‰ «” ›«œÂ ò‰Ìœ", vbExclamation, ""
      Exit Sub
   End If
 End If
 '
 Dim radif As Byte
 With Grid1
      radif = Val(.Cell(.Rows - 1, 3).Text) + 1
     .AddItem ""
     .Cell(.Rows - 1, 1).Text = TxtSizeKala
     .Cell(.Rows - 1, 2).Text = TxtSizeTedad
     .Cell(.Rows - 1, 3).Text = radif
     .Cell(.Rows - 1, 4).Text = TxtParvane
     '
     strSQL = "INSERT INTO DefSize (Parvane,Radif,SizeKala,TedadSize,AHAN) "
     strSQL = strSQL & "VALUES('" & TxtParvane & "',"
     strSQL = strSQL & radif & "," & Val(TxtSizeKala) & ","
     strSQL = strSQL & Val(TxtSizeTedad) & "," & Ahan & ") "
     
     Myrs.Open strSQL, CNS
     Set Myrs = Nothing
     '
     TxtSizeKala.Text = Empty
     TxtSizeTedad.Text = Empty
     TxtSizeKala.SetFocus
     
 End With
End Sub

Private Sub CmdClose_Click()
 TxtSizeKala = Empty
 TxtSizeTedad = Empty
 
 If TxtParvane.Enabled = False And CmdClose.Caption = "»” ‰" Then
    TxtParvane.Enabled = True
    TxtParvane.SetFocus
    SendKeys "{home}+{end}"
    Grid1.Rows = 1
 ElseIf CmdClose.Caption = "«‰’—«›" Then
    CmdClose.Caption = "»” ‰"
    '
    Grid1.Enabled = True
    CmdAdd.Enabled = True
    CmdDelete.Enabled = True
    CmdFind.Enabled = True
    '
    CmdEdit.Caption = "ÊÌ—«Ì‘ ê“Ì‰Â Â«"
    '
 Else
    Unload Me
 End If
End Sub

Private Sub CmdDelete_Click()
 Dim Myrs As New Recordset
 Dim strSQL As String
 '
 If Grid1.Rows > 1 Then
    strSQL = "DELETE FROM DefSize "
    strSQL = strSQL & "WHERE Parvane='" & Grid1.Cell(Grid1.Rows - 1, 4).Text & "' "
    strSQL = strSQL & "AND Radif=" & Val(Grid1.Cell(Grid1.Rows - 1, 3).Text)
    strSQL = strSQL & " AND AHAN=" & Ahan
    '
    Myrs.Open strSQL, CNS
    Set rs = Nothing
    '
    Grid1.RemoveItem Grid1.Rows - 1
 End If
End Sub

Private Sub CmdEdit_Click()
 With Grid1
    If CmdEdit.Caption = "ÊÌ—«Ì‘ ê“Ì‰Â Â«" Then
       If .ActiveCell.Row > 0 Then
          TxtParvane = .Cell(.ActiveCell.Row, 4).Text
          TxtSizeKala = .Cell(.ActiveCell.Row, 1).Text
          TxtSizeTedad = .Cell(.ActiveCell.Row, 2).Text
          TxtSizeKala.SetFocus
          CmdClose.Caption = "«‰’—«›"
          '
          .Enabled = False
          CmdAdd.Enabled = False
          CmdDelete.Enabled = False
          CmdFind.Enabled = False
          '
          CmdEdit.Caption = "À»   €ÌÌ—« "
       End If
    ElseIf CmdEdit.Caption = "À»   €ÌÌ—« " Then
       Dim Myrs As New Recordset
       Dim strSQL As String
       '
       strSQL = "UPDATE DefSize SET "
       strSQL = strSQL & "SizeKala=" & Val(TxtSizeKala) & ", "
       strSQL = strSQL & "TedadSize=" & Val(TxtSizeTedad) & " "
       strSQL = strSQL & "WHERE Parvane='" & TxtParvane & "' "
       strSQL = strSQL & "AND Radif=" & Val(.Cell(.ActiveCell.Row, 3).Text) & " "
       strSQL = strSQL & "AND AHAN=" & Ahan
       '
       Myrs.Open strSQL, CNS
       Set Myrs = Nothing
       ''
       CmdClose.Caption = "»” ‰"
       '
       .Enabled = True
       CmdAdd.Enabled = True
       CmdDelete.Enabled = True
       CmdFind.Enabled = True
       '
       CmdEdit.Caption = "ÊÌ—«Ì‘ ê“Ì‰Â Â«"
       '
       .Cell(.ActiveCell.Row, 1).Text = TxtSizeKala
       .Cell(.ActiveCell.Row, 2).Text = TxtSizeTedad
       '
       TxtSizeKala = Empty
       TxtSizeTedad = Empty
    End If
 End With
End Sub

Private Sub CmdFind_Click()
 Dim inp As String
 
 inp = InputBox("‘„«—Â Å—Ê«‰Â —« Ê«—œ ò‰Ìœ", "Ã” ÃÊÌ Å—Ê«‰Â", "")
 If Trim(inp) <> Empty Then
    Dim Myrs As New Recordset
    Dim strSQL As String
    '
    strSQL = "SELECT * FROM DefSize "
    strSQL = strSQL & "WHERE Parvane='" & inp & "' "
    strSQL = strSQL & "AND AHAN=" & Ahan
    '
    Myrs.Open strSQL, CNS
    If Not Myrs.EOF Then
       LoadGrid Myrs
       TxtParvane = inp
       Myrs.Close
       Set Myrs = Nothing
    Else
       MsgBox "‘„«—Â Å—Ê«‰Â „Ê—œ ‰Ÿ— ÅÌœ« ‰‘œ", vbExclamation, ""
    End If
 End If
End Sub

Private Sub CmdPrint_Click()
 Dim TableName As String
 Dim mainTab As String
 Dim i As Integer
 
 If Grid1.Rows = 1 Then Exit Sub
 
 If Ahan = 1 Then
    TableName = "TabAhan_Detail"
    mainTab = "TabAhan_Master"
 ElseIf Ahan = 2 Then
    TableName = "TabKantiner_Detail"
    mainTab = "TabKantiner_Master"
 ElseIf Ahan = 3 Then
    TableName = "TabAEL_Detail"
    mainTab = "TabAEL_Master"
 End If
 
 Call MakeMasterRows
 Call MakeBorder
 If LoadDataInMasterRows(mainTab) Then
    With FrmRepSize.Grid1
         .RemoveItem 10
         .RemoveItem 10
         '
         .AddItem ""
         '
         .Range(.Rows - 1, 2, .Rows - 1, 3).Merge
         .Cell(.Rows - 1, 2).Text = "”«Ì“ ò«·«"
         .Range(.Rows - 1, 4, .Rows - 1, 6).Merge
         .Cell(.Rows - 1, 4).Text = " ⁄œ«œ ò· ”«Ì“ "
         .Range(.Rows - 1, 7, .Rows - 1, 8).Merge
         .Cell(.Rows - 1, 7).Text = "—œÌ› "
         .Range(.Rows - 1, 9, .Rows - 1, 11).Merge
         .Cell(.Rows - 1, 9).Text = "‘„«—Â Å—Ê«‰Â "
         '
         For i = 1 To Grid1.Rows - 1
             .AddItem ""
             .Range(.Rows - 1, 2, .Rows - 1, 3).Merge
             .Cell(.Rows - 1, 2).Text = Grid1.Cell(i, 1).Text
             .Range(.Rows - 1, 4, .Rows - 1, 6).Merge
             .Cell(.Rows - 1, 4).Text = Grid1.Cell(i, 2).Text
             .Range(.Rows - 1, 7, .Rows - 1, 8).Merge
             .Cell(.Rows - 1, 7).Text = Grid1.Cell(i, 3).Text
             .Range(.Rows - 1, 9, .Rows - 1, 11).Merge
             .Cell(.Rows - 1, 9).Text = Grid1.Cell(i, 4).Text
         Next
         .Range(10, 1, .Rows - 1, .Cols - 1).FontName = "B Titr"
         .Range(10, 1, .Rows - 1, .Cols - 1).FontSize = 12
         .Range(10, 1, .Rows - 1, .Cols - 1).Alignment = cellCenterCenter
         
         .Range(10, 2, .Rows - 1, 9).Borders(cellEdgeBottom) = cellThick
         .Range(10, 2, .Rows - 1, 9).Borders(cellEdgeTop) = cellThick
         .Range(10, 2, .Rows - 1, 9).Borders(cellEdgeLeft) = cellThick
         .Range(10, 2, .Rows - 1, 9).Borders(cellEdgeRight) = cellThick
         .Range(10, 2, .Rows - 1, 9).Borders(cellInsideHorizontal) = cellThin
         .Range(10, 2, .Rows - 1, 9).Borders(cellInsideVertical) = cellThin
    End With
    FrmRepSize.EnableDesc = False
    FrmRepSize.Show
 End If
End Sub

Private Sub CmdReport_Click()
 Dim TableName As String
 Dim mainTab As String
 
 If Grid1.Rows = 1 Then Exit Sub
 
 If Ahan = 1 Then
    TableName = "TabAhan_Detail"
    mainTab = "TabAhan_Master"
 ElseIf Ahan = 2 Then
    TableName = "TabKantiner_Detail"
    mainTab = "TabKantiner_Master"
 ElseIf Ahan = 3 Then
    TableName = "TabAEL_Detail"
    mainTab = "TabAEL_Master"
 End If
 '
 
 '
 Call MakeMasterRows
 Call MakeBorder
 If LoadDataInMasterRows(mainTab) Then
    Call MakeReport(TableName, mainTab, ReportDate1, ReportDate2)
    Call MakeDetailRows
    FrmRepSize.Ahan = Ahan
    FrmRepSize.EnableDesc = True
    FrmRepSize.Show
 End If
 '
 
End Sub


Private Sub CmdReportDate_Click()
 FrmGetPrintDate.WichF = 4
 FrmGetPrintDate.ParCode = TxtParvane.Text
 FrmGetPrintDate.Caption = "  Å—Ì‰  «“ „‘Œ’«  ”«Ì“  " & TxtParvane
 FrmGetPrintDate.Show 1, FrmDefSize
 
 CmdReport_Click
End Sub

Private Sub Form_Load()
 ClearField
 SetGrid
End Sub

Sub SetGrid()
 Dim i As Integer
 With Grid1
      .Cols = 5
      .Rows = 1
      '
      .DefaultFont.Name = "B Nazanin"
      .DefaultFont.Bold = True
      .DefaultFont.Size = 12
      
      .DefaultRowHeight = 27
      .AllowUserResizing = False
      '
      '.BackColor2 = vbWhite
      .Column(0).Width = 0
      .Column(1).Width = 105
      .Column(2).Width = 105
      .Column(3).Width = 100
      .Column(4).Width = 120
      '
      For i = 1 To .Cols - 1
          .Column(i).Alignment = cellCenterCenter
      Next
      '
      'Make Titr
      .Cell(0, 1).Text = "”«Ì“ ò«·«"
      .Cell(0, 2).Text = " ⁄œ«œ ò· ”«Ì“"
      .Cell(0, 3).Text = "—œÌ›"
      .Cell(0, 4).Text = "‘„«—Â Å—Ê«‰Â"
      '
      .ReadOnly = True
 End With
End Sub

Sub ClearField()
   TxtParvane = Empty
   TxtSizeKala = Empty
   TxtSizeTedad = Empty
End Sub

Private Sub TxtParvane_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"

End Sub

Private Sub TxtParvane_LostFocus()
 If Trim(TxtParvane) <> Empty Then
    Dim Myrs As New Recordset
    Dim strSQL As String
    
    strSQL = "SELECT * FROM DefSize "
    strSQL = strSQL & "WHERE Parvane='" & TxtParvane & "' "
    strSQL = strSQL & "AND AHAN=" & Ahan
    
    Myrs.Open strSQL, CNS
    If Not Myrs.EOF Then LoadGrid Myrs
       
    Myrs.Close
    Set Myrs = Nothing
    '
    TxtParvane.Enabled = False
 End If
End Sub

Sub LoadGrid(F As Recordset)
  With Grid1
      .Rows = 1
      Do While Not F.EOF
        .AddItem ""
        .Cell(.Rows - 1, 1).Text = IIf(IsNull(F("SizeKala")), 0, F("SizeKala"))
        .Cell(.Rows - 1, 2).Text = IIf(IsNull(F("TedadSize")), 0, F("TedadSize"))
        .Cell(.Rows - 1, 3).Text = F("Radif")
        .Cell(.Rows - 1, 4).Text = F("Parvane")
        F.MoveNext
      Loop
  End With
End Sub

Private Sub TxtSizeKala_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"
'
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Private Sub TxtSizeTedad_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"
'
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Sub MakeMasterRows()
  Dim i As Integer
  
   With FrmRepSize.Grid1
       .Cols = 12
       .Rows = 12
       .DefaultRowHeight = 22
'''''''''''''''''''''''''''''''''''''
       .Column(1).Width = 80 'Mobile
       .Column(2).Width = 80 'Total
       .Column(3).Width = 40 'size
       .Column(4).Width = 40 'Shakhe
       .Column(5).Width = 30 'tedad
       .Column(6).Width = 50 'Weight
       .Column(7).Width = 105 'Kamioon
       .Column(8).Width = 90 'Anbar
       .Column(9).Width = 60 'Date
       .Column(10).Width = 55 'Barname
       .Column(11).Width = 35 'Radif
'''''''''''''''''''''''''''''''''''''''''
       .RowHeight(0) = 0
       .RowHeight(11) = 60
       .Column(0).Width = 0
       ''''''''''
       For i = 1 To 8
           .RowHeight(i) = 20
       Next
       .RowHeight(3) = 30
       .RowHeight(9) = 18
       ''''''''''''''''
       .Cell(11, 5).WrapText = True
       .Cell(11, 2).WrapText = True
       '''''''''''''''''''
       '''''
       .Range(1, 3, 3, 6).Merge 'keshti
       .Range(1, 1, 3, 2).Merge 'Barname
       .Range(1, 7, 2, 11).Merge 'List
       .Range(3, 7, 3, 11).Merge 'SArvar
       .Range(10, 10, 10, 11).Merge 'SAHEB KALA(Titr)
       .Range(10, 8, 10, 9).Merge 'SAHEB KALA(Value)
       '.Range(10, 6, 10, 7).Merge 'Etebar(Titr)
       .Range(10, 4, 10, 6).Merge 'Etebar(Value)
       .Range(10, 2, 10, 3).Merge 'Gharardad(Titr)
       
       '
       For i = 4 To 8
           .Range(i, 1, i, 2).Merge
       Next
       '
       For i = 4 To 8
           .Range(i, 7, i, 8).Merge
       Next
       '
       For i = 4 To 8
           .Range(i, 3, i, 6).Merge
       Next
       '
       For i = 4 To 8
           .Range(i, 9, i, 11).Merge
       Next
       
       ''''''''''''''
       .Range(1, 1, 10, 7).FontName = "B Zar"
       .Range(1, 1, 10, 7).FontBold = True
       .Range(1, 1, 10, 7).FontSize = 12
       ''''''''''''''List e bargiri
       .Range(1, 4, 3, 7).FontName = "B Zar"
       .Range(1, 4, 3, 7).FontBold = True
       .Range(1, 4, 3, 7).FontSize = 12
       '
       .Range(1, 7, 2, 11).FontName = "B Zar"
       .Range(1, 7, 2, 11).FontBold = True
       .Range(1, 7, 2, 11).FontSize = 12
       
       ''''''''''''''
       .Range(11, 1, 11, 11).FontName = "B Titr"
       .Range(11, 1, 11, 11).FontBold = True
       .Range(11, 1, 11, 11).FontSize = 8
       '
       .Range(7, 7, 7, 8).FontName = "B Zar"
       .Range(7, 7, 7, 8).FontBold = True
       .Range(7, 7, 7, 8).FontSize = 12
       
       '''''''''''''
       .Cell(1, 1).Text = ":»«—‰«„Â"
       .Cell(1, 1).Alignment = cellRightCenter
       .Cell(1, 1).Font.Name = "Traditional Arabic"
       .Cell(1, 3).Text = ":ò‘ Ì"
       .Cell(1, 4).Alignment = cellRightCenter
       '''
       .Cell(4, 3).Text = "”«Ì“ ò‹‹«·«"
       .Cell(5, 3).Text = "Ê“‰ ‰‹«Œ‹‹«·’"
       .Cell(6, 3).Text = "Ê“‰ Œ‹‹‹‹«·’"
       .Cell(7, 3).Text = " ⁄œ«œ »‰œ· Ê ò· ‘«ŒÂ"
       .Cell(8, 3).Text = "‰«„  —ŒÌ’ ò«—"
       ''''''''''''''''
       .Cell(4, 9).Text = "‰‹‹Ê⁄ ò‹‹‹«·«"
       .Cell(5, 9).Text = "‘„‹‹«—Â Å‹—Ê«‰Â"
       .Cell(6, 9).Text = "‘„«—Â ò‹Ê ‹«é"
       .Cell(7, 9).Text = "‘„«—Â ﬁ»÷ «‰»«—"
       .Cell(8, 9).Text = " ‹«—ÌŒ Å‹‹—Ê«‰Â"
       '
       .Cell(11, 11).Text = "—œÌ›"
       .Range(11, 9, 11, 10).Merge 'Size
       .Cell(11, 9).Text = "”«Ì“ ò«·«"
       
       .Cell(11, 8).Text = " «—ÌŒ "
       .Cell(11, 7).Text = " ⁄œ«œ —› Â"
       .Range(11, 5, 11, 6).Merge 'Baghimande
       .Cell(11, 5).Text = " ⁄œ«œ »«ﬁÌ „«‰œÂ"
       
       .Range(11, 3, 11, 4).Merge 'Tedad Kol
       .Cell(11, 3).Text = " ⁄œ«œ ò·"
       
       .Range(11, 1, 11, 2).Merge  'Nothing
       .Cell(11, 1).Text = "„‹‹·‹«Õ‹‹Ÿ‹« "
       '.Cell(11, 4).Text = "”«Ì“"
       '.Cell(11, 3).Text = "‘«ŒÂ"
       '.Cell(11, 2).Text = "ò· ò—«ÌÂ"
       '.Cell(11, 1).Text = "„Ê»«Ì·"

       '
       
       .Range(4, 1, 8, 11).Alignment = cellRightCenter ' Size,No,Par,...
       .Range(7, 1, 10, 11).Alignment = cellRightCenter 'Saheb,Etebar
       .Range(11, 1, 11, 11).Alignment = cellCenterCenter 'Titr
       '.Range(12, 1, 11, 12).Alignment = cellRightCenter
       '
       
       
       .Cell(1, 3).Alignment = cellRightCenter 'BArname
       .Cell(1, 7).Alignment = cellRightCenter 'List
       .Cell(3, 7).Alignment = cellRightCenter 'Sarvar
       .Cell(10, 7).Alignment = cellRightCenter 'saheb KALA

       '
       If Ahan = 1 Or Ahan = 2 Then
          .Cell(1, 7).Text = "·Ì”  »«—êÌ—Ì ‘—ò  Õ„· Ê ‰ﬁ· „Â—Ê—“«‰  —«»—" 'List
       Else
          .Cell(1, 7).Text = "‘‹‹‹‹‹—ò‹‹  ”‹‹—Ê‘  ‹—Œ‹Ì‹‹’ Å‹‹«—”‹Â" ' List
          .Range(1, 7, 2, 11).FontSize = 14
       End If
       '
       Dim inp As String
        inp = InputBox("", ":ﬁ«»·  ÊÃÂ ”—Ê— ê—«„Ì ")
        If Trim(inp) <> Empty Then
           .Cell(3, 7).Text = " ﬁ«»·  ÊÃÂ ”—Ê— ê—«„Ì: " & inp 'Sarvar
        Else
           .Cell(3, 7).Text = ": ﬁ«»·  ÊÃÂ ”—Ê— ê—«„Ì" 'Sarvar
        End If
       ''''''''''''''''''''''''
       .Cell(10, 10).Text = "’«Õ» ò«·«" 'Saheb KALA
       .Cell(10, 7).Text = "‘„«—Â «⁄ »«—" 'Etebar
       .Cell(10, 2).Text = "‘„«—Â ﬁ—«—œ«œ" 'Etebar
       ''''
       .Cell(10, 4).Font.Size = 10 'Etebar
       .Cell(10, 4).Alignment = cellCenterCenter 'Etebar
       .Cell(10, 1).Font.Size = 10 'Gharardad
       .Cell(10, 1).Alignment = cellCenterCenter 'gharadad
       '
       .Range(1, 1, .Rows - 1, .Cols - 1).FontName = "B Zar"
       .Range(1, 1, .Rows - 1, .Cols - 1).FontSize = 12
       .Range(1, 1, .Rows - 1, .Cols - 1).FontBold = True
   End With

End Sub

Private Sub MakeBorder()
   With FrmRepSize.Grid1
       .Range(1, 1, 3, 6).Selected
       .Selection.Borders(cellEdgeBottom) = cellThick
       .Selection.Borders(cellEdgeTop) = cellThick
       .Selection.Borders(cellEdgeRight) = cellThick
       .Selection.Borders(cellEdgeLeft) = cellThick
       '
       .Range(4, 1, 8, 2).Selected
       .Selection.Borders(cellEdgeLeft) = cellThick
       '
       .Range(1, 1, 3, 2).Selected
       .Selection.Borders(cellEdgeRight) = cellThick
       
       '
       .Range(1, 1, 8, 4).Selected
       .Selection.Borders(cellEdgeRight) = cellThick
       '
      
       .Range(1, 1, 8, 11).Selected
       .Selection.Borders(cellEdgeBottom) = cellThick
       '
       .Range(4, 7, 8, 11).Selected
       .Selection.Borders(cellEdgeTop) = cellThick
       .Selection.Borders(cellEdgeRight) = cellThick
       '
       .Range(1, 7, 2, 11).Selected
       .Selection.Borders(cellEdgeTop) = cellNone
       .Selection.Borders(cellEdgeBottom) = cellNone
       '
       .Range(4, 7, 8, 11).Selected
       .Selection.Borders(cellInsideHorizontal) = cellThin
       .Selection.Borders(cellInsideVertical) = cellThick
       '
       .Range(4, 1, 8, 3).Selected
       .Selection.Borders(cellInsideHorizontal) = cellThin
       .Selection.Borders(cellInsideVertical) = cellThick
       '
       .Range(1, 1, 3, 3).Selected
       .Selection.Borders(cellEdgeBottom) = cellThick
       '
       .Range(10, 10, 10, 11).Selected
       .Selection.Borders(cellEdgeLeft) = cellThick
       '
       .Range(10, 6, 10, 7).Selected
       .Selection.Borders(cellEdgeRight) = cellThick
       '
       .Range(10, 1, 10, 1).Selected
       .Selection.Borders(cellEdgeRight) = cellThick
       
       '
       .Range(10, 1, 10, 11).Selected
       .Selection.Borders(cellEdgeTop) = cellThick
       .Selection.Borders(cellEdgeBottom) = cellThick
       '
       .Range(10, 1, 10, 2).Selected
       .Selection.Borders(cellEdgeLeft) = cellThick
       .Selection.Borders(cellEdgeRight) = cellThick
       '
       .Range(10, 3, 10, 4).Selected
       .Selection.Borders(cellEdgeRight) = cellThick
       '
       .Range(10, 5, 10, 11).Selected
       .Selection.Borders(cellEdgeRight) = cellThick
       '
       .Range(11, 1, 11, 11).Selected
       .Selection.Borders(cellEdgeLeft) = cellThick
       .Selection.Borders(cellEdgeRight) = cellThick
       .Selection.Borders(cellEdgeBottom) = cellThick
       .Selection.Borders(cellInsideHorizontal) = cellThin
       .Selection.Borders(cellInsideVertical) = cellThick
       '''
       .Range(10, 10, 10, 11).Selected
       .Selection.Borders(cellEdgeLeft) = cellNone
       '
       .Range(10, 4, 10, 5).Selected
       .Selection.Borders(cellEdgeRight) = cellNone
       '
       .Range(10, 1, 10, 1).Selected
       .Selection.Borders(cellEdgeRight) = cellNone
       
   End With

End Sub


Sub MakeReport(TabName As String, mainTab As String, D1 As String, D2 As String)
 Dim Myrs As New Recordset
 Dim strSQL As String
 '
 ''add date to grid
 Dim Tedad As Long
 Dim SumTedad As Long
 Dim i As Long, j As Long
 Dim MergeCounter As Integer
 
 MergeCounter = 12 ' First Detail Row
 j = 1
     
  '''QUERY 1
  strSQL = "SELECT DefSize.Parvane, " & TabName & ".DBarname, "
  strSQL = strSQL & TabName & ".Size0, SUM(" & TabName & ".Tedad) AS SumOfTedad "
  strSQL = strSQL & "FROM DefSize INNER JOIN " & TabName & " ON "
  strSQL = strSQL & "(DefSize.SizeKala = " & TabName & ".Size0) AND "
  strSQL = strSQL & "(DefSize.Parvane = " & TabName & ".Parvane) "
  strSQL = strSQL & "GROUP BY DefSize.Parvane, " & TabName & ".DBarname, "
  strSQL = strSQL & TabName & ".Size0 "
  strSQL = strSQL & "HAVING (((DefSize.Parvane)='" & TxtParvane & "')) "
  strSQL = strSQL & "ORDER BY " & TabName & ".Size0"
  
  '
  Myrs.Open strSQL, CNS
  With FrmRepSize.Grid1
     Do While Not Myrs.EOF
        ' Find each size 's Tedad Kol
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.Cell(i, 1).Text) = Val(Myrs("Size0")) Then
               Tedad = Val(Grid1.Cell(i, 2).Text)
               Exit For
            End If
        Next
        '
        If Val(Myrs("Size0")) <> Val(.Cell(.Rows - 1, 9).Text) Then
           SumTedad = 0
           On Error Resume Next
           .Range(MergeCounter, 1, .Rows - 1, 2).Merge
           .Range(MergeCounter, 3, .Rows - 1, 3).Merge
           .Cell(MergeCounter, 1).WrapText = True
           .Range(.Rows - 1, 1, .Rows - 1, .Cols - 1).Borders(cellEdgeBottom) = cellThick
           
           MergeCounter = .Rows
        End If
        '
        If D1 <> Empty And D2 <> Empty Then ''
           If Myrs("DBarname") >= D1 And Myrs("DBarname") <= D2 Then
             .AddItem ""
             .Cell(.Rows - 1, 11).Text = j
             .Range(.Rows - 1, 9, .Rows - 1, 10).Merge 'Size
             .Cell(.Rows - 1, 9).Text = Val(Myrs("Size0"))
             .Cell(.Rows - 1, 8).Text = Myrs("DBarname")
             .Cell(.Rows - 1, 7).Text = Myrs("SumOfTedad")
             .Range(.Rows - 1, 5, .Rows - 1, 6).Merge 'Baghimande
             SumTedad = SumTedad + Myrs("SumOfTedad")
             .Cell(.Rows - 1, 5).Text = Tedad - SumTedad
             .Range(.Rows - 1, 3, .Rows - 1, 4).Merge 'Tedad Kol
             .Cell(.Rows - 1, 3).Text = Tedad
          End If
        ElseIf D1 = Empty And D2 = Empty Then ''
             .AddItem ""
             .Cell(.Rows - 1, 11).Text = j
             .Range(.Rows - 1, 9, .Rows - 1, 10).Merge 'Size
             .Cell(.Rows - 1, 9).Text = Val(Myrs("Size0"))
             .Cell(.Rows - 1, 8).Text = Myrs("DBarname")
             .Cell(.Rows - 1, 7).Text = Myrs("SumOfTedad")
             .Range(.Rows - 1, 5, .Rows - 1, 6).Merge 'Baghimande
             SumTedad = SumTedad + Myrs("SumOfTedad")
             .Cell(.Rows - 1, 5).Text = Tedad - SumTedad
             .Range(.Rows - 1, 3, .Rows - 1, 4).Merge 'Tedad Kol
             .Cell(.Rows - 1, 3).Text = Tedad
        End If ''
       
       Myrs.MoveNext
       j = j + 1
     Loop
     
     If .Rows <= 12 Then Exit Sub
     
     .Range(MergeCounter, 1, .Rows - 1, 2).Merge
     .Range(MergeCounter, 3, .Rows - 1, 3).Merge
     .Cell(MergeCounter, 1).WrapText = True
     .Range(.Rows - 1, 1, .Rows - 1, .Cols - 1).Borders(cellEdgeBottom) = cellThick
      Myrs.Close
      '''''''''''''''''''''''''''''''
     '''Load Description
     strSQL = "SELECT * FROM DefSize "
     strSQL = strSQL & "WHERE Parvane='" & TxtParvane & "' AND "
     strSQL = strSQL & "AHAN =" & Ahan & " AND "
     strSQL = strSQL & "Description IS NOT NULL "
     Myrs.Open strSQL, CNS
     Do While Not Myrs.EOF
        For i = 12 To .Rows - 1
            If Trim(.Cell(i, 9).Text) = Trim(Myrs("SizeKala")) Then
               .Cell(i, 1).Text = Myrs("Description")
               Exit For
            End If
        Next
        Myrs.MoveNext
     Loop
     '''
     Myrs.Close
     ''''''''''''''''''''''''''''
      ''''Tonaj Baghimande
      Dim SumWeight As Long
      Dim NakhalesWeight As Long
      
      Myrs.Open "SELECT SUM(Weight) FROM " & TabName & _
                " WHERE Parvane='" & TxtParvane & "'", CNS
      SumWeight = Myrs(0)
      Myrs.Close
      '--
      Myrs.Open "SELECT NWeight FROM " & mainTab & _
                " WHERE Parvane='" & TxtParvane & "'", CNS
      NakhalesWeight = Myrs(0)
      Myrs.Close
      '--
      .AddItem ""
      .Range(.Rows - 1, 1, .Rows - 1, 8).Merge
      .Range(.Rows - 1, 9, .Rows - 1, 11).Merge
      .Cell(.Rows - 1, 9).Text = " ‰«é »«ﬁÌ„«‰œÂ"
      .Cell(.Rows - 1, 1).Text = NakhalesWeight - SumWeight
      ''''''''''
     Set Myrs = Nothing
  
  End With
End Sub

Sub MakeDetailRows()
On Error Resume Next
    With FrmRepSize.Grid1
         .Range(12, 1, .Rows - 1, .Cols - 1).Alignment = cellCenterCenter
         .Range(12, 1, .Rows - 1, .Cols - 1).FontName = "B Nazanin"
         .Range(12, 1, .Rows - 1, .Cols - 1).FontSize = 12
         .Range(12, 1, .Rows - 1, .Cols - 1).FontBold = True
         '
         .Range(12, 1, .Rows - 1, .Cols - 1).Borders(cellEdgeLeft) = cellThick
         .Range(12, 1, .Rows - 1, .Cols - 1).Borders(cellEdgeRight) = cellThick
         .Range(12, 1, .Rows - 1, .Cols - 1).Borders(cellEdgeBottom) = cellThick
         '
         .Range(12, 1, .Rows - 1, .Cols - 1).Borders(cellInsideVertical) = cellThick
    End With
End Sub

Function LoadDataInMasterRows(TabName As String) As Boolean
Dim Myrs As New Recordset
Dim ks As String
'''''''''Load Data In Master Rows
LoadDataInMasterRows = True

  Myrs.Open "SELECT * FROM " & TabName & " " & _
          "WHERE Parvane='" & TxtParvane & "'", CNS
  
  If Myrs.EOF Then
     MsgBox "ç‰Ì‰ Å—Ê«‰Â «Ì „ÊÃÊœ ‰Ì” ", vbExclamation, ""
     Myrs.Close
     LoadDataInMasterRows = False
     Exit Function
  End If
  
  With FrmRepSize.Grid1
    .Cell(1, 3).Text = "ò‘ Ì :" & Space(1) & Trim(Myrs("Keshti"))
    .Cell(1, 1).Text = "»«—‰«„Â :" & Space(0) & Trim(Myrs("Barname"))
    .Cell(1, 1).Font.Name = "B Zar"
    .Cell(6, 7).Font.Size = 12
    
    .Cell(4, 7).Text = Trim(Myrs("Typekala"))
    .Cell(5, 7).Text = Trim(Myrs("Parvane"))
    .Cell(6, 7).Text = Trim(Myrs("DKootaj")) & Space(13 - Len(Trim(Myrs("Kootaj")))) & Trim(Myrs("Kootaj"))
    .Cell(6, 7).Font.Name = "B Zar"
    .Cell(6, 7).Font.Size = 12
    '
    .Cell(7, 7).Text = Trim(Myrs("DGhabz")) & Space(17 - Len(Trim(Myrs("Ghabz")))) & Trim(Myrs("Ghabz"))
    .Cell(8, 7).Text = Trim(Myrs("Dparvane"))
    .Cell(7, 7).Font.Name = "B Zar"
    .Cell(7, 7).Font.Size = 12
    
    ks = Trim(Myrs("Sizekala"))
    ks = Replace(ks, "x", Chr$(215))
    .Cell(4, 1).Text = ks
    .Cell(4, 1).Font.Name = "B Zar"
    .Cell(4, 1).Font.Size = 12
    
    .Cell(5, 1).Text = Trim(Myrs("NWeight"))
    .Cell(6, 1).Text = Trim(Myrs("Weight"))
    .Cell(7, 1).Text = Val(Myrs("Shakhe")) & Space(17) & Val(Myrs("Bandel"))
    .Cell(8, 1).Text = Trim(Myrs("Tarkhiskar"))
    
    '
    .Cell(10, 8).Text = IIf(IsNull(Myrs("Saheb")), "", Trim(Myrs("Saheb")))
    .Cell(10, 4).Text = IIf(IsNull(Myrs("Etebar")), "", Trim(Myrs("Etebar")))
    .Cell(10, 1).Text = IIf(IsNull(Myrs("Gharardad")), "", Trim(Myrs("Gharardad")))
  End With
  Myrs.Close
  
  Set Myrs = Nothing
End Function
