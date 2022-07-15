VERSION 5.00
Object = "{9DBDC544-49CA-11D7-B1ED-C2237039C523}#1.1#0"; "FarDate.Ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form FrmEditDetailA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«’·«Õ «ÿ·«⁄«  Ã«‰»Ì ¬Â‰ ¬·« "
   ClientHeight    =   10095
   ClientLeft      =   4815
   ClientTop       =   1230
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "B Zar"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmEditDetailA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin FlexCell.Grid Grid2 
      Height          =   495
      Left            =   3360
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin FlexCell.Grid Grid1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   13150
      Cols            =   12
      DefaultFontName =   "B Zar"
      DefaultFontSize =   12
      DefaultFontBold =   -1  'True
      DefaultRowHeight=   32
      GridColor       =   -2147483630
      GridLiness      =   -1  'True
      Rows            =   15
   End
   Begin FarDate1.FarDate TxtDate 
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   480
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
   End
   Begin PrjShayan.TypeButton CmdOK 
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   480
      Width           =   615
      _ExtentX        =   1085
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
      MICON           =   "FrmEditDetailA.frx":169B2
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
      Left            =   5280
      TabIndex        =   12
      Top             =   9120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "ç«Å ÃœÊ·"
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
      MICON           =   "FrmEditDetailA.frx":169CE
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ·›‰  „«”   09112320258 "
      BeginProperty Font 
         Name            =   "B Jadid"
         Size            =   27.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   3300
      TabIndex        =   11
      Top             =   5400
      Width           =   5565
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„Õ„œ Õ”‰ ¬—ÊÌ‰ ›—"
      BeginProperty Font 
         Name            =   "B Jadid"
         Size            =   27.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   2400
      TabIndex        =   10
      Top             =   3960
      Width           =   5415
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "»«  ‘ò— «“ Õ”‰ «‰ Œ«» ‘„«"
      BeginProperty Font 
         Name            =   "B Jadid"
         Size            =   27.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   2760
      TabIndex        =   9
      Top             =   2640
      Width           =   5745
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ê Ì« »——ÊÌ œò„Â «Ì òÂ œ— ” Ê‰ „Ê»«Ì· ÊÃÊœ œ«—œ ò·Ìò ò‰Ìœ"
      BeginProperty Font 
         Name            =   "Traditional Arabic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   5910
      TabIndex        =   6
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " —« »›‘«—ÌœF4    »—«Ì –ŒÌ—Â ò—œ‰ „Ê«—œ  €ÌÌ— œ«œÂ ‘œÂ  Å” «“  ò„Ì· Â— ”ÿ— ò·Ìœ"
      BeginProperty Font 
         Name            =   "Traditional Arabic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   3975
      TabIndex        =   5
      Top             =   0
      Width           =   7200
   End
   Begin VB.Label LblWeight 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " :Ã„⁄ Ê“‰"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   9360
      Width           =   1110
   End
   Begin VB.Label LblShakhe 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " :Ã„⁄ ò· ‘«ŒÂ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   8760
      Width           =   1605
   End
   Begin VB.Label LblBandel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " : ⁄œ«œ »‰œ· "
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9540
      TabIndex        =   2
      Top             =   9360
      Width           =   1305
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " :Ã„⁄ ò· ò—«ÌÂ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9210
      TabIndex        =   1
      Top             =   8760
      Width           =   1605
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   8760
      Width           =   10575
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "mnuEdit"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "òÅÌ"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "«‰ ﬁ«·"
      End
      Begin VB.Menu ln1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "ﬁ—«— œ«œ‰"
      End
      Begin VB.Menu ln2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Ã” ÃÊ"
         Begin VB.Menu mnuRow 
            Caption         =   "—œÌ›"
         End
         Begin VB.Menu ln3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBarname 
            Caption         =   "»«—‰«„Â"
         End
      End
   End
End
Attribute VB_Name = "FrmEditDetailA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Ahan As Boolean 'Ahan=True    AEL= False
Public ParvaneCode As String

Private Sub Command1_Click()

End Sub

Private Sub CmdPrint_Click()
 Dim TableName As String
 Dim i As Integer, j As Integer
 
 With Grid2
    .OpenFile App.Path & "\AhanSarBarg.cel"
    
    If Ahan Then
       TableName = "TabAhan_Master"
    Else
       TableName = "TabAEL_Master"
    End If
    '
    Call LoadDataInMasterRows(TableName)
    '
    For i = 1 To Grid1.Rows - 1
        For j = 1 To 11
            .Cell(.Rows - 1, j).Text = Grid1.Cell(i, j).Text
        Next
        .AddItem ""
    Next

    '''MAKE LAST ROW and Show Totals
    .Range(.Rows - 1, 10, .Rows - 1, 11).Merge
    .Cell(.Rows - 1, 10).Text = "Ã„⁄ ò—«ÌÂ "
    .Range(.Rows - 1, 8, .Rows - 1, 9).Merge  ' Keraye Value
    .Cell(.Rows - 1, 8).Text = Sum(2)
     '
    .Range(.Rows - 1, 1, .Rows - 1, 2).Merge
    .Cell(.Rows - 1, 1).Text = Sum(3) 'Shakhe Value
    .Range(.Rows - 1, 3, .Rows - 1, 5).Merge
    .Cell(.Rows - 1, 3).Text = "Ã„⁄ ò· ‘«ŒÂ"
     '
     .AddItem ""
     '
    .Range(.Rows - 1, 10, .Rows - 1, 11).Merge
    .Cell(.Rows - 1, 10).Text = " ⁄œ«œ »‰œ· "
    .Range(.Rows - 1, 8, .Rows - 1, 9).Merge  ' Bandel Value
    .Cell(.Rows - 1, 8).Text = Sum(5)
     '
    .Range(.Rows - 1, 1, .Rows - 1, 2).Merge
    .Cell(.Rows - 1, 1).Text = Sum(6) 'Weight Value
    .Range(.Rows - 1, 3, .Rows - 1, 5).Merge
    .Cell(.Rows - 1, 3).Text = "Ã„⁄ Ê“‰ "
    '''''''''
    '''''''''
    .Range(11, 1, .Rows - 1, .Cols - 1).Borders(cellInsideHorizontal) = cellThin
    .Range(11, 1, .Rows - 1, .Cols - 1).Borders(cellInsideVertical) = cellThin
    .Range(11, 1, .Rows - 1, .Cols - 1).Borders(cellEdgeLeft) = cellThick
    .Range(11, 1, .Rows - 1, .Cols - 1).Borders(cellEdgeRight) = cellThick
    .Range(11, 1, .Rows - 1, .Cols - 1).Borders(cellEdgeBottom) = cellThick
    ''
    .Range(11, 1, .Rows - 1, .Cols - 1).Alignment = cellCenterCenter
    '
    .Range(12, 7, .Rows - 1, 7).FontSize = 10
    '
    .PrintPreview 100
 End With
 
End Sub

Private Sub Form_Activate()
  If Grid1.Rows > 1 Then Grid1.Cell(1, 10).SetFocus
    
  FrmFindAhan.LblWait = ""

End Sub

Private Sub Form_Load()
 On Error Resume Next
 
 Dim Table As String
 Dim i As Byte
   RightToLeft = True
   '
   BackColor = RGB(58, 120, 200)
   '
   TxtDate.Text = TxtDate.Today
   '
   If Ahan Then
      Table = "TabAhan_Detail"
   Else
      Table = "TabAEL_Detail"
   End If
 
   With Grid1
       .BackColorFixed = RGB(90, 158, 214)
       .BackColorFixedSel = RGB(110, 180, 230)
       .BackColorBkg = RGB(90, 158, 214)
       .BackColorScrollBar = RGB(231, 235, 247)
       .BackColor1 = RGB(231, 235, 247)
       .BackColor2 = RGB(239, 243, 255)
       .GridColor = RGB(148, 190, 231)
       '
       .Rows = 1
       '
       .RowHeight(0) = 60
       '
       .Cell(0, 10).WrapText = True
       '
       
       ''''''''''''''''''''''''''''''''''
       .Column(0).Width = 10
       .Column(1).Width = 80 'Mobile
       .Column(2).Width = 80 'Total
       .Column(3).Width = 55 'size
       .Column(4).Width = 55 'Shakhe
       .Column(5).Width = 35 'tedad
       .Column(6).Width = 40 'Weight
       .Column(7).Width = 85 'Kamioon
       .Column(8).Width = 90 'Anbar
       .Column(9).Width = 60 'Date
       .Column(10).Width = 60 'Barname
       .Column(11).Width = 40 'Radif
       '''''''''''''''''''''''''''''''''''
       
       '
       .Cell(0, 11).Text = "—œÌ›"
       .Cell(0, 10).Text = "‘„«—Â »«—‰«„Â"
       .Cell(0, 9).Text = " «—ÌŒ"
       .Cell(0, 8).Text = "«‰»«—  Œ·ÌÂ"
       .Cell(0, 7).Text = "‘„«—Â ò«„ÌÊ‰"
       .Cell(0, 6).Text = "Ê“‰"
       .Cell(0, 5).Text = " ⁄œ«œ"
       .Cell(0, 3).Text = "‘«ŒÂ"
       .Cell(0, 4).Text = "”«Ì“"
       .Cell(0, 2).Text = "ò· ò—«ÌÂ"
       .Cell(0, 1).Text = "„Ê»«Ì·"
       
       '''''''''''''''''''''''''''''''''''
       rs.Open "SELECT COUNT(Parvane),SUM(Total),SUM(Tedad),SUM(Shakhe),SUM(Weight) " & _
               "FROM " & Table & " " & _
               "WHERE Parvane='" & ParvaneCode & "' ", CNS
       '
       If rs.EOF Then
          rs.Close
          Exit Sub
       End If
       '
      
       
       
       LblTotal = " Ã„⁄ ò· ò—«ÌÂ" & "    " & rs(1)
       LblBandel = "  ⁄œ«œ »‰œ·" & "    " & rs(2)
       LblShakhe = " Ã„⁄ ò· ‘«ŒÂ" & "    " & rs(3)
       LblWeight = " Ã„⁄  Ê“‰ " & "    " & rs(4)
       
       rs.Close
       
       
       ''''
       rs.Open "SELECT * FROM " & Table & " " & _
               "WHERE Parvane='" & ParvaneCode & "' " & _
               "ORDER BY Count0", CNS
               
       While Not rs.EOF
           
           .AddItem rs(11) & vbTab & rs(10) & vbTab & rs(8) & _
                      vbTab & rs(9) & vbTab & rs(7) & vbTab & rs(6) _
                     & vbTab & rs(5) & vbTab & rs(4) & vbTab & rs(3) _
                     & vbTab & rs(2) & vbTab & rs(1)
           rs.MoveNext
           i = i + 1
           
       Wend
       
       rs.Close
       '
       For i = 1 To 11
           .Column(i).AutoFit
       Next
       '
       .Range(1, 1, .Rows - 1, 11).Alignment = cellCenterCenter
       .Range(1, 11, .Rows - 1, 11).Locked = True
       .Range(1, 11, .Rows - 1, 11).ForeColor = vbMagenta
       '
       .Range(1, 1, .Rows - 1, 6).Mask = cellNumeric
       .Range(1, 4, .Rows - 1, 4).Mask = cellDefaultMask 'SIZE
       '
       .Column(1).CellType = cellButton
       '
       .Range(1, 7, .Rows - 1, 8).FontName = "Titr"
       .Range(1, 7, .Rows - 1, 8).FontBold = True
       .Range(1, 7, .Rows - 1, 8).FontSize = 12
       

   End With
End Sub

Function Sum(Column As Byte) As Currency
  Dim i As Integer
  Dim SumX As Currency
      
      SumX = 0
      For i = 1 To Grid1.Rows - 1
          SumX = SumX + Val(Grid1.Cell(i, Column).Text)
      Next
      Sum = SumX
End Function

Private Sub Grid1_ButtonClick(ByVal Row As Long, ByVal Col As Long)
 Dim Count0 As Integer, i As Integer
 Dim ArrFieldName(1 To 10) As String
 Dim strSQL As String
 Dim Table, Tonaj As String
  '
   If Ahan Then
      Table = "TabAhan_Detail"
      Tonaj = "TabAhan_Tonaj"
   Else
      Table = "TabAEL_Detail"
      Tonaj = "TabAEL_Tonaj"
   End If
  
   ArrFieldName(1) = "BarNumber='"
   ArrFieldName(2) = "',DBarname='"
   ArrFieldName(3) = "',Anbar='"
   ArrFieldName(4) = "',Kamioon='"
   ArrFieldName(5) = "',Weight="
   ArrFieldName(6) = ",Tedad="
   ArrFieldName(7) = ",Size0='"
   ArrFieldName(8) = "',Shakhe="
   ArrFieldName(9) = ",Total="
   ArrFieldName(10) = ",Mobile='"
   '
   For i = 1 To 10
       strSQL = strSQL + ArrFieldName(i) & Grid1.Cell(Row, 11 - i).Text
   Next
   strSQL = strSQL + "'"
   
   With Grid1
        Count0 = Val(.Cell(Row, 11).Text)
        '
        rs.Open "UPDATE " & Table & " SET " & strSQL & " " & _
                "WHERE Parvane='" & ParvaneCode & "' AND Count0=" & Count0, CNS
        '
        Call UpdateTonaj(Table, Tonaj, Row)
        '
        MsgBox "«ÿ·«⁄«  »« „ÊﬁÌ  À»  ‘œ", vbInformation, "‘„«—Â —œÌ›   " & Count0
                 
   End With
End Sub

Private Sub Grid1_CellChange(ByVal Row As Long, ByVal Col As Long)
   Select Case Col
        Case 2:      LblTotal = " Ã„⁄ ò· ò—«ÌÂ" & "    " & Sum(2)
        Case 3:      LblShakhe = " Ã„⁄ ò· ‘«ŒÂ" & "    " & Sum(3)
        Case 5:      LblBandel = "  ⁄œ«œ »‰œ·" & "    " & Sum(5)
        Case 6:      LblWeight = " Ã„⁄  Ê“‰ " & "    " & Sum(6)
   End Select
End Sub

Private Sub Grid1_GotFocus()
 Dim oldKB As Long
 
  oldKB = GetKeyboardLayout(0)
  'Change keyboard Engish
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  With Grid1.ActiveCell
       If KeyCode = 13 Then
          If .Col = 1 Then
             If .Row <> Grid1.Rows - 1 Then
                Grid1.Cell(.Row + 1, 10).SetFocus
                KeyCode = 0
             Else
                Grid1.Cell(1, 10).SetFocus
             End If
          Else
             Grid1.Cell(.Row, .Col - 1).SetFocus
             KeyCode = 0
          End If
       End If
  End With
End Sub

Private Sub hfg_Click()

End Sub

Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then Grid1_ButtonClick Grid1.ActiveCell.Row, Grid1.ActiveCell.Col
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then PopupMenu mnuEdit
End Sub

Private Sub kTxtBarnameDate_Validate(Cancel As Boolean)
End Sub

Private Sub CmdOk_Click()
 Dim Table As String
 Dim i As Integer
   
   Grid1.Visible = False
   '
   If Ahan Then
      Table = "TabAhan_Detail"
   Else
      Table = "TabAEL_Detail"
   End If
   '
   Grid1.Rows = 1
   '
   rs.Open "SELECT COUNT(Parvane),SUM(Total),SUM(Tedad),SUM(Shakhe),SUM(Weight) " & _
           "FROM " & Table & " " & _
           "WHERE Parvane='" & ParvaneCode & "' " & _
           "AND DBarname='" & Mid(TxtDate.Text, 3) & "'", CNS
   '
   If rs.EOF Then
      rs.Close
      Exit Sub
   End If
   '
  
   
   
   LblTotal = " Ã„⁄ ò· ò—«ÌÂ" & "    " & rs(1)
   LblBandel = "  ⁄œ«œ »‰œ·" & "    " & rs(2)
   LblShakhe = " Ã„⁄ ò· ‘«ŒÂ" & "    " & rs(3)
   LblWeight = " Ã„⁄  Ê“‰ " & "    " & rs(4)
   
   rs.Close
   
   
   ''''
   rs.Open "SELECT * FROM " & Table & " " & _
           "WHERE Parvane='" & ParvaneCode & "' " & _
           "AND DBarname='" & Mid(TxtDate.Text, 3) & "' " & _
           "ORDER BY Count0", CNS
           
   If rs.EOF Then
      MsgBox "«ÿ·«⁄« Ì œ— «Ì‰  «—ÌŒ „ÊÃÊœ ‰„Ì »«‘œ", vbExclamation, ""
      rs.Close
      Exit Sub
   End If
   
   While Not rs.EOF

       Grid1.AddItem rs(11) & vbTab & rs(10) & vbTab & rs(8) & _
                  vbTab & rs(9) & vbTab & rs(7) & vbTab & rs(6) _
                 & vbTab & rs(5) & vbTab & rs(4) & vbTab & rs(3) _
                 & vbTab & rs(2) & vbTab & rs(1)
       rs.MoveNext
       i = i + 1

   Wend

   rs.Close
   
   For i = 1 To 11
       Grid1.Column(i).AutoFit
   Next
      
   With Grid1
       .Range(1, 1, .Rows - 1, 11).Alignment = cellCenterCenter
       .Range(1, 11, .Rows - 1, 11).Locked = True
       .Range(1, 11, .Rows - 1, 11).ForeColor = vbMagenta
       '
       .Range(1, 1, .Rows - 1, 6).Mask = cellNumeric
       '
       .Column(1).CellType = cellButton
       '
       .Range(1, 7, .Rows - 1, 8).FontName = "Titr"
       .Range(1, 7, .Rows - 1, 8).FontBold = True
       .Range(1, 7, .Rows - 1, 8).FontSize = 12
       '
       .Visible = True
  End With
End Sub

Private Sub mnuBarname_Click()
 Dim inp As String, b As Boolean
 Dim i As Integer
  
  inp = InputBox("‘„«—Â »«—‰«„Â —« Ê«—œ ‰„«ÌÌœ", "„Â—Ê—“«‰", "")
  '
  b = False
  For i = 1 To Grid1.Rows - 1
      If Trim(Grid1.Cell(i, 10).Text) = Trim(inp) Then
         b = True
         Exit For
      End If
  Next
  '
  If b Then
     Grid1.Cell(i, 11).SetFocus
     Grid1.Range(i, 0, i, 11).Selected
  Else
     MsgBox "‘„«—Â »«—‰«„Â „Ê—œ ‰Ÿ— Ì«›  ‰‘œ", vbInformation, "„Â—Ê—“«‰"
  End If
End Sub

Private Sub mnuCopy_Click()
   Grid1.Selection.CopyData
End Sub

Private Sub mnuCut_Click()
   Grid1.Selection.CutData
End Sub

Private Sub mnuPaste_Click()
   Grid1.Selection.PasteData
End Sub

Private Sub mnuRow_Click()
 Dim inp As Integer
 '
 On Error Resume Next
   inp = Val(InputBox("⁄œœ „Ê—œ ‰Ÿ— —« Ê«—œ ‰„«ÌÌœ", "„Â—Ê—“«‰", "0"))
   Grid1.Cell(inp, 11).SetFocus
   Grid1.Range(inp, 0, inp, 11).Selected
      
End Sub

Sub UpdateTonaj(xTable, xTonaj As String, ByVal xRow As Integer)
    Dim Weight As Long
    Dim Tedad As Long
    Dim Shakhe As Long
    Dim Price As Currency
    '

    rs.Open "SELECT SUM(Weight),SUM(Tedad) ," & _
            "SUM(Shakhe),SUM(Total) " & _
            "FROM " & xTable & " " & _
            "WHERE Parvane='" & ParvaneCode & "' " & _
            "AND DBarname='" & Grid1.Cell(xRow, 9).Text & "'", CNS
   
   If Not rs.EOF Then
      Weight = rs(0)
      Tedad = rs(1)
      Shakhe = rs(2)
      Price = rs(3)
      '''
   End If
   rs.Close
   '
   rs.Open "UPDATE " & xTonaj & " " & _
           "SET TonajEx=" & Weight & ",TotalBandel=" & Tedad & "," & _
           "TonajMod=TonajPar-" & Weight & "," & _
           "TotalShakhe=" & Shakhe & ",TotalPrice=" & Price & _
           " WHERE Parvane='" & ParvaneCode & "'" & _
           " AND BarDate='" & Grid1.Cell(xRow, 9).Text & "'", CNS


End Sub

Function LoadDataInMasterRows(TabName As String) As Boolean
Dim Myrs As New Recordset
Dim ks As String
'''''''''Load Data In Master Rows
LoadDataInMasterRows = True

  Myrs.Open "SELECT * FROM " & TabName & " " & _
          "WHERE Parvane='" & ParvaneCode & "'", CNS
  
  If Myrs.EOF Then
     MsgBox "ç‰Ì‰ Å—Ê«‰Â «Ì „ÊÃÊœ ‰Ì” ", vbExclamation, ""
     Myrs.Close
     LoadDataInMasterRows = False
     Exit Function
  End If
  
  With Grid2
    .Cell(1, 3).Text = "ò‘ Ì :" & Space(1) & Trim(Myrs("Keshti"))
    .Cell(1, 1).Text = "»«—‰«„Â :" & Space(0) & Trim(Myrs("Barname"))
    .Cell(1, 1).Font.Name = "B Zar"
    .Cell(1, 1).Font.Size = 12
    
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
    
    Dim inp As String
     inp = InputBox("", ":ﬁ«»·  ÊÃÂ ”—Ê— ê—«„Ì ")
     If Trim(inp) <> Empty Then
        .Cell(3, 7).Text = " ﬁ«»·  ÊÃÂ ”—Ê— ê—«„Ì: " & inp 'Sarvar
     Else
        .Cell(3, 7).Text = ": ﬁ«»·  ÊÃÂ ”—Ê— ê—«„Ì" 'Sarvar
     End If
    
  End With
  Myrs.Close
  
  Set Myrs = Nothing
End Function


