VERSION 5.00
Object = "{9DBDC544-49CA-11D7-B1ED-C2237039C523}#1.1#0"; "FarDate.Ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form FrmEditDetailKan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«’·«Õ «ÿ·«⁄«  Ã«‰»Ì ò«‰ Ì‰—"
   ClientHeight    =   10215
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   11130
   Icon            =   "FrmEditDetailKan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin FlexCell.Grid Grid1 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   13573
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
      TabIndex        =   5
      Top             =   360
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
      Left            =   2640
      TabIndex        =   6
      Top             =   360
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
      MICON           =   "FrmEditDetailKan.frx":169B2
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FlexCell.Grid Grid2 
      Height          =   495
      Left            =   3840
      TabIndex        =   11
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
   Begin PrjShayan.TypeButton CmdPrint 
      Height          =   495
      Left            =   1800
      TabIndex        =   12
      Top             =   9480
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
      MICON           =   "FrmEditDetailKan.frx":169CE
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LblKantiner 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ": ⁄œ«œ ò«‰ Ì‰—"
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
      Left            =   9360
      TabIndex        =   10
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "»«  ‘ò— «“ Õ”‰ «‰ Œ«» ‘„«"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2760
      TabIndex        =   9
      Top             =   2880
      Width           =   5265
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘—ò  ‰—„ «›“«—Ì ¬—«ÌÂ   "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2640
      TabIndex        =   8
      Top             =   4200
      Width           =   5415
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ·›‰  „«”  2231276-2253511"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2040
      TabIndex        =   7
      Top             =   5640
      Width           =   6840
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
      Left            =   9120
      TabIndex        =   4
      Top             =   8880
      Width           =   1605
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
      Left            =   2280
      TabIndex        =   3
      Top             =   8880
      Width           =   1110
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " —« »›‘«—ÌœF4    »—«Ì –ŒÌ—Â ò—œ‰ „Ê«—œ  €ÌÌ— œ«œÂ ‘œÂ  Å” «“  ò„Ì· Â— ”ÿ— ò·Ìœ"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2460
      TabIndex        =   2
      Top             =   -120
      Width           =   8325
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ê Ì« »——ÊÌ œò„Â «Ì òÂ œ— ” Ê‰ „Ê»«Ì· ÊÃÊœ œ«—œ ò·Ìò ò‰Ìœ"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4650
      TabIndex        =   1
      Top             =   240
      Width           =   6165
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   8880
      Width           =   10695
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "mnuEdit"
      Visible         =   0   'False
      Begin VB.Menu MnuCopy 
         Caption         =   "òÅÌ"
      End
      Begin VB.Menu MnuCut 
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
            Caption         =   "‘„«—Â »«—‰«„Â"
         End
      End
   End
End
Attribute VB_Name = "FrmEditDetailKan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ParvaneCode As String

Private Sub CmdOk_Click()
' On Error Resume Next
 
 Dim i As Integer
   
   Grid1.Visible = False
   '
   Grid1.Rows = 1
   '
   rs.Open "SELECT COUNT(Parvane),SUM(Total),SUM(Weight),SUM(Tedad) " & _
           "FROM TabKantiner_Detail " & _
           "WHERE Parvane='" & ParvaneCode & "' " & _
           "AND BarNameDate='" & Mid(TxtDate.Text, 3) & "'", CNS
   '
   If rs.EOF Then
      rs.Close
      Exit Sub
   End If
   '
   
   LblTotal = " Ã„⁄ ò· ò—«ÌÂ" & "    " & rs(1)
   LblWeight = " Ã„⁄  Ê“‰ " & "    " & rs(2)
   LblKantiner = "  ⁄œ«œ ò«‰ Ì‰—" & "    " & rs(3)
   
   rs.Close
   
   ''''
   rs.Open "SELECT * FROM TabKantiner_Detail " & _
           "WHERE Parvane='" & ParvaneCode & "' " & _
           "AND BarNameDate='" & Mid(TxtDate.Text, 3) & "' " & _
           "ORDER BY Count0", CNS
   
   If rs.EOF Then
      MsgBox "«ÿ·«⁄« Ì œ— «Ì‰  «—ÌŒ „ÊÃÊœ ‰„Ì »«‘œ", vbExclamation, ""
      rs.Close
      Exit Sub
   End If
   
   While Not rs.EOF
       
           Grid1.AddItem rs(12) & vbTab & rs(11) & vbTab & rs(6) & _
                      vbTab & rs(10) & vbTab & rs(8) & vbTab & rs(7) _
                     & vbTab & rs(5) & vbTab & rs(4) & vbTab & rs(3) _
                     & vbTab & rs(2) & vbTab & rs(1)

       rs.MoveNext
       i = i + 1
       
   Wend

   rs.Close

       Grid1.Column(0).Width = 10
       Grid1.Column(1).Width = 85 'Mobile
       Grid1.Column(2).Width = 80 'Total
       Grid1.Column(3).Width = 85 'Kantiner
       Grid1.Column(4).Width = 40 'Size
       Grid1.Column(5).Width = 35 'tedad
       Grid1.Column(6).Width = 40 'Weight
       Grid1.Column(7).Width = 90 'Kamioon
       Grid1.Column(8).Width = 90 'Anbar
       Grid1.Column(9).Width = 60 'Date
       Grid1.Column(10).Width = 60 'Barname
       Grid1.Column(11).Width = 40 'Radif
   '
   With Grid1
       .Range(1, 1, .Rows - 1, 9).Alignment = cellCenterCenter
       .Range(1, 9, .Rows - 1, 9).Locked = True
       .Range(1, 11, .Rows - 1, 11).ForeColor = vbMagenta
       '
       .Range(1, 1, .Rows - 1, 4).Mask = cellNumeric
       '
       .Column(1).CellType = cellButton
       '
       .Range(1, 5, .Rows - 1, 5).FontName = "Titr"
       .Range(1, 5, .Rows - 1, 5).FontBold = True
       .Range(1, 5, .Rows - 1, 5).FontSize = 12
       '
       .Visible = True
   End With
End Sub

Private Sub CmdPrint_Click()
 Dim TableName As String
 Dim i As Integer, j As Integer
 
 With Grid2
    .OpenFile App.Path & "\KantinerSarBarg.cel"
    '
    Call LoadDataInMasterRows("TabKantiner_Master")
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
    '.Range(.Rows - 1, 10, .Rows - 1, 11).Merge
    .Cell(.Rows - 1, 7).Text = " ⁄œ«œ ò«‰ Ì‰—"
    .Range(.Rows - 1, 5, .Rows - 1, 6).Merge ' Kantiner Value
    .Cell(.Rows - 1, 5).Text = Sum(5)
    '
    .Range(.Rows - 1, 1, .Rows - 1, 2).Merge
    .Cell(.Rows - 1, 1).Text = Sum(6) ' Weight Value
    .Range(.Rows - 1, 3, .Rows - 1, 4).Merge
    .Cell(.Rows - 1, 3).Text = "Ã„⁄ Ê“‰"
     '
    '''''''''
    .Range(11, 1, .Rows - 1, .Cols - 1).Borders(cellInsideHorizontal) = cellThin
    .Range(11, 1, .Rows - 1, .Cols - 1).Borders(cellInsideVertical) = cellThin
    .Range(11, 1, .Rows - 1, .Cols - 1).Borders(cellEdgeLeft) = cellThick
    .Range(11, 1, .Rows - 1, .Cols - 1).Borders(cellEdgeRight) = cellThick
    .Range(11, 1, .Rows - 1, .Cols - 1).Borders(cellEdgeBottom) = cellThick
    ''
    .Range(11, 1, .Rows - 1, .Cols - 1).Alignment = cellCenterCenter
    '
    .Range(12, 7, .Rows - 2, 7).FontSize = 10
    .Range(12, 3, .Rows - 2, 3).FontSize = 10
    '
    .PageSetup.LeftMargin = 0.5
    .PrintPreview 100
 End With

End Sub

Private Sub Form_Activate()
  If Grid1.Rows > 1 Then Grid1.Cell(1, 10).SetFocus
    
  FrmFindAhan.LblWait = ""

End Sub

Private Sub Form_Load()
On Error Resume Next
 
 Dim i As Byte
   RightToLeft = True
   '
   BackColor = RGB(58, 120, 200)
   '
   TxtDate.Text = TxtDate.Today
   '

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
       .Cell(0, 8).WrapText = True
       '
       
       ''''''''''''''''''''''''''''''''''
       .Cell(0, 10).WrapText = True
       .Cell(0, 3).WrapText = True

       '
       
       ''''''''''''''''''''''''''''''''''
       .Column(0).Width = 10
       .Column(1).Width = 85 'Mobile
       .Column(2).Width = 80 'Total
       .Column(3).Width = 85 'Kantiner
       .Column(4).Width = 40 'Size
       .Column(5).Width = 35 'tedad
       .Column(6).Width = 40 'Weight
       .Column(7).Width = 90 'Kamioon
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
       .Cell(0, 4).Text = "”«Ì“"
       .Cell(0, 3).Text = "‘„«—Â ò«‰ Ì‰—"
       .Cell(0, 2).Text = "ò· ò—«ÌÂ"
       .Cell(0, 1).Text = "„Ê»«Ì·"
       
       '''''''''''''''''''''''''''''''''''
       rs.Open "SELECT COUNT(Parvane),SUM(Total),SUM(Weight),SUM(Tedad) " & _
               "FROM TabKantiner_Detail " & _
               "WHERE Parvane='" & ParvaneCode & "' ", CNS
       '
       If rs.EOF Then
          rs.Close
          Exit Sub
       End If
       '
      
       
       
       LblTotal = " Ã„⁄ ò· ò—«ÌÂ" & "    " & rs(1)
       LblWeight = " Ã„⁄  Ê“‰ " & "    " & rs(2)
       LblKantiner = "  ⁄œ«œ ò«‰ Ì‰—" & "    " & rs(3)
       
       rs.Close
       
       
       ''''
       rs.Open "SELECT * FROM TabKantiner_Detail " & _
               "WHERE Parvane='" & ParvaneCode & "' " & _
               "ORDER BY Count0", CNS
               
       While Not rs.EOF
                     
           Grid1.AddItem rs(12) & vbTab & rs(11) & vbTab & rs(6) & _
                      vbTab & rs(10) & vbTab & rs(8) & vbTab & rs(7) _
                     & vbTab & rs(5) & vbTab & rs(4) & vbTab & rs(3) _
                     & vbTab & rs(2) & vbTab & rs(1)

           rs.MoveNext
           i = i + 1
           
       Wend
       
       rs.Close
       '
       For i = 100 To 11
           .Column(i).AutoFit
       Next
       '
       .Range(1, 1, .Rows - 1, 11).Alignment = cellCenterCenter
       .Range(1, 11, .Rows - 1, 11).Locked = True
       .Range(1, 11, .Rows - 1, 11).ForeColor = vbMagenta
       '
       .Column(1).CellType = cellButton
       '
       .Column(1).Mask = cellNumeric
       .Column(2).Mask = cellNumeric
       .Column(4).Mask = cellNumeric
       .Column(5).Mask = cellNumeric
       .Column(6).Mask = cellNumeric
       '
       .Range(1, 7, .Rows - 1, 8).FontName = "Titr"
       .Range(1, 7, .Rows - 1, 8).FontBold = True
       .Range(1, 7, .Rows - 1, 8).FontSize = 12
       '
       .Range(1, 3, .Rows - 1, 3).FontName = "Titr"
       .Range(1, 3, .Rows - 1, 3).FontBold = True
       .Range(1, 3, .Rows - 1, 3).FontSize = 12

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
  
   ArrFieldName(1) = "BarNumber='"
   ArrFieldName(2) = "',BarNameDate='"
   ArrFieldName(3) = "',Anbar='"
   ArrFieldName(4) = "',Kamioon='"
   ArrFieldName(5) = "',Weight="
   ArrFieldName(6) = ",Tedad="
   ArrFieldName(7) = ",size0="
   ArrFieldName(8) = ",Kantiner='"
   ArrFieldName(9) = "',Total="
   ArrFieldName(10) = ",Mobile='"
   '
   For i = 1 To 10
       strSQL = strSQL + ArrFieldName(i) & Grid1.Cell(Row, 11 - i).Text
   Next
   strSQL = strSQL + "'"
   
   With Grid1
        Count0 = Val(.Cell(Row, 11).Text)
        '
        rs.Open "UPDATE TabKantiner_Detail SET " & strSQL & " " & _
                "WHERE Parvane='" & ParvaneCode & "' AND Count0=" & Count0, CNS
        '
        Call UpdateTonaj(Row)
        '
        MsgBox "«ÿ·«⁄«  »« „ÊﬁÌ  À»  ‘œ", vbInformation, "‘„«—Â —œÌ›   " & Count0
                
   End With
End Sub

Private Sub Grid1_CellChange(ByVal Row As Long, ByVal Col As Long)
   Select Case Col
        Case 2:      LblTotal = " Ã„⁄ ò· ò—«ÌÂ" & "    " & Sum(2)
        Case 5:      LblKantiner = "  ⁄œ«œ ò«‰ Ì‰—" & "    " & Sum(5)
        Case 6:      LblWeight = " Ã„⁄  Ê“‰ " & "    " & Sum(6)
   End Select
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  With Grid1.ActiveCell
       If KeyCode = 13 Then
          If .Col = 1 Then
             If .Row <> Grid1.Rows - 1 Then
                Grid1.Cell(.Row + 1, 8).SetFocus
                KeyCode = 0
             Else
                Grid1.Cell(1, 8).SetFocus
             End If
          Else
             Grid1.Cell(.Row, .Col - 1).SetFocus
             KeyCode = 0
          End If
       End If
  End With
End Sub

Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then Grid1_ButtonClick Grid1.ActiveCell.Row, Grid1.ActiveCell.Col
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then PopupMenu mnuEdit
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
     Grid1.Cell(i, 9).SetFocus
     Grid1.Range(i, 0, i, 9).Selected
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
   Grid1.Cell(inp, 9).SetFocus
   Grid1.Range(inp, 0, inp, 9).Selected
End Sub

Sub UpdateTonaj(ByVal xRow As Integer)
    Dim Weight As Long
    Dim Tedad As Long
    Dim Shakhe As Long
    Dim Price As Currency
    '

    rs.Open "SELECT SUM(Weight),SUM(Tedad) ," & _
            "SUM(Shakhe),SUM(Total) " & _
            "FROM TabKantiner_Detail " & _
            "WHERE Parvane='" & ParvaneCode & "' " & _
            "AND BarNameDate='" & Grid1.Cell(xRow, 9).Text & "'", CNS
   
   If Not rs.EOF Then
      Weight = rs(0)
      Tedad = rs(1)
      Shakhe = rs(2)
      Price = rs(3)
      '''
   End If
   rs.Close
   '
   rs.Open "UPDATE  TabKantiner_Tonaj " & _
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
    .Cell(4, 7).Text = Trim(Myrs("Typekala"))
    .Cell(5, 7).Text = Trim(Myrs("Parvane"))
    .Cell(6, 7).Text = Trim(Myrs("DKootaj")) & Space(13 - Len(Trim(Myrs("Kootaj")))) & Trim(Myrs("Kootaj"))
    .Cell(6, 7).Font.Name = "Traditional Arabic"
    .Cell(6, 7).Font.Size = 12
    .Cell(6, 7).Font.Bold = True
    '
    .Cell(7, 7).Text = Trim(Myrs("DGhabz")) & Space(17 - Len(Trim(Myrs("Ghabz")))) & Trim(Myrs("Ghabz"))
    .Cell(8, 7).Text = Trim(Myrs("Dparvane"))
    .Cell(7, 7).Font.Name = "B Zar"
    .Cell(7, 7).Font.Size = 12

    ks = Trim(Myrs("Sizekala"))
    ks = Replace(ks, "x", Chr$(215))
    .Cell(4, 1).Text = ks
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


