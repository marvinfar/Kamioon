VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form FrmFinishAhan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "??? ??? ????? ?? ?????? ??? ????"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "B Zar"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmFinishAhan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMakeBorder 
      Caption         =   "Make Border"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin PrjShayan.TypeButton CmdPrev 
      Height          =   495
      Left            =   8760
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "??? ?????"
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
      MICON           =   "FrmFinishAhan.frx":169B2
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
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   14631
      Cols            =   12
      DefaultFontName =   "B Zar"
      DefaultFontSize =   12
      DefaultFontBold =   -1  'True
      GridColor       =   -2147483630
      GridLiness      =   -1  'True
      Rows            =   15
   End
   Begin PrjShayan.TypeButton CmdComment 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "?????? ???????"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmFinishAhan.frx":169CE
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdDelComment 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "??? ???????"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmFinishAhan.frx":169EA
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "FrmFinishAhan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rAEL As Boolean
Public fParvaneCode As String

Private Sub CmdComment_Click()
 Dim l As Integer, i As Integer
 
 l = Grid1.Rows
 Grid1.Rows = Grid1.Rows + 5
 For i = l To Grid1.Rows - 1
     Grid1.RowHeight(i) = 40
     Grid1.Range(i, 1, i, Grid1.Cols - 1).Merge
 Next
 Grid1.Range(l, 1, Grid1.Rows - 1, Grid1.Cols - 1).Alignment = cellRightCenter
End Sub

Private Sub CmdDelComment_Click()
 Dim i As Integer
 
 For i = 1 To 5
     Grid1.RemoveItem Grid1.Rows - 1
 Next
End Sub

Private Sub CmdMakeBorder_Click()
   With Grid1
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

Private Sub CmdPrev_Click()
  With Grid1.PageSetup
     
     .PaperSize = cellPaperA4  'A4 paper
     .Orientation = cellPortrait  'Portrait
     .PrintTitleRows = 11
     .LeftMargin = 1
     .RightMargin = 1
     .BottomMargin = 2.5
     .TopMargin = 1
     .CenterHorizontally = True  'Center on page horizontally
     .PrintFixedColumn = False
     .PrintFixedColumn = True
     .FooterFont.Name = "Traditional Arabic"
     .FooterFont.Bold = True
     .FooterFont.Size = 13
     .FooterMargin = 0.5
     .Footer = " ???? ????????????? ?????? ?????? ?????????? ???? ?????? ???? ????? ? ????????????? ???? ???????????? ????? ??????? ??? ??ѡ ???? ??? " & " E-Mail: mehrvarzantarabar@yahoo.com" & Space(0) & vbCrLf & _
               "?? ???? 73337-43156" & Space(10) & "?ǘ? 3239400-0181" & Space(10) & "???? 4-3239880-0181" & Space(15) & "???? &P"
    '
  End With
  
  Grid1.PrintPreview

End Sub

Private Sub Form_Load()
On Error Resume Next

 Dim i As Byte
 Dim ks As String
 Dim MTable, DTable, Tonaj As String
 Dim Weight As Long
 
  BackColor = RGB(83, 132, 178)

'
  If rAEL Then
     MTable = "TabAEL_Master"
     DTable = "TabAEL_Detail"
     Tonaj = "TabAEL_Tonaj"
  Else
     MTable = "TabAhan_Master"
     DTable = "TabAhan_Detail"
     Tonaj = "TabAhan_Tonaj"
  End If
'
  With Grid1
       .Cols = 12
       .Rows = 12
       
       Call MakeMasterRows
       Call CmdMakeBorder_Click
       '''''''''Load Data In Master Rows
       rs.Open "SELECT * FROM " & MTable & " " & _
               "WHERE Parvane='" & fParvaneCode & "'", CNS
       
       .Cell(1, 3).Text = "???? :" & Space(1) & Trim(rs("Keshti"))
       .Cell(1, 1).Text = "??????? :" & Space(0) & Trim(rs("Barname"))
       .Cell(1, 1).Font.Name = "B Zar"
       .Cell(1, 1).Font.Size = 12
       
       .Cell(4, 7).Text = Trim(rs("Typekala"))
       .Cell(5, 7).Text = Trim(rs("Parvane"))
       .Cell(6, 7).Text = Trim(rs("DKootaj")) & Space(13 - Len(Trim(rs("Kootaj")))) & Trim(rs("Kootaj"))
        .Cell(6, 7).Font.Name = "B zar"
        .Cell(6, 7).Font.Bold = True
       '
       .Cell(7, 7).Text = Trim(rs("DGhabz")) & Space(17 - Len(Trim(rs("Ghabz")))) & Trim(rs("Ghabz"))
       .Cell(8, 7).Text = Trim(rs("Dparvane"))
       .Cell(7, 7).Font.Name = "B Zar"
       .Cell(7, 7).Font.Size = 12
       
       ks = Trim(rs("Sizekala"))
       ks = Replace(ks, "x", Chr$(215))
       .Cell(4, 1).Text = ks
       .Cell(4, 1).Font.Name = "B Zar"
       .Cell(4, 1).Font.Size = 12
       
       .Cell(5, 1).Text = Trim(rs("NWeight"))
       .Cell(6, 1).Text = Trim(rs("Weight"))
       .Cell(7, 1).Text = Val(rs("Shakhe")) & Space(17) & Val(rs("Bandel"))
       .Cell(8, 1).Text = Trim(rs("Tarkhiskar"))
       
       '
       .Cell(10, 8).Text = Trim(rs("Saheb"))
       .Cell(10, 4).Text = Trim(rs("Etebar"))
       .Cell(10, 1).Text = Trim(rs("Gharardad"))
       rs.Close
       ''''''''
      rs.Open "SELECT SUM(Weight) FROM " & DTable & " " & _
              "WHERE Parvane='" & fParvaneCode & "' ", CNS
              
      Weight = rs(0)
      rs.Close
      
       rs.Open "SELECT * FROM " & DTable & " " & _
               "WHERE Parvane='" & fParvaneCode & "' " & _
               "ORDER BY Count0", CNS
       While Not rs.EOF
       
           .AddItem rs(11) & vbTab & rs(10) & vbTab & rs(8) & _
                      vbTab & rs(9) & vbTab & rs(7) & vbTab & rs(6) _
                     & vbTab & rs(5) & vbTab & rs(4) & vbTab & rs(3) _
                     & vbTab & rs(2) & vbTab & rs(1)
           rs.MoveNext

       Wend
       rs.Close
       '
       .Range(12, 11, .Rows - 1, 11).BackColor = &HE0E0E0
       .Range(11, 1, 11, 11).BackColor = &HE0E0E0
 '''
       .Range(12, 1, .Rows - 1, 11).Alignment = cellRightCenter
       .Range(12, 1, .Rows - 1, 11).FontName = "B Titr"
       .Range(12, 1, .Rows - 1, 11).FontSize = 9
       ''
       .Range(12, 1, .Rows - 1, 10).Alignment = cellCenterCenter
       .Range(12, 11, .Rows - 1, 11).Alignment = cellCenterCenter
       '
       .Range(12, 1, .Rows - 1, 11).Borders(cellInsideHorizontal) = cellThin
       .Range(12, 1, .Rows - 1, 11).Borders(cellInsideVertical) = cellThin
       '
       .Range(12, 1, .Rows - 1, 11).Borders(cellEdgeLeft) = cellThick
       .Range(12, 1, .Rows - 1, 11).Borders(cellEdgeRight) = cellThick
       .Range(12, 1, .Rows - 1, 11).Borders(cellEdgeBottom) = cellThick
       ''''
        '''Insert Bottom Rows Information''''''''''''''''''
        ''Load Ahan Tonaj
        rs.Open "SELECT MIN(TonajMod),SUM(TotalPrice) FROM " & Tonaj & " " & _
                      " WHERE Parvane='" & fParvaneCode & "' ", CNS
         
        ' Calculate Bandel And Shakhe From TabDetail
        Dim rsT As New Recordset
        
        rsT.Open "SELECT SUM(Tedad),SUM(Shakhe) FROM " & DTable & " " & _
                 " WHERE Parvane='" & fParvaneCode & "' ", CNS
                 
       .AddItem ""
       .Range(.Rows - 1, 9, .Rows - 1, 10).Merge
       .Cell(.Rows - 1, 9).Text = "  ??ǎ ??????"
       .Cell(.Rows - 1, 8).Text = .Cell(5, 1).Text
       '
       .Cell(.Rows - 1, 7).Text = "?? ??ǎ ?????"
       '
       .Range(.Rows - 1, 5, .Rows - 1, 6).Merge
       .Cell(.Rows - 1, 5).Text = Weight
       '
       .Range(.Rows - 1, 3, .Rows - 1, 4).Merge
       '
        If rs(0) < 0 Then
         .Cell(.Rows - 1, 3).Text = "????? ???"
         .Cell(.Rows - 1, 2).Text = Abs(rs(0))
       ElseIf rs(0) > 0 Then
         .Cell(.Rows - 1, 3).Text = "??? ???"
         .Cell(.Rows - 1, 2).Text = rs(0)
       Else
         .Cell(.Rows - 1, 3).Text = "???ǁ??"
         .Cell(.Rows - 1, 2).Text = rs(0)
       End If
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ''''
       .AddItem "" ''''''
       .Range(.Rows - 1, 9, .Rows - 1, 10).Merge
       .Cell(.Rows - 1, 9).Text = "??? ????"
       .Cell(.Rows - 1, 8).Text = rsT(0)
       '
       .Cell(.Rows - 1, 7).Text = "??? ?? ????"
       '
       .Range(.Rows - 1, 5, .Rows - 1, 6).Merge
       .Cell(.Rows - 1, 5).Text = rsT(1)
       '
       .Range(.Rows - 1, 3, .Rows - 1, 4).Merge
       .Cell(.Rows - 1, 3).Text = "?? ?????"
       '
       .Cell(.Rows - 1, 2).Text = rs(1)
       
       rs.Close
       
       rsT.Close
       Set rsT = Nothing
       '''''''''''''''''''''''''''''''''''''''''''''''''''''
       .RowHeight(.Rows - 1) = 40
       .RowHeight(.Rows - 2) = 40
       '
       .Range(.Rows - 2, 2, .Rows - 1, 10).Alignment = cellCenterCenter
       .Range(.Rows - 2, 2, .Rows - 1, 10).FontName = "B Titr"
       .Range(.Rows - 2, 2, .Rows - 1, 10).FontSize = 9
       '
       .Range(.Rows - 2, 2, .Rows - 1, 10).Borders(cellInsideHorizontal) = cellThick
       .Range(.Rows - 2, 2, .Rows - 1, 10).Borders(cellInsideVertical) = cellThick
       .Range(.Rows - 2, 2, .Rows - 1, 10).Borders(cellEdgeRight) = cellThick
       .Range(.Rows - 2, 2, .Rows - 1, 10).Borders(cellEdgeLeft) = cellThick
       .Range(.Rows - 2, 2, .Rows - 1, 10).Borders(cellEdgeBottom) = cellThick
       '''''''''''''
       .Range(12, 6, .Rows - 3, 7).FontSize = 11
  End With
       Call CmdPrev_Click

End Sub

Sub MakeMasterRows()
  Dim i As Integer
  
   With Grid1
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
       '
       .Range(7, 7, 7, 8).FontName = "Traditional Arabic"
       .Range(7, 7, 7, 8).FontBold = True
       .Range(7, 7, 7, 8).FontSize = 12

       ''''''''''''''
       .Range(11, 1, 11, 11).FontName = "B Titr"
       .Range(11, 1, 11, 11).FontBold = True
       .Range(11, 1, 11, 11).FontSize = 8
       '''''''''''''
       .Cell(1, 1).Text = ":???????"
       .Cell(1, 1).Font.Name = "B Badr"
       .Cell(1, 1).Font.Size = 12
       
       .Cell(1, 1).Alignment = cellRightCenter
       .Cell(1, 1).Font.Name = "Traditional Arabic"
       .Cell(1, 1).Font.Name = "Traditional Arabic"
       .Cell(1, 3).Text = ":????"
       .Cell(1, 4).Alignment = cellRightCenter
       '''
       .Cell(4, 3).Text = "???? ??????"
       .Cell(5, 3).Text = "??? ?????????"
       .Cell(6, 3).Text = "??? ????????"
       .Cell(7, 3).Text = "????? ???? ? ?? ????"
       .Cell(8, 3).Text = "??? ????? ???"
       ''''''''''''''''
       .Cell(4, 9).Text = "????? ???????"
       .Cell(5, 9).Text = "??????? ???????"
       .Cell(6, 9).Text = "????? ?????ǎ"
       .Cell(7, 9).Text = "????? ??? ?????"
       .Cell(8, 9).Text = "?????? ????????"
       '
       .Cell(11, 11).Text = "????"
       .Cell(11, 10).Text = "????? ???????"
       .Cell(11, 9).Text = "?????"
       .Cell(11, 8).Text = "????? ?????"
       .Cell(11, 7).Text = "????? ??????"
       .Cell(11, 6).Text = "???"
       .Cell(11, 5).Text = "?????"
       .Cell(11, 3).Text = "????"
       .Cell(11, 4).Text = "????"
       .Cell(11, 2).Text = "?? ?????"
       .Cell(11, 1).Text = "??????"

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
       If Not rAEL Then
          .Cell(1, 7).Text = "???? ??ѐ??? ?ј? ??? ? ??? ???????? ?????" 'List
       Else
          .Cell(1, 7).Text = "??????ј??? ?????? ????????? ????????" 'List
          .Range(1, 7, 2, 11).FontSize = 14
       End If
       '
       Dim inp As String
        inp = InputBox("", ":???? ???? ???? ????? ")
        If Trim(inp) <> Empty Then
           .Cell(3, 7).Text = " ???? ???? ???? ?????: " & inp 'Sarvar
        Else
           .Cell(3, 7).Text = ": ???? ???? ???? ?????" 'Sarvar
        End If
       ''''''''''''''''''''''''
       .Cell(10, 10).Text = "???? ????" 'Saheb KALA
       .Cell(10, 7).Text = "????? ??????" 'Etebar
       .Cell(10, 2).Text = "????? ???????" 'Etebar
       ''''
       .Cell(4, 1).Font.Name = "Tahoma"
       .Cell(4, 1).Font.Bold = True
       .Cell(4, 1).Font.Size = 9
       '
       .Cell(10, 4).Font.Size = 10 'Etebar
       .Cell(10, 4).Alignment = cellCenterCenter 'Etebar
       .Cell(10, 1).Font.Size = 10 'Gharardad
       .Cell(10, 1).Alignment = cellCenterCenter 'gharadad

   End With

End Sub

