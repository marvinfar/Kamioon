VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmKantinerRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅÌ‘ ‰„«Ì‘ ò«‰ Ì‰—"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11370
   BeginProperty Font 
      Name            =   "B Zar"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmKantiner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10410
   ScaleWidth      =   11370
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   9705
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton CmdMakeBorder 
      Caption         =   "Make Border"
      Height          =   390
      Left            =   2040
      TabIndex        =   4
      Top             =   9720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog ComDlg1 
      Left            =   360
      Top             =   9600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin PrjShayan.TypeButton CmdPrev 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   9960
      TabIndex        =   6
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "ÅÌ‘ ‰„«Ì‘"
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
      MICON           =   "FrmKantiner.frx":169B2
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdSave 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   8400
      TabIndex        =   7
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "–ŒÌ—Â"
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
      MICON           =   "FrmKantiner.frx":169CE
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdOpen 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "»«“ ò—œ‰ ê“«—‘"
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
      MICON           =   "FrmKantiner.frx":169EA
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PrjShayan.TypeButton CmdExcel 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "«‰ ﬁ«· »Â «ò”·"
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
      MICON           =   "FrmKantiner.frx":16A06
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
      Height          =   9615
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   16960
      Cols            =   5
      DefaultFontName =   "B Zar"
      DefaultFontSize =   12
      DefaultFontBold =   -1  'True
      DefaultRowHeight=   32
      Rows            =   30
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘—ò  Õ„· Ê ‰ﬁ· „Â—Ê—“«‰  —«»— Œ“—"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   27.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   885
      Left            =   1080
      TabIndex        =   3
      Top             =   1320
      Width           =   7650
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "»«  ‘ò— «“ Õ”‰ «‰ Œ«» ‘„« "
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   35.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1110
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   6840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘—ò  ò«„ÅÌÊ —Ì ¬—«ÌÂ"
      BeginProperty Font 
         Name            =   "B Zar"
         Size            =   45
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1425
      Left            =   1200
      TabIndex        =   1
      Top             =   4560
      Width           =   7530
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00004000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      FillColor       =   &H00008000&
      Height          =   1335
      Left            =   960
      Top             =   1080
      Width           =   7935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   960
      Top             =   2880
      Width           =   7935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      FillColor       =   &H00000080&
      Height          =   1335
      Left            =   960
      Top             =   4560
      Width           =   7935
   End
   Begin VB.Menu mnuA 
      Caption         =   "mnuA"
      Visible         =   0   'False
      Begin VB.Menu mnuSave2Cel 
         Caption         =   "–ŒÌ—Â ›«Ì·"
      End
      Begin VB.Menu mnuLn1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpen 
         Caption         =   "»«“ ò—œ‰ ›«Ì· –ŒÌ—Â ‘œÂ"
      End
      Begin VB.Menu mnuLn2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave2Exc 
         Caption         =   "–ŒÌ—Â ›«Ì· »Â «ò”·"
      End
   End
End
Attribute VB_Name = "FrmKantinerRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ParvaneCode, DayDate As String
Public GetPrintDate As String
Public Finish As Boolean

Private Sub CmdExcel_Click()
 On Error Resume Next
 
   ComDlg1.DialogTitle = "–ŒÌ—Â ›«Ì· ê“«—‘"
   ComDlg1.InitDir = App.Path & "\ReportExcel"
   ComDlg1.FileName = ParvaneCode & "K[" & Format(DayDate, "yy-mm-dd") & "]"
   ComDlg1.Filter = "MehrVarzan(*.xls)|*.xls"
   ComDlg1.ShowSave
   
   If InStr(1, ComDlg1.FileName, "\") > 0 Then
      Grid1.ExportToExcel (ComDlg1.FileName)
      MsgBox "›«Ì· ê“«—‘ »« „Ê›ﬁÌ  –ŒÌ—Â ‘œ", vbInformation, ""
   End If
End Sub

Private Sub CmdOpen_Click()
   ComDlg1.DialogTitle = "»«“ò—œ‰ ›«Ì· ê“«—‘"
   ComDlg1.InitDir = App.Path & "\KantinerReport"
   ComDlg1.Filter = "MehrVarzan(*.cel)|*.cel"
   ComDlg1.FileName = ""
   ComDlg1.ShowOpen
   
   If InStr(1, ComDlg1.FileName, "\") > 0 Then
      Grid1.OpenFile (ComDlg1.FileName)
   End If
 

End Sub

Private Sub CmdPrev_Click()
      With Grid1
             .Range(4, 1, 11, 9).FontName = "B Zar"
             .Range(4, 1, 11, 9).FontBold = True
             .Range(4, 1, 11, 9).FontSize = 12
             .Cell(6, 7).Font.Name = "Titr"
             .Cell(6, 7).Font.Size = 10
             .Cell(7, 7).Font.Name = "Titr"
             .Cell(7, 7).Font.Size = 10

     End With
     
     '
  With Grid1.PageSetup
     
     .PaperSize = cellPaperA4  'A4 paper
     .Orientation = cellPortrait  'Portrait
     .PrintTitleRows = 11
     .LeftMargin = 1
     .RightMargin = 1
     .BottomMargin = 2.5
     .TopMargin = 1
     .CenterHorizontally = True  'Center on page horizontally
     '.CenterVertically = True  'Center on page horizontally
     .PrintFixedColumn = False
     .PrintFixedColumn = True
     '.PrintGridlines = True
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



Private Sub CmdSave_Click()
   ComDlg1.DialogTitle = "–ŒÌ—Â ›«Ì· ê“«—‘"
   ComDlg1.InitDir = App.Path & "\KantinerReport"
   
   ComDlg1.FileName = ParvaneCode & "KAN" & Format(DayDate, "yy-mm-dd")

   ComDlg1.Filter = "MehrVarzan(*.cel)|*.cel"
   ComDlg1.ShowSave
   
   If InStr(1, ComDlg1.FileName, "\") > 0 Then
      Grid1.SaveFile (ComDlg1.FileName)
      MsgBox "›«Ì· ê“«—‘ »« „Ê›ﬁÌ  –ŒÌ—Â ‘œ", vbInformation, ""
   End If
End Sub

Private Sub Command1_Click()
  MsgBox "FR " & Grid1.Selection.FirstRow
  MsgBox "FC " & Grid1.Selection.FirstCol
  
  MsgBox "LR " & Grid1.Selection.LastRow
  MsgBox "LC " & Grid1.Selection.LastCol

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
 Dim i As Byte
 Dim ks As String
 Dim MTable, DTable, Tonaj As String
 Dim Weight As Long
 Dim cnt As Integer
 
   MTable = "TabKantiner_Master"
   DTable = "TabKantiner_Detail"
   Tonaj = "TabKantiner_Tonaj"
   '
   Me.Left = (Screen.Width \ 2) - 5000
   
   BackColor = RGB(83, 132, 178)

  With Grid1
       .Cols = 12
       .Rows = 12
       
       Call MakeMasterRows
       Call CmdMakeBorder_Click
       rs.Open "SELECT * FROM " & MTable & " " & _
               "WHERE Parvane='" & ParvaneCode & "'", CNS
       
       .Cell(1, 3).Text = "ò‘ Ì :" & Space(1) & Trim(rs("Keshti"))
       .Cell(1, 1).Text = "»«—‰«„Â :" & Space(5) & Trim(rs("Barname"))
       .Cell(1, 1).Font.Name = "B Zar"
       .Cell(1, 1).Font.Size = 12
       
       .Cell(4, 7).Text = Trim(rs("Typekala"))
       .Cell(5, 7).Text = Trim(rs("Parvane"))
       .Cell(6, 7).Text = Trim(rs("DKootaj")) & Space(13 - Len(Trim(rs("Kootaj")))) & Trim(rs("Kootaj"))
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
       .Cell(7, 1).Text = Val(rs("Bandel"))
       .Cell(8, 1).Text = Trim(rs("Tarkhiskar"))
       
       '
       .Cell(10, 8).Text = Trim(rs("Saheb"))
       .Cell(10, 4).Text = Trim(rs("Etebar"))
       .Cell(10, 1).Text = Trim(rs("Gharardad"))
       rs.Close
       ''''''''
  
       
      rs.Open "SELECT SUM(Weight) FROM " & DTable & " " & _
              "WHERE Parvane='" & ParvaneCode & "' ", CNS

      Weight = rs(0)
      rs.Close
      '
       Dim strSQL As String
       If GetPrintDate <> Empty Then ' From Print Be Tafkike Tarikh
          strSQL = "SELECT * FROM " & DTable & " " & "WHERE Parvane='" & ParvaneCode & "' " & _
                   "AND (" & GetPrintDate & ") ORDER BY Count0"
       Else
          strSQL = "SELECT * FROM " & DTable & " " & "WHERE Parvane='" & ParvaneCode & "' " & _
                   "AND BarNameDate='" & DayDate & "' " & "ORDER BY Count0,BarNameDate"
       End If
      
      '
      rs.Open strSQL, CNS
       cnt = 0
       While Not rs.EOF
           cnt = cnt + 1
           .AddItem rs(12) & vbTab & rs(11) & vbTab & rs(6) & _
                      vbTab & rs(10) & vbTab & rs(8) & vbTab & rs(7) _
                     & vbTab & rs(5) & vbTab & rs(4) & vbTab & rs(3) _
                     & vbTab & rs(2) & vbTab & cnt
           rs.MoveNext

       Wend
       rs.Close
       '
       .Range(12, 11, .Rows - 1, 11).BackColor = &HE0E0E0
       .Range(11, 1, 11, 11).BackColor = &HE0E0E0
 '''
       .Range(12, 1, .Rows - 1, 11).Alignment = cellRightCenter
       .Range(12, 1, .Rows - 1, 11).FontName = "Titr"
       .Range(12, 1, .Rows - 1, 11).FontBold = True
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
       ''''Kantiner And Mobile
       .Range(12, 3, .Rows - 1, 3).FontName = "B Titr"
       .Range(12, 3, .Rows - 1, 3).FontBold = True
       .Range(12, 3, .Rows - 1, 3).FontSize = 9
       '
       .Range(12, 1, .Rows - 1, 1).FontName = "B Titr"
       .Range(12, 1, .Rows - 1, 1).FontSize = 7
       
        '''Insert Bottom Rows Information''''''''''''''''''
        ''Load Ahan Tonaj
       If GetPrintDate <> Empty Then ' From Print Be Tafkike Tarikh
          Dim ad As String
          ad = Mid(GetPrintDate, InStr(InStr(GetPrintDate, "AND"), GetPrintDate, "'") + 1, 8)

          strSQL = "SELECT * FROM " & Tonaj & " " & _
                   "WHERE Parvane='" & ParvaneCode & "' " & _
                   "AND BarDate='" & ad & "'"
       Else
          strSQL = "SELECT * FROM " & Tonaj & " " & _
                   "WHERE Parvane='" & ParvaneCode & "' " & _
                   "AND BarDate='" & DayDate & "' "
       End If
       '
       rs.Open strSQL, CNS

       .AddItem ""
       .Range(.Rows - 1, 10, .Rows - 1, 11).Merge
       .Cell(.Rows - 1, 10).Text = " ‰«é Œ—ÊÃÌ —Ê“«‰Â"
       .Range(.Rows - 1, 8, .Rows - 1, 9).Merge
       .Cell(.Rows - 1, 8).Text = rs("TonajEx")
       '
       '.Range(.Rows - 1, 6, .Rows - 1, 7).Merge
       .Cell(.Rows - 1, 7).Text = " ò·  ‰«é Œ—ÊÃÌ"
       '
       .Range(.Rows - 1, 4, .Rows - 1, 6).Merge
       .Cell(.Rows - 1, 4).Text = Weight
       '
       .Range(.Rows - 1, 2, .Rows - 1, 3).Merge
       '
       If Finish Then
          If rs("TonajMod") < 0 Then
            .Cell(.Rows - 1, 2).Text = "«÷«›Â Ê“‰"
            .Cell(.Rows - 1, 1).Text = Abs(rs("TonajMod"))
          ElseIf rs("TonajMod") > 0 Then
            .Cell(.Rows - 1, 2).Text = "ò”— Ê“‰"
            .Cell(.Rows - 1, 1).Text = rs("TonajMod")
          Else
            .Cell(.Rows - 1, 2).Text = "Å«Ì«Å«Ì"
            .Cell(.Rows - 1, 1).Text = rs("TonajMod")
          End If
       Else
            .Cell(.Rows - 1, 2).Text = " ‰«é »«ﬁÌ„«‰œÂ"
            .Cell(.Rows - 1, 1).Text = rs("TonajMod")
       End If
       ''''
       .AddItem "" ''''''
       .Range(.Rows - 1, 10, .Rows - 1, 11).Merge
       .Cell(.Rows - 1, 10).Text = " ⁄œ«œ ò«‰ Ì‰—"
       .Range(.Rows - 1, 8, .Rows - 1, 9).Merge
       .Cell(.Rows - 1, 8).Text = rs("TotalBandel")
       
       '
       '.Range(.Rows - 1, 6, .Rows - 1, 7).Merge
       .Cell(.Rows - 1, 7).Text = "Ã„⁄ ò· ‘«ŒÂ"
       '
       .Range(.Rows - 1, 4, .Rows - 1, 6).Merge
       .Cell(.Rows - 1, 4).Text = rs("TotalShakhe")
       '
       .Range(.Rows - 1, 2, .Rows - 1, 3).Merge
       .Cell(.Rows - 1, 2).Text = "ò· ò—«ÌÂ"
       '
       .Cell(.Rows - 1, 1).Text = rs("TotalPrice")
       
       rs.Close
       
       '''''''''''''''''''''''''''''''''''''''''''''''''''''
       .RowHeight(.Rows - 1) = 40
       .RowHeight(.Rows - 2) = 40
       '
       .Range(.Rows - 2, 1, .Rows - 1, 11).Alignment = cellCenterCenter
       .Range(.Rows - 2, 1, .Rows - 1, 11).FontName = "B Titr"
       .Range(.Rows - 2, 1, .Rows - 1, 11).FontBold = True
       .Range(.Rows - 2, 1, .Rows - 1, 11).FontSize = 9
       '
       .Range(.Rows - 2, 1, .Rows - 1, 11).Borders(cellInsideHorizontal) = cellThick
       .Range(.Rows - 2, 1, .Rows - 1, 11).Borders(cellInsideVertical) = cellThick
       .Range(.Rows - 2, 1, .Rows - 1, 11).Borders(cellEdgeRight) = cellThick
       .Range(.Rows - 2, 1, .Rows - 1, 11).Borders(cellEdgeLeft) = cellThick
       .Range(.Rows - 2, 1, .Rows - 1, 11).Borders(cellEdgeBottom) = cellThick
       '''''''''''''

  End With
        Call CmdPrev_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  Call CmdSave_Click
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  
  With Grid1.ActiveCell
     If KeyCode = vbKeyRight And .Col = 6 Then
        KeyCode = 0
        Exit Sub
     End If
     
     If KeyCode = vbKeyLeft And .Col = 1 Then
        KeyCode = 0
        Exit Sub
     End If
     
     If KeyCode = vbKeyLeft And .Col = 1 And Val(Grid1.Cell(.Row, 1).Text) > 0 Then
        KeyCode = 0
        Exit Sub
     End If
    
     If KeyCode = 13 Then 'Cells Repeat
        If (.Row = Grid1.Rows - 1) And (.Col = 1 Or .Col = 1) Then
           Grid1.AddItem ""
           Grid1.Cell(Grid1.Rows - 1, 7).Text = Grid1.Rows - 1
           Grid1.Cell(Grid1.Rows - 1, 6).Text = Grid1.Cell(Grid1.Rows - 2, 6).Text
           Grid1.Cell(Grid1.Rows - 1, 5).Text = Val(Grid1.Cell(Grid1.Rows - 2, 5).Text) + 1
           Grid1.Cell(Grid1.Rows - 1, 3).Text = "IRSU"
           Grid1.Cell(Grid1.Rows - 1, 1).Text = "09"
           Grid1.Cell(Grid1.Rows - 1, 6).SetFocus
           KeyCode = 0
           Exit Sub
        End If
     
        If .Col = 1 And .Row <> Grid1.Rows - 1 Then
            
            If Val(Grid1.Cell(.Row, 1).Text) > 0 Then
               Grid1.Cell(.Row + 1, 6).SetFocus
             Else
               Grid1.Cell(.Row, 1).SetFocus
             End If
             
            KeyCode = 0
            Exit Sub
        End If
        
        If .Col = 1 And .Row <> Grid1.Rows - 1 Then
            Grid1.Cell(.Row + 1, 6).SetFocus
            KeyCode = 0
            Exit Sub
        End If
    
        Grid1.Cell(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col - 1).SetFocus
        KeyCode = 0
     End If
    
    If KeyCode = vbKeyEnd Then
          KeyCode = 0
          Exit Sub
    End If
 End With

End Sub

Sub InsertRowForPrintPreview(Row As Integer)
    With Grid1
        .InsertRow Row, 1
        .RowHeight(Row) = 60
        .Cell(Row, 1).WrapText = True
        .Cell(Row, 2).WrapText = True
        '
        .Cell(Row, 3).Font.Name = "B Zar"
        .Cell(Row, 3).Font.Bold = True
        .Cell(Row, 3).Font.Size = 12
        '
        .Column(1).Width = 95 'Mobile
        .Column(2).Width = 90 'Total
        .Column(3).Width = 115 'Kantiner
        .Column(4).Width = 110 'Kamioon
        .Column(5).Width = 80 'barname
        .Column(6).Width = 70 'Date
        .Column(7).Width = 50 'radif
        
        .Cell(Row, 1).Text = "‘„«—Â  „«” —«‰‰œÂ"
        .Cell(Row, 2).Text = "ò· ò—«ÌÂ »Â —Ì«·"
        .Cell(Row, 3).Text = "‘„«—Â ò‹‹«‰ Ì‰‹‹—"
        .Cell(Row, 4).Text = "‘„«—Â ò«„Ì‹‹‹Ê‰"
        .Cell(Row, 5).Text = "‘„«—Â »«—‰«„Â"
        .Cell(Row, 6).Text = " «—ÌŒ"
        .Cell(Row, 7).Text = "—œÌ›"
    
    End With
    
End Sub

Private Sub LblTotal_Click()

End Sub

Private Sub MnuOpen_Click()
  If Dir("c:\myCell.cel") <> Empty Then 'if File Found
     Grid1.OpenFile "c:\myCell.cel"
  Else
     MsgBox "›«Ì· ÊÃÊœ ‰œ«—œ", vbExclamation, ""
  End If
End Sub

Private Sub mnuSave2Cel_Click()
   Grid1.SaveFile "C:\myCell.cel"
   MsgBox "œ— œ—«ÌÊ C –ŒÌ—Â ‘œ"
End Sub

Private Sub mnuSave2Exc_Click()
   Grid1.ExportToExcel "C:\myCell.xls", True, True
   MsgBox "œ— œ—«ÌÊ C –ŒÌ—Â ‘œ"
End Sub

Sub MakeMasterRows()
  Dim i As Integer
  
   With Grid1
'''''''''''''''''''''''''''''''''''''
       .Column(1).Width = 60 'Mobile
       .Column(2).Width = 65 'Total
       .Column(3).Width = 120 'Kantiner
       .Column(4).Width = 40 'Size
       .Column(5).Width = 40 'tedad
       .Column(6).Width = 40 'Weight
       .Column(7).Width = 85 'Kamioon
       .Column(8).Width = 90 'Anbar
       .Column(9).Width = 60 'Date
       .Column(10).Width = 55 'Barname
       .Column(11).Width = 40 'Radif
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
       .Cell(11, 10).WrapText = True
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
       .Range(1, 7, 2, 11).FontName = "Titr"
       .Range(1, 7, 2, 11).FontBold = True
       .Range(1, 7, 2, 11).FontSize = 14
    
       ''''''''''''''
       .Range(11, 1, 11, 11).FontName = "Titr"
       .Range(11, 1, 11, 11).FontBold = True
       .Range(11, 1, 11, 11).FontSize = 8
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
       .Cell(7, 3).Text = " ⁄œ«œ ò«‰ Ì‰—"
       .Cell(8, 3).Text = "‰«„  —ŒÌ’ ò«—"
       ''''''''''''''''
       .Cell(4, 9).Text = "‰‹‹Ê⁄ ò‹‹‹«·«"
       .Cell(5, 9).Text = "‘„‹‹«—Â Å‹—Ê«‰Â"
       
       .Cell(6, 9).Text = "‘„«—Â ò‹Ê ‹«é"
       .Cell(7, 9).Text = "‘„«—Â ﬁ»÷ «‰»«—"
       .Cell(8, 9).Text = " ‹«—ÌŒ Å‹‹—Ê«‰Â"
       '
       .Cell(11, 11).Text = "—œÌ›"
       .Cell(11, 10).Text = "‘„«—Â »«—‰«„Â"
       .Cell(11, 9).Text = " «—ÌŒ"
       .Cell(11, 8).Text = "«‰»«—  Œ·ÌÂ"
       .Cell(11, 7).Text = "‘„«—Â ò«„ÌÊ‰"
       .Cell(11, 6).Text = "Ê“‰"
       .Cell(11, 5).Text = " ⁄œ«œ"
       .Cell(11, 4).Text = "”«Ì“"
       .Cell(11, 3).Text = "‘„«—Â ò«‰ Ì‰—"
       .Cell(11, 2).Text = "ò· ò—«ÌÂ"
       .Cell(11, 1).Text = "„Ê»«Ì·"

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
          .Cell(1, 7).Text = "·Ì”  »«—êÌ—Ì ‘—ò  Õ„· Ê ‰ﬁ· „Â—Ê—“«‰  —«»—" 'List
          .Cell(1, 7).Font.Name = "B Titr"
          .Cell(1, 7).Font.Size = 12

       '
       Dim inp As String
        inp = InputBox("", ":ﬁ«»·  ÊÃÂ ”—Ê— ê—«„Ì ")
        If inp <> Empty Then
           .Cell(3, 7).Text = " ﬁ«»·  ÊÃÂ ”—Ê— ê—«„Ì: " & inp 'Sarvar
        Else
           .Cell(3, 7).Text = ": ﬁ«»·  ÊÃÂ ”—Ê— ê—«„Ì" 'Sarvar
        End If
       ''''''''''''''''''''''''
       .Cell(10, 10).Text = "’«Õ» ò«·«" 'Saheb KALA
       .Cell(10, 7).Text = "‘„«—Â «⁄ »«—" 'Etebar
       .Cell(10, 2).Text = "‘„«—Â ﬁ—«—œ«œ" 'Etebar
       ''''
       .Cell(4, 1).Font.Name = "Tahoma"
       .Cell(4, 1).Font.Bold = True
       .Cell(4, 1).Font.Size = 9

   End With

End Sub


