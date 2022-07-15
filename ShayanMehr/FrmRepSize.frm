VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form FrmRepSize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›—„ ê“«—‘ ”«Ì“"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "B Zar"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmRepSize.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin FlexCell.Grid Grid1 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   14631
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin PrjShayan.TypeButton CmdPreview 
      Height          =   495
      Left            =   9120
      TabIndex        =   1
      Top             =   8400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   2
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRepSize.frx":29C12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "FrmRepSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Ahan As Byte ' 1 Ahan 2 Kant 3 AEL
Public EnableDesc As Boolean

Private Sub Preview()
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
     .Footer = " »‰œ— «‰“·Ì°€«“Ì«‰° «» œ«Ì ŒÌ«»«‰ —„÷«‰Ì°Ã‰» »«‰ò ’«œ—«  ‘⁄»Â »‰«œ— Ê ò‘ Ì—«‰Ì°òÊçÂ ‘ÂÌœ ”Ì—Ì°”«Œ „«‰  Ã«—Ì »—«œ—«‰ „Ãœ ÅÊ—° ÿ»ﬁÂ œÊ„ " & " E-Mail: mehrvarzantarabar@yahoo.com" & Space(0) & vbCrLf & _
               "„Ê»«Ì· 09126101318-09111813086" & Space(10) & "›«ò” 3239400-0181" & Space(10) & " ·›‰ 4-3239880-0181" & Space(15) & "’›ÕÂ &P"
    '
  End With
  
  Grid1.PrintPreview

End Sub


Private Sub CmdPreview_Click()
 Call Preview
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then Unload Me
End Sub

Private Sub Grid1_Click()
'MsgBox Grid1.ActiveCell.Row
'MsgBox Grid1.ActiveCell.Col
 
End Sub

Private Sub Grid1_LeaveCell(ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
 Dim Myrs As New Recordset
 Dim strSQL As String

  If Col = 1 And EnableDesc And Grid1.Cell(Row, Col).Text <> Empty Then
     strSQL = "UPDATE DefSize SET Description='" & Grid1.Cell(Row, Col).Text & "' "
     strSQL = strSQL & "WHERE Ahan=" & Ahan & " AND "
     strSQL = strSQL & "Parvane='" & Grid1.Cell(5, 7).Text & "' AND "
     strSQL = strSQL & "SizeKala='" & Val(Grid1.Cell(Row, 9).Text) & "'"
     Myrs.Open strSQL, CNS
     '
     Set Myrs = Nothing
  End If
End Sub
