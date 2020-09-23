VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin GroupedRows.LynxGrid LynxGrid1 
      Height          =   5340
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   9419
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorBkg    =   1002598
      BackColorSel    =   3634318
      ForeColorSel    =   14806777
      CustomColorFrom =   10874879
      CustomColorTo   =   4818592
      GridColor       =   6463417
      ProgressBarColor=   16744576
      FocusRectMode   =   1
      FocusRectColor  =   9895934
      ThemeColor      =   4
      ThemeStyle      =   3
      ScrollBars      =   1
      Appearance      =   0
      ScrollBarStyle  =   1
      Editable        =   -1  'True
      FullRowSelect   =   0   'False
      HotHeaderTracking=   0   'False
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   405
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05AA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdbSCB    As ADODB.Connection

Private Sub Form_Load()

  Dim rsNames     As ADODB.Recordset
  Dim rsScore     As ADODB.Recordset
  Dim lRow        As Long
  Dim lHRow       As Long
  Dim lngGroupID  As Long
  Dim lngID       As Long
  Dim dblAverage  As Double

   '// Open database
   Call OpenDB(mdbSCB, App.Path & "\Grades.mdb")
   
   '// Open recordset to NameList table
   Call OpenRS(rsNames, "SELECT NameList.* From NameList", mdbSCB)
   
   With LynxGrid1
      '// Set grid defaults
      .ImageList = ImageList1
      .AllowEdit = True
      .ScrollBarStyle = Style_Regular
      .FocusRectStyle = lgFRMedium
      .FocusRectMode = lgCol
      .FocusRowHighlight = False
      .AllowColumnSort = False
      .BackColorEvenRowsEnabled = False
      
      '// Create the Columns
      .AddColumn "First Name", 100, , , , , , , , , True
      .AddColumn "Last Name", 100, , , , , , , , , True
      .AddColumn "Average", 100, lgAlignRightCenter, lgNumeric, "0.0", , , , , , True
      .AddColumn "Date", 100, , lgDate, , , , , , , True
      .AddColumn "Score", 100, lgAlignCenterCenter, lgProgressBar
      .AddColumn "RowID", , , , , , , , , False
      .AddColumn "sID", , , , , , , , , False
   End With
   
   
   '// Are there names?
   If ADORecordCount(rsNames) Then
   
      With LynxGrid1
      
         Do
            '// Load New Group
            lngID = rsNames.Fields("sID")
            lngGroupID = lngGroupID + 1
            lHRow = .AddItem(rsNames.Fields("sFirstName") & vbTab & rsNames.Fields("sLastName"))
            .RowGroupHeader(lHRow) = True
            .RowImage(lHRow) = 1
            .RowHeight(lHRow) = 30
            .RowBackColor(lHRow) = vbWhite
            .RowData(lHRow) = lngGroupID
            
            Call OpenRS(rsScore, "SELECT Score.* From Score WHERE Score.sID = " & CStr(lngID) & ";", mdbSCB)
            If ADORecordCount(rsScore) Then
               
               '// Get Average score
               dblAverage = MakeNNull(mdbSCB.Execute("SELECT Avg(Score.Score) AS AvgOfScore From Score" & _
                  " WHERE Score.sID = " & CStr(lngID) & ";")("AvgOfScore"))
               .CellValue(lHRow, 2) = dblAverage
               
               Do
                  lRow = .AddItem
                  .RowBackColor(lRow) = &H96FFFE
                  .RowData(lRow) = lngGroupID
                  .RowVisible(lRow) = False
                  .CellText(lRow, 3) = rsScore.Fields("DateTime")
                  .CellProgressValue(lRow, 4) = CInt(rsScore.Fields("Score"))
                  .CellText(lRow, 4) = rsScore.Fields("Score")
                  .CellValue(lRow, 5) = rsScore.Fields("ID")
                  .CellValue(lRow, 6) = rsNames.Fields("sID")
                  rsScore.MoveNext
               Loop Until rsScore.EOF
            End If
            
            rsScore.Close
            rsNames.MoveNext
         Loop Until rsNames.EOF
         
         .ColForceFit
      End With
   
   End If
   
   rsNames.Close
   
   LynxGrid1.Redraw = True
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   On Error Resume Next
   mdbSCB.Close
   
   Set frmMain = Nothing
   
End Sub

Private Sub LynxGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long, vNewValue As String, Cancel As Boolean)
   
  Dim dblAverage  As Double
  Dim lngI        As Long
  
   '// Update database record
   mdbSCB.Execute "UPDATE Score SET Score.Score = " & Val(vNewValue) & " WHERE Score.ID=" & LynxGrid1.CellText(Row, 5)
   
   '// give database time to update record
   DoEvents
   
   '// Find Group Header Row and update average
   lngI = Row - 1
   Do
      If LynxGrid1.RowGroupHeader(lngI) Then
         '// recalculate average
         dblAverage = MakeNNull(mdbSCB.Execute("SELECT Avg(Score.Score) AS AvgOfScore From Score" & _
            " WHERE Score.sID = " & LynxGrid1.CellText(Row, 6) & ";")("AvgOfScore"))
            
         LynxGrid1.CellValue(lngI, 2) = dblAverage
   
         Exit Do
      End If
      lngI = lngI - 1
   Loop Until lngI < 0

End Sub

Private Sub LynxGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   
   '// don't allow edits to group row
   Cancel = LynxGrid1.RowGroupHeader(Row)
   
End Sub

Private Sub LynxGrid1_Click()
  
  Dim blnTemp  As Boolean
  Dim lRow     As Long
  Dim lGroupID As Long
   
   With LynxGrid1
   
      If .RowGroupHeader Then '// a group header?
         '// Stop the grid from Drawing
         .Redraw = False
      
         lGroupID = .RowData() '// get group number of selected row
         blnTemp = (.RowImage() = 1) '// Show or hide? (+/-)
         
         For lRow = 0 To .Rows - 1
            If Not .RowGroupHeader(lRow) Then '// not a group header
               If .RowData(lRow) = lGroupID Then '// member of group?
                  .RowVisible(lRow) = blnTemp
               End If
            End If
         Next lRow
         
         If blnTemp Then '// change group header image (+/-)
            .RowImage() = 2
         Else
            .RowImage() = 1
         End If
         
         '// Tell the grid to Draw
         .Redraw = True
      End If
   
   End With
   
End Sub

