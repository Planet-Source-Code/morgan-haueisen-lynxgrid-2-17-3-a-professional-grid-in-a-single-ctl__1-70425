VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5580
   StartUpPosition =   2  'CenterScreen
   Begin GroupedRows.LynxGrid LynxGrid1 
      Height          =   3975
      Left            =   180
      TabIndex        =   1
      Top             =   405
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   7011
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorSel    =   12937777
      ForeColorSel    =   16777215
      CustomColorFrom =   16572875
      CustomColorTo   =   14722429
      GridColor       =   16367254
      FocusRectMode   =   2
      FocusRectColor  =   9895934
      FocusRectStyle  =   2
      ThemeStyle      =   2
      ScrollBars      =   1
      Appearance      =   0
      ScrollBarStyle  =   1
      ShowRowNumbersVary=   -1  'True
      ColumnSort      =   -1  'True
      Editable        =   -1  'True
      FullRowSelect   =   0   'False
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   45
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
   Begin VB.Label Label1 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   3870
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  
  Dim lCount      As Long
  Dim lRow        As Long
  Dim lHRow       As Long
  Dim lTemp       As Long
  Dim lAvgCnt     As Long
  Dim lAvgTol     As Long
  Dim lngGroupID  As Long

   With LynxGrid1
      '// Set grid defaults
      .ImageList = ImageList1
      .AllowEdit = True
      .ScrollBarStyle = Style_Regular
      .FocusRectStyle = lgFRMedium
      .FocusRectMode = lgCol
      .FocusRowHighlight = False
      .AllowColumnSort = True
      .BackColorEvenRowsEnabled = False
      
      '// Create the Columns
      lTemp = (.VisibleWidth - 2000) \ 2
      .AddColumn "Forename", lTemp
      .AddColumn "Surname", lTemp
      .AddColumn "Present", 500, lgAlignCenterCenter, lgBoolean
      .AddColumn "OK", 500, lgAlignCenterCenter, lgButton
      .AddColumn "Progress", 1000, lgAlignCenterCenter, lgProgressBar, , , , , , , True
      
      '// Load Group #1
      lngGroupID = 1
      lHRow = .AddItem("GIRLS") '// Add Row
      .RowData(lHRow) = lngGroupID
      .RowGroupHeader(lHRow) = True
      .RowImage(lHRow) = 1
      .RowHeight(lHRow) = 30
      .RowBackColor(lHRow) = &HC5D2FF
      
      lAvgCnt = RandomInt(4, 6)
      For lCount = 1 To lAvgCnt '// Add Row Data and make it invisible
         lRow = .AddItem(GetForename(ntFemale) & vbTab & GetSurname(), , , , , False)
         .RowData(lRow) = lngGroupID
         .RowBackColor(lRow) = &HDCE9FF
         lTemp = RandomInt(20, 90)
         .CellProgressValue(lRow, 4) = lTemp
         .CellText(lRow, 4) = lTemp & "%"
         lAvgTol = lAvgTol + lTemp
      Next lCount
      .CellText(lHRow, 4) = "Avg: " & Format$(lAvgTol / lAvgCnt, "0") & "%"
      
      lAvgTol = 0

      '// Load Group #2
      lngGroupID = 2
      lHRow = .AddItem("BOYS") '// Add Row
      .RowData(lHRow) = lngGroupID
      .RowGroupHeader(lHRow) = True
      .RowImage(lHRow) = 1
      .RowHeight(lHRow) = 30
      .RowBackColor(lHRow) = &HFFD3C5
      
      lAvgCnt = RandomInt(5, 7)
      For lCount = 1 To lAvgCnt '// Add Row Data and make it invisible
         lRow = .AddItem(GetForename(ntMale) & vbTab & GetSurname(), , , , , False)
         .RowData(lRow) = lngGroupID
         .RowBackColor(lRow) = &HFFE4D6
         lTemp = RandomInt(20, 90)
         .CellProgressValue(lRow, 4) = lTemp
         .CellText(lRow, 4) = lTemp & "%"
         lAvgTol = lAvgTol + lTemp
      Next lCount
      .CellText(lHRow, 4) = "Avg: " & Format$(lAvgTol / lAvgCnt, "0") & "%"

      '// Tell the grid to Draw
      .Redraw = True
   End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmMain = Nothing
End Sub

Private Sub LynxGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   
   '// don't allow group description to be changed
   Cancel = LynxGrid1.RowGroupHeader(Row)

End Sub

Private Sub LynxGrid1_Click()
  
  Dim blnTemp As Boolean
  Dim lRow    As Long
  Dim lGroupID As Long
   
   With LynxGrid1
      '// Stop the grid from Drawing
      .Redraw = False
      
      If Not .RowGroupHeader Then '// Not a group header then return name
         Label1.Caption = .CellText(, 0) & " " & .CellText(, 1)
      
      Else
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
      End If
      
      .ColWidth(0) = (.VisibleWidth - 2000) \ 2
      .ColWidth(1) = (.VisibleWidth - 2000) \ 2
      
      '// Tell the grid to Draw
      .Redraw = True
   End With
   
End Sub

