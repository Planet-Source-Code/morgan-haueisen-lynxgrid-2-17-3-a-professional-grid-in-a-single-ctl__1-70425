VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LynxGrid Tester © 2006 Richard Mewett"
   ClientHeight    =   8940
   ClientLeft      =   1215
   ClientTop       =   2445
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin LynxGridTest.LynxGrid LynxGrid1 
      Height          =   6930
      Left            =   30
      TabIndex        =   0
      Top             =   630
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   12224
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorSel    =   12937777
      ForeColorHdr    =   8388608
      ForeColorSel    =   16777215
      CustomColorFrom =   16572875
      CustomColorTo   =   14722429
      GridColor       =   16367254
      BorderStyle     =   0
      FocusRectMode   =   2
      FocusRectColor  =   9895934
      GridLines       =   2
      ThemeStyle      =   4
      ColumnHeaderLines=   2
      Caption         =   "Employees"
      Appearance      =   0
      ScrollBarStyle  =   1
      ShowRowNumbersVary=   -1  'True
      AllowColumnResizing=   -1  'True
      AllowWordWrap   =   -1  'True
      ColumnDrag      =   -1  'True
      ColumnSort      =   -1  'True
      EditMove        =   2
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   60
      Top             =   405
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   555
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483643
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":0000
            Key             =   "MALE1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":059A
            Key             =   "MALE2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":0B34
            Key             =   "MALE3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":10CE
            Key             =   "FEMALE1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":1668
            Key             =   "FEMALE2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":1C02
            Key             =   "FEMALE3"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":219C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":2306
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":2470
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTopBar 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   11895
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print Grid"
         Height          =   465
         Left            =   9015
         TabIndex        =   5
         Top             =   90
         Width           =   2220
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single File solution"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   4
         Top             =   300
         Width           =   1620
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner-drawn editable Grid UserControl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   3
         Top             =   60
         Width           =   3315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   60
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CreateGrid()

   With LynxGrid1
      'Set ImageList to provide Item Images
      .ImageList = ImageList1

      'Create the Columns
      .AddColumn "Code", 1000, , , ">"
      .AddColumn "G", 250
      .AddColumn "Forename", 1500
      .AddColumn "Surname", 1500, , , ">" '// Allow Only UPPERCASE
      .AddColumn "Job Title", 800, , , , , , True, , , True '// This column is locked

      .AddColumn "Pension", 1000, lgAlignCenterCenter, lgBoolean
      .AddColumn "DOB", 1000, lgAlignCenterCenter, lgDate, "mm/dd/yyyy"
      .AddColumn "Premium Dollars and cents", 1600, lgAlignRightCenter, lgNumeric, "$#,#.00"
      .AddColumn "Notes", 5000
      .AddColumn "Button", 800, lgAlignCenterCenter, lgButton

      .ColImageAlignment(2) = lgAlignRightCenter

      .TotalsLineCaption(7) = "Total:"
      .TotalsLineShow = True

   End With

End Sub


Private Sub cmdPrint_Click()
   Call PrintGrid(LynxGrid1)
End Sub

Private Sub Form_Load()

   CreateGrid
   LoadDemoData

End Sub

Private Sub Form_Resize()

   If Not Me.WindowState = vbMinimized Then
      LynxGrid1.Height = Me.ScaleHeight - LynxGrid1.Top
      LynxGrid1.Width = Me.ScaleWidth - LynxGrid1.Left
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frmMain = Nothing

End Sub

Private Sub LoadDemoData()

  Dim lCount As Long
  Dim lRow   As Long
  Dim sForename As String
  Dim sGender   As String

   With LynxGrid1
      .Redraw = False

      'Add some random data

      For lCount = 1 To 50

         'Simple method to specify Gender!

         If RandomInt(0, 1) = 0 Then
            sGender = "M"
            sForename = GetForename(ntMale)
         Else
            sGender = "F"
            sForename = GetForename(ntFemale)
         End If

         '// Add data to grid and return row number
         lRow = .AddItem(Format$("XD" & Format$(.ItemCount, "000")) & vbTab & _
                         sGender & vbTab & sForename & vbTab & _
                         GetSurname() & vbTab & _
                         GetJobName() & vbTab & _
                         (RandomInt(0, 1) = 0) & vbTab & _
                         DateSerial(RandomInt(1930, 1990), RandomInt(1, 12), RandomInt(1, 28)) & vbTab & _
                         Round(100 + (Rnd * 100), 2) & vbTab & _
                         vbTab & _
                         sGender)

         If sGender = "M" Then
            'Set the Key for the ImageList Image (can use text Key or Index)
            .RowImage(lRow) = "MALE" & RandomInt(1, 3)
            .CellForeColor(lRow, 1) = vbBlue

         Else
            .RowImage(lRow) = RandomInt(3, 6)
            .RowForeColor(lRow) = vbRed
            .CellForeColor(lRow, 1) = vbGreen
         End If

         .CellImage(lRow, 2) = RandomInt(7, 9)

      Next lCount

      '// Lock Row #5
      .RowLocked(5) = True
      .CellText(5, 8) = "This Row is Locked" '// value change
      .CellImage(5, 9) = RandomInt(7, 9)

      .CellImage(8, 9) = RandomInt(7, 9)
      '.ColImageAlignment(9) = lgAlignCenterCenter

      'The grid supports per cell formatting but provides Item
      'formatting options for simplicity when only per Row formatting
      'is required (Row formatting reformats all Cells in the Row).
      .RowBackColor(5) = &H95E0F1
      .RowForeColor(5) = &H1F488A

      'Tell the grid to Draw
      .Redraw = True
   End With

End Sub

Private Sub LynxGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

   'Is the Edit allowed?
   Select Case Col
   Case 1 'Gender Column
      Cancel = True
   End Select

End Sub

Private Sub LynxGrid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

   MsgBox "clicked button on row#" & CStr(Row) & ", col#" & CStr(Col)

End Sub

Private Sub LynxGrid1_Click()

   If LynxGrid1.RowLocked(LynxGrid1.Row) Then
      MsgBox "This row is locked"

   ElseIf LynxGrid1.ColLocked(LynxGrid1.Col) Then
      MsgBox "This column is locked"
   End If

End Sub

