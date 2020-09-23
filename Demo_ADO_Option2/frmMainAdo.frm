VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ADO Demo"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   6345
      TabIndex        =   1
      Top             =   0
      Width           =   6405
      Begin VB.CommandButton Command1 
         Caption         =   "View All"
         Height          =   585
         Index           =   2
         Left            =   4365
         TabIndex        =   6
         Top             =   60
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
         Caption         =   "View Scores Above Average"
         Height          =   585
         Index           =   1
         Left            =   2475
         TabIndex        =   3
         Top             =   45
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
         Caption         =   "View Scores Below Average"
         Height          =   585
         Index           =   0
         Left            =   75
         TabIndex        =   2
         Top             =   60
         Width           =   1725
      End
      Begin VB.Label lblAverage 
         Alignment       =   2  'Center
         Height          =   600
         Left            =   1785
         TabIndex        =   5
         Top             =   45
         Width           =   705
      End
   End
   Begin Project1.LynxGrid LynxGrid1 
      Height          =   4260
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   7514
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorSel    =   65535
      ForeColorSel    =   0
      CustomColorFrom =   16572875
      CustomColorTo   =   14722429
      GridColor       =   16367254
      FocusRectMode   =   2
      FocusRectColor  =   128
      GridLines       =   3
      ThemeColor      =   5
      ColumnHeaderSmall=   -1  'True
      ScrollBarStyle  =   1
      ShowRowNumbersVary=   -1  'True
      AllowColumnResizing=   -1  'True
      ColumnDrag      =   -1  'True
      ColumnSort      =   -1  'True
      Editable        =   -1  'True
      AllowDelete     =   -1  'True
      AllowInsert     =   -1  'True
      FocusRowHighlightStyle=   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "DEMO: Add Rows, Remove Rows, Change Data.  (application created a access DB when first loaded and fills it with data)."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   75
      TabIndex        =   4
      Top             =   5100
      Width           =   6030
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDBName As String

Private Sub Command1_Click(Index As Integer)

  Dim MyDB  As ADODB.Connection
  Dim MySet As ADODB.Recordset
  Dim vAvg  As Double
  
   Call OpenDB(MyDB, mDBName)
   
   vAvg = rVal(MyDB.Execute("Select Avg(TestScores.Score) as AvgScore From TestScores;")("AvgScore"))
   
   If vAvg Then
      lblAverage.Caption = Format$(vAvg, "0")
   
      Select Case Index
      Case 1
         Call OpenRS(MySet, "Select * From TestScores" & _
            " Where TestScores.Score>" & CStr(vAvg) & " AND TestScores.Score > 0" & _
            " Order By TestScores.NameLast", MyDB)
            
         LynxGrid1.TotalsLineShow = False
      
      Case 0
         Call OpenRS(MySet, "Select * From TestScores" & _
            " Where TestScores.Score<" & CStr(vAvg) & " AND TestScores.Score > 0 " & _
            " Order By TestScores.NameLast", MyDB)
         
         LynxGrid1.TotalsLineShow = False
         
      Case Else
         Call OpenRS(MySet, "Select * From TestScores Order By TestScores.NameLast", MyDB)
         
         LynxGrid1.TotalsLineShow = True
      End Select
      
      With LynxGrid1
         .FillGridFromQuery MySet, True '// Define Grid Columns and fill with data
         
         .ColVisible(0) = False
         .TotalsLineColAvg(3) = True
         .TotalsLineCaption(3) = "Avg: "
         .ColType(3) = lgNumeric
         .ColFormat(3) = "#"
         .ColAlignment(3) = lgAlignRightCenter
         .FormatCellsAlignment , , 3, 3, lgAlignRightCenter
         .ColForceFit
      End With
      
   End If
   
   MySet.Close
   MyDB.Close

   LynxGrid1.Redraw = True
   LynxGrid1.SetFocus
   
End Sub

Private Sub Form_Load()

'''// Done in LynxGrid1.FillGridFromQuery
'''   With LynxGrid1
'''      .AddColumn "Key", , , lgNumeric, , , , , , False
'''      .AddColumn "Last Name", 2500
'''      .AddColumn "First Name", 2000
'''      .AddColumn "Score", 1000, lgAlignRightCenter, lgNumeric, "#"
'''
'''      .Redraw = True
'''   End With
   
   mDBName = App.Path
   If Not (Right$(mDBName, 1) = "\") Then mDBName = mDBName & "\"
   mDBName = mDBName & App.Title & ".mdb"
   
   If LenB(Dir$(mDBName)) = 0 Then
      Call CreateMDB(mDBName)
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmMain = Nothing
End Sub

Private Sub LynxGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long, vNewValue As String, Cancel As Boolean)
   
  Dim MyDB  As ADODB.Connection
  
   If Col = 3 Then '// Score
      If rVal(vNewValue) = 0 Then
         MsgBox "A value of zero is not allowed.", vbExclamation
         Cancel = True
      
      Else
         Call OpenDB(MyDB, mDBName)
         MyDB.Execute "UPDATE TestScores SET TestScores.Score = " & vNewValue & _
            " WHERE TestScores.Key=" & LynxGrid1.CellText(Row, 0)
         MyDB.Close
      
      End If
   
   Else
      vNewValue = Trim$(vNewValue)
      If LenB(vNewValue) = 0 Then
         MsgBox "Blank names are not allowed.", vbExclamation
         Cancel = True
      
      Else
         Call OpenDB(MyDB, mDBName)
         If Col = 1 Then
            MyDB.Execute "UPDATE TestScores SET TestScores.NameLast = '" & vNewValue & _
               "' WHERE TestScores.Key=" & LynxGrid1.CellText(Row, 0)
         Else
            MyDB.Execute "UPDATE TestScores SET TestScores.NameFirst = '" & vNewValue & _
               "' WHERE TestScores.Key=" & LynxGrid1.CellText(Row, 0)
         End If
         MyDB.Close
      
      End If
   
   End If
   
End Sub

Private Sub LynxGrid1_AfterInsert(ByVal Row As Long)
  
  Dim MyDB  As ADODB.Connection
  Dim MySet As ADODB.Recordset
  Dim lngI  As Long
  
   Call OpenDB(MyDB, mDBName)
   Call OpenRS(MySet, "Select * From TestScores", MyDB)
   
   With MySet
      .AddNew
      .Fields("NameLast") = "_LastName"
      .Fields("NameFirst") = "_FirstName"
      .Fields("Score") = 0
      .Update
      DoEvents
      .MoveLast
      lngI = .Fields("Key")
   End With
   MySet.Close
   MyDB.Close
   
   With LynxGrid1 '// Add default data to prevent blank cells
      .CellValue(Row, 0) = lngI
      .CellText(Row, 1) = "_LastName"
      .CellText(Row, 2) = "_FirstName"
      .CellValue(Row, 3) = 100
      .ForceCellEdit Row, 3, True
   End With
  
End Sub

Private Sub LynxGrid1_BeforeDelete(ByVal Row As Long, Cancel As Boolean)
  
  Dim MyDB As ADODB.Connection
   
   If MsgBox("Are you sure you want to delete the entire row?", vbQuestion + vbYesNo) = vbYes Then
      Call OpenDB(MyDB, mDBName)
      MyDB.Execute "DELETE TestScores.*, TestScores.Key From TestScores WHERE TestScores.Key=" & LynxGrid1.CellText(Row, 0)
      MyDB.Close
      LynxGrid1.RemoveRow Row
   End If
   Cancel = True
   
End Sub
