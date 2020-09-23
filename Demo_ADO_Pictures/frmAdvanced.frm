VERSION 5.00
Begin VB.Form frmAdvanced 
   Caption         =   "ShortcutBar - Advanced"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000005&
      Height          =   400
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   10005
      TabIndex        =   2
      Top             =   0
      Width           =   10065
   End
   Begin VB.PictureBox picWorking 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   255
      ScaleHeight     =   450
      ScaleWidth      =   510
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   510
   End
   Begin Project1.LynxGrid LynxGrid1 
      Align           =   1  'Align Top
      Height          =   4530
      Left            =   0
      TabIndex        =   0
      Top             =   405
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   7990
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorSel    =   12563634
      ForeColorSel    =   0
      CustomColorFrom =   12640511
      CustomColorTo   =   8438015
      GridColor       =   11246491
      BorderStyle     =   0
      FocusRectMode   =   2
      FocusRectColor  =   4406585
      ThemeColor      =   5
      ThemeStyle      =   5
      ScrollBarStyle  =   1
      ShowRowNumbers  =   -1  'True
      ShowRowNumbersVary=   -1  'True
      AllowColumnResizing=   -1  'True
      AllowRowResizing=   -1  'True
      ColumnDrag      =   -1  'True
      ColumnSort      =   -1  'True
      Editable        =   -1  'True
   End
End
Attribute VB_Name = "frmAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim lngI    As Long
  Dim lngR    As Long
   
   Call OpenDB(gdbSCB, App.Path & "\ShortcutBar.mdb")
   
   Call OpenRS(grsSCB, "SELECT TabInfo.TabName, ButtonInfo.*" & _
      " FROM ButtonInfo INNER JOIN TabInfo ON ButtonInfo.TabIndex = TabInfo.TabIndex" & _
      " ORDER BY ButtonInfo.TabIndex, ButtonInfo.ButtonIndex;", gdbSCB)
   
   If ADORecordCount(grsSCB) Then
      With LynxGrid1
         .MinRowHeight = 500
         .AddColumn "TabName", , , , , , , , , , True
         .AddColumn "TabIndex", , lgAlignCenterCenter, lgNumeric
         .AddColumn "ButtonIndex", , lgAlignCenterCenter, lgNumeric
         .AddColumn "ButtonName"
         .AddColumn "ButtonPath"
         .AddColumn "Picture", , , , , , , , , , True
         .ColImageAlignment(5) = lgAlignCenterCenter
         
         Do
            lngR = .AddItem(grsSCB.Fields(0) & vbTab & _
                     grsSCB.Fields(1) & vbTab & _
                     grsSCB.Fields(2) & vbTab & _
                     grsSCB.Fields(3) & vbTab & _
                     grsSCB.Fields(4)) ' & vbTab & "XX")
                     
            Call GetPicFromDB(grsSCB, "ButtonPic")
            Call .CellPicture(picWorking.Picture, lngR, 5)
            picWorking.Picture = Nothing
                     
            grsSCB.MoveNext
         Loop Until grsSCB.EOF
         
         .RowHeight(5) = 100
         Call ResizeGridColumns
      End With
   End If
   
   grsSCB.Close
   gdbSCB.Close
   
End Sub

Private Sub ResizeGridColumns()
  
  Dim lngI As Long
   
   On Error Resume Next
   
   With LynxGrid1
      .Redraw = False
      
      .Height = Me.ScaleHeight - Picture1.Height
      
      '// resize columns
      lngI = (.VisibleWidth - 1400) '// (width of grid - scroll bar width - ShowRowNumbers - ColWidth(5))
      .ColWidth(0) = lngI * 0.25
      .ColWidth(1) = lngI * 0.05
      .ColWidth(2) = lngI * 0.05
      .ColWidth(3) = lngI * 0.3
      .ColWidth(4) = lngI * 0.35
      .ColWidth(5) = 1400
      
      .Redraw = True
   End With

End Sub

Private Sub Form_Resize()
  
  Call ResizeGridColumns

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Set frmAdvanced = Nothing

End Sub

Private Sub LynxGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long, vNewValue As String, Cancel As Boolean)
   
   Call OpenDB(gdbSCB, App.Path & "\ShortcutBar.mdb")
   Call OpenRS(grsSCB, "Select * From ButtonInfo" & _
               " Where ButtonInfo.TabIndex=" & LynxGrid1.CellText(Row, 1) & _
               " AND ButtonInfo.ButtonIndex=" & LynxGrid1.CellText(Row, 2) & _
               " AND ButtonInfo.ButtonName='" & LynxGrid1.CellText(Row, 3) & "'" & _
               " AND ButtonInfo.ButtonPath='" & LynxGrid1.CellText(Row, 4) & "'", gdbSCB)
   
   Select Case Col
   Case 1 '// Tab Change
      If vNewValue < 0 Or vNewValue > LynxGrid1.CellValue(LynxGrid1.Rows - 1, 1) Then
         Cancel = True
      Else
         grsSCB.Fields("TabIndex") = vNewValue
         grsSCB.Update
         LynxGrid1.CellText(Row, 0) = gdbSCB.Execute("Select TabInfo.TabName as TabDesc From TabInfo" & _
            " Where TabInfo.TabIndex=" & CStr(vNewValue))("TabDesc")
      End If
      
   Case 2 '// Button Order Changed
      grsSCB.Fields("ButtonIndex") = vNewValue
      grsSCB.Update
   
   Case 3 '// Button Name
      vNewValue = Trim$(vNewValue)
      If LenB(vNewValue) = 0 Then
         Cancel = True
      Else
         grsSCB.Fields("ButtonName") = vNewValue
         grsSCB.Update
      End If
      
   Case 4 '// Button Path
      vNewValue = Trim$(vNewValue)
      If LenB(vNewValue) = 0 Then
         Cancel = True
      Else
         grsSCB.Fields("ButtonPath") = vNewValue
         grsSCB.Update
      End If
   End Select
   
   grsSCB.Close
   gdbSCB.Close
   
End Sub

Private Sub LynxGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   
   Select Case Col
   Case 0
      Cancel = True
   Case 3, 4
      If LynxGrid1.CellText(Row, 3) = "BLANK" Then
         Cancel = True
      End If
   End Select

End Sub

Private Sub LynxGrid1_Click()
   Picture1.Picture = LynxGrid1.CellPictureGet(, 5)
End Sub

