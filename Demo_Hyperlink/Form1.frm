VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hyperlink Example"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin Project1.LynxGrid grdProfile 
      Height          =   7350
      Left            =   270
      TabIndex        =   0
      Top             =   825
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   12965
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorBkg    =   15527137
      BackColorSel    =   13421748
      ForeColor       =   0
      ForeColorHdr    =   8618868
      ForeColorSel    =   4210752
      BackColorEvenRows=   15527137
      BackColorEvenRowsEnabled=   0   'False
      CustomColorFrom =   15527137
      CustomColorTo   =   13421748
      GridColor       =   13421748
      AlphaBlendSelection=   -1  'True
      BorderStyle     =   0
      FocusRectColor  =   9895934
      ThemeColor      =   5
      ThemeStyle      =   5
      ShowColumnHeaders=   0   'False
      ScrollBars      =   3
      Caption         =   "Profile"
      CaptionAlignment=   4
      ScrollBarStyle  =   1
      FocusRowHighlightKeepTextForecolor=   -1  'True
      ShowRowNumbersVary=   -1  'True
      EditMove        =   2
      FocusRowHighlightStyle=   1
      HotHeaderTracking=   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " Control Editing through code instead of using the AllowEdit Properity "
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
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   435
      Width           =   7500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " Ctrl+Click on E-mail or Single Click on Web to activate Hyperlink "
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
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   90
      Width           =   7005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" _
      Alias "ShellExecuteA" ( _
      ByVal hWnd As Long, _
      ByVal lpOperation As String, _
      ByVal lpFile As String, _
      ByVal lpParameters As String, _
      ByVal lpDirectory As String, _
      ByVal nShowCmd As Long) As Long

Private Sub Form_Load()

   '// Create Grid
   With grdProfile
      .AddColumn "Type", 2, , , , , , True, , , True
      .AddColumn "Data", 6, , , , , , True
         
      '// Add fixed row names to column #0
      .AddItem "Name"
      .AddItem "Show As"
      .AddItem "Groups"
      .AddItem "Setting"
      .AddItem "Company"
      .AddItem "Job Title"
      .AddItem "Address: Home"
      .AddItem "Address: Business"
      .AddItem "Address: Bill To"
      .AddItem "Address: Ship To"
      .AddItem "Phone: Home"
      .AddItem "Phone: Business"
      .AddItem "Phone: Mobile"
      .AddItem "Phone: Other"
      .AddItem "Phone: Fax"
      .AddItem "E-Mail: Home"
      .AddItem "E-Mail: Business"
      .AddItem "E-Mail: Other"
      .AddItem "Web (Opt#1)"
      .AddItem "Web (Opt#2)"
      .AddItem "ChangeLog"
      
      .Redraw = True
      .FormatCells , , 0, 0, lgCFBackColor, &HECECE1
      .FormatCells 1, 1, 1, 1, lgCFBackColor, &HECECE1
      .FormatCells 15, , 1, 1, lgCFForeColor, vbBlue
      .FormatCells 15, , 1, 1, lgCFFontUnderline, True
      .FormatCells 15, , 1, 1, lgCFHandPointer, True
      
      '// Add data to column #1
      .CellText(0, 1) = "Morgan Haueisen"
      .CellText(1, 1) = "Haueisen, Morgan"
      .CellText(2, 1) = "Grid Control"
      .CellText(3, 1) = "Default"
      .CellText(4, 1) = App.CompanyName
      .CellText(5, 1) = "None"
      '.CellText(6, 1) = ""
      '.CellText(7, 1) = ""
      '.CellText(8, 1) = ""
      '.CellText(9, 1) = ""
      .CellText(10, 1) = "111-111-1111"
      .CellText(11, 1) = "311-234-4567"
      '.CellText(12, 1) = ""
      '.CellText(13, 1) = ""
      '.CellText(14, 1) = ""
      .CellText(15, 1) = "MyEmail@yahoo.com"
      '.CellText(16, 1) = ""
      '.CellText(17, 1) = ""
      .CellText(18, 1) = "www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=70425&lngWId=1"
      
      .CellText(19, 1) = "LynxGrid on PSC"
      .CellTag(19, 1) = "www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=70425&lngWId=1"
      
      .CellText(20, 1) = "HistoryLog.htm"
      .CellTag(20, 1) = App.Path & "\..\HistoryLog.htm"
   
      .ColForceFit
      .Redraw = True
   End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Form1 = Nothing
End Sub

Private Sub grdProfile_AfterEdit(ByVal Row As Long, ByVal Col As Long, vNewValue As String, Cancel As Boolean)
   
   If Row = 0 Then grdProfile.CellText(1, 1) = vNewValue

End Sub

Private Sub grdProfile_CellHandClick(ByVal Row As Long, ByVal Col As Long, Shift As Integer)

  Dim lngR     As Long
  Dim strTemp  As String
  
   '// A Cell with the hand pointer visible has been clicked
   
   With grdProfile
      lngR = grdProfile.Row
      
      Select Case lngR
      Case 19, 20 '// Web page Option#2 and ChangeLog
         strTemp = grdProfile.CellTag(lngR, 1)
         If LenB(strTemp) Then
            Screen.MousePointer = vbHourglass
            DoEvents
            
            If InStr(1, strTemp, "www.") Then
               If InStr(1, strTemp, "://") = 0 Then
                  strTemp = "http://" & strTemp
               End If
            End If
            
            Call ShellExecute(Me.hWnd, "open", strTemp, vbNullString, "C:\", 5)
            Screen.MousePointer = vbDefault
         End If
      
      Case 18 '// Web page Option#1
         strTemp = grdProfile.CellText(lngR, 1)
         If LenB(strTemp) Then
            Screen.MousePointer = vbHourglass
            DoEvents
            
            If InStr(1, strTemp, "www.") Then
               If InStr(1, strTemp, "://") = 0 Then
                  strTemp = "http://" & strTemp
               End If
            End If
            
            Call ShellExecute(Me.hWnd, "open", strTemp, vbNullString, "C:\", 5)
            Screen.MousePointer = vbDefault
         End If
      End Select
   End With

End Sub

Private Sub grdProfile_CellClick(ByVal Row As Long, ByVal Col As Long, Shift As Integer)

'''   grdProfile
'''   0  "Name"
'''   1  "Show As"
'''   2  "Groups"
'''   3  "Setting"
'''   4  "Company"
'''   5  "Job Title"
'''   6  "Address: Home"
'''   7  "Address: Business"
'''   8  "Address: Bill To"
'''   9  "Address: Ship To"
'''  10  "Phone: Home"
'''  11  "Phone: Business"
'''  12  "Phone: Mobile"
'''  13  "Phone: Other"
'''  14  "Phone: Fax"
'''  15  "E-Mail: Home"
'''  16  "E-Mail: Business"
'''  17  "E-Mail: Other"
'''  18  "Web Page"
'''  19  "Web Page"
'''  20  "ChangeLog"

  Dim lngR  As Long
  Dim strTemp  As String
  
  With grdProfile
      lngR = grdProfile.Row
      
      '// Launch Hyperlink to e-mail if Ctrl+Click
      If (Shift And vbCtrlMask) Then
         '// Since these cells also have the Hand pointer visible this could be moved to CellHandClick
         Select Case lngR
         Case 15 To 17
            strTemp = .CellText(Row, Col)
            If LenB(strTemp) Then
               Screen.MousePointer = vbHourglass
               DoEvents
               Call ShellExecute(0&, vbNullString, "mailto:" & strTemp, vbNullString, vbNullString, vbNormalFocus)
               Screen.MousePointer = vbDefault
            End If
         End Select
      
      Else
         '// Edit with Single Click
         Select Case lngR
         Case 10 To 17
            grdProfile.ForceCellEdit lngR, 1
         End Select
      End If
   End With

End Sub

Private Sub grdProfile_DblClick()

'''   grdProfile
'''   0  "Name"
'''   1  "Show As"
'''   2  "Groups"
'''   3  "Setting"
'''   4  "Company"
'''   5  "Job Title"
'''   6  "Address: Home"
'''   7  "Address: Business"
'''   8  "Address: Bill To"
'''   9  "Address: Ship To"
'''  10  "Phone: Home"
'''  11  "Phone: Business"
'''  12  "Phone: Mobile"
'''  13  "Phone: Other"
'''  14  "Phone: Fax"
'''  15  "E-Mail: Home"
'''  16  "E-Mail: Business"
'''  17  "E-Mail: Other"
'''  18  "Web Page"
'''  19  "Web Page"
'''  20  "ChangeLog"

  Dim lngR     As Long
  
   '// Edit with Double Click only (not Single Click)
   
   With grdProfile
      lngR = grdProfile.Row
      Select Case lngR
      Case 0, 2 To 9
         grdProfile.ForceCellEdit lngR, 1
      End Select
   End With

End Sub

Private Sub grdProfile_EditKeyPress(ByVal Col As Long, KeyAscii As Integer)

   '// Update Cell on Enter Key
   
   With grdProfile
      If KeyAscii = 13 Then
         KeyAscii = 0
         .UpdateCell
      End If
   End With

End Sub

Private Sub grdProfile_KeyPress(KeyAscii As Integer)
   
   '// Edit cell on Enter Key
   
   With grdProfile
      If Not .Row = 1 Then
         If Not .EditPending Then
            If KeyAscii = 13 Then
               KeyAscii = 0
               .ForceCellEdit .Row, 1
            End If
         End If
      End If
   End With

End Sub
