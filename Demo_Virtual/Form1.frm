VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demo_Virtual"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5940
      TabIndex        =   1
      Top             =   270
      Width           =   1155
   End
   Begin Project1.LynxGrid grdVirtual 
      Height          =   3780
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6668
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      FocusRectColor  =   9895934
      ThemeStyle      =   4
      ScrollBars      =   1
      Appearance      =   0
      ScrollBarStyle  =   1
      ShowRowNumbers  =   -1  'True
      ShowRowNumbersVary=   -1  'True
      HotHeaderTracking=   0   'False
   End
   Begin VB.Label lblInfo 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   5940
      TabIndex        =   3
      Top             =   690
      Width           =   1995
   End
   Begin VB.Label lblTarget 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5940
      TabIndex        =   2
      Top             =   1695
      Width           =   1995
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngRowsToAdd As Long

Private Sub cmdClear_Click()
   
   mlngRowsToAdd = Int((500 * Rnd) + 100)
   lblTarget.Caption = "Maxium Rows: " & CStr(mlngRowsToAdd)
   
   grdVirtual.Clear
   grdVirtual.Redraw = True '// this will trigger the RequestRowData event

End Sub

Private Sub Form_Load()
   
  Dim lngR As Long
  Dim lngI As Long
  
   Randomize
   mlngRowsToAdd = Int((500 * Rnd) + 100)
   lblTarget.Caption = "Maxium Rows: " & CStr(mlngRowsToAdd)
   
   With grdVirtual
      .AddColumn "Column 0"
      .AddColumn "Column 1"
      .ColForceFit
      .Redraw = True '// this will trigger the RequestRowData event
   End With
   
End Sub

Private Sub grdVirtual_RequestRowData(ByVal Row As Long)

  Dim lngR As Long

   With grdVirtual
      If .Rows < mlngRowsToAdd Then
         .Redraw = False
         lngR = .AddItem
         .CellText(lngR, 0) = "Row" & CStr(lngR) & "; Col0"
         .CellText(lngR, 1) = "Row" & CStr(lngR) & "; Col1"
         .Redraw = True
      End If

      '// display first row as number 1 instead of 0
      lblInfo.Caption = "Current Rows: " & CStr(.Rows) & vbNewLine & _
         "Top Visible Row: " & CStr(.RowFirstVisible + 1) & vbNewLine & _
         "Last Visible Row: " & CStr(.RowLastVisible + 1)
   End With

End Sub

Private Sub grdVirtual_Scroll()

   '// this is here to update info on Scroll events
   
   '// display first row as number 1 instead of 0
   With grdVirtual
      lblInfo.Caption = "Current Rows: " & CStr(.Rows) & vbNewLine & _
         "Top Visible Row: " & CStr(.RowFirstVisible + 1) & vbNewLine & _
         "Last Visible Row: " & CStr(.RowLastVisible + 1)
   End With

End Sub
