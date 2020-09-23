VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LynxGrid Tester © 2006 Richard Mewett"
   ClientHeight    =   8940
   ClientLeft      =   1215
   ClientTop       =   2445
   ClientWidth     =   12585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LynxGrid.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   12585
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picToolbar 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8220
      Left            =   7470
      ScaleHeight     =   8220
      ScaleWidth      =   5100
      TabIndex        =   6
      Top             =   660
      Width           =   5100
      Begin VB.Frame fraProperties 
         Caption         =   "Properties"
         Height          =   5025
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   5100
         Begin VB.ComboBox cboHiLightStyle 
            Height          =   315
            ItemData        =   "LynxGrid.frx":2382
            Left            =   1485
            List            =   "LynxGrid.frx":2384
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   4215
            Width           =   1395
         End
         Begin VB.CheckBox chkColumnDrag 
            Caption         =   "ColumnDrag"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   505
            Width           =   1425
         End
         Begin VB.CheckBox chkDisplayEllipsis 
            Caption         =   "DisplayEllipsis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1445
            Width           =   1425
         End
         Begin VB.CheckBox chkColumnResize 
            Caption         =   "ColumnResize"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   975
            Width           =   1425
         End
         Begin VB.CheckBox chkMultiSelect 
            Caption         =   "MultiSelect"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   2855
            Width           =   1425
         End
         Begin VB.CheckBox chkCheckBoxes 
            Caption         =   "CheckBoxes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   270
            Width           =   1425
         End
         Begin VB.CheckBox chkColumnSort 
            Caption         =   "ColumnSort"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1210
            Width           =   1425
         End
         Begin VB.CheckBox chkEditable 
            Caption         =   "Editable"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1680
            Width           =   1425
         End
         Begin VB.CheckBox chkScrollTrack 
            Caption         =   "ScrollTrack"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   3090
            Width           =   1425
         End
         Begin VB.CheckBox chkHotHeaderTracking 
            Caption         =   "AllowColumnHover"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   2620
            Width           =   1815
         End
         Begin VB.CheckBox chkGridLines 
            Caption         =   "GridLines"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   2150
            Width           =   1425
         End
         Begin VB.ComboBox cboFocusRectStyle 
            Height          =   315
            ItemData        =   "LynxGrid.frx":2386
            Left            =   1470
            List            =   "LynxGrid.frx":2388
            Style           =   2  'Dropdown List
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   3870
            Width           =   1005
         End
         Begin VB.CheckBox chkFullRowSelect 
            Caption         =   "FocusRowHighlight"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1915
            Width           =   1755
         End
         Begin VB.ComboBox cboFocusRectMode 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   3510
            Width           =   1005
         End
         Begin VB.ComboBox cboThemeColor 
            Height          =   315
            Left            =   2955
            Style           =   2  'Dropdown List
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1605
         End
         Begin VB.ComboBox cboThemeStyle 
            Height          =   315
            Left            =   2940
            Style           =   2  'Dropdown List
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2985
            Width           =   1605
         End
         Begin VB.CheckBox chkAlphaBlendSelection 
            Caption         =   "Alpha Blend Selection"
            Height          =   195
            Left            =   2940
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   4725
            Width           =   1935
         End
         Begin VB.CheckBox chkHideSelection 
            Caption         =   "FocusRectHide"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   2385
            Width           =   1425
         End
         Begin VB.CheckBox chkApplySelectionToImages 
            Caption         =   "Apply Selection To Images"
            Height          =   195
            Left            =   285
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   4710
            Width           =   2295
         End
         Begin VB.TextBox txtGotoRow 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   4335
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "300"
            Top             =   3510
            Width           =   450
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Goto Row"
            Height          =   285
            Left            =   3030
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   3510
            Width           =   1275
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Freeze At Col"
            Height          =   285
            Left            =   3030
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   3885
            Width           =   1275
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   4335
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   3870
            Width           =   450
         End
         Begin VB.CheckBox chkColumnSwap 
            Caption         =   "ColumnSwap"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   740
            Width           =   1425
         End
         Begin VB.Label lblColor 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BackColor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   2625
            TabIndex        =   85
            Top             =   225
            Width           =   735
         End
         Begin VB.Label lblViewColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   3405
            TabIndex        =   84
            Top             =   195
            Width           =   495
         End
         Begin VB.Label lblColor 
            AutoSize        =   -1  'True
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   4470
            TabIndex        =   83
            Top             =   810
            Width           =   195
         End
         Begin VB.Label lblColor 
            AutoSize        =   -1  'True
            Caption         =   "From"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   4455
            TabIndex        =   82
            Top             =   525
            Width           =   345
         End
         Begin VB.Label lblViewColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   3930
            TabIndex        =   81
            Top             =   780
            Width           =   495
         End
         Begin VB.Label lblViewColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   3930
            TabIndex        =   80
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "HighlightStyle"
            Height          =   195
            Index           =   1
            Left            =   435
            TabIndex        =   72
            Top             =   4245
            Width           =   975
         End
         Begin VB.Label lblColor 
            AutoSize        =   -1  'True
            Caption         =   "Hdr Txt"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   4455
            TabIndex        =   70
            Top             =   225
            Width           =   525
         End
         Begin VB.Label lblViewColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   3930
            TabIndex        =   69
            Top             =   195
            Width           =   495
         End
         Begin VB.Label lblViewColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   3405
            TabIndex        =   48
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label lblColor 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "GridColor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   2685
            TabIndex        =   47
            Top             =   2325
            Width           =   645
         End
         Begin VB.Label lblViewColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   3405
            TabIndex        =   46
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblColor 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BackColorSel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   2370
            TabIndex        =   45
            Top             =   1110
            Width           =   960
         End
         Begin VB.Label lblViewColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   3405
            TabIndex        =   44
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblColor 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BackColorBkg"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2310
            TabIndex        =   43
            Top             =   510
            Width           =   1020
         End
         Begin VB.Label lblColor 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BackColorEdit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   2325
            TabIndex        =   42
            Top             =   810
            Width           =   1005
         End
         Begin VB.Label lblViewColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3405
            TabIndex        =   41
            Top             =   780
            Width           =   495
         End
         Begin VB.Label lblViewColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   3405
            TabIndex        =   40
            Top             =   1380
            Width           =   495
         End
         Begin VB.Label lblColor 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ForeColorEdit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   2385
            TabIndex        =   39
            Top             =   1410
            Width           =   945
         End
         Begin VB.Label lblColor 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ForeColorSel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   2430
            TabIndex        =   38
            Top             =   1710
            Width           =   900
         End
         Begin VB.Label lblViewColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   3405
            TabIndex        =   37
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FocusRectStyle"
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   36
            Top             =   3930
            Width           =   1110
         End
         Begin VB.Label lblViewColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   3405
            TabIndex        =   35
            Top             =   1980
            Width           =   495
         End
         Begin VB.Label lblColor 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FocusRectColor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   2190
            TabIndex        =   34
            Top             =   2010
            Width           =   1140
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FocusRectMode"
            Height          =   195
            Left            =   270
            TabIndex        =   33
            Top             =   3570
            Width           =   1140
         End
         Begin VB.Label lblThemeColor 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Theme Color"
            Height          =   195
            Left            =   2055
            TabIndex        =   32
            Top             =   2670
            Width           =   900
         End
         Begin VB.Label lblThemeStyle 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Theme Style"
            Height          =   195
            Left            =   2010
            TabIndex        =   31
            Top             =   3030
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   0
         TabIndex        =   53
         Top             =   4905
         Width           =   3075
         Begin VB.CommandButton cmdRemoveItem 
            Caption         =   "Remove Row"
            Height          =   345
            Left            =   1695
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   630
            Width           =   1305
         End
         Begin VB.CommandButton cmdAddItems 
            Caption         =   "Add 1000 Rows"
            Height          =   345
            Left            =   1695
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   240
            Width           =   1305
         End
         Begin VB.CommandButton cmdSort 
            Caption         =   "Sort by Name"
            Height          =   345
            Left            =   1695
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   1020
            Width           =   1305
         End
         Begin VB.CommandButton cmdRangeFormat 
            Caption         =   "Range Format"
            Height          =   345
            Left            =   1695
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1410
            Width           =   1305
         End
         Begin VB.Label lblCellValue 
            BackColor       =   &H00004000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   600
            Left            =   75
            TabIndex        =   89
            Top             =   2130
            UseMnemonic     =   0   'False
            Width           =   2925
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Count:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   90
            TabIndex        =   67
            Top             =   195
            Width           =   570
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   1845
            Width           =   510
         End
         Begin VB.Label lblRowCol 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   705
            TabIndex        =   65
            Top             =   1800
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Row/Col:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   90
            TabIndex        =   64
            Top             =   1560
            Width           =   810
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Selected"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   135
            TabIndex        =   63
            Top             =   495
            Width           =   630
         End
         Begin VB.Label lblSelectedCount 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            TabIndex        =   62
            Top             =   450
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mouse"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1635
            TabIndex        =   61
            Top             =   1845
            Width           =   480
         End
         Begin VB.Label lblMouseRowCol 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2190
            TabIndex        =   60
            Top             =   1800
            Width           =   795
         End
         Begin VB.Label lblItemCount 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            DragMode        =   1  'Automatic
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            TabIndex        =   59
            Top             =   825
            Width           =   795
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rows"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   375
            TabIndex        =   58
            Top             =   900
            Width           =   405
         End
      End
      Begin VB.CommandButton cmdFormatting 
         Caption         =   "Formatting..."
         Height          =   345
         Left            =   3285
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   5355
         Width           =   1710
      End
      Begin VB.CommandButton cmdBinding 
         Caption         =   "Binding..."
         Height          =   345
         Left            =   3285
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   5775
         Width           =   1710
      End
      Begin VB.CommandButton cmdCellWrap 
         Caption         =   "Cell Wrap..."
         Height          =   345
         Left            =   3285
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   6195
         Width           =   1710
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add a Long string"
         Height          =   675
         Left            =   3285
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   6765
         Width           =   1695
      End
      Begin LynxGridTest.LynxGrid grdFText 
         Height          =   300
         Left            =   2715
         TabIndex        =   88
         Top             =   7770
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         FocusRectColor  =   9895934
         ShowColumnHeaders=   0   'False
         ScrollBars      =   3
         Appearance      =   0
         ColumnHeaderSmall=   0   'False
         TotalsLineShow  =   0   'False
         FocusRowHighlightKeepTextForecolor=   0   'False
         ShowRowNumbers  =   0   'False
         ShowRowNumbersVary=   0   'False
         HotHeaderTracking=   0   'False
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Simulate a formatted TextBox: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   480
         TabIndex        =   79
         Top             =   7830
         Width           =   2205
      End
      Begin VB.Label lblExamples 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Examples:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3120
         TabIndex        =   68
         Top             =   5070
         Width           =   870
      End
   End
   Begin LynxGridTest.LynxGrid LynxGrid1 
      Height          =   6930
      Left            =   15
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
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
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
      FocusRectColor  =   9895934
      GridLines       =   2
      ThemeStyle      =   7
      ColumnHeaderLines=   2
      Caption         =   "Employees"
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      ScrollBarStyle  =   1
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   -1  'True
      ShowRowNumbersVary=   0   'False
      AllowColumnResizing=   -1  'True
      AllowWordWrap   =   -1  'True
      ColumnDrag      =   -1  'True
      ColumnSort      =   -1  'True
      EditMove        =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6690
      Top             =   8265
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
            Picture         =   "LynxGrid.frx":238A
            Key             =   "MALE1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":2924
            Key             =   "MALE2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":2EBE
            Key             =   "MALE3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":3458
            Key             =   "FEMALE1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":39F2
            Key             =   "FEMALE2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":3F8C
            Key             =   "FEMALE3"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":4526
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":4690
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":47FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "LynxGrid.frx":4964
      Top             =   7665
      Width           =   7305
   End
   Begin VB.PictureBox picTopBar 
      Align           =   1  'Align Top
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   12585
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   12585
      Begin VB.CommandButton Command7 
         Caption         =   "Unselect Row"
         Height          =   510
         Left            =   3675
         TabIndex        =   90
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add Col"
         Height          =   510
         Left            =   11820
         TabIndex        =   86
         ToolTipText     =   "Add a column"
         Top             =   60
         Width           =   750
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "Restore Default Col Order"
         Height          =   510
         Left            =   10530
         TabIndex        =   75
         Top             =   60
         Width           =   1305
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Column Order"
         Height          =   510
         Left            =   9195
         TabIndex        =   74
         Top             =   60
         Width           =   1350
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Column Order"
         Height          =   510
         Left            =   7890
         TabIndex        =   73
         Top             =   60
         Width           =   1320
      End
      Begin VB.CommandButton cmdAutoAjustColumns 
         Caption         =   "Force Fit"
         Height          =   510
         Left            =   7080
         TabIndex        =   77
         Top             =   60
         Width           =   825
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export"
         Height          =   510
         Left            =   6105
         TabIndex        =   76
         Top             =   60
         Width           =   990
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Toggle Row No"
         Height          =   510
         Left            =   5175
         TabIndex        =   78
         Top             =   60
         Width           =   945
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Filter"
         Height          =   510
         Left            =   4455
         TabIndex        =   87
         Top             =   60
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grid Control in a Single File solution"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   4
         Top             =   330
         Width           =   3060
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner-drawn editable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   3
         Top             =   60
         Width           =   1860
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
         ForeColor       =   &H0000FFFF&
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

Private Sub cboFocusRectMode_Click()

   LynxGrid1.FocusRectMode = cboFocusRectMode.ListIndex

End Sub

Private Sub cboFocusRectStyle_Click()

   LynxGrid1.FocusRectStyle = cboFocusRectStyle.ListIndex

End Sub

Private Sub cboHiLightStyle_click()

   LynxGrid1.FocusRowHighlightStyle = cboHiLightStyle.ListIndex

End Sub

Private Sub cboThemeColor_Click()

   Call ChangeGridBackColor(vbWindowBackground)
   LynxGrid1.ThemeColor = cboThemeColor.ListIndex
   SetColors

End Sub

Private Sub cboThemeStyle_Click()

   LynxGrid1.ThemeStyle = cboThemeStyle.ListIndex

End Sub

Private Sub chkAlphaBlendSelection_Click()

   LynxGrid1.AlphaBlendSelection = chkAlphaBlendSelection.Value

End Sub

Private Sub chkApplySelectionToImages_Click()

   LynxGrid1.ApplySelectionToImages = chkApplySelectionToImages.Value

End Sub

Private Sub chkCheckBoxes_Click()

   LynxGrid1.RowCheckBoxes = chkCheckBoxes.Value

End Sub

Private Sub chkColumnDrag_Click()

   LynxGrid1.AllowColumnDrag = chkColumnDrag.Value

   If chkColumnDrag.Value Then
      chkColumnSwap.Value = vbUnchecked
   End If

End Sub

Private Sub chkColumnResize_Click()

   If chkColumnResize.Value Then
      LynxGrid1.AllowColumnResizing = True
   Else
      LynxGrid1.AllowColumnResizing = False
   End If

End Sub

Private Sub chkColumnSort_Click()

   LynxGrid1.AllowColumnSort = chkColumnSort.Value

End Sub

Private Sub chkColumnSwap_Click()

   LynxGrid1.AllowColumnSwap = chkColumnSwap.Value

   If chkColumnSwap.Value Then
      chkColumnDrag.Value = vbUnchecked
   End If

End Sub

Private Sub chkDisplayEllipsis_Click()

   LynxGrid1.DisplayEllipsis = chkDisplayEllipsis.Value

End Sub

Private Sub chkEditable_Click()

   LynxGrid1.AllowEdit = chkEditable.Value

End Sub

Private Sub chkFullRowSelect_Click()

   LynxGrid1.FocusRowHighlight = chkFullRowSelect.Value

End Sub

Private Sub chkGridLines_Click()

   LynxGrid1.GridLines = chkGridLines.Value

End Sub

Private Sub chkHideSelection_Click()

   LynxGrid1.FocusRectHide = CBool(chkHideSelection.Value)

End Sub

Private Sub chkHotHeaderTracking_Click()

   LynxGrid1.AllowColumnHover = chkHotHeaderTracking.Value

End Sub

Private Sub chkMultiSelect_Click()

   LynxGrid1.MultiSelect = chkMultiSelect.Value

End Sub

Private Sub chkScrollTrack_Click()

   LynxGrid1.ScrollTrack = chkScrollTrack.Value

End Sub

Private Sub cmdAddItems_Click()

   Call LoadDemoData(1000)

End Sub

Private Sub cmdBinding_Click()

   frmBoundControls.Show vbModeless, Me

End Sub

Private Sub cmdCellWrap_Click()

   frmWrap.Show vbModeless, Me

End Sub

Private Sub cmdExport_Click()

   LynxGrid1.ExportGrid LynxGrid1.Caption, , False
   LynxGrid1.SetFocus

End Sub

Private Sub cmdFormatting_Click()

   frmFormatting.Show vbModeless, Me

End Sub

Private Sub cmdLoad_Click()

   LynxGrid1.ColOrderLoad Me.Name
   LynxGrid1.SetFocus

End Sub

Private Sub cmdRangeFormat_Click()

   With LynxGrid1
      .Redraw = False
      .Col = 3
'      'Change Font-Style on Forename Column
'      .FormatCells , , 2, 2, lgCFFontBold, True
'
'      'Change Colours on Surname Column
'      .FormatCells , , 3, 3, lgCFForeColor, vbYellow
'      .FormatCells , , 3, 3, lgCFBackColor, vbBlue
'      .FormatCells , , 0, 0, lgCFBackColor, &HC5C5C5
'      'Change Font on Job Title Column
'      .FormatCells , , 4, 4, lgCFFontName, "Times"
'
'      '// Keep Row#5 the same
'      .RowBackColor(5) = &H95E0F1
'      .RowForeColor(5) = &H1F488A

      .Redraw = True
      .SetFocus
   End With

End Sub

Private Sub cmdRemoveItem_Click()

   With LynxGrid1

      If .Row >= 0 Then
         .RemoveItem .Row
      End If
      .SetFocus

   End With
   lblItemCount.Caption = LynxGrid1.Rows

End Sub

Private Sub cmdRestore_Click()

   LynxGrid1.ColOrderRestore Me.Name
   LynxGrid1.SetFocus

End Sub

Private Sub cmdSave_Click()

   LynxGrid1.ColOrderSave Me.Name
   LynxGrid1.SetFocus

End Sub

Private Sub cmdSort_Click()

   'Sort the Grid by Columns 2 (Forename) & 3 (Surname)
   'NOTE: No Sort Order is specified so clicking button again automatically
   'reverses Sort Order

   LynxGrid1.Sort 3, , 2
   LynxGrid1.SetFocus

End Sub

Private Sub Command1_Click()

  Dim lngRow As Long

   LynxGrid1.SetFocus
   lngRow = Val(txtGotoRow.Text) - 1

   '// Goto only
   LynxGrid1.RowColSet lngRow, 3

   '// Goto and edit
   'LynxGrid1.ForceCellEdit lngRow, 3


End Sub

Private Sub Command2_Click()
  
  Dim lngR As Long
   
   With LynxGrid1
      lngR = .AddItem("" & vbTab & "" & vbTab & vbTab & "" & vbTab & "asd asdasdasd asdasd asdasd asdasd asdasd asd asd asdas" & _
         " dasdasdasd asdasdasd assdasdasd asdasdasd ads")
      .RowColSet lngR
      .SetFocus
   End With
   lblItemCount.Caption = LynxGrid1.Rows

End Sub

Private Sub Command3_Click()

   LynxGrid1.FreezeAtCol = Val(Text1.Text)
   LynxGrid1.Refresh

End Sub

Private Sub CreateGrid()

   'Notes:
   'Columns have an InputFilter property. This is a string which defines which
   'characters are allowed via keyboard entry using the internal TextBox editor.
   ' < = lowercase
   ' > = UPPERCASE
   ' 1234567890 = Allow only numbers

   'Date and Numeric columns default an InputFilter if one is not specified. You
   'can use the EditKeyPress Event for further control.

   With LynxGrid1
      'Set ImageList to provide Item Images
      .ImageList = ImageList1

      'Create the Columns
      .AddColumn "Code", 1000, , , ">"
      .AddColumn "G", 250, lgAlignCenterCenter
      .AddColumn "First Name", 1500
      .AddColumn "Last Name", 1500, , , ">" '// Allow Only UPPERCASE
      .AddColumn "Job Title", 800, , , , , , True, , , True   '// This column is locked

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

Private Sub cmdAutoAjustColumns_Click()
   LynxGrid1.ColForceFit
   LynxGrid1.SetFocus
End Sub

Private Sub Command4_Click()
   LynxGrid1.ShowRowNumbers = Not LynxGrid1.ShowRowNumbers
   LynxGrid1.SetFocus
End Sub

Private Sub Command5_Click()

 Dim lngR As Long
 Dim lngC As Long

   Screen.MousePointer = vbHourglass
   DoEvents

   With LynxGrid1
      .Redraw = False

      lngC = .Cols

      .AddColumn "More Notes", 2000, lgAlignRightCenter, lgString, , , , , 5
      .CellText(0, lngC) = "New Column"

      .Redraw = True
   End With

   Screen.MousePointer = vbDefault
   
End Sub

Private Sub Command6_Click()

  Dim strTemp As String
  
   strTemp = Trim$(InputBox("Filter Last Name Begins With:" & vbNewLine & "(Empty value will restore normal view)"))
   
   If LenB(strTemp) Then
      LynxGrid1.FilterOn strTemp, 3, lgSMBeginsWith, False
   Else
      LynxGrid1.FilterOff
   End If
   
End Sub

Private Sub Command7_Click()
    LynxGrid1.RowUnSelect
End Sub

Private Sub Form_Load()

   With cboHiLightStyle
      .AddItem "Solid"
      .AddItem "Gradient_V"
      .AddItem "Gradient_H"
   End With

   With cboFocusRectMode
      .AddItem "None"
      .AddItem "Row"
      .AddItem "Col"
   End With

   With cboFocusRectStyle
      .AddItem "Light"
      .AddItem "Heavy"
      .AddItem "Medium"
   End With

   With cboThemeColor
      .AddItem "Blue"
      .AddItem "Silver"
      .AddItem "Olive"
      .AddItem "Visual2005"
      .AddItem "Norton2004"
      .AddItem "CustomTheme"
      .AddItem "Autodetect"
   End With

   With cboThemeStyle
      .AddItem "Windows3D"
      .AddItem "WindowsFlat"
      .AddItem "WindowsTheme"
      .AddItem "OfficeXP"
      .AddItem "WindowsXP"
      .AddItem "Custom"
      .AddItem "Custom3D"
      .AddItem "Vista"
   End With

   'Set the controls to demo Properties
   With LynxGrid1
      chkCheckBoxes.Value = Abs(.RowCheckBoxes)
      chkColumnDrag.Value = Abs(.AllowColumnDrag)
      chkColumnSwap.Value = Abs(.AllowColumnSwap)
      chkColumnResize.Value = Abs(.AllowColumnResizing)
      chkColumnSort.Value = Abs(.AllowColumnSort)
      chkDisplayEllipsis.Value = Abs(.DisplayEllipsis)
      chkEditable.Value = Abs(.AllowEdit)
      chkFullRowSelect.Value = Abs(.FocusRowHighlight)
      chkHideSelection.Value = Abs(.FocusRectHide)
      chkHotHeaderTracking.Value = Abs(.AllowColumnHover)
      chkMultiSelect.Value = Abs(.MultiSelect)
      chkScrollTrack.Value = Abs(.ScrollTrack)

      cboFocusRectMode.ListIndex = .FocusRectMode
      cboFocusRectStyle.ListIndex = .FocusRectStyle

      cboThemeColor.ListIndex = .ThemeColor
      cboThemeStyle.ListIndex = .ThemeStyle

      chkAlphaBlendSelection.Value = Abs(.AlphaBlendSelection)
      chkApplySelectionToImages.Value = Abs(.ApplySelectionToImages)

      cboHiLightStyle.ListIndex = .FocusRowHighlightStyle

   End With

   With grdFText
      .AddColumn "", .Width, lgAlignRightCenter, lgNumeric, "$ #,0.00"
      .AddItem 3467.98
      .AllowEdit = True
      .ScrollBars = Scroll_None
      .ColumnHeaders = False
      .FocusRectMode = lgCol
      .FocusRectStyle = lgFRLight
      .FocusRowHighlight = False
      .Redraw = True
   End With
   
   CreateGrid
   LoadDemoData

End Sub

Private Sub Form_Resize()

   If Not Me.WindowState = vbMinimized Then
      picToolbar.Left = Me.ScaleWidth - picToolbar.Width - 50
      
      '// Use Move so that redraw is only fired once.
      With LynxGrid1
         .Move .Left, _
               .Top, _
               Me.ScaleWidth - .Left - picToolbar.Width - 100, _
               Me.ScaleHeight - .Top - txtHelp.Height - 100
      End With
      
      txtHelp.Move LynxGrid1.Left, LynxGrid1.Top + LynxGrid1.Height + 50, LynxGrid1.Width
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frmMain = Nothing

End Sub

Private Sub lblViewColor_Click(Index As Integer)

   Select Case cboThemeColor.ListIndex
   Case 5, 6
   Case Else
      MsgBox "You must change the Theme Color mode to Custom before changing this color.", vbInformation
      Exit Sub
   End Select
   
   Select Case cboThemeStyle.ListIndex
   Case 0, 1, 2, 4, 7
      If Index = 8 Or Index = 9 Then
         MsgBox "The Theme style selected prevents these colors from being changed." & _
         "  Only Themes OfficeXP, Custom, and Custom3D can be changed", vbInformation
         Exit Sub
      End If
   End Select
   
   gclrBack = Val(lblViewColor(Index).BackColor)
   frmColorPicker.Init gclrBack, gclrBack, Me
   Unload frmColorPicker
   lblViewColor(Index).BackColor = gclrBack
   
   With LynxGrid1
      .Redraw = False
      
      Select Case Index
      Case 0: .BackColorBkg = lblViewColor(Index).BackColor
      Case 1: .BackColorEdit = lblViewColor(Index).BackColor
      Case 2: .BackColorSel = lblViewColor(Index).BackColor
      Case 3: .ForeColorEdit = lblViewColor(Index).BackColor
      Case 4: .ForeColorSel = lblViewColor(Index).BackColor
      Case 5: .FocusRectColor = lblViewColor(Index).BackColor
      Case 6: .GridColor = lblViewColor(Index).BackColor
      Case 7: .ForeColorHdr = lblViewColor(Index).BackColor
      Case 8: .ThemeCustomColorFrom = lblViewColor(Index).BackColor
      Case 9: .ThemeCustomColorTo = lblViewColor(Index).BackColor
      Case 10: Call ChangeGridBackColor(lblViewColor(Index).BackColor)
      End Select

      .Redraw = True
      
   End With
   
End Sub

Private Sub ChangeGridBackColor(ByVal vNewColor As Long)
      
  Dim lngR As Long
  Dim lngC As Long
  Dim lngB As Long
   
   With LynxGrid1
      .Redraw = False
      For lngR = 0 To .Rows - 1
         For lngC = 0 To .Cols - 1
            If .CellBackColor(lngR, lngC) = .BackColor Then
               .CellBackColor(lngR, lngC) = vNewColor
            End If
      Next lngC, lngR
      .BackColor = vNewColor
      .Redraw = True
   End With

End Sub

Private Sub LoadDemoData(Optional ByVal vNewRows As Long = 500)

  Dim lCount      As Long
  Dim lRow        As Long
  Dim sForename   As String
  Dim sGender     As String

   Screen.MousePointer = vbHourglass
   DoEvents
   
   With LynxGrid1
      .Redraw = False

      'Add some random data
      For lCount = 1 To vNewRows

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
            .RowImage(lRow) = "MALE2" ' & RandomInt(1, 3)
            .CellForeColor(lRow, 1) = vbBlue

         Else
            .RowImage(lRow) = "FEMALE2" 'RandomInt(3, 6)
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
      
      '// Show Hand on row #7, Col #8
      .CellText(7, 8) = "Hand Pointer Here" '// value change
      .CellHandPointer(7, 8) = True

      'The grid supports per cell formatting but provides Item
      'formatting options for simplicity when only per Row formatting
      'is required (Row formatting reformats all Cells in the Row).
      .RowBackColor(5) = &H95E0F1
      .RowForeColor(5) = &H1F488A

      'Tell the grid to Draw
      .Redraw = True
   End With
   
   lblItemCount.Caption = LynxGrid1.Rows
   Screen.MousePointer = vbDefault

End Sub

Private Sub LynxGrid1_AfterDelete()
   lblItemCount.Caption = LynxGrid1.Rows
End Sub

Private Sub LynxGrid1_AfterInsert(ByVal Row As Long)
   lblItemCount.Caption = LynxGrid1.Rows
End Sub

Private Sub LynxGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

   'Is the Edit allowed?
   Select Case Col
   Case 1 'Gender Column
      Cancel = True
   End Select

End Sub

Private Sub LynxGrid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

   MsgBox "clicked button on row#" & CStr(Row + 1) & ", col#" & CStr(Col)

End Sub

Private Sub LynxGrid1_CellHandClick(ByVal Row As Long, ByVal Col As Long, Shift As Integer)
   MsgBox "this cell has the hand pointer (" & CStr(Row) & ", " & CStr(Col) & ")"
End Sub

Private Sub LynxGrid1_Click()

   If LynxGrid1.RowLocked(LynxGrid1.Row) Then
      MsgBox "This row is locked"

   ElseIf LynxGrid1.ColLocked(LynxGrid1.Col) Then
      MsgBox "This column is locked"
   End If
   
End Sub

Private Sub LynxGrid1_ColumnOrderChanged(ByVal ToCol As Long, ByVal FromCol As Long)

   lblRowCol.Caption = LynxGrid1.Row & "," & LynxGrid1.Col
   lblCellValue.Caption = LynxGrid1.CellText

End Sub

Private Sub LynxGrid1_MouseLeave()

   lblMouseRowCol.Caption = "-"

End Sub

Private Sub LynxGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

   lblMouseRowCol.Caption = LynxGrid1.MouseRow & "," & LynxGrid1.MouseCol
   lblMouseRowCol.Refresh

End Sub

Private Sub LynxGrid1_RowColChanged()

   lblRowCol.Caption = LynxGrid1.Row & "," & LynxGrid1.Col
   lblCellValue.Caption = LynxGrid1.CellText

End Sub

Private Sub LynxGrid1_SelectionChanged()

   lblSelectedCount.Caption = LynxGrid1.SelectedCount()
   lblSelectedCount.Refresh

End Sub

Private Sub SetColors()

   With LynxGrid1
      lblViewColor(0).BackColor = .BackColorBkg
      lblViewColor(1).BackColor = .BackColorEdit
      lblViewColor(2).BackColor = .BackColorSel
      lblViewColor(3).BackColor = .ForeColorEdit
      lblViewColor(4).BackColor = .ForeColorSel
      lblViewColor(5).BackColor = .FocusRectColor
      lblViewColor(6).BackColor = .GridColor
      lblViewColor(7).BackColor = .ForeColorHdr
      lblViewColor(8).BackColor = .ThemeCustomColorFrom
      lblViewColor(9).BackColor = .ThemeCustomColorTo
      lblViewColor(10).BackColor = .BackColor
   End With
   
End Sub

