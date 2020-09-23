VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMultiPgPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//************************************/
'// Author: Morgan Haueisen
'//         morganh@hartcom.net
'// Copyright (c) 1998-2005
'//************************************/
Option Explicit

Private Declare Function StretchBlt Lib "gdi32" ( _
      ByVal hdc As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long, _
      ByVal hSrcDC As Long, _
      ByVal XSrc As Long, _
      ByVal YSrc As Long, _
      ByVal nSrcWidth As Long, _
      ByVal nSrcHeight As Long, _
      ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'// Flag indicating Printing or Previewing
Private mblnPrintFlag   As Boolean
Private mblnNoPrinters  As Boolean
'// Object used for Print Preview
Private mObjPrint       As Control
'// Storage for the Printer's orignal scale mode
Private intPntrSM       As Integer
'// Storage for the Object's orignal scale mode
Private intObjSM        As Integer
'// Default Scale Mode
Private mintScaleMode   As Integer
'// The actual printable area (something a little less then the paper size)
Private msngPgWidth     As Single
Private msngPgHeight    As Single
Private mintOrientation As Integer
'// Remember ColorMode
Private mbytColorMode   As Byte
'// Remember Header Information
Private mstrTitleMain   As String
Private mstrTitleSub    As String
Private mblnTitleItalic As Boolean
Private mintPageNumber  As Integer
Private mstrTempDir     As String
Private mDLeftMar       As Single

Public Enum PageOrientation
   PagePortrait = vbPRORPortrait
   PageLandscape = vbPRORLandscape
End Enum

Public Enum PrinterColorModeTypes
   cmMonochrome = vbPRCMMonochrome
   cmColor = vbPRCMColor
End Enum

'// -----rotate fonts
Private Const LF_FACESIZE = 32

Private Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName As String * LF_FACESIZE
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" _
      Alias "CreateFontIndirectA" ( _
      ByRef lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" _
      Alias "TextOutA" ( _
      ByVal hdc As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal lpString As String, _
      ByVal nCount As Long) As Long

Public Property Let BackColor(ByVal NewColor As Long)

  Dim lngR As Long

   If mblnPrintFlag Then
      If NewColor = -1 Then
         Call MakeTrans '// I am not sure why this is necessary
         lngR = SetBkColor(Printer.hdc, vbWhite)
         Printer.FontTransparent = True
      Else
         Printer.FontTransparent = False
         lngR = SetBkColor(Printer.hdc, GetSystemColor(NewColor))
      End If

   Else
      If NewColor = -1 Then
         mObjPrint.FontTransparent = True
         lngR = SetBkColor(mObjPrint.hdc, vbWhite)
      Else
         mObjPrint.FontTransparent = False
         lngR = SetBkColor(mObjPrint.hdc, GetSystemColor(NewColor))
      End If
   End If

End Property

Public Property Get BackColor() As Long

   If mblnPrintFlag Then
      BackColor = GetBkColor(Printer.hdc)
   Else
      BackColor = mObjPrint.BackColor
   End If

End Property

Private Sub Class_Initialize()

   On Error GoTo Err_Proc

   mblnNoPrinters = (Err.Number = 482)
   If mblnNoPrinters Then mblnPrintFlag = False
   
   ColorMode = cmColor
   Orientation = PagePortrait
   PrintCopies = 1
   SendToPrinter = True
   mblnTitleItalic = True

   '// Default Scale Mode
   '// vbInches or vbCentimeters
   
   '// mintScaleMode = vbCentimeters
   '// mDLeftMar = 0.41
   
   mintScaleMode = vbInches
   mDLeftMar = 0.17

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "Class_Initialize"
   Err.Clear
   Resume Exit_Here

End Sub

Public Property Get ColorMode() As PrinterColorModeTypes

   ColorMode = mbytColorMode

End Property

Public Property Let ColorMode(ByVal vNewValue As PrinterColorModeTypes)

   mbytColorMode = vNewValue
   If mblnPrintFlag Then Printer.ColorMode = vNewValue

End Property

Public Property Let CurrentX(ByVal NewXvalue As Single)

   If mblnPrintFlag Then
      Printer.CurrentX = NewXvalue
   Else
      mObjPrint.CurrentX = NewXvalue
   End If

End Property

Public Property Get CurrentX() As Single

   If mblnPrintFlag Then
      CurrentX = Printer.CurrentX
   Else
      CurrentX = mObjPrint.CurrentX
   End If

End Property

Public Property Get CurrentY() As Single

   If mblnPrintFlag Then
      CurrentY = Printer.CurrentY
   Else
      CurrentY = mObjPrint.CurrentY
   End If

End Property

Public Property Let CurrentY(ByVal NewYvalue As Single)

   If mblnPrintFlag Then
      Printer.CurrentY = NewYvalue
   Else
      mObjPrint.CurrentY = NewYvalue
   End If

End Property

Public Property Let DrawStyle(ByVal vNewValue As DrawStyleConstants)

   On Error Resume Next
   Printer.DrawStyle = vNewValue
   mObjPrint.DrawStyle = vNewValue

End Property

Public Property Get DrawStyle() As DrawStyleConstants

   DrawStyle = Printer.DrawStyle

End Property

Public Property Let DrawWidth(ByVal NewWidth As Integer)

   If NewWidth < 1 Then NewWidth = 1

   If mblnPrintFlag Then
      Printer.DrawWidth = NewWidth
   Else
      NewWidth = NewWidth / 2
      If NewWidth < 1 Then NewWidth = 1
      mObjPrint.DrawWidth = NewWidth
   End If

End Property

Public Property Get DrawWidth() As Integer

   If mblnPrintFlag Then
      DrawWidth = Printer.DrawWidth
   Else
      DrawWidth = mObjPrint.DrawWidth * 2
   End If

End Property

Private Sub Err_Handler(Optional ByVal vblnDisplayError As Boolean = True, _
                        Optional ByVal vstrErrNumber As String = vbNullString, _
                        Optional ByVal vstrErrDescription As String = vbNullString, _
                        Optional ByVal vstrModuleName As String = vbNullString, _
                        Optional ByVal vstrProcName As String = vbNullString)

  Dim strTemp As String
  Dim lngFN   As Long

   '// Purpose: Error handling
   On Error Resume Next

   '// Show Error Message
   If vblnDisplayError Then
      strTemp = "Error occured: "
      If LenB(vstrErrNumber) Then strTemp = strTemp & vstrErrNumber & vbNewLine Else strTemp = strTemp & vbNewLine
      If LenB(vstrErrDescription) Then strTemp = strTemp & "Description: " & vstrErrDescription & vbNewLine
      If LenB(vstrModuleName) Then strTemp = strTemp & "Module: " & vstrModuleName & vbNewLine
      If LenB(vstrProcName) Then strTemp = strTemp & "Function: " & vstrProcName
      MsgBox strTemp, vbCritical, App.Title & " - ERROR"
   End If

   '// Write error log
   lngFN = FreeFile
   Open App.Path & "\ErrorLog.txt" For Append As #lngFN
   Write #lngFN, Now, vstrErrNumber, vstrErrDescription, vstrModuleName, vstrProcName, _
         App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision, _
         Environ("username"), Environ("computername")
   Close #lngFN

End Sub

Public Property Let FontBold(YesNo As Boolean)

   If mblnPrintFlag Then
      Printer.FontBold = YesNo
   Else
      mObjPrint.FontBold = YesNo
   End If

End Property

Public Property Get FontBold() As Boolean

   If mblnPrintFlag Then
      FontBold = Printer.FontBold
   Else
      FontBold = mObjPrint.FontBold
   End If

End Property

Public Property Let FontItalic(YesNo As Boolean)

   If mblnPrintFlag Then
      Printer.FontItalic = YesNo
   Else
      mObjPrint.FontItalic = YesNo
   End If

End Property

Public Property Get FontItalic() As Boolean

   If mblnPrintFlag Then
      FontItalic = Printer.FontItalic
   Else
      FontItalic = mObjPrint.FontItalic
   End If

End Property

Public Property Get FontName() As String

   If mblnPrintFlag Then
      FontName = Printer.FontName
   Else
      FontName = mObjPrint.FontName
   End If

End Property

Public Property Let FontName(ByVal newFont As String)

   If mblnPrintFlag Then
      Printer.FontName = newFont
      Printer.Print vbNullString;
   Else
      mObjPrint.FontName = newFont
      mObjPrint.Print vbNullString;
   End If

End Property

Public Property Let FontSize(pSize As Integer)

   If mblnPrintFlag Then
      Printer.FontSize = pSize
   Else
      mObjPrint.FontSize = pSize
   End If

End Property

Public Property Get FontSize() As Integer

   If mblnPrintFlag Then
      FontSize = Printer.FontSize
   Else
      FontSize = mObjPrint.FontSize
   End If

End Property

Public Property Let FontStrikethru(YesNo As Boolean)

   If mblnPrintFlag Then
      Printer.FontStrikethru = YesNo
   Else
      mObjPrint.FontStrikethru = YesNo
   End If

End Property

Public Property Get FontStrikethru() As Boolean

   If mblnPrintFlag Then
      FontStrikethru = Printer.FontStrikethru
   Else
      FontStrikethru = mObjPrint.FontStrikethru
   End If

End Property

Public Property Let FontTransparent(ByVal vNewValue As Boolean)

   If mblnPrintFlag Then
      If vNewValue Then
         '// I am not sure why this is necessary but it doesn't work without it
         Call MakeTrans
         Me.BackColor = -1
      Else
         Printer.FontTransparent = False
      End If

   Else
      mObjPrint.FontTransparent = vNewValue
   End If

End Property

Public Property Get FontTransparent() As Boolean

   If mblnPrintFlag Then
      FontTransparent = Printer.FontTransparent
   Else
      FontTransparent = mObjPrint.FontTransparent
   End If

End Property

Public Property Get FontUnderline() As Boolean

   If mblnPrintFlag Then
      FontUnderline = Printer.FontUnderline
   Else
      FontUnderline = mObjPrint.FontUnderline
   End If

End Property

Public Property Let FontUnderline(YesNo As Boolean)

   If mblnPrintFlag Then
      Printer.FontUnderline = YesNo
   Else
      mObjPrint.FontUnderline = YesNo
   End If

End Property

Public Property Let ForeColor(NewColor As Long)

   If mblnPrintFlag Then
      Printer.ForeColor = GetSystemColor(NewColor)
   Else
      mObjPrint.ForeColor = GetSystemColor(NewColor)
   End If

End Property

Public Property Get ForeColor() As Long

   If mblnPrintFlag Then
      ForeColor = Printer.ForeColor
   Else
      ForeColor = mObjPrint.ForeColor
   End If

End Property

Public Function GetFormalCase(ByVal TextString As String) As String

  Dim lngI As Long

   '// Cap the first letter if each word
   On Local Error Resume Next

   TextString = UCase$(Left$(TextString, 1)) & LCase$(Mid$(TextString, 2))

   '// Look for space
   lngI = InStr(TextString, " ")

   If lngI Then

      Do
         Mid$(TextString, lngI + 1, 1) = UCase$(Mid$(TextString, lngI + 1, 1))
         lngI = lngI + 1
         lngI = InStr(lngI, TextString, " ")
         If lngI = 0 Or lngI + 1 > Len(TextString) Then Exit Do
      Loop

   End If

   '// Look for .
   lngI = InStr(TextString, ".")

   If lngI Then

      Do
         Mid$(TextString, lngI + 1, 1) = UCase$(Mid$(TextString, lngI + 1, 1))
         lngI = lngI + 1
         lngI = InStr(lngI, TextString, ".")
         If lngI = 0 Or lngI + 1 > Len(TextString) Then Exit Do
      Loop

   End If

   '// Look for Mc
   lngI = InStr(TextString, "Mc")

   If lngI Then

      Do
         Mid$(TextString, lngI + 2, 1) = UCase$(Mid$(TextString, lngI + 2, 1))
         lngI = lngI + 2
         lngI = InStr(lngI, TextString, "Mc")
         If lngI = 0 Or lngI + 2 > Len(TextString) Then Exit Do
      Loop

   End If

   '// Look for O'
   lngI = InStr(TextString, "O'")

   If lngI Then

      Do
         Mid$(TextString, lngI + 2, 1) = UCase$(Mid$(TextString, lngI + 2, 1))
         lngI = lngI + 2
         lngI = InStr(lngI, TextString, "O'")
         If lngI = 0 Or lngI + 2 > Len(TextString) Then Exit Do
      Loop

   End If

   '// Look for -
   lngI = InStr(TextString, "-")

   If lngI Then

      Do
         Mid$(TextString, lngI + 1, 1) = UCase$(Mid$(TextString, lngI + 1, 1))
         lngI = lngI + 1
         lngI = InStr(lngI, TextString, "-")
         If lngI = 0 Or lngI + 1 > Len(TextString) Then Exit Do
      Loop

   End If

   GetFormalCase = LTrim$(TextString)

End Function

Public Function GetPage() As Variant

   If mblnPrintFlag Then
      GetPage = Printer.Page
   Else
      GetPage = mintPageNumber + 1
   End If

End Function

Public Function GetPaperHeight() As Single

   GetPaperHeight = msngPgHeight

End Function

Public Function GetPaperWidth() As Single

   GetPaperWidth = msngPgWidth

End Function

Public Function GetRemoveCRLF(ByVal TextString As String) As String

  Dim lngI           As Long
  Dim blnFoundString As Boolean
  Dim blnFoundFirst  As Boolean

   On Error GoTo Err_Proc

   Do
      blnFoundString = False

      lngI = InStr(TextString, vbCr)

      If lngI Then
         Mid$(TextString, lngI, 1) = " "
         blnFoundString = True
         blnFoundFirst = True
      End If

      lngI = InStr(TextString, vbLf)

      If lngI = 1 Then
         TextString = Mid$(TextString, lngI + 1)
      ElseIf lngI > 1 Then

         If blnFoundFirst Then
            TextString = Mid$(TextString, 1, lngI - 1) & Mid$(TextString, lngI + 1)
         Else
            Mid$(TextString, lngI, 1) = " "
         End If

         blnFoundString = True
      End If

      blnFoundFirst = False

   Loop Until blnFoundString = False
   GetRemoveCRLF = TextString

Exit_Here:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "GetRemoveCRLF"
   Err.Clear
   Resume Exit_Here

End Function

Private Function GetSystemColor(ByVal vlngColor As Long) As String

   '// Is the color a VB color constant?
   If vlngColor < 0 Then
      '// Retrieves the current color of the specified display element
      vlngColor = GetSysColor(vlngColor And &HFF&)
   End If

   GetSystemColor = "&H" & Right$("000000" & Hex(vlngColor), 6)

End Function

Public Function GetStripQuotes(ByVal TextString As String) As String

   If Left$(TextString, 1) = Chr$(34) Then TextString = Mid$(TextString, 2)
   If Right$(TextString, 1) = Chr$(34) Then TextString = Left$(TextString, Len(TextString) - 1)
   GetStripQuotes = TextString

End Function

Public Function GetTextHeight(Optional TextString As String = "Sample Text") As Single

   If mblnPrintFlag Then
      GetTextHeight = Printer.TextHeight(TextString)
   Else
      GetTextHeight = mObjPrint.TextHeight(TextString)
   End If

End Function

Public Function GetTextWidth(TextString As Variant) As Single

   If mblnPrintFlag Then
      GetTextWidth = Printer.TextWidth(TextString)
   Else
      GetTextWidth = mObjPrint.TextWidth(TextString)
   End If

End Function

Private Sub MakeTrans()

  Dim sngX As Single
  Dim sngY As Single

   On Error GoTo Err_Proc

   sngX = Me.CurrentX
   sngY = Me.CurrentY
   '// I am not sure why this is necessary
   '// but it doesn't work without it.
   Sleep 1
   Me.BackColor = vbWhite
   Sleep 1
   pQuarterSpace
   Sleep 1
   Me.CurrentX = sngX
   Me.CurrentY = sngY
   Sleep 1

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "MakeTrans"
   Err.Clear
   Resume Exit_Here

End Sub

Public Property Get Orientation() As PageOrientation

   Orientation = mintOrientation

End Property

Public Property Let Orientation(ByVal vNewValue As PageOrientation)

   On Error Resume Next
   Printer.Orientation = vNewValue
   mintOrientation = vNewValue

End Property

Public Sub p15Space()

  Dim eFont As Integer
  Dim hFont As Integer

   On Error GoTo Err_Proc
   pPrint
   eFont = Me.FontSize
   hFont = eFont \ 2
   Me.FontSize = hFont
   pPrint
   Me.FontSize = eFont

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "p15Space"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pBox(Optional ByVal bLeft As Single = -1, _
                Optional ByVal bTop As Single = -1, _
                Optional ByVal bWidth As Single = -1, _
                Optional ByVal bHeight As Single = -1, _
                Optional ByVal ColorLine As Long = -1, _
                Optional ByVal ColorFill As Long = -1, _
                Optional FilledBox As FillStyleConstants = vbFSTransparent)

  Dim sngX As Single
  Dim sngY As Single

   On Error GoTo Err_Proc

   sngY = Me.CurrentY
   sngX = Me.CurrentX

   If ColorLine = -1 Then ColorLine = Me.ForeColor
   If ColorFill = -1 Then ColorFill = ColorLine
   If bLeft = -1 Then
      bLeft = sngX
      If bLeft < mDLeftMar Then bLeft = mDLeftMar
   End If
   If bTop = -1 Then bTop = sngY
   If bWidth = -1 Then bWidth = msngPgWidth - bLeft
   If bHeight = -1 Then bHeight = GetTextHeight

   If FilledBox <> vbFSTransparent Then
      If mblnPrintFlag Then
         Printer.FillColor = ColorFill
         Printer.FillStyle = FilledBox
         Printer.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), ColorLine, B
         Printer.FillStyle = vbFSTransparent
      Else
         mObjPrint.FillColor = ColorFill
         mObjPrint.FillStyle = FilledBox
         mObjPrint.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), ColorLine, B
         mObjPrint.FillStyle = vbFSTransparent
      End If

   Else

      If mblnPrintFlag Then
         Printer.FillStyle = vbFSTransparent
         Printer.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), ColorLine, B
      Else
         mObjPrint.FillStyle = vbFSTransparent
         mObjPrint.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), ColorLine, B
      End If

   End If

   Me.CurrentX = sngX
   Me.CurrentY = sngY

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pBox"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pBullet(Optional ByVal LeftMargin As Single = -1)

  Dim eFontN As String
  Dim eFontB As Boolean
  Dim eFontI As Boolean
  Dim eFontU As Boolean
  Dim eFontK As Boolean

   On Error GoTo Err_Proc

   If LeftMargin = -1 Then
      LeftMargin = Me.CurrentX
      If LeftMargin < mDLeftMar Then LeftMargin = mDLeftMar
   End If

   '// Save current setting
   eFontN = Me.FontName
   eFontB = Me.FontBold
   eFontI = Me.FontItalic
   eFontU = Me.FontUnderline
   eFontK = Me.FontStrikethru

   '// Change settings
   '// FontName = "Wingdings"
   Me.FontName = "Symbol"
   Me.FontBold = False
   Me.FontItalic = False
   Me.FontUnderline = False
   Me.FontStrikethru = False

   'pPrint Chr(164) & "  ", LeftMargin, True
   pPrint Chr(183) & "  ", LeftMargin, True

   '// Restore setting
   Me.FontName = eFontN
   Me.FontBold = eFontB
   Me.FontItalic = eFontI
   Me.FontUnderline = eFontU
   Me.FontStrikethru = eFontK

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pBullet"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pCancled()

   Me.FontSize = 12
   Me.FontBold = True
   Me.ForeColor = vbRed
   pPrint
   pPrint "**** PRINTING CANCLED ****", 0.5

End Sub

Public Sub pCenter(ByVal PrintVar As String, _
                   Optional ByVal SameLine As Boolean = False, _
                   Optional COLWIDTH As Single = -1, _
                   Optional LeftMargin As Single = 0)

   On Error GoTo Err_Proc

  Dim lngI    As Long
  Dim strTemp As String
  
   On Error GoTo Err_Proc

   If COLWIDTH = -1 Then COLWIDTH = msngPgWidth - LeftMargin
   
   Do
      lngI = InStr(PrintVar, vbNewLine)
      If lngI Then
         strTemp = Mid$(PrintVar, 1, lngI - 1)
         PrintVar = Mid$(PrintVar, lngI + 2)
      Else
         strTemp = PrintVar
      End If
   
      If GetTextWidth(strTemp) > msngPgWidth Then
         Me.pCenterMultiline strTemp, LeftMargin, LeftMargin + COLWIDTH, , (SameLine And lngI = 0)
      Else
         Me.CurrentX = LeftMargin + ((COLWIDTH - GetTextWidth(strTemp)) / 2)
         Me.pPrint strTemp, , (SameLine And lngI = 0)
      End If
   
   Loop Until lngI = 0

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pCenter"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pCenterMultiline(ByVal PrintVar As String, _
                            Optional ByVal LeftMargin As Single = 0, _
                            Optional ByVal RightMargin As Single = -1, _
                            Optional ByVal RemoveCrLf As Boolean = True, _
                            Optional ByVal SameLine As Boolean = False)

  Dim intStartChar  As Integer
  Dim sngCharLength As Single
  Dim sngCurrentPos As Single
  Dim sngY          As Single
  Dim sngTxtLen     As Single
  Dim strTemp       As String
  Dim blnStrip      As Boolean
  Dim sngColW       As Single
  Dim lngX          As Long
  Dim lngY          As Long

   On Error GoTo Err_Proc

   If RightMargin = -1 Then RightMargin = msngPgWidth
   sngColW = RightMargin - LeftMargin
   sngY = Me.CurrentY

   sngTxtLen = Len(PrintVar)
   intStartChar = 1
   sngCurrentPos = 0
   sngCharLength = sngTxtLen

   If LenB(PrintVar) = 0 Then
      pPrint
      Exit Sub
   End If

   blnStrip = (InStrB(PrintVar, vbCr) > 0 Or InStrB(PrintVar, vbLf) > 0)

   For lngX = 1 To sngTxtLen
      lngY = lngX - sngCurrentPos

      '// Mark space between words
      If Mid$(PrintVar, lngX, 1) < Chr(33) Then sngCharLength = lngY

      If (GetTextWidth(Mid$(PrintVar, intStartChar, lngY)) >= sngColW) _
         Or (Not RemoveCrLf And Mid$(PrintVar, lngX, 1) = vbCr) Then

         '// If there are no spaces then break line here */
         If sngCharLength > lngY Then sngCharLength = lngY - 1

         strTemp = Trim$(Mid$(PrintVar, intStartChar, sngCharLength))
         If blnStrip Then strTemp = GetRemoveCRLF(strTemp)
         Me.CurrentX = LeftMargin + ((sngColW - GetTextWidth(strTemp)) / 2)

         If mblnPrintFlag Then
            Printer.Print strTemp
         Else
            mObjPrint.Print strTemp
         End If

         sngCurrentPos = sngCharLength + sngCurrentPos
         intStartChar = sngCurrentPos + 1
         sngCharLength = sngTxtLen

      End If
   Next lngX

   strTemp = Trim$(Mid$(PrintVar, intStartChar))
   If blnStrip Then strTemp = GetRemoveCRLF(strTemp)
   Me.CurrentX = LeftMargin + ((sngColW - GetTextWidth(strTemp)) / 2)

   If mblnPrintFlag Then
      Printer.Print strTemp
   Else
      mObjPrint.Print strTemp
   End If

   If SameLine Then Me.CurrentY = sngY

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pCenterMultiline"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pCharacter(ByVal vstrChar As String, _
                      Optional ByVal RightMargin As Single = -1, _
                      Optional ByVal LeftMargin As Single = -1, _
                      Optional ByVal SameLine As Boolean = False)

  Dim strFontName    As String
  Dim lngForeColor   As Long
  Dim sngCharWidth   As Single

   On Error GoTo Err_Proc

   strFontName = Me.FontName
   lngForeColor = Me.ForeColor

   If LeftMargin <> -1 Then
      Me.CurrentX = LeftMargin
      If LeftMargin < mDLeftMar Then LeftMargin = mDLeftMar
   End If
   If RightMargin = -1 Then RightMargin = msngPgWidth

   If Me.CurrentX >= RightMargin Then GoTo ExitSpaceSub
   sngCharWidth = GetTextWidth(vstrChar)

   Do
      If mblnPrintFlag Then
         Printer.Print vstrChar;
      Else
         mObjPrint.Print vstrChar;
      End If

   Loop Until Me.CurrentX + sngCharWidth >= RightMargin
   
   If Not SameLine Then
      If mblnPrintFlag Then
         Printer.Print vbNullString
      Else
         mObjPrint.Print vbNullString
      End If
   End If

ExitSpaceSub:
   Me.FontName = strFontName
   Me.ForeColor = lngForeColor

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pCharacter"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pCircle(ByVal bLeft As Single, _
                   ByVal bTop As Single, _
                   ByVal bRadius As Single, _
                   Optional ByVal ColorLine As Long = -1, _
                   Optional ByVal ColorFill As Long = -1, _
                   Optional FilledCircle As FillStyleConstants = vbFSTransparent, _
                   Optional AspectRatio As Single = 1)

   On Error GoTo Err_Proc

   If ColorLine = -1 Then ColorLine = Me.ForeColor
   If ColorFill = -1 Then ColorFill = ColorLine

   If mblnPrintFlag Then
      If FilledCircle <> vbFSTransparent Then
         Printer.FillStyle = FilledCircle
         Printer.FillColor = ColorFill
      End If

      Printer.Circle (bLeft, bTop), bRadius, ColorLine, , , AspectRatio

      Printer.FillStyle = vbFSTransparent
   Else

      If FilledCircle <> vbFSTransparent Then
         mObjPrint.FillStyle = FilledCircle
         mObjPrint.FillColor = ColorFill
      End If

      mObjPrint.Circle (bLeft, bTop), bRadius, ColorLine, , , AspectRatio

      mObjPrint.FillStyle = vbFSTransparent
   End If

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pCircle"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pDots(ByVal RightMargin As Single, Optional ByVal LeftMargin As Single = -1)

   Call pCharacter(".", RightMargin, LeftMargin, True)
   
End Sub

Public Sub pDoubleLine(Optional ByVal LeftPos As Single = 0, _
                       Optional ByVal RightPos As Single = -1, _
                       Optional ByVal LineWidth As Integer = 1, _
                       Optional IncludeSpace As Boolean = True)

  Dim intDrawWidth As Integer
  Dim sngY         As Single
  Dim sngX         As Single

   On Error GoTo Err_Proc

   sngX = Me.CurrentX

   Select Case mintScaleMode
   Case vbCentimeters
      sngY = 0.07

   Case Else '// vbinches
      sngY = 0.03
   End Select

   intDrawWidth = Me.DrawWidth
   If LineWidth > 0 Then Me.DrawWidth = LineWidth
   'If IncludeSpace Then me.CurrentY = me.CurrentY + sngY

   If mblnPrintFlag Then
      sngX = Printer.CurrentX
      If RightPos <= LeftPos Then RightPos = msngPgWidth
      Printer.CurrentY = Printer.CurrentY + sngY
      Printer.Line (LeftPos, Printer.CurrentY)-(RightPos, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + sngY
      Printer.Line (LeftPos, Printer.CurrentY)-(RightPos, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + sngY
      Printer.CurrentX = sngX
   Else
      sngX = mObjPrint.CurrentX
      If RightPos <= LeftPos Then RightPos = msngPgWidth
      mObjPrint.CurrentY = mObjPrint.CurrentY + sngY
      mObjPrint.Line (LeftPos, mObjPrint.CurrentY)-(RightPos, mObjPrint.CurrentY)
      mObjPrint.CurrentY = mObjPrint.CurrentY + sngY
      mObjPrint.Line (LeftPos, mObjPrint.CurrentY)-(RightPos, mObjPrint.CurrentY)
      mObjPrint.CurrentY = mObjPrint.CurrentY + sngY
      mObjPrint.CurrentX = sngX
   End If

   If LineWidth > 0 Then Me.DrawWidth = intDrawWidth
   If IncludeSpace Then Me.CurrentY = Me.CurrentY + sngY

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pDoubleLine"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pEndDoc(Optional ByVal oModal As Byte = 1, _
                   Optional ByRef OwnerForm As Form, _
                   Optional ByRef FormToHide As Form)

   If mblnPrintFlag Then
      Printer.EndDoc
      Printer.ScaleMode = intPntrSM
      Me.SendToPrinter = False
   Else
      On Local Error Resume Next
      mObjPrint.ScaleMode = intObjSM

      SavePicture mObjPrint.Image, mstrTempDir & "PPview" & CStr(mintPageNumber) & ".bmp"

      frmMultiPgPreview.PageNumber = mintPageNumber
      mObjPrint.Picture = LoadPicture(mstrTempDir & "PPview0.bmp")
      FormToHide.Hide
      frmMultiPgPreview.Show oModal, OwnerForm
   End If
   
   Screen.MousePointer = vbDefault

End Sub

Public Function pEndOfPage(Optional ByVal Less As Single = 0, _
                           Optional ByVal SaveRoomForFooter As Boolean = True) As Boolean

  Dim sngN           As Single
  Dim sngTextH       As Single
  Dim intFontSize    As Integer
  Const C_TextString As String = "CgGyJj"
  
   On Error GoTo Err_Proc

   If mblnPrintFlag Then
      '// Make sure there is room for the footer
      sngTextH = Printer.TextHeight(C_TextString)

      If SaveRoomForFooter Then
         intFontSize = Printer.FontSize
         Printer.FontSize = 10
         sngTextH = Printer.TextHeight(C_TextString) * 2.5
         Printer.FontSize = intFontSize
      End If

      sngN = Printer.CurrentY + Printer.TextHeight(C_TextString) + sngTextH + Less
   Else
      '// Make sure there is room for the footer
      sngTextH = mObjPrint.TextHeight(C_TextString)

      If SaveRoomForFooter Then
         intFontSize = mObjPrint.FontSize
         mObjPrint.FontSize = 10
         sngTextH = mObjPrint.TextHeight(C_TextString) * 2
         mObjPrint.FontSize = intFontSize
      End If

      sngN = mObjPrint.CurrentY + mObjPrint.TextHeight(C_TextString) + sngTextH + Less
   End If

   If sngN >= msngPgHeight Then
      pEndOfPage = True
   Else
      pEndOfPage = False
   End If

Exit_Here:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pEndOfPage"
   Err.Clear
   Resume Exit_Here

End Function

Public Sub pFontName(Optional ByVal pFontName As String = "Times New Roman")

   On Error GoTo Err_Proc

   If mblnPrintFlag Then
      Printer.FontName = pFontName
      Printer.Print vbNullString;
   Else
      mObjPrint.FontName = pFontName
      mObjPrint.Print vbNullString;
   End If

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pFontName"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pFooter(Optional FooterText As String = vbNullString)

  Dim eFontS  As Integer
  Dim eFontN  As String
  Dim eFontB  As Boolean
  Dim eFontI  As Boolean
  Dim eFontU  As Boolean
  Dim eFontK  As Boolean
  Dim tMargin As Single

   On Error GoTo Err_Proc

   Select Case mintScaleMode
   Case vbCentimeters
      tMargin = 1.25

   Case Else '// vbinches
      tMargin = 0.5
   End Select

   '// Save current setting
   eFontN = Me.FontName
   eFontS = Me.FontSize
   eFontB = Me.FontBold
   eFontI = Me.FontItalic
   eFontU = Me.FontUnderline
   eFontK = Me.FontStrikethru

   '// Change settings
   pFontName
   Me.FontSize = 10
   Me.FontBold = False
   Me.FontItalic = False
   Me.FontUnderline = False
   Me.FontStrikethru = False
   Me.CurrentY = msngPgHeight - (GetTextHeight * 2)
   pLine 0, , 6

   If LenB(FooterText) = 0 Then
      FooterText = "Printed: " & Format$(Now, "ddd. mmmm d, yyyy \@ h:mm ampm")
   End If

   pPrint FooterText, tMargin, True
   pRightJust "Pg. " & GetPage, msngPgWidth - tMargin

   '// Restore setting
   Me.FontName = eFontN
   Me.FontSize = eFontS
   Me.FontBold = eFontB
   Me.FontItalic = eFontI
   Me.FontUnderline = eFontU
   Me.FontStrikethru = eFontK

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pFooter"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pHalfSpace()

  Dim eFont As Integer
  Dim hFont As Integer

   On Error GoTo Err_Proc

   eFont = Me.FontSize
   hFont = eFont \ 2
   If hFont < 1 Then hFont = 1
   Me.FontSize = hFont
   pPrint
   Me.FontSize = eFont

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pHalfSpace"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pHeader(Optional ByVal MainTitle As String = vbNullString, _
                   Optional SubTitle As Variant, _
                   Optional ItalicMain As Variant)

  Dim eFontS As Integer
  Dim eFontN As String
  Dim eFontB As Boolean
  Dim eFontI As Boolean
  Dim eFontU As Boolean
  Dim eFontK As Boolean

   On Error GoTo Err_Proc

   If LenB(MainTitle) = 0 And LenB(mstrTitleMain) = 0 Then Exit Sub

   '// Save current setting
   eFontN = Me.FontName
   eFontS = Me.FontSize
   eFontB = Me.FontBold
   eFontI = Me.FontItalic
   eFontU = Me.FontUnderline
   eFontK = Me.FontStrikethru
   Me.CurrentY = 0
   Me.CurrentX = 0

   If MainTitle > vbNullString Then mstrTitleMain = MainTitle
   If Not IsMissing(SubTitle) Then mstrTitleSub = SubTitle
   If Not IsMissing(ItalicMain) Then mblnTitleItalic = CBool(ItalicMain)

   '// Change settings
   pFontName
   Me.FontSize = 16
   Me.FontBold = True
   Me.FontItalic = mblnTitleItalic
   pCenter mstrTitleMain

   Me.FontSize = 11
   Me.FontBold = False
   Me.FontItalic = False
   If mstrTitleSub > vbNullString Then pCenter mstrTitleSub
   pDoubleLine
   pHalfSpace

   '// Restore setting
   Me.FontName = eFontN
   Me.FontSize = eFontS
   Me.FontBold = eFontB
   Me.FontItalic = eFontI
   Me.FontUnderline = eFontU
   Me.FontStrikethru = eFontK

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pHeader"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pLine(Optional ByVal LeftMargin As Single = -1, _
                 Optional ByVal RightMargin As Single = -1, _
                 Optional ByVal LineWidth As Integer = 1, _
                 Optional ByVal IncludeSpace As Boolean = True)

  Dim intDrawWidth As Integer
  Dim sngY         As Single
  Dim sngX         As Single

   On Error GoTo Err_Proc

   Select Case mintScaleMode
   Case vbCentimeters
      sngY = 0.07

   Case Else '// vbinches
      sngY = 0.03
   End Select

   If LeftMargin = -1 Then LeftMargin = Me.CurrentX
   If IncludeSpace Then Me.CurrentY = Me.CurrentY + sngY

   If LineWidth > 0 Then
      intDrawWidth = Me.DrawWidth
      Me.DrawWidth = LineWidth
   End If

   If mblnPrintFlag Then
      sngX = Printer.CurrentX
      If RightMargin <= LeftMargin Then RightMargin = msngPgWidth
      Printer.Line (LeftMargin, Printer.CurrentY)-(RightMargin, Printer.CurrentY)
      Printer.CurrentX = sngX
   Else
      sngX = mObjPrint.CurrentX
      If RightMargin <= LeftMargin Then RightMargin = msngPgWidth
      mObjPrint.Line (LeftMargin, mObjPrint.CurrentY)-(RightMargin, mObjPrint.CurrentY)
      mObjPrint.CurrentX = sngX
   End If

   If LineWidth > 0 Then Me.DrawWidth = intDrawWidth
   If IncludeSpace Then Me.CurrentY = Me.CurrentY + sngY

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pLine"
   Err.Clear
   Resume Exit_Here

End Sub

Public Function pMultiline(ByVal PrintVar As Variant, _
                           Optional ByVal LeftMargin As Single = -1, _
                           Optional ByVal RightMargin As Single = -1, _
                           Optional ByVal IndentChar As String = vbNullString, _
                           Optional ByVal SameLine As Boolean = True, _
                           Optional ByVal UsePageBreaks As Boolean = False) As Single

  Dim StartChar   As Integer
  Dim SecondLine  As Boolean
  Dim CharLength  As Single
  Dim CurrentPos  As Single
  Dim sngY        As Single
  Dim TxtLen      As Single
  Dim IndentText  As String
  Dim tString     As String
  Dim NeedsStrip  As Boolean
  Dim X           As Long
  Dim Y           As Long

   On Error GoTo Err_Proc

   If LeftMargin = -1 Then LeftMargin = Me.CurrentX
   'If LeftMargin > msngPgWidth - 0.1 Then LeftMargin = msngPgWidth - 0.5
   If LeftMargin >= msngPgWidth Then GoTo Exit_Here
   
   If RightMargin < LeftMargin Or RightMargin > msngPgWidth Then RightMargin = msngPgWidth - 0.1
   RightMargin = RightMargin - LeftMargin
   
   sngY = Me.CurrentY

   If PrintVar = vbNullString Then
      If SameLine Then
         pPrint vbNullString, , True
      Else
         pPrint
      End If
      GoTo Exit_Here
   End If

   TxtLen = Len(PrintVar)
   StartChar = 1
   CurrentPos = 0
   CharLength = TxtLen
   IndentText = vbNullString

   If InStr(PrintVar, vbCr) Or InStr(PrintVar, vbLf) Then NeedsStrip = True

   For X = 1 To TxtLen
      Y = X - CurrentPos
      If Mid$(PrintVar, X, 1) < Chr(33) Then CharLength = Y

      If (GetTextWidth(IndentText) + GetTextWidth(Mid$(PrintVar, StartChar, Y)) >= RightMargin) Or (Mid$(PrintVar, X, 1) = vbCr) Then
            
         '// If there are no spaces then break line here */
         If CharLength > Y Then CharLength = Y - Len(IndentText) - 1

         If NeedsStrip Then
            tString = IndentText & RTrim$(GetRemoveCRLF(Mid$(PrintVar, StartChar, CharLength)))
         Else
            tString = IndentText & Mid$(PrintVar, StartChar, CharLength)
         End If

         Me.CurrentX = LeftMargin

         If mblnPrintFlag Then
            Printer.Print tString
         Else
            mObjPrint.Print tString
         End If

         CurrentPos = CharLength + CurrentPos
         StartChar = CurrentPos + 1
         CharLength = TxtLen

         If Not SecondLine Then
            SecondLine = True
            IndentText = IndentChar
         End If

         If UsePageBreaks Then
            If pEndOfPage Then
               pFooter
               pNewPage
               pHeader
               SameLine = False
            End If
         End If
      
      End If
   Next X

   If NeedsStrip Then
      tString = IndentText & Trim$(GetRemoveCRLF(Mid$(PrintVar, StartChar)))
   Else
      tString = IndentText & Mid$(PrintVar, StartChar)
   End If

   Me.CurrentX = LeftMargin

   If mblnPrintFlag Then
      Printer.Print tString
   Else
      mObjPrint.Print tString
   End If

Exit_Here:
   pMultiline = Me.CurrentY

   If SameLine Then
      Me.CurrentY = sngY
   End If

   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "pMultiline"
   Err.Clear
   Resume Exit_Here

End Function

Public Sub pNewPage()

   On Local Error Resume Next

   If mblnPrintFlag Then
      Printer.NewPage
   Else
      SavePicture mObjPrint.Image, mstrTempDir & "PPview" & CStr(mintPageNumber) & ".bmp"
      mObjPrint.Cls
      mintPageNumber = mintPageNumber + 1
   End If

End Sub

Public Sub pPrint(Optional ByVal PrintVar As String = vbNullString, _
                  Optional ByVal LeftMargin As Single = -1, _
                  Optional ByVal SameLine As Boolean = False)

   On Error GoTo Err_Proc

   If LenB(PrintVar) = 0 Then
      '// Empty String */
   Else
      If LeftMargin = -1 Then
         LeftMargin = Me.CurrentX
         If LeftMargin < mDLeftMar Then LeftMargin = mDLeftMar
      End If

      If (GetTextWidth(PrintVar) + LeftMargin > msngPgWidth) Or InStrB(PrintVar, vbCr) > 0 Or InStrB(PrintVar, vbLf) > 0 Then
         Call pMultiline(PrintVar, LeftMargin, , , SameLine)
         Exit Sub
      End If

   End If

   If LeftMargin >= 0 Then Me.CurrentX = LeftMargin

   If SameLine Then
      If mblnPrintFlag Then
         Printer.Print PrintVar;
      Else
         mObjPrint.Print PrintVar;
      End If

   Else

      If mblnPrintFlag Then
         Printer.Print PrintVar
      Else
         mObjPrint.Print PrintVar
      End If

   End If

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pPrint"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pPrintedDate(Optional PrintCentered As Boolean = False, _
                        Optional ByVal LeftMargin As Single = -1, _
                        Optional ByVal SameLine As Boolean = False)

  Dim PrintVar    As String
  Dim FSize       As Integer
  Dim FBold       As Boolean
  Dim FItalic     As Boolean
  Dim FUnderline  As Boolean
  Dim FStrikethru As Boolean

   On Error GoTo Err_Proc

   FSize = Me.FontSize
   FBold = Me.FontBold
   FItalic = Me.FontItalic
   FUnderline = Me.FontUnderline
   FStrikethru = Me.FontStrikethru

   Me.FontSize = 9
   Me.FontBold = False
   Me.FontUnderline = False
   Me.FontItalic = False
   Me.FontStrikethru = False
   PrintVar = "Printed: " & Format$(Now, "ddd. mmm. d, yyyy \@ h:mm ampm")

   If PrintCentered Then
      Call pCenter(PrintVar, SameLine)
   Else
      Call pPrint(PrintVar, LeftMargin, SameLine)
   End If

   Me.FontSize = FSize
   Me.FontBold = FBold
   Me.FontItalic = FItalic
   Me.FontUnderline = FUnderline
   Me.FontStrikethru = FStrikethru

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pPrintedDate"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pPrintPicture(ByRef newPic As StdPicture, _
                         Optional LeftMargin As Single = -1, _
                         Optional TopMargin As Single = -1, _
                         Optional pWidth As Single = 0, _
                         Optional pHeight As Single = 0, _
                         Optional ByVal ScaleToFit As Boolean = False, _
                         Optional ByVal MaintainRatio As Boolean = True)

  Dim sngX      As Single
  Dim sngY      As Single
  Dim sngWidth  As Single
  Dim sngHeight As Single
  Dim sngAspect As Single
  Dim picBox    As PictureBox

   On Error GoTo Err_Proc

   If pWidth = 0 Then pWidth = pHeight
   If pHeight = 0 Then pHeight = pWidth
   If pWidth = 0 And pHeight = 0 Then ScaleToFit = True

   If mblnPrintFlag Then
      Load frmMultiPgPreview
      Set picBox = frmMultiPgPreview.picPrintPic
      picBox.Picture = newPic

      sngAspect = picBox.ScaleHeight / picBox.ScaleWidth

      If ScaleToFit Then
         sngWidth = Printer.ScaleWidth
         sngHeight = Printer.ScaleHeight
      Else
         sngWidth = pWidth
         sngHeight = pHeight
      End If

      If MaintainRatio Then
         If sngHeight / sngWidth > sngAspect Then
            sngHeight = sngAspect * sngWidth

            If LeftMargin = -1 Then
               sngX = Printer.ScaleLeft
            Else
               sngX = LeftMargin
            End If

            If TopMargin = -1 Then
               sngY = (Printer.ScaleHeight - sngHeight) / 2
            Else
               sngY = TopMargin
            End If

         Else
            sngWidth = sngHeight / sngAspect

            If LeftMargin = -1 Then
               sngX = (Printer.ScaleWidth - sngWidth) / 2
            Else
               sngX = LeftMargin
            End If

            If TopMargin = -1 Then
               sngY = Printer.ScaleTop
            Else
               sngY = TopMargin
            End If

         End If
      Else

         If LeftMargin = -1 Then
            sngX = Printer.ScaleLeft
         Else
            sngX = LeftMargin
         End If

         If TopMargin = -1 Then
            sngY = (Printer.ScaleHeight - sngHeight) / 2
         Else
            sngY = TopMargin
         End If

      End If

      Printer.PaintPicture picBox.Picture, sngX, sngY, sngWidth, sngHeight, , , , , vbSrcCopy
      Unload frmMultiPgPreview

   Else

      Set picBox = frmMultiPgPreview.picPrintPic
      picBox.Picture = newPic

      sngAspect = picBox.ScaleHeight / picBox.ScaleWidth

      If ScaleToFit Then
         sngWidth = mObjPrint.ScaleWidth
         sngHeight = mObjPrint.ScaleHeight
      Else
         sngWidth = pWidth
         sngHeight = pHeight
      End If

      If MaintainRatio Then
         If sngHeight / sngWidth > sngAspect Then
            sngHeight = sngAspect * sngWidth

            If LeftMargin = -1 Then
               sngX = mObjPrint.ScaleLeft
            Else
               sngX = LeftMargin
            End If

            If TopMargin = -1 Then
               sngY = (mObjPrint.ScaleHeight - sngHeight) / 2
            Else
               sngY = TopMargin
            End If

         Else
            sngWidth = sngHeight / sngAspect

            If LeftMargin = -1 Then
               sngX = (mObjPrint.ScaleWidth - sngWidth) / 2
            Else
               sngX = LeftMargin
            End If

            If TopMargin = -1 Then
               sngY = mObjPrint.ScaleTop
            Else
               sngY = TopMargin
            End If

         End If
      Else

         If LeftMargin = -1 Then
            sngX = mObjPrint.ScaleLeft
         Else
            sngX = LeftMargin
         End If

         If TopMargin = -1 Then
            sngY = (mObjPrint.ScaleHeight - sngHeight) / 2
         Else
            sngY = TopMargin
         End If

      End If

      mObjPrint.PaintPicture picBox.Picture, sngX, sngY, sngWidth, sngHeight, , , , , vbSrcCopy

      picBox.Picture = Nothing

   End If

   Set picBox = Nothing

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pPrintPicture"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pPrintRotate(Optional ByVal PrintVar As String = vbNullString, _
                        Optional ByVal Degree As Integer = 0, _
                        Optional ByVal LeftMargin As Single = -1)

  Dim f         As LOGFONT
  Dim hPrevFont As Long
  Dim hFont     As Long
  Dim ObjhDC    As Long
  Dim lngX      As Long
  Dim lngY      As Long
  
   '// By Diomidis Kiriakopoulos modified by Me

   If PrintVar = vbNullString Then Exit Sub

   If Degree > 359 Then
      Degree = 359
   ElseIf Degree < 0 Then
      Degree = 0
   End If

   If LeftMargin = -1 Then
      LeftMargin = Me.CurrentX
      If LeftMargin < mDLeftMar Then LeftMargin = mDLeftMar
   Else
      Me.CurrentX = LeftMargin
   End If

   On Local Error GoTo GetOut

   '// Save the hDC.
   If mblnPrintFlag Then
      Printer.Print vbNullString;
      ObjhDC = Printer.hdc
      f.lfHeight = (FontSize * -20) / Printer.TwipsPerPixelY
      lngX = Printer.CurrentX * 600 '// Printer dots/inch
      lngY = Printer.CurrentY * 600
   Else
      ObjhDC = mObjPrint.hdc
      f.lfHeight = (FontSize * -20) / Screen.TwipsPerPixelY
   End If

   f.lfEscapement = 10 * Degree            '// rotation angle, in tenths
   f.lfFaceName = Me.FontName & vbNullChar '// null terminated
   hFont = CreateFontIndirect(f)

   hPrevFont = SelectObject(ObjhDC, hFont)

   If mblnPrintFlag Then
      '// Draw the text.
      TextOut ObjhDC, lngX, lngY, PrintVar, Len(PrintVar)
   Else
      mObjPrint.Print PrintVar;
   End If

   '//  Clean up, restore original font
   hFont = SelectObject(ObjhDC, hPrevFont)
   DeleteObject hFont

GetOut:
   On Local Error GoTo 0

End Sub

Public Sub pQuarterSpace()

  Dim eFont As Integer
  Dim hFont As Integer

   On Error GoTo Err_Proc

   eFont = Me.FontSize
   hFont = eFont \ 4
   If hFont < 1 Then hFont = 1
   Me.FontSize = hFont
   pPrint
   Me.FontSize = eFont

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pQuarterSpace"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pRightJust(ByVal PrintVar As Variant, _
                      Optional ByVal RightMargin As Single = -1, _
                      Optional ByVal SameLine As Boolean = False)

  Dim sngTxtWidth As Single

   On Error GoTo Err_Proc

   sngTxtWidth = GetTextWidth(PrintVar)
   If RightMargin = -1 Then RightMargin = Me.CurrentX + sngTxtWidth
   Me.CurrentX = RightMargin - sngTxtWidth

   If SameLine Then
      If mblnPrintFlag Then
         Printer.Print PrintVar;
      Else
         mObjPrint.Print PrintVar;
      End If

   Else

      If mblnPrintFlag Then
         Printer.Print PrintVar
      Else
         mObjPrint.Print PrintVar
      End If

   End If

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pRightJust"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pRightTab(ByVal PrintVar As Variant, _
                     Optional ByVal xFromRight As Single = 0.1, _
                     Optional ByVal SameLine As Boolean = False)

   On Error GoTo Err_Proc

   Me.CurrentX = msngPgWidth - (GetTextWidth(PrintVar) + xFromRight)

   If SameLine Then
      If mblnPrintFlag Then
         Printer.Print PrintVar;
      Else
         mObjPrint.Print PrintVar;
      End If

   Else

      If mblnPrintFlag Then
         Printer.Print PrintVar
      Else
         mObjPrint.Print PrintVar
      End If

   End If

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pRightTab"
   Err.Clear
   Resume Exit_Here

End Sub

Public Property Get PrintCopies() As Integer

   On Local Error Resume Next
   PrintCopies = Printer.Copies
   On Local Error GoTo 0

End Property

Public Property Let PrintCopies(pNumber As Integer)

   On Local Error Resume Next
   Printer.Copies = pNumber
   On Local Error GoTo 0

End Property

Public Sub pSpaces(Optional ByVal RightMargin As Single = -1, _
                   Optional ByVal LeftMargin As Single = -1, _
                   Optional ByVal UseSymbol As Boolean = False)

  Dim xFontname  As String
  Dim xForeColor As Long
  Dim strTemp    As String

   On Error GoTo Err_Proc

   xFontname = Me.FontName
   xForeColor = Me.ForeColor

   If UseSymbol Then
      Me.FontName = "Symbol"
      strTemp = "\"
      Me.ForeColor = vbGreen
   Else
      strTemp = " "
   End If

   Call pCharacter(strTemp, RightMargin, LeftMargin)

ExitSpaceSub:
   Me.FontName = xFontname
   Me.ForeColor = xForeColor

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pSpaces"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pStartDoc()

   mintPageNumber = 0
   mstrTempDir = Environ("TEMP") & "\"

   On Local Error Resume Next

   If Not mblnNoPrinters Then
      '// Set the Printer's scale mode
      intPntrSM = Printer.ScaleMode
      Printer.ScaleMode = mintScaleMode
   
      '// Get the physical printable area
      msngPgWidth = Printer.ScaleWidth
      msngPgHeight = Printer.ScaleHeight
   
   Else
      msngPgWidth = 8
      msngPgHeight = 10.5
      MsgBox "No Printers Found!", vbExclamation
   End If
   
   If mblnPrintFlag Then
      '// Initialize printer
      Printer.Print vbNullString;
      
   Else
      '// Initialize the preview object
      Load frmMultiPgPreview
      Set mObjPrint = frmMultiPgPreview.Picture1

      '// Scale Object to Printer's printable area
      intObjSM = mObjPrint.ScaleMode
      mObjPrint.ScaleMode = mintScaleMode

      '// Full Page size (1440 twips = 1 inch or 567 twips = 1 centimeter)
      Select Case mintScaleMode
      Case vbCentimeters
         mObjPrint.Width = (msngPgWidth + 0.6) * 567
         mObjPrint.Height = (msngPgHeight + 0.6) * 567

      Case Else '// vbinches
         mObjPrint.Width = (msngPgWidth + 0.25) * 1440
         mObjPrint.Height = (msngPgHeight + 0.25) * 1440
      End Select

      '// Set default properties of the scroll bars
      frmMultiPgPreview.VScroll1.Max = Val(mObjPrint.Height * 0.5)
      frmMultiPgPreview.VScroll1.Min = -500
      frmMultiPgPreview.VScroll1.SmallChange = Val(frmMultiPgPreview.VScroll1.Max * 0.1)
      frmMultiPgPreview.VScroll1.LargeChange = Val(frmMultiPgPreview.VScroll1.Max * 0.5)

      frmMultiPgPreview.HScroll1.Max = Val(mObjPrint.Width * 0.25)
      frmMultiPgPreview.HScroll1.Min = -500
      frmMultiPgPreview.HScroll1.SmallChange = Val(frmMultiPgPreview.HScroll1.Max * 0.1)
      frmMultiPgPreview.HScroll1.LargeChange = Val(frmMultiPgPreview.HScroll1.Max * 0.5)

      '// Set default properties of the object to match printer
      mObjPrint.Scale (0, 0)-(msngPgWidth, msngPgHeight)
      mObjPrint.FontName = Printer.FontName
      mObjPrint.FontSize = Printer.FontSize
      mObjPrint.ForeColor = Printer.ForeColor
      mObjPrint.Picture = Nothing

   End If

End Sub

Public Sub pVerticalLine(Optional ByVal LeftPos As Single = -1, _
                         Optional ByVal TopPos As Single = -1, _
                         Optional ByVal BottomPos As Single = -1, _
                         Optional ByVal LineWidth As Integer = 0)

  Dim eDrawWidth As Integer
  Dim sngY       As Single
  Dim sngX       As Single
  Dim eFontSize  As Integer

   On Error GoTo Err_Proc

   eDrawWidth = Me.DrawWidth
   sngX = Me.CurrentX
   sngY = Me.CurrentY
   If LineWidth > 0 Then Me.DrawWidth = LineWidth

   If BottomPos = -1 Then
      eFontSize = Me.FontSize
      FontSize = 10
      BottomPos = msngPgHeight - (GetTextHeight * 2)
      Me.FontSize = eFontSize
   End If

   If LeftPos = -1 Then LeftPos = Me.CurrentX
   If LeftPos > msngPgWidth Then LeftPos = msngPgWidth - 0.01
   If TopPos = -1 Then TopPos = Me.CurrentY

   If mblnPrintFlag Then
      Printer.Line (LeftPos, TopPos)-(LeftPos, BottomPos)
   Else
      mObjPrint.Line (LeftPos, TopPos)-(LeftPos, BottomPos)
   End If

   'Me.CurrentX = sngX
   Me.CurrentY = sngY

   If LineWidth > 0 Then Me.DrawWidth = eDrawWidth

Exit_Here:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pVerticalLine"
   Err.Clear
   Resume Exit_Here

End Sub

Public Property Get ReportSubTitle() As String

   ReportSubTitle = mstrTitleSub

End Property

Public Property Let ReportSubTitle(ByVal vNewValue As String)

   mstrTitleSub = vNewValue

End Property

Public Property Let ReportTitle(ByVal vNewValue As String)

   mstrTitleMain = vNewValue

End Property

Public Property Get ReportTitle() As String

   ReportTitle = mstrTitleMain

End Property

Public Property Let SendToPrinter(ByVal vNewValue As Boolean)

   If Not mblnNoPrinters Then
      mblnPrintFlag = vNewValue
   End If

End Property

Public Property Get SendToPrinter() As Boolean

   SendToPrinter = mblnPrintFlag

End Property

Public Property Get DefaultLeftMargin() As Single

   DefaultLeftMargin = mDLeftMar
   
End Property

Public Property Let DefaultLeftMargin(ByVal vNewValue As Single)
   
   mDLeftMar = vNewValue

End Property
