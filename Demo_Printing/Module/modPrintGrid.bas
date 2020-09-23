Attribute VB_Name = "modPrintGrid"
Option Explicit
'// Here is an example of grid printing.
'// It does not print pictures in a cell or handle columns that fall off the edge of the paper.

Public cPrint  As clsMultiPgPreview

Public Sub PrintGrid(ByRef rGridObj As LynxGrid, _
                     Optional ByVal vbVisibleColsOnly As Boolean = True, _
                     Optional ByVal vbVisibleRowsOnly As Boolean = True)

  Dim lngR        As Long
  Dim lngC        As Long
  Dim lngC1       As Long
  Dim sngTemp1    As Single
  Dim sngTemp2    As Single
  Dim strText     As String
  Dim lMaxCol     As Long
  Dim lColp()     As Single
  Dim sngY1       As Single
  Dim sngY2       As Single
  Const C_LeftM   As Single = 0.05
  Const C_PIXPINCH As Single = 0.00075
  
   On Local Error GoTo Err_Proc

   Set cPrint = New clsMultiPgPreview
   cPrint.Orientation = PageLandscape
   cPrint.SendToPrinter = False

SendToPrinter:
   Screen.MousePointer = vbHourglass
   DoEvents
   
   cPrint.pStartDoc
   
   With rGridObj
   
      lMaxCol = .Cols - 1
   
      ReDim lColp(0 To .Cols - 1, 0 To 1) As Single
      
      '// Header
      GoSub PrintHeader
   
      '// Grid Data
      For lngR = 0 To .Rows - 1
         If Not vbVisibleRowsOnly Or .RowVisible(lngR) Then
            For lngC = 0 To lMaxCol
               If vbVisibleColsOnly Then
                  If .ColVisible(lngC) Then
                     cPrint.ForeColor = .CellForeColor(lngR, lngC)
                     cPrint.BackColor = .CellBackColor(lngR, lngC)
                     If LenB(.ColFormat(lngC)) Then
                        sngY1 = cPrint.pMultiline(Format$(.CellText(lngR, lngC), .ColFormat(lngC)), lColp(lngC, 0), lColp(lngC, 1), , True)
                     Else
                        sngY1 = cPrint.pMultiline(.CellText(lngR, lngC), lColp(lngC, 0), lColp(lngC, 1), , True)
                     End If
                     cPrint.ForeColor = vbBlack
                     cPrint.BackColor = vbWhite
                  End If
               
               Else
                  cPrint.ForeColor = .CellForeColor(lngR, lngC)
                  cPrint.BackColor = .CellBackColor(lngR, lngC)
                  If LenB(.ColFormat(lngC)) Then
                     sngY1 = cPrint.pMultiline(Format$(.CellText(lngR, lngC), .ColFormat(lngC)), lColp(lngC, 0), lColp(lngC, 1), , True)
                  Else
                     sngY1 = cPrint.pMultiline(.CellText(lngR, lngC), lColp(lngC, 0), lColp(lngC, 1), , True)
                  End If
                  cPrint.ForeColor = vbBlack
                  cPrint.BackColor = vbWhite
               End If
               If sngY1 > sngY2 Then sngY2 = sngY1
            Next lngC
         End If

         cPrint.CurrentY = sngY2
         sngY1 = 0
         sngY2 = 0
         cPrint.pLine C_LeftM, lColp(lMaxCol, 1) + C_LeftM
         If cPrint.pEndOfPage Then
            cPrint.pFooter
            cPrint.pNewPage
            GoSub PrintHeader
         End If
      Next lngR

      '// Totals Line
      If .TotalsLineShow Then
         cPrint.pLine C_LeftM, lColp(lMaxCol, 1) + C_LeftM, 4
         For lngC = 0 To lMaxCol
            If vbVisibleColsOnly Then
               If .ColVisible(lngC) Then
                  If .ColType(lngC) = lgNumeric Then
                     If LenB(.ColFormat(lngC)) Then
                        cPrint.pPrint Format$(.TotalsCol(lngC), .ColFormat(lngC)), lColp(lngC, 0), True
                     Else
                        cPrint.pPrint .TotalsCol(lngC), lColp(lngC, 0), True
                     End If
                  End If
               End If
            Else
               If .ColType(lngC) = lgNumeric Then
                  If LenB(.ColFormat(lngC)) Then
                     cPrint.pPrint Format$(.TotalsCol(lngC), .ColFormat(lngC)), lColp(lngC, 0), True
                  Else
                     cPrint.pPrint .TotalsCol(lngC), lColp(lngC, 0), True
                  End If
               End If
            End If
         Next lngC
         cPrint.pPrint
         cPrint.pLine C_LeftM, lColp(lMaxCol, 1) + C_LeftM, 4
      End If
   
      Screen.MousePointer = vbDefault
      cPrint.pFooter
      cPrint.pEndDoc
      If cPrint.SendToPrinter Then
         GoTo SendToPrinter
      End If
   
   End With
   
   Set cPrint = Nothing
   Erase lColp
   Screen.MousePointer = vbDefault
   Exit Sub


PrintHeader:
   With rGridObj
      cPrint.pFontName
      cPrint.FontSize = 8
      cPrint.FontBold = True
      
      If LenB(.Caption) Then
         cPrint.FontSize = 12
         cPrint.pCenter .Caption
         cPrint.pQuarterSpace
         cPrint.FontSize = 8
      End If
   
      sngTemp1 = C_LeftM
      For lngC = 0 To lMaxCol
         If vbVisibleColsOnly Then
            If .ColVisible(lngC) Then
               cPrint.pVerticalLine sngTemp1
               lColp(lngC, 0) = sngTemp1 + C_LeftM
               sngTemp2 = (.COLWIDTH(lngC) * C_PIXPINCH) ' / 2.2)
               If .ColType(lngC) = lgBoolean Then
                  sngTemp2 = 0.4
               ElseIf sngTemp2 < 0.4 Then
                  sngTemp2 = 0.4
               End If
               sngTemp1 = sngTemp1 + sngTemp2
               lColp(lngC, 1) = sngTemp1 - C_LeftM
               sngY1 = cPrint.pMultiline(.ColHeading(lngC), lColp(lngC, 0), lColp(lngC, 1), , True)
            End If
   
         Else
            cPrint.pVerticalLine sngTemp1
            lColp(lngC, 0) = sngTemp1 + C_LeftM
            sngTemp2 = (.COLWIDTH(lngC) * C_PIXPINCH) ' / 2.2)
            If .ColType(lngC) = lgBoolean Then
               sngTemp2 = 0.4
            ElseIf sngTemp2 < 0.4 Then
               sngTemp2 = 0.4
            End If
            sngTemp1 = sngTemp1 + sngTemp2
            lColp(lngC, 1) = sngTemp1 - C_LeftM
            sngY1 = cPrint.pMultiline(.ColHeading(lngC), lColp(lngC, 0), lColp(lngC, 1), , True)
         End If
         If sngY1 > sngY2 Then sngY2 = sngY1
      Next lngC
      
      cPrint.pVerticalLine sngTemp1
      cPrint.CurrentY = sngY2
      sngY1 = 0
      sngY2 = 0
      cPrint.FontBold = False
      cPrint.pDoubleLine C_LeftM, lColp(lMaxCol, 1) + C_LeftM
   End With
   
   Return
   
Err_Proc:
   MsgBox "Error# " & Err.Number & vbNewLine & Err.Description, vbCritical, "LynxGrid.Export"
   Close

End Sub


