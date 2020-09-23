Attribute VB_Name = "modADO"
Option Explicit

Public gstrDBName           As String

Public gdbSCB               As ADODB.Connection
Public grsSCB               As ADODB.Recordset

Public Function ADOFindFirst(ByRef MySet As ADODB.Recordset, _
                             ByVal Filter As String) As Boolean
  
  Dim mhRS As ADODB.Recordset
  Dim mhMatch As Boolean

    On Error GoTo Err_Proc
    
    Set mhRS = New ADODB.Recordset
    Set mhRS = MySet.Clone
    mhRS.Filter = Filter
    
    If mhRS.RecordCount > 0 Then
        mhRS.MoveFirst
        MySet.Bookmark = mhRS.Bookmark
        mhMatch = True
    Else
        If MySet.RecordCount > 0 Then
            MySet.MoveLast
            MySet.MoveNext
        End If
        mhMatch = False
    End If
    mhRS.Close
    
    ADOFindFirst = mhMatch
    
Exit Function

Err_Proc:
    On Error Resume Next
    mhRS.Close
    Set mhRS = Nothing
    ADOFindFirst = False

End Function

Public Function ADORecordCount(ByRef MySet As ADODB.Recordset) As Long

  Dim BkMark As Variant
  Dim RC     As Long

   On Local Error Resume Next

   With MySet
      BkMark = .Bookmark
      .MoveLast
      RC = .RecordCount
   End With 'MySet

   If RC = 1 Then
      If IsNull(MySet.Fields(0)) Then
         RC = 0
      End If

   End If
   
   ADORecordCount = RC
   MySet.Bookmark = BkMark
   On Local Error GoTo 0

End Function

Public Sub OpenDB(ByRef vConnection As ADODB.Connection, ByVal vDBName As String)

   Set vConnection = New ADODB.Connection
   vConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vDBName & ";Jet OLEDB:Database Password=;Jet" & _
      " OLEDB:Engine Type=5;"

End Sub

Public Sub OpenRS(ByRef vRecordset As ADODB.Recordset, _
                  ByVal oSourceTable As String, _
                  ByRef vConnection As ADODB.Connection, _
                  Optional oCursorType As CursorTypeEnum = adOpenStatic, _
                  Optional oLockType As LockTypeEnum = adLockOptimistic, _
                  Optional ByVal oOptions As Integer = -1)

   Set vRecordset = New ADODB.Recordset
   vRecordset.Open oSourceTable, vConnection, oCursorType, oLockType, oOptions

End Sub

