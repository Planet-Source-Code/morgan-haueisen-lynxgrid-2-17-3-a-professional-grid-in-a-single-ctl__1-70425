Attribute VB_Name = "modCreateDB"
Option Explicit

Private CAT As ADOX.Catalog

Private Sub CreateIndexes()

   On Error GoTo ErrTrap
  Dim IDX As ADOX.Index

   ' ===[Create Index 'PrimaryKey']===
   Set IDX = New ADOX.Index

   With IDX
      .Name = "PrimaryKey"
      .Columns.Append "Key"
      .PrimaryKey = True
      .Unique = True
      .Clustered = False
      .IndexNulls = adIndexNullsDisallow
   End With

   CAT.Tables("TestScores").Indexes.Append IDX
   ' ===[Create Index 'Key']===
   Set IDX = New ADOX.Index

   With IDX
      .Name = "Key"
      .Columns.Append "Key"
      .PrimaryKey = False
      .Unique = False
      .Clustered = False
      .IndexNulls = adIndexNullsAllow
   End With

   CAT.Tables("TestScores").Indexes.Append IDX

   Set IDX = Nothing

   Exit Sub

ErrTrap:
   'MsgBox Err.Number & " / " & Err.Description,,"Error In CreateIndexes"
   'Exit Sub
   'Resume

End Sub

Public Sub CreateMDB(ByVal dbPathFilename As String)

   On Error GoTo ErrTrap

   Set CAT = New ADOX.Catalog

   '/* Engine Type = 4; (Access97)
   '/* Engine Type = 5; (Access2000)

   CAT.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPathFilename & ";Jet OLEDB:Database Password=;Jet" & _
      " OLEDB:Engine Type=5;"

   Call CreateTables
   Call CreateIndexes

   Set CAT = Nothing

   Call LoadDB(dbPathFilename)

   '  MsgBox "Database created.", vbApplicationModal + vbInformation, App.Title
   Exit Sub

ErrTrap:
   '  MsgBox Err.Number & " / " & Err.Description

End Sub

Private Sub CreateTables()

   On Error GoTo ErrTrap
  Dim TBL As ADOX.Table

   ' ===[Create Table 'TestScores']===
   Set TBL = New ADOX.Table
   Set TBL.ParentCatalog = CAT

   With TBL
      .Name = "TestScores"
      .Columns.Append "Key", adInteger, 0
      .Columns("Key").Properties("AutoIncrement") = True
      .Columns("Key").Properties("NullAble") = True

      .Columns.Append "NameLast", adVarWChar, 50
      .Columns("NameLast").Properties("NullAble") = True
      .Columns("NameLast").Properties("Jet OLEDB:Allow Zero Length") = True

      .Columns.Append "NameFirst", adVarWChar, 50
      .Columns("NameFirst").Properties("NullAble") = True
      .Columns("NameFirst").Properties("Jet OLEDB:Allow Zero Length") = True

      .Columns.Append "Score", adInteger, 0
      .Columns("Score").Properties("NullAble") = True
      .Columns("Score").Properties("Default") = 0

   End With

   CAT.Tables.Append TBL

   Set TBL = Nothing

   Exit Sub

ErrTrap:
   'MsgBox Err.Number & " / " & Err.Description,,"Error In CreateTables"
   'Exit Sub
   'Resume

End Sub

Private Sub LoadDB(ByVal dbPathFilename As String)

  Dim MyDB  As ADODB.Connection
  Dim MySet As ADODB.Recordset
  Dim lngI  As Long

   Randomize

   Call OpenDB(MyDB, dbPathFilename)

   Call OpenRS(MySet, "Select * From TestScores", MyDB)
   
   MyDB.BeginTrans

   For lngI = 0 To 500

      With MySet
         .AddNew
         .Fields("Score") = RandomInt(50, 100)

         If RandomInt(0, 1) = 0 Then
            .Fields("NameFirst") = GetForename(ntMale)
          Else
            .Fields("NameFirst") = GetForename(ntFemale)
         End If

         .Fields("NameLast") = GetSurname()
         .Update
      End With

   Next lngI

   MyDB.CommitTrans
   MySet.Close
   MyDB.Close

End Sub

