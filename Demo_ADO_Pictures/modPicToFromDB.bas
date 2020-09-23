Attribute VB_Name = "modPicToFromDB"
Option Explicit

Public Sub GetPicFromDB(ByRef MySet As ADODB.Recordset, _
                        ByVal sField As String)
                
  Dim smPic       As ADODB.Stream
  Dim strFileName As String
  
    On Error Resume Next
    
    strFileName = App.Path & "\Temp.bmp"
          
    '// Create an instance of the stream object
    Set smPic = New ADODB.Stream
    '// set the type to binary to load the image as a binary stream
    smPic.Type = adTypeBinary
    smPic.Open
    '// Load the binary image data from the DB into the stream object
    smPic.Write MySet.Fields(sField).Value

    '// Check the size of the ado stream to make sure there is data
    If smPic.Size > 0 Then
        '// Write the content of the stream object to a file
        '// The file will be created if doesn't exists. Otherwise over writes the existing file
        smPic.SaveToFile strFileName, adSaveCreateOverWrite
       
        '// Load the temp Picture into the Image control
        frmAdvanced.picWorking.Picture = LoadPicture(strFileName)
        Kill strFileName
        
    Else
      '// No picture found in DB
      frmAdvanced.picWorking.Picture = Nothing
    End If
    
    '// Close and destroy the stream object
    smPic.Close
    Set smPic = Nothing
    
End Sub

Public Sub SavePicToDB(ByRef MySet As ADODB.Recordset, _
                       ByVal sField As String, _
                       ByRef oTargetObj As StdPicture)

  Dim strFileName As String
  Dim smPic       As ADODB.Stream

   On Error Resume Next

   Set smPic = New ADODB.Stream
   '// set the type to binary to load the image as a binary stream
   smPic.Type = adTypeBinary
   smPic.Open

   strFileName = App.Path & "\Temp.bmp"
   SavePicture oTargetObj, strFileName

   '// Load the content of the picture into the stream object
   smPic.LoadFromFile strFileName
   '// Check the size of the ado stream to make sure there is data
   If smPic.Size > 0 Then
      MySet.Fields(sField) = smPic.Read
   End If

   '// Close and destroy the stream object reference
   smPic.Close
   Set smPic = Nothing

   Kill strFileName

End Sub

