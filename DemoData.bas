Attribute VB_Name = "mDemodata"
Option Explicit

'####################################################################################
'Title:     DemoData
'Author:    Richard Mewett
'Created:   31/05/06

'Just a simple module for use with demos of my UserControls - nothing fancy here!

Public Enum NameTypeEnum
   ntRandom = 0
   ntMale = 1
   ntFemale = 2
End Enum

'These are for generating the demo data
Private Const M_FORENAMES = "Alan,Alfie,Andrew,Ben,Bill,Bob,Boris,Brian,Charles,Chris,David,Gavin,Geoff,Grant,Harry,Ian," & _
   "James,Jon,Mark,Matthew,Michael,Patrick,Paul,Peter,Richard,Robert,Samuel,Simon,Tony,Trevor,William"
Private Const F_FORENAMES = "Alicia,Alison,Amanda,Barbara,Caroline,Charlotte,Dawn,Hannah,Harriet,Hayley,Jane,Jennifer,Karen," & _
   "Katie,Kerry,Kim,Lara,Laura,Lucy,Mary,Mellisa,Patricia,Paula,Rachel,Sarah,Stephanie,Susan,Tracy,Vanessa"
Private Const SURNAMES = "Anderson-Allen,Black,Bloggs,Brown,Clarke,Cole,Davis,Dawson,Evans,Gate,Johnson,Jones,Lawson,Lee," & _
   "Richards,Ryan,Smith,Stephens,Temple,Turner,Wallace,White,Williams"

Private Const JOBS = "Accountant,Architect,Artist,Banker,Builder,Carpenter,Dentist,Director,Doctor,Engineer,Estate Agent," & _
   "Fire Fighter,Gardener,Manager,Mechanic,Miner,Nurse,Optician,Pilot,Plumber,Police,Programmer,Scientist,Secretary,Shop" & _
   " Assistant,Solicitor,Surgeon,Teacher,Truck Driver,Vet"

Private mCalled As Boolean

Private mMF() As String
Private mFF() As String
Private mSurnames() As String

Private mJobs() As String

Public gclrBack As OLE_COLOR

Public Function GetForename(Optional nType As NameTypeEnum) As String

   Initialise

   Select Case nType
   Case ntRandom

      If RandomInt(0, 1) = 0 Then
         GetForename = mMF(RandomInt(LBound(mMF), UBound(mMF)))
      Else
         GetForename = mFF(RandomInt(LBound(mFF), UBound(mFF)))
      End If

   Case ntMale
      GetForename = mMF(RandomInt(LBound(mMF), UBound(mMF)))

   Case ntFemale
      GetForename = mFF(RandomInt(LBound(mFF), UBound(mFF)))

   End Select

End Function

Public Function GetJobName(Optional Index As Long = -1) As String

   Initialise

   If Index = -1 Then
      GetJobName = mJobs(RandomInt(LBound(mJobs), UBound(mJobs)))
   Else
      GetJobName = mJobs(Index)
   End If
   
End Function

Public Function GetSurname() As String

   Initialise

   GetSurname = mSurnames(RandomInt(LBound(mSurnames), UBound(mSurnames)))

End Function

Private Sub Initialise()

   If Not mCalled Then
      mCalled = True
      Randomize Timer

      mMF() = Split(M_FORENAMES, ",")
      mFF() = Split(F_FORENAMES, ",")
      mSurnames() = Split(SURNAMES, ",")

      mJobs() = Split(JOBS, ",")
   End If

End Sub

Public Function JobCount() As Long

   Initialise

   JobCount = UBound(mJobs)

End Function

Public Sub Main()

   frmMain.Show

End Sub

Public Function RandomInt(lowerbound As Long, upperbound As Long) As Long

   RandomInt = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)

End Function

Public Function rVal(ByVal vString As String) As Double

   '// MLH - New
   '// Returns the numbers contained in a string as a numeric value
   '// The Val function recognizes only the period (.) as a valid decimal separator.
   '// The CDbl errors on empty strings or values containing non-numeric values
   '// Returns the numbers contained in a string as a numeric value

  Dim lngI     As Long
  Dim lngS     As Long
  Dim bytAscV  As Byte
  Dim strTemp  As String
  
  On Error Resume Next

   vString = Trim$(UCase$(vString))
   
   If Left$(vString, 4) = "TRUE" Then
      rVal = True
      
   ElseIf Left$(vString, 5) = "FALSE" Then
      rVal = False
   
   Else
      Select Case Left$(vString, 2) '// Hex or Octal?
      Case Is = "&H", Is = "&O"
         lngS = 3
         strTemp = Left$(vString, 2)
      Case Else
         lngS = 1
      End Select
      
      For lngI = lngS To Len(vString)
         bytAscV = Asc(Mid$(vString, lngI, 1))
         Select Case bytAscV
         Case 48 To 57 '// 1234567890
            strTemp = strTemp & Mid$(vString, lngI, 1)
         
         Case 44, 45, 46 '// , - .
            strTemp = strTemp & Mid$(vString, lngI, 1)
         
         Case 36 '// $
            '// Ignore
            
         Case Is > 57, Is < 44
            If Left$(strTemp, 2) = "&H" Then '// Hex Values ?
               Select Case bytAscV
               Case 65 To 70 '// ABCDEF
                  strTemp = strTemp & Mid$(vString, lngI, 1)
               Case Else
                  Exit For
               End Select
            Else
               Exit For
            End If
         End Select
      Next lngI
      
      If LenB(strTemp) Then
         rVal = CDbl(strTemp)
      End If
   End If
   
   On Error GoTo 0

End Function

'''Public Function GetNameOfPerson(Optional nType As NameTypeEnum) As String

   '''
   '''   Select Case nType
   '''    Case ntRandom
   '''
   '''      If RandomInt(0, 1) = 0 Then
   '''         GetNameOfPerson = GetForename(ntMale) & " " & GetSurname()
   '''       Else
   '''         GetNameOfPerson = GetForename(ntFemale) & " " & GetSurname()
   '''      End If
   '''
   '''    Case ntMale
   '''      GetNameOfPerson = GetForename(ntMale) & " " & GetSurname()
   '''
   '''    Case ntFemale
   '''      GetNameOfPerson = GetForename(ntFemale) & " " & GetSurname()
   '''
   '''   End Select
   '''

   '''End Function

