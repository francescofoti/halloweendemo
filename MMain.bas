Attribute VB_Name = "MMain"
Option Compare Database
Option Explicit

Public Const APP_VERSION As String = "00.00.01"

Public Const IMSG_APP_ALREADY_RUNNING As String = "This application is already running."

Public Sub ShowError(ByVal psErrCtx As String, ByVal plErrNumber As Long, ByVal psErrDesc As String)
  Dim sMsg    As String
  sMsg = "An unexpected error occured." & vbCrLf & vbCrLf
  sMsg = sMsg & "Context: " & IIf(Len(psErrCtx), psErrCtx, "(unknown)") & vbCrLf
  sMsg = sMsg & "Number: " & plErrNumber & vbCrLf
  sMsg = sMsg & "Message:" & vbCrLf
  sMsg = sMsg & psErrDesc & vbCrLf
  MsgBox sMsg, vbCritical
End Sub

'TestDDELink
'Return 1 if this database (psDatabaseName) is already opened by another MSAccess instance.
'Found on the Internet ages ago, source lost, please tweet us @idevinfo if you can attribute it.
Function TestDDELink(ByVal psDatabaseName As String) As Integer
  Dim vDDEChannel As Long
  On Error Resume Next
  Application.SetOption "Ignore DDE Requests", True
  ' Open a channel between database instances
  vDDEChannel = DDEInitiate("MSAccess", psDatabaseName)
  'If the database is NOT already opened, then it will not be possible to create the DDE channel
  If Err Then
    TestDDELink = 0
  Else
    TestDDELink = 1
    DDETerminate vDDEChannel
    DDETerminateAll
  End If
  Application.SetOption ("Ignore DDE Requests"), False
End Function

'This is the starting point of the application
Public Function Main() As Integer
  Const LOCAL_ERR_CTX As String = "Main"
  Dim fOK         As Boolean
  
  On Error Resume Next
  
  'Test for single instance application
  If TestDDELink(Application.CurrentDb.Name) Then
    MsgBox IMSG_APP_ALREADY_RUNNING, vbInformation
    DoCmd.Quit acQuitSaveNone
  End If
  
  'Init Consoul library
  If Not FindConsoulLibrary() Then
    Exit Function
  End If
  
  'Launch demo
  EmojiDemo
  
Main_Exit:
  Exit Function
  
Main_Err:
  ShowError LOCAL_ERR_CTX, Err.Number, "Fatal error : " & Err.Description
  Resume Main_Exit
End Function

