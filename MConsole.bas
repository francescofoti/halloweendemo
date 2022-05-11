Attribute VB_Name = "MConsole"
'MConsole.bas
'
'(C) 2018, devinfo.net, Développement Informatique Services, Francesco Foti
'VB/A wrappers for the Consoul library
'Library docs: https://consoul.net/docs/reference/index.htm
'VBA SDK docs: https://consoul.net/docs/sdk/vba/index.htm

#If MSACCESS Then
Option Compare Database
#End If

Option Explicit

Private Const FORMNAME_CONSOLE As String = "Console"

'The "Console" form will act as a console provider globally
'by implementing IConsoleOutput and announcing itself here.
Private miiConOut As IConsoleOutput

Public Function IICOnsoleOutput() As IConsoleOutput
  Set IICOnsoleOutput = miiConOut
End Function

Public Sub SetConsoleOutput(ByRef piiConOut As IConsoleOutput)
  Set miiConOut = piiConOut
End Sub

Public Sub ConOut(ByVal psText As String)
  If Not miiConOut Is Nothing Then
    miiConOut.ConOut psText
  Else
    Debug.Print psText;
  End If
End Sub

Public Function ConOutLn(ByVal psText As String, Optional ByVal piQBColorText As Variant) As Integer
  If Not miiConOut Is Nothing Then
    If Not IsMissing(piQBColorText) Then
      ConOutLn = miiConOut.ConOutLn(psText, piQBColorText)
    Else
      ConOutLn = miiConOut.ConOutLn(psText)
    End If
  Else
    Debug.Print psText
  End If
End Function

Public Sub OpenConsoleWindow()
  On Error Resume Next
  DoCmd.OpenForm FORMNAME_CONSOLE
End Sub

Public Sub CloseConsoleWindow()
  On Error Resume Next
  DoCmd.Close acForm, FORMNAME_CONSOLE
End Sub

Public Function GetConsoleForm() As Form_Console
  On Error Resume Next
  Set GetConsoleForm = Forms(FORMNAME_CONSOLE)
End Function

Public Function GetConsole() As CConsoul
  On Error Resume Next
  Set GetConsole = Forms(FORMNAME_CONSOLE).Console
End Function

