Attribute VB_Name = "MFileSystem"
Option Compare Database
Option Explicit

'https://francescofoti.com/2017/10/manipulating-filenames-in-excel-or-access-or-vba/
Private Const PATH_SEP      As String = "\"
Private Const PATH_SEP_INV  As String = "/"
Private Const EXT_SEP       As String = "."
Private Const DRIVE_SEP     As String = ":"

Public Function CombinePath(ByVal psPath1 As String, ByVal psFilename As String) As String
  If left$(psFilename, 1) <> PATH_SEP Then
    CombinePath = NormalizePath(psPath1) & psFilename
  Else
    CombinePath = DenormalizePath(psPath1) & psFilename
  End If
End Function

' Make sure path ends in a backslash.
'Private function as we want to promote CombinePath().
Public Function NormalizePath(ByVal spath As String) As String
  If Right$(spath, 1) <> PATH_SEP Then
    NormalizePath = spath & PATH_SEP
  Else
    NormalizePath = spath
  End If
End Function

' Make sure path doesn't end in a backslash
Public Function DenormalizePath(ByVal spath As String) As String
  If Right$(spath, 1) = PATH_SEP Then
    spath = left$(spath, Len(spath) - 1)
  End If
  DenormalizePath = spath
End Function

' Test the existence of a file
Public Function ExistFile(psSpec As String) As Boolean
  On Error Resume Next
  Call FileLen(psSpec)
  ExistFile = (Err.Number = 0&)
End Function

Public Function StripFileName(ByVal psFilename As String) As String
  Dim i           As Long
  Dim fLoop       As Boolean
  Dim sChar       As String * 1
  
  i = Len(psFilename)
  If i Then fLoop = True
  While fLoop
    If i > 0 Then
      sChar = Mid$(psFilename, i, 1)
      If (sChar = PATH_SEP) Or (sChar = DRIVE_SEP) Or (sChar = PATH_SEP_INV) Then fLoop = False
    End If
    If i > 1& Then
      i = i - 1&
    Else
      i = 0&
      fLoop = False
    End If
  Wend
  If i Then
    StripFileName = left$(psFilename, i)
  Else
    StripFileName = ""
  End If
End Function


