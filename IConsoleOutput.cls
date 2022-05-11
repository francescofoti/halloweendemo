VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConsoleOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'IConsoleOutput.bas
'
'(C) 2018, DIS Développement Informatique Services, Francesco Foti
'VB/A wrappers for the Consoul library (https://www.devinfo.net/projects/consoul.htm)

#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

Public Sub ConOut(ByVal psInfo As String)

End Sub

Public Function ConOutLn(ByVal psInfo As String, Optional ByVal piQBColorText As Variant) As Integer

End Function

