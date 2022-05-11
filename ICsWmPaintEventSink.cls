VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICsWmPaintEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

'Return 1 from implementation to stop broadcast
#If Win64 Then
Public Function OnConsolePaint( _
  ByVal phWnd As LongPtr, _
  ByVal pwCbkMode As Integer, _
  ByVal phDC As LongPtr, _
  ByVal lprcLinePos As LongPtr, _
  ByVal lprcLineRect As LongPtr, _
  ByVal lprcPaint As LongPtr) As Integer

End Function
#Else
Public Function OnConsolePaint( _
  ByVal phWnd As Long, _
  ByVal pwCbkMode As Integer, _
  ByVal phDC As Long, _
  ByVal lprcLinePos As Long, _
  ByVal lprcLineRect As Long, _
  ByVal lprcPaint As Long) As Integer

End Function
#End If
