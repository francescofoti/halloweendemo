VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICsVirtualLineEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

#If Win64 Then
Public Function OnVirtualLineGetData(ByVal phWnd As LongPtr, ByVal piLine As Long) As Integer

End Function
#Else
Public Function OnVirtualLineGetData(ByVal phWnd As Long, ByVal piLine As Long) As Integer

End Function
#End If
