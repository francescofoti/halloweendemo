VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICsZonePaintEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

#If Win64 Then
Public Function OnZonePaint( _
    ByVal phWnd As LongPtr, _
    ByVal phDC As LongPtr, _
    ByVal piZoneID As Integer, _
    ByVal piLine As Long, _
    ByVal piLeft As Integer, _
    ByVal piTop As Integer, _
    ByVal piRight As Integer, _
    ByVal piBottom As Integer _
  ) As Integer
  
End Function
#Else
Public Function OnZonePaint( _
    ByVal phWnd As Long, _
    ByVal phDC As Long, _
    ByVal piZoneID As Integer, _
    ByVal piLine As Long, _
    ByVal piLeft As Integer, _
    ByVal piTop As Integer, _
    ByVal piRight As Integer, _
    ByVal piBottom As Integer _
  ) As Integer
  
End Function
#End If

