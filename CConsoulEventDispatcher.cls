VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConsoulEventDispatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

' We receive mouse events via a sub in a module.
' This class registers ICsMouseEventSink target interface for hWnds
' and sends them to the registered ICsMouseEventSink target

'The event's sinks list hold interface pointers for each type
'of event that we'll dispatch for an hwnd
Private mlstEventSinks As CList

'Those are bit flags
Public Enum eConsoulEventType
  eCsUnsupportedEventType = 0
  eCsMouseEvent       'ICsMouseEventSink
  eCsVirtualLineEvent 'ICsVirtualLineEventSink
  eCsWmPaint          'ICsWmPaintEventSink
  eCsZonePaint        'ICsZonePaintEventSink
End Enum

Private Sub Class_Initialize()
  Set mlstEventSinks = New CList
  'We'll store the hwnd as a double, as in 64bits this is a LongPtr (or LongLong) and CList/CRow classes
  'do not know of these "new" data types (maybe will fix that one day)
  mlstEventSinks.ArrayDefine Array("hwnd", "ICsMouseEventSink", "ICsVirtualLineEventSink", _
                                   "ICsWmPaintEventSink", "ICsZonePaintEventSink"), _
                             Array(vbDouble, vbObject, vbObject, vbObject, vbObject)
End Sub

Private Sub Class_Terminate()
  Set mlstEventSinks = Nothing
End Sub

'We can register one interface at a time for an hwnd, ie
'we can call this method multiple times for an hwnd, one for
'interface we want to register for it.
'We must receive here the type of event from the caller, because we cannot
'determine and choose the interface type from pIIEventSink ourselves.
'UnregisterEventSink unregisters them all.
#If Win64 Then
Public Sub RegisterEventSink(ByVal phWnd As LongPtr, ByRef pIIEventSink As Object, ByVal peEventType As eConsoulEventType)
#Else
Public Sub RegisterEventSink(ByVal phWnd As Long, ByRef pIIEventSink As Object, ByVal peEventType As eConsoulEventType)
#End If
  Dim iHwnd       As Long
  
  If peEventType <> eCsUnsupportedEventType Then
    iHwnd = mlstEventSinks.Find("hwnd", phWnd)
    If iHwnd = 0 Then
      iHwnd = mlstEventSinks.AddValues(phWnd, Nothing, Nothing, Nothing, Nothing)
    End If
    Select Case peEventType
    Case eConsoulEventType.eCsMouseEvent
      Set mlstEventSinks("ICsMouseEventSink", iHwnd) = pIIEventSink
    Case eConsoulEventType.eCsVirtualLineEvent
      Set mlstEventSinks("ICsVirtualLineEventSink", iHwnd) = pIIEventSink
    Case eConsoulEventType.eCsWmPaint
      Set mlstEventSinks("ICsWmPaintEventSink", iHwnd) = pIIEventSink
    Case eConsoulEventType.eCsZonePaint
      Set mlstEventSinks("ICsZonePaintEventSink", iHwnd) = pIIEventSink
    Case Else
      Err.Raise 13, "RegisterEventSink", "Interface type is not supported"
    End Select
  Else
    Err.Raise 13, "RegisterEventSink", "Interface type is not supported"
  End If
End Sub

#If Win64 Then
Public Sub UnregisterEventSink(ByVal phWnd As LongPtr)
#Else
Public Sub UnregisterEventSink(ByVal phWnd As Long)
#End If
  Dim iHwnd       As Long
  iHwnd = mlstEventSinks.Find("hwnd", phWnd)
  If iHwnd > 0 Then
    mlstEventSinks.Remove iHwnd
  End If
End Sub

#If Win64 Then
Public Sub BroadcastMouseEvent( _
    ByVal phWnd As LongPtr, _
    ByVal piEvtCode As Integer, _
    ByVal pwParam As Integer, _
    ByVal piZoneID As Integer, _
    ByVal piRow As Integer, _
    ByVal piCol As Integer, _
    ByVal piPosX As Integer, _
    ByVal piPosY As Integer _
  )
#Else
Public Sub BroadcastMouseEvent( _
    ByVal phWnd As Long, _
    ByVal piEvtCode As Integer, _
    ByVal pwParam As Integer, _
    ByVal piZoneID As Integer, _
    ByVal piRow As Integer, _
    ByVal piCol As Integer, _
    ByVal piPosX As Integer, _
    ByVal piPosY As Integer _
  )
#End If
  Dim iiEventSink As ICsMouseEventSink
  Dim iResponse   As Integer
  Dim i           As Long
  On Error Resume Next
  For i = 1 To mlstEventSinks.Count
    If mlstEventSinks("hwnd", i) = phWnd Then
      Set iiEventSink = mlstEventSinks("ICsMouseEventSink", i)
      If Not iiEventSink Is Nothing Then
        iResponse = iiEventSink.OnMouseButton(phWnd, piEvtCode, pwParam, piZoneID, piRow, piCol, piPosX, piPosY)
        If CBool(iResponse) = True Then Exit For  'Called interface implementor handled the message and stopped propagation
        Set iiEventSink = Nothing
      End If
    End If
  Next i
End Sub

#If Win64 Then
Public Sub BroadcastVirtualLineEvent(ByVal phWnd As LongPtr, ByVal piLine As Long)
#Else
Public Sub BroadcastVirtualLineEvent(ByVal phWnd As Long, ByVal piLine As Long)
#End If
  Dim iiEventSink As ICsVirtualLineEventSink
  Dim iResponse   As Integer
  Dim i           As Long
  On Error Resume Next
  For i = 1 To mlstEventSinks.Count
    If mlstEventSinks("hwnd", i) = phWnd Then
      Set iiEventSink = mlstEventSinks("ICsVirtualLineEventSink", i)
      If Not iiEventSink Is Nothing Then
        iResponse = iiEventSink.OnVirtualLineGetData(phWnd, piLine)
        If CBool(iResponse) = True Then Exit For  'Called interface implementor handled the message and stopped propagation
        Set iiEventSink = Nothing
      End If
    End If
  Next i
End Sub

#If Win64 Then
Public Sub BroadcastConsolePaint( _
    ByVal phWnd As LongPtr, _
    ByVal pwCbkMode As Integer, _
    ByVal phDC As LongPtr, _
    ByVal lprcLinePos As LongPtr, _
    ByVal lprcLineRect As LongPtr, _
    ByVal lprcPaint As LongPtr _
  )
#Else
Public Sub BroadcastConsolePaint( _
    ByVal phWnd As Long, _
    ByVal pwCbkMode As Integer, _
    ByVal phDC As Long, _
    ByVal lprcLinePos As Long, _
    ByVal lprcLineRect As Long, _
    ByVal lprcPaint As Long _
  )
#End If
  Dim iiEventSink As ICsWmPaintEventSink
  Dim iResponse   As Integer
  Dim i           As Long
  On Error Resume Next
  For i = 1 To mlstEventSinks.Count
    If mlstEventSinks("hwnd", i) = phWnd Then
      Set iiEventSink = mlstEventSinks("ICsWmPaintEventSink", i)
      If Not iiEventSink Is Nothing Then
        iResponse = iiEventSink.OnConsolePaint(phWnd, pwCbkMode, phDC, lprcLinePos, lprcLineRect, lprcPaint)
        If CBool(iResponse) = True Then Exit For  'Called interface implementor handled the message and stopped propagation
        Set iiEventSink = Nothing
      End If
    End If
  Next i
End Sub

#If Win64 Then
Public Sub BroadcastZonePaint( _
    ByVal phWnd As LongPtr, _
    ByVal phDC As LongPtr, _
    ByVal piZoneID As Integer, _
    ByVal piLine As Long, _
    ByVal piLeft As Integer, _
    ByVal piTop As Integer, _
    ByVal piRight As Integer, _
    ByVal piBottom As Integer _
  )
#Else
Public Sub BroadcastZonePaint( _
    ByVal phWnd As Long, _
    ByVal phDC As Long, _
    ByVal piZoneID As Integer, _
    ByVal piLine As Long, _
    ByVal piLeft As Integer, _
    ByVal piTop As Integer, _
    ByVal piRight As Integer, _
    ByVal piBottom As Integer _
  )
#End If
  Dim iiEventSink As ICsZonePaintEventSink
  Dim iResponse   As Integer
  Dim i           As Long
  On Error Resume Next
  For i = 1 To mlstEventSinks.Count
    If mlstEventSinks("hwnd", i) = phWnd Then
      Set iiEventSink = mlstEventSinks("ICsZonePaintEventSink", i)
      If Not iiEventSink Is Nothing Then
        iResponse = iiEventSink.OnZonePaint(phWnd, phDC, piZoneID, piLine, piLeft, piTop, piRight, piBottom)
        If CBool(iResponse) = True Then Exit For  'Called interface implementor handled the message and stopped propagation
        Set iiEventSink = Nothing
      End If
    End If
  Next i
End Sub

