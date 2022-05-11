Attribute VB_Name = "MCCallbacks"
'MCCallbacks.bas
'
'(C) 2018, devinfo.net, Développement Informatique Services, Francesco Foti
'VB/A wrappers for the Consoul library
'Library docs: https://consoul.net/docs/reference/index.htm
'VBA SDK docs: https://consoul.net/docs/sdk/vba/index.htm
#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

Public ConsoulEventDispatcher As New CConsoulEventDispatcher

#If Win64 Then
Public Function OnConsoulMouseButton( _
      ByVal phWnd As LongPtr, _
      ByVal piEvtCode As Integer, _
      ByVal pwParam As Integer, _
      ByVal piZoneID As Integer, _
      ByVal piRow As Integer, _
      ByVal piCol As Integer, _
      ByVal piPosX As Integer, _
      ByVal piPosY As Integer _
  ) As Integer
#Else
Public Function OnConsoulMouseButton( _
    ByVal phWnd As Long, _
    ByVal piEvtCode As Integer, _
    ByVal pwParam As Integer, _
    ByVal piZoneID As Integer, _
    ByVal piRow As Integer, _
    ByVal piCol As Integer, _
    ByVal piPosX As Integer, _
    ByVal piPosY As Integer _
  ) As Integer
#End If
  On Error Resume Next
  ConsoulEventDispatcher.BroadcastMouseEvent phWnd, piEvtCode, pwParam, piZoneID, piRow, piCol, piPosX, piPosY
End Function

#If Win64 Then
Public Function OnConsoulVirtualLine( _
    ByVal phWnd As LongPtr, _
    ByVal piLine As Long _
  ) As Integer
#Else
Public Function OnConsoulVirtualLine( _
    ByVal phWnd As Long, _
    ByVal piLine As Long _
  ) As Integer
#End If
  On Error Resume Next
  ConsoulEventDispatcher.BroadcastVirtualLineEvent phWnd, piLine
End Function

#If Win64 Then
Public Function OnConsoulWmPaint( _
    ByVal phWnd As LongPtr, _
    ByVal pwCbkMode As Integer, _
    ByVal phDC As LongPtr, _
    ByVal lprcLinePos As LongPtr, _
    ByVal lprcLineRect As LongPtr, _
    ByVal lprcPaint As LongPtr _
  ) As Integer
#Else
Public Function OnConsoulWmPaint( _
    ByVal phWnd As Long, _
    ByVal pwCbkMode As Integer, _
    ByVal phDC As Long, _
    ByVal lprcLinePos As Long, _
    ByVal lprcLineRect As Long, _
    ByVal lprcPaint As Long _
  ) As Integer
#End If
  On Error Resume Next
  ConsoulEventDispatcher.BroadcastConsolePaint phWnd, pwCbkMode, phDC, lprcLinePos, lprcLineRect, lprcPaint
End Function

#If Win64 Then
Public Function OnConsoulZonePaint( _
    ByVal phWnd As LongPtr, _
    ByVal phDC As LongPtr, _
    ByVal piZoneID As Integer, _
    ByVal piLine As Long, _
    ByVal piLeft As Integer, _
    ByVal piTop As Integer, _
    ByVal piRight As Integer, _
    ByVal piBottom As Integer _
  ) As Integer
#Else
Public Function OnConsoulZonePaint( _
    ByVal phWnd As Long, _
    ByVal phDC As Long, _
    ByVal piZoneID As Integer, _
    ByVal piLine As Long, _
    ByVal piLeft As Integer, _
    ByVal piTop As Integer, _
    ByVal piRight As Integer, _
    ByVal piBottom As Integer _
  ) As Integer
#End If
  On Error Resume Next
  ConsoulEventDispatcher.BroadcastZonePaint phWnd, phDC, piZoneID, piLine, piLeft, piTop, piRight, piBottom
End Function
