Attribute VB_Name = "MWin32API"
Option Compare Database
Option Explicit

Public Type RECT
  left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type POINTAPI
  X As Long
  Y As Long
End Type

'Mouse button (and shift/ctrl) states for mouse events
Public Const MK_CONTROL   As Integer = &H8    'The CTRL key is down
Public Const MK_LBUTTON   As Integer = &H1    'The left mouse button is down
Public Const MK_MBUTTON   As Integer = &H10   'The middle mouse button is down
Public Const MK_RBUTTON   As Integer = &H2    'The right mouse button is down
Public Const MK_SHIFT     As Integer = &H4    'The SHIFT key is down
Public Const MK_XBUTTON1  As Integer = &H20   'The first X button is down
Public Const MK_XBUTTON2  As Integer = &H40   'The second X button is down

' ShowWindow() Commands
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10

' Scroll Bar Constants
Public Const SB_HORZ = 0
Public Const SB_VERT = 1
Public Const SB_CTL = 2
Public Const SB_BOTH = 3

Public Const WM_VSCROLL = &H115
Public Const SB_PAGEUP = 2
Public Const SB_PAGEDOWN = 3
Public Const SB_TOP = 6
Public Const SB_BOTTOM = 7
Public Const SB_LINEUP = 0
Public Const SB_LINEDOWN = 1

' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

'Close window
Public Const WM_CLOSE = &H10

' SetWindowPos() hwndInsertAfter values
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const GWL_STYLE = (-16)
Public Const WS_BORDER = &H800000

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const TwipsPerInch = 1440

#If Win64 Then
  Public Declare PtrSafe Function apiGetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long
  Public Declare PtrSafe Function apiWindowFromPoint Lib "user32" Alias "WindowFromPoint" (ByVal xpoint As Long, ByVal ypoint As Long) As LongPtr
  Public Declare PtrSafe Function apiGetFocus Lib "user32" Alias "GetFocus" () As LongPtr
  Public Declare PtrSafe Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
  Public Declare PtrSafe Function apiGetClientRect Lib "user32" Alias "GetClientRect" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
  Public Declare PtrSafe Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
  Public Declare PtrSafe Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
  Public Declare PtrSafe Function apiShowScrollBar Lib "user32" Alias "ShowScrollBar" (ByVal hWnd As LongPtr, ByVal wBar As Long, ByVal bShow As Long) As Long
  Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
  Public Declare PtrSafe Function PostMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
  Public Declare PtrSafe Function apiBringWindowToTop Lib "user32" Alias "BringWindowToTop" (ByVal hWnd As LongPtr) As Long
  Public Declare PtrSafe Function apiSetRect Lib "user32" Alias "SetRect" (lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
  Public Declare PtrSafe Function apiCopyRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, lpSourceRect As RECT) As Long
  Public Declare PtrSafe Function apiGetWindowLong Lib "user32" Alias "apiGetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
  Public Declare PtrSafe Function apiSetWindowLong Lib "user32" Alias "apiSetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
  Public Declare PtrSafe Function apiGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As LongPtr) As LongPtr
  Public Declare PtrSafe Function apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
  Public Declare PtrSafe Function apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
#Else
  Public Declare Function apiGetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long
  Public Declare Function apiWindowFromPoint Lib "user32" Alias "WindowFromPoint" (ByVal xpoint As Long, ByVal ypoint As Long) As Long
  Public Declare Function apiGetFocus Lib "user32" Alias "GetFocus" () As LongPtr
  Public Declare Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
  Public Declare Function apiGetClientRect Lib "user32" Alias "GetClientRect" (ByVal hWnd As Long, lpRect As RECT) As Long
  Public Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
  Public Declare Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
  Public Declare Function apiShowScrollBar Lib "user32" Alias "ShowScrollBar" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
  Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  Public Declare Function PostMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  Public Declare Function apiBringWindowToTop Lib "user32" Alias "BringWindowToTop" (ByVal hWnd As Long) As Long
  Public Declare Function apiSetRect Lib "user32" Alias "SetRect" (lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
  Public Declare Function apiCopyRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, lpSourceRect As RECT) As Long
  Public Declare Function apiGetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
  Public Declare Function apiSetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Public Declare Function apiGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As Long) As Long
  Public Declare Function apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Long, ByVal nIndex As Long) As Long
  Public Declare Function apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As Long, ByVal hdc As Long) As Long
#End If

#If Win64 Then
Public Function GetClientRect(ByVal plHWnd As LongPtr, ByRef pRetRECT As RECT) As Long
#Else
Public Function GetClientRect(ByVal plHWnd As Long, ByRef pRetRECT As RECT) As Long
#End If
  GetClientRect = apiGetClientRect(plHWnd, pRetRECT)
End Function

#If Win64 Then
Public Function MaximizeWindow(ByVal hWnd As LongPtr) As Long
#Else
Public Function MaximizeWindow(ByVal hWnd As Long) As Long
#End If
  On Error Resume Next
  MaximizeWindow = apiShowWindow(hWnd, SW_MAXIMIZE)
End Function

#If Win64 Then
Public Function RestoreWindow(ByVal hWnd As LongPtr) As Long
#Else
Public Function RestoreWindow(ByVal hWnd As Long) As Long
#End If
  On Error Resume Next
  RestoreWindow = apiShowWindow(hWnd, SW_RESTORE)
End Function

'Modififed from From https://stackoverflow.com/questions/23042374/access-2010-vba-api-twips-pixel
Public Function TwipsToPixelsY(ByVal Y As Long) As Integer
  #If Win64 Then
  Dim ScreenDC As LongPtr
  #Else
  Dim ScreenDC As Long
  #End If
  ScreenDC = apiGetDC(0)
  TwipsToPixelsY = Y / TwipsPerInch * apiGetDeviceCaps(ScreenDC, LOGPIXELSY)
  apiReleaseDC 0, ScreenDC
End Function

Public Function TwipsToPixelsX(ByVal X As Long) As Integer
  #If Win64 Then
  Dim ScreenDC As LongPtr
  #Else
  Dim ScreenDC As Long
  #End If
  ScreenDC = apiGetDC(0)
  TwipsToPixelsX = X / TwipsPerInch * apiGetDeviceCaps(ScreenDC, LOGPIXELSX)
  apiReleaseDC 0, ScreenDC
End Function

Public Function PixelsToTwipsX(ByVal X As Integer) As Long
  #If Win64 Then
  Dim ScreenDC As LongPtr
  #Else
  Dim ScreenDC As Long
  #End If
  ScreenDC = apiGetDC(0)
  PixelsToTwipsX = CLng(CDbl(X) * TwipsPerInch / apiGetDeviceCaps(ScreenDC, LOGPIXELSX))
  apiReleaseDC 0, ScreenDC
End Function

Public Function PixelsToTwipsY(ByVal Y As Integer) As Long
  #If Win64 Then
  Dim ScreenDC As LongPtr
  #Else
  Dim ScreenDC As Long
  #End If
  ScreenDC = apiGetDC(0)
  PixelsToTwipsY = CLng(CDbl(Y) * TwipsPerInch / apiGetDeviceCaps(ScreenDC, LOGPIXELSY))
  apiReleaseDC 0, ScreenDC
End Function


