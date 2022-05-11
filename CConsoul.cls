VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConsoul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CConsoul.bas
'  VB/A wrapper for the Consoul library
'
'(C) 2018, devinfo.net, Développement Informatique Services, Francesco Foti
'Library docs: https://consoul.net/docs/reference/index.htm
'VBA SDK docs: https://consoul.net/docs/sdk/vba/index.htm

#If MSACCESS Then
Option Compare Database
#End If
Option Explicit

Public Enum eConsoulCreateAttributes
  LW_CREATE_FILLEDVEMPTY = 1
  LW_AUTOADJUST_ON_MAXCHARWIDTH = 2
  LW_NO_AUTOREDRAW = 4
  LW_RENDERMODEBYLINE = 8
  LW_SEND_MOUSEMOVE = 16  'Use SendMessage(), not for VB/A w/o subclassing
  LW_TRACK_ZONES = 32
  LW_SENDMESSAGE_NOCALLBACKS = 64
  LW_MULTIZONECLICK = 128
  LW_BKCOLSPILL = 256
End Enum

Public Enum eWmMouseButton
  WM_LBUTTONDOWN = &H201
  WM_LBUTTONUP = &H202
  WM_LBUTTONDBLCLK = &H203
  WM_RBUTTONDOWN = &H204
  WM_RBUTTONUP = &H205
  WM_RBUTTONDBLCLK = &H206
  WM_MBUTTONDOWN = &H207
  WM_MBUTTONUP = &H208
  WM_MBUTTONDBLCLK = &H209
  WM_XBUTTONDOWN = &H20B
  WM_XBUTTONUP = &H20C
  WM_XBUTTONDBLCLK = &H20D
End Enum

Public Enum eRenderMode
  rmContinuous = 0
  rmByLine = 1
End Enum

Public Enum ePaintCallbackMode  'bit 0 and 1 used
  WMPAINTCBK_NONE = 0
  WMPAINTCBK_BEFORE = 1
  WMPAINTCBK_AFTER = 2
End Enum

Public Enum eLineVerticalLocation
  elsTop
  elsBottom
End Enum

#If Win64 Then
  'Consoul API 64bits declarations
  Private Declare PtrSafe Function CSClear Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSCreateLogWindow Lib "consoul_010205_64.dll" _
    (ByVal hwndParent As LongPtr, _
     ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, _
     ByVal lBackColor As Long, ByVal lForeColor As Long, _
     ByVal sFontName As LongPtr, ByVal iFontSize As Integer, _
     ByVal iQueueSize As Integer, ByVal pwCreateAttribs As Integer) As LongPtr
  Private Declare PtrSafe Function CSDestroyLogWindow Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetAutoRedraw Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetCaretBlinkMs Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetCaretHeight Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetCaretPos Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByRef wRow As Integer, _
      ByRef wCol As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSGetCaretWidth Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetCharHeight Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetCharWidth Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetFontUnicodeRange Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal lIndex As Long, _
      ByVal lRetLow As LongPtr, _
      ByVal lRetCount As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetHoverCursor Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As LongPtr
  Private Declare PtrSafe Function CSGetLineHeight Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetLinePadding Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wTopOrBottom As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSGetLineSpacing Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wTopOrBottom As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSGetLineText Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wLine As Integer, _
      ByVal wBufferCharLen As Integer, _
      ByVal lpStringBuffer As LongPtr, _
      ByVal wAsVt100 As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSGetMaxCharWidth Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetMaxVisibleCols Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetMaxVisibleRows Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetMouseCursor Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As LongPtr
  Private Declare PtrSafe Function CSGetMultiZoneClick Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetQueueCapacity Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetRenderMode Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetTopLine Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetTrackZones Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetUseCallbacks Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetVersion Lib "consoul_010205_64.dll" ( _
      ByVal lpStringBuffer As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetZoneTag Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wLine As Integer, _
      ByVal wZoneID As Integer, _
      ByVal wBufferCharLen As Integer, _
      ByVal lpStringBuffer As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetZoneText Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wLine As Integer, _
      ByVal wZoneID As Integer, _
      ByVal wBufferCharLen As Integer, _
      ByVal lpStringBuffer As LongPtr, _
      ByVal wAsVt100 As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSIsCaretVisible Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSLineCount Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSLoadFontUnicodeRanges Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Long
  Private Declare Function CSPaintOnDC Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal hdc As LongPtr, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long, _
      ByVal iStartLine As Integer, _
      ByVal iEndLine As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSPopLines Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wPopCount As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSPushLine Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal lpszString As LongPtr, _
      ByVal wNoParsing As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSRedrawLine Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal iLine As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSRefreshWindow Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSReplaceZone Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wLine As Integer, _
      ByVal wZoneID As Integer, _
      ByVal lpszReplaceBy As LongPtr, _
      ByVal wNoParsing As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSResetAllLines Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSSetAlphaTransparency Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal bPercent As Byte _
    ) As Integer
  Private Declare PtrSafe Function CSSetAutoRedraw Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wAutoRedraw As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSSetCaretBlinkMs Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wMillisecs As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSSetCaretHeight Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wHeight As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSSetCaretPos Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wRow As Integer, _
      ByVal wCol As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSSetCaretWidth Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wWidth As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSSetColorTransparency Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal bPercent As Byte, _
      ByVal lColorRef As Long, _
      ByVal bEnabled As Byte _
    ) As Integer
  Private Declare PtrSafe Function CSSetHoverCursor Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal hCursor As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSSetLine Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal iLine As Integer, _
      ByVal lpszString As LongPtr, _
      ByVal wNoParsing As Integer, _
      ByVal wNoUpdate As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSSetLinePadding Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wSpacing As Integer, _
      ByVal wTopOrBottom As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSSetLineSpacing Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wSpacing As Integer, _
      ByVal wTopOrBottom As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSSetMouseCursor Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal hCursor As LongPtr _
    ) As Integer
  Private Declare PtrSafe Function CSGetMultiZoneClick Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wMulti As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSSetRenderMode Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wRenderMode As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSSetTrackZones Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wTrack As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSSetUseCallbacks Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wUseCallbacks As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSShowCaret Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal wShow As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSTextWidth Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal lpszText As LongPtr, _
      ByVal lDTFlags As Long _
    ) As Integer
  Private Declare PtrSafe Function CSUnloadFontUnicodeRanges Lib "consoul_010205_64.dll" ( _
      ByVal hWnd As LongPtr _
    ) As Long
  '(Undocumented functions)
  Private Declare PtrSafe Function CSClickAtCaretPos Lib "consoul_010205_64.dll" (ByVal hWnd As LongPtr, ByVal plClickMsg As Long) As Integer
  Private Declare PtrSafe Function CSGetLongestLineWidth Lib "consoul_010205_64.dll" (ByVal hWnd As LongPtr) As Long
  Private Declare PtrSafe Function CSSetAutoAdjustWidth Lib "consoul_010205_64.dll" (ByVal hWnd As LongPtr, ByVal pwAutoAdjustWidth As Integer) As Integer
  Private Declare PtrSafe Function CSGetAutoAdjustWidth Lib "consoul_010205_64.dll" (ByVal hWnd As LongPtr) As Integer
  'Set callbacks
  Private Declare PtrSafe Function CSSetOnMouseButtonCallback Lib "consoul_010205_64.dll" (ByVal hWnd As LongPtr, ByVal pEvtProc As LongPtr) As Integer
  Private Declare PtrSafe Function CSSetOnVirtuaLineCallback Lib "consoul_010205_64.dll" (ByVal hWnd As LongPtr, ByVal pEvtProc As LongPtr) As Integer
  Private Declare PtrSafe Function CSSetOnWmPaintCallback Lib "consoul_010205_64.dll" (ByVal hWnd As LongPtr, ByVal pwCbkMode As Integer, ByVal pEvtProc As LongPtr) As Integer
  Private Declare PtrSafe Function CSSetOnDrawZoneCallback Lib "consoul_010205_64.dll" (ByVal hWnd As LongPtr, ByVal pEvtProc As LongPtr) As Integer
  'Win32 API rarely used outside this class
  Private Declare PtrSafe Function SetScrollPos Lib "user32" (ByVal hWnd As LongPtr, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
#Else
  'Consoul API 32bits declarations
  Private Declare Function CSClear Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSCreateLogWindow Lib "consoul_010205_32.dll" ( _
      ByVal hwndParent As Long, _
      ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, _
      ByVal lBackColor As Long, ByVal lForeColor As Long, _
      ByVal sFontName As Long, ByVal iFontSize As Integer, _
      ByVal iQueueSize As Integer, ByVal wCreateAttribs As Integer _
    ) As Long
  Private Declare Function CSDestroyLogWindow Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetAutoRedraw Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetCaretBlinkMs Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetCaretHeight Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetCaretPos Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByRef wRow As Integer, _
      ByRef wCol As Integer _
    ) As Integer
  Private Declare Function CSGetCaretWidth Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetCharHeight Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetCharWidth Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetFontUnicodeRange Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal lIndex As Long, _
      ByVal lRetLow As Long, _
      ByVal lRetCount As Long _
    ) As Integer
  Private Declare Function CSGetHoverCursor Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Long
  Private Declare Function CSGetLineHeight Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetLinePadding Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wTopOrBottom As Integer _
    ) As Integer
  Private Declare Function CSGetLineSpacing Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wTopOrBottom As Integer _
    ) As Integer
  Private Declare Function CSGetLineText Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wLine As Integer, _
      ByVal wBufferCharLen As Integer, _
      ByVal lpszStringBuffer As Long, _
      ByVal wAsVt100 As Integer _
    ) As Integer
  Private Declare Function CSGetMaxCharWidth Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetMaxVisibleCols Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetMaxVisibleRows Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetMouseCursor Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Long
  Private Declare Function CSGetMultiZoneClick Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetQueueCapacity Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetRenderMode Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetTopLine Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetTrackZones Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetUseCallbacks Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSGetVersion Lib "consoul_010205_32.dll" ( _
      ByVal lpStringBuffer As Long _
    ) As Integer
  Private Declare Function CSGetZoneTag Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wLine As Integer, _
      ByVal wZoneID As Integer, _
      ByVal wBufferCharLen As Integer, _
      ByVal lpStringBuffer As Long _
    ) As Integer
  Private Declare Function CSGetZoneText Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wLine As Integer, _
      ByVal wZoneID As Integer, _
      ByVal wBufferCharLen As Integer, _
      ByVal lpStringBuffer As Long, _
      ByVal wAsVt100 As Integer _
    ) As Integer
  Private Declare Function CSIsCaretVisible Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSLineCount Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSLoadFontUnicodeRanges Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Long
  Private Declare Function CSPaintOnDC Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal hdc As Long, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long, _
      ByVal iStartLine As Integer, _
      ByVal iEndLine As Integer _
    ) As Integer
  Private Declare Function CSPopLines Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wPopCount As Integer _
    ) As Integer
  Private Declare Function CSPushLine Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal lpszString As Long, _
      ByVal wNoParsing As Integer _
    ) As Integer
  Private Declare Function CSRedrawLine Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal iLine As Integer _
    ) As Integer
  Private Declare Function CSRefreshWindow Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSReplaceZone Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wLine As Integer, _
      ByVal wZoneID As Integer, _
      ByVal lpszReplaceBy As Long, _
      ByVal wNoParsing As Integer _
    ) As Integer
  Private Declare Function CSResetAllLines Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  Private Declare Function CSSetAlphaTransparency Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal bPercent As Byte _
    ) As Integer
  Private Declare Function CSSetAutoRedraw Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wAutoRedraw As Integer _
    ) As Integer
  Private Declare Function CSSetCaretBlinkMs Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wMillisecs As Integer _
    ) As Integer
  Private Declare Function CSSetCaretHeight Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wHeight As Integer _
    ) As Integer
  Private Declare Function CSSetCaretPos Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wRow As Integer, _
      ByVal wCol As Integer _
    ) As Integer
  Private Declare Function CSSetCaretWidth Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wWidth As Integer _
    ) As Integer
  Private Declare PtrSafe Function CSSetColorTransparency Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal bPercent As Byte, _
      ByVal lColorRef As Long, _
      ByVal bEnabled As Byte _
    ) As Integer
  Private Declare Function CSSetHoverCursor Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal hCursor As Long _
    ) As Integer
  Private Declare Function CSSetLine Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal iLine As Integer, _
      ByVal lpszString As Long, _
      ByVal wNoParsing As Integer, _
      ByVal wNoUpdate As Integer _
    ) As Integer
  Private Declare Function CSSetLinePadding Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wSpacing As Integer, _
      ByVal wTopOrBottom As Integer _
    ) As Integer
  Private Declare Function CSSetLineSpacing Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wSpacing As Integer, _
      ByVal wTopOrBottom As Integer _
    ) As Integer
  Private Declare Function CSSetMouseCursor Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal hCursor As Long _
    ) As Integer
  Private Declare Function CSSetMultiZoneClick Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wMulti As Integer _
    ) As Integer
  Private Declare Function CSSetRenderMode Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wRenderMode As Integer _
    ) As Integer
  Private Declare Function CSSetTrackZones Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wTrack As Integer _
    ) As Integer
  Private Declare Function CSSetUseCallbacks Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wUseCallbacks As Integer _
    ) As Integer
  Private Declare Function CSShowCaret Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal wShow As Integer _
    ) As Integer
  Private Declare Function CSTextWidth Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long, _
      ByVal lpszText As Long, _
      ByVal lDTFlags As Long _
    ) As Integer
  Private Declare Function CSUnloadFontUnicodeRanges Lib "consoul_010205_32.dll" ( _
      ByVal hWnd As Long _
    ) As Integer
  '(Undocumented functions)
  Private Declare Function CSClickAtCaretPos Lib "consoul_010205_32.dll" (ByVal hWnd As Long, ByVal plClickMsg As Long) As Integer
  Private Declare Function CSGetLongestLineWidth Lib "consoul_010205_32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function CSSetAutoAdjustWidth Lib "consoul_010205_32.dll" (ByVal hWnd As Long, ByVal pwAutoAdjustWidth As Integer) As Integer
  Private Declare Function CSGetAutoAdjustWidth Lib "consoul_010205_32.dll" (ByVal hWnd As Long) As Integer
  'Set callbacks
  Private Declare Function CSSetOnMouseButtonCallback Lib "consoul_010205_32.dll" (ByVal hWnd As Long, ByVal pEvtProc As Long) As Integer
  Private Declare Function CSSetOnVirtuaLineCallback Lib "consoul_010205_32.dll" (ByVal hWnd As Long, ByVal pEvtProc As Long) As Integer
  Private Declare Function CSSetOnWmPaintCallback Lib "consoul_010205_32.dll" (ByVal hWnd As Long, ByVal pwCbkMode As Integer, ByVal pEvtProc As Long) As Integer
  Private Declare Function CSSetOnDrawZoneCallback Lib "consoul_010205_32.dll" (ByVal hWnd As Long, ByVal pEvtProc As Long) As Integer
  'Win32 API rarely used outside this class
  Private Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
#End If

'Private class members
#If Win64 Then
  Private mhWnd             As LongPtr
  Private mhWndParent       As LongPtr
#Else
  Private mhWnd             As Long
  Private mhWndParent       As Long
#End If
Private msBuffer          As String

'Public properties (free r/w access, w/o property proc)
Public FontName             As String
Public FontSize             As Integer
Public WindowTitle          As String
Public BackColor            As Long
Public ForeColor            As Long
Public MaxCapacity          As Integer
Public WithFrame            As Boolean

'Track window position and dimensions
Private mrcWindow As RECT

Private Sub Class_Initialize()
  ' Set some defaults
  Me.FontName = "Courier New"
  Me.FontSize = 14
  Me.WindowTitle = "Consoul Window"
  Me.BackColor = vbBlack
  Me.ForeColor = RGB(200, 200, 200)
  Me.WithFrame = False
End Sub

Private Sub Class_Terminate()
  If mhWnd <> 0 Then
    Call CSDestroyLogWindow(mhWnd)
    mhWnd = 0
  End If
End Sub

' -------- Class properties & events

#If Win64 Then
Public Property Get hwndParent() As LongPtr
#Else
Public Property Get hwndParent() As Long
#End If
  hwndParent = mhWndParent
End Property

#If Win64 Then
Public Property Get hWnd() As LongPtr
#Else
Public Property Get hWnd() As Long
#End If
  hWnd = mhWnd
End Property

'The number of text lines in the windows queue
Public Property Get LineCount() As Integer
  LineCount = CSLineCount(mhWnd)
End Property

Public Property Get TopLine() As Integer
  If mhWnd <> 0 Then TopLine = CSGetTopLine(mhWnd)
End Property

Public Property Get CharWidth() As Integer
  If mhWnd <> 0 Then CharWidth = CSGetCharWidth(mhWnd)
End Property

Public Property Get MaxCharWidth() As Integer
  If mhWnd <> 0 Then MaxCharWidth = CSGetMaxCharWidth(mhWnd)
End Property

Public Property Get CharHeight() As Integer
  If mhWnd <> 0 Then CharHeight = CSGetCharHeight(mhWnd)
End Property

Public Property Get MaxVisibleCols() As Integer
  If mhWnd <> 0 Then MaxVisibleCols = CSGetMaxVisibleCols(mhWnd)
End Property

Public Property Get MaxVisibleRows() As Integer
  If mhWnd <> 0 Then MaxVisibleRows = CSGetMaxVisibleRows(mhWnd)
End Property

Public Property Get ConsoulVersion() As String
  Dim sBuffer As String
  Dim i       As Integer
  sBuffer = Space$(25)  'at least 25 chars as specs tell
  Call CSGetVersion(StrPtr(sBuffer))
  sBuffer = StrConv(sBuffer, vbUnicode)
  i = InStr(1, sBuffer, Chr$(0))
  If i > 0 Then
    sBuffer = left$(sBuffer, i - 1)
  End If
  ConsoulVersion = sBuffer
End Property

Public Property Get MultiZoneClick() As Boolean
  If mhWnd Then
    MultiZoneClick = CBool(CSGetMultiZoneClick(mhWnd))
  End If
End Property

Public Property Let TrackZones(ByVal pfAllowed As Boolean)
  If mhWnd Then
    CSSetTrackZones mhWnd, Abs(CInt(pfAllowed))
  End If
End Property

Public Property Get TrackZones() As Boolean
  If mhWnd Then
    TrackZones = CBool(CSGetTrackZones(mhWnd))
  End If
End Property

Public Property Let MultiZoneClick(ByVal pfAllowed As Boolean)
  If mhWnd Then
    CSSetMultiZoneClick mhWnd, Abs(CInt(pfAllowed))
  End If
End Property

Public Property Get CaretWidth() As Integer
  If mhWnd Then
    CaretWidth = CSGetCaretWidth(mhWnd)
  End If
End Property

Public Property Let CaretWidth(ByVal piNewWidth As Integer)
  If mhWnd Then
    Call CSSetCaretWidth(mhWnd, piNewWidth)
  End If
End Property

Public Property Get CaretHeight() As Integer
  If mhWnd Then
    CaretHeight = CSGetCaretHeight(mhWnd)
  End If
End Property

Public Property Let CaretHeight(ByVal piNewHeight As Integer)
  If mhWnd Then
    Call CSSetCaretHeight(mhWnd, piNewHeight)
  End If
End Property

Public Property Get AutoAdjustWidth() As Boolean
  If mhWnd Then
    AutoAdjustWidth = CSGetAutoAdjustWidth(mhWnd)
  End If
End Property

Public Property Let AutoAdjustWidth(ByVal pfAutoAdjustWidth As Boolean)
  If mhWnd Then
    Call CSSetAutoAdjustWidth(mhWnd, Abs(CInt(pfAutoAdjustWidth)))
  End If
End Property

Public Property Get left() As Integer
  left = mrcWindow.left
End Property

Public Property Get Top() As Integer
  Top = mrcWindow.Top
End Property

Public Property Get Width() As Integer
  Width = (mrcWindow.Bottom - mrcWindow.Top)
End Property

Public Property Get Height() As Integer
  Height = (mrcWindow.Right - mrcWindow.left)
End Property

Public Sub GetSizeAndPosRect(ByRef prcRetSize As RECT)
  apiCopyRect prcRetSize, mrcWindow
End Sub

Public Sub SetSizeAndPosRect(ByRef prcSizeAndPos As RECT, Optional ByVal bRepaint As Long = 1)
  With prcSizeAndPos
    Me.MoveWindow .left, .Top, .Right - .left, .Bottom - .Top, bRepaint
  End With
End Sub

#If Win64 Then
Public Function Attach( _
    ByVal phWndParent As LongPtr, _
    ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, _
    Optional ByVal OnMouseBtnEvtProc As LongPtr = 0, _
    Optional ByVal OnZoneDrawEvtProc As LongPtr = 0, _
    Optional ByVal OnVirtuaLineEvtProc As LongPtr = 0, _
    Optional ByVal piCreateAttributes As eConsoulCreateAttributes = LW_RENDERMODEBYLINE) As Boolean
#Else
Public Function Attach( _
    ByVal phWndParent As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, _
    Optional ByVal OnMouseBtnEvtProc As LongPtr = 0, _
    Optional ByVal OnZoneDrawEvtProc As LongPtr = 0, _
    Optional ByVal OnVirtuaLineEvtProc As LongPtr = 0, _
    Optional ByVal piCreateAttributes As eConsoulCreateAttributes = LW_RENDERMODEBYLINE) As Boolean
#End If
  
  mhWnd = CSCreateLogWindow( _
    phWndParent, _
    X, Y, Width, Height, _
    Me.BackColor, Me.ForeColor, _
    StrPtr(Me.FontName), Me.FontSize, _
    Me.MaxCapacity, _
    CInt(piCreateAttributes))
  If mhWnd <> 0 Then
    'if no max capacity was provided, consoul decides for us
    If Me.MaxCapacity = 0 Then Me.MaxCapacity = CSGetQueueCapacity(mhWnd)
    mhWndParent = phWndParent 'Used when creating log windows to set parent
    If OnMouseBtnEvtProc <> 0 Then
      Call CSSetOnMouseButtonCallback(mhWnd, OnMouseBtnEvtProc)
    End If
    If OnZoneDrawEvtProc <> 0 Then
      Call CSSetOnDrawZoneCallback(mhWnd, OnZoneDrawEvtProc)
    End If
    If OnVirtuaLineEvtProc <> 0 Then
      Call CSSetOnVirtuaLineCallback(mhWnd, OnVirtuaLineEvtProc)
    End If
  End If

  Attach = True
End Function

#If Win64 Then
Public Sub SetMouseButtonCallback(ByVal lpProcAddress As LongPtr)
#Else
Public Sub SetMouseButtonCallback(ByVal lpProcAddress As Long)
#End If
  If mhWnd <> 0 Then
    Call CSSetOnMouseButtonCallback(mhWnd, lpProcAddress)
  End If
End Sub

#If Win64 Then
Public Sub SetDrawZoneCallback(ByVal lpProcAddress As LongPtr)
#Else
Public Sub SetDrawZoneCallback(ByVal lpProcAddress As Long)
#End If
  If mhWnd <> 0 Then
    Call CSSetOnDrawZoneCallback(mhWnd, lpProcAddress)
  End If
End Sub

#If Win64 Then
Public Sub SetVirtuaLineCallback(ByVal lpProcAddress As LongPtr)
#Else
Public Sub SetVirtuaLineCallback(ByVal lpProcAddress As Long)
#End If
  If mhWnd <> 0 Then
    Call CSSetOnVirtuaLineCallback(mhWnd, lpProcAddress)
  End If
End Sub

#If Win64 Then
Public Sub SetWmPaintCallback(ByVal pwCbkModeBits As ePaintCallbackMode, ByVal lpProcAddress As LongPtr)
#Else
Public Sub SetWmPaintCallback(ByVal pwCbkModeBits As ePaintCallbackMode, ByVal lpProcAddress As Long)
#End If
  If mhWnd <> 0 Then
    Call CSSetOnWmPaintCallback(mhWnd, pwCbkModeBits, lpProcAddress)
  End If
End Sub

Public Sub Detach()
  If mhWnd <> 0 Then CSDestroyLogWindow mhWnd
  mhWnd = 0
End Sub

Public Property Get AutoRedraw() As Boolean
  If mhWnd <> 0 Then
    AutoRedraw = CBool(CSGetAutoRedraw(mhWnd))
    Exit Property
  End If
  AutoRedraw = True
End Property

Public Property Let AutoRedraw(ByVal pfFlag As Boolean)
  If mhWnd <> 0 Then
    Call CSSetAutoRedraw(mhWnd, Abs(CInt(pfFlag)))
  End If
End Property

Public Property Get RenderMode() As eRenderMode
  If mhWnd <> 0 Then
    RenderMode = CSGetRenderMode(mhWnd)
  End If
End Property

Public Property Let RenderMode(ByVal peMode As eRenderMode)
  If mhWnd <> 0 Then
    Call CSSetRenderMode(mhWnd, CInt(peMode))
  End If
End Property

Public Property Get TextWidth(ByVal psText As String, Optional ByVal plDTFlags As Long = 0&) As Integer
  If mhWnd <> 0 Then
    TextWidth = CSTextWidth(mhWnd, StrPtr(psText), plDTFlags)
  End If
End Property

' -------- Win32 API wrappers / facilitators

Public Sub GetClientRect(ByRef prcClient As RECT)
  If mhWnd <> 0 Then Call apiGetClientRect(mhWnd, prcClient)
End Sub

Public Sub GetParentClientRect(ByRef prcClient As RECT)
  If mhWnd <> 0 Then
    If mhWndParent > 0 Then
      Call apiGetClientRect(mhWndParent, prcClient)
    End If
  End If
End Sub

Public Function GetClientWidth() As Integer
  If mhWnd = 0 Then Exit Function
  Dim rcClient As RECT
  Call apiGetClientRect(mhWnd, rcClient)
  GetClientWidth = rcClient.Right - rcClient.left
End Function

Public Function GetClientHeight() As Integer
  If mhWnd = 0 Then Exit Function
  Dim rcClient As RECT
  Call apiGetClientRect(mhWnd, rcClient)
  GetClientHeight = rcClient.Bottom - rcClient.Top
End Function

Public Sub ShowWindow(ByVal pfVisible As Boolean)
  If mhWnd <> 0 Then apiShowWindow mhWnd, IIf(pfVisible, SW_SHOW, SW_HIDE)
End Sub

Public Sub MoveWindow(ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal bRepaint As Long = 1)
  If mhWnd <> 0 Then
    apiMoveWindow mhWnd, X, Y, nWidth, nHeight, bRepaint
    With mrcWindow
      .left = X
      .Top = Y
      .Right = X + nWidth
      .Bottom = Y + nHeight
    End With
  End If
End Sub

Public Sub BringWindowToTop()
  If mhWnd <> 0 Then apiBringWindowToTop (mhWnd)
End Sub

#If Win64 Then
Public Sub SetWindowPos(ByVal phWndInsertAfter As LongPtr, ByVal px As Long, ByVal py As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal pwFlags As Long = 0&)
#Else
Public Sub SetWindowPos(ByVal phWndInsertAfter As Long, ByVal px As Long, ByVal py As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal pwFlags As Long = 0&)
#End If
  If mhWnd <> 0 Then
    Dim X       As Integer
    Dim Y       As Integer
    Dim iWidth  As Integer
    Dim iHeight As Integer
    X = mrcWindow.left
    Y = mrcWindow.Top
    iWidth = mrcWindow.Right - mrcWindow.left
    iHeight = mrcWindow.Bottom - mrcWindow.Top
    apiSetWindowPos mhWnd, phWndInsertAfter, px, py, nWidth, nHeight, pwFlags
    If Not (pwFlags And SWP_NOREPOSITION) Then  'adjust member variables tracking position
      X = px
      Y = py
    End If
    If Not (pwFlags And SWP_NOSIZE) Then  'adjust member variables tracking size
      iWidth = nWidth
      iHeight = nHeight
    End If
    apiSetRect mrcWindow, X, Y, X + iWidth, X + iHeight
  End If
End Sub


' -------- Console output methods

'Append a chunk of text to the internal console line buffer.
'The internal buffer is pushed in the console's window queue
'at the next OutputLn call.
Public Sub Output(ByVal psLine As String)
  If mhWnd <> 0 Then msBuffer = msBuffer & psLine
End Sub

'Push the msBuffer (text line buffer) in the window's queue, and,
'Consoul will immediately render the console window
'if AutoRedraw is True, scroll it if necessary.
Public Function OutputLn(ByVal psLine As String, Optional ByVal piQBColorText As Variant, Optional ByVal pfNoParsing As Variant) As Integer
  Dim iLine   As Integer
  If mhWnd <> 0 Then
    If IsMissing(pfNoParsing) Then pfNoParsing = False
    If Not IsMissing(piQBColorText) Then
      msBuffer = VT_FCOLOR(QBColor(piQBColorText)) & msBuffer
    End If
    msBuffer = msBuffer & psLine & VT_RESET()
    iLine = CSPushLine(mhWnd, StrPtr(msBuffer), Abs(CInt(pfNoParsing)))
    msBuffer = ""
  End If
  OutputLn = iLine
End Function

Public Function PushVirtualLine() As Integer
  PushVirtualLine = CSPushLine(mhWnd, 0&, 1)
End Function

'Change the text of an existing console line.
'piLine is the index of the line in the window queue [1..LineCount].
'Call RedrawLine after that to update the console window display accordingly.
Public Sub SetLine(ByVal piLine As Integer, ByVal psText As String, Optional ByVal pfNoParsing As Boolean = False, Optional ByVal pfRedrawLine As Boolean = False)
  If mhWnd <> 0 Then
    Call CSSetLine(mhWnd, piLine, StrPtr(psText), Abs(CInt(pfNoParsing)), Abs(CInt(Not pfRedrawLine)))
  End If
End Sub

Public Sub ResetVirtualLine(ByVal piLine As Integer, Optional ByVal pfRedrawLine As Boolean = False)
  If mhWnd <> 0 Then
    Call CSSetLine(mhWnd, piLine, 0&, 0, Abs(CInt(Not pfRedrawLine)))
  End If
End Sub

Public Sub MakeLineVirtual(ByVal piLine As Integer, Optional ByVal pfRedrawLine As Boolean = False)
  If mhWnd <> 0 Then
    Call CSSetLine(mhWnd, piLine, 0, 0, Abs(CInt(Not pfRedrawLine)))
  End If
End Sub

Public Sub RedrawLine(ByVal piLine As Integer)
  Call CSRedrawLine(mhWnd, piLine)
End Sub

Public Function ReplaceZone(ByVal piLine As Integer, ByVal piZoneID As Integer, ByVal psReplaceBy As String, Optional ByVal pfNoParsing As Boolean = False) As Integer
  If mhWnd <> 0 Then
    ReplaceZone = CSReplaceZone(mhWnd, piLine, piZoneID, StrPtr(psReplaceBy), Abs(CInt(pfNoParsing)))
  End If
End Function

Public Sub Clear()
  If mhWnd <> 0 Then
    Call CSClear(mhWnd)
  End If
End Sub

Public Sub RefreshWindow()
  If mhWnd <> 0 Then
    Call CSRefreshWindow(mhWnd)
  End If
End Sub

Public Function GetCaretPos(ByRef piRetRow As Integer, ByRef piRetCol As Integer) As Integer
  If mhWnd <> 0 Then
    GetCaretPos = CSGetCaretPos(mhWnd, piRetRow, piRetCol)
  End If
End Function

Public Function SetCaretPos(ByVal piRow As Integer, ByVal piCol As Integer) As Integer
  If mhWnd <> 0 Then
    SetCaretPos = CSSetCaretPos(mhWnd, piRow, piCol)
  End If
End Function

Public Function IsCaretVisible() As Boolean
  If mhWnd <> 0 Then
    IsCaretVisible = CSIsCaretVisible(mhWnd)
  End If
End Function

Public Function ShowCaret(ByVal pfVisible As Boolean) As Integer
  If mhWnd <> 0 Then
    ShowCaret = CSShowCaret(mhWnd, CInt(Abs(pfVisible)))
  End If
End Function

Public Property Get CaretBlinkMs() As Integer
  If mhWnd <> 0 Then
    CaretBlinkMs = CSGetCaretBlinkMs(mhWnd)
  End If
End Property

Public Property Let CaretBlinkMs(ByVal piMillisecs As Integer)
  If mhWnd <> 0 Then
    Call CSSetCaretBlinkMs(mhWnd, piMillisecs)
  End If
End Property

Public Function ClickAtCaretPos(ByVal peMsg As eWmMouseButton) As Integer
  If mhWnd <> 0 Then
    ClickAtCaretPos = CSClickAtCaretPos(mhWnd, CInt(peMsg))
  End If
End Function

Public Sub ShowScrollBar(ByVal pfShow As Boolean)
  If mhWnd Then
    Call apiShowScrollBar(mhWnd, SB_VERT, Abs(CInt(pfShow)))
  End If
End Sub

Public Sub ScrollUp()
  If mhWnd Then
    PostMessage mhWnd, WM_VSCROLL, SB_LINEUP, ByVal 0&
  End If
End Sub

Public Sub ScrollDown()
  If mhWnd Then
    PostMessage mhWnd, WM_VSCROLL, SB_LINEDOWN, ByVal 0&
  End If
End Sub

Public Sub ScrollTop()
  If mhWnd Then
    PostMessage mhWnd, WM_VSCROLL, SB_TOP, ByVal 0&
  End If
End Sub

Public Sub ScrollPageUp()
  If mhWnd Then
    PostMessage mhWnd, WM_VSCROLL, SB_PAGEUP, ByVal 0&
  End If
End Sub

Public Sub ScrollPageDown()
  If mhWnd Then
    PostMessage mhWnd, WM_VSCROLL, SB_PAGEDOWN, ByVal 0&
  End If
End Sub

Public Sub ScrollBottom()
  If mhWnd Then
    PostMessage mhWnd, WM_VSCROLL, SB_BOTTOM, ByVal 0&
  End If
End Sub

Public Sub ScrollTo(ByVal plLine As Long)
  If mhWnd Then
    SetScrollPos mhWnd, SB_VERT, CInt(plLine), 1
  End If
End Sub

Public Function IsRowVisible(ByVal piRow As Integer) As Boolean
  IsRowVisible = (piRow >= Me.TopLine) And (piRow <= (Me.TopLine + Me.MaxVisibleRows - 1))
End Function

Public Sub ShowBorder(ByVal pfShowBorder As Boolean)
  #If Win64 Then
    Dim lWinStyle As LongPtr
  #Else
    Dim lWinStyle As Long
  #End If
  If mhWnd = 0 Then Exit Sub
  
  lWinStyle = apiGetWindowLong(mhWnd, GWL_STYLE)
  If pfShowBorder Then
    apiSetWindowLong mhWnd, GWL_STYLE, lWinStyle Or WS_BORDER
  Else
    apiSetWindowLong mhWnd, GWL_STYLE, lWinStyle And Not WS_BORDER
  End If
  apiSetWindowPos mhWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER
End Sub

Public Function GetUnicodeCharCodes(ByRef palRetCharCodes() As Long) As Long
  If mhWnd = 0 Then Exit Function
  Dim i           As Long
  Dim lRangeCt    As Long
  Dim lTotGlyphCt As Long
  Dim iStart      As Integer
  Dim iGlyphCt    As Integer
  Dim k           As Long
  
  On Error GoTo GetUnicodeCharCodes_Err
  
  lRangeCt = CSLoadFontUnicodeRanges(mhWnd)
  For i = 1 To lRangeCt
    'Debug.Print "VP(Low)="; VarPtr(lStart); ", VP(count)="; VarPtr(lGlyphCt)
    If CSGetFontUnicodeRange(mhWnd, i, VarPtr(iStart), VarPtr(iGlyphCt)) Then
      'Debug.Print "VP(iStart)="; VarPtr(iStart); ", VP(iGlyphCt)="; VarPtr(iGlyphCt)
      If iGlyphCt Then
        ReDim Preserve palRetCharCodes(1 To lTotGlyphCt + iGlyphCt) As Long
        For k = 1 To iGlyphCt
          palRetCharCodes(lTotGlyphCt + k) = iStart + k - 1
        Next k
        lTotGlyphCt = lTotGlyphCt + iGlyphCt
      End If
    End If
  Next i
  CSUnloadFontUnicodeRanges mhWnd
  GetUnicodeCharCodes = lTotGlyphCt
  
GetUnicodeCharCodes_Exit:
  Exit Function

GetUnicodeCharCodes_Err:
  Resume GetUnicodeCharCodes_Exit
End Function

Public Sub SetAlphaTransparency(ByVal pbPercent As Byte)
  If mhWnd Then
    Call CSSetAlphaTransparency(mhWnd, pbPercent)
  End If
End Sub

Public Sub SetColorTransparency(ByVal pbPercent As Byte, ByVal plColor As Long, ByVal pfMakeTransparent As Boolean)
  If mhWnd Then
    Call CSSetColorTransparency(mhWnd, pbPercent, plColor, Abs(CByte(pfMakeTransparent)))
  End If
End Sub

Public Function GetZoneTag(ByVal piLine As Integer, ByVal piZoneID As Integer) As String
  Dim sBuffer     As String
  Dim iLen        As Integer
  Dim i           As Integer
  
  iLen = CSGetZoneTag(mhWnd, piLine, piZoneID, 0, 0)
  If iLen > 0 Then
    sBuffer = Space$(iLen + 1)
    iLen = CSGetZoneTag(mhWnd, piLine, piZoneID, Len(sBuffer), StrPtr(sBuffer))
    If iLen > 0 Then
      i = InStr(1, sBuffer, Chr$(0))
      If i > 0 Then
        sBuffer = left$(sBuffer, i - 1)
      End If
    End If
    sBuffer = Replace(sBuffer, Chr$(1), VT_EOM)
  End If
  
  GetZoneTag = sBuffer
End Function

Public Function GetZoneText(ByVal piLine As Integer, ByVal piZoneID As Integer, Optional ByVal pfInVT100 As Boolean = False) As String
  Dim sBuffer     As String
  Dim iLen        As Integer
  Dim i           As Integer
  
  iLen = CSGetZoneText(mhWnd, piLine, piZoneID, 0, 0, CInt(Abs(pfInVT100)))
  If iLen > 0 Then
    sBuffer = Space$(iLen + 1)
    iLen = CSGetZoneText(mhWnd, piLine, piZoneID, Len(sBuffer), StrPtr(sBuffer), CInt(Abs(pfInVT100)))
    If iLen > 0 Then
      i = InStr(1, sBuffer, Chr$(0))
      If i > 0 Then
        sBuffer = left$(sBuffer, i - 1)
      End If
    End If
    sBuffer = Replace(sBuffer, Chr$(1), VT_EOM)
  End If
  
  GetZoneText = sBuffer
End Function

Public Function GetLineText(ByVal piLine As Integer, Optional ByVal pfInVT100 As Boolean = False) As String
  Dim sBuffer     As String
  Dim iLen        As Integer
  Dim i           As Integer
  
  iLen = CSGetLineText(mhWnd, piLine, 0, 0, CInt(Abs(pfInVT100)))
  If iLen > 0 Then
    sBuffer = Space$(iLen + 1)
    iLen = CSGetLineText(mhWnd, piLine, Len(sBuffer), StrPtr(sBuffer), CInt(Abs(pfInVT100)))
    If iLen > 0 Then
      i = InStr(1, sBuffer, Chr$(0))
      If i > 0 Then
        sBuffer = left$(sBuffer, i - 1)
      End If
    End If
    sBuffer = Replace(sBuffer, Chr$(1), VT_EOM)
  End If
  
  GetLineText = sBuffer
End Function

#If Win64 Then
Public Function PaintOnDC(ByVal phDC As LongPtr, ByVal piStartLine As Integer, ByVal piEndLine As Integer, ByVal plWidth As Long, ByVal plHeight As Long) As Integer
#Else
Public Function PaintOnDC(ByVal phDC As Long, ByVal piStartLine As Integer, ByVal piEndLine As Integer, ByVal plWidth As Long, ByVal plHeight As Long) As Integer
#End If
  PaintOnDC = CSPaintOnDC(mhWnd, phDC, plWidth, plHeight, piStartLine, piEndLine)
End Function

Public Function GetLongestLineWidth() As Long
  GetLongestLineWidth = CSGetLongestLineWidth(mhWnd)
End Function

Public Function ResetAllLines()
  If mhWnd Then
    Call CSResetAllLines(mhWnd)
  End If
End Function

Public Sub CloseWindow()
  If mhWnd Then
    PostMessage mhWnd, WM_CLOSE, 0, ByVal 0&
    mhWnd = 0
  End If
End Sub

#If Win64 Then
Public Function GetMouseCursor() As LongPtr
#Else
Public Function GetMouseCursor() As Long
#End If
  If mhWnd Then
    GetMouseCursor = CSGetMouseCursor(mhWnd)
  End If
End Function

#If Win64 Then
Public Function SetMouseCursor(ByVal phCursor As LongPtr) As LongPtr
#Else
Public Function SetMouseCursor(ByVal phCursor As Long) As Long
#End If
  If mhWnd Then
    SetMouseCursor = CSSetMouseCursor(mhWnd, phCursor)
  End If
End Function

#If Win64 Then
Public Function GetHoverCursor() As LongPtr
#Else
Public Function GetHoverCursor() As Long
#End If
  If mhWnd Then
    GetHoverCursor = CSGetHoverCursor(mhWnd)
  End If
End Function

#If Win64 Then
Public Function SetHoverCursor(ByVal phCursor As LongPtr) As LongPtr
#Else
Public Function SetHoverCursor(ByVal phCursor As Long) As Long
#End If
  If mhWnd Then
    SetHoverCursor = CSSetHoverCursor(mhWnd, phCursor)
  End If
End Function

Public Property Get LineSpacing(ByVal peTopOrBottom As eLineVerticalLocation) As Integer
  If mhWnd Then
    Dim iTopOrBottom As Integer
    iTopOrBottom = CInt(peTopOrBottom)
    LineSpacing = CSGetLineSpacing(mhWnd, iTopOrBottom)
  End If
End Property

Public Property Let LineSpacing(ByVal peTopOrBottom As eLineVerticalLocation, ByVal piSpacing As Integer)
  If mhWnd Then
    Dim iTopOrBottom As Integer
    iTopOrBottom = CInt(peTopOrBottom)
    Call CSSetLineSpacing(mhWnd, piSpacing, iTopOrBottom)
  End If
End Property

Public Property Get LinePadding(ByVal peTopOrBottom As eLineVerticalLocation) As Integer
  If mhWnd Then
    Dim iTopOrBottom As Integer
    iTopOrBottom = CInt(peTopOrBottom)
    LinePadding = CSGetLinePadding(mhWnd, iTopOrBottom)
  End If
End Property

Public Property Let LinePadding(ByVal peTopOrBottom As eLineVerticalLocation, ByVal piPadding As Integer)
  If mhWnd Then
    Dim iTopOrBottom As Integer
    iTopOrBottom = CInt(peTopOrBottom)
    Call CSSetLinePadding(mhWnd, piPadding, iTopOrBottom)
  End If
End Property

Public Property Get LineHeight() As Integer
  If mhWnd Then
    LineHeight = CSGetLineHeight(mhWnd)
  End If
End Property

