VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Consoul Window object managed by this form instance
Private moConsole As CConsoul
Private mconEmoji As CConsoul

'IConsoleOutput Interface Will link the moConsole Consoul object
'globally thru the MConsole module, and route global ConOut() and
'ConoutLn() calls to this instance.
Implements IConsoleOutput

Implements ICsZonePaintEventSink

Private Const CONSOLE_FONTNAME    As String = "Lucida Console"
Private Const CONSOLE_FONTSIZE    As Integer = 24
Private Const CONSOLE_MAXCAPACITY As Integer = 2000

Private mlBackColor As Long
Private mlForeColor As Long

#If Win64 Then
Private Declare PtrSafe Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
#Else
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
#End If

Public Property Get Console() As CConsoul
  Set Console = moConsole
End Property

' Create the Consoul window
Private Function CreateConsoul() As Boolean
  On Error GoTo CreateConsoul_Err
  
  Set moConsole = New CConsoul
  moConsole.FontName = CONSOLE_FONTNAME
  moConsole.FontSize = CONSOLE_FONTSIZE
  moConsole.MaxCapacity = CONSOLE_MAXCAPACITY
  moConsole.BackColor = mlBackColor
  moConsole.ForeColor = mlForeColor
  'We create a zero width/height console, and ajust size later
  If Not moConsole.Attach(Me.hWnd, 0, 0, 0, 0, piCreateAttributes:=LW_RENDERMODEBYLINE) Then
    MsgBox "Failed to create consoul window", vbCritical
    GoTo CreateConsoul_Exit
  End If
  moConsole.ShowWindow True
  
  'Set zone tracking for mouse cursor zone hover effect (after creation)
  moConsole.TrackZones = True
  'link console to callback that'll dispatch via CConsoulEventDispatcher
  moConsole.SetDrawZoneCallback AddressOf MCCallbacks.OnConsoulZonePaint
  
  CreateConsoul = True
  
CreateConsoul_Exit:
  Exit Function

CreateConsoul_Err:
  ShowError "CreateConsoul", Err.Number, "Failed to create consoul's output windows: " & Err.Description
End Function

Private Function CreateEmojiConsoul() As Boolean
  On Error GoTo CreateEmojiConsoul_Err
  
  Set mconEmoji = New CConsoul
  mconEmoji.FontName = Me.txtEmoji.FontName
  mconEmoji.FontSize = CONSOLE_FONTSIZE
  mconEmoji.MaxCapacity = 1
  mconEmoji.BackColor = mlBackColor
  mconEmoji.ForeColor = mlForeColor
  'We create a zero width/height console, and ajust size later
  If Not mconEmoji.Attach(Me.hWnd, 0, 0, 0, 0, piCreateAttributes:=LW_RENDERMODEBYLINE) Then
    MsgBox "Failed to create emoji consoul window", vbCritical
    GoTo CreateEmojiConsoul_Exit
  End If
  mconEmoji.ShowWindow True
  'we need to output something to force Consoul to compute char/line height
  mconEmoji.OutputLn " "
  moConsole.OutputLn " "
  moConsole.Clear
  'now we adjust the line spacing in the main console to match the char height of
  'the emoji console
  If moConsole.LineHeight < mconEmoji.CharHeight Then
    Dim iHalf As Integer
    iHalf = (mconEmoji.CharHeight - moConsole.LineHeight) \ 2
    moConsole.LineSpacing(elsTop) = iHalf + iHalf Mod 2
    moConsole.LineSpacing(elsBottom) = iHalf
  End If
  'Optional: Get rid of unneeded space above emojis, example:
  'mconEmoji.LineSpacing(elsTop) = -6
  
  CreateEmojiConsoul = True
  
CreateEmojiConsoul_Exit:
  Exit Function

CreateEmojiConsoul_Err:
  ShowError "CreateEmojiConsoul", Err.Number, "Failed to create consoul's output windows: " & Err.Description
End Function

'Provide a way to use the console output by handling an IConsoleOutput interface reference
Public Property Get IICOnsoleOutput() As IConsoleOutput
  Set IICOnsoleOutput = Me
End Property

Private Sub cmdSetZone_Click()
  Dim sEmojis As String
  sEmojis = Me.txtEmoji & ""
  If Len(sEmojis) = 0 Then
    Exit Sub
  End If
  mconEmoji.OutputLn VT_FCOLOR(QBColor(QBCOLOR_YELLOW)) & sEmojis 'Note that we have a mconEmoji.MaxCapacity of 1, so this will replace the line
  mconEmoji.RefreshWindow
  'we don't know on which line is our zone, so we'll refresh the full
  'Consoul window (there's no so many lines...)
  moConsole.RefreshWindow
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    CloseConsoleWindow
  End If
End Sub

Private Sub Form_Load()
  On Error Resume Next
  MaximizeWindow Me.hWnd
  SetConsoleOutput Me
  
  mlBackColor = RGB(0, 13, 54) 'more "night blue" than "deep black"
  mlForeColor = QBColor(QBCOLOR_WHITE)
  Me.Section(AcSection.acDetail).BackColor = mlBackColor
  'we need to create the emoji console after the main console
  Me.txtEmoji.FontSize = CONSOLE_FONTSIZE
  Call CreateConsoul
  Call CreateEmojiConsoul
  
  'Register ourselves as target for zone paint events for this console
  ConsoulEventDispatcher.RegisterEventSink moConsole.hWnd, Me, eCsZonePaint
  Form_Resize 'adjust console window size & position
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  If moConsole Is Nothing Then Exit Sub
  
  'Adjust to full form client area of the "Detail" section
  'We keep an arbitrary pixels margin around the consoul window.
  Dim rcClient      As RECT
  Dim iWidth        As Integer
  Dim iHeight       As Integer
  Dim iHeaderHeight As Integer
  Dim iFooterHeight As Integer
  Dim iLeft         As Integer
  
  Const MARGIN As Long = 8&
  
  GetClientRect Me.hWnd, rcClient
  
  iHeaderHeight = TwipsToPixelsY(Me.Section(AcSection.acHeader).Height)
  iFooterHeight = TwipsToPixelsY(Me.Section(AcSection.acFooter).Height)
  
  iWidth = rcClient.Right - rcClient.left - 2 * MARGIN
  iHeight = rcClient.Bottom - rcClient.Top - 2 * MARGIN - iHeaderHeight - iFooterHeight
  
  moConsole.MoveWindow MARGIN, MARGIN + iHeaderHeight, iWidth, iHeight
  'we put the emoji console at the right of cmdSetZone
  iLeft = TwipsToPixelsX(Me.cmdSetZone.left + Me.cmdSetZone.Width) + MARGIN
  mconEmoji.MoveWindow iLeft, MARGIN, iWidth - iLeft, iHeaderHeight - 2 * MARGIN
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set mconEmoji = Nothing
  SetConsoleOutput Nothing
  If Not moConsole Is Nothing Then
    ConsoulEventDispatcher.UnregisterEventSink moConsole.hWnd
  End If
  Set moConsole = Nothing
End Sub

Private Sub IConsoleOutput_ConOut(ByVal psInfo As String)
  If moConsole Is Nothing Then Exit Sub
  moConsole.Output psInfo
End Sub

Private Function IConsoleOutput_ConOutLn(ByVal psInfo As String, Optional ByVal piQBColorText As Variant) As Integer
  If moConsole Is Nothing Then Exit Function
  If IsMissing(piQBColorText) Then
    IConsoleOutput_ConOutLn = moConsole.OutputLn(psInfo)
  Else
    IConsoleOutput_ConOutLn = moConsole.OutputLn(psInfo, piQBColorText)
  End If
  DoEvents 'Leaves time to the child window to process messages
End Function

Private Function ICsZonePaintEventSink_OnZonePaint(ByVal phWnd As Long, ByVal phDC As Long, ByVal piZoneID As Integer, ByVal piLine As Long, ByVal piLeft As Integer, ByVal piTop As Integer, ByVal piRight As Integer, ByVal piBottom As Integer) As Integer
  If piZoneID <> ZONEID_DEMO Then Exit Function
  
  Dim ptPrev As POINTAPI
  Dim lResult As Long
  
  'The PaintOnDC method only paints beginning at pixel [0,0],
  'so we temporarly move the coordinates system origin to the
  'beginning of our zone.
  lResult = SetViewportOrgEx(phDC, piLeft, piTop, ptPrev)
  'Paint the emoji console here
  mconEmoji.PaintOnDC phDC, 1, 1, piRight - piLeft, piBottom - piTop
  'Restore the viewport origin
  lResult = SetViewportOrgEx(phDC, ptPrev.X, ptPrev.Y, ptPrev)
End Function
