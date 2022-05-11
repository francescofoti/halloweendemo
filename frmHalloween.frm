VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmHalloween"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const PUMPKIN_FONTNAME    As String = "Lucida Console"
Private Const PUMPKIN_FONTSIZE    As Integer = 8
Private Const PUMPKIN_LINES       As Integer = 25
Private Const PUMPKIN_COLS        As Integer = 27
Private Const PUMPKIN_COUNT       As Integer = 20

Private masPumpkinLines() As String
Private mcolPumpkins      As Collection

Implements ICsMouseEventSink

Private Sub DestroyPumpkinSprites()
  Set mcolPumpkins = Nothing
End Sub

Private Function CreatePumpkinSprites() As Boolean
  Dim i         As Integer
  Dim fOK       As Boolean
  Dim iPosX     As Integer
  Dim iPosY     As Integer
  Dim rcWindow  As RECT
  Dim oPumpkin  As CConsoulSprite
  
  GetClientRect Me.hWnd, rcWindow
  
  Set mcolPumpkins = New Collection
  
  For i = 1 To PUMPKIN_COUNT
    'As we set a random color for each pumkin, we have to recreate the pumpkin lines array for the sprite
    BuildPumpkinLinesArray masPumpkinLines
    Set oPumpkin = New CConsoulSprite
    'When we create a sprite, we give it ourself as a ICsMouseEventSink interface, so we're called back when any sprite is clicked
    fOK = oPumpkin.CreateSprite(Me.hWnd, PUMPKIN_LINES, PUMPKIN_LINES, PUMPKIN_FONTNAME, PUMPKIN_FONTSIZE, masPumpkinLines, Me)
    If fOK Then
      mcolPumpkins.Add oPumpkin
      'random position
      iPosX = GetRandom(0, rcWindow.Right - rcWindow.left)
      iPosY = GetRandom(0, rcWindow.Bottom - rcWindow.Top)
      'random speed
      oPumpkin.MoveIncrH = GetRandom(1, 5)
      oPumpkin.MoveIncrV = GetRandom(1, 25)
      oPumpkin.Move iPosX, iPosY
      'show sprite
      oPumpkin.Visible = True
      oPumpkin.Transparent = True
      'Set random H/V directions
      oPumpkin.MoveDirH = GetRandom(eSpriteMoveDir.eMoveStandStill, eSpriteMoveDir.eMoveDecrease)
      oPumpkin.MoveDirV = GetRandom(eSpriteMoveDir.eMoveStandStill, eSpriteMoveDir.eMoveDecrease)
    Else
      MsgBox "Failed to create sprite #" & i, vbCritical
      DestroyPumpkinSprites
      Exit For
    End If
  Next i
  
End Function

Private Sub Form_Load()
  On Error Resume Next
  MaximizeWindow Me.hWnd
  Call CreatePumpkinSprites
  Me.TimerInterval = 5
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  '/**/
End Sub

Private Sub Form_Timer()
  AnimatePumpkins
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.TimerInterval = 0
  DestroyPumpkinSprites
End Sub

'Pumpkin drawing in VT100, generated with AsciiPaint
Sub BuildPumpkinLinesArray(ByRef paPumpkinLines() As String)
  ReDim paPumpkinLines(1 To PUMPKIN_LINES) As String
  Dim i               As Integer
  Dim sTemp           As String
  Dim sPumpkinColor   As String
  
  i = GetRandom(1, 4)
  Select Case i
  Case 1
    sPumpkinColor = "0080FF"    'orange
  Case 2
    sPumpkinColor = "FF8080"  'blue
  Case 3
    sPumpkinColor = "80FF80"  'green
  Case 4
    sPumpkinColor = "80FFF0"  'yellow
  End Select
  i = 0
  
  sTemp = String$(14, &H20) & ChrW$(&H1B) & "[38;$8000m" & ChrW$(&H1B) & "[48;$3EC170m" & String$(3, &H2591) & ChrW$(&H1B) & "[0m" & String$(10, &H20)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = String$(14, &H20) & ChrW$(&H1B) & "[38;$8000m" & ChrW$(&H1B) & "[48;$3EC170m" & String$(3, &H2591) & ChrW$(&H1B) & "[0m" & String$(10, &H20)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = String$(13, &H20) & ChrW$(&H1B) & "[38;$8000m" & ChrW$(&H1B) & "[48;$3EC170m" & String$(3, &H2591) & ChrW$(&H1B) & "[0m" & String$(11, &H20)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = String$(13, &H20) & ChrW$(&H1B) & "[38;$8000m" & ChrW$(&H1B) & "[48;$3EC170m" & String$(3, &H2591) & ChrW$(&H1B) & "[0m" & String$(11, &H20)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = String$(12, &H20) & ChrW$(&H1B) & "[38;$8000m" & ChrW$(&H1B) & "[48;$3EC170m" & String$(3, &H2591) & ChrW$(&H1B) & "[0m" & String$(12, &H20)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = String$(7, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & ChrW$(&HA0) & String$(2, &H20) & ChrW$(&H1B) & "[38;$3EC170m" & ChrW$(&H1B) & "[48;$8000m" & String$(2, &H2591) & ChrW$(&H1B) & "[38;$8000m" & ChrW$(&H1B) & "[48;$3EC170m" & String$(3, &H2591) & ChrW$(&H1B) & "[48;$8000m" & String$(2, &H2591) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(3, &H20) & ChrW$(&H1B) & "[0m" & String$(3, &H20) & ChrW$(&H1B) & "[0m" & String$(4, &H20)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = String$(4, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(7, &H20) & ChrW$(&H1B) & "[38;$3EC170m" & ChrW$(&H1B) & "[48;$8000m" & String$(5, &H2591) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(7, &H20) & ChrW$(&H1B) & "[0m" & String$(4, &H20)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = String$(3, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(21, &H20) & ChrW$(&H1B) & "[0m" & String$(3, &H20)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = String$(2, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(23, &H20) & ChrW$(&H1B) & "[0m" & String$(2, &H20)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = " " & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(6, &H20) & ChrW$(&H1B) & "[0m" & String$(2, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(9, &H20) & ChrW$(&H1B) & "[0m" & String$(2, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(6, &H20) & ChrW$(&H1B) & "[0m "
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(7, &H20) & ChrW$(&H1B) & "[0m" & String$(3, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(7, &H20) & ChrW$(&H1B) & "[0m" & String$(3, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(6, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m" & ChrW$(&H2591)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(7, &H20) & ChrW$(&H1B) & "[0m" & String$(4, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(5, &H20) & ChrW$(&H1B) & "[0m" & String$(4, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(6, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m" & ChrW$(&H2591)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(26, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m" & ChrW$(&H2591)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(11, &H20) & ChrW$(&H1B) & "[0m" & String$(5, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(10, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m" & ChrW$(&H2591)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(12, &H20) & ChrW$(&H1B) & "[0m" & String$(3, &H20) & ChrW$(&H1B) & "[48;$73E6m " & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(10, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m" & ChrW$(&H2591)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(13, &H20) & ChrW$(&H1B) & "[0m " & ChrW$(&H1B) & "[48;$73E6m " & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(11, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m "
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(4, &H20) & ChrW$(&H1B) & "[0m " & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(8, &H20) & ChrW$(&H1B) & "[48;$73E6m " & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(8, &H20) & ChrW$(&H1B) & "[0m " & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(3, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m "
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(4, &H20) & ChrW$(&H1B) & "[0m" & String$(2, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(2, &H20) & ChrW$(&H1B) & "[0m" & String$(4, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(2, &H20) & ChrW$(&H1B) & "[48;$73E6m " & ChrW$(&H1B) & "[0m" & String$(6, &H20) & ChrW$(&H1B) & "[0m" & String$(2, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(3, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m "
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(5, &H20) & ChrW$(&H1B) & "[0m" & String$(7, &H20) & ChrW$(&H1B) & "[48;$73E6m" & String$(3, &H20) & ChrW$(&H1B) & "[0m" & String$(7, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(4, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m "
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = " " & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(5, &H20) & ChrW$(&H1B) & "[0m" & String$(3, &H20) & ChrW$(&H1B) & "[48;$80FFm " & ChrW$(&H1B) & "[48;$73E6m " & ChrW$(&H1B) & "[0m" & String$(6, &H20) & ChrW$(&H1B) & "[0m" & String$(4, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(3, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m" & String$(2, &H20) & ChrW$(&H1B) & "[0m "
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = String$(2, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(5, &H20) & ChrW$(&H1B) & "[0m" & String$(2, &H20) & ChrW$(&H1B) & "[48;$80FFm " & ChrW$(&H1B) & "[48;$73E6m " & ChrW$(&H1B) & "[0m" & String$(5, &H20) & ChrW$(&H1B) & "[48;$80FFm " & ChrW$(&H1B) & "[48;$73E6m " & ChrW$(&H1B) & "[0m" & String$(2, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(3, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m" & String$(2, &H20) & ChrW$(&H1B) & "[0m " & ChrW$(&H1B) & "[0m "
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = String$(2, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(8, &H20) & ChrW$(&H1B) & "[48;$73E6m " & ChrW$(&H1B) & "[0m" & String$(5, &H20) & ChrW$(&H1B) & "[48;$80FFm " & ChrW$(&H1B) & "[48;$73E6m " & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(5, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m" & String$(2, &H20) & ChrW$(&H1B) & "[0m " & ChrW$(&H1B) & "[0m "
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = String$(3, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(18, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m" & String$(3, &H20) & ChrW$(&H1B) & "[0m" & String$(3, &H20)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = String$(4, &H20) & ChrW$(&H1B) & "[48;$" & sPumpkinColor & "m" & String$(15, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m" & String$(4, &H20) & ChrW$(&H1B) & "[0m" & String$(4, &H20)
  i = i + 1: paPumpkinLines(i) = sTemp
  sTemp = String$(7, &H20) & ChrW$(&H1B) & "[38;$73E6m" & ChrW$(&H1B) & "[48;$73E6m" & String$(13, &H20) & ChrW$(&H1B) & "[0m" & String$(7, &H20)
  i = i + 1: paPumpkinLines(i) = sTemp
  
End Sub

Private Sub AnimatePumpkins()
  Dim i       As Integer
  Dim xPos    As Integer
  Dim yPos    As Integer
  Dim rcClient  As RECT
  Dim oPumpkin  As CConsoulSprite
  
  GetClientRect Me.hWnd, rcClient
  
  If mcolPumpkins Is Nothing Then
    Me.TimerInterval = 0
    MsgBox "No sprites defined", vbCritical
    Exit Sub
  End If
  
  For Each oPumpkin In mcolPumpkins
    xPos = oPumpkin.X + oPumpkin.MoveIncrH
    yPos = oPumpkin.Y + oPumpkin.MoveIncrV
    If xPos > (rcClient.Right - oPumpkin.Width) Then
      oPumpkin.MoveDirH = eMoveDecrease
    End If
    If xPos < 0 Then
      oPumpkin.MoveDirH = eMoveIncrease
    End If
    If yPos > (rcClient.Bottom - oPumpkin.Height) Then
      oPumpkin.MoveDirV = eMoveDecrease
    End If
    If yPos < 0 Then
      oPumpkin.MoveDirV = eMoveIncrease
    End If
    oPumpkin.Move xPos, yPos
  Next
End Sub

Private Function ICsMouseEventSink_OnMouseButton(ByVal phWnd As Long, ByVal piEvtCode As Integer, ByVal pwParam As Integer, ByVal piZoneID As Integer, ByVal piRow As Integer, ByVal piCol As Integer, ByVal piPosX As Integer, ByVal piPosY As Integer) As Integer
  Dim oPumpkin  As CConsoulSprite
  Dim i         As Integer
  
  'we know the hWnd is one of our sprites
  For i = 1 To mcolPumpkins.Count
    Set oPumpkin = mcolPumpkins(i)
    If oPumpkin.hWnd = phWnd Then
      Set oPumpkin = Nothing
      mcolPumpkins.Remove i
      Exit For
    End If
    Set oPumpkin = Nothing
  Next
End Function
