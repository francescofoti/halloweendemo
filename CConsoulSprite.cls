VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConsoulSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private miX       As Integer
Private miY       As Integer
Private miWidth   As Integer
Private miHeight  As Integer

Public Enum eSpriteMoveDir
  eMoveStandStill = 0
  eMoveIncrease = 1
  eMoveDecrease = 2
End Enum

Private moConsole             As CConsoul
Private masLines()            As String
Private miRows                As Integer
Private miCols                As Integer
Private mvTransparentColor    As Variant
Private miTransparentLevel    As Integer
Private miAlphaTransparency   As Integer
Private mfTransparent         As Boolean
Private mvBackColor           As Variant
Private msFontName            As String
Private miFontSize            As Integer
Private mfVisible             As Boolean
Private mbMoveDirH            As eSpriteMoveDir
Private mbMoveDirV            As eSpriteMoveDir
Private miMoveIncrH           As Integer
Private miMoveIncrV           As Integer

Private Sub DestroySprite()
  On Error Resume Next
  If Not moConsole Is Nothing Then
    ConsoulEventDispatcher.UnregisterEventSink moConsole.hWnd
    moConsole.Detach
  End If
  Set moConsole = Nothing
  Erase masLines()
  miX = 0
  miY = 0
  miWidth = 0
  miHeight = 0
  miRows = 0
  miCols = 0
  mfTransparent = False
End Sub

Private Sub Class_Initialize()
  mvBackColor = vbBlack
  mvTransparentColor = vbBlack
  miTransparentLevel = 100
End Sub

Private Sub Class_Terminate()
  DestroySprite
End Sub

Public Property Get hWnd() As LongPtr
  If Not moConsole Is Nothing Then
    hWnd = moConsole.hWnd
  End If
End Property

Public Property Get X() As Integer
  X = miX
End Property

Public Property Get Y() As Integer
  Y = miY
End Property

Public Property Get Width() As Integer
  Width = miWidth
End Property

Public Property Get Height() As Integer
  Height = miHeight
End Property

Public Property Get BackColor() As Variant
  BackColor = mvBackColor
End Property

Public Property Let BackColor(ByVal pvBackColor As Variant)
  mvBackColor = pvBackColor
End Property

Public Property Get TransparentColor() As Variant
  TransparentColor = mvTransparentColor
End Property

Public Property Let TransparentColor(ByVal pvTransparentColor As Variant)
  mvTransparentColor = pvTransparentColor
End Property

Public Property Get TransparentLevel() As Integer
  TransparentLevel = miTransparentLevel
End Property

Public Property Let TransparentLevel(ByVal piLevel As Integer)
  miTransparentLevel = piLevel
End Property

Public Property Get AlphaTransparency() As Integer
  AlphaTransparency = miAlphaTransparency
End Property

Public Property Let AlphaTransparency(ByVal piLevel As Integer)
  miAlphaTransparency = piLevel
End Property

Public Property Let Transparent(ByVal pfTransparent As Boolean)
  If pfTransparent Then
    If miAlphaTransparency > 0 Then
      moConsole.SetAlphaTransparency miAlphaTransparency
    End If
    If Not IsNull(mvTransparentColor) Then
      moConsole.SetColorTransparency miTransparentLevel, mvTransparentColor, True
    End If
  Else
    moConsole.SetAlphaTransparency 0
  End If
  mfTransparent = pfTransparent
End Property

Public Property Get Transparent() As Boolean
  Transparent = mfTransparent
End Property

Public Property Get Console() As CConsoul
  Set Console = moConsole
End Property

Public Property Get FontName() As String
  FontName = msFontName
End Property

Public Property Get FontSize() As Integer
  FontSize = miFontSize
End Property

Public Property Get Visible() As Boolean
  Visible = mfVisible
End Property

Public Property Let Visible(ByVal pfVisible As Boolean)
  mfVisible = pfVisible
  If Not moConsole Is Nothing Then
    moConsole.ShowWindow pfVisible
  End If
End Property

Public Function CreateSprite( _
    ByVal phWndParent As Long, _
    ByVal piRows As Integer, _
    ByVal piCols As Integer, _
    ByVal psFontName As String, _
    ByVal piFontSize As Integer, _
    ByRef pasLines() As String, _
    ByRef pIICsMouseEventSink As ICsMouseEventSink _
  ) As Boolean
  On Error GoTo CreateSprite_Err
  Const LOCAL_ERR_CTX As String = "CreateSprite"
  Dim i As Integer
  
  DestroySprite
  ReDim masLines(1 To piRows) As String
  Set moConsole = New CConsoul
  
  moConsole.FontName = psFontName
  moConsole.FontSize = piFontSize
  moConsole.MaxCapacity = piRows
  If Not IsNull(mvBackColor) Then
    moConsole.BackColor = mvBackColor
  End If
  
  'We create a zero width/height console, and ajust size later
  If Not moConsole.Attach(phWndParent, 0, 0, 0, 0, AddressOf MCCallbacks.OnConsoulMouseButton, piCreateAttributes:=LW_RENDERMODEBYLINE) Then
    GoTo CreateSprite_Exit
  End If
  For i = 1 To piRows
    masLines(i) = pasLines(i)
    moConsole.OutputLn pasLines(i)
  Next i
  'Trick to force the rendering and compute the char dimensions
  moConsole.ShowWindow True
  moConsole.MoveWindow 0, 0, 1, 1
  moConsole.ShowBorder False
  moConsole.ShowScrollBar False
  moConsole.ShowWindow False
  
  miRows = piRows
  miCols = piCols
  msFontName = psFontName
  miFontSize = piFontSize
  miWidth = moConsole.MaxCharWidth * miCols
  miWidth = moConsole.TextWidth(moConsole.GetLineText(moConsole.GetLongestLineWidth))
  miHeight = moConsole.CharHeight * miRows
  
  ConsoulEventDispatcher.RegisterEventSink moConsole.hWnd, pIICsMouseEventSink, eCsMouseEvent
  
  CreateSprite = True
  
CreateSprite_Exit:
  Exit Function

CreateSprite_Err:
  If Not moConsole Is Nothing Then
    Set moConsole = Nothing
  End If
  Resume CreateSprite_Exit
End Function

Public Sub SetPosition(ByVal piX As Integer, ByVal piY As Integer)
  miX = piX
  miY = piY
  If Not moConsole Is Nothing Then
    moConsole.MoveWindow piX, piY, miWidth, miHeight, 1
  End If
End Sub

'--- obsoletetion going on:
Public Sub Move(ByVal piX As Integer, ByVal piY As Integer)
  miX = piX
  miY = piY
  If Not moConsole Is Nothing Then
    moConsole.MoveWindow piX, piY, miWidth, miHeight, 1
  End If
End Sub

Public Property Get MoveDirV() As eSpriteMoveDir
  MoveDirV = mbMoveDirV
End Property

Public Property Let MoveDirV(ByVal pbMoveDirV As eSpriteMoveDir)
  mbMoveDirV = pbMoveDirV
End Property

Public Property Get MoveDirH() As eSpriteMoveDir
  MoveDirH = mbMoveDirH
End Property

Public Property Let MoveDirH(ByVal pbMoveDirH As eSpriteMoveDir)
  mbMoveDirH = pbMoveDirH
End Property

Private Function GetMoveDirIncrement(ByVal peMoveDir As eSpriteMoveDir) As Integer
  Select Case peMoveDir
  Case eSpriteMoveDir.eMoveStandStill
    GetMoveDirIncrement = 0
  Case eSpriteMoveDir.eMoveIncrease
    GetMoveDirIncrement = 1
  Case eSpriteMoveDir.eMoveDecrease
    GetMoveDirIncrement = -1
  End Select
End Function

Public Property Get MoveIncrV() As Integer
  MoveIncrV = GetMoveDirIncrement(mbMoveDirV) * miMoveIncrV
End Property

Public Property Let MoveIncrV(ByVal piMoveIncrV As Integer)
  miMoveIncrV = piMoveIncrV
End Property

Public Property Get MoveIncrH() As Integer
  MoveIncrH = GetMoveDirIncrement(mbMoveDirH) * miMoveIncrH
End Property

Public Property Let MoveIncrH(ByVal piMoveIncrH As Integer)
  miMoveIncrH = piMoveIncrH
End Property

