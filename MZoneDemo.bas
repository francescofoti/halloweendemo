Attribute VB_Name = "MZoneDemo"
Option Compare Database
Option Explicit

#If Win64 Then
  Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong
#Else
  Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#End If

Public Const ZONEID_DEMO As Integer = 100

Public Function GetRandom(ByVal iLo As Long, ByVal iHi As Long) As Long
  GetRandom = Int(iLo + (Rnd * (iHi - iLo + 1)))
End Function

Public Sub EmojiDemo()
  OpenConsoleWindow
  
  ConOutLn "Consoul library version " & GetConsole().ConsoulVersion
  ConOutLn ""
  'If we do not have the same backcolor/forecolor in the mconEmoji and moConsole
  '(in Form_Canvas), we have to restore backcolor and forecolor after painting
  'the zone, as we just paint one console over the other and do not restore
  'proper DC properties after the PaintOnDC call:
  'ConOutLn "Your emoji(s) appear in yellow " & _
           VTX_ZONE_BEGIN(ZONEID_DEMO) & "     " & VTX_ZONE_END(ZONEID_DEMO) & _
           VT_BCOLOR(GetConsole().BackColor) & VT_FCOLOR(GetConsole().ForeColor) & "<-- here."
  'If we have the same color on both consoles, that's ok:
  ConOutLn "Your emoji(s) appear in yellow " & _
           VTX_ZONE_BEGIN(ZONEID_DEMO) & "     " & VTX_ZONE_END(ZONEID_DEMO) & "<-- here."
  ConOutLn ""
  ConOutLn "Press [Esc] to close the console window"
End Sub

