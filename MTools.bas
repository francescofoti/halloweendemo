Attribute VB_Name = "MTools"
Option Compare Database
Option Explicit

'Split a string into a new array.
'Returns the number of elements in the array.
'If lMaxItems is specified, then the returned asRetItems() array
'will have at maximum lMaxItems, the last one holding the remaining
'chunk that wasn't splitted because the lMaxItems limit was reached.
Public Function SplitString(ByRef asRetItems() As String, _
  ByVal sToSplit As String, _
  Optional sSep As String = " ", _
  Optional lMaxItems As Long = 0&, _
  Optional eCompare As VbCompareMethod = vbBinaryCompare) _
  As Long

  Dim lPos        As Long
  Dim lDelimLen   As Long
  Dim lRetCount   As Long
  
  On Error Resume Next
  Erase asRetItems
  On Error GoTo SplitString_Err
  
  If Len(sToSplit) Then
    lDelimLen = Len(sSep)
    If lDelimLen Then
      lPos = InStr(1, sToSplit, sSep, eCompare)
      Do While lPos
        lRetCount = lRetCount + 1&
        ReDim Preserve asRetItems(1& To lRetCount)
        asRetItems(lRetCount) = left$(sToSplit, lPos - 1&)
        sToSplit = Mid$(sToSplit, lPos + lDelimLen)
        If lMaxItems Then
          If lRetCount = lMaxItems - 1& Then Exit Do
        End If
        lPos = InStr(1, sToSplit, sSep, eCompare)
      Loop
    End If
    lRetCount = lRetCount + 1&
    ReDim Preserve asRetItems(1& To lRetCount)
    asRetItems(lRetCount) = sToSplit
  End If
  SplitString = lRetCount
SplitString_Err:
End Function

