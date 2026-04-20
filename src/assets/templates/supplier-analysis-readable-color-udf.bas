Attribute VB_Name = "SupplierAnalysisReadableColor"
Option Explicit

' Import into supplier-analysis-macros.xlsm (VBA Editor: File > Import File) if these UDFs are missing.
' The Angular export references SumByReadableColor and CountByReadableColor.

Private Function CellIsWinnerHighlight(ByVal c As Range) As Boolean
    On Error Resume Next
    Dim ic As Long
    ic = c.Interior.Color
    On Error GoTo 0
    ' Export: single winner yellow; tie uses light green (RGB 198,239,206 -> BGR &HCEEFC6)
    CellIsWinnerHighlight = (ic = &HFFFF&) Or (ic = &HCEEFC6)
End Function

Private Function CellMatchesReadable(ByVal c As Range, ByVal colorName As String) As Boolean
    Dim n As String
    n = LCase$(Trim$(colorName))
    Select Case n
        Case "yellow"
            CellMatchesReadable = CellIsWinnerHighlight(c)
        Case "notyellow", "not yellow", "transparent"
            CellMatchesReadable = Not CellIsWinnerHighlight(c)
        Case Else
            CellMatchesReadable = False
    End Select
End Function

Public Function SumByReadableColor(ByVal rangeForColor As Range, ByVal rangeToSum As Range, ByVal colorName As String) As Double
    Dim i As Long, n As Long
    Dim c As Range, s As Range
    SumByReadableColor = 0
    n = rangeForColor.Cells.Count
    If n <> rangeToSum.Cells.Count Then Exit Function
    For i = 1 To n
        Set c = rangeForColor.Cells(i)
        Set s = rangeToSum.Cells(i)
        If CellMatchesReadable(c, colorName) Then
            If IsNumeric(s.Value2) Then SumByReadableColor = SumByReadableColor + CDbl(s.Value2)
        End If
    Next i
End Function

Public Function CountByReadableColor(ByVal rangeForColor As Range, ByVal rangeToCount As Range, ByVal colorName As String) As Long
    Dim i As Long, n As Long
    Dim c As Range, t As Range
    CountByReadableColor = 0
    n = rangeForColor.Cells.Count
    If n <> rangeToCount.Cells.Count Then Exit Function
    For i = 1 To n
        Set c = rangeForColor.Cells(i)
        Set t = rangeToCount.Cells(i)
        If CellMatchesReadable(c, colorName) Then CountByReadableColor = CountByReadableColor + 1
    Next i
End Function
