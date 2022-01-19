Sub SelectionSqrt()
    Dim cell As Range
    Dim ErrMsg As String
    If TypeName(Selection) <> “Range” Then Exit Sub
    On Error GoTo ErrorHandler
    For Each cell In Selection
        cell.Value = Sqr(cell.Value)
    Next cell
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 5 'Negative number"
            Resume Next
        Case 13 'Type mismatch
            Resume Next
        Case 1004 'Locked cell, protected sheet
            MsgBox "Cell is locked. Try again.", vbCritical, cell.Address
            Exit Sub
        Case Else
            ErrMsg = Error(Err.Number)
            MsgBox "ERROR: " & ErrMag, vbCritical, cell, Address
         Exit Sub
    End Select
End Sub
