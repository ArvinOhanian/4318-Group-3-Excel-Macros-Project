Sub ID()
' ID Macro
' Outputs last 4 digits of IDs
' Checks if there are any repeats
' Will change output to last 5 digits if no repeats
' Note: Original IDs must be left of selected region
' Ex: If Original IDs are in A, output will be in B
' Therefore highlight columns right of original inputs
' Numbers must be formatted similar to Bronco ID
    Dim myRange As Range
    Dim myCell As Range
    Dim rep As Boolean
    Dim temp As String
    Set myRange = Selection
    'set up table
    ActiveCell.Offset(-1, -1).Activate
    If IsEmpty(ActiveCell) Then
        ActiveCell.Value = "ID"
    End If
    ActiveCell.Offset(0, 1).Activate
    If IsEmpty(ActiveCell) Then
        ActiveCell.Value = "Output"
    End If
    ActiveCell.Offset(1, 0).Activate
    'initialize rep
    rep = False
    '4 digit ID
    temp = "=MID(RC[-1],6,4)"
    For Each myCell In myRange
        'used to not overwrite data
        If IsEmpty(myCell) Then
            myCell.Value = temp
        End If
    Next myCell
    'first check
    For Each myCell In myRange
        If WorksheetFunction.CountIf(myRange, myCell.Value) > 1 Then
            rep = True
        End If
    Next myCell
    '5 digit ID if there are repeats
    temp = "=MID(RC[-1],5,5)"
    For Each myCell In myRange
        If (rep) Then
            myCell.Value = temp
        End If
    Next myCell
    'reset rep to false
    rep = False
    'second check
    For Each myCell In myRange
        If WorksheetFunction.CountIf(myRange, myCell.Value) > 1 Then
        rep = True
    End If
    Next myCell
    '6 digit ID if there are still repeats
    temp = "=MID(RC[-1],4,6)"
    For Each myCell In myRange
        If (rep) Then
            myCell.Value = "=MID(RC[-1],4,6)"
        End If
    Next myCell
    rep = False
End Sub
