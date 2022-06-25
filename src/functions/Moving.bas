Attribute VB_Name = "F_Moving"
Option Explicit
Option Private Module

Function moveUp()
    Dim r As Long
    If gCount = 1 Then
        keybd_event vbKeyUp, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyUp, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        r = ActiveCell.Row - gCount
        If r < 1 Then
            r = 1
        End If
        ActiveSheet.Cells(r, ActiveCell.Column).Select
    End If
End Function

Function moveDown()
    Dim r As Long
    If gCount = 1 Then
        keybd_event vbKeyDown, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyDown, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        r = ActiveCell.Row + gCount
        If r > ActiveSheet.Rows.Count Then
            r = ActiveSheet.Rows.Count
        End If
        ActiveSheet.Cells(r, ActiveCell.Column).Select
    End If
End Function

Function moveLeft()
    Dim c As Long
    If gCount = 1 Then
        keybd_event vbKeyLeft, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyLeft, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        c = ActiveCell.Column - gCount
        If c < 1 Then
            c = 1
        End If
        ActiveSheet.Cells(ActiveCell.Row, c).Select
    End If
End Function

Function moveRight()
    Dim c As Long
    If gCount = 1 Then
        keybd_event vbKeyRight, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyRight, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        c = ActiveCell.Column + gCount
        If c > ActiveSheet.Columns.Count Then
            c = ActiveSheet.Columns.Count
        End If
        ActiveSheet.Cells(ActiveCell.Row, c).Select
    End If
End Function

Function moveToTopRow()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        If gCount = 1 Then
            .Cells(1, ActiveCell.Column).Select
        Else
            .Cells(gCount, ActiveCell.Column).Select
        End If
    End With
End Function

Function moveToLastRow()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        If gCount = 1 Then
            .Cells(.UsedRange.Item(.UsedRange.Count).Row, ActiveCell.Column).Select
        Else
            .Cells(gCount, ActiveCell.Column).Select
        End If
    End With
End Function

Function moveToFirstColumn()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.Row, 1).Select
    End With
End Function

Function moveToLeftEnd()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.Row, .UsedRange.Item(1).Column).Select
    End With
End Function

Function moveToRightEnd()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.Row, .UsedRange.Item(.UsedRange.Count).Column).Select
    End With
End Function

Function moveToTopOfCurrentRegion()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.CurrentRegion.Item(1).Row, ActiveCell.Column).Select
    End With
End Function

Function moveToBottomOfCurrentRegion()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.CurrentRegion.Item(ActiveCell.CurrentRegion.Count).Row, ActiveCell.Column).Select
    End With
End Function

Function moveToA1()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(1, 1).Select
    End With
End Function

Function moveToSpecifiedCell(ByVal Address As String) As Boolean
    Call recordToJumpList

    On Error GoTo Catch
    Address = Trim(Address)

    If reMatch(Address, "^[0-9]{1,7}$") Then
        ActiveSheet.Cells(CInt(Address), ActiveCell.Column).Select
        moveToSpecifiedCell = True

    ElseIf reMatch(Address, "^[a-z]{1,3}$", IgnoreCase:=True) Then
        ActiveSheet.Range(Address & ActiveCell.Row).Select
        moveToSpecifiedCell = True

    ElseIf reMatch(Address, "^[a-z]{1,3}[0-9]{1,7}(:[a-z]{1,3}[0-9]{1,7})?$", IgnoreCase:=True) Then
        ActiveSheet.Range(Address).Select
        moveToSpecifiedCell = True

    ElseIf reMatch(Address, "^[a-z]{1,3}:[a-z]{1,3}$", IgnoreCase:=True) Then
        ActiveSheet.Range(Address).Select
        moveToSpecifiedCell = True

    ElseIf reMatch(Address, "[0-9]{1,7}:[0-9]{1,7}") Then
        ActiveSheet.Range(Address).Select
        moveToSpecifiedCell = True

    End If

Catch:
End Function

Function moveToSpecifiedRow(ByVal n As String) As Boolean
    On Error GoTo Catch
    n = Trim(n)
    If reMatch(n, "^[0-9]{1,7}$") Then
        If CLng(n) > ActiveSheet.Rows.Count Then
            Exit Function
        End If

        Call recordToJumpList

        ActiveSheet.Cells(CLng(n), ActiveCell.Column).Select
        moveToSpecifiedRow = True
    End If

Catch:
End Function
