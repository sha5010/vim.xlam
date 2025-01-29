Attribute VB_Name = "F_Mode"
Option Explicit
Option Private Module

Function ChangeToNormalMode(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If gVim Is Nothing Then
        ChangeToNormalMode = True
        Exit Function
    End If

    Call gVim.Mode.Change(MODE_NORMAL)
    Call SetStatusBar
    Exit Function

Catch:
    Call ErrorHandler("ChangeToNormalMode")
End Function

Function ChangeToShapeInsertMode(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If gVim Is Nothing Then
        ChangeToShapeInsertMode = True
        Exit Function
    End If

    Call gVim.Mode.Change(MODE_SHAPEINSERT)
    Call SetStatusBar("-- SHAPE INSERT (ESC to exit) --")
    Exit Function

Catch:
    Call ErrorHandler("ChangeToShapeInsertMode")
End Function

Private Sub ToggleVisualInner(ByVal visualLine As Boolean)
    On Error GoTo Catch

    If gVim Is Nothing Then
        Exit Sub
    End If

    With gVim.Mode
        If .Current <> MODE_VISUAL Then
            Call .Change(MODE_VISUAL)
            If Not .Visual Is Nothing Then
                .Visual.IsVisualLine = visualLine
            End If
            Exit Sub
        End If

        If Not .Visual Is Nothing Then
            If .Visual.IsVisualLine <> visualLine Then
                .Visual.IsVisualLine = visualLine
            Else
                Call ChangeToNormalMode
            End If
        End If
    End With
    Exit Sub

Catch:
    Call ErrorHandler("ToggleVisualInner")
End Sub

Function ToggleVisualMode(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    Call ToggleVisualInner(False)
    Exit Function

Catch:
    Call ErrorHandler("ToggleVisualMode")
End Function

Function ToggleVisualLine(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    Call ToggleVisualInner(True)
    Exit Function

Catch:
    Call ErrorHandler("ToggleVisualLine")
End Function

Function SwapVisualBase(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If gVim Is Nothing Then
        SwapVisualBase = True
        Exit Function
    ElseIf gVim.Mode.Current <> MODE_VISUAL Then
        SwapVisualBase = True
        Exit Function
    End If

    If Not gVim.Mode.Visual Is Nothing Then
        Call gVim.Mode.Visual.SwapBase
    End If
    Exit Function

Catch:
    Call ErrorHandler("SwapVisualBase")
End Function

Function StopVisualMode(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If gVim Is Nothing Then
        StopVisualMode = True
        Exit Function
    ElseIf gVim.Mode.Current <> MODE_VISUAL Then
        Exit Function
    End If

    Call ChangeToNormalMode
    Exit Function

Catch:
    Call ErrorHandler("StopVisualMode")
End Function
