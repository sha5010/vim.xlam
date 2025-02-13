Attribute VB_Name = "F_Sheets"
Option Explicit
Option Private Module

Function NextSheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Long

    With ActiveWorkbook
        i = .ActiveSheet.Index
        Dim cnt As Long: cnt = gVim.Count1

        Do While cnt > 0
            i = (i Mod .Sheets.Count) + 1
            If .Sheets(i).Visible = xlSheetVisible Then
                cnt = cnt - 1
            End If

            If i = .ActiveSheet.Index Then
                Dim visibleSheets As Long
                visibleSheets = gVim.Count1 - cnt
                cnt = cnt Mod visibleSheets
            End If
        Loop
        .Sheets(i).Activate
    End With
    Exit Function

Catch:
    For i = 1 To gVim.Count1
        Call KeyStroke(Ctrl_ + PageDown_)
    Next i
    Call ErrorHandler("NextSheet")
End Function

Function PreviousSheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Long

    With ActiveWorkbook
        i = .ActiveSheet.Index
        Dim cnt As Long: cnt = gVim.Count1

        Do While cnt > 0
            i = ((i - 2 + .Sheets.Count) Mod .Sheets.Count) + 1
            If .Sheets(i).Visible = xlSheetVisible Then
                cnt = cnt - 1
            End If

            If i = .ActiveSheet.Index Then
                Dim visibleSheets As Long
                visibleSheets = gVim.Count1 - cnt
                cnt = cnt Mod visibleSheets
            End If
        Loop
        .Sheets(i).Activate
    End With
    Exit Function

Catch:
    For i = 1 To gVim.Count1
        Call KeyStroke(Ctrl_ + PageUp_)
    Next i
    Call ErrorHandler("PreviousSheet")
End Function

Function RenameSheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim ret As String
    Dim beforeName As String

    With ActiveWorkbook
        beforeName = .ActiveSheet.Name
        ret = InputBox(gVim.Msg.EnterNewSheetName, gVim.Msg.RenameSheetTitle, beforeName)
        Call DisableIME

        If ret <> "" Then
            'Exit if same name
            If ret = beforeName Then
                Exit Function

            'Error when new sheet name already exists
            ElseIf IsSheetExists(ret) Then
                MsgBox gVim.Msg.SheetAlreadyExists(ret), vbExclamation
                Exit Function
            End If
            .Sheets(.ActiveSheet.Index).Name = ret

            Call SetStatusBarTemporarily(gVim.Msg.ChangedSheetName & _
                ": """ & beforeName & """ -> """ & ret & """", 3000)
        End If
    End With
    Exit Function

Catch:
    Call ErrorHandler("RenameSheet")
End Function

Function MoveSheetForward(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim idx As Integer
    Dim cnt As Integer
    Dim n As Integer
    Dim i As Integer
    Dim warpFlag As Boolean

    With ActiveWorkbook
        idx = .ActiveSheet.Index
        cnt = gVim.Count1
        n = .Sheets.Count
        i = idx
        Do
            i = (i Mod n) + 1
            If i = 1 Then
                warpFlag = True
            ElseIf i = idx Then
                warpFlag = False
            End If

            If .Sheets(i).Visible = xlSheetVisible Then
                cnt = cnt - 1
            End If

            If cnt = 0 Then
                If warpFlag Then
                    .Sheets(idx).Move Before:=.Sheets(i)
                Else
                    .Sheets(idx).Move After:=.Sheets(i)
                End If
                Exit Do
            End If
        Loop
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveSheetBack")
End Function

Function MoveSheetBack(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim idx As Integer
    Dim cnt As Integer
    Dim n As Integer
    Dim i As Integer
    Dim warpFlag As Boolean

    With ActiveWorkbook
        idx = .ActiveSheet.Index
        cnt = gVim.Count1
        n = .Sheets.Count
        i = idx
        Do
            i = (i - 1) Mod n
            If i = 0 Then
                i = n
                warpFlag = True
            ElseIf i = idx Then
                warpFlag = False
            End If

            If .Sheets(i).Visible = xlSheetVisible Then
                cnt = cnt - 1
            End If

            If cnt = 0 Then
                If warpFlag Then
                    .Sheets(idx).Move After:=.Sheets(i)
                Else
                    .Sheets(idx).Move Before:=.Sheets(i)
                End If
                Exit Do
            End If
        Loop
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveSheetBack")
End Function

Function InsertWorksheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    With ActiveWorkbook
        .Worksheets.Add Before:=.ActiveSheet
    End With
    Exit Function

Catch:
    Call ErrorHandler("InsertWorksheet")
End Function

Function AppendWorksheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    With ActiveWorkbook
        .Worksheets.Add After:=.ActiveSheet
    End With
    Exit Function

Catch:
    Call ErrorHandler("AppendWorksheet")
End Function

Function DeleteSheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    'error if target sheet is last visible one
    If ActiveSheet.Visible = xlSheetVisible And GetVisibleSheetsCount() = 1 Then
        MsgBox gVim.Msg.DeleteOrHideAllSheets, vbExclamation
        Exit Function
    End If

    ActiveSheet.Delete
    Exit Function

Catch:
    Call ErrorHandler("DeleteSheet")
End Function

Function ActivateSheet(Optional ByVal sheetNum As String) As Boolean
    On Error GoTo Catch

    If Len(sheetNum) = 0 Then
        ActivateSheet = True
        Exit Function
    ElseIf sheetNum Like "*[!0-9]*" Then
        Exit Function
    End If

    Dim idx As Long
    idx = CLng(Right(sheetNum, 10))

    With ActiveWorkbook
        If idx < 1 Then
            idx = 1
        ElseIf .Sheets.Count < idx Then
            idx = .Sheets.Count
        End If

        If .Sheets(idx).Visible <> xlSheetVisible Then
            .Sheets(idx).Visible = xlSheetVisible
        End If

        .Sheets(idx).Select
    End With
    Exit Function

Catch:
    Call ErrorHandler("ActivateSheet")
End Function

Function ActivateFirstSheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Integer

    With ActiveWorkbook
        For i = 1 To .Sheets.Count
            If .Sheets(i).Visible = xlSheetVisible Then
                .Sheets(i).Activate
                Exit Function
            End If
        Next i
    End With
    Exit Function

Catch:
    Call ErrorHandler("ActivateFirstSheet")
End Function

Function ActivateLastSheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Integer

    With ActiveWorkbook
        For i = .Sheets.Count To 1 Step -1
            If .Sheets(i).Visible = xlSheetVisible Then
                .Sheets(i).Activate
                Exit Function
            End If
        Next i
    End With
    Exit Function

Catch:
    Call ErrorHandler("ActivateLastSheet")
End Function

Function CloneSheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    ActiveSheet.Copy After:=ActiveSheet

    Exit Function

Catch:
    Call ErrorHandler("CloneSheet")
End Function

Function ShowSheetPicker(Optional ByVal g As String) As Boolean
    UF_SheetPicker.Show
End Function

Function ChangeSheetTabColor(Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    If ActiveSheet Is Nothing Then
        Exit Function
    End If

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If Not resultColor Is Nothing Then
        With ActiveSheet.Tab
            If resultColor.IsNull Then
                .ColorIndex = xlNone
            ElseIf resultColor.IsThemeColor Then
                .ThemeColor = resultColor.ThemeColor
                .TintAndShade = resultColor.TintAndShade
            Else
                .Color = resultColor.Color
            End If

            Call RepeatRegister("ChangeSheetTabColor", resultColor)
        End With
    End If
    Exit Function

Catch:
    Call ErrorHandler("ChangeSheetTabColor")
End Function

Function ExportSheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    Application.Dialogs(xlDialogWorkbookCopy).Show
    Exit Function

Catch:
    Call ErrorHandler("ExportSheet")
End Function

Function PrintPreviewOfActiveSheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    ActiveSheet.PrintPreview
    Exit Function

Catch:
    Call ErrorHandler("PrintPreviewOfActiveSheet")
End Function
