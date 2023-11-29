Attribute VB_Name = "F_WSFunc"
Option Explicit
Option Private Module

Function NextWorksheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Integer

    With ActiveWorkbook
        i = .ActiveSheet.Index - 1
        Do
            i = (i + 1) Mod .Worksheets.Count
            If .Worksheets(i + 1).Visible = xlSheetVisible Then
                .Worksheets(i + 1).Activate
                Exit Function
            End If
        Loop
    End With
    Exit Function

Catch:
    Call KeyStroke(True, Ctrl_ + PageDown_)
    Call ErrorHandler("NextWorksheet")
End Function

Function PreviousWorksheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Integer

    With ActiveWorkbook
        i = .ActiveSheet.Index - 1
        Do
            i = (i - 1 + .Worksheets.Count) Mod .Worksheets.Count
            If .Worksheets(i + 1).Visible = xlSheetVisible Then
                .Worksheets(i + 1).Activate
                Exit Function
            End If
        Loop
    End With
    Exit Function

Catch:
    Call KeyStroke(True, Ctrl_ + PageUp_)
    Call ErrorHandler("PreviousWorksheet")
End Function

Function RenameWorksheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim ret As String
    Dim beforeName As String

    With ActiveWorkbook
        beforeName = .ActiveSheet.Name
        ret = InputBox("新しいシート名を入力してください。", "シート名の変更", beforeName)

        If ret <> "" Then
            'Exit if same name
            If ret = beforeName Then
                Exit Function

            'Error when new sheet name already exists
            ElseIf IsSheetExists(ret) Then
                MsgBox "すでに """ & ret & """ シートが存在します。", vbExclamation
                Exit Function
            End If
            .Worksheets(.ActiveSheet.Index).Name = ret

            Call SetStatusBarTemporarily("シート名を変更しました： """ & _
                beforeName & """ → """ & ret & """", 3000)
        End If
    End With
    Exit Function

Catch:
    Call ErrorHandler("RenameWorksheet")
End Function

Function MoveWorksheetForward(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim idx As Integer
    Dim cnt As Integer
    Dim n As Integer
    Dim i As Integer
    Dim warpFlag As Boolean

    With ActiveWorkbook
        idx = .ActiveSheet.Index
        cnt = gVim.Count1
        n = .Worksheets.Count
        i = idx
        Do
            i = (i Mod n) + 1
            If i = 1 Then
                warpFlag = True
            ElseIf i = idx Then
                warpFlag = False
            End If

            If .Worksheets(i).Visible = xlSheetVisible Then
                cnt = cnt - 1
            End If

            If cnt = 0 Then
                If warpFlag Then
                    .Worksheets(idx).Move Before:=.Worksheets(i)
                Else
                    .Worksheets(idx).Move After:=.Worksheets(i)
                End If
                Exit Do
            End If
        Loop
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveWorksheetBack")
End Function

Function MoveWorksheetBack(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim idx As Integer
    Dim cnt As Integer
    Dim n As Integer
    Dim i As Integer
    Dim warpFlag As Boolean

    With ActiveWorkbook
        idx = .ActiveSheet.Index
        cnt = gVim.Count1
        n = .Worksheets.Count
        i = idx
        Do
            i = (i - 1) Mod n
            If i = 0 Then
                i = n
                warpFlag = True
            ElseIf i = idx Then
                warpFlag = False
            End If

            If .Worksheets(i).Visible = xlSheetVisible Then
                cnt = cnt - 1
            End If

            If cnt = 0 Then
                If warpFlag Then
                    .Worksheets(idx).Move After:=.Worksheets(i)
                Else
                    .Worksheets(idx).Move Before:=.Worksheets(i)
                End If
                Exit Do
            End If
        Loop
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveWorksheetBack")
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

Function DeleteWorksheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    'error if target sheet is last visible one
    If ActiveSheet.Visible = xlSheetVisible And GetVisibleSheetsCount() = 1 Then
        MsgBox "シートをすべて削除、または非表示にすることはできません。", vbExclamation
        Exit Function
    End If

    ActiveSheet.Delete
    Exit Function

Catch:
    Call ErrorHandler("DeleteWorksheet")
End Function

Function ActivateWorksheet(Optional ByVal sheetNum As String) As Boolean
    On Error GoTo Catch

    If Len(sheetNum) = 0 Then
        ActivateWorksheet = True
        Exit Function
    ElseIf sheetNum Like "*[!0-9]*" Then
        Exit Function
    End If

    Dim idx As Long
    idx = CLng(Right(sheetNum, 10))

    With ActiveWorkbook
        If idx < 1 Then
            idx = 1
        ElseIf .Worksheets.Count < idx Then
            idx = .Worksheets.Count
        End If

        If .Worksheets(idx).Visible <> xlSheetVisible Then
            .Worksheets(idx).Visible = xlSheetVisible
        End If

        .Worksheets(idx).Select
    End With
    Exit Function

Catch:
    Call ErrorHandler("ActivateWorksheet")
End Function

Function ActivateFirstWorksheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Integer

    With ActiveWorkbook
        For i = 1 To .Worksheets.Count
            If .Worksheets(i).Visible = xlSheetVisible Then
                .Worksheets(i).Activate
                Exit Function
            End If
        Next i
    End With
    Exit Function

Catch:
    Call ErrorHandler("ActivateFirstWorksheet")
End Function

Function ActivateLastWorksheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Integer

    With ActiveWorkbook
        For i = .Worksheets.Count To 1 Step -1
            If .Worksheets(i).Visible = xlSheetVisible Then
                .Worksheets(i).Activate
                Exit Function
            End If
        Next i
    End With
    Exit Function

Catch:
    Call ErrorHandler("ActivateLastWorksheet")
End Function

Function CloneWorksheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    ActiveSheet.Copy After:=ActiveSheet

    Exit Function

Catch:
    Call ErrorHandler("CloneWorksheet")
End Function

Function ShowSheetPicker(Optional ByVal g As String) As Boolean
    UF_SheetPicker.Show
End Function

Function ChangeWorksheetTabColor(Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    If ActiveSheet Is Nothing Then
        Exit Function
    End If

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.ShowColorPicker()
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

            Call RepeatRegister("ChangeWorksheetTabColor", resultColor)
        End With
    End If
    Exit Function

Catch:
    Call ErrorHandler("ChangeWorksheetTabColor")
End Function

Function ExportWorksheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    Application.Dialogs(xlDialogWorkbookCopy).Show
    Exit Function

Catch:
    Call ErrorHandler("ExportWorksheet")
End Function

Function PrintPreviewOfActiveSheet(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    ActiveSheet.PrintPreview
    Exit Function

Catch:
    Call ErrorHandler("PrintPreviewOfActiveSheet")
End Function
