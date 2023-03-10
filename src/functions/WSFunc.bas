Attribute VB_Name = "F_WSFunc"
Option Explicit
Option Private Module

Function nextWorksheet()
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
    Call keystroke(True, Ctrl_ + PageDown_)
    If Err.Number <> 0 Then
        Call errorHandler("nextWorksheet")
    End If
End Function

Function previousWorksheet()
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
    Call keystroke(True, Ctrl_ + PageUp_)
    If Err.Number <> 0 Then
        Call errorHandler("previousWorksheet")
    End If
End Function

Function renameWorksheet()
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
            ElseIf isSheetExists(ret) Then
                MsgBox "すでに """ & ret & """ シートが存在します。", vbExclamation
                Exit Function
            End If
            .Worksheets(.ActiveSheet.Index).Name = ret

            Call setStatusBarTemporarily("シート名を変更しました： """ & _
                beforeName & """ → """ & ret & """", 3)
        End If
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("renameWorksheet")
    End If
End Function

Function moveWorksheetForward()
    On Error GoTo Catch

    Dim idx As Integer
    Dim cnt As Integer
    Dim n As Integer
    Dim i As Integer
    Dim warpFlag As Boolean

    With ActiveWorkbook
        idx = .ActiveSheet.Index
        cnt = gCount
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
    If Err.Number <> 0 Then
        Call errorHandler("moveWorksheetBack")
    End If
End Function

Function moveWorksheetBack()
    On Error GoTo Catch

    Dim idx As Integer
    Dim cnt As Integer
    Dim n As Integer
    Dim i As Integer
    Dim warpFlag As Boolean

    With ActiveWorkbook
        idx = .ActiveSheet.Index
        cnt = gCount
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
    If Err.Number <> 0 Then
        Call errorHandler("moveWorksheetBack")
    End If
End Function

Function insertWorksheet()
    On Error GoTo Catch
    With ActiveWorkbook
        .Worksheets.Add Before:=.ActiveSheet
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("insertWorksheet")
    End If
End Function

Function appendWorksheet()
    On Error GoTo Catch
    With ActiveWorkbook
        .Worksheets.Add After:=.ActiveSheet
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("appendWorksheet")
    End If
End Function

Function deleteWorksheet()
    On Error GoTo Catch

    'error if target sheet is last visible one
    If ActiveSheet.Visible = xlSheetVisible And getVisibleSheetsCount() = 1 Then
        MsgBox "シートをすべて削除、または非表示にすることはできません。", vbExclamation
        Exit Function
    End If

    ActiveSheet.Delete
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("deleteWorksheet")
    End If
End Function

Function activateWorksheet(ByVal n As String) As Boolean
    On Error GoTo Catch

    Dim idx As Integer

    If Not IsNumeric(n) Or InStr(n, ".") > 0 Then
        Exit Function
    End If

    idx = CInt(n)

    With ActiveWorkbook
        If idx < 1 Or .Worksheets.Count < idx Then
            Exit Function
        End If

        If .Worksheets(idx).Visible <> xlSheetVisible Then
            .Worksheets(idx).Visible = xlSheetVisible
        End If

        .Worksheets(idx).Select
        activateWorksheet = True
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("activateWorksheet")
    End If
End Function

Function activateFirstWorksheet()
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
    If Err.Number <> 0 Then
        Call errorHandler("activateFirstWorksheet")
    End If
End Function

Function activateLastWorksheet()
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
    If Err.Number <> 0 Then
        Call errorHandler("activateLastWorksheet")
    End If
End Function

Function cloneWorksheet()
    On Error GoTo Catch

    ActiveSheet.Copy After:=ActiveSheet

    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("cloneWorksheet")
    End If
End Function

Function showSheetPicker()
    UF_SheetPicker.Show
End Function

Function changeWorksheetTabColor(Optional ByVal resultColor As cls_FontColor)
    On Error GoTo Catch

    If ActiveSheet Is Nothing Then
        Exit Function
    End If

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.showColorPicker()
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

            Call repeatRegister("changeWorksheetTabColor", resultColor)
        End With
    End If
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("changeWorksheetTabColor")
    End If
End Function

Function exportWorksheet()
    On Error GoTo Catch
    Application.Dialogs(xlDialogWorkbookCopy).Show
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("exportWorksheet")
    End If
End Function

Function printPreviewOfActiveSheet()
    On Error GoTo Catch
    ActiveSheet.PrintPreview
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("printPreviewOfActiveSheet")
    End If
End Function
