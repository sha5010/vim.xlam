Attribute VB_Name = "F_WSFunc"
Option Explicit
Option Private Module

Function nextWorksheet()
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
End Function

Function previousWorksheet()
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
End Function

Function renameWorksheet()
    Dim ret As String
    Dim beforeName As String

    With ActiveWorkbook
        beforeName = .ActiveSheet.Name
        ret = InputBox("新しいシート名を入力してください。", "シート名の変更", beforeName)

        If ret <> "" Then
            On Error GoTo Catch
            .Worksheets(.ActiveSheet.Index).Name = ret

            Call setStatusBarTemporarily("シート名を変更しました： """ & _
                beforeName & """ → """ & ret & """", 3)
        End If
    End With
    Exit Function

Catch:
    Call debugPrint("Cannot rename worksheet. ErrNo: " & Err.Number & "  Description: " & Err.Description, "renameWorksheet")
End Function

Function moveWorksheetForward()
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
End Function

Function moveWorksheetBack()
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
End Function

Function insertWorksheet()
    With ActiveWorkbook
        .Worksheets.Add Before:=.ActiveSheet
    End With
End Function

Function appendWorksheet()
    With ActiveWorkbook
        .Worksheets.Add After:=.ActiveSheet
    End With
End Function

Function deleteWorksheet()
    ActiveSheet.Delete
End Function

Function activateWorksheet(ByVal n As String) As Boolean
    Dim idx As Integer

    On Error GoTo Catch

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
    Call debugPrint("Cannot activate worksheet. ErrNo: " & Err.Number & "  Description: " & Err.Description, "activateWorksheet")
End Function

Function activateFirstWorksheet()
    Dim i As Integer

    On Error GoTo Catch

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
    Call debugPrint("Cannot activate worksheet. ErrNo: " & Err.Number & "  Description: " & Err.Description, "activateFirstWorksheet")
End Function

Function activateLastWorksheet()
    Dim i As Integer

    On Error GoTo Catch

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
    Call debugPrint("Cannot activate worksheet. ErrNo: " & Err.Number & "  Description: " & Err.Description, "activateLastWorksheet")
End Function

Function cloneWorksheet()
    On Error GoTo Catch

    ActiveSheet.Copy After:=ActiveSheet

    Exit Function

Catch:
    Call debugPrint("Cannot clone worksheet. ErrNo: " & Err.Number & "  Description: " & Err.Description, "cloneWorksheet")
End Function

Function showSheetPicker()
    UF_SheetPicker.Show
End Function

Function changeWorksheetTabColor(Optional ByVal resultColor As cls_FontColor)
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
End Function

Function exportWorksheet()
    Application.Dialogs(xlDialogWorkbookCopy).Show
End Function

Function printPreviewOfActiveSheet()
    ActiveSheet.PrintPreview
End Function
