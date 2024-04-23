VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ColorPicker 
   Caption         =   "ColorPicker"
   ClientHeight    =   3165
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   3510
   OleObjectBlob   =   "ColorPicker.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const KEY_LIST1 As String = "qwertyuiop"    ' Theme color selection key (section 1)
Private Const KEY_LIST2 As String = "asdfghjkl;"    ' Default color selection key (section 2)
Private Const KEY_LIST3 As String = "zxcvbn"        ' Custom color selection key (section 3)
Private Const KEY_DETAIL As String = "12345"        ' Theme color luminance selection key
Private Const KEY_NULL As String = "n"              ' Auto or Null selection key
Private Const BORDER_COLOR As Long = &HE4E4E4       ' Color box border color
Private Const DISABLED_COLOR As Long = &HDDDDDD     ' Disabled text and background color

Private Const IDX_LIST1 As Long = 2                 ' Theme color row number
Private Const IDX_LIST2 As Long = 18                ' Default color row number
Private Const IDX_LIST3 As Long = 23                ' Custom color row number
Private Const IDX_DETAIL_TOP As Long = 5            ' Theme color luminance top row number
Private Const IDX_RESULT As Long = 27               ' Result and status text row number

Private Const TEXT_PREFIX As String = " "           ' Text prefix when user type some key
Private Const PLACEHOLDER As String = " <SPACE> or #  ->  RGB Color"    ' Placeholder when user do not type any key

Private BOX_GAP As Double

Private cColorTable As Collection
Private cColorObject As Collection
Private cLabelTable As Collection
Private cResultLabel As MSForms.Label
Private cTextLabel As MSForms.Label
Private cFocusLabel As MSForms.Label
Private cCmdBuf As String
Private cResultColor As cls_FontColor

'/*
' * Generates a label name based on the provided coordinates (x, y).
' *
' * @param {Long} x - The x-coordinate.
' * @param {Long} y - The y-coordinate.
' * @returns {String} - The generated label name.
' */
Private Function GetLabelName(ByVal x As Long, ByVal y As Long) As String
    GetLabelName = "Label_    "
    If x < 10 Then
        Mid(GetLabelName, 7) = "0"
        Mid(GetLabelName, 8) = CStr(x)
    Else
        Mid(GetLabelName, 7) = CStr(x)
    End If

    If y < 10 Then
        Mid(GetLabelName, 9) = "0"
        Mid(GetLabelName, 10) = CStr(y)
    Else
        Mid(GetLabelName, 9) = CStr(y)
    End If
End Function
'/*
' * Extracts coordinates (x, y) from the provided label name.
' *
' * @param {String} labelName - The label name to extract coordinates from.
' * @param {Long} x - Output parameter for the x-coordinate.
' * @param {Long} y - Output parameter for the y-coordinate.
' */
Private Sub GetXYFromLabelName(ByVal labelName As String, ByRef x As Long, ByRef y As Long)
    If Not labelName Like "Label_[0-9][0-9][0-9][0-9]" Then
        x = -1
        y = -1
    Else
        x = CLng(Mid(labelName, 7, 2))
        y = CLng(Mid(labelName, 9, 2))
    End If
End Sub

'/*
' * Creates a label and sets its size, position, and font.
' *
' * @param {Long} x - The x-coordinate.
' * @param {Long} y - The y-coordinate.
' * @param {Long} xSize - Optional. The width of the label in cells (default: 1).
' * @param {Long} ySize - Optional. The height of the label in cells (default: 1).
' * @returns {MSForms.Label} - The created label.
' */
Private Function PutLabel(ByVal x As Long, ByVal y As Long, _
                 Optional ByVal xSize As Long = 1, Optional ByVal ySize As Long = 1) As MSForms.Label

    Set PutLabel = Me.Controls.Add("Forms.Label.1", GetLabelName(x, y), True)
    With PutLabel
        ' Size & Place
        .Left = BOX_GAP + x * (gVim.Config.ColorPickerSize + BOX_GAP)
        .Top = BOX_GAP + (gVim.Config.ColorPickerSize / 2) * y
        .Width = gVim.Config.ColorPickerSize * xSize + BOX_GAP * (xSize - 1)
        .Height = gVim.Config.ColorPickerSize * ySize

        ' Font
        .TextAlign = fmTextAlignCenter
        .Font.Name = "Consolas"
        .Font.Size = gVim.Config.ColorPickerSize / 4 * 3
    End With
End Function

'/*
' * Creates a colored label and associates it with a font color.
' *
' * @param {Long} x - The x-coordinate.
' * @param {Long} y - The y-coordinate.
' * @param {cls_FontColor} associatedColor - The font color associated with the label.
' * @param {Long} BorderColor - Optional. The border color (default: xlNone).
' * @param {Long} xSize - Optional. The width of the label in cells (default: 1).
' * @param {Long} ySize - Optional. The height of the label in cells (default: 1).
' * @returns {MSForms.Label} - The created colored label.
' */
Private Function PutColor(ByVal x As Long, ByVal y As Long, ByRef associatedColor As cls_FontColor, _
                 Optional ByVal BorderColor As Long = xlNone, _
                 Optional ByVal xSize As Long = 1, _
                 Optional ByVal ySize As Long = 1) As MSForms.Label

    Set PutColor = PutLabel(x, y, xSize, ySize)
    With PutColor
        .BackColor = associatedColor.Color
        .Tag = .BackColor

        If BorderColor <> xlNone Then
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = BorderColor
        End If
    End With

    cColorTable.Add PutColor, PutColor.Name
    cColorObject.Add associatedColor, PutColor.Name
End Function

'/*
' * Creates a label with the specified caption.
' *
' * @param {Long} x - The x-coordinate.
' * @param {Long} y - The y-coordinate.
' * @param {String} Caption - The caption text for the label.
' * @param {Long} xSize - Optional. The width of the label in cells (default: 1).
' * @param {Long} ySize - Optional. The height of the label in cells (default: 1).
' * @returns {MSForms.Label} - The created label with the specified caption.
' */
Private Function PutText(ByVal x As Long, ByVal y As Long, ByVal Caption As String, _
                Optional ByVal xSize As Long = 1, Optional ySize As Long = 1) As MSForms.Label

    Set PutText = PutLabel(x, y, xSize, ySize)
    PutText.Caption = Caption

    cLabelTable.Add PutText, PutText.Name
End Function

'/*
' * Activates the UserForm and initializes the status of class-level variables.
' */
Private Sub UserForm_Activate()
    ' Clear command buffer and result color
    cCmdBuf = ""
    Set cResultColor = Nothing

    ' Set the position to the center of the active window
    With ActiveWindow
        Me.Top = .Top + (.Height - Me.Height) / 2
        Me.Left = .Left + (.Width - Me.Width) / 2
    End With
End Sub

'/*
' * Initializes the UserForm, setting up color tables, labels, and form size.
' */
Private Sub UserForm_Initialize()
    ' Initialize variables
    Set cColorTable = New Collection
    Set cLabelTable = New Collection
    Set cColorObject = New Collection

    BOX_GAP = gVim.Config.ColorPickerSize / 4

    Dim i As Long
    Dim j As Long
    Dim Color As cls_FontColor
    Dim themeColorLuminance As Long
    Dim cnt As Long: cnt = 1

    Dim defaultColor As Variant
    defaultColor = Array(192, 255, 49407, 65535, 5296274, 5287936, 15773696, 12611584, 6299648, 10498160)

    Dim customColor As Variant
    customColor = Array(gVim.Config.CustomColor1, _
                        gVim.Config.CustomColor2, _
                        gVim.Config.CustomColor3, _
                        gVim.Config.CustomColor4, _
                        gVim.Config.CustomColor5)

    Dim lncBlack     As Variant: lncBlack = Array(0, 50, 35, 25, 15, 5)             ' Basecolor's luminance = 0
    Dim lncDarkGray  As Variant: lncDarkGray = Array(0, 90, 75, 50, 25, 10)         ' Basecolor's luminance = 1 - 50
    Dim lncDefault   As Variant: lncDefault = Array(0, 80, 60, 40, -25, -50)        ' Basecolor's luminance = 51 - 203
    Dim lncLightGray As Variant: lncLightGray = Array(0, -10, -25, -50, -75, -90)   ' Basecolor's luminance = 204 - 254
    Dim lncWhite     As Variant: lncWhite = Array(0, -5, -15, -25, -35, -50)        ' Basecolor's luminance = 255

    ' Loop through theme colors and set up color variations
    For i = 0 To 9
        ' Put theme colors
        Set Color = New cls_FontColor
        Call Color.Setup(msoThemeColorIndex:=i + 1)
        themeColorLuminance = Color.Luminance

        Call PutText(i + 1, IDX_LIST1 - 2, Caption:=Mid(KEY_LIST1, i + 1, 1))
        Call PutColor(i + 1, IDX_LIST1, Color, BorderColor:=BORDER_COLOR)

        ' Put brightness variation of theme color
        For j = 1 To 5
            Set Color = New cls_FontColor
            Call Color.Setup(msoThemeColorIndex:=i + 1)

            Select Case themeColorLuminance
                Case 51 To 203
                    Color.AddLuminance = lncDefault(j)
                Case 1 To 50
                    Color.AddLuminance = lncDarkGray(j)
                Case 204 To 254
                    Color.AddLuminance = lncLightGray(j)
                Case 0
                    Color.AddLuminance = lncBlack(j)
                Case 255
                    Color.AddLuminance = lncWhite(j)
            End Select

            Call PutColor(i + 1, IDX_DETAIL_TOP + (j - 1) * 2, Color)
        Next j

        ' Border label
        With PutLabel(i + 1, IDX_DETAIL_TOP, ySize:=5)
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = BORDER_COLOR
            .BackStyle = fmBackStyleTransparent
        End With

        ' Put default color
        Set Color = New cls_FontColor
        Call Color.Setup(colorCode:=defaultColor(i))

        Call PutText(i + 1, IDX_LIST2 - 2, Caption:=Mid(KEY_LIST2, i + 1, 1))
        Call PutColor(i + 1, IDX_LIST2, Color, BorderColor:=BORDER_COLOR)
    Next i

    ' Labels on both sides
    For j = 1 To 5
        Call ChangeState(PutText(0, IDX_DETAIL_TOP + (j - 1) * 2, Caption:=Mid(KEY_DETAIL, j, 1)), False)
        Call ChangeState(PutText(11, IDX_DETAIL_TOP + (j - 1) * 2, Caption:=Mid(KEY_DETAIL, j, 1)), False)
    Next j

    ' Custom colors
    For i = 1 To 5
        Set Color = New cls_FontColor
        Call Color.Setup(colorCode:=customColor(i - 1))

        Call PutText(i, IDX_LIST3 - 2, Caption:=Mid(KEY_LIST3, i, 1))
        Call PutColor(i, IDX_LIST3, Color, BorderColor:=BORDER_COLOR)
    Next i

    ' Auto or None
    Set Color = New cls_FontColor
    Call Color.Setup(colorCode:=0)  ' dummy
    Color.IsNull = True

    Call PutText(6, IDX_LIST3 - 2, Caption:=KEY_NULL, xSize:=5)
    With PutColor(6, IDX_LIST3, Color, xSize:=5)
        .Caption = "Auto, Null"
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = &HA0A0A0
    End With

    ' Result label
    Set cResultLabel = PutLabel(0, IDX_RESULT)
    With cResultLabel
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = BORDER_COLOR
        .BackStyle = fmBackStyleOpaque
        .BackColor = Me.BackColor
    End With

    ' Text label
    Set cTextLabel = PutLabel(1, IDX_RESULT, xSize:=11)
    With cTextLabel
        .Caption = PLACEHOLDER
        .TextAlign = fmTextAlignLeft
    End With

    ' Focus label
    Set cFocusLabel = PutLabel(0, 0)
    With cFocusLabel
        .BorderColor = &H1048EF
        .BorderStyle = fmBorderStyleSingle
        .BackStyle = fmBackStyleTransparent
        .Visible = False
    End With

    ' Calculate form margin
    Dim marginWidth  As Double: marginWidth = Me.Width - Me.InsideWidth
    Dim marginHeight As Double: marginHeight = Me.Height - Me.InsideHeight

    ' Set form size
    Me.Width = marginWidth + gVim.Config.ColorPickerSize * 12 + BOX_GAP * 13
    Me.Height = marginHeight + gVim.Config.ColorPickerSize * (IDX_RESULT / 2 + 1) + BOX_GAP * 2

    ' Flickering prevention
    Me.DrawBuffer = WorksheetFunction.Min(CLng(Me.InsideHeight * Me.InsideWidth / 9 * 16), 1048576)
End Sub


'/*
' * Handles key presses in the user form.
' *
' * @param {MSForms.ReturnInteger} KeyAscii - The ASCII value of the pressed key.
' */
Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim addChar As String

    ' Reset input if Escape key is pressed
    If KeyAscii = Escape_ Then
        If cCmdBuf <> "" Then
            cCmdBuf = ""
        Else
            ' Close form
            Call Quit(False)
            Exit Sub
        End If

    ' Delete one character if Backspace key is pressed
    ElseIf KeyAscii = BackSpace_ Then
        If cCmdBuf <> "" Then
            cCmdBuf = Left(cCmdBuf, Len(cCmdBuf) - 1)
        End If

    ' Close form if Enter key is pressed and result color is selected
    ElseIf KeyAscii = Enter_ Then
        If Not cResultColor Is Nothing Then
            Call Quit(True)
        End If
        Exit Sub

    ' Handle other key presses
    ElseIf KeyAscii > 31 Then
        addChar = LCase(Chr(KeyAscii))

        ' Process key based on current input state
        If Len(cCmdBuf) = 0 Then
            If addChar = " " Or addChar = "#" Or InStr(KEY_LIST1 & KEY_LIST2 & KEY_LIST3 & KEY_NULL, addChar) > 0 Then
                If addChar = " " Then
                    cCmdBuf = cCmdBuf & "#"
                Else
                    cCmdBuf = cCmdBuf & addChar
                End If
            End If
        ElseIf Left(cCmdBuf, 1) = "#" And InStr("0123456789abcdef", addChar) > 0 And Len(cCmdBuf) < 7 Then
            cCmdBuf = cCmdBuf & addChar
        ElseIf Len(cCmdBuf) = 1 And InStr(KEY_DETAIL, addChar) > 0 Then
            cCmdBuf = cCmdBuf & addChar
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    ' Check and process the command
    Call CheckCommand
End Sub

'/*
' * Updates the focus and enables/disables labels based on the selected color coordinates.
' *
' * @param {Long} [x=-1] - The x-coordinate of the selected color.
' * @param {Long} [y=-1] - The y-coordinate of the selected color.
' */
Private Sub UpdateFocus(Optional ByVal x As Long = -1, Optional ByVal y As Long = -1)
    Dim labelObj As MSForms.Label
    Dim labelX As Long
    Dim labelY As Long
    Dim labelEnabled As Boolean

    ' Iterate through color labels
    For Each labelObj In cColorTable
        If x < 0 Then
            ' No color selected, enable all labels
            Call ChangeState(labelObj, True)
        Else
            ' Enable labels based on selected coordinates
            Call GetXYFromLabelName(labelObj.Name, labelX, labelY)
            labelEnabled = (y < 0 And labelY < IDX_LIST2 And labelX = x) Or (labelX = x And labelY = y)
            Call ChangeState(labelObj, labelEnabled)
        End If
    Next

    ' Iterate through text labels
    For Each labelObj In cLabelTable
        Call GetXYFromLabelName(labelObj.Name, labelX, labelY)

        If labelX = 0 Or labelX = 11 Then
            ' Side text label
            If x < 0 Then
                Call ChangeState(labelObj, False)
            Else
                Call ChangeState(labelObj, (y < 0 Or y = labelY))
            End If
        Else
            ' Other text label
            If x < 0 Then
                Call ChangeState(labelObj, True)
            ElseIf y < 0 Or y < IDX_LIST2 Then
                Call ChangeState(labelObj, (x = labelX And labelY < IDX_LIST2 - 2))
            ElseIf y < IDX_LIST3 Then
                Call ChangeState(labelObj, (x = labelX And labelY = IDX_LIST2 - 2))
            ElseIf y < IDX_RESULT Then
                Call ChangeState(labelObj, (x = labelX And labelY = IDX_LIST3 - 2))
            End If
        End If
    Next
End Sub

'/*
' * Checks the entered command and updates the focus accordingly.
' */
Private Sub CheckCommand()
    Dim x As Long
    Dim y As Long
    Dim colorY As Long

    ' Check if command buffer is empty
    If cCmdBuf = "" Then
        ' No command, update focus and reset labels
        Call UpdateFocus
        cTextLabel.Caption = PLACEHOLDER
        cFocusLabel.Visible = False
        Set cResultColor = Nothing

    ElseIf InStr(cCmdBuf, "#") = 1 Then
        ' Handle '#' command
        If cCmdBuf = "#" Then
            Call UpdateFocus(0, 0)
        End If

        Dim colorCode As String
        Dim colorValue As Long

        ' Extract color code
        colorCode = Mid(cCmdBuf, 2)
        colorValue = HexColorCodeToLong(colorCode)

        ' Validate color value
        If colorValue < 0 Then
            Set cResultColor = Nothing
        Else
            ' Create a new color object
            Set cResultColor = New cls_FontColor
            Call cResultColor.Setup(colorCode:=colorValue)
        End If

        ' Update labels
        cTextLabel.Caption = TEXT_PREFIX & cCmdBuf
        cFocusLabel.Visible = False

    ElseIf Len(cCmdBuf) > 0 Then
        ' Handle other commands
        x = InStr(KEY_LIST2 & KEY_LIST3, cCmdBuf)
        If x > 0 Then
            ' Commands in KEY_LIST2 or KEY_LIST3
            If x > 10 Then
                y = IDX_LIST3
            Else
                y = IDX_LIST2
            End If
            x = (x - 1) Mod 10 + 1
            colorY = y

        Else
            ' Commands in KEY_LIST1
            x = InStr(KEY_LIST1, Left(cCmdBuf, 1))
            If Len(cCmdBuf) = 1 Then
                ' Single character command
                y = -1
                colorY = IDX_LIST1
                Set cResultColor = cColorObject(GetLabelName(x, IDX_LIST1))
            Else
                ' Detailed commands
                y = IDX_LIST1 + 1 + InStr(KEY_DETAIL, Mid(cCmdBuf, 2, 1)) * 2
                colorY = y
            End If
        End If

        ' Update focus and labels based on the command
        Call UpdateFocus(x, y)

        Set cResultColor = cColorObject(GetLabelName(x, colorY))
        With cColorTable(GetLabelName(x, colorY))
            cFocusLabel.Left = .Left
            cFocusLabel.Top = .Top
            cFocusLabel.Width = .Width
            cFocusLabel.Height = .Height
            cFocusLabel.Visible = True
        End With

        ' Update text label with the command and result
        If Not cResultColor.IsNull Then
            cTextLabel.Caption = Left(TEXT_PREFIX & cCmdBuf & "     ", 6) & "#" & ColorCodeToHex(cResultColor.Color)
        Else
            cTextLabel.Caption = TEXT_PREFIX & cCmdBuf
        End If
    End If

    ' Update the result label's appearance based on the result color
    If cResultColor Is Nothing Then
        cResultLabel.BackColor = Me.BackColor
        cResultLabel.BorderStyle = fmBorderStyleNone
    ElseIf cResultColor.IsNull Then
        cResultLabel.BackColor = Me.BackColor
        cResultLabel.BorderStyle = fmBorderStyleNone
    Else
        cResultLabel.BackColor = cResultColor.Color
    End If
End Sub

'/*
' * Exits the user form, optionally returning the result.
' *
' * @param {Boolean} returnResult - Flag indicating whether to return the result.
' */
Private Sub Quit(ByVal returnResult As Boolean)
    If Not returnResult Then
        Set cResultColor = Nothing
    End If

    ' Reset labels and hide the form
    cTextLabel.Caption = PLACEHOLDER
    Me.Hide
End Sub

'/*
' * Changes the state of the target label based on the Enabled parameter.
' *
' * @param {MSForms.Label} targetLabel - The label to change the state.
' * @param {Boolean} Enabled - Flag indicating whether the label should be enabled.
' */
Private Sub ChangeState(ByRef targetLabel As MSForms.Label, ByVal Enabled As Boolean)
    With targetLabel
        ' Update label appearance based on the Enabled parameter
        If Enabled Then
            If .Tag <> "" Then   ' Only color label has the tag
                If Not .BackColor = .Tag Then .BackColor = .Tag
            End If
            If Not .ForeColor = vbBlack Then .ForeColor = vbBlack
        Else
            If .Tag <> "" Then
                If Not .BackColor = DISABLED_COLOR Then .BackColor = DISABLED_COLOR
            End If
            If Not .ForeColor = DISABLED_COLOR Then .ForeColor = DISABLED_COLOR
        End If
    End With
End Sub

Public Function Launch() As cls_FontColor
    Dim colorCode As String
    Dim colorValue As Long

    UF_Cmd.Hide
    Me.Show
    Set Launch = cResultColor
    Unload Me
End Function
