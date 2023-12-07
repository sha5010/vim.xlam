VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_FontColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private IsThemeColor_ As Boolean
Private ThemeColorIndex As Long
Private ObjectThemeColorIndex As Long
Private AddLuminance_ As Long
Private Color_ As Long
Private IsSetup As Boolean
Private IsNull_ As Boolean

'##################################################################################################'
' Source: https://www.mrexcel.com/board/threads/applying-tintandshade-to-a-userform-label.1189219/
'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv'

#If VBA7 Then
    Private Declare PtrSafe Function ColorAdjustLuma Lib "SHLWAPI.DLL" (ByVal clrRGB As Long, ByVal n As Long, ByVal fScale As Long) As Long
    Private Declare PtrSafe Function TranslateColor Lib "oleAut32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As LongPtr, Col As Long) As Long
#Else
    Private Declare Function ColorAdjustLuma Lib "SHLWAPI.DLL" (ByVal clrRGB As Long, ByVal n As Long, ByVal fScale As Long) As Long
    Private Declare Function TranslateColor Lib "oleAut32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
#End If

Private Function ApplyTintAndShade(ByVal ColRef As Long, Optional ByVal AddLuminance As Long = 0) As Long
    '(AddLuminance must be between -100 and +100)
    Call TranslateColor(ColRef, 0, ColRef)
    ApplyTintAndShade = ColorAdjustLuma(ColRef, AddLuminance * 10, True)
End Function

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
'##################################################################################################'

Private Sub Class_Initialize()
    IsSetup = False
    IsThemeColor_ = False
    ThemeColorIndex = 0
    AddLuminance_ = 0
    Color_ = 0
    IsNull_ = False
End Sub

Public Sub Setup(Optional ByVal msoThemeColorIndex As Long = xlNone, Optional ByVal colorCode As Long = xlNone)
    If msoThemeColorIndex = xlNone And colorCode = xlNone Then
        Err.Raise 50001, Description:="msoThemeColorIndex か colorCode のどちらかを指定しなければなりません。"
        Exit Sub
    End If

    IsSetup = True

    If msoThemeColorIndex <> xlNone Then
        IsThemeColor_ = True
        ThemeColorIndex = msoThemeColorIndex

        Select Case msoThemeColorIndex
            Case 1, 3
                ObjectThemeColorIndex = ThemeColorIndex + 1
            Case 2, 4
                ObjectThemeColorIndex = ThemeColorIndex - 1
            Case Else
                ObjectThemeColorIndex = ThemeColorIndex
        End Select

        On Error GoTo Catch

        Color_ = ActiveWorkbook.Theme.ThemeColorScheme(ObjectThemeColorIndex)
        Exit Sub

Catch:
        Dim savedMode As Long
        Dim savedTintAndShade As Double
        Dim savedColor As Long

        With ActiveSheet.Tab
            If .ThemeColor > 0 Then
                savedColor = .ThemeColor
                savedTintAndShade = .TintAndShade
                savedMode = 0
            ElseIf .ColorIndex = xlNone Then
                savedMode = 1
            Else
                savedColor = .Color
                savedMode = 2
            End If

            .ThemeColor = msoThemeColorIndex
            Color_ = .Color

            If savedMode = 0 Then
                .ThemeColor = savedColor
                .TintAndShade = savedTintAndShade
            ElseIf savedMode = 1 Then
                .ColorIndex = xlNone
            Else
                .Color = savedColor
            End If
        End With

    Else
        Color_ = colorCode
    End If
End Sub

Public Property Get AddLuminance() As Long
    AddLuminance = AddLuminance_
End Property

Public Property Let AddLuminance(ByVal AddLuminance As Long)
    AddLuminance_ = AddLuminance
End Property

Public Property Get Color() As Long
    If AddLuminance <> 0 Then
        Color = ApplyTintAndShade(Color_, AddLuminance)
    Else
        Color = Color_
    End If
End Property

Public Property Get ThemeColor() As Long
    ThemeColor = ThemeColorIndex
End Property

Public Property Get ObjectThemeColor() As Long
    ObjectThemeColor = ObjectThemeColorIndex
End Property

Public Property Get TintAndShade() As Double
    TintAndShade = AddLuminance_ / 100
End Property

Public Property Get IsThemeColor() As Boolean
    IsThemeColor = IsThemeColor_
End Property

Public Property Let IsNull(ByVal IsNull As Boolean)
    IsNull_ = IsNull
End Property

Public Property Get IsNull() As Boolean
    IsNull = IsNull_
End Property

Public Property Get Red() As Byte
    Red = Color_ Mod 256
End Property

Public Property Get Green() As Byte
    Green = Color_ ¥ 256 Mod 256
End Property

Public Property Get Blue() As Byte
    Blue = Color_ ¥ 256 ¥ 256
End Property

Public Property Get Hue() As Long
    Dim maxValue As Long: maxValue = WorksheetFunction.Max(Red, Green, Blue)
    Dim minValue As Long: minValue = WorksheetFunction.Min(Red, Green, Blue)

    If maxValue = minValue Then
        Hue = 0
    ElseIf Red = maxValue Then
        Hue = CLng(60 * ((Green - Blue) / (maxValue - minValue)))
    ElseIf Green = maxValue Then
        Hue = CLng(60 * ((Blue - Red) / (maxValue - minValue))) + 120
    ElseIf Blue = maxValue Then
        Hue = CLng(60 * ((Red - Green) / (maxValue - minValue))) + 240
    End If
End Property

Public Property Get Saturation() As Byte
    Dim maxValue As Long: maxValue = WorksheetFunction.Max(Red, Green, Blue)
    Dim minValue As Long: minValue = WorksheetFunction.Min(Red, Green, Blue)
    Dim cnt As Double: cnt = (maxValue + minValue) / 2

    If cnt <= 127 Then
        Saturation = (maxValue - minValue) / (maxValue + minValue)
    Else
        Saturation = (maxValue - minValue) / (510 - maxValue - minValue)
    End If
End Property

Public Property Get Luminance() As Byte
    Dim maxValue As Long: maxValue = WorksheetFunction.Max(Red, Green, Blue)
    Dim minValue As Long: minValue = WorksheetFunction.Min(Red, Green, Blue)

    Luminance = (maxValue + minValue) / 2
End Property
