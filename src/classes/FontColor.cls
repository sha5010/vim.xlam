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

Private IsThemeColor_ As Boolean
Private ThemeColorIndex As Integer
Private ObjectThemeColorIndex As Integer
Private Luminance_ As Integer
Private Color_ As Long
Private IsSetup As Boolean
Private IsNull_ As Boolean

Private Function ApplyTintAndShade(ByVal ColRef As Long, Optional ByVal Luminance As Long = 0) As Long
    '(Luminance must be between -100 and +100)
    Call TranslateColor(ColRef, 0, ColRef)
    ApplyTintAndShade = ColorAdjustLuma(ColRef, Luminance * 10, True)
End Function

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
'##################################################################################################'

Private Sub Class_Initialize()
    IsSetup = False
    IsThemeColor_ = False
    ThemeColorIndex = 0
    Luminance_ = 0
    Color_ = 0
    IsNull_ = False
End Sub

Public Sub Setup(Optional ByVal msoThemeColorIndex As Integer = xlNone, Optional ByVal colorCode As Long = xlNone)
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
        Dim savedMode As Integer
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

Private Sub CheckSetup()
    If Not IsSetup Then
        Err.Raise 50000, Description:="セットアップされていません。"
        End
    End If
End Sub

Public Property Get Luminance() As Integer
    Call CheckSetup
    Luminance = Luminance_
End Property

Public Property Let Luminance(ByVal Luminance As Integer)
    Call CheckSetup
    Luminance_ = Luminance
End Property

Public Property Get Color() As Long
    Call CheckSetup
    Color = ApplyTintAndShade(Color_, Luminance)
End Property

Public Property Get ThemeColor() As Integer
    ThemeColor = ThemeColorIndex
End Property

Public Property Get ObjectThemeColor() As Integer
    ObjectThemeColor = ObjectThemeColorIndex
End Property

Public Property Get TintAndShade() As Double
    TintAndShade = Luminance_ / 100
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
