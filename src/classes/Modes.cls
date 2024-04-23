VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Modes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent As cls_Vim

Private cCurrentMode As String
Private cKeyMapFallback As Dictionary
Private cKeyMapBits As Dictionary

' Event hooks for each modes
Private cModeNormal As cls_ModeNormal
Attribute cModeNormal.VB_VarHelpID = -1
Private cModeVisual As cls_ModeVisual
Private cModeShapeInsert As cls_ModeShapeInsert

'/*
' * Initializes the class and sets up the key map fallback.
' */
Private Sub Class_Initialize()
    Set cKeyMapFallback = New Dictionary
    With cKeyMapFallback
        .Add MODE_DUMMY, ""
        .Add MODE_NORMAL, MODE_DUMMY
        .Add MODE_VISUAL, MODE_NORMAL
        .Add MODE_CMDLINE, ""
        .Add MODE_SHAPEINSERT, ""
    End With

    Set cKeyMapBits = New Dictionary

    Dim i As Long: i = 0
    Dim mode_ As Variant
    For Each mode_ In cKeyMapFallback.Keys()
        cKeyMapBits.Add mode_, 2 ^ i
        i = i + 1
    Next
End Sub

Private Sub Class_Terminate()
    Call ClearAppInstances
    Set Parent = Nothing
End Sub


'/*
' * Gets the current mode.
' *
' * @returns {String} The current mode.
' */
Public Property Get Current() As String
    Current = cCurrentMode
End Property

'/*
' * Gets an array of available modes.
' *
' * @returns {Variant} An array containing the available modes.
' */
Public Property Get AvailableModes() As Variant
    AvailableModes = cKeyMapFallback.Keys()
End Property

'/*
' * Gets the fallback mode for the specified mode.
' *
' * @param {String} [mode_] - The mode to get the fallback for. If not provided, uses the current mode.
' * @returns {String} The fallback mode.
' */
Public Function KeyMapFallback(Optional mode_ As String = "") As String
    If mode_ = "" Then
        mode_ = cCurrentMode
    End If
    KeyMapFallback = cKeyMapFallback(mode_)
End Function

'/*
' * Checks if the specified mode exists in the key map fallback.
' *
' * @param {String} mode_ - The mode to check.
' * @returns {Boolean} True if the mode exists, False otherwise.
' */
Public Function Exists(ByVal mode_ As String) As Boolean
    Exists = cKeyMapFallback.Exists(mode_)
End Function

'/*
' * Clears the event hook instances for different modes.
' */
Private Sub ClearAppInstances()
    Set cModeNormal = Nothing
    Set cModeVisual = Nothing
    Set cModeShapeInsert = Nothing
End Sub

'/*
' * Changes the current mode to the specified mode.
' *
' * @param {String} newMode - The new mode to switch to.
' */
Public Sub Change(ByVal newMode As String)
    ' Check if the new mode is supported
    If Not cKeyMapFallback.Exists(newMode) Then
        Err.Raise 50000, Description:="Unsupported mode: " & newMode

    ' Check if the new mode is current mode
    ElseIf newMode = cCurrentMode Then
        Exit Sub
    End If

    ' Do nothing if change from/to cmdline mode
    If newMode = MODE_CMDLINE Or cCurrentMode = MODE_CMDLINE Then
        cCurrentMode = newMode
        Exit Sub
    End If

    ' Clear existing event hook instances
    Call ClearAppInstances

    ' Set up event hooks based on the new mode
    Select Case newMode
        Case MODE_NORMAL
            Set cModeNormal = New cls_ModeNormal
            Call DisableIME
        Case MODE_VISUAL
            Set cModeVisual = New cls_ModeVisual
        Case MODE_SHAPEINSERT
            Set cModeShapeInsert = New cls_ModeShapeInsert
    End Select

    ' Update the current mode
    cCurrentMode = newMode

    ' Update the keymap if the parent is available
    If Not Parent Is Nothing Then
        Call Parent.KeyMap.BindAll
    End If
End Sub

Public Sub Disable()
    Call ClearAppInstances
    cCurrentMode = MODE_DUMMY
End Sub

Public Function ModeBit(ByVal mode_ As String, _
               Optional ByVal isFallback As Boolean = False) As Long

    If isFallback And mode_ = "" Then
        ModeBit = 0
        Exit Function
    End If

    If Not cKeyMapBits.Exists(mode_) Then
        Err.Raise 50000, Description:="Unsupported mode: " & mode_
    End If

    If isFallback Then
        ModeBit = cKeyMapBits(mode_) Or ModeBit(cKeyMapFallback(mode_), isFallback:=True)
    Else
        ModeBit = cKeyMapBits(mode_)
    End If
End Function

Public Function CheckBits(ByVal sourceValue As Long, ByVal targetBits As Long) As Boolean
    CheckBits = (sourceValue And targetBits) > 0
End Function

Public Property Get Normal() As cls_ModeNormal
    Set Normal = cModeNormal
End Property

Public Property Get Visual() As cls_ModeVisual
    Set Visual = cModeVisual
End Property

Public Property Get ShapeInsert() As cls_ModeShapeInsert
    Set ShapeInsert = cModeShapeInsert
End Property
