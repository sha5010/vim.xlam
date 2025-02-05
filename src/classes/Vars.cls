VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Vars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent As cls_Vim

Public LastYanked As Range      ' Hold target cells with Copy, Cut command
Public ExtendRange As Range     ' Hold multi selected cells

Public FromInsertCmd As Boolean ' Whether started by insert mode

' Selection lock feature
Private cLockedRowsBegin    As Long
Private cLockedRowsEnd      As Long
Private cLockedColumnsBegin As Long
Private cLockedColumnsend   As Long
Public LockedSheet          As Worksheet

Public Property Get LockedRowsBegin() As Long
    If cLockedRowsBegin > cLockedRowsEnd Then
        Dim tmp As Long
        tmp = cLockedRowsBegin
        cLockedRowsBegin = cLockedRowsEnd
        cLockedRowsEnd = tmp
    End If

    If cLockedRowsBegin <= 0 Then
        cLockedRowsBegin = 0
    ElseIf cLockedRowsBegin > ActiveSheet.Rows.Count Then
        cLockedColumnsBegin = ActiveSheet.Rows.Count
    End If
    LockedRowsBegin = cLockedRowsBegin
End Property
Public Property Get LockedRowsEnd() As Long
    If cLockedRowsBegin > cLockedRowsEnd Then
        Dim tmp As Long
        tmp = cLockedRowsBegin
        cLockedRowsBegin = cLockedRowsEnd
        cLockedRowsEnd = tmp
    End If

    If cLockedRowsEnd <= 0 Then
        cLockedRowsEnd = 0
    ElseIf cLockedRowsEnd > ActiveSheet.Rows.Count Then
        cLockedColumnsend = ActiveSheet.Rows.Count
    End If
    LockedRowsEnd = cLockedRowsEnd
End Property
Public Sub SetLockedRows(ByVal startRow As Long, ByVal endRow As Long)
    If startRow > endRow Then
        Dim tmp As Long
        tmp = startRow
        startRow = endRow
        endRow = tmp
    End If

    If startRow <= 0 Then
        cLockedRowsBegin = 0
        cLockedRowsEnd = 0
    ElseIf endRow > ActiveSheet.Rows.Count Then
        cLockedRowsBegin = 0
        cLockedRowsEnd = 0
    Else
        cLockedRowsBegin = startRow
        cLockedRowsEnd = endRow
    End If
End Sub

Public Property Get LockedColumnsBegin() As Long
    If cLockedColumnsBegin > cLockedColumnsend Then
        Dim tmp As Long
        tmp = cLockedColumnsBegin
        cLockedColumnsBegin = cLockedColumnsend
        cLockedColumnsend = tmp
    End If

    If cLockedColumnsBegin <= 0 Then
        cLockedColumnsBegin = 0
    ElseIf cLockedColumnsBegin > ActiveSheet.Columns.Count Then
        cLockedColumnsBegin = ActiveSheet.Columns.Count
    End If
    LockedColumnsBegin = cLockedColumnsBegin
End Property
Public Property Get LockedColumnsEnd() As Long
    If cLockedColumnsBegin > cLockedColumnsend Then
        Dim tmp As Long
        tmp = cLockedColumnsBegin
        cLockedColumnsBegin = cLockedColumnsend
        cLockedColumnsend = tmp
    End If

    If cLockedColumnsend <= 0 Then
        cLockedColumnsend = 0
    ElseIf cLockedColumnsend > ActiveSheet.Columns.Count Then
        cLockedColumnsend = ActiveSheet.Columns.Count
    End If
    LockedColumnsEnd = cLockedColumnsend
End Property
Public Sub SetLockedColumns(ByVal startColumn As Long, ByVal endColumn As Long)
    If startColumn > endColumn Then
        Dim tmp As Long
        tmp = startColumn
        startColumn = endColumn
        endColumn = tmp
    End If

    If startColumn <= 0 Then
        cLockedColumnsBegin = 0
        cLockedColumnsend = 0
    ElseIf endColumn > ActiveSheet.Columns.Count Then
        cLockedColumnsBegin = 0
        cLockedColumnsend = 0
    Else
        cLockedColumnsBegin = startColumn
        cLockedColumnsend = endColumn
    End If
End Sub

Public Function GetLockedRange() As String
    If LockedSheet Is Nothing Or (cLockedColumnsBegin = 0 And cLockedRowsBegin = 0) Then
        GetLockedRange = ""
        Exit Function
    End If

    With LockedSheet
        Dim isAbs As Boolean
        isAbs = (Application.ReferenceStyle = xlR1C1)
        If cLockedColumnsBegin = 0 Then
            GetLockedRange = .Range(.Rows(LockedRowsBegin), .Rows(LockedRowsEnd)).Address(isAbs, isAbs, Application.ReferenceStyle)
        ElseIf cLockedRowsBegin = 0 Then
            GetLockedRange = .Range(.Columns(LockedColumnsBegin), .Columns(LockedColumnsEnd)).Address(isAbs, isAbs, Application.ReferenceStyle)
        Else
            GetLockedRange = .Range(.Cells(LockedRowsBegin, LockedColumnsBegin), .Cells(LockedRowsEnd, LockedColumnsEnd)).Address(isAbs, isAbs, Application.ReferenceStyle)
        End If
    End With
End Function
