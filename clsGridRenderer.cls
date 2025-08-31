VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGridRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsGridRenderer.cls — View (Excel renderer)
' Purpose: Paints a Byte grid onto a worksheet range. Initializes cell sizing,
'          clears the target area, and performs diff-based rendering to minimize
'          work per frame (no per-cell screen refresh).


Option Explicit

Implements IRenderer

Private mWS As Worksheet
Private mrngTopLeft As Range
Private mlngHeight As Long, mlngWidth As Long
Private mlngAlive As Long, mlngDead As Long

' keep last drawn state to repaint only diffs
Private mAbytLast() As Byte
Private mblnHasLast As Boolean


Public Sub Init(ByVal ws As Worksheet, ByVal topLeft As Range, _
                ByVal height As Long, ByVal width As Long)
    Call IRenderer_Init(ws, topLeft, height, width)
End Sub


Public Sub Clear()
    Call IRenderer_Clear
End Sub


Public Sub Render(ByVal grid As Variant)
    Call IRenderer_Render(grid)
End Sub


Private Sub IRenderer_Init(ByVal ws As Worksheet, ByVal topLeft As Range, _
                ByVal height As Long, ByVal width As Long)
    Set mWS = ws
    Set mrngTopLeft = topLeft
    mlngHeight = height: mlngWidth = width
    mlngAlive = vbBlack: mlngDead = vbWhite
    mblnHasLast = False

    With mWS.Range(mrngTopLeft, mrngTopLeft.Offset(mlngHeight - 1, mlngWidth - 1))
        .ColumnWidth = 2.5
        .RowHeight = 15
        .Interior.Color = mlngDead
        .Borders.LineStyle = xlLineStyleNone
    End With
End Sub

Private Sub IRenderer_Clear()
    mWS.Range(mrngTopLeft, mrngTopLeft.Offset(mlngHeight - 1, mlngWidth - 1)).Interior.Color = mlngDead
    mblnHasLast = False
End Sub

Private Sub IRenderer_Render(ByVal grid As Variant)
    ' NOTE: Do NOT enable ScreenUpdating here—controller will pulse it once per frame.
    Dim prevSU As Boolean: prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim r As Long, c As Long
    Dim base As Range: Set base = mrngTopLeft

    If mblnHasLast Then
        For r = 1 To mlngHeight
            DoEvents
            For c = 1 To mlngWidth
                If grid(r, c) <> mAbytLast(r, c) Then
                    base.Offset(r - 1, c - 1).Interior.Color = IIf(grid(r, c) = 1, mlngAlive, mlngDead)
                End If
            Next c
        Next r
    Else
        For r = 1 To mlngHeight
            DoEvents
            For c = 1 To mlngWidth
                base.Offset(r - 1, c - 1).Interior.Color = IIf(grid(r, c) = 1, mlngAlive, mlngDead)
            Next c
        Next r
    End If

    mAbytLast = grid
    mblnHasLast = True

    Application.ScreenUpdating = prevSU
End Sub


Private Sub Class_Terminate()
    Set mWS = Nothing
    Set mrngTopLeft = Nothing
End Sub
