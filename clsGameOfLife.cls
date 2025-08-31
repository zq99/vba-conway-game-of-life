VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGameOfLife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsGameOfLife.cls — Model (rules & state)
' Purpose: Implements Conway’s Game of Life core logic independent of Excel.
'          Owns the grid state and double-buffer stepping; supports wrap-around
'          edges and random seeding; exposes a snapshot for renderers.

' --- Conway's Game of Life: concept & rules ---
' A 2D cellular automaton (John Conway). Each cell is either alive (1) or dead (0)
' and updates in lockstep based on its 8 neighbors (Moore neighborhood).
'
' Rules (B3/S23):
'   Birth   (B3): a dead cell becomes alive if it has exactly 3 live neighbors.
'   Survival(S23): a live cell stays alive with 2 or 3 live neighbors; otherwise it dies.
'
' Deterministic evolution (no randomness after seeding). On finite boards the state
' eventually repeats (still lifes, oscillators, or cycles). With wrapEdges=True this
' class treats the world as a torus (opposite edges touch).
'
' Typical emergent patterns: still lifes (block, beehive), oscillators (blinker, toad),
' and spaceships (glider).


Option Explicit

Implements ICellularAutomaton


Private mlngWidth As Long, mlngHeight As Long, mblnWrap As Boolean
Private mAbytGrid() As Byte, mAbytNext() As Byte
Private mAlngDR(1 To 8) As Long, mAlngDC(1 To 8) As Long


Public Sub Init(ByVal width As Long, ByVal height As Long, ByVal wrapEdges As Boolean)
    Call ICellularAutomaton_Init(width, height, wrapEdges)
End Sub

Public Sub SeedRandom(ByVal aliveProb As Double, Optional ByVal seed As Variant)
    Call ICellularAutomaton_SeedRandom(aliveProb, seed)
End Sub

Public Sub StepOne()
    Call ICellularAutomaton_StepOne
End Sub


Public Function Snapshot() As Variant
    Snapshot = ICellularAutomaton_Snapshot
End Function

Private Sub ICellularAutomaton_Init(ByVal width As Long, ByVal height As Long, ByVal wrapEdges As Boolean)
    mlngWidth = width: mlngHeight = height: mblnWrap = wrapEdges
    ReDim mAbytGrid(1 To mlngHeight, 1 To mlngWidth)
    ReDim mAbytNext(1 To mlngHeight, 1 To mlngWidth)
    Dim i As Long, dr As Long, dc As Long
    i = 0
    For dr = -1 To 1
        For dc = -1 To 1
            If Not (dr = 0 And dc = 0) Then
                i = i + 1: mAlngDR(i) = dr: mAlngDC(i) = dc
            End If
        Next dc
    Next dr
End Sub

Private Sub ICellularAutomaton_SeedRandom(ByVal aliveProb As Double, Optional ByVal seed As Variant)
    If IsMissing(seed) Then Randomize Else Randomize seed
    Dim r As Long, c As Long
    For r = 1 To mlngHeight
        For c = 1 To mlngWidth
            mAbytGrid(r, c) = IIf(Rnd < aliveProb, 1, 0)
        Next c
    Next r
End Sub

Private Sub ICellularAutomaton_StepOne()

' StepOne — advance the automaton by one generation.
' Counts each cell’s 8 neighbors and applies Life rules (B3/S23):
'   - Alive stays alive with 2 or 3 neighbors; otherwise dies.
'   - Dead becomes alive with exactly 3 neighbors.
' Honors toroidal wrap when mWrap=True. Writes next state into mAbytNext, then swaps
' buffers with mAbytGrid. Pure state update—no rendering or I/O.

    Dim r As Long, c As Long, i As Long, rr As Long, cc As Long
    Dim cnt As Integer, alive As Byte
    For r = 1 To mlngHeight
        For c = 1 To mlngWidth
            cnt = 0
            For i = 1 To 8
                rr = r + mAlngDR(i): cc = c + mAlngDC(i)
                If mblnWrap Then
                    If rr < 1 Then rr = mlngHeight
                    If rr > mlngHeight Then rr = 1
                    If cc < 1 Then cc = mlngWidth
                    If cc > mlngWidth Then cc = 1
                    If mAbytGrid(rr, cc) = 1 Then cnt = cnt + 1
                Else
                    If rr >= 1 And rr <= mlngHeight And cc >= 1 And cc <= mlngWidth Then
                        If mAbytGrid(rr, cc) = 1 Then cnt = cnt + 1
                    End If
                End If
            Next i
            alive = mAbytGrid(r, c)
            If alive = 1 Then
                mAbytNext(r, c) = IIf(cnt = 2 Or cnt = 3, 1, 0)
            Else
                mAbytNext(r, c) = IIf(cnt = 3, 1, 0)
            End If
        Next c
    Next r
    Dim tmp As Variant
    tmp = mAbytGrid: mAbytGrid = mAbytNext: mAbytNext = tmp
End Sub

Private Function ICellularAutomaton_Snapshot() As Variant
    ICellularAutomaton_Snapshot = mAbytGrid
End Function

