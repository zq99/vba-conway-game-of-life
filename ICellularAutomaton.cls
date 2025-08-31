VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICellularAutomaton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub Init(ByVal width As Long, ByVal height As Long, ByVal wrapEdges As Boolean): End Sub
Public Sub SeedRandom(ByVal aliveProb As Double, Optional ByVal seed As Variant): End Sub
Public Sub StepOne(): End Sub
Public Function Snapshot() As Variant: End Function

